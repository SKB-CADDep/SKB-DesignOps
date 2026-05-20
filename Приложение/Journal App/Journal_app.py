import sys
import os
import re
import shutil
import sqlite3
import getpass
from datetime import datetime
from collections import defaultdict
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QHBoxLayout, QTableWidget, QTableWidgetItem,
    QPushButton, QLabel, QComboBox, QMessageBox,
    QTabWidget, QFileDialog, QProgressBar, QLineEdit, QSpinBox, QFrame,
    QAbstractItemView, QInputDialog, QHeaderView
)
from PySide6.QtGui import QColor
from PySide6.QtCore import Qt, QTimer, QThread, Signal

import pandas as pd


# ================= НАСТРОЙКИ =================

NETWORK_PATH = Path(r"\\fileserver\УТЗ\Электронная библиотека УТЗ\01_Техническая литература")
AUTO_REFRESH_INTERVAL_MS = 1_000
RATING_MONTH_SLOTS = 3

MONTH_NAMES_RU = {
    1: "Январь",
    2: "Февраль",
    3: "Март",
    4: "Апрель",
    5: "Май",
    6: "Июнь",
    7: "Июль",
    8: "Август",
    9: "Сентябрь",
    10: "Октябрь",
    11: "Ноябрь",
    12: "Декабрь",
}

# Общая БД на файловом сервере
DB_DIR = Path(r"\\fileserver\УТЗ\10 Служба технического директора\05 СКБт\!Общая\Сканы для ГПЯ\!Обработка документов\Журналы")
DB_PATH = str(DB_DIR / "journal_app.db")

# Резервная копия БД
DB_BACKUP_DIR = DB_DIR / "_backup"
DB_BACKUP_PATH = DB_BACKUP_DIR / "journal_app_backup.db"
BACKUP_DELAY_MS = 2_000  # небольшая задержка, чтобы склеивать серию правок в один бэкап

CURRENT_USER = getpass.getuser()

ADMIN_USERS = [
    "Administrator",
    "admin",
    "nyagavrilova"
]

# ================= БАЗА ДАННЫХ =================

def init_db():
    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS issues (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        journal TEXT,
        year TEXT,
        issue TEXT,
        path TEXT,
        status TEXT DEFAULT 'free',
        taken_by TEXT,
        taken_at TEXT,
        completed_at TEXT,
        updated_at TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS export_issues (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        journal TEXT NOT NULL,
        year TEXT NOT NULL,
        issue TEXT NOT NULL,
        status TEXT DEFAULT 'free',
        taken_by TEXT,
        taken_at TEXT,
        completed_at TEXT,
        updated_at TEXT,
        UNIQUE(journal, year, issue)
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS change_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        table_name TEXT NOT NULL,
        row_id INTEGER,
        action TEXT NOT NULL,
        created_at TEXT NOT NULL
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS app_meta (
        key TEXT PRIMARY KEY,
        value TEXT NOT NULL
    )
    """)

    cur.execute("""
    INSERT OR IGNORE INTO app_meta (key, value)
    VALUES ('db_version', ?)
    """, (datetime.now().isoformat(),))

    cur.execute("""
    CREATE TABLE IF NOT EXISTS monthly_plans (
        month TEXT PRIMARY KEY,
        plan_count INTEGER NOT NULL DEFAULT 0,
        created_by TEXT,
        updated_at TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS rating_months (
        slot INTEGER PRIMARY KEY,
        month TEXT NOT NULL,
        updated_at TEXT
    )
    """)

    default_rating_months = [
        f"{datetime.now().year}-06",
        f"{datetime.now().year}-07",
        f"{datetime.now().year}-08"
    ]

    cur.execute("SELECT slot FROM rating_months")
    existing_slots = {row[0] for row in cur.fetchall()}

    for slot, month in enumerate(default_rating_months, start=1):
        if slot not in existing_slots:
            cur.execute("""
                INSERT INTO rating_months (slot, month, updated_at)
                VALUES (?, ?, ?)
            """, (slot, month, datetime.now().isoformat()))

    cur.execute("PRAGMA table_info(issues)")
    issue_columns = {row[1] for row in cur.fetchall()}
    if "updated_at" not in issue_columns:
        cur.execute("ALTER TABLE issues ADD COLUMN updated_at TEXT")

    cur.execute("""
    CREATE TRIGGER IF NOT EXISTS trg_issues_ai
    AFTER INSERT ON issues
    BEGIN
        INSERT INTO change_log(table_name, row_id, action, created_at)
        VALUES ('issues', NEW.id, 'I', datetime('now'));
    END;
    """)
    cur.execute("""
    CREATE TRIGGER IF NOT EXISTS trg_issues_au
    AFTER UPDATE ON issues
    BEGIN
        INSERT INTO change_log(table_name, row_id, action, created_at)
        VALUES ('issues', NEW.id, 'U', datetime('now'));
    END;
    """)
    cur.execute("""
    CREATE TRIGGER IF NOT EXISTS trg_issues_ad
    AFTER DELETE ON issues
    BEGIN
        INSERT INTO change_log(table_name, row_id, action, created_at)
        VALUES ('issues', OLD.id, 'D', datetime('now'));
    END;
    """)

    cur.execute("""
    CREATE TRIGGER IF NOT EXISTS trg_export_issues_ai
    AFTER INSERT ON export_issues
    BEGIN
        INSERT INTO change_log(table_name, row_id, action, created_at)
        VALUES ('export_issues', NEW.id, 'I', datetime('now'));
    END;
    """)
    cur.execute("""
    CREATE TRIGGER IF NOT EXISTS trg_export_issues_au
    AFTER UPDATE ON export_issues
    BEGIN
        INSERT INTO change_log(table_name, row_id, action, created_at)
        VALUES ('export_issues', NEW.id, 'U', datetime('now'));
    END;
    """)
    cur.execute("""
    CREATE TRIGGER IF NOT EXISTS trg_export_issues_ad
    AFTER DELETE ON export_issues
    BEGIN
        INSERT INTO change_log(table_name, row_id, action, created_at)
        VALUES ('export_issues', OLD.id, 'D', datetime('now'));
    END;
    """)

    cur.execute("""
    CREATE TRIGGER IF NOT EXISTS trg_monthly_plans_ai
    AFTER INSERT ON monthly_plans
    BEGIN
        INSERT INTO change_log(table_name, row_id, action, created_at)
        VALUES ('monthly_plans', NULL, 'I', datetime('now'));
    END;
    """)
    cur.execute("""
    CREATE TRIGGER IF NOT EXISTS trg_monthly_plans_au
    AFTER UPDATE ON monthly_plans
    BEGIN
        INSERT INTO change_log(table_name, row_id, action, created_at)
        VALUES ('monthly_plans', NULL, 'U', datetime('now'));
    END;
    """)
    cur.execute("""
    CREATE TRIGGER IF NOT EXISTS trg_monthly_plans_ad
    AFTER DELETE ON monthly_plans
    BEGIN
        INSERT INTO change_log(table_name, row_id, action, created_at)
        VALUES ('monthly_plans', NULL, 'D', datetime('now'));
    END;
    """)

    cur.execute("""
    CREATE TRIGGER IF NOT EXISTS trg_rating_months_ai
    AFTER INSERT ON rating_months
    BEGIN
        INSERT INTO change_log(table_name, row_id, action, created_at)
        VALUES ('rating_months', NULL, 'I', datetime('now'));
    END;
    """)
    cur.execute("""
    CREATE TRIGGER IF NOT EXISTS trg_rating_months_au
    AFTER UPDATE ON rating_months
    BEGIN
        INSERT INTO change_log(table_name, row_id, action, created_at)
        VALUES ('rating_months', NULL, 'U', datetime('now'));
    END;
    """)
    cur.execute("""
    CREATE TRIGGER IF NOT EXISTS trg_rating_months_ad
    AFTER DELETE ON rating_months
    BEGIN
        INSERT INTO change_log(table_name, row_id, action, created_at)
        VALUES ('rating_months', NULL, 'D', datetime('now'));
    END;
    """)

    conn.commit()
    conn.close()


def ensure_db_dir():
    if not DB_DIR.exists():
        raise sqlite3.OperationalError(
            f"Папка для БД не найдена: {DB_DIR}"
        )

def get_db_connection():
    ensure_db_dir()
    conn = sqlite3.connect(DB_PATH, timeout=3)
    conn.execute("PRAGMA busy_timeout = 3000")
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

def restore_db_from_backup_if_needed():
    """
    Если основная БД удалена, но есть резервная копия,
    восстанавливаем её до запуска приложения.
    """
    if Path(DB_PATH).exists() or not DB_BACKUP_PATH.exists():
        return False

    ensure_db_dir()
    DB_BACKUP_DIR.mkdir(parents=True, exist_ok=True)

    try:
        shutil.copy2(DB_BACKUP_PATH, DB_PATH)
    except OSError as exc:
        raise sqlite3.OperationalError(
            f"Не удалось восстановить БД из резервной копии: {exc}"
        ) from exc

    return True


class DatabaseBackupThread(QThread):
    backup_finished = Signal(bool, str)

    def run(self):
        temp_path = DB_BACKUP_PATH.with_name(DB_BACKUP_PATH.name + ".tmp")

        try:
            source_path = Path(DB_PATH)
            if not source_path.exists():
                raise FileNotFoundError(f"Не найдена исходная БД: {source_path}")

            DB_BACKUP_DIR.mkdir(parents=True, exist_ok=True)

            if temp_path.exists():
                temp_path.unlink()

            source = sqlite3.connect(DB_PATH, timeout=5)
            source.execute("PRAGMA busy_timeout = 5000")

            destination = sqlite3.connect(str(temp_path), timeout=5)
            try:
                # Делает консистентную копию БД без блокировки UI
                source.backup(destination, pages=200, sleep=0.05)
                destination.commit()
            finally:
                destination.close()
                source.close()

            os.replace(temp_path, DB_BACKUP_PATH)
            self.backup_finished.emit(True, str(DB_BACKUP_PATH))

        except Exception as exc:
            try:
                if temp_path.exists():
                    temp_path.unlink()
            except OSError:
                pass

            self.backup_finished.emit(False, str(exc))


def extract_issue_metadata_from_pdf(pdf_path):
    pdf_path = Path(pdf_path)
    parents = list(pdf_path.parents)

    for idx, parent in enumerate(parents):
        year_match = re.search(r"(19|20)\d{2}", parent.name)
        if not year_match:
            continue

        year = year_match.group(0)
        journal = parents[idx + 1].name.strip() if idx + 1 < len(parents) else ""

        if journal.startswith(("Ж", "ж")):
            return journal, year

    return None, None


def import_pdf_paths(pdf_paths, fallback_journal=None, fallback_year=None):
    conn = get_db_connection()
    cur = conn.cursor()

    fallback_journal = (fallback_journal or "").strip()
    fallback_year = (fallback_year or "").strip()

    inserted = 0

    for pdf_path in sorted({Path(p) for p in pdf_paths}, key=lambda p: p.as_posix().lower()):
        if not pdf_path.is_file() or pdf_path.suffix.lower() != ".pdf":
            continue

        journal, year = extract_issue_metadata_from_pdf(pdf_path)

        if (not journal or not year) and fallback_journal and re.fullmatch(r"\d{4}", fallback_year):
            journal, year = fallback_journal, fallback_year

        if not journal or not year:
            continue

        path = str(pdf_path)

        cur.execute("SELECT id FROM issues WHERE path=?", (path,))
        if not cur.fetchone():
            cur.execute("""
                INSERT INTO issues (journal, year, issue, path, updated_at)
                VALUES (?, ?, ?, ?, ?)
            """, (journal, year, pdf_path.stem, path, datetime.now().isoformat()))
            inserted += 1

    conn.commit()
    conn.close()
    return inserted


def load_from_folder(root_path=NETWORK_PATH, fallback_journal=None, fallback_year=None):
    root = Path(root_path)
    if not root.exists():
        return 0

    return import_pdf_paths(root.rglob("*.pdf"), fallback_journal, fallback_year)

# ================= ГЛАВНОЕ ОКНО =================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Журналы WindChill")
        self.resize(1100, 600)

        self.is_admin = CURRENT_USER in ADMIN_USERS
        self.search_timer = QTimer(self)
        self.search_timer.setSingleShot(True)
        self.search_timer.setInterval(300)
        self.search_timer.timeout.connect(self.load_data)

        self.export_search_timer = QTimer(self)
        self.export_search_timer.setSingleShot(True)
        self.export_search_timer.setInterval(300)
        self.export_search_timer.timeout.connect(self.load_export_data)

        self.refresh_timer = QTimer(self)
        self.refresh_timer.setSingleShot(False)
        self.refresh_timer.setInterval(AUTO_REFRESH_INTERVAL_MS)
        self.refresh_timer.timeout.connect(self.refresh_ui)
        self.last_db_signature = None
        self._preserve_scroll_on_reload = False
        self._preserve_export_scroll_on_reload = False

        self._backup_enabled = True
        self._backup_requested = False
        self.backup_thread = None

        self.backup_timer = QTimer(self)
        self.backup_timer.setSingleShot(True)
        self.backup_timer.setInterval(BACKUP_DELAY_MS)
        self.backup_timer.timeout.connect(self.start_backup_thread)

        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        self.main_tab = QWidget()
        self.export_tab = QWidget()
        self.rating_tab = QWidget()

        self.tabs.addTab(self.main_tab, "📚 Журналы")
        self.tabs.addTab(self.export_tab, "📤 Скачать из Elibrary и выгрузить в WindChill")
        self.tabs.addTab(self.rating_tab, "🏆 Рейтинг")
        self.tabs.currentChanged.connect(self.on_tab_changed)

        self.setup_main_tab()
        self.setup_export_tab()
        self.setup_rating_tab()
        self.load_data()
        self.load_export_data()
        self.remember_db_signature()
        self.refresh_timer.start()

    def closeEvent(self, event):
        self.backup_timer.stop()

        # Если есть ожидающий бэкап — запускаем его перед закрытием
        if self._backup_requested and not (self.backup_thread and self.backup_thread.isRunning()):
            self.start_backup_thread()

        # Дожидаемся завершения текущего бэкапа, чтобы не потерять последние изменения
        if self.backup_thread and self.backup_thread.isRunning():
            self.backup_thread.wait(5000)

        super().closeEvent(event)

    def sync_from_network(self):
        if not self.is_admin:
            return

        folder_path = QFileDialog.getExistingDirectory(
            self,
            "Выберите папку с PDF",
            str(NETWORK_PATH)
        )
        if not folder_path:
            return

        imported = load_from_folder(folder_path)

        if imported == 0:
            fallback_journal, fallback_year = self.ask_manual_import_metadata()
            if fallback_journal and fallback_year:
                imported = load_from_folder(
                    folder_path,
                    fallback_journal=fallback_journal,
                    fallback_year=fallback_year
                )

        self.load_journal_filter_options()
        self.load_data()
        self.remember_db_signature()

        QMessageBox.information(
            self,
            "Импорт завершён",
            f"Добавлено новых записей: {imported}"
        )

    def schedule_load_data(self):
        self.search_timer.start()

    def get_default_rating_months(self):
        year = datetime.now().year
        return [
            f"{year}-06",
            f"{year}-07",
            f"{year}-08"
        ]

    def get_rating_months(self):
        default_months = self.get_default_rating_months()

        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("""
            SELECT slot, month
            FROM rating_months
            ORDER BY slot
        """)
        rows = cur.fetchall()
        conn.close()

        months_by_slot = {slot: month for slot, month in rows}
        result = []

        for slot in range(1, RATING_MONTH_SLOTS + 1):
            month = months_by_slot.get(slot, default_months[slot - 1])
            try:
                datetime.strptime(month, "%Y-%m")
            except ValueError:
                month = default_months[slot - 1]
            result.append(month)

        return result

    def save_rating_months(self):
        if not self.is_admin or not hasattr(self, "rating_month_edits"):
            return

        months = []
        for edit in self.rating_month_edits:
            month = edit.text().strip()
            try:
                datetime.strptime(month, "%Y-%m")
            except ValueError:
                QMessageBox.warning(
                    self,
                    "Ошибка",
                    "Введите месяц в формате YYYY-MM, например 2026-06."
                )
                return
            months.append(month)

        conn = get_db_connection()
        cur = conn.cursor()

        for slot, month in enumerate(months, start=1):
            cur.execute("""
                INSERT INTO rating_months (slot, month, updated_at)
                VALUES (?, ?, ?)
                ON CONFLICT(slot) DO UPDATE SET
                    month=excluded.month,
                    updated_at=excluded.updated_at
            """, (slot, month, datetime.now().isoformat()))

        conn.commit()
        conn.close()

        self.update_rating()

    def get_month_display_name(self, month_str):
        try:
            dt = datetime.strptime(month_str, "%Y-%m")
        except ValueError:
            return "Не задан"

        month_name = MONTH_NAMES_RU.get(dt.month, "Месяц")
        return f"{month_name} {dt.year}"

    def get_month_ranking(self, month_str):
        if not month_str:
            return []

        try:
            datetime.strptime(month_str, "%Y-%m")
        except ValueError:
            return []

        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("""
            SELECT taken_by,
                   COUNT(*) AS cnt,
                   MIN(completed_at) AS first_done_at,
                   MIN(id) AS first_issue_id
            FROM issues
            WHERE status='done'
              AND completed_at LIKE ?
              AND taken_by IS NOT NULL
              AND TRIM(taken_by) != ''
            GROUP BY taken_by
            ORDER BY cnt DESC, first_done_at ASC, first_issue_id ASC
        """, (month_str + "%",))
        rows = cur.fetchall()
        conn.close()

        return [(taken_by, cnt) for taken_by, cnt, _, _ in rows]

    def format_month_ranking_text(self, ranking_rows):
        if not ranking_rows:
            return "Нет завершенных журналов"

        def format_place(place: int) -> str:
            if place == 1:
                return "🥇 1 место"
            if place == 2:
                return "🥈 2 место"
            if place == 3:
                return "🥉 3 место"
            return f"{place} место"

        lines = []
        current_place = 1
        current_count = None
        current_users = []

        def flush_group(place, users, count):
            if not users:
                return
            users_text = "; ".join(f"{user} ({count})" for user in users)
            lines.append(f"{format_place(place)} — {users_text}")

        for user, count in ranking_rows:
            if current_count is None:
                current_count = count
                current_users = [user]
                continue

            if count == current_count:
                current_users.append(user)
            else:
                flush_group(current_place, current_users, current_count)
                current_place += 1
                current_count = count
                current_users = [user]

        flush_group(current_place, current_users, current_count)

        return "\n".join(lines)

    def on_tab_changed(self, index):
        if index == self.tabs.indexOf(self.main_tab) and hasattr(self, "table"):
            self.load_journal_filter_options()
            self.load_data()

        elif index == self.tabs.indexOf(self.export_tab) and hasattr(self, "export_table"):
            self.load_export_journal_filter_options()
            self.load_export_data()

        elif index == self.tabs.indexOf(self.rating_tab) and hasattr(self, "rating_month_rank_labels"):
            self.update_rating()

    def get_db_signature(self):
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("""
            SELECT COALESCE(MAX(id), 0)
            FROM change_log
        """)
        row = cur.fetchone()
        conn.close()
        return row[0] if row else 0

    def get_changes_since(self, last_id):
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("""
            SELECT id, table_name, row_id, action
            FROM change_log
            WHERE id > ?
            ORDER BY id
        """, (last_id,))
        rows = cur.fetchall()
        conn.close()
        return rows

    def remember_db_signature(self):
        try:
            self.last_db_signature = self.get_db_signature()
            if self._backup_enabled:
                self.schedule_backup()
        except sqlite3.OperationalError:
            self.last_db_signature = None

    def schedule_backup(self):
        self._backup_requested = True
        self.backup_timer.start(BACKUP_DELAY_MS)

    def start_backup_thread(self):
        if not self._backup_requested:
            return

        if self.backup_thread and self.backup_thread.isRunning():
            return

        self._backup_requested = False
        self.backup_thread = DatabaseBackupThread()
        self.backup_thread.backup_finished.connect(self.on_backup_finished)
        self.backup_thread.start()

    def on_backup_finished(self, success, message):
        if self.backup_thread is not None:
            self.backup_thread.deleteLater()
            self.backup_thread = None

        if not success:
            print(f"Ошибка резервного копирования: {message}")

        # Если пока копировали пришли новые изменения — ставим ещё один бэкап
        if self._backup_requested:
            self.backup_timer.start(BACKUP_DELAY_MS)

    def refresh_ui(self):
        if not self.isVisible():
            return

        try:
            latest_change_id = self.get_db_signature()
        except sqlite3.OperationalError:
            return

        if latest_change_id == self.last_db_signature:
            return

        changes = self.get_changes_since(self.last_db_signature or 0)
        self.last_db_signature = latest_change_id

        if not changes:
            return

        current_tab = self.tabs.currentWidget()

        issue_changes = [row_id for _, table_name, row_id, _ in changes if table_name == "issues"]
        export_changes = [row_id for _, table_name, row_id, _ in changes if table_name == "export_issues"]
        rating_changed = any(
            table_name in {"issues", "monthly_plans", "rating_months"}
            for _, table_name, _, _ in changes
        )

        if current_tab == self.main_tab:
            if issue_changes:
                self.load_journal_filter_options()

                issue_actions = {action for _, table_name, _, action in changes if table_name == "issues"}
                if "I" in issue_actions or "D" in issue_actions:
                    self._preserve_scroll_on_reload = True
                    try:
                        self.load_data()
                    finally:
                        self._preserve_scroll_on_reload = False
                else:
                    for row_id in sorted(set(issue_changes)):
                        self.refresh_issue_row(row_id)

        elif current_tab == self.export_tab:
            if export_changes:
                self.load_export_journal_filter_options()

                export_actions = {action for _, table_name, _, action in changes if table_name == "export_issues"}
                if "I" in export_actions or "D" in export_actions:
                    self._preserve_export_scroll_on_reload = True
                    try:
                        self.load_export_data()
                    finally:
                        self._preserve_export_scroll_on_reload = False
                else:
                    for row_id in sorted(set(export_changes)):
                        self.refresh_export_issue_row(row_id)

        elif current_tab == self.rating_tab:
            if rating_changed:
                self.update_rating()

    def load_journal_filter_options(self):
        current = self.journal_filter_box.currentText() if hasattr(self, "journal_filter_box") else "Все журналы"

        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("""
            SELECT DISTINCT journal
            FROM issues
            WHERE journal IS NOT NULL AND TRIM(journal) != ''
            ORDER BY journal COLLATE NOCASE
        """)
        journals = [row[0] for row in cur.fetchall()]
        conn.close()

        self.journal_filter_box.blockSignals(True)
        self.journal_filter_box.clear()
        self.journal_filter_box.addItem("Все журналы")
        self.journal_filter_box.addItems(journals)

        index = self.journal_filter_box.findText(current)
        self.journal_filter_box.setCurrentIndex(index if index >= 0 else 0)
        self.journal_filter_box.blockSignals(False)

    def get_issue_lock_info(self, issue_id):
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("SELECT status, taken_by FROM issues WHERE id=?", (issue_id,))
        row = cur.fetchone()
        conn.close()
        return row if row else (None, None)

    def warn_issue_locked(self, taken_by):
        QMessageBox.warning(
            self,
            "Внимание",
            f"Журнал уже в работе у пользователя {taken_by}."
        )

    def ask_manual_import_metadata(self):
        journal, ok = QInputDialog.getText(
            self,
            "Импорт журналов",
            "Не удалось определить журнал автоматически.\nВведите название журнала:"
        )
        if not ok or not journal.strip():
            return None, None

        year, ok = QInputDialog.getText(
            self,
            "Импорт журналов",
            "Введите год в формате YYYY, например 2024:"
        )
        if not ok or not re.fullmatch(r"\d{4}", year.strip()):
            QMessageBox.warning(
                self,
                "Ошибка",
                "Год должен быть в формате YYYY, например 2024."
            )
            return None, None

        return journal.strip(), year.strip()

    def delete_selected_issue(self):
        if not self.is_admin:
            return

        if not self.table.selectionModel():
            QMessageBox.warning(
                self,
                "Внимание",
                "Сначала выберите журналы в таблице."
            )
            return

        selected_rows = sorted(
            {index.row() for index in self.table.selectionModel().selectedRows()}
        )

        if not selected_rows:
            QMessageBox.warning(
                self,
                "Внимание",
                "Сначала выберите журналы в таблице."
            )
            return

        issues_to_delete = []

        for row in selected_rows:
            journal_item = self.table.item(row, 0)
            year_item = self.table.item(row, 1)
            issue_item = self.table.item(row, 2)

            if not journal_item:
                continue

            issue_id = journal_item.data(Qt.UserRole)
            if issue_id is None:
                continue

            journal = journal_item.text().strip()
            year = year_item.text().strip() if year_item else ""
            issue = issue_item.text().strip() if issue_item else ""

            issues_to_delete.append((issue_id, journal, year, issue))

        if not issues_to_delete:
            QMessageBox.warning(
                self,
                "Внимание",
                "Не удалось определить выбранные записи."
            )
            return

        preview_lines = [
            f"• {journal} | {year} | {issue}"
            for _, journal, year, issue in issues_to_delete[:5]
        ]

        preview_text = "\n".join(preview_lines)
        if len(issues_to_delete) > 5:
            preview_text += f"\n... и ещё {len(issues_to_delete) - 5}"

        reply = QMessageBox.question(
            self,
            "Удаление журналов",
            f"Удалить выбранные записи?\n\n"
            f"Количество: {len(issues_to_delete)}\n\n"
            f"{preview_text}\n\n"
            f"Записи будут удалены из общей базы.\n"
            f"Если PDF-файлы останутся в сетевой папке, "
            f"они могут снова появиться после импорта.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply != QMessageBox.Yes:
            return

        conn = get_db_connection()
        cur = conn.cursor()
        for issue_id, _, _, _ in issues_to_delete:
            cur.execute("DELETE FROM issues WHERE id=?", (issue_id,))
        conn.commit()
        conn.close()

        self.load_journal_filter_options()
        self._preserve_scroll_on_reload = True
        try:
            self.load_data()
        finally:
            self._preserve_scroll_on_reload = False

        self.update_rating()
        self.remember_db_signature()

    def get_issue_row_data(self, issue_id):
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("""
            SELECT id, journal, year, issue, path, status, taken_by, taken_at, completed_at
            FROM issues
            WHERE id=?
        """, (issue_id,))
        row = cur.fetchone()
        conn.close()
        return row

    def find_issue_row(self, issue_id):
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            if item and item.data(Qt.UserRole) == issue_id:
                return row
        return -1

    def issue_matches_current_view(self, journal, year, issue, path, status, taken_by, taken_at, completed_at):
        filter_value = self.filter_box.currentText()
        journal_filter = self.journal_filter_box.currentText()
        search_text = self.search_box.text().strip().casefold()

        if filter_value == "Свободные" and status != "free":
            return False
        if filter_value == "В работе" and status != "in_progress":
            return False
        if filter_value == "Готово" and status != "done":
            return False
        if filter_value == "Мои" and taken_by != CURRENT_USER:
            return False

        if journal_filter != "Все журналы" and (journal or "").strip().casefold() != journal_filter.strip().casefold():
            return False

        row_search = " ".join([
            str(journal or ""),
            str(year or ""),
            str(issue or ""),
            str(path or ""),
            "Свободен" if status == "free" else
            "В работе" if status == "in_progress" else
            "Готово" if status == "done" else str(status),
            str(taken_by or ""),
            str(taken_at or ""),
            str(completed_at or "")
        ]).casefold()

        if search_text and search_text not in row_search:
            return False

        return True

    def refresh_issue_row(self, issue_id):
        row_data = self.get_issue_row_data(issue_id)
        if not row_data:
            row = self.find_issue_row(issue_id)
            if row >= 0:
                self.table.removeRow(row)
            return

        id_, journal, year, issue, path, status, taken_by, taken_at, completed_at = row_data

        row = self.find_issue_row(issue_id)

        # Если текущие фильтры/поиск уже не подходят — просто убираем строку из таблицы
        if not self.issue_matches_current_view(journal, year, issue, path, status, taken_by, taken_at, completed_at):
            if row >= 0:
                self.table.removeRow(row)
            return

        # Если строка не найдена, но должна быть видна — делаем мягкий полный reload с сохранением прокрутки
        if row < 0:
            self._preserve_scroll_on_reload = True
            try:
                self.load_data()
            finally:
                self._preserve_scroll_on_reload = False
            return

        self.table.blockSignals(True)
        try:
            journal_item = self.table.item(row, 0) or QTableWidgetItem()
            journal_item.setText(journal or "")
            journal_item.setData(Qt.UserRole, id_)
            if self.is_admin:
                journal_item.setFlags(journal_item.flags() | Qt.ItemIsEditable)
            self.table.setItem(row, 0, journal_item)

            year_item = self.table.item(row, 1) or QTableWidgetItem()
            year_item.setText(year or "")
            year_item.setData(Qt.UserRole, id_)
            if self.is_admin:
                year_item.setFlags(year_item.flags() | Qt.ItemIsEditable)
            self.table.setItem(row, 1, year_item)

            issue_item = self.table.item(row, 2) or QTableWidgetItem()
            issue_item.setText(issue or "")
            issue_item.setData(Qt.UserRole, id_)
            if self.is_admin:
                issue_item.setFlags(issue_item.flags() | Qt.ItemIsEditable)
            self.table.setItem(row, 2, issue_item)

            who_item = self.table.item(row, 4) or QTableWidgetItem()
            who_item.setText(taken_by or "")
            who_item.setData(Qt.UserRole, id_)
            if self.is_admin:
                who_item.setFlags(who_item.flags() | Qt.ItemIsEditable)
            self.table.setItem(row, 4, who_item)

            if self.is_admin:
                status_combo = QComboBox()
                status_combo.addItem("⚪ Свободен", "free")
                status_combo.addItem("🟡 В работе", "in_progress")
                status_combo.addItem("✅ Готово", "done")
                status_combo.setCurrentIndex(status_combo.findData(status))
                status_combo.currentIndexChanged.connect(
                    lambda _, issue_id=id_, combo=status_combo: self.change_status(issue_id, combo.currentData())
                )
                self.table.setCellWidget(row, 3, status_combo)
            else:
                status_item = QTableWidgetItem()
                if status == "free":
                    status_item.setText("⚪ Свободен")
                elif status == "in_progress":
                    status_item.setText("🟡 В работе")
                    status_item.setBackground(QColor("#fff3cd"))
                elif status == "done":
                    status_item.setText("✅ Готово")
                    status_item.setBackground(QColor("#d4edda"))
                self.table.setItem(row, 3, status_item)

            btn = QPushButton()
            if status == "free":
                btn.setText("Взять")
                btn.clicked.connect(lambda _, i=id_: self.take_issue(i))
            elif status == "in_progress" and taken_by == CURRENT_USER:
                btn.setText("Готово")
                btn.clicked.connect(lambda _, i=id_: self.complete_issue(i))
            else:
                btn.setText("—")
                btn.setEnabled(False)

            self.table.setCellWidget(row, 5, btn)
        finally:
            self.table.blockSignals(False)

    # ================= ГЛАВНАЯ ВКЛАДКА =================

    def setup_main_tab(self):
        layout = QVBoxLayout()

        # ================= КАРТОЧКА ФИЛЬТРОВ =================
        self.filter_frame = QFrame()
        self.filter_frame.setObjectName("ratingCard")
        filter_frame_layout = QVBoxLayout(self.filter_frame)

        filter_layout = QHBoxLayout()

        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Поиск по всем столбцам...")
        self.search_box.textChanged.connect(self.schedule_load_data)

        self.journal_filter_box = QComboBox()

        self.filter_box = QComboBox()
        self.filter_box.addItems(["Все", "Свободные", "В работе", "Готово", "Мои"])

        if self.is_admin:
            refresh_btn = QPushButton("🔄 Обновить")
            refresh_btn.clicked.connect(self.load_data)

            delete_btn = QPushButton("🗑 Удалить выбранный журнал")
            delete_btn.clicked.connect(self.delete_selected_issue)

        export_btn = QPushButton("📊 Экспорт отчета")
        export_btn.clicked.connect(self.export_report)

        filter_layout.addWidget(QLabel("Поиск:"))
        filter_layout.addWidget(self.search_box)
        filter_layout.addWidget(QLabel("Журнал:"))
        filter_layout.addWidget(self.journal_filter_box)
        filter_layout.addWidget(QLabel("Статус:"))
        filter_layout.addWidget(self.filter_box)
        filter_layout.addStretch()

        if self.is_admin:
            sync_btn = QPushButton("📥 Импорт журналов")
            sync_btn.clicked.connect(self.sync_from_network)
            filter_layout.addWidget(sync_btn)

        if self.is_admin:
            filter_layout.addWidget(delete_btn)
            filter_layout.addWidget(refresh_btn)

        filter_layout.addWidget(export_btn)

        filter_frame_layout.addLayout(filter_layout)

        # ================= КАРТОЧКА ТАБЛИЦЫ =================
        self.table_frame = QFrame()
        self.table_frame.setObjectName("ratingCard")
        table_frame_layout = QVBoxLayout(self.table_frame)

        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(
            ["Журнал", "Год", "Выпуск", "Статус", "Кто", "Действие"]
        )
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)

        # 🔒 Запрет редактирования таблицы
        if self.is_admin:
            self.table.setEditTriggers(
                QTableWidget.DoubleClicked |
                QTableWidget.EditKeyPressed |
                QTableWidget.AnyKeyPressed
            )
            self.table.itemChanged.connect(self.on_table_item_changed)
        else:
            self.table.setEditTriggers(QTableWidget.NoEditTriggers)

        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setAlternatingRowColors(True)

        self.table.setColumnWidth(0, 240)
        self.table.setColumnWidth(1, 300)
        self.table.setColumnWidth(2, 800)
        self.table.setColumnWidth(3, 110)
        self.table.setColumnWidth(4, 120)
        self.table.setColumnWidth(5, 90)

        self.table.setStyleSheet("""
            QTableWidget {
                background-color: #ffffff;
                alternate-background-color: #f7f7f7;
            }
            QHeaderView::section {
                background-color: #e9eef5;
                color: #1f2937;
                padding: 6px 8px;
                border: 1px solid #d7dde5;
                font-weight: bold;
            }
            QTableWidget::item:selected {
                background-color: #0078d7;
                color: white;
            }
            QTableWidget::item:selected:active {
                background-color: #0078d7;
                color: white;
            }
        """)

        table_frame_layout.addWidget(self.table)

        self.journal_filter_box.currentTextChanged.connect(self.schedule_load_data)
        self.filter_box.currentTextChanged.connect(self.schedule_load_data)
        self.load_journal_filter_options()

        layout.addWidget(self.filter_frame)
        layout.addWidget(self.table_frame)

        self.main_tab.setLayout(layout)

        self.main_tab.setStyleSheet("""
            QFrame#ratingCard {
                background-color: #ffffff;
                border: 1px solid #d7dde5;
                border-radius: 10px;
                padding: 10px;
            }
        """)

    def setup_export_tab(self):
        layout = QVBoxLayout()

        # ================= КАРТОЧКА ФИЛЬТРОВ =================
        self.export_filter_frame = QFrame()
        self.export_filter_frame.setObjectName("ratingCard")
        filter_frame_layout = QVBoxLayout(self.export_filter_frame)

        filter_layout = QHBoxLayout()

        self.export_search_box = QLineEdit()
        self.export_search_box.setPlaceholderText("Поиск по всем столбцам...")
        self.export_search_box.textChanged.connect(self.schedule_export_load_data)

        self.export_journal_filter_box = QComboBox()

        self.export_filter_box = QComboBox()
        self.export_filter_box.addItems(["Все", "Свободные", "В работе", "Готово", "Мои"])

        filter_layout.addWidget(QLabel("Поиск:"))
        filter_layout.addWidget(self.export_search_box)
        filter_layout.addWidget(QLabel("Журнал:"))
        filter_layout.addWidget(self.export_journal_filter_box)
        filter_layout.addWidget(QLabel("Статус:"))
        filter_layout.addWidget(self.export_filter_box)
        filter_layout.addStretch()

        if self.is_admin:
            import_btn = QPushButton("📥 Импорт из Excel")
            import_btn.clicked.connect(self.import_export_from_excel)

            delete_btn = QPushButton("🗑 Удалить выбранное")
            delete_btn.clicked.connect(self.delete_selected_export_issue)

            refresh_btn = QPushButton("🔄 Обновить")
            refresh_btn.clicked.connect(self.load_export_data)

            filter_layout.addWidget(import_btn)
            filter_layout.addWidget(delete_btn)
            filter_layout.addWidget(refresh_btn)

        filter_frame_layout.addLayout(filter_layout)

        # ================= СТАТИСТИКА ВЫГРУЗКИ =================
        self.export_stats_frame = QFrame()
        self.export_stats_frame.setObjectName("ratingCard")
        stats_layout = QVBoxLayout(self.export_stats_frame)

        stats_title = QLabel("Статистика выгрузки")
        stats_title.setStyleSheet("font-weight: 700; font-size: 14px; color: #1f2937;")

        self.export_summary_label = QLabel()
        self.export_summary_label.setWordWrap(True)
        self.export_summary_label.setStyleSheet("font-size: 14px;")

        self.export_progress_table = QTableWidget()
        self.export_progress_table.setColumnCount(4)
        self.export_progress_table.setHorizontalHeaderLabels(
            ["Журнал", "Выгружено", "Всего", "%"]
        )
        self.export_progress_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.export_progress_table.setSelectionMode(QAbstractItemView.NoSelection)
        self.export_progress_table.setFocusPolicy(Qt.NoFocus)
        self.export_progress_table.setAlternatingRowColors(True)
        self.export_progress_table.setShowGrid(False)
        self.export_progress_table.setWordWrap(False)
        self.export_progress_table.setTextElideMode(Qt.ElideRight)
        self.export_progress_table.verticalHeader().setVisible(False)
        self.export_progress_table.verticalHeader().setDefaultSectionSize(28)

        header = self.export_progress_table.horizontalHeader()
        header.setStretchLastSection(False)
        header.setDefaultAlignment(Qt.AlignCenter)
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)

        self.export_stats_frame.setMinimumWidth(380)
        self.export_progress_table.setMinimumWidth(360)
        self.export_progress_table.setStyleSheet("""
            QTableWidget {
                background-color: #ffffff;
                alternate-background-color: #f7f7f7;
            }
            QHeaderView::section {
                background-color: #e9eef5;
                color: #1f2937;
                padding: 6px 8px;
                border: 1px solid #d7dde5;
                font-weight: bold;
            }
            QTableWidget::item:selected {
                background-color: #0078d7;
                color: white;
            }
            QTableWidget::item:selected:active {
                background-color: #0078d7;
                color: white;
            }
        """)

        stats_layout.addWidget(stats_title)
        stats_layout.addWidget(self.export_summary_label)
        stats_layout.addWidget(self.export_progress_table)

        # ================= КАРТОЧКА ТАБЛИЦЫ =================
        self.export_table_frame = QFrame()
        self.export_table_frame.setObjectName("ratingCard")
        table_frame_layout = QVBoxLayout(self.export_table_frame)

        self.export_table = QTableWidget()
        self.export_table.setColumnCount(6)
        self.export_table.setHorizontalHeaderLabels(
            ["Журнал", "Год", "Выпуск", "Статус", "Кто", "Действие"]
        )
        self.export_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.export_table.setSelectionMode(QAbstractItemView.ExtendedSelection)

        if self.is_admin:
            self.export_table.setEditTriggers(
                QTableWidget.DoubleClicked |
                QTableWidget.EditKeyPressed |
                QTableWidget.AnyKeyPressed
            )
            self.export_table.itemChanged.connect(self.on_export_table_item_changed)
        else:
            self.export_table.setEditTriggers(QTableWidget.NoEditTriggers)

        self.export_table.horizontalHeader().setStretchLastSection(False)
        self.export_table.setAlternatingRowColors(True)

        self.export_table.setColumnWidth(0, 240)
        self.export_table.setColumnWidth(1, 90)
        self.export_table.setColumnWidth(2, 100)
        self.export_table.setColumnWidth(3, 110)
        self.export_table.setColumnWidth(4, 120)
        self.export_table.setColumnWidth(5, 95)

        self.export_table.setStyleSheet("""
            QTableWidget {
                background-color: #ffffff;
                alternate-background-color: #f7f7f7;
            }
            QHeaderView::section {
                background-color: #e9eef5;
                color: #1f2937;
                padding: 6px 8px;
                border: 1px solid #d7dde5;
                font-weight: bold;
            }
            QTableWidget::item:selected {
                background-color: #0078d7;
                color: white;
            }
            QTableWidget::item:selected:active {
                background-color: #0078d7;
                color: white;
            }
        """)

        table_frame_layout.addWidget(self.export_table)

        self.export_journal_filter_box.currentTextChanged.connect(self.schedule_export_load_data)
        self.export_filter_box.currentTextChanged.connect(self.schedule_export_load_data)

        self.load_export_journal_filter_options()

        layout.addWidget(self.export_filter_frame)

        content_layout = QHBoxLayout()
        content_layout.setSpacing(10)
        content_layout.addWidget(self.export_table_frame, 3)
        content_layout.addWidget(self.export_stats_frame, 2)

        layout.addLayout(content_layout)

        self.export_tab.setLayout(layout)

        self.export_tab.setStyleSheet("""
            QFrame#ratingCard {
                background-color: #ffffff;
                border: 1px solid #d7dde5;
                border-radius: 10px;
                padding: 10px;
            }
        """)

    def schedule_export_load_data(self):
        self.export_search_timer.start()

    def load_export_journal_filter_options(self):
        current = self.export_journal_filter_box.currentText() if hasattr(self, "export_journal_filter_box") else "Все журналы"

        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("""
            SELECT DISTINCT journal
            FROM export_issues
            WHERE journal IS NOT NULL AND TRIM(journal) != ''
            ORDER BY journal COLLATE NOCASE
        """)
        journals = [row[0] for row in cur.fetchall()]
        conn.close()

        self.export_journal_filter_box.blockSignals(True)
        self.export_journal_filter_box.clear()
        self.export_journal_filter_box.addItem("Все журналы")
        self.export_journal_filter_box.addItems(journals)

        index = self.export_journal_filter_box.findText(current)
        self.export_journal_filter_box.setCurrentIndex(index if index >= 0 else 0)
        self.export_journal_filter_box.blockSignals(False)

    def update_export_progress_table(self):
        conn = get_db_connection()
        cur = conn.cursor()

        cur.execute("""
            SELECT COUNT(*),
                   COALESCE(SUM(CASE WHEN status='done' THEN 1 ELSE 0 END), 0)
            FROM export_issues
        """)
        total_row = cur.fetchone() or (0, 0)
        total_count = total_row[0] or 0
        done_count = total_row[1] or 0
        percent_total = int((done_count / total_count) * 100) if total_count else 0

        cur.execute("""
            SELECT journal,
                   SUM(CASE WHEN status='done' THEN 1 ELSE 0 END) AS done_count,
                   COUNT(*) AS total_count
            FROM export_issues
            GROUP BY journal
            ORDER BY journal COLLATE NOCASE
        """)
        journal_rows = cur.fetchall()
        conn.close()

        self.export_summary_label.setText(
            f"Всего: {total_count} | Выгружено: {done_count} | Прогресс: {percent_total}%"
        )

        self.export_progress_table.setRowCount(0)
        self.export_progress_table.setUpdatesEnabled(False)
        self.export_progress_table.blockSignals(True)

        for journal, done_count, total_count in journal_rows:
            journal_name = journal or "Без названия"
            percent = int((done_count / total_count) * 100) if total_count else 0

            row = self.export_progress_table.rowCount()
            self.export_progress_table.insertRow(row)
            self.export_progress_table.setRowHeight(row, 28)

            journal_item = QTableWidgetItem(journal_name)
            journal_item.setToolTip(journal_name)
            journal_item.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)

            done_item = QTableWidgetItem(str(done_count))
            total_item = QTableWidgetItem(str(total_count))
            percent_item = QTableWidgetItem(f"{percent}%")

            done_item.setTextAlignment(Qt.AlignCenter)
            total_item.setTextAlignment(Qt.AlignCenter)
            percent_item.setTextAlignment(Qt.AlignCenter)

            if percent >= 80:
                percent_item.setBackground(QColor("#d4edda"))
            elif percent >= 50:
                percent_item.setBackground(QColor("#fff3cd"))
            else:
                percent_item.setBackground(QColor("#f8d7da"))

            self.export_progress_table.setItem(row, 0, journal_item)
            self.export_progress_table.setItem(row, 1, done_item)
            self.export_progress_table.setItem(row, 2, total_item)
            self.export_progress_table.setItem(row, 3, percent_item)

        self.export_progress_table.blockSignals(False)
        self.export_progress_table.setUpdatesEnabled(True)

    def get_export_issue_lock_info(self, issue_id):
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("SELECT status, taken_by FROM export_issues WHERE id=?", (issue_id,))
        row = cur.fetchone()
        conn.close()
        return row if row else (None, None)

    def find_export_issue_row(self, issue_id):
        for row in range(self.export_table.rowCount()):
            item = self.export_table.item(row, 0)
            if item and item.data(Qt.UserRole) == issue_id:
                return row
        return -1

    def issue_matches_export_current_view(self, journal, year, issue, status, taken_by, taken_at, completed_at):
        filter_value = self.export_filter_box.currentText()
        journal_filter = self.export_journal_filter_box.currentText()
        search_text = self.export_search_box.text().strip().casefold()

        if filter_value == "Свободные" and status != "free":
            return False
        if filter_value == "В работе" and status != "in_progress":
            return False
        if filter_value == "Готово" and status != "done":
            return False
        if filter_value == "Мои" and taken_by != CURRENT_USER:
            return False

        if journal_filter != "Все журналы" and (journal or "").strip().casefold() != journal_filter.strip().casefold():
            return False

        status_text = (
            "Свободен" if status == "free" else
            "В работе" if status == "in_progress" else
            "Готово" if status == "done" else str(status)
        )

        row_search = " ".join([
            str(journal or ""),
            str(year or ""),
            str(issue or ""),
            status_text,
            str(taken_by or ""),
            str(taken_at or ""),
            str(completed_at or "")
        ]).casefold()

        if search_text and search_text not in row_search:
            return False

        return True

    def get_export_issue_row_data(self, issue_id):
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("""
            SELECT id, journal, year, issue, status, taken_by, taken_at, completed_at
            FROM export_issues
            WHERE id=?
        """, (issue_id,))
        row = cur.fetchone()
        conn.close()
        return row

    def refresh_export_issue_row(self, issue_id):
        row_data = self.get_export_issue_row_data(issue_id)
        if not row_data:
            row = self.find_export_issue_row(issue_id)
            if row >= 0:
                self.export_table.removeRow(row)
            return

        id_, journal, year, issue, status, taken_by, taken_at, completed_at = row_data

        row = self.find_export_issue_row(issue_id)

        if not self.issue_matches_export_current_view(journal, year, issue, status, taken_by, taken_at, completed_at):
            if row >= 0:
                self.export_table.removeRow(row)
            self.update_export_progress_table()
            return

        if row < 0:
            self.load_export_data()
            return

        self.export_table.blockSignals(True)
        try:
            journal_item = self.export_table.item(row, 0) or QTableWidgetItem()
            journal_item.setText(journal or "")
            journal_item.setData(Qt.UserRole, id_)
            if self.is_admin:
                journal_item.setFlags(journal_item.flags() | Qt.ItemIsEditable)
            self.export_table.setItem(row, 0, journal_item)

            year_item = self.export_table.item(row, 1) or QTableWidgetItem()
            year_item.setText(year or "")
            year_item.setData(Qt.UserRole, id_)
            if self.is_admin:
                year_item.setFlags(year_item.flags() | Qt.ItemIsEditable)
            self.export_table.setItem(row, 1, year_item)

            issue_item = self.export_table.item(row, 2) or QTableWidgetItem()
            issue_item.setText(issue or "")
            issue_item.setData(Qt.UserRole, id_)
            if self.is_admin:
                issue_item.setFlags(issue_item.flags() | Qt.ItemIsEditable)
            self.export_table.setItem(row, 2, issue_item)

            who_item = self.export_table.item(row, 4) or QTableWidgetItem()
            who_item.setText(taken_by or "")
            who_item.setData(Qt.UserRole, id_)
            if self.is_admin:
                who_item.setFlags(who_item.flags() | Qt.ItemIsEditable)
            self.export_table.setItem(row, 4, who_item)

            if self.is_admin:
                status_combo = QComboBox()
                status_combo.addItem("⚪ Свободен", "free")
                status_combo.addItem("🟡 В работе", "in_progress")
                status_combo.addItem("✅ Готово", "done")
                idx = status_combo.findData(status)
                status_combo.setCurrentIndex(idx if idx >= 0 else 0)
                status_combo.currentIndexChanged.connect(
                    lambda _, issue_id=id_, combo=status_combo: self.change_export_status(issue_id, combo.currentData())
                )
                self.export_table.setCellWidget(row, 3, status_combo)
            else:
                status_item = QTableWidgetItem()
                if status == "free":
                    status_item.setText("⚪ Свободен")
                elif status == "in_progress":
                    status_item.setText("🟡 В работе")
                    status_item.setBackground(QColor("#fff3cd"))
                elif status == "done":
                    status_item.setText("✅ Готово")
                    status_item.setBackground(QColor("#d4edda"))
                self.export_table.setItem(row, 3, status_item)

            btn = QPushButton()
            if status == "free":
                btn.setText("Взять")
                btn.clicked.connect(lambda _, i=id_: self.take_export_issue(i))
            elif status == "in_progress" and taken_by == CURRENT_USER:
                btn.setText("Готово")
                btn.clicked.connect(lambda _, i=id_: self.complete_export_issue(i))
            else:
                btn.setText("—")
                btn.setEnabled(False)

            self.export_table.setCellWidget(row, 5, btn)
        finally:
            self.export_table.blockSignals(False)

        self.update_export_progress_table()

    def load_export_data(self):
        scroll_value = self.export_table.verticalScrollBar().value() if self._preserve_export_scroll_on_reload else 0

        self.load_export_journal_filter_options()

        conn = get_db_connection()
        cur = conn.cursor()

        filter_value = self.export_filter_box.currentText()
        journal_filter = self.export_journal_filter_box.currentText()
        search_text = self.export_search_box.text().strip().casefold()

        query = """
            SELECT id, journal, year, issue, status, taken_by, taken_at, completed_at
            FROM export_issues
            WHERE 1=1
        """
        params = []

        if filter_value == "Свободные":
            query += " AND status = 'free'"
        elif filter_value == "В работе":
            query += " AND status = 'in_progress'"
        elif filter_value == "Готово":
            query += " AND status = 'done'"
        elif filter_value == "Мои":
            query += " AND taken_by = ?"
            params.append(CURRENT_USER)

        query += " ORDER BY journal COLLATE NOCASE, year, issue"

        cur.execute(query, params)
        rows = cur.fetchall()
        conn.close()

        self.export_table.setRowCount(0)
        self.export_table.setUpdatesEnabled(False)
        self.export_table.blockSignals(True)

        for row_data in rows:
            id_, journal, year, issue, status, taken_by, taken_at, completed_at = row_data

            status_text = (
                "Свободен" if status == "free" else
                "В работе" if status == "in_progress" else
                "Готово" if status == "done" else str(status)
            )

            row_search = " ".join([
                str(journal or ""),
                str(year or ""),
                str(issue or ""),
                status_text,
                str(taken_by or ""),
                str(taken_at or ""),
                str(completed_at or "")
            ]).casefold()

            if search_text and search_text not in row_search:
                continue

            if journal_filter != "Все журналы" and (journal or "").strip().casefold() != journal_filter.strip().casefold():
                continue

            row = self.export_table.rowCount()
            self.export_table.insertRow(row)

            journal_item = QTableWidgetItem(journal or "")
            journal_item.setData(Qt.UserRole, id_)

            year_item = QTableWidgetItem(year or "")
            year_item.setData(Qt.UserRole, id_)

            issue_item = QTableWidgetItem(issue or "")
            issue_item.setData(Qt.UserRole, id_)

            who_item = QTableWidgetItem(taken_by or "")
            who_item.setData(Qt.UserRole, id_)

            if self.is_admin:
                journal_item.setFlags(journal_item.flags() | Qt.ItemIsEditable)
                year_item.setFlags(year_item.flags() | Qt.ItemIsEditable)
                issue_item.setFlags(issue_item.flags() | Qt.ItemIsEditable)
                who_item.setFlags(who_item.flags() | Qt.ItemIsEditable)

            self.export_table.setItem(row, 0, journal_item)
            self.export_table.setItem(row, 1, year_item)
            self.export_table.setItem(row, 2, issue_item)

            if self.is_admin:
                status_combo = QComboBox()
                status_combo.addItem("⚪ Свободен", "free")
                status_combo.addItem("🟡 В работе", "in_progress")
                status_combo.addItem("✅ Готово", "done")
                idx = status_combo.findData(status)
                status_combo.setCurrentIndex(idx if idx >= 0 else 0)
                status_combo.currentIndexChanged.connect(
                    lambda _, issue_id=id_, combo=status_combo: self.change_export_status(issue_id, combo.currentData())
                )
                self.export_table.setCellWidget(row, 3, status_combo)
            else:
                status_item = QTableWidgetItem(status_text)
                if status == "in_progress":
                    status_item.setBackground(QColor("#fff3cd"))
                elif status == "done":
                    status_item.setBackground(QColor("#d4edda"))
                self.export_table.setItem(row, 3, status_item)

            self.export_table.setItem(row, 4, who_item)

            btn = QPushButton()
            if status == "free":
                btn.setText("Взять")
                btn.clicked.connect(lambda _, i=id_: self.take_export_issue(i))
            elif status == "in_progress" and taken_by == CURRENT_USER:
                btn.setText("Готово")
                btn.clicked.connect(lambda _, i=id_: self.complete_export_issue(i))
            else:
                btn.setText("—")
                btn.setEnabled(False)

            self.export_table.setCellWidget(row, 5, btn)

        self.export_table.blockSignals(False)
        self.export_table.setUpdatesEnabled(True)

        if self._preserve_export_scroll_on_reload:
            self.export_table.verticalScrollBar().setValue(scroll_value)

        self.update_export_progress_table()

    def delete_selected_export_issue(self):
        if not self.is_admin:
            return

        if not self.export_table.selectionModel():
            QMessageBox.warning(
                self,
                "Внимание",
                "Сначала выберите записи в таблице."
            )
            return

        selected_rows = sorted(
            {index.row() for index in self.export_table.selectionModel().selectedRows()}
        )

        if not selected_rows:
            QMessageBox.warning(
                self,
                "Внимание",
                "Сначала выберите записи в таблице."
            )
            return

        rows_to_delete = []

        for row in selected_rows:
            journal_item = self.export_table.item(row, 0)
            year_item = self.export_table.item(row, 1)
            issue_item = self.export_table.item(row, 2)

            if not journal_item:
                continue

            issue_id = journal_item.data(Qt.UserRole)
            if issue_id is None:
                continue

            journal = journal_item.text().strip()
            year = year_item.text().strip() if year_item else ""
            issue = issue_item.text().strip() if issue_item else ""

            rows_to_delete.append((issue_id, journal, year, issue))

        if not rows_to_delete:
            QMessageBox.warning(
                self,
                "Внимание",
                "Не удалось определить выбранные записи."
            )
            return

        preview_lines = [
            f"• {journal} | {year} | {issue}"
            for _, journal, year, issue in rows_to_delete[:5]
        ]

        preview_text = "\n".join(preview_lines)
        if len(rows_to_delete) > 5:
            preview_text += f"\n... и ещё {len(rows_to_delete) - 5}"

        reply = QMessageBox.question(
            self,
            "Удаление записей",
            f"Удалить выбранные записи?\n\n"
            f"Количество: {len(rows_to_delete)}\n\n"
            f"{preview_text}",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply != QMessageBox.Yes:
            return

        conn = get_db_connection()
        cur = conn.cursor()
        for issue_id, _, _, _ in rows_to_delete:
            cur.execute("DELETE FROM export_issues WHERE id=?", (issue_id,))
        conn.commit()
        conn.close()

        self.load_export_data()
        self.remember_db_signature()

    def on_export_table_item_changed(self, item):
        if not self.is_admin:
            return

        issue_id = item.data(Qt.UserRole)
        if issue_id is None:
            return

        field_map = {
            0: "journal",
            1: "year",
            2: "issue",
            4: "taken_by",
        }

        field = field_map.get(item.column())
        if field is None:
            return

        new_value = item.text().strip()

        conn = get_db_connection()
        cur = conn.cursor()

        try:
            cur.execute(
                f"UPDATE export_issues SET {field}=?, updated_at=? WHERE id=?",
                (new_value, datetime.now().isoformat(), issue_id)
            )
            conn.commit()
        except sqlite3.IntegrityError:
            conn.rollback()
            conn.close()
            QMessageBox.warning(
                self,
                "Ошибка",
                "Такая запись уже существует."
            )
            self.refresh_export_issue_row(issue_id)
            return

        conn.close()
        self.refresh_export_issue_row(issue_id)
        self.remember_db_signature()

    def change_export_status(self, issue_id, status):
        conn = get_db_connection()
        cur = conn.cursor()

        current_status, taken_by = self.get_export_issue_lock_info(issue_id)
        if current_status == "in_progress" and taken_by and taken_by != CURRENT_USER and not self.is_admin:
            self.warn_issue_locked(taken_by)
            conn.close()
            return

        cur.execute("SELECT taken_by, taken_at FROM export_issues WHERE id=?", (issue_id,))
        row = cur.fetchone()
        taken_by, taken_at = row if row else (None, None)

        if status == "free":
            cur.execute("""
                UPDATE export_issues
                SET status='free',
                    taken_by=NULL,
                    taken_at=NULL,
                    completed_at=NULL,
                    updated_at=?
                WHERE id=?
            """, (datetime.now().isoformat(), issue_id))

        elif status == "in_progress":
            if not taken_by:
                taken_by = CURRENT_USER
            if not taken_at:
                taken_at = datetime.now().isoformat()

            cur.execute("""
                UPDATE export_issues
                SET status='in_progress',
                    taken_by=?,
                    taken_at=?,
                    completed_at=NULL,
                    updated_at=?
                WHERE id=?
            """, (taken_by, taken_at, datetime.now().isoformat(), issue_id))

        elif status == "done":
            if not taken_by:
                taken_by = CURRENT_USER

            cur.execute("""
                UPDATE export_issues
                SET status='done',
                    taken_by=?,
                    completed_at=?,
                    updated_at=?
                WHERE id=?
            """, (taken_by, datetime.now().isoformat(), datetime.now().isoformat(), issue_id))

        else:
            conn.close()
            return

        conn.commit()
        conn.close()
        self.refresh_export_issue_row(issue_id)
        self.remember_db_signature()

    def take_export_issue(self, issue_id):
        conn = get_db_connection()
        cur = conn.cursor()

        current_status, taken_by = self.get_export_issue_lock_info(issue_id)
        if current_status == "in_progress" and taken_by and taken_by != CURRENT_USER and not self.is_admin:
            self.warn_issue_locked(taken_by)
            conn.close()
            return

        current_status, taken_by = self.get_export_issue_lock_info(issue_id)
        if current_status != "free":
            if current_status == "in_progress" and taken_by and taken_by != CURRENT_USER and not self.is_admin:
                self.warn_issue_locked(taken_by)
            else:
                QMessageBox.warning(
                    self,
                    "Внимание",
                    "Эта запись уже в работе."
                )
            conn.close()
            return

        cur.execute("""
            SELECT COUNT(*) FROM export_issues
            WHERE taken_by=? AND status='in_progress'
        """, (CURRENT_USER,))
        count = cur.fetchone()[0]

        if count > 0:
            QMessageBox.warning(
                self,
                "Внимание",
                "Вы уже взяли одну запись в работу.\n"
                "Сначала завершите её."
            )
            conn.close()
            return

        cur.execute("""
            UPDATE export_issues
            SET status='in_progress',
                taken_by=?,
                taken_at=?,
                updated_at=?
            WHERE id=? AND status='free'
        """, (CURRENT_USER, datetime.now().isoformat(), datetime.now().isoformat(), issue_id))

        conn.commit()
        conn.close()

        self.refresh_export_issue_row(issue_id)
        self.remember_db_signature()

    def complete_export_issue(self, issue_id):
        conn = get_db_connection()
        cur = conn.cursor()

        cur.execute("""
            UPDATE export_issues
            SET status='done',
                completed_at=?,
                updated_at=?
            WHERE id=?
        """, (datetime.now().isoformat(), datetime.now().isoformat(), issue_id))

        conn.commit()
        conn.close()

        self.refresh_export_issue_row(issue_id)
        self.remember_db_signature()

    def import_export_from_excel(self):
        if not self.is_admin:
            return

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите Excel-файл",
            "",
            "Excel (*.xlsx *.xls)"
        )
        if not file_path:
            return

        try:
            df = pd.read_excel(file_path)
        except Exception as exc:
            QMessageBox.critical(
                self,
                "Ошибка импорта",
                f"Не удалось прочитать Excel-файл.\n\n{exc}"
            )
            return

        if df.empty:
            QMessageBox.warning(
                self,
                "Импорт",
                "Файл Excel пуст."
            )
            return

        def normalize_name(value):
            return re.sub(r"[^a-zа-я0-9]+", "", str(value).strip().casefold())

        column_map = {}
        for column in df.columns:
            key = normalize_name(column)
            if key in {"journal", "журнал", "названиежурнала", "наименованижурнала", "наименование"}:
                column_map[column] = "journal"
            elif key in {"year", "год"}:
                column_map[column] = "year"
            elif key in {"issue", "выпуск", "номер", "номервыпуска"}:
                column_map[column] = "issue"
            elif key in {"status", "статус"}:
                column_map[column] = "status"
            elif key in {"takenby", "кто", "пользователь"}:
                column_map[column] = "taken_by"
            elif key in {"takenat", "датавзятия", "взято"}:
                column_map[column] = "taken_at"
            elif key in {"completedat", "датаокончания", "готово", "выгружено"}:
                column_map[column] = "completed_at"

        df = df.rename(columns=column_map)

        required_columns = {"journal", "year", "issue"}
        if not required_columns.issubset(set(df.columns)):
            QMessageBox.warning(
                self,
                "Импорт",
                "В Excel должны быть обязательные колонки:\n"
                "Журнал, Год, Выпуск"
            )
            return

        def cell_to_text(value):
            if pd.isna(value):
                return ""
            if isinstance(value, datetime):
                return value.isoformat(sep=" ", timespec="seconds")
            if isinstance(value, float) and value.is_integer():
                return str(int(value))
            return str(value).strip()

        def normalize_status(value):
            key = normalize_name(value)
            status_map = {
                "free": "free",
                "свободен": "free",
                "свободно": "free",
                "inprogress": "in_progress",
                "вработе": "in_progress",
                "done": "done",
                "готово": "done",
                "выгружен": "done",
            }
            return status_map.get(key, "free")

        inserted = 0
        conn = get_db_connection()
        cur = conn.cursor()

        for _, row in df.iterrows():
            journal = cell_to_text(row.get("journal", ""))
            year_raw = row.get("year", "")
            issue = cell_to_text(row.get("issue", ""))

            if isinstance(year_raw, datetime):
                year = year_raw.strftime("%Y")
            else:
                year = cell_to_text(year_raw)

            if not journal or not year or not issue:
                continue

            status = normalize_status(row.get("status", "free"))
            taken_by = cell_to_text(row.get("taken_by", ""))
            taken_at = cell_to_text(row.get("taken_at", ""))
            completed_at = cell_to_text(row.get("completed_at", ""))

            cur.execute("""
                INSERT OR IGNORE INTO export_issues
                    (journal, year, issue, status, taken_by, taken_at, completed_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                journal,
                year,
                issue,
                status,
                taken_by or None,
                taken_at or None,
                completed_at or None,
                datetime.now().isoformat()
            ))

            if cur.rowcount > 0:
                inserted += 1

        conn.commit()
        conn.close()

        self.load_export_data()
        self.remember_db_signature()

        QMessageBox.information(
            self,
            "Импорт завершён",
            f"Добавлено новых записей: {inserted}"
        )

    def load_data(self):
        scroll_value = self.table.verticalScrollBar().value() if self._preserve_scroll_on_reload else 0
        conn = get_db_connection()
        cur = conn.cursor()

        filter_value = self.filter_box.currentText()
        journal_filter = self.journal_filter_box.currentText()
        search_text = self.search_box.text().strip().casefold()

        query = """
            SELECT id, journal, year, issue, path, status, taken_by, taken_at, completed_at
            FROM issues
            WHERE 1=1
        """
        params = []

        if filter_value == "Свободные":
            query += " AND status = 'free'"
        elif filter_value == "В работе":
            query += " AND status = 'in_progress'"
        elif filter_value == "Готово":
            query += " AND status = 'done'"
        elif filter_value == "Мои":
            query += " AND taken_by = ?"
            params.append(CURRENT_USER)

        query += " ORDER BY journal COLLATE NOCASE, year, issue"

        cur.execute(query, params)
        rows = cur.fetchall()
        conn.close()

        self.table.setRowCount(0)
        self.table.setUpdatesEnabled(False)
        self.table.blockSignals(True)

        for row_data in rows:
            id_, journal, year, issue, path, status, taken_by, taken_at, completed_at = row_data

            row_search = " ".join([
                str(journal or ""),
                str(year or ""),
                str(issue or ""),
                str(path or ""),
                "Свободен" if status == "free" else
                "В работе" if status == "in_progress" else
                "Готово" if status == "done" else str(status),
                str(taken_by or ""),
                str(taken_at or ""),
                str(completed_at or "")
            ]).casefold()

            if search_text and search_text not in row_search:
                continue

            if journal_filter != "Все журналы" and (journal or "").strip().casefold() != journal_filter.strip().casefold():
                continue

            row = self.table.rowCount()
            self.table.insertRow(row)

            journal_item = QTableWidgetItem(journal)
            journal_item.setData(Qt.UserRole, id_)

            year_item = QTableWidgetItem(year)
            year_item.setData(Qt.UserRole, id_)

            issue_item = QTableWidgetItem(issue)
            issue_item.setData(Qt.UserRole, id_)

            who_item = QTableWidgetItem(taken_by or "")
            who_item.setData(Qt.UserRole, id_)

            if self.is_admin:
                journal_item.setFlags(journal_item.flags() | Qt.ItemIsEditable)
                year_item.setFlags(year_item.flags() | Qt.ItemIsEditable)
                issue_item.setFlags(issue_item.flags() | Qt.ItemIsEditable)
                who_item.setFlags(who_item.flags() | Qt.ItemIsEditable)

            self.table.setItem(row, 0, journal_item)
            self.table.setItem(row, 1, year_item)
            self.table.setItem(row, 2, issue_item)

            if self.is_admin:
                status_combo = QComboBox()
                status_combo.addItem("⚪ Свободен", "free")
                status_combo.addItem("🟡 В работе", "in_progress")
                status_combo.addItem("✅ Готово", "done")
                status_combo.setCurrentIndex(status_combo.findData(status))
                status_combo.currentIndexChanged.connect(
                    lambda _, issue_id=id_, combo=status_combo: self.change_status(issue_id, combo.currentData())
                )
                self.table.setCellWidget(row, 3, status_combo)
            else:
                status_item = QTableWidgetItem()
                if status == "free":
                    status_item.setText("⚪ Свободен")
                elif status == "in_progress":
                    status_item.setText("🟡 В работе")
                    status_item.setBackground(QColor("#fff3cd"))
                elif status == "done":
                    status_item.setText("✅ Готово")
                    status_item.setBackground(QColor("#d4edda"))
                self.table.setItem(row, 3, status_item)

            self.table.setItem(row, 4, who_item)

            btn = QPushButton()
            if status == "free":
                btn.setText("Взять")
                btn.clicked.connect(lambda _, i=id_: self.take_issue(i))
            elif status == "in_progress" and taken_by == CURRENT_USER:
                btn.setText("Готово")
                btn.clicked.connect(lambda _, i=id_: self.complete_issue(i))
            else:
                btn.setText("—")
                btn.setEnabled(False)

            self.table.setCellWidget(row, 5, btn)

        self.table.blockSignals(False)
        self.table.setUpdatesEnabled(True)

        if self._preserve_scroll_on_reload:
            self.table.verticalScrollBar().setValue(scroll_value)

    def on_table_item_changed(self, item):
        if not self.is_admin:
            return

        issue_id = item.data(Qt.UserRole)
        if issue_id is None:
            return

        field_map = {
            0: "journal",
            1: "year",
            2: "issue",
            4: "taken_by",
        }

        field = field_map.get(item.column())
        if field is None:
            return

        new_value = item.text().strip()

        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute(
            f"UPDATE issues SET {field}=?, updated_at=? WHERE id=?",
            (new_value, datetime.now().isoformat(), issue_id)
        )
        conn.commit()
        conn.close()
        self.remember_db_signature()

    def change_status(self, issue_id, status):
        conn = get_db_connection()
        cur = conn.cursor()

        current_status, taken_by = self.get_issue_lock_info(issue_id)
        if current_status == "in_progress" and taken_by and taken_by != CURRENT_USER and not self.is_admin:
            self.warn_issue_locked(taken_by)
            conn.close()
            return

        cur.execute("SELECT taken_by, taken_at FROM issues WHERE id=?", (issue_id,))
        row = cur.fetchone()
        taken_by, taken_at = row if row else (None, None)

        if status == "free":
            cur.execute("""
                UPDATE issues
                SET status='free',
                    taken_by=NULL,
                    taken_at=NULL,
                    completed_at=NULL,
                    updated_at=?
                WHERE id=?
            """, (datetime.now().isoformat(), issue_id))

        elif status == "in_progress":
            if not taken_by:
                taken_by = CURRENT_USER
            if not taken_at:
                taken_at = datetime.now().isoformat()

            cur.execute("""
                UPDATE issues
                SET status='in_progress',
                    taken_by=?,
                    taken_at=?,
                    completed_at=NULL,
                    updated_at=?
                WHERE id=?
            """, (taken_by, taken_at, datetime.now().isoformat(), issue_id))

        elif status == "done":
            if not taken_by:
                taken_by = CURRENT_USER

            cur.execute("""
                UPDATE issues
                SET status='done',
                    taken_by=?,
                    completed_at=?,
                    updated_at=?
                WHERE id=?
            """, (taken_by, datetime.now().isoformat(), datetime.now().isoformat(), issue_id))

        else:
            conn.close()
            return

        conn.commit()
        conn.close()
        self.refresh_issue_row(issue_id)
        self.remember_db_signature()

    def take_issue(self, issue_id):
        conn = get_db_connection()
        cur = conn.cursor()

        current_status, taken_by = self.get_issue_lock_info(issue_id)
        if current_status == "in_progress" and taken_by and taken_by != CURRENT_USER and not self.is_admin:
            self.warn_issue_locked(taken_by)
            conn.close()
            return

        current_status, taken_by = self.get_issue_lock_info(issue_id)
        if current_status != "free":
            if current_status == "in_progress" and taken_by and taken_by != CURRENT_USER and not self.is_admin:
                self.warn_issue_locked(taken_by)
            else:
                QMessageBox.warning(
                    self,
                    "Внимание",
                    "Этот журнал уже в работе."
                )
            conn.close()
            return

        # 🔎 Проверяем, есть ли уже журнал в работе
        cur.execute("""
            SELECT COUNT(*) FROM issues
            WHERE taken_by=? AND status='in_progress'
        """, (CURRENT_USER,))
        count = cur.fetchone()[0]

        if count > 0:
            QMessageBox.warning(
                self,
                "Внимание",
                "Вы уже взяли один журнал в работу.\n"
                "Сначала завершите его."
            )
            conn.close()
            return

        # ✅ Если нет — берем в работу
        cur.execute("""
            UPDATE issues
            SET status='in_progress',
                taken_by=?,
                taken_at=?,
                updated_at=?
            WHERE id=? AND status='free'
        """, (CURRENT_USER, datetime.now().isoformat(), datetime.now().isoformat(), issue_id))

        conn.commit()
        conn.close()

        self.refresh_issue_row(issue_id)
        self.remember_db_signature()

    def complete_issue(self, issue_id):
        conn = get_db_connection()
        cur = conn.cursor()

        cur.execute("""
        UPDATE issues
        SET status='done',
            completed_at=?,
            updated_at=?
        WHERE id=?
        """, (datetime.now().isoformat(), datetime.now().isoformat(), issue_id))

        conn.commit()
        conn.close()

        self.refresh_issue_row(issue_id)
        self.remember_db_signature()

    def export_report(self):
        conn = get_db_connection()
        df = pd.read_sql_query("SELECT * FROM issues WHERE status='done' AND LOWER(path) LIKE '%.pdf'", conn)
        conn.close()

        file_path, _ = QFileDialog.getSaveFileName(self, "Сохранить отчет", "", "Excel (*.xlsx)")
        if file_path:
            df.to_excel(file_path, index=False)

    def get_selected_month(self):
        if self.is_admin and hasattr(self, "plan_month_edit"):
            month = self.plan_month_edit.text().strip()
            try:
                datetime.strptime(month, "%Y-%m")
                return month
            except ValueError:
                return datetime.now().strftime("%Y-%m")

        return datetime.now().strftime("%Y-%m")

    def get_previous_month(self, month_str):
        dt = datetime.strptime(month_str, "%Y-%m")
        if dt.month == 1:
            return f"{dt.year - 1}-12"
        return f"{dt.year}-{dt.month - 1:02d}"

    def get_best_user_for_month(self, month_str):
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("""
            SELECT taken_by, COUNT(*) as cnt
            FROM issues
            WHERE status='done'
              AND completed_at LIKE ?
              AND taken_by IS NOT NULL
              AND TRIM(taken_by) != ''
            GROUP BY taken_by
            ORDER BY cnt DESC, taken_by COLLATE NOCASE
            LIMIT 1
        """, (month_str + "%",))
        row = cur.fetchone()
        conn.close()
        return row

    def save_monthly_plan(self):
        if not self.is_admin:
            return

        month = self.plan_month_edit.text().strip()
        try:
            datetime.strptime(month, "%Y-%m")
        except ValueError:
            QMessageBox.warning(
                self,
                "Ошибка",
                "Введите месяц в формате YYYY-MM, например 2026-05."
            )
            return

        plan_count = self.plan_count_spin.value()

        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("""
            INSERT INTO monthly_plans (month, plan_count, created_by, updated_at)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(month) DO UPDATE SET
                plan_count=excluded.plan_count,
                created_by=excluded.created_by,
                updated_at=excluded.updated_at
        """, (month, plan_count, CURRENT_USER, datetime.now().isoformat()))
        conn.commit()
        conn.close()

        self.remember_db_signature()

    # ================= РЕЙТИНГ =================

    def setup_rating_tab(self):
        layout = QVBoxLayout()

        self.plan_frame = QFrame()
        self.plan_frame.setObjectName("ratingCard")
        plan_layout = QVBoxLayout(self.plan_frame)

        plan_title = QLabel("План на месяц")
        plan_title.setStyleSheet("font-weight: 700; font-size: 14px; color: #1f2937;")

        self.plan_summary_label = QLabel()
        self.plan_summary_label.setWordWrap(True)
        self.plan_summary_label.setStyleSheet("font-size: 14px;")

        plan_layout.addWidget(plan_title)
        plan_layout.addWidget(self.plan_summary_label)

        plan_controls = QHBoxLayout()
        plan_controls.addStretch()

        if self.is_admin:
            self.plan_month_edit = QLineEdit()
            self.plan_month_edit.setPlaceholderText("YYYY-MM")
            self.plan_month_edit.setMaximumWidth(90)
            self.plan_month_edit.editingFinished.connect(self.update_rating)

            self.plan_count_spin = QSpinBox()
            self.plan_count_spin.setRange(0, 100000)
            self.plan_count_spin.setMaximumWidth(110)

            save_plan_btn = QPushButton("Сохранить план")
            save_plan_btn.clicked.connect(self.save_monthly_plan)

            plan_controls.addWidget(QLabel("Месяц:"))
            plan_controls.addWidget(self.plan_month_edit)
            plan_controls.addWidget(QLabel("План:"))
            plan_controls.addWidget(self.plan_count_spin)
            plan_controls.addWidget(save_plan_btn)

        plan_layout.addLayout(plan_controls)
        layout.addWidget(self.plan_frame)

        self.info_frame = QFrame()
        self.info_frame.setObjectName("ratingCard")
        info_layout = QVBoxLayout(self.info_frame)

        info_title = QLabel("Всего записей")
        info_title.setStyleSheet("font-weight: 700; font-size: 14px; color: #1f2937;")

        self.rating_info_label = QLabel()
        self.rating_info_label.setWordWrap(True)
        self.rating_info_label.setStyleSheet("font-size: 14px;")

        info_layout.addWidget(info_title)
        info_layout.addWidget(self.rating_info_label)

        self.journal_frame = QFrame()
        self.journal_frame.setObjectName("ratingCard")
        journal_layout = QVBoxLayout(self.journal_frame)

        journal_title = QLabel("Прогресс по журналам")
        journal_title.setStyleSheet("font-weight: 700; font-size: 14px; color: #1f2937;")

        self.journal_progress_table = QTableWidget()
        self.journal_progress_summary_label = QLabel()
        self.journal_progress_summary_label.setWordWrap(False)
        self.journal_progress_summary_label.setStyleSheet("font-size: 14px; color: #1f2937;")
        self.journal_progress_table.setColumnCount(4)
        self.journal_progress_table.setHorizontalHeaderLabels(
            ["Журнал", "Готово", "Всего", "%"]
        )
        self.journal_progress_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.journal_progress_table.setAlternatingRowColors(True)
        self.journal_progress_table.horizontalHeader().setStretchLastSection(False)
        self.journal_progress_table.setColumnWidth(0, 360)
        self.journal_progress_table.setColumnWidth(1, 80)
        self.journal_progress_table.setColumnWidth(2, 80)
        self.journal_progress_table.setColumnWidth(3, 80)

        self.journal_progress_table.setStyleSheet("""
            QTableWidget {
                background-color: #ffffff;
                alternate-background-color: #f7f7f7;
            }
            QHeaderView::section {
                background-color: #e9eef5;
                color: #1f2937;
                padding: 6px 8px;
                border: 1px solid #d7dde5;
                font-weight: bold;
            }
            QTableWidget::item:selected {
                background-color: #0078d7;
                color: white;
            }
            QTableWidget::item:selected:active {
                background-color: #0078d7;
                color: white;
            }
        """)

        journal_layout.addWidget(journal_title)
        journal_layout.addWidget(self.journal_progress_summary_label)
        journal_layout.addWidget(self.journal_progress_table)

        self.rating_board_frame = QFrame()
        self.rating_board_frame.setObjectName("ratingCard")
        board_layout = QVBoxLayout(self.rating_board_frame)
        board_layout.setContentsMargins(0, 0, 0, 0)
        board_layout.setSpacing(6)

        board_header = QHBoxLayout()
        board_title = QLabel("Рейтинг по месяцам")
        board_title.setStyleSheet("font-weight: 700; font-size: 14px; color: #1f2937;")
        board_header.addWidget(board_title)
        board_header.addStretch()

        if self.is_admin:
            save_months_btn = QPushButton("Сохранить месяцы")
            save_months_btn.clicked.connect(self.save_rating_months)
            board_header.addWidget(save_months_btn)

        board_layout.addLayout(board_header)

        self.rating_month_edits = []
        self.rating_month_title_labels = []
        self.rating_month_rank_labels = []

        columns_layout = QHBoxLayout()
        columns_layout.setContentsMargins(0, 0, 0, 0)
        columns_layout.setSpacing(10)

        saved_months = self.get_rating_months()

        for slot in range(RATING_MONTH_SLOTS):
            column_frame = QFrame()
            column_frame.setObjectName("ratingMonthCard")
            column_layout = QVBoxLayout(column_frame)

            title_label = QLabel()
            title_label.setAlignment(Qt.AlignCenter)
            title_label.setFixedHeight(28)
            title_label.setStyleSheet("""
                QLabel {
                    background-color: #e9eef5;
                    color: #1f2937;
                    padding: 3px 8px;
                    border: 1px solid #d7dde5;
                    border-radius: 6px;
                    font-weight: bold;
                    font-size: 13px;
                }
            """)
            self.rating_month_title_labels.append(title_label)
            column_layout.addWidget(title_label)

            if self.is_admin:
                edit_row = QHBoxLayout()
                edit_row.addWidget(QLabel("Период:"))
                month_edit = QLineEdit()
                month_edit.setPlaceholderText("YYYY-MM")
                month_edit.setMaximumWidth(90)
                month_edit.setText(saved_months[slot])
                self.rating_month_edits.append(month_edit)
                edit_row.addWidget(month_edit)
                edit_row.addStretch()
                column_layout.addLayout(edit_row)

            rank_label = QLabel()
            rank_label.setWordWrap(True)
            rank_label.setAlignment(Qt.AlignTop | Qt.AlignLeft)
            rank_label.setStyleSheet("font-size: 13px;")
            self.rating_month_rank_labels.append(rank_label)
            column_layout.addWidget(rank_label)

            columns_layout.addWidget(column_frame)

        board_layout.addLayout(columns_layout)

        content_layout = QHBoxLayout()
        content_layout.addWidget(self.journal_frame, 1)
        content_layout.addWidget(self.rating_board_frame, 1)
        content_layout.setAlignment(self.rating_board_frame, Qt.AlignTop)

        layout.addLayout(content_layout)

        self.rating_tab.setLayout(layout)

        self.rating_tab.setStyleSheet("""
            QFrame#ratingCard {
                background-color: #ffffff;
                border: 1px solid #d7dde5;
                border-radius: 10px;
                padding: 6px;
            }
            QFrame#ratingMonthCard {
                background-color: #f9fafb;
                border: 1px solid #d7dde5;
                border-radius: 8px;
                padding: 8px;
            }
        """)

        if self.is_admin:
            self.plan_month_edit.setText(datetime.now().strftime("%Y-%m"))

    def update_rating(self):
        conn = get_db_connection()
        cur = conn.cursor()

        month = self.get_selected_month()

        cur.execute("SELECT plan_count FROM monthly_plans WHERE month=?", (month,))
        row = cur.fetchone()
        plan_count = row[0] if row else 0

        cur.execute("""
            SELECT COUNT(*)
            FROM issues
            WHERE status='done' AND completed_at LIKE ?
        """, (month + "%",))
        month_done = cur.fetchone()[0]

        month_remaining = max(plan_count - month_done, 0)
        month_percent = int((month_done / plan_count) * 100) if plan_count else 0

        cur.execute("SELECT COUNT(*) FROM issues")
        total_issues = cur.fetchone()[0]

        cur.execute("SELECT COUNT(*) FROM issues WHERE status='done'")
        done_issues = cur.fetchone()[0]

        overall_percent = int((done_issues / total_issues) * 100) if total_issues else 0

        cur.execute("""
            SELECT journal,
                   SUM(CASE WHEN status='done' THEN 1 ELSE 0 END) AS done_count,
                   COUNT(*) AS total_count
            FROM issues
            GROUP BY journal
            ORDER BY journal COLLATE NOCASE
        """)
        journal_rows = cur.fetchall()

        rating_months = self.get_rating_months()

        conn.close()

        if self.is_admin and hasattr(self, "plan_count_spin"):
            self.plan_count_spin.blockSignals(True)
            self.plan_count_spin.setValue(plan_count)
            self.plan_count_spin.blockSignals(False)

        if self.is_admin and hasattr(self, "rating_month_edits"):
            for edit, month_value in zip(self.rating_month_edits, rating_months):
                edit.blockSignals(True)
                edit.setText(month_value)
                edit.blockSignals(False)

        self.plan_summary_label.setText(
            f"Месяц: {month} | План: {plan_count} | Факт: {month_done} | "
            f"Осталось: {month_remaining} | Выполнение: {month_percent}%"
        )

        self.rating_info_label.setText(
            f"Всего записей: {total_issues} | Готово: {done_issues} | "
            f"Общий прогресс: {overall_percent}%"
        )
        if hasattr(self, "journal_progress_summary_label"):
            self.journal_progress_summary_label.setText(
                f"Всего: {total_issues} | Готово: {done_issues} | Прогресс: {overall_percent}%"
            )

        self.journal_progress_table.setRowCount(0)
        self.journal_progress_table.setUpdatesEnabled(False)
        self.journal_progress_table.blockSignals(True)

        for journal, done_count, total_count in journal_rows:
            journal_name = journal or "Без названия"
            percent = int((done_count / total_count) * 100) if total_count else 0

            row = self.journal_progress_table.rowCount()
            self.journal_progress_table.insertRow(row)

            journal_item = QTableWidgetItem(journal_name)
            done_item = QTableWidgetItem(str(done_count))
            total_item = QTableWidgetItem(str(total_count))
            percent_item = QTableWidgetItem(f"{percent}%")

            done_item.setTextAlignment(Qt.AlignCenter)
            total_item.setTextAlignment(Qt.AlignCenter)
            percent_item.setTextAlignment(Qt.AlignCenter)

            if percent >= 80:
                percent_item.setBackground(QColor("#d4edda"))
            elif percent >= 50:
                percent_item.setBackground(QColor("#fff3cd"))
            else:
                percent_item.setBackground(QColor("#f8d7da"))

            self.journal_progress_table.setItem(row, 0, journal_item)
            self.journal_progress_table.setItem(row, 1, done_item)
            self.journal_progress_table.setItem(row, 2, total_item)
            self.journal_progress_table.setItem(row, 3, percent_item)

        self.journal_progress_table.blockSignals(False)
        self.journal_progress_table.setUpdatesEnabled(True)

        for index, month_value in enumerate(rating_months):
            if index < len(self.rating_month_title_labels):
                self.rating_month_title_labels[index].setText(
                    self.get_month_display_name(month_value)
                )

            if index < len(self.rating_month_rank_labels):
                ranking_rows = self.get_month_ranking(month_value)
                self.rating_month_rank_labels[index].setText(
                    self.format_month_ranking_text(ranking_rows)
                )


# ================= ЗАПУСК =================

if __name__ == "__main__":
    try:
        restore_db_from_backup_if_needed()
        init_db()
    except sqlite3.OperationalError as exc:
        app = QApplication(sys.argv)
        QMessageBox.critical(
            None,
            "Ошибка запуска",
            f"Не удалось открыть общую базу данных.\n\n"
            f"Путь: {DB_PATH}\n\n"
            f"Проверьте, что у вас есть доступ на запись к сетевой папке.\n\n"
            f"Техническая ошибка:\n{exc}"
        )
        sys.exit(1)

    # Автоскан сети отключен, чтобы не тормозить старт

    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

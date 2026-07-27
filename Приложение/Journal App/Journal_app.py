import sys
import os
import re
import json
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
    QAbstractItemView, QInputDialog, QHeaderView, QListWidget, QDialog
)
from PySide6.QtGui import QColor
from PySide6.QtCore import Qt, QTimer

import pandas as pd


class NoWheelComboBox(QComboBox):
    def wheelEvent(self, event):
        event.ignore()


ISSUES_TEXT_FIT_COLUMNS = (0, 1, 2)
ISSUES_TABLE_COLUMN_COMPLETED_AT = 9

ISSUES_TABLE_FIXED_COLUMN_SAMPLES = {
    3: ["⚪ Свободен", "🟡 В работе", "✅ Готово"],
    4: ["nyagavrilova", "username"],
    5: ["Взять", "Готово", "—"],
    6: ["nyagavrilova", "Кто проверил"],
    7: ["Взять на проверку", "Проверено", "—"],
    8: ["⚪ Свободно", "🟡 На проверке", "✅ Проверено"],
    9: ["30.06.2026", "01.07.2026"],
}

ISSUES_TABLE_COLUMN_EXTRA_PADDING = {
    3: 36,
    5: 24,
    7: 24,
    8: 36,
}

RATING_PROGRESS_COLUMN_SAMPLES = [
    ["Техническая литература", "Электронная библиотека"],
    ["12345", "Всего, шт"],
    ["12345", "Выгружено, шт"],
    ["100%", "Выгружено, %"],
    ["12345", "На проверку, шт"],
    ["12345", "Проверено, шт"],
    ["100%", "Проверено, %"],
]

ISSUES_REPORT_COLUMNS = [
    ("journal", "Журнал"),
    ("year", "Год"),
    ("issue", "Выпуск"),
    ("status", "Статус"),
    ("taken_by", "Кто"),
    ("completed_at", "Дата готовности"),
    ("review_taken_by", "Кто проверил"),
    ("review_status", "Статус проверки"),
]

ISSUE_STATUS_TEXTS = {
    "free": "⚪ Свободен",
    "in_progress": "🟡 В работе",
    "done": "✅ Готово",
}


def set_percent_cell_background(item, percent):
    if percent >= 80:
        item.setBackground(QColor("#d4edda"))
    elif percent >= 50:
        item.setBackground(QColor("#fff3cd"))
    else:
        item.setBackground(QColor("#f8d7da"))


def format_report_datetime(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    if not text:
        return ""
    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00")).strftime("%d.%m.%Y")
    except ValueError:
        return text


def format_completed_at_for_table(value):
    if value is None or not str(value).strip():
        return ""
    text = str(value).strip()
    try:
        dt = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError:
        return text
    if dt.hour or dt.minute or dt.second:
        return dt.strftime("%d.%m.%Y %H:%M")
    return dt.strftime("%d.%m.%Y")


def parse_completed_at_input(text):
    value = text.strip()
    if not value:
        return None

    for fmt in ("%d.%m.%Y %H:%M:%S", "%d.%m.%Y %H:%M", "%d.%m.%Y"):
        try:
            return datetime.strptime(value, fmt).isoformat()
        except ValueError:
            continue

    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, fmt).isoformat()
        except ValueError:
            continue

    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00")).isoformat()
    except ValueError as exc:
        raise ValueError(
            "Введите дату в формате ДД.ММ.ГГГГ или ГГГГ-ММ-ДД, например 30.06.2026."
        ) from exc


def parse_date_filter_input(text):
    """Разбор даты из поля фильтра. Пустое или неполное значение — фильтр не применяется."""
    value = text.strip()
    if not value:
        return None

    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    return None


def parse_completed_at_date(value):
    if value is None or not str(value).strip():
        return None
    text = str(value).strip()

    for fmt in ("%d.%m.%Y %H:%M:%S", "%d.%m.%Y %H:%M", "%d.%m.%Y",
                "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue

    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00")).date()
    except ValueError:
        return None


def completed_at_matches_date_filter(completed_at, filter_date):
    if filter_date is None:
        return True
    record_date = parse_completed_at_date(completed_at)
    return record_date == filter_date if record_date is not None else False


def append_completed_at_date_sql(query, params, filter_date):
    """Фильтр по дню готовности в SQL (prefix по ISO), без обхода всех строк в Python."""
    if filter_date is None:
        return query, params
    query += " AND completed_at LIKE ?"
    params.append(filter_date.isoformat() + "%")
    return query, params


STATUS_SEARCH_TEXTS = {
    "free": ("free", "свободен", "свободно", "⚪ свободен"),
    "in_progress": ("in_progress", "в работе", "🟡 в работе"),
    "done": ("done", "готово", "✅ готово", "выгружен", "выгружено"),
}

REVIEW_STATUS_SEARCH_TEXTS = {
    "free": ("free", "свободно", "проверка свободна", "⚪ свободно"),
    "in_progress": ("in_progress", "на проверке", "🟡 на проверке"),
    "done": ("done", "проверено", "✅ проверено"),
}


def build_issue_search_text(
    journal="",
    year="",
    issue="",
    path="",
    status="",
    taken_by="",
    review_status=None,
    review_taken_by="",
):
    """Текст строки для текстового поиска (без дат — для них отдельный фильтр)."""
    status_key = status or ""
    review_key = review_status or "free"
    parts = [
        str(journal or ""),
        str(year or ""),
        str(issue or ""),
        str(path or ""),
        *STATUS_SEARCH_TEXTS.get(status_key, (str(status_key),)),
        str(taken_by or ""),
        *REVIEW_STATUS_SEARCH_TEXTS.get(review_key, (str(review_key),)),
        str(review_taken_by or ""),
    ]
    return " ".join(parts).casefold()


def issue_row_matches_search(search_text, **fields):
    if not search_text:
        return True
    return search_text in build_issue_search_text(**fields)


def filter_issue_rows(rows, search_text, with_path=False):
    """Текстовый поиск по уже выбранным из БД строкам (даты фильтруются в SQL)."""
    if not search_text:
        return rows

    filtered = []
    for row_data in rows:
        if with_path:
            (
                _id,
                journal,
                year,
                issue,
                path,
                status,
                taken_by,
                _taken_at,
                _completed_at,
                review_status,
                review_taken_by,
                _review_taken_at,
                _review_completed_at,
            ) = row_data
        else:
            (
                _id,
                journal,
                year,
                issue,
                status,
                taken_by,
                _taken_at,
                _completed_at,
                review_status,
                review_taken_by,
                _review_taken_at,
                _review_completed_at,
            ) = row_data
            path = ""

        if issue_row_matches_search(
            search_text,
            journal=journal,
            year=year,
            issue=issue,
            path=path,
            status=status,
            taken_by=taken_by,
            review_status=review_status,
            review_taken_by=review_taken_by,
        ):
            filtered.append(row_data)
    return filtered


def _column_sample_texts(column_samples, column):
    if isinstance(column_samples, dict):
        return column_samples.get(column, [])
    if column_samples and column < len(column_samples):
        return column_samples[column]
    return []


def fit_table_columns_to_content(table, columns, horizontal_padding=28):
    if not columns:
        return

    header = table.horizontalHeader()
    font_metrics = table.fontMetrics()

    for column in columns:
        header_item = table.horizontalHeaderItem(column)
        header_text = header_item.text() if header_item else ""
        max_width = font_metrics.horizontalAdvance(header_text) + horizontal_padding

        for row in range(table.rowCount()):
            item = table.item(row, column)
            if item is None:
                continue
            text = item.text() or ""
            max_width = max(
                max_width,
                font_metrics.horizontalAdvance(text) + horizontal_padding,
            )

        table.setColumnWidth(column, max_width)
        header.setSectionResizeMode(column, QHeaderView.Interactive)


def apply_static_column_widths(
    table,
    column_samples=None,
    stretch_last=False,
    horizontal_padding=28,
    column_extra_padding=None,
    columns_to_configure=None,
):
    header = table.horizontalHeader()
    font_metrics = header.fontMetrics()
    columns = columns_to_configure or range(table.columnCount())

    for column in columns:
        header_item = table.horizontalHeaderItem(column)
        header_text = header_item.text() if header_item else ""
        width = font_metrics.horizontalAdvance(header_text) + horizontal_padding

        for sample_text in _column_sample_texts(column_samples, column):
            width = max(
                width,
                font_metrics.horizontalAdvance(sample_text) + horizontal_padding,
            )

        if column_extra_padding and column in column_extra_padding:
            width += column_extra_padding[column]

        table.setColumnWidth(column, width)
        header.setSectionResizeMode(column, QHeaderView.Interactive)

    header.setStretchLastSection(stretch_last)


# ================= НАСТРОЙКИ =================

NETWORK_PATH = Path(r"\\fileserver\УТЗ\Электронная библиотека УТЗ\01_Техническая литература")
AUTO_REFRESH_INTERVAL_MS = 500
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

APP_VERSION = "1.1.2"
VERSIONS_DIR_NAME = "versions"
VERSION_CONFIG_NAME = "version.json"
UPDATE_CHECK_INTERVAL_MS = 5 * 60 * 1000


def resolve_app_paths():
    """Корень развёртывания (БД, version.json) и папка текущего exe."""
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
    else:
        exe_dir = Path(__file__).resolve().parent

    if exe_dir.parent.name.lower() == VERSIONS_DIR_NAME:
        app_root = exe_dir.parent.parent
    else:
        app_root = exe_dir

    return app_root, exe_dir


def parse_version_parts(version_str):
    parts = []
    for part in re.split(r"[.\-]", str(version_str or "").strip()):
        if part.isdigit():
            parts.append(int(part))
        elif part:
            break
    return tuple(parts) if parts else (0,)


def is_version_newer(available_version, current_version):
    return parse_version_parts(available_version) > parse_version_parts(current_version)


def load_deployed_version_config():
    config_path = APP_ROOT / VERSION_CONFIG_NAME
    if not config_path.exists():
        return {}

    try:
        with config_path.open(encoding="utf-8-sig") as config_file:
            data = json.load(config_file)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


APP_ROOT, APP_EXE_DIR = resolve_app_paths()
DB_PATH = str(APP_ROOT / "journal_app.db")
DB_DIR = APP_ROOT

CURRENT_USER = getpass.getuser()

ADMIN_USERS = [
    "Administrator",
    "admin",
    "nyagavrilova",
    "pyagavrilov"
]

# ================= БАЗА ДАННЫХ =================

def init_db():
    # === Очистка старых WAL-файлов (если остались от предыдущего режима) ===
    wal_path = Path(str(DB_PATH) + "-wal")
    shm_path = Path(str(DB_PATH) + "-shm")

    if wal_path.exists() or shm_path.exists():
        # Пытаемся принудительно применить WAL к основной БД
        try:
            temp_conn = sqlite3.connect(DB_PATH, timeout=5)
            temp_conn.execute("PRAGMA journal_mode = DELETE")
            temp_conn.execute("PRAGMA wal_checkpoint(TRUNCATE)")
            temp_conn.commit()
            temp_conn.close()
        except sqlite3.OperationalError:
            # Если БД занята — ждём и пробуем ещё раз
            import time
            time.sleep(2)
            try:
                temp_conn = sqlite3.connect(DB_PATH, timeout=5)
                temp_conn.execute("PRAGMA journal_mode = DELETE")
                temp_conn.execute("PRAGMA wal_checkpoint(TRUNCATE)")
                temp_conn.commit()
                temp_conn.close()
            except sqlite3.OperationalError:
                pass  # не критично, файлы удалятся позже

        # Удаляем WAL/SHM файлы если они ещё есть и никто не держит БД
        try:
            if shm_path.exists():
                shm_path.unlink()
            if wal_path.exists():
                wal_path.unlink()
        except OSError:
            pass  # файлы заняты другим процессом — оставляем

    backup_db()  # бэкап перед инициализацией
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
        review_status TEXT DEFAULT 'free',
        review_taken_by TEXT,
        review_taken_at TEXT,
        review_completed_at TEXT,
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
        review_status TEXT DEFAULT 'free',
        review_taken_by TEXT,
        review_taken_at TEXT,
        review_completed_at TEXT,
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
    CREATE TABLE IF NOT EXISTS reviewers (
        username TEXT PRIMARY KEY,
        added_by TEXT,
        updated_at TEXT
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
    for column_name, column_sql in {
        "review_status": "TEXT DEFAULT 'free'",
        "review_taken_by": "TEXT",
        "review_taken_at": "TEXT",
        "review_completed_at": "TEXT",
    }.items():
        if column_name not in issue_columns:
            cur.execute(f"ALTER TABLE issues ADD COLUMN {column_name} {column_sql}")

    cur.execute("PRAGMA table_info(export_issues)")
    export_issue_columns = {row[1] for row in cur.fetchall()}
    for column_name, column_sql in {
        "review_status": "TEXT DEFAULT 'free'",
        "review_taken_by": "TEXT",
        "review_taken_at": "TEXT",
        "review_completed_at": "TEXT",
    }.items():
        if column_name not in export_issue_columns:
            cur.execute(f"ALTER TABLE export_issues ADD COLUMN {column_name} {column_sql}")

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

    cur.execute("""
    CREATE TRIGGER IF NOT EXISTS trg_reviewers_ai
    AFTER INSERT ON reviewers
    BEGIN
        INSERT INTO change_log(table_name, row_id, action, created_at)
        VALUES ('reviewers', NULL, 'I', datetime('now'));
    END;
    """)
    cur.execute("""
    CREATE TRIGGER IF NOT EXISTS trg_reviewers_au
    AFTER UPDATE ON reviewers
    BEGIN
        INSERT INTO change_log(table_name, row_id, action, created_at)
        VALUES ('reviewers', NULL, 'U', datetime('now'));
    END;
    """)
    cur.execute("""
    CREATE TRIGGER IF NOT EXISTS trg_reviewers_ad
    AFTER DELETE ON reviewers
    BEGIN
        INSERT INTO change_log(table_name, row_id, action, created_at)
        VALUES ('reviewers', NULL, 'D', datetime('now'));
    END;
    """)

    # Индексы для ускорения запросов
    cur.execute("CREATE INDEX IF NOT EXISTS idx_issues_status ON issues(status)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_issues_taken_by ON issues(taken_by)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_issues_journal ON issues(journal)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_issues_completed_at ON issues(completed_at)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_issues_review_status ON issues(review_status)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_issues_review_taken_by ON issues(review_taken_by)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_export_issues_status ON export_issues(status)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_export_issues_taken_by ON export_issues(taken_by)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_export_issues_journal ON export_issues(journal)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_export_issues_completed_at ON export_issues(completed_at)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_export_issues_review_status ON export_issues(review_status)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_export_issues_review_taken_by ON export_issues(review_taken_by)")

    conn.commit()
    conn.close()

    # Проверка целостности: если БД пустая — возможно, это локальный мусор
    check_conn = sqlite3.connect(DB_PATH, timeout=5)
    check_cur = check_conn.cursor()
    check_cur.execute("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='issues'")
    has_issues_table = check_cur.fetchone()[0] > 0
    check_conn.close()

    if not has_issues_table:
        # БД существует, но таблиц нет — удаляем и пересоздаём
        try:
            Path(DB_PATH).unlink()
        except OSError:
            pass
        # Пересоздаём с нуля
        conn = get_db_connection()
        cur = conn.cursor()
        # ... (здесь нужно создать все таблицы заново)
        # Но проще — вызвать init_db рекурсивно
        conn.close()
        init_db()
        return

_db_access_checked = False

def ensure_db_dir():
    global _db_access_checked
    if not DB_DIR.exists():
        raise sqlite3.OperationalError(
            f"Папка приложения не найдена: {DB_DIR}"
        )
    # Проверка прав — просто пробуем открыть БД
    # Удаление тестового файла убрано, т.к. может не быть прав на удаление
    if not _db_access_checked:
        try:
            if not Path(DB_PATH).exists():
                # Создаём БД если её нет (проверка прав на создание)
                test_conn = sqlite3.connect(DB_PATH, timeout=5)
                test_conn.close()
                # Не удаляем — пусть остаётся, это рабочая БД
            else:
                # БД существует — просто проверяем что можем открыть
                test_conn = sqlite3.connect(DB_PATH, timeout=5)
                test_conn.close()
            _db_access_checked = True
        except (OSError, sqlite3.OperationalError) as e:
            raise sqlite3.OperationalError(
                f"Нет доступа к базе данных: {DB_PATH}\n"
                f"Проверьте права доступа к сетевой папке.\n"
                f"Ошибка: {e}"
            )

def get_db_connection():
    ensure_db_dir()
    conn = sqlite3.connect(DB_PATH, timeout=10)
    conn.execute("PRAGMA busy_timeout = 5000")
    conn.execute("PRAGMA journal_mode = DELETE")
    conn.execute("PRAGMA synchronous = NORMAL")
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

# ================= БЭКАП БД =================

BACKUP_MAX_COUNT = 10


def backup_db():
    """Создаёт бэкап БД с датой/временем в имени. Хранит не более BACKUP_MAX_COUNT копий."""
    if not Path(DB_PATH).exists():
        return

    backup_dir = DB_DIR / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"journal_app_backup_{timestamp}.db"
    backup_path = backup_dir / backup_name

    try:
        shutil.copy2(DB_PATH, backup_path)
    except OSError:
        return  # тихо пропускаем, если нет доступа

    # Ротация: оставляем только последние BACKUP_MAX_COUNT
    existing = sorted(
        backup_dir.glob("journal_app_backup_*.db"),
        key=lambda p: p.stat().st_mtime,
        reverse=True
    )
    for old in existing[BACKUP_MAX_COUNT:]:
        try:
            old.unlink()
        except OSError:
            pass

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

    # Принудительно очищаем WAL-файлы оставшиеся от старого режима
    try:
        cleanup_conn = sqlite3.connect(DB_PATH, timeout=5)
        cleanup_conn.execute("PRAGMA journal_mode = DELETE")
        cleanup_conn.execute("PRAGMA wal_checkpoint(TRUNCATE)")
        cleanup_conn.close()
    except Exception:
        pass

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
        self.setWindowTitle(f"Журналы WindChill (v{APP_VERSION})")
        self.resize(1100, 600)

        self.is_admin = CURRENT_USER in ADMIN_USERS
        self.is_reviewer = self.user_is_reviewer(CURRENT_USER)
        self.search_timer = QTimer(self)
        self.search_timer.setSingleShot(True)
        self.search_timer.setInterval(300)
        self.search_timer.timeout.connect(self.load_data)

        self.export_search_timer = QTimer(self)
        self.export_search_timer.setSingleShot(True)
        self.export_search_timer.setInterval(300)
        self.export_search_timer.timeout.connect(self.load_export_data)

        self.date_filter_timer = QTimer(self)
        self.date_filter_timer.setSingleShot(True)
        self.date_filter_timer.setInterval(400)
        self.date_filter_timer.timeout.connect(self.load_data)

        self.export_date_filter_timer = QTimer(self)
        self.export_date_filter_timer.setSingleShot(True)
        self.export_date_filter_timer.setInterval(400)
        self.export_date_filter_timer.timeout.connect(self.load_export_data)

        self.refresh_timer = QTimer(self)
        self.refresh_timer.setSingleShot(False)
        self.refresh_timer.setInterval(AUTO_REFRESH_INTERVAL_MS)
        self.refresh_timer.timeout.connect(self.refresh_ui)
        self.last_db_signature = None
        self._preserve_scroll_on_reload = False
        self._preserve_export_scroll_on_reload = False

        self.tabs = QTabWidget()

        self.main_tab = QWidget()
        self.export_tab = QWidget()
        self.rating_tab = QWidget()

        self.tabs.addTab(self.main_tab, "📚 Журналы")
        self.tabs.addTab(self.export_tab, "📤 Скачать из Elibrary и выгрузить в WindChill")
        self.tabs.addTab(self.rating_tab, "🏆 Рейтинг")
        self.tabs.currentChanged.connect(self.on_tab_changed)

        self.central_container = QWidget()
        self.central_layout = QVBoxLayout(self.central_container)
        self.central_layout.setContentsMargins(0, 0, 0, 0)
        self.central_layout.setSpacing(0)

        self.update_banner = QFrame()
        self.update_banner.setObjectName("updateBanner")
        self.update_banner.setVisible(False)
        banner_layout = QHBoxLayout(self.update_banner)
        banner_layout.setContentsMargins(12, 8, 12, 8)
        self.update_banner_label = QLabel()
        self.update_banner_label.setWordWrap(True)
        banner_layout.addWidget(self.update_banner_label, 1)
        self.central_layout.addWidget(self.update_banner)
        self.central_layout.addWidget(self.tabs)
        self.setCentralWidget(self.central_container)

        self.update_banner.setStyleSheet("""
            QFrame#updateBanner {
                background-color: #fff3cd;
                border-bottom: 1px solid #ffeeba;
            }
        """)
        self.update_banner_label.setStyleSheet("color: #856404;")

        self.update_check_timer = QTimer(self)
        self.update_check_timer.setInterval(UPDATE_CHECK_INTERVAL_MS)
        self.update_check_timer.timeout.connect(self.check_for_updates)

        self.setup_main_tab()
        self.setup_export_tab()
        self.setup_rating_tab()
        self._tabs_loaded = {0}  # первая вкладка уже загружена
        self.load_data()
        self.remember_db_signature()
        self.refresh_timer.start()
        self.check_for_updates()
        self.update_check_timer.start()

    def check_for_updates(self):
        config = load_deployed_version_config()
        deployed_version = str(config.get("current", "")).strip()
        notes = str(config.get("notes", "")).strip()

        if not deployed_version or not is_version_newer(deployed_version, APP_VERSION):
            self.update_banner.setVisible(False)
            return

        message = (
            f"Доступна новая версия {deployed_version} "
            f"(сейчас установлена {APP_VERSION}). "
            f"Закройте программу и запустите её снова через JournalLauncher."
        )
        if notes:
            message += f"\n{notes}"

        self.update_banner_label.setText(message)
        self.update_banner.setVisible(True)

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

    def user_is_reviewer(self, username):
        if not username:
            return False

        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute(
            "SELECT 1 FROM reviewers WHERE LOWER(username)=LOWER(?)",
            (username.strip(),)
        )
        exists = cur.fetchone() is not None
        conn.close()
        return exists

    def load_reviewers(self):
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("SELECT username FROM reviewers ORDER BY username COLLATE NOCASE")
        reviewers = [row[0] for row in cur.fetchall()]
        conn.close()
        return reviewers

    def refresh_reviewers_list(self):
        if not self.is_admin or not getattr(self, "reviewers_list", None):
            return

        self.reviewers_list.clear()
        self.reviewers_list.addItems(self.load_reviewers())
        self.is_reviewer = self.user_is_reviewer(CURRENT_USER)

    def get_review_action_button(self, issue_id, status, review_status, review_taken_by, is_export=False):
        btn = QPushButton()

        if (
            self.is_reviewer
            and status == "done"
            and (review_status or "free") == "free"
        ):
            btn.setText("Взять на проверку")
            btn.setStyleSheet("""
                QPushButton {
                    background-color: #0d6efd;
                    color: white;
                    font-weight: 600;
                    border-radius: 6px;
                    padding: 4px 8px;
                }
                QPushButton:hover {
                    background-color: #0b5ed7;
                }
            """)
            btn.clicked.connect(
                lambda _, i=issue_id: self.take_review_issue(i, is_export=is_export)
            )
        elif (
            self.is_reviewer
            and status == "done"
            and review_status == "in_progress"
            and review_taken_by == CURRENT_USER
        ):
            btn.setText("Проверено")
            btn.setStyleSheet("""
                QPushButton {
                    background-color: #198754;
                    color: white;
                    font-weight: 600;
                    border-radius: 6px;
                    padding: 4px 8px;
                }
                QPushButton:hover {
                    background-color: #157347;
                }
            """)
            btn.clicked.connect(
                lambda _, i=issue_id: self.confirm_review_complete(i, is_export=is_export)
            )
        else:
            btn.setText("—")
            btn.setEnabled(False)
            if not self.is_reviewer:
                btn.setToolTip("Кнопка доступна только пользователям из списка проверяющих.")
            elif status != "done":
                btn.setToolTip("Проверку можно взять только после статуса «Готово».")
            elif review_status == "done":
                btn.setToolTip("Запись уже проверена.")
            elif review_status == "in_progress" and review_taken_by:
                btn.setToolTip(f"Запись уже на проверке у пользователя {review_taken_by}.")

        return btn

    def get_review_status_text(self, review_status, status=None):
        if status != "done":
            return "—"
        if review_status == "in_progress":
            return "🟡 На проверке"
        if review_status == "done":
            return "✅ Проверено"
        return "⚪ Свободно"

    def get_review_status_combo(self, issue_id, review_status, is_export=False):
        combo = NoWheelComboBox()
        combo.addItem("⚪ Свободно", "free")
        combo.addItem("🟡 На проверке", "in_progress")
        combo.addItem("✅ Проверено", "done")
        idx = combo.findData(review_status or "free")
        combo.setCurrentIndex(idx if idx >= 0 else 0)
        combo.currentIndexChanged.connect(
            lambda _, row_id=issue_id, widget=combo: self.change_review_status(
                row_id,
                widget.currentData(),
                is_export=is_export
            )
        )
        return combo

    def set_review_status_cell(self, table, row, issue_id, status, review_status, is_export=False):
        table.removeCellWidget(row, 8)

        if status == "done" and self.is_admin:
            table.setCellWidget(
                row,
                8,
                self.get_review_status_combo(issue_id, review_status, is_export=is_export)
            )
            return

        review_status_item = table.item(row, 8) or QTableWidgetItem()
        review_status_item.setText(self.get_review_status_text(review_status, status))
        review_status_item.setData(Qt.UserRole, issue_id)

        if status == "done" and review_status == "in_progress":
            review_status_item.setBackground(QColor("#fff3cd"))
        elif status == "done" and review_status == "done":
            review_status_item.setBackground(QColor("#d4edda"))
        else:
            review_status_item.setData(Qt.BackgroundRole, None)

        table.setItem(row, 8, review_status_item)

    def set_completed_at_cell(self, table, row, issue_id, completed_at, status):
        if not self.is_admin:
            return

        item = table.item(row, ISSUES_TABLE_COLUMN_COMPLETED_AT) or QTableWidgetItem()
        item.setData(Qt.UserRole, issue_id)
        item.setToolTip("ДД.ММ.ГГГГ или ГГГГ-ММ-ДД")

        if status == "done":
            item.setText(format_completed_at_for_table(completed_at))
            item.setFlags(item.flags() | Qt.ItemIsEditable)
        else:
            item.setText("")
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)

        table.setItem(row, ISSUES_TABLE_COLUMN_COMPLETED_AT, item)

    def save_completed_at(self, issue_id, text, is_export=False):
        table_name = "export_issues" if is_export else "issues"
        refresh_row = self.refresh_export_issue_row if is_export else self.refresh_issue_row

        try:
            parsed_value = parse_completed_at_input(text)
        except ValueError as exc:
            QMessageBox.warning(self, "Ошибка", str(exc))
            refresh_row(issue_id)
            return

        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute(
            f"UPDATE {table_name} SET completed_at=?, updated_at=? WHERE id=?",
            (parsed_value, datetime.now().isoformat(), issue_id),
        )
        conn.commit()
        conn.close()
        self.remember_db_signature()
        refresh_row(issue_id)

    def change_review_status(self, issue_id, review_status, is_export=False):
        if not self.is_admin:
            return

        table_name = "export_issues" if is_export else "issues"
        now = datetime.now().isoformat()

        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute(
            f"SELECT status, review_taken_by, review_taken_at FROM {table_name} WHERE id=?",
            (issue_id,)
        )
        row = cur.fetchone()
        if not row:
            conn.close()
            return

        status, review_taken_by, review_taken_at = row

        if review_status == "free":
            cur.execute(f"""
                UPDATE {table_name}
                SET review_status='free',
                    review_taken_by=NULL,
                    review_taken_at=NULL,
                    review_completed_at=NULL,
                    updated_at=?
                WHERE id=?
            """, (now, issue_id))
        elif review_status == "in_progress":
            if status != "done":
                conn.close()
                QMessageBox.warning(
                    self,
                    "Внимание",
                    "Проверку можно начать только для записей со статусом «Готово»."
                )
                if is_export:
                    self.refresh_export_issue_row(issue_id)
                else:
                    self.refresh_issue_row(issue_id)
                return

            cur.execute(f"""
                UPDATE {table_name}
                SET review_status='in_progress',
                    review_taken_by=?,
                    review_taken_at=?,
                    review_completed_at=NULL,
                    updated_at=?
                WHERE id=?
            """, (
                review_taken_by or CURRENT_USER,
                review_taken_at or now,
                now,
                issue_id
            ))
        elif review_status == "done":
            if status != "done":
                conn.close()
                QMessageBox.warning(
                    self,
                    "Внимание",
                    "Проверить можно только записи со статусом «Готово»."
                )
                if is_export:
                    self.refresh_export_issue_row(issue_id)
                else:
                    self.refresh_issue_row(issue_id)
                return

            cur.execute(f"""
                UPDATE {table_name}
                SET review_status='done',
                    review_taken_by=?,
                    review_taken_at=?,
                    review_completed_at=?,
                    updated_at=?
                WHERE id=?
            """, (
                review_taken_by or CURRENT_USER,
                review_taken_at or now,
                now,
                now,
                issue_id
            ))
        else:
            conn.close()
            return

        conn.commit()
        conn.close()

        if is_export:
            self.refresh_export_issue_row(issue_id)
        else:
            self.refresh_issue_row(issue_id)
        self.remember_db_signature()

    def take_review_issue(self, issue_id, is_export=False):
        if not self.is_reviewer:
            return

        table_name = "export_issues" if is_export else "issues"
        conn = get_db_connection()
        cur = conn.cursor()

        cur.execute(f"""
            UPDATE {table_name}
            SET review_status='in_progress',
                review_taken_by=?,
                review_taken_at=?,
                review_completed_at=NULL,
                updated_at=?
            WHERE id=?
              AND status='done'
              AND COALESCE(review_status, 'free')='free'
        """, (CURRENT_USER, datetime.now().isoformat(), datetime.now().isoformat(), issue_id))

        if cur.rowcount == 0:
            cur.execute(f"SELECT review_status, review_taken_by FROM {table_name} WHERE id=?", (issue_id,))
            row = cur.fetchone()
            conn.close()

            if row and row[0] == "in_progress" and row[1] and row[1] != CURRENT_USER:
                QMessageBox.warning(
                    self,
                    "Внимание",
                    f"Запись уже на проверке у пользователя {row[1]}."
                )
            else:
                QMessageBox.warning(self, "Внимание", "Эту запись нельзя взять на проверку.")
            return

        conn.commit()
        conn.close()

        if is_export:
            self.refresh_export_issue_row(issue_id)
        else:
            self.refresh_issue_row(issue_id)
        self.remember_db_signature()

    def confirm_review_complete(self, issue_id, is_export=False):
        reply = QMessageBox.question(
            self,
            "Подтверждение",
            "Подтвердить, что запись проверена?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.complete_review_issue(issue_id, is_export=is_export)

    def complete_review_issue(self, issue_id, is_export=False):
        if not self.is_reviewer:
            return

        table_name = "export_issues" if is_export else "issues"
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute(f"""
            UPDATE {table_name}
            SET review_status='done',
                review_completed_at=?,
                updated_at=?
            WHERE id=?
              AND review_status='in_progress'
              AND review_taken_by=?
        """, (datetime.now().isoformat(), datetime.now().isoformat(), issue_id, CURRENT_USER))
        conn.commit()
        conn.close()

        if is_export:
            self.refresh_export_issue_row(issue_id)
        else:
            self.refresh_issue_row(issue_id)
        self.remember_db_signature()

    def add_reviewer(self):
        if not self.is_admin:
            return

        username, ok = QInputDialog.getText(
            self,
            "Добавить проверяющего",
            "Введите Windows-логин пользователя:"
        )
        username = username.strip() if ok else ""
        if not username:
            return

        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("""
            INSERT INTO reviewers (username, added_by, updated_at)
            VALUES (?, ?, ?)
            ON CONFLICT(username) DO UPDATE SET
                added_by=excluded.added_by,
                updated_at=excluded.updated_at
        """, (username, CURRENT_USER, datetime.now().isoformat()))
        conn.commit()
        conn.close()

        self.refresh_reviewers_list()
        self.load_data()
        if hasattr(self, "export_table"):
            self.load_export_data()
        self.remember_db_signature()

    def delete_selected_reviewer(self):
        if not self.is_admin or not getattr(self, "reviewers_list", None):
            return

        selected_items = self.reviewers_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Внимание", "Выберите проверяющего для удаления.")
            return

        username = selected_items[0].text()
        reply = QMessageBox.question(
            self,
            "Удалить проверяющего",
            f"Удалить пользователя {username} из списка проверяющих?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("DELETE FROM reviewers WHERE username=?", (username,))
        conn.commit()
        conn.close()

        self.refresh_reviewers_list()
        self.load_data()
        if hasattr(self, "export_table"):
            self.load_export_data()
        self.remember_db_signature()

    def schedule_load_data(self):
        self.search_timer.start()

    def schedule_completed_at_filter(self):
        """Перезагрузка только при пустом поле или полной валидной дате."""
        text = self.completed_at_filter_box.text().strip()
        if text and parse_date_filter_input(text) is None:
            return
        self.date_filter_timer.start()

    def schedule_export_load_data(self):
        self.export_search_timer.start()

    def schedule_export_completed_at_filter(self):
        text = self.export_completed_at_filter_box.text().strip()
        if text and parse_date_filter_input(text) is None:
            return
        self.export_date_filter_timer.start()

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
                   MIN(completed_at) AS first_done_at
            FROM (
                SELECT taken_by, completed_at, id FROM issues
                WHERE status='done' AND completed_at LIKE ? AND taken_by IS NOT NULL AND TRIM(taken_by) != ''
                UNION ALL
                SELECT taken_by, completed_at, id FROM export_issues
                WHERE status='done' AND completed_at LIKE ? AND taken_by IS NOT NULL AND TRIM(taken_by) != ''
            )
            GROUP BY taken_by
            ORDER BY cnt DESC, first_done_at ASC
        """, (month_str + "%", month_str + "%"))
        rows = cur.fetchall()
        conn.close()

        return [(taken_by, cnt) for taken_by, cnt, _ in rows]

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
        if index == self.tabs.indexOf(self.rating_tab) and hasattr(self, "rating_month_rank_labels"):
            self.update_rating()  # рейтинг всегда обновляется при открытии
            return

        if index in self._tabs_loaded:
            return
        self._tabs_loaded.add(index)

        if index == self.tabs.indexOf(self.export_tab) and hasattr(self, "export_table"):
            self.load_export_journal_filter_options()
            self.load_export_data()

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
        except sqlite3.OperationalError:
            self.last_db_signature = None

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
        reviewers_changed = any(table_name == "reviewers" for _, table_name, _, _ in changes)
        rating_changed = any(
            table_name in {"issues", "monthly_plans", "rating_months"}
            for _, table_name, _, _ in changes
        )

        if reviewers_changed:
            self.is_reviewer = self.user_is_reviewer(CURRENT_USER)
            if self.is_admin and hasattr(self, "reviewers_list"):
                self.refresh_reviewers_list()

        if current_tab == self.main_tab:
            if reviewers_changed:
                self._preserve_scroll_on_reload = True
                try:
                    self.load_data()
                finally:
                    self._preserve_scroll_on_reload = False
            elif issue_changes:
                issue_actions = {action for _, table_name, _, action in changes if table_name == "issues"}
                if "I" in issue_actions or "D" in issue_actions:
                    self.load_journal_filter_options()
                    self._preserve_scroll_on_reload = True
                    try:
                        self.load_data()
                    finally:
                        self._preserve_scroll_on_reload = False
                else:
                    for row_id in sorted(set(issue_changes)):
                        self.refresh_issue_row(row_id)

        elif current_tab == self.export_tab:
            if reviewers_changed:
                self._preserve_export_scroll_on_reload = True
                try:
                    self.load_export_data()
                finally:
                    self._preserve_export_scroll_on_reload = False
            elif export_changes:
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
            if rating_changed or export_changes:
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
            SELECT id, journal, year, issue, path, status, taken_by, taken_at, completed_at,
                   review_status, review_taken_by, review_taken_at, review_completed_at
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

    def issue_matches_current_view(
        self,
        journal,
        year,
        issue,
        path,
        status,
        taken_by,
        taken_at,
        completed_at,
        review_status=None,
        review_taken_by=None,
        review_taken_at=None,
        review_completed_at=None
    ):
        filter_value = self.filter_box.currentText()
        review_filter_value = (
            self.review_filter_box.currentText()
            if hasattr(self, "review_filter_box")
            else "Все"
        )
        journal_filter = self.journal_filter_box.currentText()
        search_text = self.search_box.text().strip().casefold()
        completed_at_filter_date = parse_date_filter_input(
            self.completed_at_filter_box.text()
            if hasattr(self, "completed_at_filter_box")
            else ""
        )

        if filter_value == "Свободные" and status != "free":
            return False
        if filter_value == "В работе" and status != "in_progress":
            return False
        if filter_value == "Готово" and status != "done":
            return False
        if filter_value == "Мои" and taken_by != CURRENT_USER:
            return False

        normalized_review_status = review_status or "free"
        if review_filter_value == "Свободные" and (status != "done" or normalized_review_status != "free"):
            return False
        if review_filter_value == "В работе" and (status != "done" or normalized_review_status != "in_progress"):
            return False
        if review_filter_value == "Готово" and (status != "done" or normalized_review_status != "done"):
            return False
        if review_filter_value == "Мои" and (status != "done" or review_taken_by != CURRENT_USER):
            return False

        if journal_filter != "Все журналы" and (journal or "").strip().casefold() != journal_filter.strip().casefold():
            return False

        if not completed_at_matches_date_filter(completed_at, completed_at_filter_date):
            return False

        return issue_row_matches_search(
            search_text,
            journal=journal,
            year=year,
            issue=issue,
            path=path,
            status=status,
            taken_by=taken_by,
            review_status=review_status,
            review_taken_by=review_taken_by,
        )

    def check_issue_integrity(self):
        """Подсвечивает дубли: один и тот же журнал/год/выпуск в работе или готов у разных пользователей."""
        if not self.is_admin:
            return

        conn = get_db_connection()
        cur = conn.cursor()

        # Ищем дубли по journal + year + issue, где статус in_progress или done
        cur.execute("""
            SELECT journal, year, issue, GROUP_CONCAT(taken_by, '; ') AS users, COUNT(*) AS cnt
            FROM issues
            WHERE status IN ('in_progress', 'done')
              AND taken_by IS NOT NULL
              AND TRIM(taken_by) != ''
            GROUP BY journal, year, issue
            HAVING COUNT(DISTINCT taken_by) > 1
        """)
        duplicate_groups = cur.fetchall()
        conn.close()

        if not duplicate_groups:
            return

        # Собираем множество id проблемных строк
        conn = get_db_connection()
        cur = conn.cursor()

        problem_ids = set()
        problem_info = {}  # id -> "Пользователь1; Пользователь2"

        for journal, year, issue, users, cnt in duplicate_groups:
            cur.execute("""
                SELECT id, taken_by FROM issues
                WHERE journal=? AND year=? AND issue=?
                  AND status IN ('in_progress', 'done')
            """, (journal, year, issue))
            rows = cur.fetchall()
            for issue_id, taken_by in rows:
                problem_ids.add(issue_id)
                problem_info[issue_id] = users

        conn.close()

        # Подсвечиваем в таблице
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            if not item:
                continue
            issue_id = item.data(Qt.UserRole)

            who_item = self.table.item(row, 4)

            if issue_id in problem_ids:
                # Подсвечиваем всю строку
                for col in range(self.table.columnCount()):
                    cell = self.table.item(row, col)
                    if cell:
                        cell.setBackground(QColor("#f8d7da"))
                if who_item:
                    who_item.setToolTip(
                        f"⚠ ДУБЛЬ! Этот выпуск также у: {problem_info[issue_id]}"
                    )
            else:
                # Сбрасываем фон
                for col in range(self.table.columnCount()):
                    cell = self.table.item(row, col)
                    if cell:
                        cell.setBackground(QColor())
                if who_item:
                    who_item.setToolTip("")

    def check_export_issue_integrity(self):
        """Подсвечивает дубли в export_issues."""
        if not self.is_admin:
            return

        conn = get_db_connection()
        cur = conn.cursor()

        cur.execute("""
            SELECT journal, year, issue, GROUP_CONCAT(taken_by, '; ') AS users, COUNT(*) AS cnt
            FROM export_issues
            WHERE status IN ('in_progress', 'done')
              AND taken_by IS NOT NULL
              AND TRIM(taken_by) != ''
            GROUP BY journal, year, issue
            HAVING COUNT(DISTINCT taken_by) > 1
        """)
        duplicate_groups = cur.fetchall()
        conn.close()

        if not duplicate_groups:
            return

        conn = get_db_connection()
        cur = conn.cursor()

        problem_ids = set()
        problem_info = {}

        for journal, year, issue, users, cnt in duplicate_groups:
            cur.execute("""
                SELECT id, taken_by FROM export_issues
                WHERE journal=? AND year=? AND issue=?
                  AND status IN ('in_progress', 'done')
            """, (journal, year, issue))
            rows = cur.fetchall()
            for issue_id, taken_by in rows:
                problem_ids.add(issue_id)
                problem_info[issue_id] = users

        conn.close()

        for row in range(self.export_table.rowCount()):
            item = self.export_table.item(row, 0)
            if not item:
                continue
            issue_id = item.data(Qt.UserRole)

            who_item = self.export_table.item(row, 4)

            if issue_id in problem_ids:
                for col in range(self.export_table.columnCount()):
                    cell = self.export_table.item(row, col)
                    if cell:
                        cell.setBackground(QColor("#f8d7da"))
                if who_item:
                    who_item.setToolTip(
                        f"⚠ ДУБЛЬ! Этот выпуск также у: {problem_info[issue_id]}"
                    )
            else:
                for col in range(self.export_table.columnCount()):
                    cell = self.export_table.item(row, col)
                    if cell:
                        cell.setBackground(QColor())
                if who_item:
                    who_item.setToolTip("")

    def refresh_issue_row(self, issue_id):
        row_data = self.get_issue_row_data(issue_id)
        if not row_data:
            row = self.find_issue_row(issue_id)
            if row >= 0:
                self.table.removeRow(row)
            return

        (
            id_,
            journal,
            year,
            issue,
            path,
            status,
            taken_by,
            taken_at,
            completed_at,
            review_status,
            review_taken_by,
            review_taken_at,
            review_completed_at,
        ) = row_data

        row = self.find_issue_row(issue_id)

        # Если текущие фильтры/поиск уже не подходят — просто убираем строку из таблицы
        if not self.issue_matches_current_view(
            journal,
            year,
            issue,
            path,
            status,
            taken_by,
            taken_at,
            completed_at,
            review_status,
            review_taken_by,
            review_taken_at,
            review_completed_at
        ):
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

            reviewer_item = self.table.item(row, 6) or QTableWidgetItem()
            reviewer_item.setText(review_taken_by or "")
            reviewer_item.setData(Qt.UserRole, id_)
            if review_status == "in_progress":
                reviewer_item.setBackground(QColor("#fff3cd"))
            elif review_status == "done":
                reviewer_item.setBackground(QColor("#d4edda"))
            else:
                reviewer_item.setData(Qt.BackgroundRole, None)
            self.table.setItem(row, 6, reviewer_item)

            if self.is_admin:
                status_combo = NoWheelComboBox()
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
                btn.clicked.connect(lambda _, i=id_: self.confirm_complete(i, is_export=False))
            else:
                btn.setText("—")
                btn.setEnabled(False)

            self.table.setCellWidget(row, 5, btn)
            self.table.setCellWidget(
                row,
                7,
                self.get_review_action_button(id_, status, review_status, review_taken_by)
            )
            self.set_review_status_cell(self.table, row, id_, status, review_status)
            self.set_completed_at_cell(self.table, row, id_, completed_at, status)
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

        self.completed_at_filter_box = QLineEdit()
        self.completed_at_filter_box.setPlaceholderText("ДД.ММ.ГГГГ")
        self.completed_at_filter_box.setClearButtonEnabled(True)
        self.completed_at_filter_box.setMaximumWidth(120)
        self.completed_at_filter_box.textChanged.connect(self.schedule_completed_at_filter)

        if self.is_admin or self.is_reviewer:
            self.review_filter_box = QComboBox()
            self.review_filter_box.addItems(["Все", "Свободные", "В работе", "Мои", "Готово"])

        if self.is_admin:
            refresh_btn = QPushButton("🔄 Обновить")
            refresh_btn.clicked.connect(self.load_data)

            delete_btn = QPushButton("🗑 Удалить выбранный журнал")
            delete_btn.clicked.connect(self.delete_selected_issue)

            reviewers_btn = QPushButton("👥 Проверяющие")
            reviewers_btn.clicked.connect(self.open_reviewers_dialog)

        export_btn = QPushButton("📊 Экспорт отчета")
        export_btn.clicked.connect(self.export_report)

        filter_layout.addWidget(QLabel("Поиск:"))
        filter_layout.addWidget(self.search_box)
        filter_layout.addWidget(QLabel("Журнал:"))
        filter_layout.addWidget(self.journal_filter_box)
        filter_layout.addWidget(QLabel("Статус:"))
        filter_layout.addWidget(self.filter_box)
        filter_layout.addWidget(QLabel("Дата готовности:"))
        filter_layout.addWidget(self.completed_at_filter_box)
        if self.is_admin or self.is_reviewer:
            filter_layout.addWidget(QLabel("Статус проверки:"))
            filter_layout.addWidget(self.review_filter_box)
        filter_layout.addStretch()

        if self.is_admin:
            sync_btn = QPushButton("📥 Импорт журналов")
            sync_btn.clicked.connect(self.sync_from_network)
            filter_layout.addWidget(sync_btn)

        if self.is_admin:
            filter_layout.addWidget(delete_btn)
            filter_layout.addWidget(reviewers_btn)
            filter_layout.addWidget(refresh_btn)

        filter_layout.addWidget(export_btn)

        filter_frame_layout.addLayout(filter_layout)

        # ================= КАРТОЧКА ТАБЛИЦЫ =================
        self.table_frame = QFrame()
        self.table_frame.setObjectName("ratingCard")
        table_frame_layout = QVBoxLayout(self.table_frame)

        self.table = QTableWidget()
        self.table.setColumnCount(10)
        self.table.setHorizontalHeaderLabels(
            [
                "Журнал", "Год", "Выпуск", "Статус", "Кто", "Действие",
                "Кто проверил", "Проверка", "Статус проверки", "Дата готовности",
            ]
        )
        self.table.setColumnHidden(ISSUES_TABLE_COLUMN_COMPLETED_AT, not self.is_admin)
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

        self.table.setAlternatingRowColors(True)

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
        apply_static_column_widths(
            self.table,
            ISSUES_TABLE_FIXED_COLUMN_SAMPLES,
            column_extra_padding=ISSUES_TABLE_COLUMN_EXTRA_PADDING,
            columns_to_configure=range(3, self.table.columnCount()),
        )
        fit_table_columns_to_content(self.table, ISSUES_TEXT_FIT_COLUMNS)
        # Ускорение прокрутки для больших таблиц
        self.table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)

        table_frame_layout.addWidget(self.table)

        self.journal_filter_box.currentTextChanged.connect(self.schedule_load_data)
        self.filter_box.currentTextChanged.connect(self.schedule_load_data)
        if self.is_admin or self.is_reviewer:
            self.review_filter_box.currentTextChanged.connect(self.schedule_load_data)
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
            QLineEdit {
                border: 1px solid #d7dde5;
                border-radius: 6px;
                padding: 4px 8px;
                background-color: #ffffff;
            }
            QComboBox {
                border: 1px solid #d7dde5;
                border-radius: 6px;
                padding: 4px 8px;
                background-color: #ffffff;
            }
            QComboBox::drop-down {
                border: none;
                border-radius: 6px;
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

        self.export_completed_at_filter_box = QLineEdit()
        self.export_completed_at_filter_box.setPlaceholderText("ДД.ММ.ГГГГ")
        self.export_completed_at_filter_box.setClearButtonEnabled(True)
        self.export_completed_at_filter_box.setMaximumWidth(120)
        self.export_completed_at_filter_box.textChanged.connect(self.schedule_export_completed_at_filter)

        if self.is_admin or self.is_reviewer:
            self.export_review_filter_box = QComboBox()
            self.export_review_filter_box.addItems(["Все", "Свободные", "В работе", "Мои", "Готово"])

        filter_layout.addWidget(QLabel("Поиск:"))
        filter_layout.addWidget(self.export_search_box)
        filter_layout.addWidget(QLabel("Журнал:"))
        filter_layout.addWidget(self.export_journal_filter_box)
        filter_layout.addWidget(QLabel("Статус:"))
        filter_layout.addWidget(self.export_filter_box)
        filter_layout.addWidget(QLabel("Дата готовности:"))
        filter_layout.addWidget(self.export_completed_at_filter_box)
        if self.is_admin or self.is_reviewer:
            filter_layout.addWidget(QLabel("Статус проверки:"))
            filter_layout.addWidget(self.export_review_filter_box)
        filter_layout.addStretch()

        export_excel_btn = QPushButton("📊 Экспорт в Excel")
        export_excel_btn.clicked.connect(self.export_export_tab_to_excel)
        filter_layout.addWidget(export_excel_btn)

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

        self.export_progress_table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)

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
        apply_static_column_widths(
            self.export_progress_table,
            RATING_PROGRESS_COLUMN_SAMPLES,
        )

        stats_layout.addWidget(stats_title)
        stats_layout.addWidget(self.export_summary_label)
        stats_layout.addWidget(self.export_progress_table)

        # ================= КАРТОЧКА ТАБЛИЦЫ =================
        self.export_table_frame = QFrame()
        self.export_table_frame.setObjectName("ratingCard")
        table_frame_layout = QVBoxLayout(self.export_table_frame)

        self.export_table = QTableWidget()
        self.export_table.setColumnCount(10)
        self.export_table.setHorizontalHeaderLabels(
            [
                "Журнал", "Год", "Выпуск", "Статус", "Кто", "Действие",
                "Кто проверил", "Проверка", "Статус проверки", "Дата готовности",
            ]
        )
        self.export_table.setColumnHidden(ISSUES_TABLE_COLUMN_COMPLETED_AT, not self.is_admin)
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

        self.export_table.setAlternatingRowColors(True)

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
        apply_static_column_widths(
            self.export_table,
            ISSUES_TABLE_FIXED_COLUMN_SAMPLES,
            column_extra_padding=ISSUES_TABLE_COLUMN_EXTRA_PADDING,
            columns_to_configure=range(3, self.export_table.columnCount()),
        )
        fit_table_columns_to_content(self.export_table, ISSUES_TEXT_FIT_COLUMNS)
        self.export_table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)

        table_frame_layout.addWidget(self.export_table)

        self.export_journal_filter_box.currentTextChanged.connect(self.schedule_export_load_data)
        self.export_filter_box.currentTextChanged.connect(self.schedule_export_load_data)
        if self.is_admin or self.is_reviewer:
            self.export_review_filter_box.currentTextChanged.connect(self.schedule_export_load_data)

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
            QLineEdit {
                border: 1px solid #d7dde5;
                border-radius: 6px;
                padding: 4px 8px;
                background-color: #ffffff;
            }
            QComboBox {
                border: 1px solid #d7dde5;
                border-radius: 6px;
                padding: 4px 8px;
                background-color: #ffffff;
            }
            QComboBox::drop-down {
                border: none;
                border-radius: 6px;
            }
        """)

    def open_reviewers_dialog(self):
        if not self.is_admin:
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Проверяющие")
        dialog.resize(420, 420)

        layout = QVBoxLayout()

        frame = QFrame()
        frame.setObjectName("ratingCard")
        frame_layout = QVBoxLayout(frame)

        title = QLabel("Проверяющие")
        title.setStyleSheet("font-weight: 700; font-size: 14px; color: #1f2937;")

        hint = QLabel(
            "Добавляйте Windows-логины сотрудников, которым доступна кнопка проверки."
        )
        hint.setWordWrap(True)

        self.reviewers_list = QListWidget()

        buttons_layout = QHBoxLayout()
        add_btn = QPushButton("Добавить проверяющего")
        add_btn.clicked.connect(self.add_reviewer)

        delete_btn = QPushButton("Удалить выбранного")
        delete_btn.clicked.connect(self.delete_selected_reviewer)

        buttons_layout.addWidget(add_btn)
        buttons_layout.addWidget(delete_btn)
        buttons_layout.addStretch()

        frame_layout.addWidget(title)
        frame_layout.addWidget(hint)
        frame_layout.addWidget(self.reviewers_list)
        frame_layout.addLayout(buttons_layout)

        layout.addWidget(frame)
        dialog.setLayout(layout)

        dialog.setStyleSheet("""
            QFrame#ratingCard {
                background-color: #ffffff;
                border: 1px solid #d7dde5;
                border-radius: 10px;
                padding: 10px;
            }
            QListWidget {
                background-color: #ffffff;
                border: 1px solid #d7dde5;
                border-radius: 6px;
                padding: 4px;
            }
        """)

        self.refresh_reviewers_list()
        dialog.exec()
        self.reviewers_list = None

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

    def issue_matches_export_current_view(
        self,
        journal,
        year,
        issue,
        status,
        taken_by,
        taken_at,
        completed_at,
        review_status=None,
        review_taken_by=None,
        review_taken_at=None,
        review_completed_at=None
    ):
        filter_value = self.export_filter_box.currentText()
        review_filter_value = (
            self.export_review_filter_box.currentText()
            if hasattr(self, "export_review_filter_box")
            else "Все"
        )
        journal_filter = self.export_journal_filter_box.currentText()
        search_text = self.export_search_box.text().strip().casefold()
        completed_at_filter_date = parse_date_filter_input(
            self.export_completed_at_filter_box.text()
            if hasattr(self, "export_completed_at_filter_box")
            else ""
        )

        if filter_value == "Свободные" and status != "free":
            return False
        if filter_value == "В работе" and status != "in_progress":
            return False
        if filter_value == "Готово" and status != "done":
            return False
        if filter_value == "Мои" and taken_by != CURRENT_USER:
            return False

        normalized_review_status = review_status or "free"
        if review_filter_value == "Свободные" and (status != "done" or normalized_review_status != "free"):
            return False
        if review_filter_value == "В работе" and (status != "done" or normalized_review_status != "in_progress"):
            return False
        if review_filter_value == "Готово" and (status != "done" or normalized_review_status != "done"):
            return False
        if review_filter_value == "Мои" and (status != "done" or review_taken_by != CURRENT_USER):
            return False

        if journal_filter != "Все журналы" and (journal or "").strip().casefold() != journal_filter.strip().casefold():
            return False

        if not completed_at_matches_date_filter(completed_at, completed_at_filter_date):
            return False

        return issue_row_matches_search(
            search_text,
            journal=journal,
            year=year,
            issue=issue,
            status=status,
            taken_by=taken_by,
            review_status=review_status,
            review_taken_by=review_taken_by,
        )

    def get_export_issue_row_data(self, issue_id):
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("""
            SELECT id, journal, year, issue, status, taken_by, taken_at, completed_at,
                   review_status, review_taken_by, review_taken_at, review_completed_at
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

        (
            id_,
            journal,
            year,
            issue,
            status,
            taken_by,
            taken_at,
            completed_at,
            review_status,
            review_taken_by,
            review_taken_at,
            review_completed_at,
        ) = row_data

        row = self.find_export_issue_row(issue_id)

        if not self.issue_matches_export_current_view(
            journal,
            year,
            issue,
            status,
            taken_by,
            taken_at,
            completed_at,
            review_status,
            review_taken_by,
            review_taken_at,
            review_completed_at
        ):
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

            reviewer_item = self.export_table.item(row, 6) or QTableWidgetItem()
            reviewer_item.setText(review_taken_by or "")
            reviewer_item.setData(Qt.UserRole, id_)
            if review_status == "in_progress":
                reviewer_item.setBackground(QColor("#fff3cd"))
            elif review_status == "done":
                reviewer_item.setBackground(QColor("#d4edda"))
            else:
                reviewer_item.setData(Qt.BackgroundRole, None)
            self.export_table.setItem(row, 6, reviewer_item)

            if self.is_admin:
                status_combo = NoWheelComboBox()
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
                btn.clicked.connect(lambda _, i=id_: self.confirm_complete(i, is_export=True))
            else:
                btn.setText("—")
                btn.setEnabled(False)

            self.export_table.setCellWidget(row, 5, btn)
            self.export_table.setCellWidget(
                row,
                7,
                self.get_review_action_button(id_, status, review_status, review_taken_by, is_export=True)
            )
            self.set_review_status_cell(self.export_table, row, id_, status, review_status, is_export=True)
            self.set_completed_at_cell(self.export_table, row, id_, completed_at, status)
        finally:
            self.export_table.blockSignals(False)

        self.update_export_progress_table()

    def load_export_data(self):
        scroll_value = self.export_table.verticalScrollBar().value() if self._preserve_export_scroll_on_reload else 0

        self.load_export_journal_filter_options()

        conn = get_db_connection()
        cur = conn.cursor()

        filter_value = self.export_filter_box.currentText()
        review_filter_value = (
            self.export_review_filter_box.currentText()
            if hasattr(self, "export_review_filter_box")
            else "Все"
        )
        journal_filter = self.export_journal_filter_box.currentText()
        search_text = self.export_search_box.text().strip().casefold()
        completed_at_filter_date = parse_date_filter_input(self.export_completed_at_filter_box.text())

        query = """
            SELECT id, journal, year, issue, status, taken_by, taken_at, completed_at,
                   review_status, review_taken_by, review_taken_at, review_completed_at
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

        if review_filter_value == "Свободные":
            query += " AND status = 'done' AND COALESCE(review_status, 'free') = 'free'"
        elif review_filter_value == "В работе":
            query += " AND status = 'done' AND review_status = 'in_progress'"
        elif review_filter_value == "Готово":
            query += " AND status = 'done' AND review_status = 'done'"
        elif review_filter_value == "Мои":
            query += " AND status = 'done' AND review_taken_by = ?"
            params.append(CURRENT_USER)

        if journal_filter != "Все журналы":
            query += " AND journal = ?"
            params.append(journal_filter)

        query, params = append_completed_at_date_sql(query, params, completed_at_filter_date)

        query += " ORDER BY journal COLLATE NOCASE, year, issue"

        cur.execute(query, params)
        rows = cur.fetchall()
        conn.close()

        if search_text:
            rows = filter_issue_rows(
                rows,
                search_text,
                with_path=False,
            )

        status_texts = {"free": "⚪ Свободен", "in_progress": "🟡 В работе", "done": "✅ Готово"}
        status_bg = {"in_progress": QColor("#fff3cd"), "done": QColor("#d4edda")}

        self.export_table.setRowCount(0)
        self.export_table.setUpdatesEnabled(False)
        self.export_table.blockSignals(True)

        for row_idx, row_data in enumerate(rows):
            (
                id_,
                journal,
                year,
                issue,
                status,
                taken_by,
                taken_at,
                completed_at,
                review_status,
                review_taken_by,
                review_taken_at,
                review_completed_at,
            ) = row_data

            self.export_table.insertRow(row_idx)

            journal_item = QTableWidgetItem(journal or "")
            journal_item.setData(Qt.UserRole, id_)
            year_item = QTableWidgetItem(year or "")
            year_item.setData(Qt.UserRole, id_)
            issue_item = QTableWidgetItem(issue or "")
            issue_item.setData(Qt.UserRole, id_)
            who_item = QTableWidgetItem(taken_by or "")
            who_item.setData(Qt.UserRole, id_)
            reviewer_item = QTableWidgetItem(review_taken_by or "")
            reviewer_item.setData(Qt.UserRole, id_)

            if self.is_admin:
                journal_item.setFlags(journal_item.flags() | Qt.ItemIsEditable)
                year_item.setFlags(year_item.flags() | Qt.ItemIsEditable)
                issue_item.setFlags(issue_item.flags() | Qt.ItemIsEditable)
                who_item.setFlags(who_item.flags() | Qt.ItemIsEditable)

                status_combo = NoWheelComboBox()
                status_combo.addItem("⚪ Свободен", "free")
                status_combo.addItem("🟡 В работе", "in_progress")
                status_combo.addItem("✅ Готово", "done")
                idx = status_combo.findData(status)
                status_combo.setCurrentIndex(idx if idx >= 0 else 0)
                status_combo.currentIndexChanged.connect(
                    lambda _, issue_id=id_, combo=status_combo: self.change_export_status(issue_id, combo.currentData())
                )
                self.export_table.setCellWidget(row_idx, 3, status_combo)
            else:
                status_item = QTableWidgetItem(status_texts.get(status, status))
                bg = status_bg.get(status)
                if bg:
                    status_item.setBackground(bg)
                self.export_table.setItem(row_idx, 3, status_item)

            self.export_table.setItem(row_idx, 0, journal_item)
            self.export_table.setItem(row_idx, 1, year_item)
            self.export_table.setItem(row_idx, 2, issue_item)
            self.export_table.setItem(row_idx, 4, who_item)
            if review_status == "in_progress":
                reviewer_item.setBackground(QColor("#fff3cd"))
            elif review_status == "done":
                reviewer_item.setBackground(QColor("#d4edda"))
            self.export_table.setItem(row_idx, 6, reviewer_item)

            btn = QPushButton()
            if status == "free":
                btn.setText("Взять")
                btn.clicked.connect(lambda _, i=id_: self.take_export_issue(i))
            elif status == "in_progress" and taken_by == CURRENT_USER:
                btn.setText("Готово")
                btn.clicked.connect(lambda _, i=id_: self.confirm_complete(i, is_export=True))
            else:
                btn.setText("—")
                btn.setEnabled(False)

            self.export_table.setCellWidget(row_idx, 5, btn)
            self.export_table.setCellWidget(
                row_idx,
                7,
                self.get_review_action_button(id_, status, review_status, review_taken_by, is_export=True)
            )
            self.set_review_status_cell(self.export_table, row_idx, id_, status, review_status, is_export=True)
            self.set_completed_at_cell(self.export_table, row_idx, id_, completed_at, status)

        self.export_table.blockSignals(False)
        self.export_table.setUpdatesEnabled(True)
        fit_table_columns_to_content(self.export_table, ISSUES_TEXT_FIT_COLUMNS)

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

        if item.column() == ISSUES_TABLE_COLUMN_COMPLETED_AT:
            self.save_completed_at(issue_id, item.text(), is_export=True)
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
                    review_status='free',
                    review_taken_by=NULL,
                    review_taken_at=NULL,
                    review_completed_at=NULL,
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
                    review_status='free',
                    review_taken_by=NULL,
                    review_taken_at=NULL,
                    review_completed_at=NULL,
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

        cur.execute("""
            SELECT COUNT(*) FROM export_issues
            WHERE taken_by=? AND status='in_progress'
        """, (CURRENT_USER,))
        if cur.fetchone()[0] > 0:
            conn.close()
            QMessageBox.warning(
                self,
                "Внимание",
                "Вы уже взяли одну запись в работу.\nСначала завершите её."
            )
            return

        cur.execute("""
            UPDATE export_issues
            SET status='in_progress',
                taken_by=?,
                taken_at=?,
                updated_at=?
            WHERE id=? AND status='free'
        """, (CURRENT_USER, datetime.now().isoformat(), datetime.now().isoformat(), issue_id))

        if cur.rowcount == 0:
            cur.execute("SELECT status, taken_by FROM export_issues WHERE id=?", (issue_id,))
            row = cur.fetchone()
            conn.close()

            if row:
                status, taken_by = row
                if status == "in_progress" and taken_by and taken_by != CURRENT_USER:
                    self.warn_issue_locked(taken_by)
                else:
                    QMessageBox.warning(self, "Внимание", "Эта запись уже в работе.")
            else:
                QMessageBox.warning(self, "Ошибка", "Запись не найдена.")
            return

        conn.commit()
        conn.close()

        self.refresh_export_issue_row(issue_id)
        self.remember_db_signature()

    def complete_export_issue(self, issue_id):
        self.confirm_complete(issue_id, is_export=True)

    def _do_complete_export_issue(self, issue_id):
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
            elif key in {"reviewstatus", "статуспроверки", "проверка"}:
                column_map[column] = "review_status"
            elif key in {"reviewtakenby", "ктопроверил", "проверяющий"}:
                column_map[column] = "review_taken_by"
            elif key in {"reviewtakenat", "датавзятиянапроверку", "взятона проверку", "взятонапроверку"}:
                column_map[column] = "review_taken_at"
            elif key in {"reviewcompletedat", "датапроверки", "проверено"}:
                column_map[column] = "review_completed_at"

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

        def normalize_review_status(value):
            key = normalize_name(value)
            status_map = {
                "free": "free",
                "свободен": "free",
                "свободно": "free",
                "inprogress": "in_progress",
                "напроверке": "in_progress",
                "взятонапроверку": "in_progress",
                "done": "done",
                "проверено": "done",
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
            review_status = normalize_review_status(row.get("review_status", "free"))
            review_taken_by = cell_to_text(row.get("review_taken_by", ""))
            review_taken_at = cell_to_text(row.get("review_taken_at", ""))
            review_completed_at = cell_to_text(row.get("review_completed_at", ""))

            cur.execute("""
                INSERT OR IGNORE INTO export_issues
                    (journal, year, issue, status, taken_by, taken_at, completed_at,
                     review_status, review_taken_by, review_taken_at, review_completed_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                journal,
                year,
                issue,
                status,
                taken_by or None,
                taken_at or None,
                completed_at or None,
                review_status,
                review_taken_by or None,
                review_taken_at or None,
                review_completed_at or None,
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
        review_filter_value = (
            self.review_filter_box.currentText()
            if hasattr(self, "review_filter_box")
            else "Все"
        )
        journal_filter = self.journal_filter_box.currentText()
        search_text = self.search_box.text().strip().casefold()
        completed_at_filter_date = parse_date_filter_input(self.completed_at_filter_box.text())

        query = """
            SELECT id, journal, year, issue, path, status, taken_by, taken_at, completed_at,
                   review_status, review_taken_by, review_taken_at, review_completed_at
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

        if review_filter_value == "Свободные":
            query += " AND status = 'done' AND COALESCE(review_status, 'free') = 'free'"
        elif review_filter_value == "В работе":
            query += " AND status = 'done' AND review_status = 'in_progress'"
        elif review_filter_value == "Готово":
            query += " AND status = 'done' AND review_status = 'done'"
        elif review_filter_value == "Мои":
            query += " AND status = 'done' AND review_taken_by = ?"
            params.append(CURRENT_USER)

        if journal_filter != "Все журналы":
            query += " AND journal = ?"
            params.append(journal_filter)

        query, params = append_completed_at_date_sql(query, params, completed_at_filter_date)

        query += " ORDER BY journal COLLATE NOCASE, year, issue"

        cur.execute(query, params)
        rows = cur.fetchall()
        conn.close()

        if search_text:
            rows = filter_issue_rows(
                rows,
                search_text,
                with_path=True,
            )

        # Кэшируем статусные тексты для скорости
        status_texts = {"free": "⚪ Свободен", "in_progress": "🟡 В работе", "done": "✅ Готово"}
        status_bg = {"in_progress": QColor("#fff3cd"), "done": QColor("#d4edda")}

        self.table.setRowCount(0)
        self.table.setUpdatesEnabled(False)
        self.table.blockSignals(True)

        for row_idx, row_data in enumerate(rows):
            (
                id_,
                journal,
                year,
                issue,
                path,
                status,
                taken_by,
                taken_at,
                completed_at,
                review_status,
                review_taken_by,
                review_taken_at,
                review_completed_at,
            ) = row_data

            self.table.insertRow(row_idx)

            journal_item = QTableWidgetItem(journal)
            journal_item.setData(Qt.UserRole, id_)

            year_item = QTableWidgetItem(year)
            year_item.setData(Qt.UserRole, id_)

            issue_item = QTableWidgetItem(issue)
            issue_item.setData(Qt.UserRole, id_)

            who_item = QTableWidgetItem(taken_by or "")
            who_item.setData(Qt.UserRole, id_)
            reviewer_item = QTableWidgetItem(review_taken_by or "")
            reviewer_item.setData(Qt.UserRole, id_)

            if self.is_admin:
                journal_item.setFlags(journal_item.flags() | Qt.ItemIsEditable)
                year_item.setFlags(year_item.flags() | Qt.ItemIsEditable)
                issue_item.setFlags(issue_item.flags() | Qt.ItemIsEditable)
                who_item.setFlags(who_item.flags() | Qt.ItemIsEditable)

                status_combo = NoWheelComboBox()
                status_combo.addItem("⚪ Свободен", "free")
                status_combo.addItem("🟡 В работе", "in_progress")
                status_combo.addItem("✅ Готово", "done")
                status_combo.setCurrentIndex(status_combo.findData(status))
                status_combo.currentIndexChanged.connect(
                    lambda _, issue_id=id_, combo=status_combo: self.change_status(issue_id, combo.currentData())
                )
                self.table.setCellWidget(row_idx, 3, status_combo)
            else:
                status_item = QTableWidgetItem(status_texts.get(status, status))
                bg = status_bg.get(status)
                if bg:
                    status_item.setBackground(bg)
                self.table.setItem(row_idx, 3, status_item)

            self.table.setItem(row_idx, 0, journal_item)
            self.table.setItem(row_idx, 1, year_item)
            self.table.setItem(row_idx, 2, issue_item)
            self.table.setItem(row_idx, 4, who_item)
            if review_status == "in_progress":
                reviewer_item.setBackground(QColor("#fff3cd"))
            elif review_status == "done":
                reviewer_item.setBackground(QColor("#d4edda"))
            self.table.setItem(row_idx, 6, reviewer_item)

            btn = QPushButton()
            if status == "free":
                btn.setText("Взять")
                btn.clicked.connect(lambda _, i=id_: self.take_issue(i))
            elif status == "in_progress" and taken_by == CURRENT_USER:
                btn.setText("Готово")
                btn.clicked.connect(lambda _, i=id_: self.confirm_complete(i, is_export=False))
            else:
                btn.setText("—")
                btn.setEnabled(False)

            self.table.setCellWidget(row_idx, 5, btn)
            self.table.setCellWidget(
                row_idx,
                7,
                self.get_review_action_button(id_, status, review_status, review_taken_by)
            )
            self.set_review_status_cell(self.table, row_idx, id_, status, review_status)
            self.set_completed_at_cell(self.table, row_idx, id_, completed_at, status)

        self.table.blockSignals(False)
        self.table.setUpdatesEnabled(True)
        fit_table_columns_to_content(self.table, ISSUES_TEXT_FIT_COLUMNS)

        if self._preserve_scroll_on_reload:
            self.table.verticalScrollBar().setValue(scroll_value)

    def on_table_item_changed(self, item):
        if not self.is_admin:
            return

        issue_id = item.data(Qt.UserRole)
        if issue_id is None:
            return

        if item.column() == ISSUES_TABLE_COLUMN_COMPLETED_AT:
            self.save_completed_at(issue_id, item.text(), is_export=False)
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
                    review_status='free',
                    review_taken_by=NULL,
                    review_taken_at=NULL,
                    review_completed_at=NULL,
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
                    review_status='free',
                    review_taken_by=NULL,
                    review_taken_at=NULL,
                    review_completed_at=NULL,
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

        # Проверяем, нет ли уже журнала в работе у этого пользователя
        cur.execute("""
            SELECT COUNT(*) FROM issues
            WHERE taken_by=? AND status='in_progress'
        """, (CURRENT_USER,))
        if cur.fetchone()[0] > 0:
            conn.close()
            QMessageBox.warning(
                self,
                "Внимание",
                "Вы уже взяли один журнал в работу.\nСначала завершите его."
            )
            return

        # Атомарно пытаемся захватить. UPDATE в SQLite атомарен сам по себе.
        cur.execute("""
            UPDATE issues
            SET status='in_progress',
                taken_by=?,
                taken_at=?,
                updated_at=?
            WHERE id=? AND status='free'
        """, (CURRENT_USER, datetime.now().isoformat(), datetime.now().isoformat(), issue_id))

        if cur.rowcount == 0:
            # Не удалось — выясняем причину для понятного сообщения
            cur.execute("SELECT status, taken_by FROM issues WHERE id=?", (issue_id,))
            row = cur.fetchone()
            conn.close()

            if row:
                status, taken_by = row
                if status == "in_progress" and taken_by and taken_by != CURRENT_USER:
                    self.warn_issue_locked(taken_by)
                else:
                    QMessageBox.warning(self, "Внимание", "Этот журнал уже в работе.")
            else:
                QMessageBox.warning(self, "Ошибка", "Запись не найдена.")
            return

        conn.commit()
        conn.close()

        self.refresh_issue_row(issue_id)
        self.remember_db_signature()

    def confirm_complete(self, issue_id, is_export=False):
        """Защита от случайного нажатия «Готово»"""
        reply = QMessageBox.question(
            self,
            "Подтверждение",
            "Вы загрузили все статьи выпуска данного журнала в WindChill?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            if is_export:
                self._do_complete_export_issue(issue_id)
            else:
                self._do_complete_issue(issue_id)

    def _do_complete_issue(self, issue_id):
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

    def complete_issue(self, issue_id):
        self.confirm_complete(issue_id, is_export=False)

    def export_report(self):
        conn = get_db_connection()
        column_sql = ", ".join(column for column, _ in ISSUES_REPORT_COLUMNS)
        df = pd.read_sql_query(
            f"""
            SELECT {column_sql}
            FROM issues
            WHERE status='done' AND LOWER(path) LIKE '%.pdf%'
            ORDER BY journal COLLATE NOCASE, year, issue
            """,
            conn,
        )
        conn.close()

        if df.empty:
            QMessageBox.information(
                self,
                "Экспорт отчёта",
                "Нет готовых PDF-записей для экспорта."
            )
            return

        df["status"] = df["status"].map(lambda value: ISSUE_STATUS_TEXTS.get(value, value))
        df["review_status"] = df["review_status"].map(
            lambda value: self.get_review_status_text(value, "done")
        )
        df["completed_at"] = df["completed_at"].map(format_report_datetime)

        df = df.rename(columns={column: title for column, title in ISSUES_REPORT_COLUMNS})

        file_path, _ = QFileDialog.getSaveFileName(self, "Сохранить отчет", "", "Excel (*.xlsx)")
        if file_path:
            df.to_excel(file_path, index=False)

    def export_export_tab_to_excel(self):
        """Экспорт записей вкладки «Скачать из Elibrary» в Excel (с учётом текущих фильтров)."""
        filter_value = self.export_filter_box.currentText()
        review_filter_value = (
            self.export_review_filter_box.currentText()
            if hasattr(self, "export_review_filter_box")
            else "Все"
        )
        journal_filter = self.export_journal_filter_box.currentText()
        search_text = self.export_search_box.text().strip().casefold()
        completed_at_filter_date = parse_date_filter_input(self.export_completed_at_filter_box.text())

        column_sql = ", ".join(column for column, _ in ISSUES_REPORT_COLUMNS)
        query = f"""
            SELECT {column_sql}
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

        if review_filter_value == "Свободные":
            query += " AND status = 'done' AND COALESCE(review_status, 'free') = 'free'"
        elif review_filter_value == "В работе":
            query += " AND status = 'done' AND review_status = 'in_progress'"
        elif review_filter_value == "Готово":
            query += " AND status = 'done' AND review_status = 'done'"
        elif review_filter_value == "Мои":
            query += " AND status = 'done' AND review_taken_by = ?"
            params.append(CURRENT_USER)

        if journal_filter != "Все журналы":
            query += " AND journal = ?"
            params.append(journal_filter)

        query, params = append_completed_at_date_sql(query, params, completed_at_filter_date)

        query += " ORDER BY journal COLLATE NOCASE, year, issue"

        conn = get_db_connection()
        df = pd.read_sql_query(query, conn, params=params)
        conn.close()

        if search_text and not df.empty:
            mask = df.apply(
                lambda row: issue_row_matches_search(
                    search_text,
                    journal=row.get("journal"),
                    year=row.get("year"),
                    issue=row.get("issue"),
                    status=row.get("status"),
                    taken_by=row.get("taken_by"),
                    review_status=row.get("review_status"),
                    review_taken_by=row.get("review_taken_by"),
                ),
                axis=1,
            )
            df = df[mask]

        if df.empty:
            QMessageBox.information(
                self,
                "Экспорт в Excel",
                "Нет записей для экспорта по текущим фильтрам."
            )
            return

        df["status"] = df["status"].map(lambda value: ISSUE_STATUS_TEXTS.get(value, value))
        df["review_status"] = df["review_status"].map(
            lambda value: self.get_review_status_text(value, "done")
        )
        df["completed_at"] = df["completed_at"].map(format_report_datetime)

        df = df.rename(columns={column: title for column, title in ISSUES_REPORT_COLUMNS})

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить Excel",
            "export_elibrary.xlsx",
            "Excel (*.xlsx)",
        )
        if not file_path:
            return

        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"

        try:
            df.to_excel(file_path, index=False)
        except Exception as exc:
            QMessageBox.critical(
                self,
                "Ошибка экспорта",
                f"Не удалось сохранить Excel-файл.\n\n{exc}"
            )
            return

        QMessageBox.information(
            self,
            "Экспорт в Excel",
            f"Сохранено записей: {len(df)}\n\n{file_path}"
        )

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
        self.update_rating()  # сразу показываем изменения

        QMessageBox.information(
            self,
            "Сохранено",
            f"План на {month} сохранён: {plan_count}"
        )

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

        journal_title = QLabel("Статистика по журналам")
        journal_title.setStyleSheet("font-weight: 700; font-size: 14px; color: #1f2937;")

        self.journal_progress_table = QTableWidget()
        self.journal_progress_summary_label = QLabel()
        self.journal_progress_summary_label.setWordWrap(False)
        self.journal_progress_summary_label.setStyleSheet("font-size: 14px; color: #1f2937;")
        self.journal_progress_table.setColumnCount(7)
        self.journal_progress_table.setHorizontalHeaderLabels([
            "Журнал",
            "Всего, шт",
            "Выгружено, шт",
            "Выгружено, %",
            "На проверку, шт",
            "Проверено, шт",
            "Проверено, %",
        ])
        self.journal_progress_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.journal_progress_table.setAlternatingRowColors(True)
        self.journal_progress_table.verticalHeader().setDefaultSectionSize(28)

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
        apply_static_column_widths(
            self.journal_progress_table,
            RATING_PROGRESS_COLUMN_SAMPLES,
        )

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

        left_content_layout = QVBoxLayout()
        left_content_layout.addWidget(self.journal_frame)
        left_content_layout.addStretch()

        content_layout = QHBoxLayout()
        content_layout.addLayout(left_content_layout, 1)
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
            QLineEdit {
                border: 1px solid #d7dde5;
                border-radius: 6px;
                padding: 4px 8px;
                background-color: #ffffff;
            }
            QComboBox {
                border: 1px solid #d7dde5;
                border-radius: 6px;
                padding: 4px 8px;
                background-color: #ffffff;
            }
            QComboBox::drop-down {
                border: none;
                border-radius: 6px;
            }
        """)

        if self.is_admin:
            self.plan_month_edit.setText(datetime.now().strftime("%Y-%m"))

    def update_rating(self):
        self.is_reviewer = self.user_is_reviewer(CURRENT_USER)

        conn = get_db_connection()
        cur = conn.cursor()

        month = self.get_selected_month()

        cur.execute("SELECT plan_count FROM monthly_plans WHERE month=?", (month,))
        row = cur.fetchone()
        plan_count = row[0] if row else 0

        # Факт за месяц: issues + export_issues
        cur.execute("""
            SELECT (
                SELECT COUNT(*) FROM issues
                WHERE status='done' AND completed_at LIKE ?
            ) + (
                SELECT COUNT(*) FROM export_issues
                WHERE status='done' AND completed_at LIKE ?
            )
        """, (month + "%", month + "%"))
        month_done = cur.fetchone()[0] or 0

        month_remaining = max(plan_count - month_done, 0)
        month_percent = int((month_done / plan_count) * 100) if plan_count else 0

        # Всего записей и готово: суммируем обе таблицы
        cur.execute("SELECT COUNT(*) FROM issues")
        total_issues = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM export_issues")
        total_export = cur.fetchone()[0]

        cur.execute("SELECT COUNT(*) FROM issues WHERE status='done'")
        done_issues = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM export_issues WHERE status='done'")
        done_export = cur.fetchone()[0]

        total_all = total_issues + total_export
        done_all = done_issues + done_export

        overall_percent = int((done_all / total_all) * 100) if total_all else 0

        cur.execute("SELECT COUNT(*) FROM issues WHERE review_status='done'")
        reviewed_issues = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM export_issues WHERE review_status='done'")
        reviewed_export = cur.fetchone()[0]
        reviewed_all = reviewed_issues + reviewed_export
        to_review_all = max(done_all - reviewed_all, 0)
        reviewed_percent = int((reviewed_all / done_all) * 100) if done_all else 0

        # Статистика по журналам: объединяем обе таблицы
        cur.execute("""
            SELECT journal,
                   COUNT(*) AS total_count,
                   SUM(CASE WHEN status='done' THEN 1 ELSE 0 END) AS done_count,
                   SUM(CASE WHEN review_status='done' THEN 1 ELSE 0 END) AS reviewed_count
            FROM (
                SELECT journal, status, review_status FROM issues
                UNION ALL
                SELECT journal, status, review_status FROM export_issues
            )
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
            f"Всего записей: {total_all} | Выгружено: {done_all} | "
            f"Общий прогресс: {overall_percent}%"
        )

        self.journal_progress_summary_label.setText(
            f"Всего: {total_all} | Выгружено: {done_all} ({overall_percent}%) | "
            f"На проверку: {to_review_all} | Проверено: {reviewed_all} ({reviewed_percent}%)"
        )

        self.journal_progress_table.setRowCount(0)
        self.journal_progress_table.setUpdatesEnabled(False)
        self.journal_progress_table.blockSignals(True)

        for journal, total_count, done_count, reviewed_count in journal_rows:
            journal_name = journal or "Без названия"
            exported_percent = int((done_count / total_count) * 100) if total_count else 0
            to_review_count = max(done_count - reviewed_count, 0)
            reviewed_percent_value = int((reviewed_count / done_count) * 100) if done_count else 0

            row = self.journal_progress_table.rowCount()
            self.journal_progress_table.insertRow(row)

            journal_item = QTableWidgetItem(journal_name)
            total_item = QTableWidgetItem(str(total_count))
            done_item = QTableWidgetItem(str(done_count))
            exported_percent_item = QTableWidgetItem(f"{exported_percent}%")
            to_review_item = QTableWidgetItem(str(to_review_count))
            reviewed_item = QTableWidgetItem(str(reviewed_count))
            reviewed_percent_item = QTableWidgetItem(f"{reviewed_percent_value}%")

            for item in (
                total_item,
                done_item,
                exported_percent_item,
                to_review_item,
                reviewed_item,
                reviewed_percent_item,
            ):
                item.setTextAlignment(Qt.AlignCenter)

            set_percent_cell_background(exported_percent_item, exported_percent)
            set_percent_cell_background(reviewed_percent_item, reviewed_percent_value)

            self.journal_progress_table.setItem(row, 0, journal_item)
            self.journal_progress_table.setItem(row, 1, total_item)
            self.journal_progress_table.setItem(row, 2, done_item)
            self.journal_progress_table.setItem(row, 3, exported_percent_item)
            self.journal_progress_table.setItem(row, 4, to_review_item)
            self.journal_progress_table.setItem(row, 5, reviewed_item)
            self.journal_progress_table.setItem(row, 6, reviewed_percent_item)

        self.journal_progress_table.blockSignals(False)
        self.journal_progress_table.setUpdatesEnabled(True)

        rows_count = max(self.journal_progress_table.rowCount(), 1)
        row_height = self.journal_progress_table.verticalHeader().defaultSectionSize()
        header_height = self.journal_progress_table.horizontalHeader().height()
        table_height = header_height + rows_count * row_height + 8
        self.journal_progress_table.setMinimumHeight(table_height)
        self.journal_frame.setMinimumHeight(table_height + 110)

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

    def closeEvent(self, event):
        """Корректно закрываем соединение при выходе"""
        self.refresh_timer.stop()
        self.update_check_timer.stop()
        # Даём время на завершение операций
        QTimer.singleShot(100, lambda: None)
        super().closeEvent(event)


# ================= ЗАПУСК =================

if __name__ == "__main__":
    # Повторные попытки инициализации БД при высокой нагрузке
    max_retries = 5
    for attempt in range(max_retries):
        try:
            init_db()
            break
        except sqlite3.OperationalError as exc:
            if attempt == max_retries - 1:
                app = QApplication(sys.argv)
                QMessageBox.critical(
                    None,
                    "Ошибка запуска",
                    f"Не удалось открыть общую базу данных после {max_retries} попыток.\n\n"
                    f"Путь: {DB_PATH}\n\n"
                    f"1. Проверьте, что файловый сервер доступен.\n"
                    f"2. Проверьте права на запись в папку.\n"
                    f"3. Попробуйте перезапустить программу.\n\n"
                    f"Техническая ошибка:\n{exc}"
                )
                sys.exit(1)
            import time
            time.sleep(1)  # ждём секунду перед повтором

    # Отключаем следование тёмной теме Windows
    app = QApplication(sys.argv)
    app.setStyle("windowsvista")

    window = MainWindow()
    window.show()
    sys.exit(app.exec())

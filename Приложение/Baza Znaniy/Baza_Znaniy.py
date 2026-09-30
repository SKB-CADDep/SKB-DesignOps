import sys
import json
import re
import html
import sqlite3
import getpass
from datetime import datetime
from pathlib import Path
from urllib.parse import quote, unquote

from PySide6.QtCore import Qt, QTimer, QUrl
from PySide6.QtGui import QDesktopServices, QFont, QPixmap, QPalette, QColor
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QPushButton, QTreeWidget, QTreeWidgetItem, QSplitter, QTabWidget,
    QScrollArea, QFrame, QToolButton, QListWidget, QListWidgetItem, QTextEdit,
    QTextBrowser, QCheckBox, QComboBox, QMessageBox, QSpinBox, QGroupBox,
    QSizePolicy, QAbstractItemView, QInputDialog, QDialog, QDialogButtonBox,
    QFormLayout, QPlainTextEdit, QGraphicsDropShadowEffect, QFileDialog
)

# Определяем путь к базе данных рядом с EXE-файлом
import sys

if getattr(sys, 'frozen', False):
    # PyInstaller-сборка: база рядом с .exe
    BASE_DIR = Path(sys.executable).parent
else:
    # Обычный запуск скрипта: база рядом со скриптом
    BASE_DIR = Path(__file__).parent
DB_PATH = BASE_DIR / "knowledge_base.db"

ROLE_KIND = Qt.UserRole + 1
ROLE_ID = Qt.UserRole + 2
ADMIN_USERNAME = "nyagavrilova"

APP_QSS = """
QMainWindow {
    background: #f7f9fd;
    color: #263238;
}

QWidget {
    background: #f7f9fd;
    color: #263238;
}

QTabWidget::pane {
    border: 1px solid #d9e3f0;
    background: #fcfdff;
    border-radius: 10px;
}

QTabBar::tab {
    background: #edf4fb;
    color: #263238;
    padding: 9px 14px;
    border: 1px solid #d9e3f0;
    border-bottom: none;
    border-top-left-radius: 8px;
    border-top-right-radius: 8px;
    margin-right: 2px;
}

QTabBar::tab:selected {
    background: #ffffff;
}

QLineEdit, QTextEdit, QTextBrowser, QTreeWidget, QListWidget, QComboBox, QSpinBox {
    background: #ffffff;
    color: #263238;
    border: 1px solid #d9e3f0;
    border-radius: 8px;
    padding: 7px;
    selection-background-color: #dbeafe;
    selection-color: #263238;
}

QLineEdit:focus, QTextEdit:focus, QTextBrowser:focus, QComboBox:focus, QSpinBox:focus {
    border: 1px solid #b9cde8;
}

QComboBox QAbstractItemView {
    background: #ffffff;
    color: #263238;
    selection-background-color: #dbeafe;
    selection-color: #263238;
    border: 1px solid #d9e3f0;
    outline: 0;
}

QComboBox::drop-down {
    border: none;
    background: transparent;
}

QPushButton {
    padding: 8px 12px;
    border-radius: 8px;
    background: #dceeff;
    color: #263238;
    border: 1px solid #c3d8f2;
}

QPushButton:hover {
    background: #cfe6ff;
}

QPushButton:pressed {
    background: #bfdcff;
}

QPushButton:disabled {
    background: #eef4fb;
    color: #90a4b8;
    border: 1px solid #d9e3f0;
}

QTreeWidget {
    background: #ffffff;
    alternate-background-color: #f7fafc;
}

QListWidget {
    background: #ffffff;
    alternate-background-color: #f7fafc;
}

QListWidget::item {
    background: #ffffff;
    border: 1px solid #e1e8f0;
    border-radius: 8px;
    margin: 3px 4px;
    padding: 6px 8px;
}

QListWidget::item:selected,
QTreeWidget::item:selected {
    background: #dbeafe;
    color: #263238;
}

QHeaderView::section {
    background: #eef5fb;
    color: #263238;
    padding: 6px 8px;
    border: 1px solid #d9e3f0;
    font-weight: 600;
}

QGroupBox {
    font-weight: 600;
    border: 1px solid #d9e3f0;
    border-radius: 8px;
    margin-top: 10px;
    padding-top: 10px;
    background: #ffffff;
}

QGroupBox::title {
    subcontrol-origin: margin;
    left: 12px;
    padding: 0 4px;
    color: #263238;
}

QScrollBar:vertical {
    background: #f4f7fb;
    width: 12px;
    margin: 0px;
}

QScrollBar::handle:vertical {
    background: #c9d9ea;
    min-height: 24px;
    border-radius: 5px;
}

QScrollBar::handle:vertical:hover {
    background: #b6c9e0;
}

QStatusBar {
    background: #f7f9fd;
    color: #263238;
}

QListWidget::indicator:unchecked {
    border: 2px solid #C9D9EA;
    background: white;
    border-radius: 4px;
}
QListWidget::indicator:checked {
    border: 2px solid #C9D9EA;
    background: #C9D9EA;
    border-radius: 4px;
}

* {
    font-size: 13px;
}
QGroupBox::title {
    font-size: 13px;
}
QTabBar::tab {
    font-size: 13px;
}
QTreeWidget {
    font-size: 12px;
}

/* Синие ссылки во всех rich‑text виджетах */
a {
    color: #0000FF;
}
"""


def ts():
    return datetime.now().strftime("%d.%m.%Y %H:%M")


def apply_light_palette(app):
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor("#f7f9fd"))
    palette.setColor(QPalette.WindowText, QColor("#263238"))
    palette.setColor(QPalette.Base, QColor("#ffffff"))
    palette.setColor(QPalette.AlternateBase, QColor("#f7fafc"))
    palette.setColor(QPalette.Text, QColor("#263238"))
    palette.setColor(QPalette.Button, QColor("#dceeff"))
    palette.setColor(QPalette.ButtonText, QColor("#263238"))
    palette.setColor(QPalette.Highlight, QColor("#dbeafe"))
    palette.setColor(QPalette.HighlightedText, QColor("#263238"))
    palette.setColor(QPalette.ToolTipBase, QColor("#ffffff"))
    palette.setColor(QPalette.ToolTipText, QColor("#263238"))
    app.setPalette(palette)


def p(text):
    return f"<p>{html.escape(str(text))}</p>"


def html_list(items):
    return "<ul>" + "".join(f"<li>{html.escape(str(item))}</li>" for item in items) + "</ul>"


def internal_link(title, text=None):
    label = html.escape(text or title)
    return f"<a href='instruction://{quote(title)}'>{label}</a>"


def escape_html(text):
    return html.escape(str(text)).replace("\n", "<br>")


def render_markdown(text):
    """
    Преобразует Markdown-подобный текст в HTML.
    Поддерживаемые стили GitHub:
    - # Заголовок 1
    - ## Заголовок 2
    - ### Заголовок 3
    - #### Заголовок 4
    - **жирный**
    - *курсив*
    - ~~зачёркнутый~~
    - `код`
    - ```многострочный код```
    - - элемент списка
    - 1. нумерованный список
    - > цитата
    - --- горизонтальная линия
    - [текст](ссылка)
    - обычный текст с переносами строк
    """
    if not text:
        return ""

    lines = text.split("\n")
    result = []
    in_code_block = False
    code_lines = []
    in_paragraph = False
    paragraph_lines = []

    def flush_paragraph():
        nonlocal in_paragraph, paragraph_lines
        if paragraph_lines:
            result.append("<p>" + "<br>".join(paragraph_lines) + "</p>")
            paragraph_lines = []
            in_paragraph = False

    def process_inline(text):
        """Обработка inline-элементов: жирный, курсив, код, зачёркнутый, ссылки."""
        # Экранируем HTML
        text = html.escape(text)

        # ```код``` -> <code>код</code>
        text = re.sub(r'```(.+?)```', r'<pre><code>\1</code></pre>', text)

        # `код` -> <code>код</code>
        text = re.sub(r'`([^`]+)`', r'<code>\1</code>', text)

        # **жирный**
        text = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', text)

        # *курсив* (но не **)
        text = re.sub(r'(?<!\*)\*([^*\n]+)\*(?!\*)', r'<i>\1</i>', text)

        # ~~зачёркнутый~~
        text = re.sub(r'~~(.+?)~~', r'<s>\1</s>', text)

        # [текст](ссылка) – синие ссылки
        text = re.sub(r'\[([^\]]+)\]\(([^)]+)\)', r'<a href="\2" style="color: #0000FF;">\1</a>', text)

        return text

    i = 0
    while i < len(lines):
        line = lines[i]

        # Многострочный код
        if line.strip().startswith("```"):
            if not in_code_block:
                flush_paragraph()
                in_code_block = True
                code_lines = []
            else:
                # Завершаем код
                result.append("<pre><code>" + "\n".join(code_lines) + "</code></pre>")
                in_code_block = False
                code_lines = []
            i += 1
            continue

        if in_code_block:
            code_lines.append(line)
            i += 1
            continue

        stripped = line.strip()

        # Горизонтальная линия
        if re.match(r'^[-*_]{3,}$', stripped):
            flush_paragraph()
            result.append("<hr>")
            i += 1
            continue

        # Заголовки
        if stripped.startswith("#### "):
            flush_paragraph()
            result.append("<h4>" + process_inline(stripped[5:]) + "</h4>")
            i += 1
            continue
        if stripped.startswith("### "):
            flush_paragraph()
            result.append("<h3>" + process_inline(stripped[4:]) + "</h3>")
            i += 1
            continue
        if stripped.startswith("## "):
            flush_paragraph()
            result.append("<h2>" + process_inline(stripped[3:]) + "</h2>")
            i += 1
            continue
        if stripped.startswith("# "):
            flush_paragraph()
            result.append("<h1>" + process_inline(stripped[2:]) + "</h1>")
            i += 1
            continue

        # Цитата
        if stripped.startswith("> "):
            flush_paragraph()
            result.append("<blockquote>" + process_inline(stripped[2:]) + "</blockquote>")
            i += 1
            continue

        # Ненумерованный список
        if re.match(r'^[\-\*\+]\s', stripped):
            flush_paragraph()
            list_items = []
            while i < len(lines) and re.match(r'^[\-\*\+]\s', lines[i].strip()):
                list_items.append("<li>" + process_inline(re.sub(r'^[\-\*\+]\s+', '', lines[i].strip())) + "</li>")
                i += 1
            result.append("<ul>" + "".join(list_items) + "</ul>")
            continue

        # Нумерованный список
        if re.match(r'^\d+\.\s', stripped):
            flush_paragraph()
            list_items = []
            while i < len(lines) and re.match(r'^\d+\.\s', lines[i].strip()):
                list_items.append("<li>" + process_inline(re.sub(r'^\d+\.\s+', '', lines[i].strip())) + "</li>")
                i += 1
            result.append("<ol>" + "".join(list_items) + "</ol>")
            continue

        # Пустая строка — завершаем параграф
        if not stripped:
            flush_paragraph()
            i += 1
            continue

        # Обычный текст — накапливаем в параграф
        in_paragraph = True
        paragraph_lines.append(process_inline(stripped))
        i += 1

    flush_paragraph()

    return "\n".join(result)


def strip_html_tags(text):
    text = re.sub(r"<[^>]+>", " ", str(text))
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def split_non_empty_lines(text):
    return [line.strip() for line in str(text).splitlines() if line.strip()]


def split_titles(text):
    parts = re.split(r"[,\n;]+", str(text))
    return [part.strip() for part in parts if part.strip()]


def short_button_text(text, limit=48):
    text = str(text).strip()
    return text if len(text) <= limit else text[:limit - 1].rstrip() + "…"


def parse_manual_sections(text):
    raw = str(text).strip()
    if not raw:
        return []

    blocks = re.split(r"\n\s*\n", raw)
    result = []

    for block in blocks:
        lines = [line.rstrip() for line in block.splitlines() if line.strip()]
        if not lines:
            continue

        title = lines[0].strip()
        body = "\n".join(lines[1:]).strip()
        body_html = escape_html(body) if body else "<p></p>"
        result.append(section(title, body_html))

    return result


def section(title, body_html, image_path=None):
    return {
        "title": title,
        "body": body_html,
        "image_path": image_path or ""
    }


def make_sections(intro, steps, note_html):
    return [
        section("Коротко", p(intro)),
        section("Порядок действий", html_list(steps)),
        section("Важно", note_html),
    ]


def make_card(category, task_title, instruction_title, short_desc, checklist, steps, note_html, related):
    return {
        "category": category,
        "task_title": task_title,
        "instruction_title": instruction_title,
        "short_desc": short_desc,
        "checklist": checklist,
        "sections": make_sections(short_desc, steps, note_html),
        "related": related,
    }


# ================== ДЕМО-ДАННЫЕ ==================
# TODO: сюда потом можно вставить твои реальные 12 инструкций из Word.

def template_card(category, task_title, instruction_title, related=None):
    return make_card(
        category=category,
        task_title=task_title,
        instruction_title=instruction_title,
        short_desc=f"ВСТАВЬ СЮДА КРАТКОЕ ОПИСАНИЕ ИЗ WORD ДЛЯ «{instruction_title}».",
        checklist=[
            "Открыть нужную инструкцию",
            "Выполнить шаги из инструкции",
            "Проверить результат и, если нужно, оставить комментарий",
        ],
        steps=[
            "ВСТАВЬ СЮДА ОСНОВНОЙ ТЕКСТ / ПЕРВУЮ ГЛАВУ ИЗ WORD.",
            "ВСТАВЬ СЮДА СЛЕДУЮЩИЙ РАЗДЕЛ ИНСТРУКЦИИ.",
            "ВСТАВЬ СЮДА ЗАКЛЮЧЕНИЕ, ПРИМЕЧАНИЯ И ОСОБЫЕ СЛУЧАИ.",
        ],
        note_html=f"<p><b>Важно:</b> сюда можно вставить ссылки на связанные инструкции и предупреждения для «{instruction_title}».</p>",
        related=related or [],
    )


SAMPLE_CARDS = [
    template_card(
        category="PDF и документы",
        task_title="Извлечь содержимое PDF-файлов",
        instruction_title="Инструкция по работе с программой для извлечения содержимого PDF-файлов - PDF Contents Extractor",
        related=[
            "Инструкция по подготовке документов для выгрузки в WindChill",
            "Инструкция по разбору документов для выгрузки в WindChill",
            "Инструкция по работе с программой split_spreads_ui по разрезанию разворотов в книгах",
        ],
    ),
    template_card(
        category="WindChill",
        task_title="Загрузить документы в WindChill",
        instruction_title="Инструкция по загрузке документов в WindChill",
        related=[
            "Инструкция по подготовке документов для выгрузки в WindChill",
            "Инструкция по разбору документов для выгрузки в WindChill",
        ],
    ),
    template_card(
        category="Утилизация",
        task_title="Работа со шредером",
        instruction_title="Инструкция для шредера",
        related=[
            "Инструкция по обработке макулатуры",
        ],
    ),
    template_card(
        category="Утилизация",
        task_title="Обработать макулатуру",
        instruction_title="Инструкция по обработке макулатуры",
        related=[
            "Инструкция для шредера",
        ],
    ),
    template_card(
        category="WindChill",
        task_title="Подготовить документы для выгрузки в WindChill",
        instruction_title="Инструкция по подготовке документов для выгрузки в WindChill",
        related=[
            "Инструкция по разбору документов для выгрузки в WindChill",
            "Инструкция по загрузке документов в WindChill",
        ],
    ),
    template_card(
        category="WindChill",
        task_title="Разобрать документы для выгрузки в WindChill",
        instruction_title="Инструкция по разбору документов для выгрузки в WindChill",
        related=[
            "Инструкция по подготовке документов для выгрузки в WindChill",
            "Инструкция по загрузке документов в WindChill",
        ],
    ),
    template_card(
        category="PDF и документы",
        task_title="Разрезать развороты в книгах",
        instruction_title="Инструкция по работе с программой split_spreads_ui по разрезанию разворотов в книгах",
        related=[
            "Инструкция по работе с программой для извлечения содержимого PDF-файлов - PDF Contents Extractor",
        ],
    ),
    template_card(
        category="Журналы",
        task_title="Смотреть прогресс по журналам",
        instruction_title="Инструкция по работе с приложением прогресса",
        related=[
            "Инструкция по загрузке статей в научную библиотеку УТЗ в WNC",
            "Инструкция по разделению закладок на отдельные файлы",
            "Инструкция по скачиванию статей с сайта Elibrary",
            "Инструкция по созданию закладок в PDF файле для последующего разъединения статей на отдельные файлы",
        ],
    ),
    template_card(
        category="Журналы",
        task_title="Загрузить статьи в научную библиотеку УТЗ",
        instruction_title="Инструкция по загрузке статей в научную библиотеку УТЗ в WNC",
        related=[
            "Инструкция по скачиванию статей с сайта Elibrary",
            "Инструкция по созданию закладок в PDF файле для последующего разъединения статей на отдельные файлы",
            "Инструкция по разделению закладок на отдельные файлы",
        ],
    ),
    template_card(
        category="Журналы",
        task_title="Разделить закладки на отдельные файлы",
        instruction_title="Инструкция по разделению закладок на отдельные файлы",
        related=[
            "Инструкция по созданию закладок в PDF файле для последующего разъединения статей на отдельные файлы",
            "Инструкция по загрузке статей в научную библиотеку УТЗ в WNC",
        ],
    ),
    template_card(
        category="Журналы",
        task_title="Скачать статьи с eLibrary",
        instruction_title="Инструкция по скачиванию статей с сайта Elibrary",
        related=[
            "Инструкция по созданию закладок в PDF файле для последующего разъединения статей на отдельные файлы",
        ],
    ),
    template_card(
        category="Журналы",
        task_title="Создать закладки в PDF",
        instruction_title="Инструкция по созданию закладок в PDF файле для последующего разъединения статей на отдельные файлы",
        related=[
            "Инструкция по разделению закладок на отдельные файлы",
            "Инструкция по скачиванию статей с сайта Elibrary",
            "Инструкция по загрузке статей в научную библиотеку УТЗ в WNC",
        ],
    ),
]


# ================== БАЗА ДАННЫХ ==================

class KnowledgeBaseDB:
    def __init__(self, path):
        self.path = Path(path)
        self.conn = sqlite3.connect(str(self.path), timeout=10)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON")
        self.init_schema()
        self.migrate_ratings_schema()
        self.seed_demo_data()

    def close(self):
        try:
            self.conn.close()
        except Exception:
            pass

    def init_schema(self):
        cur = self.conn.cursor()

        cur.execute("""
            CREATE TABLE IF NOT EXISTS categories (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                sort_order INTEGER NOT NULL DEFAULT 0
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS instructions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                category_id INTEGER NOT NULL,
                title TEXT NOT NULL UNIQUE,
                short_desc TEXT NOT NULL DEFAULT '',
                sections_json TEXT NOT NULL DEFAULT '[]',
                related_ids_json TEXT NOT NULL DEFAULT '[]',
                updated_at TEXT NOT NULL,
                FOREIGN KEY(category_id) REFERENCES categories(id) ON DELETE CASCADE
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS tasks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                category_id INTEGER NOT NULL,
                title TEXT NOT NULL UNIQUE,
                instruction_id INTEGER,
                checklist_json TEXT NOT NULL DEFAULT '[]',
                sort_order INTEGER NOT NULL DEFAULT 0,
                FOREIGN KEY(category_id) REFERENCES categories(id) ON DELETE CASCADE,
                FOREIGN KEY(instruction_id) REFERENCES instructions(id) ON DELETE SET NULL
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS comments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                instruction_id INTEGER NOT NULL,
                anchor TEXT NOT NULL DEFAULT '',
                author TEXT NOT NULL,
                is_anonymous INTEGER NOT NULL DEFAULT 1,
                text TEXT NOT NULL,
                created_at TEXT NOT NULL,
                FOREIGN KEY(instruction_id) REFERENCES instructions(id) ON DELETE CASCADE
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS ratings (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                instruction_id INTEGER NOT NULL,
                rating INTEGER NOT NULL CHECK(rating BETWEEN 1 AND 10),
                created_at TEXT NOT NULL,
                FOREIGN KEY(instruction_id) REFERENCES instructions(id) ON DELETE CASCADE
            )
        """)
        self.conn.commit()

    def migrate_ratings_schema(self):
        row = self.conn.execute("""
            SELECT sql
            FROM sqlite_master
            WHERE type='table' AND name='ratings'
        """).fetchone()

        sql = row["sql"] if row and row["sql"] else ""
        if "BETWEEN 1 AND 10" in sql:
            return

        cur = self.conn.cursor()
        cur.execute("PRAGMA foreign_keys = OFF")
        cur.execute("ALTER TABLE ratings RENAME TO ratings_old")
        cur.execute("""
            CREATE TABLE ratings (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                instruction_id INTEGER NOT NULL,
                rating INTEGER NOT NULL CHECK(rating BETWEEN 1 AND 10),
                created_at TEXT NOT NULL,
                FOREIGN KEY(instruction_id) REFERENCES instructions(id) ON DELETE CASCADE
            )
        """)
        cur.execute("""
            INSERT INTO ratings (id, instruction_id, rating, created_at)
            SELECT id, instruction_id, rating, created_at
            FROM ratings_old
        """)
        cur.execute("DROP TABLE ratings_old")
        cur.execute("PRAGMA foreign_keys = ON")
        self.conn.commit()

    def seed_demo_data(self):
        cur = self.conn.cursor()
        cur.execute("SELECT COUNT(*) AS cnt FROM instructions")
        if cur.fetchone()["cnt"] > 0:
            return

        categories = list(dict.fromkeys(card["category"] for card in SAMPLE_CARDS))
        category_ids = {}

        for idx, category_name in enumerate(categories, start=1):
            cur.execute(
                "INSERT INTO categories (name, sort_order) VALUES (?, ?)",
                (category_name, idx)
            )
            category_ids[category_name] = cur.lastrowid

        instruction_ids = {}

        for card in SAMPLE_CARDS:
            cur.execute("""
                INSERT INTO instructions
                    (category_id, title, short_desc, sections_json, related_ids_json, updated_at)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (
                category_ids[card["category"]],
                card["instruction_title"],
                card["short_desc"],
                json.dumps(card["sections"], ensure_ascii=False),
                json.dumps([], ensure_ascii=False),
                ts(),
            ))
            instruction_ids[card["instruction_title"]] = cur.lastrowid

        # Обновляем связи между инструкциями
        for card in SAMPLE_CARDS:
            related_ids = [
                instruction_ids[title]
                for title in card["related"]
                if title in instruction_ids
            ]
            cur.execute(
                "UPDATE instructions SET related_ids_json=?, updated_at=? WHERE title=?",
                (json.dumps(related_ids, ensure_ascii=False), ts(), card["instruction_title"])
            )

        # Задачи
        for idx, card in enumerate(SAMPLE_CARDS, start=1):
            cur.execute("""
                INSERT INTO tasks
                    (category_id, title, instruction_id, checklist_json, sort_order)
                VALUES (?, ?, ?, ?, ?)
            """, (
                category_ids[card["category"]],
                card["task_title"],
                instruction_ids[card["instruction_title"]],
                json.dumps(card["checklist"], ensure_ascii=False),
                idx,
            ))

        self.conn.commit()

    def all_instruction_titles(self):
        """Возвращает список названий всех инструкций."""
        rows = self.conn.execute("SELECT title FROM instructions ORDER BY title").fetchall()
        return [row["title"] for row in rows]

    def category_by_id(self, category_id):
        row = self.conn.execute("""
            SELECT c.id, c.name, c.sort_order,
                   COALESCE(COUNT(t.id), 0) AS task_count
            FROM categories c
            LEFT JOIN tasks t ON t.category_id = c.id
            WHERE c.id=?
            GROUP BY c.id, c.name, c.sort_order
        """, (category_id,)).fetchone()
        return dict(row) if row else None

    def count_tasks_in_category(self, category_id):
        row = self.conn.execute(
            "SELECT COUNT(*) AS cnt FROM tasks WHERE category_id=?",
            (category_id,)
        ).fetchone()
        return int(row["cnt"]) if row else 0

    def tasks_for_category(self, category_id):
        rows = self.conn.execute("""
            SELECT id AS task_id, title AS task_title
            FROM tasks
            WHERE category_id=?
            ORDER BY sort_order, id
        """, (category_id,)).fetchall()
        return [dict(row) for row in rows]

    def instruction_by_id(self, instruction_id):
        row = self.conn.execute("""
            SELECT i.id AS instruction_id,
                   i.category_id,
                   c.name AS category_name,
                   i.title AS instruction_title,
                   i.short_desc,
                   i.sections_json,
                   i.related_ids_json,
                   i.updated_at
            FROM instructions i
            JOIN categories c ON c.id = i.category_id
            WHERE i.id=?
        """, (instruction_id,)).fetchone()

        if not row:
            return None

        data = dict(row)
        data["sections"] = json.loads(data["sections_json"] or "[]")
        related_ids = json.loads(data["related_ids_json"] or "[]")
        data["related_ids"] = related_ids
        data["related_titles"] = []

        for rid in related_ids:
            r = self.conn.execute(
                "SELECT title FROM instructions WHERE id=?",
                (rid,)
            ).fetchone()
            if r:
                data["related_titles"].append(r["title"])

        return data

    def instruction_by_title(self, title):
        row = self.conn.execute("""
            SELECT id
            FROM instructions
            WHERE LOWER(title)=LOWER(?)
            LIMIT 1
        """, (title,)).fetchone()
        if not row:
            return None
        return self.instruction_by_id(row["id"])

    def task_by_instruction_id(self, instruction_id):
        row = self.conn.execute("""
            SELECT id
            FROM tasks
            WHERE instruction_id=?
            LIMIT 1
        """, (instruction_id,)).fetchone()
        if not row:
            return None
        return self.task_bundle(row["id"])

    def _row_to_task(self, row):
        data = dict(row)
        data["checklist"] = json.loads(data["checklist_json"] or "[]")
        data["sections"] = json.loads(data["sections_json"] or "[]")
        related_ids = json.loads(data["related_ids_json"] or "[]")
        data["related_ids"] = related_ids
        data["related_titles"] = []

        for rid in related_ids:
            rr = self.conn.execute("SELECT title FROM instructions WHERE id=?", (rid,)).fetchone()
            if rr:
                data["related_titles"].append(rr["title"])

        search_parts = [
            data.get("category_name") or "",
            data.get("task_title") or "",
            data.get("instruction_title") or "",
            data.get("short_desc") or "",
            " ".join(data.get("checklist", [])),
            " ".join(
                (section.get("title") or "") + " " + strip_html_tags(section.get("body") or "")
                for section in data.get("sections", [])
            ),
            " ".join((data.get("related_titles") or [])),
        ]
        data["search_blob"] = " ".join(search_parts).casefold()
        return data

    def all_tasks(self):
        rows = self.conn.execute("""
            SELECT t.id AS task_id,
                   t.title AS task_title,
                   t.instruction_id,
                   t.checklist_json,
                   t.sort_order,
                   c.id AS category_id,
                   c.name AS category_name,
                   c.sort_order AS category_sort_order,
                   i.id AS instruction_real_id,
                   i.title AS instruction_title,
                   i.short_desc,
                   i.sections_json,
                   i.related_ids_json
            FROM tasks t
            JOIN categories c ON c.id = t.category_id
            LEFT JOIN instructions i ON i.id = t.instruction_id
            ORDER BY c.sort_order, t.sort_order, t.id
        """).fetchall()

        return [self._row_to_task(row) for row in rows]

    def task_bundle(self, task_id):
        for task in self.all_tasks():
            if task["task_id"] == task_id:
                return task
        return None

    def search_tasks(self, search_text):
        tasks = self.all_tasks()
        if not search_text:
            return tasks

        needle = search_text.casefold().strip()
        return [task for task in tasks if needle in task["search_blob"]]

    def comments_for_instruction(self, instruction_id):
        rows = self.conn.execute("""
            SELECT id, instruction_id, anchor, author, is_anonymous, text, created_at
            FROM comments
            WHERE instruction_id=?
            ORDER BY id DESC
        """, (instruction_id,)).fetchall()
        return [dict(row) for row in rows]

    def add_comment(self, instruction_id, anchor, author, is_anonymous, text):
        self.conn.execute("""
            INSERT INTO comments
                (instruction_id, anchor, author, is_anonymous, text, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            instruction_id,
            anchor or "",
            author,
            1 if is_anonymous else 0,
            text,
            ts()
        ))
        self.conn.commit()

    def comment_by_id(self, comment_id):
        row = self.conn.execute("""
            SELECT id, instruction_id, anchor, author, is_anonymous, text, created_at
            FROM comments
            WHERE id=?
        """, (comment_id,)).fetchone()
        return dict(row) if row else None

    def update_comment(self, comment_id, text):
        self.conn.execute("""
            UPDATE comments
            SET text=?
            WHERE id=?
        """, (text, comment_id))
        self.conn.commit()

    def delete_comment(self, comment_id):
        self.conn.execute(
            "DELETE FROM comments WHERE id=?",
            (comment_id,)
        )
        self.conn.commit()

    def comments_export_rows(self):
        rows = self.conn.execute("""
            SELECT c.id,
                   cat.name AS category_name,
                   i.title AS instruction_title,
                   c.anchor,
                   c.author,
                   c.is_anonymous,
                   c.text,
                   c.created_at
            FROM comments c
            JOIN instructions i ON i.id = c.instruction_id
            JOIN categories cat ON cat.id = i.category_id
            ORDER BY cat.sort_order, i.title, c.id
        """).fetchall()
        return [dict(row) for row in rows]

    def ratings_export_rows(self):
        rows = self.conn.execute("""
            SELECT r.id,
                   cat.name AS category_name,
                   i.title AS instruction_title,
                   r.rating,
                   r.created_at
            FROM ratings r
            JOIN instructions i ON i.id = r.instruction_id
            JOIN categories cat ON cat.id = i.category_id
            ORDER BY cat.sort_order, i.title, r.id
        """).fetchall()
        return [dict(row) for row in rows]

    def export_feedback_xlsx(self, path):
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font
            from openpyxl.utils import get_column_letter
        except ImportError as exc:
            raise RuntimeError(
                "Для экспорта в Excel нужен пакет openpyxl. "
                "Установите его командой: pip install openpyxl"
            ) from exc

        wb = Workbook()

        comments_rows = self.comments_export_rows()
        ratings_rows = self.ratings_export_rows()

        summary_ws = wb.active
        summary_ws.title = "Сводка"
        summary_ws.append(["Показатель", "Значение"])
        summary_ws["A1"].font = Font(bold=True)
        summary_ws["B1"].font = Font(bold=True)

        total_avg = round(
            sum(row["rating"] for row in ratings_rows) / len(ratings_rows), 2
        ) if ratings_rows else 0.0

        summary_data = [
            ("Всего комментариев", len(comments_rows)),
            ("Всего оценок", len(ratings_rows)),
            ("Средняя оценка по всем", total_avg),
        ]
        for item in summary_data:
            summary_ws.append(list(item))

        comments_ws = wb.create_sheet("Комментарии")
        comments_headers = [
            "id", "category_name", "instruction_title", "anchor",
            "author", "is_anonymous", "text", "created_at"
        ]
        comments_ws.append(comments_headers)
        for cell in comments_ws[1]:
            cell.font = Font(bold=True)

        for row in comments_rows:
            comments_ws.append([
                row["id"],
                row["category_name"],
                row["instruction_title"],
                row["anchor"],
                row["author"],
                row["is_anonymous"],
                row["text"],
                row["created_at"],
            ])

        ratings_ws = wb.create_sheet("Оценки")
        ratings_headers = [
            "id", "category_name", "instruction_title", "rating", "created_at"
        ]
        ratings_ws.append(ratings_headers)
        for cell in ratings_ws[1]:
            cell.font = Font(bold=True)

        for row in ratings_rows:
            ratings_ws.append([
                row["id"],
                row["category_name"],
                row["instruction_title"],
                row["rating"],
                row["created_at"],
            ])

        for ws in (summary_ws, comments_ws, ratings_ws):
            for col_idx, column_cells in enumerate(ws.iter_cols(1, ws.max_column), start=1):
                max_len = 0
                for cell in column_cells:
                    value = cell.value
                    if value is not None:
                        max_len = max(max_len, len(str(value)))
                ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 60)

        wb.save(str(path))

    def rating_stats(self, instruction_id):
        row = self.conn.execute("""
            SELECT ROUND(AVG(rating), 2) AS avg_rating,
                   COUNT(*) AS cnt
            FROM ratings
            WHERE instruction_id=?
        """, (instruction_id,)).fetchone()

        avg_rating = float(row["avg_rating"]) if row and row["avg_rating"] is not None else 0.0
        cnt = int(row["cnt"]) if row and row["cnt"] is not None else 0
        return avg_rating, cnt

    def add_rating(self, instruction_id, rating):
        self.conn.execute("""
            INSERT INTO ratings (instruction_id, rating, created_at)
            VALUES (?, ?, ?)
        """, (instruction_id, int(rating), ts()))
        self.conn.commit()

    def ratings_for_instruction(self, instruction_id):
        rows = self.conn.execute("""
            SELECT id, rating, created_at
            FROM ratings
            WHERE instruction_id=?
            ORDER BY id DESC
        """, (instruction_id,)).fetchall()
        return [dict(row) for row in rows]

    def delete_rating(self, rating_id):
        self.conn.execute(
            "DELETE FROM ratings WHERE id=?",
            (rating_id,)
        )
        self.conn.commit()

    def add_category(self, name):
        name = name.strip()
        row = self.conn.execute(
            "SELECT COALESCE(MAX(sort_order), 0) + 1 AS next_order FROM categories"
        ).fetchone()
        next_order = int(row["next_order"]) if row else 1

        cur = self.conn.execute(
            "INSERT INTO categories (name, sort_order) VALUES (?, ?)",
            (name, next_order)
        )
        self.conn.commit()
        return cur.lastrowid

    def add_task(self, category_id, title, instruction_title=""):
        title = title.strip()
        instruction_id = None

        if instruction_title.strip():
            instruction = self.instruction_by_title(instruction_title.strip())
            if not instruction:
                raise ValueError(f"Инструкция не найдена: {instruction_title}")
            instruction_id = instruction["instruction_id"]

        row = self.conn.execute(
            "SELECT COALESCE(MAX(sort_order), 0) + 1 AS next_order FROM tasks WHERE category_id=?",
            (category_id,)
        ).fetchone()
        next_order = int(row["next_order"]) if row else 1

        cur = self.conn.execute("""
            INSERT INTO tasks
                (category_id, title, instruction_id, checklist_json, sort_order)
            VALUES (?, ?, ?, ?, ?)
        """, (
            category_id,
            title,
            instruction_id,
            json.dumps([], ensure_ascii=False),
            next_order
        ))
        self.conn.commit()
        return cur.lastrowid

    def rename_category(self, category_id, new_name):
        self.conn.execute(
            "UPDATE categories SET name=? WHERE id=?",
            (new_name.strip(), category_id)
        )
        self.conn.commit()

    def rename_task(self, task_id, new_title):
        self.conn.execute(
            "UPDATE tasks SET title=? WHERE id=?",
            (new_title.strip(), task_id)
        )
        self.conn.commit()

    def delete_category(self, category_id):
        self.conn.execute(
            "DELETE FROM categories WHERE id=?",
            (category_id,)
        )
        self.conn.commit()

    def delete_task(self, task_id):
        self.conn.execute(
            "DELETE FROM tasks WHERE id=?",
            (task_id,)
        )
        self.conn.commit()

    def move_category_up(self, category_id):
        """Перемещает категорию вверх (уменьшает sort_order, меняясь с предыдущей)."""
        current = self.conn.execute(
            "SELECT sort_order FROM categories WHERE id=?", (category_id,)
        ).fetchone()
        if not current:
            return
        cur_order = current["sort_order"]
        # Найти ближайшую категорию с меньшим sort_order
        prev = self.conn.execute(
            "SELECT id, sort_order FROM categories WHERE sort_order < ? ORDER BY sort_order DESC LIMIT 1",
            (cur_order,)
        ).fetchone()
        if not prev:
            return  # уже наверху
        # Обмен sort_order
        self.conn.execute("UPDATE categories SET sort_order=? WHERE id=?", (prev["sort_order"], category_id))
        self.conn.execute("UPDATE categories SET sort_order=? WHERE id=?", (cur_order, prev["id"]))
        self.conn.commit()

    def move_category_down(self, category_id):
        """Перемещает категорию вниз (увеличивает sort_order, меняясь со следующей)."""
        current = self.conn.execute(
            "SELECT sort_order FROM categories WHERE id=?", (category_id,)
        ).fetchone()
        if not current:
            return
        cur_order = current["sort_order"]
        # Найти ближайшую категорию с большим sort_order
        next_cat = self.conn.execute(
            "SELECT id, sort_order FROM categories WHERE sort_order > ? ORDER BY sort_order ASC LIMIT 1",
            (cur_order,)
        ).fetchone()
        if not next_cat:
            return  # уже внизу
        self.conn.execute("UPDATE categories SET sort_order=? WHERE id=?", (next_cat["sort_order"], category_id))
        self.conn.execute("UPDATE categories SET sort_order=? WHERE id=?", (cur_order, next_cat["id"]))
        self.conn.commit()

    def move_task_up(self, task_id):
        """Перемещает задачу вверх в пределах её категории."""
        task = self.conn.execute(
            "SELECT category_id, sort_order FROM tasks WHERE id=?", (task_id,)
        ).fetchone()
        if not task:
            return
        cat_id, cur_order = task["category_id"], task["sort_order"]
        prev = self.conn.execute(
            "SELECT id, sort_order FROM tasks WHERE category_id=? AND sort_order < ? ORDER BY sort_order DESC LIMIT 1",
            (cat_id, cur_order)
        ).fetchone()
        if not prev:
            return
        self.conn.execute("UPDATE tasks SET sort_order=? WHERE id=?", (prev["sort_order"], task_id))
        self.conn.execute("UPDATE tasks SET sort_order=? WHERE id=?", (cur_order, prev["id"]))
        self.conn.commit()

    def move_task_down(self, task_id):
        """Перемещает задачу вниз в пределах её категории."""
        task = self.conn.execute(
            "SELECT category_id, sort_order FROM tasks WHERE id=?", (task_id,)
        ).fetchone()
        if not task:
            return
        cat_id, cur_order = task["category_id"], task["sort_order"]
        next_task = self.conn.execute(
            "SELECT id, sort_order FROM tasks WHERE category_id=? AND sort_order > ? ORDER BY sort_order ASC LIMIT 1",
            (cat_id, cur_order)
        ).fetchone()
        if not next_task:
            return
        self.conn.execute("UPDATE tasks SET sort_order=? WHERE id=?", (next_task["sort_order"], task_id))
        self.conn.execute("UPDATE tasks SET sort_order=? WHERE id=?", (cur_order, next_task["id"]))
        self.conn.commit()

    def update_task_view_data(self, task_id, short_desc, instruction_title, checklist):
        task = self.task_bundle(task_id)
        if not task:
            raise ValueError("Задача не найдена.")

        instruction_id = task.get("instruction_real_id")
        if not instruction_id:
            raise ValueError("У задачи нет привязанной инструкции.")

        short_desc = short_desc.strip()
        instruction_title = instruction_title.strip()

        if not short_desc:
            raise ValueError("Описание не может быть пустым.")

        if not instruction_title:
            raise ValueError("Название инструкции не может быть пустым.")

        other = self.conn.execute("""
            SELECT id
            FROM instructions
            WHERE LOWER(title)=LOWER(?) AND id<>?
            LIMIT 1
        """, (instruction_title, instruction_id)).fetchone()

        if other:
            raise ValueError("Инструкция с таким названием уже существует.")

        self.conn.execute("""
            UPDATE instructions
            SET title=?, short_desc=?, updated_at=?
            WHERE id=?
        """, (
            instruction_title,
            short_desc,
            ts(),
            instruction_id
        ))

        self.conn.execute("""
            UPDATE tasks
            SET checklist_json=?
            WHERE id=?
        """, (
            json.dumps(checklist, ensure_ascii=False),
            task_id
        ))

        self.conn.commit()

    def add_instruction(self, category_id, task_id, instruction_title, short_desc, sections, related_titles):
        """
        Создаёт инструкцию и привязывает её к существующей задаче task_id.
        Больше не создаёт новую задачу.
        """
        instruction_title = instruction_title.strip()
        short_desc = short_desc.strip()

        if not instruction_title:
            raise ValueError("Название инструкции не может быть пустым.")
        if not task_id:
            raise ValueError("Не указана задача для привязки инструкции.")

        # Проверяем, что задача существует
        task = self.task_bundle(task_id)
        if not task:
            raise ValueError("Задача не найдена.")

        if not sections:
            sections = [{"title": "Коротко", "body": p(short_desc or instruction_title), "image_path": ""}]

        related_ids = []
        missing_related = []
        for rel_title in related_titles:
            inst = self.instruction_by_title(rel_title)
            if inst:
                related_ids.append(inst["instruction_id"])
            else:
                missing_related.append(rel_title)

        cur = self.conn.execute("""
            INSERT INTO instructions
                (category_id, title, short_desc, sections_json, related_ids_json, updated_at)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            category_id,
            instruction_title,
            short_desc,
            json.dumps(sections, ensure_ascii=False),
            json.dumps(related_ids, ensure_ascii=False),
            ts(),
        ))
        instruction_id = cur.lastrowid

        # Привязываем инструкцию к задаче
        self.conn.execute(
            "UPDATE tasks SET instruction_id = ? WHERE id = ?",
            (instruction_id, task_id)
        )
        self.conn.commit()

        return instruction_id, missing_related

    def get_help_content(self):
        """Возвращает содержимое справки из таблицы settings."""
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL DEFAULT ''
            )
        """)
        row = self.conn.execute(
            "SELECT value FROM settings WHERE key='help_content'"
        ).fetchone()
        if row:
            return row["value"]
        # Значение по умолчанию
        default_help = (
            "## Как использовать приложение\n\n"
            "### Навигация\n"
            "- В левой панели выберите **категорию**, затем **задачу**\n"
            "- Используйте **поиск** для быстрого нахождения нужной задачи\n\n"
            "### Вкладки\n"
            "- **📝 Задача** — описание и чек-лист\n"
            "- **📘 Инструкция** — подробная инструкция с блоками\n"
            "- **💬 Комментарии и оценка** — обратная связь\n\n"
            "### Чек-лист\n"
            "- Отмечайте выполненные пункты ✅\n"
            "- Прогресс сохраняется автоматически\n\n"
            "### Форматирование текста (Markdown)\n"
            "- `# Заголовок 1`\n"
            "- `## Заголовок 2`\n"
            "- `**жирный**`\n"
            "- `*курсив*`\n"
            "- `- список`\n"
            "- `1. нумерованный список`\n"
            "- `` `код` ``\n"
            "- `> цитата`\n"
            "- `[текст](ссылка)`\n\n"
            "### Для администратора\n"
            "- Кнопка **✎** в заголовке задачи — редактирование задачи\n"
            "- Кнопка **Редактировать инструкцию** — изменение инструкции\n"
            "- Кнопка **✎** в верхней плашке — изменение приветствия\n"
            "- Кнопка **Экспорт в Excel** — выгрузка комментариев и оценок\n"
        )
        self.conn.execute(
            "INSERT OR IGNORE INTO settings (key, value) VALUES ('help_content', ?)",
            (default_help,)
        )
        self.conn.commit()
        return default_help

    def save_help_content(self, content):
        """Сохраняет содержимое справки."""
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL DEFAULT ''
            )
        """)
        self.conn.execute(
            "INSERT OR REPLACE INTO settings (key, value) VALUES ('help_content', ?)",
            (content,)
        )
        self.conn.commit()

    def get_banner_text(self):
        """Возвращает сохранённый текст баннера или значение по умолчанию."""
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL DEFAULT ''
            )
        """)
        row = self.conn.execute(
            "SELECT value FROM settings WHERE key='banner_text'"
        ).fetchone()
        if row:
            return row["value"]
        default = (
            "Слева выбери задачу. Сначала смотри чек‑лист, потом подробную инструкцию. "
            "Комментарии и оценки — анонимные."
        )
        self.conn.execute(
            "INSERT OR IGNORE INTO settings (key, value) VALUES ('banner_text', ?)",
            (default,)
        )
        self.conn.commit()
        return default

    def save_banner_text(self, text):
        """Сохраняет текст баннера."""
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL DEFAULT ''
            )
        """)
        self.conn.execute(
            "INSERT OR REPLACE INTO settings (key, value) VALUES ('banner_text', ?)",
            (text,)
        )
        self.conn.commit()

    def update_instruction(self, instruction_id, category_id, title, short_desc, sections, related_titles):
        """Обновляет существующую инструкцию."""
        title = title.strip()
        short_desc = short_desc.strip()
        if not title:
            raise ValueError("Название инструкции не может быть пустым.")

        # Проверяем, что инструкция существует
        existing = self.instruction_by_id(instruction_id)
        if not existing:
            raise ValueError("Инструкция не найдена.")

        # Проверка уникальности названия (исключая саму себя)
        other = self.conn.execute(
            "SELECT id FROM instructions WHERE LOWER(title)=LOWER(?) AND id<>?",
            (title, instruction_id)
        ).fetchone()
        if other:
            raise ValueError("Инструкция с таким названием уже существует.")

        related_ids = []
        missing_related = []
        for rel_title in related_titles:
            inst = self.instruction_by_title(rel_title)
            if inst:
                related_ids.append(inst["instruction_id"])
            else:
                missing_related.append(rel_title)

        self.conn.execute("""
            UPDATE instructions
            SET category_id=?,
                title=?,
                short_desc=?,
                sections_json=?,
                related_ids_json=?,
                updated_at=?
            WHERE id=?
        """, (
            category_id,
            title,
            short_desc,
            json.dumps(sections, ensure_ascii=False),
            json.dumps(related_ids, ensure_ascii=False),
            ts(),
            instruction_id
        ))
        self.conn.commit()
        return missing_related


# ================== UI ==================
class InstructionEditorDialog(QDialog):
    def __init__(self, db: KnowledgeBaseDB, categories, default_category_id=None, instruction_data=None, parent=None):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("Редактор инструкции" if instruction_data is None else "Редактирование инструкции")
        self.setModal(True)
        self.resize(900, 800)

        layout = QVBoxLayout(self)

        # Категория и задача
        form = QFormLayout()
        self.category_combo = QComboBox()
        for cat in categories:
            self.category_combo.addItem(cat["name"], cat["id"])
        if default_category_id is not None:
            idx = self.category_combo.findData(default_category_id)
            if idx >= 0:
                self.category_combo.setCurrentIndex(idx)

        self.task_combo = QComboBox()
        self._populate_tasks()

        # При смене категории обновляем список задач
        self.category_combo.currentIndexChanged.connect(self._populate_tasks)

        self.title_edit = QLineEdit()
        self.title_edit.setPlaceholderText("Название инструкции")
        self.short_desc_edit = QLineEdit()
        self.short_desc_edit.setPlaceholderText("Краткое описание (одна строка)")

        form.addRow("Категория:", self.category_combo)
        form.addRow("Задача:", self.task_combo)
        form.addRow("Название инструкции:", self.title_edit)
        form.addRow("Краткое описание:", self.short_desc_edit)
        layout.addLayout(form)

        # Секции (без изменений)
        sections_group = QGroupBox("Секции")
        sections_layout = QVBoxLayout(sections_group)

        self.sections_widget = QWidget()
        self.sections_layout = QVBoxLayout(self.sections_widget)
        self.sections_layout.setContentsMargins(0, 0, 0, 0)
        self.sections_layout.setSpacing(8)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(self.sections_widget)
        sections_layout.addWidget(scroll)

        # Подсказка по markdown
        markdown_hint = QLabel(
            '<span style="color:#5b6577; font-size:11px;">'
            'Поддерживается Markdown: <b>#</b> Заголовок, <b>**</b>жирный<b>**</b>, '
            '<b>*</b>курсив<b>*</b>, <b>`</b>код<b>`</b>, '
            '<b>-</b> список, <b>1.</b> нумерованный список, '
            '<b>&gt;</b> цитата, <b>[текст](url)</b> ссылка'
            '</span>'
        )
        markdown_hint.setWordWrap(True)
        markdown_hint.setTextFormat(Qt.RichText)
        sections_layout.addWidget(markdown_hint)

        btn_row = QHBoxLayout()
        add_btn = QPushButton("Добавить секцию")
        add_btn.clicked.connect(self.add_section)
        btn_row.addWidget(add_btn)
        btn_row.addStretch()
        sections_layout.addLayout(btn_row)

        layout.addWidget(sections_group)

        # Связанные инструкции (без изменений)
        related_group = QGroupBox("Связанные инструкции")
        related_layout = QVBoxLayout(related_group)

        self.related_combo = QComboBox()
        self.related_combo.setEditable(False)
        self._populate_related_combo()
        add_rel_btn = QPushButton("Добавить связь")
        add_rel_btn.clicked.connect(self.add_related)

        rel_combo_layout = QHBoxLayout()
        rel_combo_layout.addWidget(self.related_combo, 1)
        rel_combo_layout.addWidget(add_rel_btn)

        self.related_list = QListWidget()
        self.related_list.setAlternatingRowColors(True)
        del_rel_btn = QPushButton("Удалить выбранную связь")
        del_rel_btn.clicked.connect(self.remove_selected_related)

        related_layout.addLayout(rel_combo_layout)
        related_layout.addWidget(self.related_list)
        related_layout.addWidget(del_rel_btn)

        layout.addWidget(related_group)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.validate_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        # Заполнение при редактировании
        if instruction_data is not None:
            self._populate_from_data(instruction_data)
        else:
            self.add_section()

    def _populate_tasks(self):
        """Заполняет выпадающий список задач для выбранной категории."""
        self.task_combo.clear()
        cat_id = self.category_combo.currentData()
        if cat_id is None:
            return
        tasks = self.db.tasks_for_category(cat_id)
        for t in tasks:
            self.task_combo.addItem(t["task_title"], t["task_id"])
        # Если нет задач, показываем предупреждение
        if self.task_combo.count() == 0:
            self.task_combo.addItem("(нет задач)", None)

    def _populate_from_data(self, data: dict):
        """Заполняет поля редактора существующими данными инструкции."""
        # Категория
        idx = self.category_combo.findData(data.get("category_id"))
        if idx >= 0:
            self.category_combo.setCurrentIndex(idx)

        # Задача – при редактировании показываем задачу, к которой привязана инструкция
        task_id = data.get("task_id")
        if task_id is not None:
            # Найдём задачу в списке
            for i in range(self.task_combo.count()):
                if self.task_combo.itemData(i) == task_id:
                    self.task_combo.setCurrentIndex(i)
                    break
        # Блокируем изменение задачи при редактировании
        self.task_combo.setEnabled(data.get("is_edit", False) == False)

        self.title_edit.setText(data.get("title", ""))
        self.short_desc_edit.setText(data.get("short_desc", ""))

        # Секции
        sections = data.get("sections", [])
        for sec in sections:
            self.add_section()
            frame = self.sections_layout.itemAt(self.sections_layout.count() - 1).widget()

            # Заголовок
            if hasattr(frame, 'title_edit'):
                frame.title_edit.setText(sec.get("title", ""))

            # Блоки
            blocks = sec.get("blocks", [])
            # Удаляем дефолтный текстовый блок, если есть данные
            if blocks and hasattr(frame, 'blocks_layout'):
                # Очищаем blocks_layout
                while frame.blocks_layout.count() > 0:
                    item = frame.blocks_layout.takeAt(0)
                    if item.widget():
                        item.widget().deleteLater()

                for block in blocks:
                    block_type = block.get("type")
                    if block_type == "text":
                        self._add_text_block(frame.blocks_layout)
                        # Заполняем последний добавленный блок
                        last_block = frame.blocks_layout.itemAt(frame.blocks_layout.count() - 1).widget()
                        if hasattr(last_block, 'text_edit'):
                            content = block.get("content", "")
                            last_block.text_edit.setPlainText(content)

                    elif block_type == "images":
                        self._add_images_block(frame.blocks_layout)
                        last_block = frame.blocks_layout.itemAt(frame.blocks_layout.count() - 1).widget()
                        # Восстанавливаем ширину, если она сохранена
                        if hasattr(last_block, 'width_combo'):
                            saved_width = str(block.get("image_width", 760))
                            if saved_width == "0":
                                last_block.width_combo.setCurrentText("Оригинал")
                            else:
                                last_block.width_combo.setCurrentText(saved_width)
                        # Загружаем пути изображений
                        if hasattr(last_block, 'img_list'):
                            for img_path in block.get("paths", []):
                                if img_path:
                                    last_block.img_list.addItem(img_path)

        if not sections:
            self.add_section()

        # Связанные инструкции
        for rel_title in data.get("related_titles", []):
            found = False
            for i in range(self.related_combo.count()):
                if self.related_combo.itemText(i) == rel_title:
                    found = True
                    break
            if found:
                self.related_list.addItem(rel_title)

    def _populate_related_combo(self):
        self.related_combo.clear()
        titles = self.db.all_instruction_titles()
        self.related_combo.addItems(titles)

    def add_section(self):
        frame = QFrame()
        frame.setFrameStyle(QFrame.Box | QFrame.Plain)
        frame.setStyleSheet("QFrame { border: 1px solid #c0c0c0; border-radius: 6px; background: #fdfdfd; }")

        layout = QVBoxLayout(frame)

        # Заголовок секции
        title_edit = QLineEdit()
        title_edit.setPlaceholderText("Заголовок секции")
        layout.addWidget(title_edit)

        # Список блоков контента (текст / изображения)
        blocks_widget = QWidget()
        blocks_layout = QVBoxLayout(blocks_widget)
        blocks_layout.setContentsMargins(0, 0, 0, 0)
        blocks_layout.setSpacing(6)
        layout.addWidget(blocks_widget)

        # Кнопки для добавления блоков
        add_block_layout = QHBoxLayout()
        add_text_btn = QPushButton("+ Текст")
        add_text_btn.clicked.connect(lambda: self._add_text_block(blocks_layout))
        add_images_btn = QPushButton("+ Изображения")
        add_images_btn.clicked.connect(lambda: self._add_images_block(blocks_layout))
        add_block_layout.addWidget(add_text_btn)
        add_block_layout.addWidget(add_images_btn)
        add_block_layout.addStretch()
        layout.addLayout(add_block_layout)

        # Сохраняем ссылки в frame для последующего сбора данных
        frame.title_edit = title_edit
        frame.blocks_layout = blocks_layout

        # Кнопки управления секцией
        ctrl_layout = QHBoxLayout()
        up_btn = QPushButton("↑")
        up_btn.setToolTip("Переместить выше")
        up_btn.clicked.connect(lambda: self._move_section(frame, -1))
        down_btn = QPushButton("↓")
        down_btn.setToolTip("Переместить ниже")
        down_btn.clicked.connect(lambda: self._move_section(frame, 1))
        del_btn = QPushButton("Удалить секцию")
        del_btn.clicked.connect(lambda: self._delete_section(frame))

        ctrl_layout.addWidget(up_btn)
        ctrl_layout.addWidget(down_btn)
        ctrl_layout.addStretch()
        ctrl_layout.addWidget(del_btn)
        layout.addLayout(ctrl_layout)

        self.sections_layout.addWidget(frame)

        # Добавляем один текстовый блок по умолчанию
        self._add_text_block(blocks_layout)

    def _browse_image(self, line_edit: QLineEdit):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите изображение", "",
            "Изображения (*.png *.jpg *.jpeg *.bmp *.gif);;Все файлы (*)"
        )
        if file_path:
            line_edit.setText(file_path)

    def _add_text_block(self, blocks_layout):
        """Добавляет текстовый блок в секцию."""
        block_frame = QFrame()
        block_frame.setFrameStyle(QFrame.StyledPanel | QFrame.Plain)
        block_frame.setStyleSheet("QFrame { background: #ffffff; border: 1px solid #e0e0e0; border-radius: 4px; }")
        block_layout = QVBoxLayout(block_frame)
        block_layout.setContentsMargins(6, 6, 6, 6)
        block_layout.setSpacing(4)

        # Метка типа блока
        header = QHBoxLayout()
        header.addWidget(QLabel("📝 Текст"))
        header.addStretch()
        del_btn = QPushButton("✕")
        del_btn.setFixedSize(24, 24)
        del_btn.setToolTip("Удалить блок")
        del_btn.clicked.connect(lambda: self._delete_block(block_frame, blocks_layout))
        header.addWidget(del_btn)
        block_layout.addLayout(header)

        text_edit = QPlainTextEdit()
        text_edit.setPlaceholderText("Текст блока...")
        text_edit.setMinimumHeight(60)
        block_layout.addWidget(text_edit)

        # Сохраняем тип блока
        block_frame.block_type = "text"
        block_frame.text_edit = text_edit

        blocks_layout.addWidget(block_frame)

    def _add_images_block(self, blocks_layout):
        """Добавляет блок изображений в секцию."""
        block_frame = QFrame()
        block_frame.setFrameStyle(QFrame.StyledPanel | QFrame.Plain)
        block_frame.setStyleSheet("QFrame { background: #ffffff; border: 1px solid #e0e0e0; border-radius: 4px; }")
        block_layout = QVBoxLayout(block_frame)
        block_layout.setContentsMargins(6, 6, 6, 6)
        block_layout.setSpacing(4)

        # Метка типа блока
        header = QHBoxLayout()
        header.addWidget(QLabel("🖼 Изображения"))
        header.addStretch()
        del_btn = QPushButton("✕")
        del_btn.setFixedSize(24, 24)
        del_btn.setToolTip("Удалить блок")
        del_btn.clicked.connect(lambda: self._delete_block(block_frame, blocks_layout))
        header.addWidget(del_btn)
        block_layout.addLayout(header)

        # Список изображений
        img_list = QListWidget()
        img_list.setAlternatingRowColors(True)
        img_list.setMaximumHeight(100)
        block_layout.addWidget(img_list)

        # Кнопки управления изображениями
        img_btn_layout = QHBoxLayout()
        add_img_btn = QPushButton("Добавить изображение")
        add_img_btn.clicked.connect(lambda _, lst=img_list: self._add_image_to_list(lst))
        remove_img_btn = QPushButton("Удалить")
        remove_img_btn.clicked.connect(lambda _, lst=img_list: self._remove_selected_image(lst))
        up_img_btn = QPushButton("↑")
        up_img_btn.clicked.connect(lambda _, lst=img_list: self._move_image(lst, -1))
        down_img_btn = QPushButton("↓")
        down_img_btn.clicked.connect(lambda _, lst=img_list: self._move_image(lst, 1))

        img_btn_layout.addWidget(add_img_btn)
        img_btn_layout.addWidget(remove_img_btn)
        img_btn_layout.addWidget(up_img_btn)
        img_btn_layout.addWidget(down_img_btn)
        img_btn_layout.addStretch()
        block_layout.addLayout(img_btn_layout)

        # Выбор ширины отображения
        width_layout = QHBoxLayout()
        width_layout.addWidget(QLabel("Ширина:"))
        width_combo = QComboBox()
        width_combo.addItems(["400", "600", "760", "1000", "Оригинал"])
        width_combo.setCurrentText("760")
        width_layout.addWidget(width_combo)
        width_layout.addStretch()
        block_layout.addLayout(width_layout)

        # Сохраняем тип блока и ссылки
        block_frame.block_type = "images"
        block_frame.img_list = img_list
        block_frame.width_combo = width_combo

        blocks_layout.addWidget(block_frame)

    def _delete_block(self, block_frame, blocks_layout):
        """Удаляет блок из секции."""
        blocks_layout.removeWidget(block_frame)
        block_frame.deleteLater()

    def _add_image_to_list(self, img_list):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите изображение", "",
            "Изображения (*.png *.jpg *.jpeg *.bmp *.gif);;Все файлы (*)"
        )
        if file_path:
            img_list.addItem(file_path)

    def _remove_selected_image(self, img_list):
        row = img_list.currentRow()
        if row >= 0:
            img_list.takeItem(row)

    def _move_image(self, img_list, direction):
        row = img_list.currentRow()
        if row < 0:
            return
        new_row = row + direction
        if 0 <= new_row < img_list.count():
            item = img_list.takeItem(row)
            img_list.insertItem(new_row, item)
            img_list.setCurrentRow(new_row)

    def _move_section(self, frame, direction: int):
        idx = self.sections_layout.indexOf(frame)
        if idx == -1:
            return
        new_idx = idx + direction
        if 0 <= new_idx < self.sections_layout.count():
            self.sections_layout.removeWidget(frame)
            self.sections_layout.insertWidget(new_idx, frame)

    def _delete_section(self, frame):
        if self.sections_layout.count() <= 1:
            QMessageBox.warning(self, "Внимание", "Должна остаться хотя бы одна секция.")
            return
        self.sections_layout.removeWidget(frame)
        frame.deleteLater()

    def add_related(self):
        title = self.related_combo.currentText().strip()
        if not title:
            return
        for i in range(self.related_list.count()):
            if self.related_list.item(i).text() == title:
                QMessageBox.information(self, "Внимание", "Эта инструкция уже добавлена.")
                return
        self.related_list.addItem(title)

    def remove_selected_related(self):
        row = self.related_list.currentRow()
        if row >= 0:
            self.related_list.takeItem(row)

    def _gather_sections(self):
        sections = []
        for i in range(self.sections_layout.count()):
            frame = self.sections_layout.itemAt(i).widget()
            if frame is None:
                continue

            title = frame.title_edit.text().strip() if hasattr(frame, 'title_edit') else ""
            blocks = []

            if hasattr(frame, 'blocks_layout'):
                for j in range(frame.blocks_layout.count()):
                    block_frame = frame.blocks_layout.itemAt(j).widget()
                    if block_frame is None:
                        continue

                    block_type = getattr(block_frame, 'block_type', None)

                    if block_type == "text":
                        text = block_frame.text_edit.toPlainText().strip() if hasattr(block_frame, 'text_edit') else ""
                        if text:
                            blocks.append({
                                "type": "text",
                                "content": text  # сохраняем как есть, markdown применится при отображении
                            })

                    elif block_type == "images":
                        image_paths = []
                        if hasattr(block_frame, 'img_list'):
                            for k in range(block_frame.img_list.count()):
                                path = block_frame.img_list.item(k).text().strip()
                                if path:
                                    image_paths.append(path)
                        # Собираем выбранную ширину
                        image_width = None
                        if hasattr(block_frame, 'width_combo'):
                            w = block_frame.width_combo.currentText()
                            if w == "Оригинал":
                                image_width = 0  # 0 – не масштабировать
                            else:
                                image_width = int(w)
                        else:
                            image_width = 760  # default
                        if image_paths:
                            blocks.append({
                                "type": "images",
                                "paths": image_paths,
                                "image_width": image_width
                            })

            if title or blocks:
                sections.append({
                    "title": title,
                    "blocks": blocks
                })
        return sections

    def validate_and_accept(self):
        if not self.title_edit.text().strip():
            QMessageBox.warning(self, "Ошибка", "Введите название инструкции.")
            return
        # Проверка, что задача выбрана (только при создании, если поле доступно)
        if self.task_combo.isEnabled() and self.task_combo.currentData() is None:
            QMessageBox.warning(self, "Ошибка", "Выберите задачу, к которой привязать инструкцию.")
            return
        sections = self._gather_sections()
        if not sections:
            QMessageBox.warning(self, "Ошибка", "Добавьте хотя бы одну секцию.")
            return
        self.accept()

    def get_data(self):
        sections = self._gather_sections()
        related_titles = [self.related_list.item(i).text() for i in range(self.related_list.count())]
        return {
            "category_id": self.category_combo.currentData(),
            "task_id": self.task_combo.currentData(),  # None при редактировании, если задача отключена
            "title": self.title_edit.text().strip(),
            "short_desc": self.short_desc_edit.text().strip(),
            "sections": sections,
            "related_titles": related_titles
        }


class CollapsibleSection(QFrame):
    def __init__(self, title, blocks=None, link_handler=None):
        super().__init__()
        self.setObjectName("sectionCard")
        self.setStyleSheet("""
QFrame#sectionCard {
    background: #ffffff;
    border: 1px solid #d9e3f0;
    border-radius: 10px;
}

QFrame#sectionCard QToolButton {
    font-size: 15px;
    font-weight: bold;
    background: #f7f9fd;
    color: #263238;
    border-bottom: 1px solid #d9e3f0;
    border-top-left-radius: 10px;
    border-top-right-radius: 10px;
}

QFrame#sectionCard QToolButton:hover {
    background: #eef5fb;
}

            QToolButton {
                font-weight: 600;
                text-align: left;
                padding: 10px 12px;
                background: #eef5fb;
                color: #263238;
                border-bottom: 1px solid #d9e3f0;
                border-top-left-radius: 10px;
                border-top-right-radius: 10px;
            }

            QToolButton:hover {
                background: #e4f0ff;
            }

            QFrame#sectionCard QWidget#qt_scrollarea_viewport,
QFrame#sectionCard QWidget {
    background: #ffffff;
}
        """)

        self.link_handler = link_handler

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.toggle = QToolButton(text=title)
        self.toggle.setCheckable(True)
        self.toggle.setChecked(False)
        self.toggle.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.toggle.setArrowType(Qt.RightArrow)
        self.toggle.toggled.connect(self.on_toggled)

        layout.addWidget(self.toggle)

        self.content = QWidget()
        content_layout = QVBoxLayout(self.content)
        content_layout.setContentsMargins(12, 8, 12, 12)
        content_layout.setSpacing(8)

        # Отображаем блоки
        for block in (blocks or []):
            block_type = block.get("type")
            if block_type == "text":
                raw_text = block.get("content", "")
                # Применяем markdown-рендеринг
                body_html = render_markdown(raw_text)
                body_label = QLabel()
                body_label.setWordWrap(True)
                body_label.setTextFormat(Qt.RichText)
                body_label.setTextInteractionFlags(Qt.TextBrowserInteraction)
                body_label.linkActivated.connect(self.on_link_activated)
                body_label.setText(body_html)
                content_layout.addWidget(body_label)

            elif block_type == "images":
                image_width = block.get("image_width", 760) or 760
                if image_width == 0:  # Оригинал
                    scale_mode = Qt.SmoothTransformation
                    use_width = 0
                else:
                    scale_mode = Qt.SmoothTransformation
                    use_width = image_width
                for image_path in block.get("paths", []):
                    img = QPixmap(str(image_path))
                    if not img.isNull():
                        image_label = QLabel()
                        if use_width == 0:
                            pix = img
                        else:
                            pix = img.scaledToWidth(use_width, Qt.SmoothTransformation)
                        image_label.setPixmap(pix)
                        image_label.setAlignment(Qt.AlignCenter)
                        image_label.setStyleSheet("border: 1px solid #d9e3f0; border-radius: 6px;")
                        shadow = QGraphicsDropShadowEffect()
                        shadow.setBlurRadius(12)
                        shadow.setOffset(2, 2)
                        shadow.setColor(QColor(0, 0, 0, 40))
                        image_label.setGraphicsEffect(shadow)
                        content_layout.addWidget(image_label)

        self.content.setVisible(False)
        layout.addWidget(self.content)

    def on_toggled(self, checked):
        self.content.setVisible(checked)
        self.toggle.setArrowType(Qt.DownArrow if checked else Qt.RightArrow)

    def on_link_activated(self, link):
        if self.link_handler:
            self.link_handler(link)


class TaskEditorDialog(QDialog):
    """Диалог редактирования задачи (описание, инструкция, чек-лист)."""

    def __init__(self, db: KnowledgeBaseDB, task, instruction, parent=None):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("Редактирование задачи")
        self.setModal(True)
        self.resize(700, 600)

        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        # Описание
        layout.addWidget(QLabel("Описание (поддерживается Markdown):"))
        self.desc_edit = QPlainTextEdit()
        self.desc_edit.setPlaceholderText("Описание задачи...")
        self.desc_edit.setMinimumHeight(120)
        self.desc_edit.setPlainText(task.get("short_desc", ""))
        layout.addWidget(self.desc_edit)

        # Название инструкции
        form = QFormLayout()
        self.instruction_edit = QLineEdit()
        self.instruction_edit.setPlaceholderText("Название инструкции")
        if instruction:
            self.instruction_edit.setText(instruction.get("instruction_title", ""))
        form.addRow("Инструкция:", self.instruction_edit)
        layout.addLayout(form)

        # Чек-лист
        layout.addWidget(QLabel("Чек-лист (по одному пункту на строку):"))
        self.checklist_edit = QPlainTextEdit()
        self.checklist_edit.setPlaceholderText("Пункт 1\nПункт 2\nПункт 3")
        checklist_text = "\n".join(task.get("checklist", []))
        self.checklist_edit.setPlainText(checklist_text)
        layout.addWidget(self.checklist_edit)

        # Кнопки
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.validate_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def validate_and_accept(self):
        desc = self.desc_edit.toPlainText().strip()
        instruction_title = self.instruction_edit.text().strip()
        checklist = split_non_empty_lines(self.checklist_edit.toPlainText())

        if not desc:
            QMessageBox.warning(self, "Ошибка", "Введите описание.")
            return
        if not instruction_title:
            QMessageBox.warning(self, "Ошибка", "Введите название инструкции.")
            return
        if not checklist:
            QMessageBox.warning(self, "Ошибка", "Чек-лист не может быть пустым.")
            return

        self.accept()

    def get_data(self):
        return {
            "short_desc": self.desc_edit.toPlainText().strip(),
            "instruction_title": self.instruction_edit.text().strip(),
            "checklist": split_non_empty_lines(self.checklist_edit.toPlainText())
        }


class FeedbackManagerDialog(QDialog):
    """Диалог управления оценками и комментариями (для администратора)."""

    def __init__(self, db: KnowledgeBaseDB, instruction_id, parent=None):
        super().__init__(parent)
        self.db = db
        self.instruction_id = instruction_id
        self.setWindowTitle("Управление оценками и комментариями")
        self.setModal(True)
        self.resize(750, 550)

        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        # Вкладки: Оценки | Комментарии
        tabs = QTabWidget()

        # --- Оценки ---
        ratings_tab = QWidget()
        ratings_layout = QVBoxLayout(ratings_tab)
        self.ratings_list = QListWidget()
        self.ratings_list.setAlternatingRowColors(True)
        ratings_layout.addWidget(self.ratings_list)

        del_rating_btn = QPushButton("Удалить выбранную оценку")
        del_rating_btn.clicked.connect(self._delete_rating)
        ratings_layout.addWidget(del_rating_btn)

        tabs.addTab(ratings_tab, "Оценки")

        # --- Комментарии ---
        comments_tab = QWidget()
        comments_layout = QVBoxLayout(comments_tab)
        self.comments_list = QListWidget()
        self.comments_list.setAlternatingRowColors(True)
        comments_layout.addWidget(self.comments_list)

        comm_btn_layout = QHBoxLayout()
        edit_comm_btn = QPushButton("Редактировать")
        edit_comm_btn.clicked.connect(self._edit_comment)
        del_comm_btn = QPushButton("Удалить")
        del_comm_btn.clicked.connect(self._delete_comment)
        comm_btn_layout.addWidget(edit_comm_btn)
        comm_btn_layout.addWidget(del_comm_btn)
        comm_btn_layout.addStretch()
        comments_layout.addLayout(comm_btn_layout)

        tabs.addTab(comments_tab, "Комментарии")

        layout.addWidget(tabs)

        # Закрыть
        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(self.close)
        layout.addWidget(close_btn, alignment=Qt.AlignRight)

        self._load_data()

    def _load_data(self):
        # Оценки
        self.ratings_list.clear()
        ratings = self.db.ratings_for_instruction(self.instruction_id)
        for r in ratings:
            item = QListWidgetItem(f"{r['rating']} / 10  —  {r['created_at']}")
            item.setData(Qt.UserRole, r["id"])
            self.ratings_list.addItem(item)

        # Комментарии
        self.comments_list.clear()
        comments = self.db.comments_for_instruction(self.instruction_id)
        for c in comments:
            text = c["text"][:100] + "…" if len(c["text"]) > 100 else c["text"]
            item = QListWidgetItem(f"[{c['created_at']}] {c['author']}: {text}")
            item.setData(Qt.UserRole, c["id"])
            self.comments_list.addItem(item)

    def _delete_rating(self):
        item = self.ratings_list.currentItem()
        if not item:
            QMessageBox.warning(self, "Внимание", "Выберите оценку.")
            return
        if QMessageBox.question(self, "Подтверждение", "Удалить оценку?") == QMessageBox.Yes:
            self.db.delete_rating(item.data(Qt.UserRole))
            self._load_data()

    def _edit_comment(self):
        item = self.comments_list.currentItem()
        if not item:
            QMessageBox.warning(self, "Внимание", "Выберите комментарий.")
            return
        comment = self.db.comment_by_id(item.data(Qt.UserRole))
        if not comment:
            return
        text, ok = QInputDialog.getMultiLineText(self, "Редактировать", "Текст:", comment["text"])
        if ok and text.strip():
            self.db.update_comment(comment["id"], text.strip())
            self._load_data()

    def _delete_comment(self):
        item = self.comments_list.currentItem()
        if not item:
            QMessageBox.warning(self, "Внимание", "Выберите комментарий.")
            return
        if QMessageBox.question(self, "Подтверждение", "Удалить комментарий?") == QMessageBox.Yes:
            self.db.delete_comment(item.data(Qt.UserRole))
            self._load_data()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("База знаний по задачам")
        self.resize(1500, 920)
        self.setMinimumSize(1250, 760)

        self.db = KnowledgeBaseDB(DB_PATH)

        self.current_category = None
        self.username = getpass.getuser()
        self.is_admin = self.username.casefold() == ADMIN_USERNAME.casefold()
        self.banner_text = self.db.get_banner_text()

        self.current_category = None
        self.current_task = None
        self.current_instruction = None
        self.current_section_titles = []
        self.task_state_cache = {}

        self.search_timer = QTimer(self)
        self.search_timer.setSingleShot(True)
        self.search_timer.setInterval(250)
        self.search_timer.timeout.connect(self.reload_nav_tree)

        self.build_ui()
        self.apply_admin_mode()
        self.clear_views()
        self.reload_nav_tree()

    def build_ui(self):
        self.setStyleSheet(APP_QSS)

        root = QWidget()
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(12, 12, 12, 12)
        root_layout.setSpacing(10)

        self.banner_widget = QWidget()
        self.banner_widget.setObjectName("bannerWidget")
        banner_layout = QHBoxLayout(self.banner_widget)
        banner_layout.setContentsMargins(10, 8, 10, 8)
        banner_layout.setSpacing(8)

        self.banner_label = QLabel(self.banner_text)
        self.banner_label.setWordWrap(True)
        self.banner_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

        self.banner_edit_button = QToolButton()
        self.banner_edit_button.setText("✎")
        self.banner_edit_button.setToolTip("Редактировать верхний текст")
        self.banner_edit_button.clicked.connect(self.edit_banner_text)
        self.banner_edit_button.setVisible(self.is_admin)
        self.banner_edit_button.setFixedSize(28, 28)

        banner_layout.addWidget(self.banner_label, 1)
        banner_layout.addWidget(self.banner_edit_button, 0, Qt.AlignTop)

        self.banner_widget.setStyleSheet("""
            QWidget#bannerWidget {
                background: #eaf2ff;
                border: 1px solid #c6d8ff;
                border-radius: 8px;
            }

            QWidget#bannerWidget QLabel {
                background: transparent;
                border: none;
                color: #1f2937;
                font-size: 12px;
            }

            QWidget#bannerWidget QToolButton {
                background: transparent;
                border: none;
                color: #1f2937;
            }

            QWidget#bannerWidget QToolButton:hover {
                background: rgba(255, 255, 255, 0.25);
                border-radius: 6px;
            }
        """)
        self.banner_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Maximum)
        root_layout.addWidget(self.banner_widget)

        splitter = QSplitter(Qt.Horizontal)

        # Левая панель
        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(8)

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Поиск по задачам, инструкциям и тексту блоков...")
        self.search_edit.textChanged.connect(self.search_timer.start)

        self.tasks_group = QGroupBox("Задачи")
        tasks_layout = QVBoxLayout(self.tasks_group)
        tasks_layout.setContentsMargins(8, 8, 8, 8)
        tasks_layout.setSpacing(8)

        self.nav_tree = QTreeWidget()
        self.nav_tree.setHeaderHidden(True)
        self.nav_tree.setSelectionMode(QAbstractItemView.SingleSelection)
        self.nav_tree.setAlternatingRowColors(True)
        self.nav_tree.currentItemChanged.connect(self.on_nav_item_changed)

        tasks_layout.addWidget(self.search_edit)
        tasks_layout.addWidget(self.nav_tree)

        left_layout.addWidget(self.tasks_group)

        self.tree_admin_group = QGroupBox("Управление деревом")
        self.tree_admin_group.setVisible(self.is_admin)
        tree_admin_layout = QVBoxLayout(self.tree_admin_group)
        tree_admin_layout.setSpacing(6)

        row1 = QHBoxLayout()
        self.add_category_button = QPushButton("Добавить категорию")
        self.add_category_button.clicked.connect(self.add_tree_category)
        row1.addWidget(self.add_category_button)

        self.add_task_button = QPushButton("Добавить задачу")
        self.add_task_button.clicked.connect(self.add_tree_task)
        row1.addWidget(self.add_task_button)

        self.add_instruction_button = QPushButton("Добавить инструкцию")
        self.add_instruction_button.clicked.connect(self.add_tree_instruction)
        row1.addWidget(self.add_instruction_button)

        tree_admin_layout.addLayout(row1)

        row2 = QHBoxLayout()
        self.edit_tree_button = QPushButton("Редактировать")
        self.edit_tree_button.clicked.connect(self.edit_tree_item)
        row2.addWidget(self.edit_tree_button)

        self.delete_tree_button = QPushButton("Удалить")
        self.delete_tree_button.clicked.connect(self.delete_tree_item)
        row2.addWidget(self.delete_tree_button)

        self.save_tree_button = QPushButton("Сохранить")
        self.save_tree_button.clicked.connect(self.save_tree_changes)
        row2.addWidget(self.save_tree_button)

        tree_admin_layout.addLayout(row2)
        row3 = QHBoxLayout()
        self.move_up_button = QPushButton("⬆ Вверх")
        self.move_up_button.setToolTip("Переместить выбранный элемент вверх")
        self.move_up_button.clicked.connect(self.move_tree_item_up)
        row3.addWidget(self.move_up_button)

        self.move_down_button = QPushButton("⬇ Вниз")
        self.move_down_button.setToolTip("Переместить выбранный элемент вниз")
        self.move_down_button.clicked.connect(self.move_tree_item_down)
        row3.addWidget(self.move_down_button)

        tree_admin_layout.addLayout(row3)
        left_layout.addWidget(self.tree_admin_group)

        # Правая часть — вкладки
        self.tabs = QTabWidget()

        self.build_task_tab()
        self.build_instruction_tab()
        self.build_feedback_tab()

        self.build_help_tab()

        self.tabs.addTab(self.task_tab, "📝 Задача")
        self.tabs.addTab(self.instruction_tab, "📘 Инструкция")
        self.tabs.addTab(self.feedback_tab, "💬 Комментарии и оценка")
        self.tabs.addTab(self.help_tab, "❓ Справка")
        splitter.addWidget(left)
        splitter.addWidget(self.tabs)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 3)
        splitter.setSizes([360, 1040])

        root_layout.addWidget(splitter)
        self.setCentralWidget(root)

        self.statusBar().showMessage("Выберите задачу слева")

    def build_task_tab(self):
        self.task_tab = QWidget()
        layout = QVBoxLayout(self.task_tab)
        layout.setSpacing(10)

        header_row = QHBoxLayout()

        self.task_title_label = QLabel("Выберите задачу слева")
        self.task_title_label.setWordWrap(True)
        self.task_title_label.setFont(QFont("Segoe UI", 14, QFont.Bold))
        self.task_title_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

        header_row.addWidget(self.task_title_label, 1)
        self.edit_task_button = QPushButton("Редактировать задачу")
        self.edit_task_button.setToolTip("Редактировать описание, инструкцию и чек-лист")
        self.edit_task_button.setVisible(self.is_admin)
        self.edit_task_button.clicked.connect(self.open_task_editor)
        header_row.addWidget(self.edit_task_button)

        layout.addLayout(header_row)

        self.task_category_label = QLabel("")
        self.task_category_label.setStyleSheet("color: #5b6577;")

        # Описание
        self.task_desc_group = QGroupBox("Описание")
        self.task_desc_group.setVisible(False)
        self.task_desc_group.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        desc_layout_inner = QVBoxLayout(self.task_desc_group)
        self.task_desc_label = QLabel("")
        self.task_desc_label.setWordWrap(True)
        self.task_desc_label.setTextFormat(Qt.RichText)
        self.task_desc_label.setOpenExternalLinks(True)
        self.task_desc_label.linkActivated.connect(self.handle_link_activated)
        self.task_desc_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.task_desc_label.setStyleSheet("color: #263238; background: transparent; border: none;")
        desc_layout_inner.addWidget(self.task_desc_label)

        self.task_instruction_label = QLabel("")
        self.task_instruction_label.setWordWrap(True)
        self.task_instruction_label.setStyleSheet("font-weight: 600; color: #1f2937;")

        layout.addWidget(self.task_category_label)
        layout.addWidget(self.task_desc_group)

        self.category_tasks_group = QGroupBox("Задачи")
        self.category_tasks_group.setVisible(False)
        self.category_tasks_layout = QVBoxLayout(self.category_tasks_group)
        self.category_tasks_layout.setContentsMargins(8, 8, 8, 8)
        self.category_tasks_layout.setSpacing(6)
        layout.addWidget(self.category_tasks_group)

        layout.addWidget(self.task_instruction_label)

        checklist_group = QGroupBox("Чек-лист")
        self.task_checklist_group = checklist_group
        self.task_checklist_group.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        checklist_layout = QVBoxLayout(checklist_group)

        self.task_checklist = QListWidget()
        self.task_checklist.itemChanged.connect(self.on_checklist_item_changed)
        self.task_checklist.setSelectionMode(QAbstractItemView.NoSelection)
        self.task_checklist.setAlternatingRowColors(True)
        self.task_checklist.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.task_checklist.setMinimumHeight(80)
        self.task_checklist.setSizeAdjustPolicy(QAbstractItemView.AdjustToContents)

        checklist_layout.addWidget(self.task_checklist)

        layout.addWidget(checklist_group)

        button_row = QHBoxLayout()
        button_row.addStretch()

        self.open_instruction_button = QPushButton("Открыть инструкцию")
        self.open_instruction_button.clicked.connect(self.open_current_instruction)
        button_row.addWidget(self.open_instruction_button)

        layout.addLayout(button_row)
        layout.addStretch()

    def build_instruction_tab(self):
        self.instruction_tab = QWidget()
        layout = QVBoxLayout(self.instruction_tab)
        layout.setSpacing(10)

        # Заголовок с кнопкой редактирования
        header_row = QHBoxLayout()
        self.instruction_title_label = QLabel("Инструкция")
        self.instruction_title_label.setWordWrap(True)
        self.instruction_title_label.setFont(QFont("Segoe UI", 14, QFont.Bold))

        header_row.addWidget(self.instruction_title_label, 1)
        header_row.addStretch()

        self.edit_instruction_btn = QPushButton("Редактировать инструкцию")
        self.edit_instruction_btn.setToolTip("Изменить содержание инструкции")
        self.edit_instruction_btn.setVisible(False)  # будет показана, когда загружена инструкция
        self.edit_instruction_btn.clicked.connect(self.edit_current_instruction)
        header_row.addWidget(self.edit_instruction_btn)

        layout.addLayout(header_row)

        self.instruction_desc_label = QLabel("")
        self.instruction_desc_label.setWordWrap(True)
        self.instruction_desc_label.setStyleSheet("color: #5b6577;")
        layout.addWidget(self.instruction_desc_label)
        self.instruction_scroll = QScrollArea()
        self.instruction_scroll.setWidgetResizable(True)
        self.instruction_scroll.setFrameShape(QFrame.NoFrame)

        # --- Sticky header (приклеенный заголовок) ---
        self.sticky_header_label = QLabel()
        self.sticky_header_label.setVisible(False)
        self.sticky_header_label.setWordWrap(False)
        self.sticky_header_label.setStyleSheet("""
            QLabel {
                background: #eef5fb;
                color: #263238;
                font-size: 15px;
                font-weight: bold;
                padding: 10px 12px;
                border: 1px solid #d9e3f0;
                border-radius: 10px;
            }
        """)
        layout.addWidget(self.sticky_header_label)

        self.instruction_container = QWidget()
        self.instruction_container_layout = QVBoxLayout(self.instruction_container)
        self.instruction_container_layout.setContentsMargins(0, 0, 0, 0)
        self.instruction_container_layout.setSpacing(10)
        self.instruction_container_layout.setAlignment(Qt.AlignTop)

        self.instruction_scroll.setWidget(self.instruction_container)
        layout.addWidget(self.instruction_scroll)

        # Подключаемся к сигналу прокрутки
        self.instruction_scroll.verticalScrollBar().valueChanged.connect(self._update_sticky_header)

    def build_feedback_tab(self):
        self.feedback_tab = QWidget()
        layout = QVBoxLayout(self.feedback_tab)
        layout.setSpacing(10)

        # --- Заголовок ---
        title = QLabel("Комментарии и оценка к инструкции")
        title.setFont(QFont("Segoe UI", 14, QFont.Bold))
        layout.addWidget(title)

        note = QLabel("Оценка анонимная. Старайся быть честным — так инструкция станет лучше.")
        note.setStyleSheet("color: #5b6577;")
        note.setWordWrap(True)
        layout.addWidget(note)

        # --- Статистика ---
        self.feedback_stats_label = QLabel("Выберите инструкцию слева")
        self.feedback_stats_label.setFont(QFont("Segoe UI", 12, QFont.Bold))
        layout.addWidget(self.feedback_stats_label)

        # --- Оценка ---
        rating_layout = QHBoxLayout()
        rating_layout.addWidget(QLabel("Оценка (1–10):"))
        self.rating_spin = QSpinBox()
        self.rating_spin.setRange(1, 10)
        self.rating_spin.setValue(10)
        self.rating_spin.setMaximumWidth(90)
        rating_layout.addWidget(self.rating_spin)
        rating_layout.addStretch()
        layout.addLayout(rating_layout)

        hint_label = QLabel("Где 1 — очень плохо, а 10 — отличная и самодостаточная инструкция")
        hint_label.setWordWrap(True)
        hint_label.setStyleSheet("font-size: 11px; color: #5b6577; background: transparent; border: none;")
        layout.addWidget(hint_label)

        # --- Комментарий ---
        comment_title = QLabel("Комментарий к инструкции")
        comment_title.setFont(QFont("Segoe UI", 13, QFont.Bold))
        layout.addWidget(comment_title)

        top_row = QHBoxLayout()
        top_row.addWidget(QLabel("Блок:"))
        self.comment_anchor_box = QComboBox()
        self.comment_anchor_box.setMinimumWidth(280)
        top_row.addWidget(self.comment_anchor_box)
        self.comment_author_edit = QLineEdit()
        self.comment_author_edit.setPlaceholderText("Имя (необязательно)")
        top_row.addWidget(self.comment_author_edit)
        layout.addLayout(top_row)

        comment_note = QLabel("Можно оставить комментарий к целой инструкции или к конкретному блоку.")
        comment_note.setStyleSheet("color: #5b6577;")
        comment_note.setWordWrap(True)
        layout.addWidget(comment_note)

        self.comment_text_edit = QTextEdit()
        self.comment_text_edit.setPlaceholderText("Что непонятно? Что нужно поправить или доработать?")
        self.comment_text_edit.setMinimumHeight(100)
        layout.addWidget(self.comment_text_edit)

        # --- Кнопки ---
        btn_row = QHBoxLayout()

        self.submit_feedback_btn = QPushButton("Поставить оценку и добавить комментарий")
        self.submit_feedback_btn.clicked.connect(self.submit_feedback)
        btn_row.addWidget(self.submit_feedback_btn)

        self.manage_feedback_btn = QPushButton("Редактировать")
        self.manage_feedback_btn.setToolTip("Редактировать и удалять оценки и комментарии")
        self.manage_feedback_btn.setVisible(self.is_admin)
        self.manage_feedback_btn.clicked.connect(self.open_feedback_manager)
        btn_row.addWidget(self.manage_feedback_btn)

        btn_row.addStretch()
        layout.addLayout(btn_row)

        # --- Экспорт (только админ) ---
        self.export_feedback_button = QPushButton("Экспорт в Excel")
        self.export_feedback_button.clicked.connect(self.export_feedback_to_excel)
        self.export_feedback_button.setVisible(self.is_admin)
        layout.addWidget(self.export_feedback_button, alignment=Qt.AlignRight)

        # --- Список комментариев (видимый всем) ---
        self.feedback_comments_browser = QTextBrowser()
        self.feedback_comments_browser.setOpenExternalLinks(False)
        self.feedback_comments_browser.setStyleSheet("""
            QTextBrowser {
                background: white;
                border: 1px solid #d8dee9;
                border-radius: 8px;
            }
        """)
        layout.addWidget(self.feedback_comments_browser)

        layout.addStretch()

    def build_help_tab(self):
        self.help_tab = QWidget()
        layout = QVBoxLayout(self.help_tab)
        layout.setSpacing(10)

        # Заголовок с кнопкой редактирования для админа
        header_row = QHBoxLayout()
        help_title = QLabel("Справка")
        help_title.setFont(QFont("Segoe UI", 14, QFont.Bold))
        header_row.addWidget(help_title, 1)
        header_row.addStretch()

        self.edit_help_btn = QPushButton("Редактировать справку")
        self.edit_help_btn.setToolTip("Изменить текст справки")
        self.edit_help_btn.setVisible(self.is_admin)
        self.edit_help_btn.clicked.connect(self.edit_help_content)
        header_row.addWidget(self.edit_help_btn)

        layout.addLayout(header_row)

        # Область просмотра справки
        self.help_browser = QTextBrowser()
        self.help_browser.setOpenExternalLinks(False)
        self.help_browser.anchorClicked.connect(self.handle_link_activated)
        self.help_browser.setStyleSheet("""
            QTextBrowser {
                background: white;
                border: 1px solid #d8dee9;
                border-radius: 8px;
                padding: 12px;
            }
        """)
        layout.addWidget(self.help_browser)

        # Загружаем содержимое
        self.refresh_help_tab()

    def build_comments_tab(self):
        self.comments_tab = QWidget()
        layout = QVBoxLayout(self.comments_tab)
        layout.setSpacing(10)

        title = QLabel("Комментарии к инструкции")
        title.setFont(QFont("Segoe UI", 13, QFont.Bold))

        note = QLabel("Можно оставить комментарий к целой инструкции или к конкретному блоку.")
        note.setStyleSheet("color: #5b6577;")
        note.setWordWrap(True)

        layout.addWidget(title)
        layout.addWidget(note)

        self.comments_browser = QTextBrowser()
        self.comments_browser.setOpenExternalLinks(False)
        self.comments_browser.setStyleSheet("""
            QTextBrowser {
                background: white;
                border: 1px solid #d8dee9;
                border-radius: 8px;
            }
        """)
        layout.addWidget(self.comments_browser)

        self.comment_form_widget = QWidget()
        form_layout = QVBoxLayout(self.comment_form_widget)
        form_layout.setSpacing(8)

        top_row = QHBoxLayout()

        top_row.addWidget(QLabel("Блок:"))
        self.comment_anchor_box = QComboBox()
        self.comment_anchor_box.setMinimumWidth(280)
        top_row.addWidget(self.comment_anchor_box)

        top_row.addSpacing(14)
        self.comment_anonymous_check = QCheckBox("Анонимно")
        self.comment_anonymous_check.setChecked(True)
        self.comment_anonymous_check.toggled.connect(
            lambda checked: self.comment_author_edit.setEnabled(not checked)
        )
        top_row.addWidget(self.comment_anonymous_check)

        self.comment_author_edit = QLineEdit()
        self.comment_author_edit.setPlaceholderText("Имя (необязательно)")
        self.comment_author_edit.setEnabled(False)
        top_row.addWidget(self.comment_author_edit)

        form_layout.addLayout(top_row)

        self.comment_text_edit = QTextEdit()
        self.comment_text_edit.setPlaceholderText("Что непонятно? Что нужно поправить или доработать?")
        self.comment_text_edit.setMinimumHeight(90)
        form_layout.addWidget(self.comment_text_edit)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.add_comment_button = QPushButton("Добавить комментарий")
        self.add_comment_button.clicked.connect(self.add_comment)
        btn_row.addWidget(self.add_comment_button)
        form_layout.addLayout(btn_row)

        layout.addWidget(self.comment_form_widget)

    def build_rating_tab(self):
        self.rating_tab = QWidget()
        layout = QVBoxLayout(self.rating_tab)
        layout.setSpacing(10)

        title = QLabel("Анонимная оценка понятности")
        title.setFont(QFont("Segoe UI", 13, QFont.Bold))

        note = QLabel("Оценка анонимная. Старайся быть честным — так инструкция станет лучше.")
        note.setStyleSheet("color: #5b6577;")
        note.setWordWrap(True)

        layout.addWidget(title)
        layout.addWidget(note)

        self.rating_summary_label = QLabel("Выберите инструкцию слева")
        self.rating_summary_label.setWordWrap(True)
        self.rating_summary_label.setFont(QFont("Segoe UI", 12, QFont.Bold))
        layout.addWidget(self.rating_summary_label)

        self.rating_form_widget = QWidget()
        form_layout = QHBoxLayout(self.rating_form_widget)
        form_layout.setContentsMargins(0, 0, 0, 0)
        form_layout.setSpacing(10)

        form_layout.addWidget(QLabel("Оценка (1–5):"))
        self.rating_spin = QSpinBox()
        self.rating_spin.setRange(1, 5)
        self.rating_spin.setValue(5)
        self.rating_spin.setMaximumWidth(90)
        form_layout.addWidget(self.rating_spin)

        self.add_rating_button = QPushButton("Поставить оценку")
        self.add_rating_button.clicked.connect(self.add_rating)
        form_layout.addWidget(self.add_rating_button)

        form_layout.addStretch()
        layout.addWidget(self.rating_form_widget)
        layout.addStretch()

    # ================== Навигация ==================

    def reload_nav_tree(self):
        search = self.search_edit.text().strip()
        tasks = self.db.search_tasks(search)

        categories = self.db.conn.execute("""
            SELECT id, name, sort_order
            FROM categories
            ORDER BY sort_order, name
        """).fetchall()

        grouped = {c["id"]: [] for c in categories}
        for task in tasks:
            grouped.setdefault(task["category_id"], []).append(task)

        self.nav_tree.blockSignals(True)
        self.nav_tree.clear()

        item_map = {}
        first_task_item = None
        first_category_item = None

        for cat in categories:
            visible_tasks = grouped.get(cat["id"], [])

            if search and not visible_tasks:
                continue

            cat_item = QTreeWidgetItem([cat["name"]])
            cat_item.setData(0, ROLE_KIND, "category")
            cat_item.setData(0, ROLE_ID, cat["id"])
            font = cat_item.font(0)
            font.setBold(True)
            cat_item.setFont(0, font)

            item_map[cat["id"]] = cat_item
            if first_category_item is None:
                first_category_item = cat_item

            for task in visible_tasks:
                task_item = QTreeWidgetItem([task["task_title"]])
                task_item.setData(0, ROLE_KIND, "task")
                task_item.setData(0, ROLE_ID, task["task_id"])
                cat_item.addChild(task_item)
                item_map[task["task_id"]] = task_item
                if first_task_item is None:
                    first_task_item = task_item

            if visible_tasks:
                cat_item.setExpanded(True)

            self.nav_tree.addTopLevelItem(cat_item)

        if not item_map:
            empty = QTreeWidgetItem(["Ничего не найдено"])
            empty.setFlags(Qt.NoItemFlags)
            self.nav_tree.addTopLevelItem(empty)

        self.nav_tree.blockSignals(False)

        target_task_id = self.current_task["task_id"] if self.current_task else None
        target_category_id = self.current_category["id"] if self.current_category else None

        if target_task_id and target_task_id in item_map:
            self.nav_tree.setCurrentItem(item_map[target_task_id])
        elif target_category_id and target_category_id in item_map:
            self.nav_tree.setCurrentItem(item_map[target_category_id])
        elif first_task_item:
            self.nav_tree.setCurrentItem(first_task_item)
        elif first_category_item:
            self.nav_tree.setCurrentItem(first_category_item)
        else:
            self.clear_views()
            self.statusBar().showMessage("Ничего не найдено")

    def on_nav_item_changed(self, current, previous):
        if not current:
            self.clear_views()
            self.update_tree_admin_controls()
            return

        kind = current.data(0, ROLE_KIND)
        entity_id = current.data(0, ROLE_ID)

        if kind == "task":
            self.show_task(entity_id)
            self.tabs.setCurrentIndex(0)

        elif kind == "category":
            self.show_category(entity_id)
            self.tabs.setCurrentIndex(0)

        else:
            self.clear_views()

        self.update_tree_admin_controls()

    # ================== Отображение контекста ==================

    def clear_views(self):
        self.current_category = None
        self.current_task = None
        self.current_instruction = None
        self.current_section_titles = []

        self.task_title_label.setText("Выберите задачу слева")
        self.task_category_label.setText("")
        self.task_instruction_label.setText("")

        self.task_checklist.blockSignals(True)
        self.task_checklist.clear()
        self.task_checklist.blockSignals(False)
        self.task_checklist.setEnabled(False)
        self.open_instruction_button.setEnabled(False)

        if hasattr(self, "task_checklist_group"):
            self.task_checklist_group.setVisible(False)
        if hasattr(self, "category_tasks_group"):
            self.category_tasks_group.setVisible(False)
        self.task_instruction_label.setVisible(False)
        self.open_instruction_button.setVisible(False)
        if hasattr(self, "task_desc_group"):
            self.task_desc_group.setVisible(False)
        if hasattr(self, 'sticky_header_label'):
            self.sticky_header_label.setVisible(False)
        if hasattr(self, '_section_cards'):
            self._section_cards = []

        self.instruction_title_label.setText("Инструкция")
        self.instruction_desc_label.setText("Выберите задачу или инструкцию слева.")
        clear_layout(self.instruction_container_layout)
        placeholder = QLabel("Здесь будет подробная инструкция с раскрывающимися блоками.")
        placeholder.setWordWrap(True)
        self.instruction_container_layout.addWidget(placeholder)

        if hasattr(self, 'feedback_stats_label'):
            self.feedback_stats_label.setText("Выберите инструкцию слева")
        if hasattr(self, 'submit_feedback_btn'):
            self.submit_feedback_btn.setEnabled(False)
        if hasattr(self, 'feedback_comments_browser'):
            self.feedback_comments_browser.setHtml("<i>Выберите инструкцию слева.</i>")

        self.update_tree_admin_controls()
        self.statusBar().showMessage("Выберите задачу слева")

    def show_category(self, category_id):
        category = self.db.category_by_id(category_id)
        if not category:
            return

        self.current_category = category
        self.current_task = None
        self.current_instruction = None
        self.current_section_titles = []

        count = self.db.count_tasks_in_category(category_id)

        self.task_title_label.setText(category["name"])
        self.task_category_label.setText(f"Раздел: {category['name']}")
        if hasattr(self, 'task_desc_group'):
            self.task_desc_group.setVisible(True)
            self.task_desc_label.setText(render_markdown(
                f"В этом разделе задач: {count}. Выбери конкретную задачу слева или ниже по кнопке."
            ))
        self.task_instruction_label.setText("Инструкция: —")

        self.task_checklist.blockSignals(True)
        self.task_checklist.clear()
        self.task_checklist.blockSignals(False)
        self.task_checklist.setEnabled(False)
        self.open_instruction_button.setEnabled(False)

        self.task_instruction_label.setVisible(False)
        self.task_checklist_group.setVisible(False)
        self.open_instruction_button.setVisible(False)

        if hasattr(self, "category_tasks_group"):
            self.category_tasks_group.setVisible(True)
            self.refresh_category_task_buttons()

        self.instruction_title_label.setText("Инструкция")
        self.instruction_desc_label.setText("Категория выбрана. Теперь открой конкретную задачу.")
        clear_layout(self.instruction_container_layout)
        placeholder = QLabel("Подробная инструкция появится после выбора задачи.")
        placeholder.setWordWrap(True)
        self.instruction_container_layout.addWidget(placeholder)

        self.refresh_feedback_tab()

        self.statusBar().showMessage(f"Категория: {category['name']}")

    def show_task(self, task_id):
        task = self.db.task_bundle(task_id)
        if not task:
            return

        self.current_task = task
        self.current_category = self.db.category_by_id(task["category_id"])
        self.current_instruction = None
        if task.get("instruction_id"):
            self.current_instruction = self.db.instruction_by_id(task["instruction_id"])
        self.current_section_titles = [s.get("title", "") for s in
                                       (self.current_instruction["sections"] if self.current_instruction else [])]

        # Задача
        self.task_title_label.setText(task["task_title"])
        self.task_category_label.setText(f"Раздел: {task['category_name']}")
        if hasattr(self, 'task_desc_group'):
            self.task_desc_group.setVisible(True)
            html_desc = render_markdown(task["short_desc"])
            self.task_desc_label.setText(html_desc)
        if self.current_instruction:
            self.task_instruction_label.setText(f"Инструкция: {self.current_instruction['instruction_title']}")
        else:
            self.task_instruction_label.setText("Инструкция: не привязана")

        self.task_checklist.blockSignals(True)
        self.task_checklist.clear()

        saved_states = self.task_state_cache.get(task["task_id"], [])
        for idx, item_text in enumerate(task["checklist"]):
            item = QListWidgetItem(item_text)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            state = saved_states[idx] if idx < len(saved_states) else Qt.Unchecked
            item.setCheckState(state)
            self.task_checklist.addItem(item)

        self.task_checklist.blockSignals(False)
        self.task_checklist.setEnabled(True)
        self.open_instruction_button.setEnabled(self.current_instruction is not None)

        self.task_instruction_label.setVisible(True)
        self.task_checklist_group.setVisible(True)
        self.open_instruction_button.setVisible(True)
        if hasattr(self, "category_tasks_group"):
            self.category_tasks_group.setVisible(False)

        # Инструкция
        self.refresh_instruction_tab()
        # Комментарии и оценка
        self.refresh_feedback_tab()

        self.tabs.setCurrentIndex(0)
        self.statusBar().showMessage(f"Открыта задача: {task['task_title']}")

    def _update_sticky_header(self):
        """Обновляет приклеенный заголовок на основе позиции прокрутки."""
        if not hasattr(self, '_section_cards') or not self._section_cards:
            self.sticky_header_label.setVisible(False)
            return

        scroll_y = self.instruction_scroll.verticalScrollBar().value()
        sticky_title = None

        for card in self._section_cards:
            # Позиция кнопки-заголовка относительно контейнера
            toggle_pos = card.toggle.mapTo(self.instruction_container, card.toggle.pos())
            toggle_top = toggle_pos.y()

            if toggle_top <= scroll_y:
                sticky_title = card.toggle.text()
            else:
                break  # дальше заголовки ещё не видны

        if sticky_title:
            self.sticky_header_label.setText(sticky_title)
            self.sticky_header_label.setVisible(True)
        else:
            self.sticky_header_label.setVisible(False)

    def show_instruction_only(self, instruction_id):
        instruction = self.db.instruction_by_id(instruction_id)
        if not instruction:
            return

        self.current_instruction = instruction
        self.current_category = self.db.category_by_id(instruction["category_id"])
        self.current_task = None
        self.current_section_titles = [s.get("title", "") for s in instruction["sections"]]

        self.refresh_task_tab()
        self.refresh_instruction_tab()
        self.refresh_feedback_tab()
        self.tabs.setCurrentIndex(1)
        self.statusBar().showMessage(f"Открыта инструкция: {instruction['instruction_title']}")

    def move_tree_item_up(self):
        """Перемещает выбранный элемент дерева вверх."""
        if not self.is_admin:
            return
        kind, entity_id = self.selected_tree_entity()
        if kind == "category":
            self.db.move_category_up(entity_id)
        elif kind == "task":
            self.db.move_task_up(entity_id)
        else:
            return
        self.reload_nav_tree()

    def move_tree_item_down(self):
        """Перемещает выбранный элемент дерева вниз."""
        if not self.is_admin:
            return
        kind, entity_id = self.selected_tree_entity()
        if kind == "category":
            self.db.move_category_down(entity_id)
        elif kind == "task":
            self.db.move_task_down(entity_id)
        else:
            return
        self.reload_nav_tree()

    # ================== Вкладка задачи ==================

    def open_task_editor(self):
        """Открывает отдельное окно для редактирования задачи."""
        if not self.is_admin or not self.current_task:
            return

        saved_task_id = self.current_task["task_id"]

        dialog = TaskEditorDialog(
            db=self.db,
            task=self.current_task,
            instruction=self.current_instruction,
            parent=self
        )

        if dialog.exec() != QDialog.Accepted:
            return

        data = dialog.get_data()
        try:
            self.db.update_task_view_data(
                saved_task_id,
                data["short_desc"],
                data["instruction_title"],
                data["checklist"]
            )
        except Exception as exc:
            QMessageBox.warning(self, "Ошибка", f"Не удалось сохранить изменения: {exc}")
            return

        self.reload_nav_tree()
        self.show_task(saved_task_id)
        self.statusBar().showMessage("Задача обновлена")

    def refresh_task_tab(self):
        if self.current_task:
            self.task_title_label.setText(self.current_task["task_title"])
            self.task_category_label.setText(f"Раздел: {self.current_task['category_name']}")
            if hasattr(self, 'task_desc_group'):
                self.task_desc_group.setVisible(True)
                self.task_desc_label.setText(render_markdown(self.current_task["short_desc"]))
            if self.current_instruction:
                self.task_instruction_label.setText(f"Инструкция: {self.current_instruction['instruction_title']}")
            else:
                self.task_instruction_label.setText("Инструкция: не привязана")
            self.task_checklist.setEnabled(True)
            self.open_instruction_button.setEnabled(self.current_instruction is not None)

            self.task_instruction_label.setVisible(True)
            self.task_checklist_group.setVisible(True)
            self.open_instruction_button.setVisible(True)
            if hasattr(self, "category_tasks_group"):
                self.category_tasks_group.setVisible(False)
            return

        if self.current_instruction and not self.current_task:
            self.task_title_label.setText(self.current_instruction["instruction_title"])
            self.task_category_label.setText(f"Раздел: {self.current_instruction['category_name']}")
            if hasattr(self, 'task_desc_group'):
                self.task_desc_group.setVisible(True)
                self.task_desc_label.setText(render_markdown(self.current_instruction["short_desc"]))
            self.task_instruction_label.setText("Открыто по ссылке без отдельной задачи")
            self.task_checklist.blockSignals(True)
            self.task_checklist.clear()
            self.task_checklist.blockSignals(False)
            self.task_checklist.setEnabled(False)
            self.open_instruction_button.setEnabled(True)

            self.task_instruction_label.setVisible(True)
            self.task_checklist_group.setVisible(False)
            self.open_instruction_button.setVisible(True)
            if hasattr(self, "category_tasks_group"):
                self.category_tasks_group.setVisible(False)
            return

        if self.current_category:
            self.task_title_label.setText(self.current_category["name"])
            self.task_category_label.setText(f"Раздел: {self.current_category['name']}")
            if hasattr(self, 'task_desc_group'):
                self.task_desc_group.setVisible(True)
                self.task_desc_label.setText(render_markdown(
                    f"В этом разделе задач: {self.current_category['task_count']}. "
                    "Выбери конкретную задачу слева или ниже по кнопке."
                ))
            self.task_instruction_label.setText("Выберите конкретную задачу слева")
            self.task_checklist.blockSignals(True)
            self.task_checklist.clear()
            self.task_checklist.blockSignals(False)
            self.task_checklist.setEnabled(False)
            self.open_instruction_button.setEnabled(False)

            self.task_instruction_label.setVisible(False)
            self.task_checklist_group.setVisible(False)
            self.open_instruction_button.setVisible(False)
            if hasattr(self, "category_tasks_group"):
                self.category_tasks_group.setVisible(True)
                self.refresh_category_task_buttons()
            return

        self.task_title_label.setText("Выберите задачу слева")
        self.task_category_label.setText("")
        if hasattr(self, 'task_desc_group'):
            self.task_desc_group.setVisible(False)
        self.task_instruction_label.setText("")
        self.task_checklist.setEnabled(False)
        self.open_instruction_button.setEnabled(False)

    def on_checklist_item_changed(self, item):
        if not self.current_task:
            return

        states = [self.task_checklist.item(i).checkState() for i in range(self.task_checklist.count())]
        self.task_state_cache[self.current_task["task_id"]] = states

    def open_current_instruction(self):
        if self.current_instruction:
            self.tabs.setCurrentIndex(1)
        else:
            QMessageBox.warning(self, "Внимание", "Для этой задачи пока не привязана инструкция.")

    def refresh_category_task_buttons(self):
        if not hasattr(self, "category_tasks_layout"):
            return

        clear_layout(self.category_tasks_layout)

        if not self.current_category:
            return

        tasks = self.db.tasks_for_category(self.current_category["id"])
        if not tasks:
            lbl = QLabel("В этом разделе пока нет задач.")
            lbl.setWordWrap(True)
            lbl.setStyleSheet("color: #5b6577;")
            self.category_tasks_layout.addWidget(lbl)
            return

        for task in tasks:
            btn = QPushButton(short_button_text(task["task_title"]))
            btn.setToolTip(task["task_title"])
            btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            btn.clicked.connect(lambda _, tid=task["task_id"]: self.show_task(tid))
            self.category_tasks_layout.addWidget(btn)

    def apply_admin_mode(self):
        if hasattr(self, "banner_edit_button"):
            self.banner_edit_button.setVisible(self.is_admin)
        if hasattr(self, "tree_admin_group"):
            self.tree_admin_group.setVisible(self.is_admin)
        if hasattr(self, "edit_task_button"):
            self.edit_task_button.setVisible(self.is_admin)
        if hasattr(self, "export_feedback_button"):
            self.export_feedback_button.setVisible(self.is_admin)
        if hasattr(self, "edit_help_btn"):
            self.edit_help_btn.setVisible(self.is_admin)

        if self.is_admin:
            self.statusBar().showMessage(f"Администратор: {self.username}")

        if hasattr(self, "manage_feedback_btn"):
            self.manage_feedback_btn.setVisible(self.is_admin)

        if hasattr(self, "move_up_button"):
            self.move_up_button.setVisible(self.is_admin)
        if hasattr(self, "move_down_button"):
            self.move_down_button.setVisible(self.is_admin)

        self.update_tree_admin_controls()

    def edit_banner_text(self):
        if not self.is_admin:
            return

        text, ok = QInputDialog.getMultiLineText(
            self,
            "Редактирование верхнего текста",
            "Текст в верхней плашке:",
            self.banner_label.text()
        )
        if not ok:
            return

        text = text.strip()
        if not text:
            QMessageBox.warning(self, "Внимание", "Текст не может быть пустым.")
            return

        self.banner_text = text
        self.banner_label.setText(text)
        self.banner_label.adjustSize()
        self.db.save_banner_text(text)
        self.banner_widget.adjustSize()

    def selected_tree_entity(self):
        item = self.nav_tree.currentItem()
        if not item:
            return None, None
        return item.data(0, ROLE_KIND), item.data(0, ROLE_ID)

    def update_tree_admin_controls(self):
        if not hasattr(self, "add_category_button"):
            return
        if not self.is_admin:
            return

        kind, _ = self.selected_tree_entity()
        can_modify = kind in {"category", "task"}
        can_add_item = kind in {"category", "task"}

        self.add_category_button.setEnabled(True)
        self.add_task_button.setEnabled(can_add_item)
        self.add_instruction_button.setEnabled(can_add_item)
        self.edit_tree_button.setEnabled(can_modify)
        self.delete_tree_button.setEnabled(can_modify)
        can_move = kind in {"category", "task"}
        if hasattr(self, "move_up_button"):
            self.move_up_button.setEnabled(can_move)
        if hasattr(self, "move_down_button"):
            self.move_down_button.setEnabled(can_move)
        self.save_tree_button.setEnabled(True)

    def add_tree_category(self):
        if not self.is_admin:
            return

        name, ok = QInputDialog.getText(self, "Добавить категорию", "Название категории:")
        if not ok:
            return

        name = name.strip()
        if not name:
            QMessageBox.warning(self, "Внимание", "Название категории не может быть пустым.")
            return

        try:
            self.db.add_category(name)
        except Exception as exc:
            QMessageBox.warning(self, "Ошибка", f"Не удалось добавить категорию: {exc}")
            return

        self.reload_nav_tree()

    def add_tree_task(self):
        if not self.is_admin:
            return

        kind, entity_id = self.selected_tree_entity()
        if kind == "category":
            category_id = entity_id
        elif kind == "task":
            task = self.db.task_bundle(entity_id)
            category_id = task["category_id"] if task else None
        else:
            category_id = None

        if not category_id:
            QMessageBox.warning(self, "Внимание", "Выберите категорию или задачу, чтобы добавить новую задачу.")
            return

        title, ok = QInputDialog.getText(self, "Добавить задачу", "Название задачи:")
        if not ok:
            return

        title = title.strip()
        if not title:
            QMessageBox.warning(self, "Внимание", "Название задачи не может быть пустым.")
            return

        instruction_title, ok = QInputDialog.getText(
            self,
            "Добавить задачу",
            "Название инструкции (необязательно):"
        )
        if not ok:
            return

        try:
            self.db.add_task(category_id, title, instruction_title)
        except Exception as exc:
            QMessageBox.warning(self, "Ошибка", f"Не удалось добавить задачу: {exc}")
            return

        self.reload_nav_tree()

    def add_tree_instruction(self):
        if not self.is_admin:
            return

        kind, entity_id = self.selected_tree_entity()
        if kind == "category":
            category_id = entity_id
        elif kind == "task":
            task = self.db.task_bundle(entity_id)
            category_id = task["category_id"] if task else None
        else:
            category_id = None

        categories = self.db.conn.execute("""
            SELECT id, name
            FROM categories
            ORDER BY sort_order, name
        """).fetchall()

        if not categories:
            QMessageBox.warning(self, "Внимание", "Сначала добавь хотя бы одну категорию.")
            return

        dialog = InstructionEditorDialog(
            db=self.db,
            categories=[dict(row) for row in categories],
            default_category_id=category_id,
            parent=self
        )

        if dialog.exec() != QDialog.Accepted:
            return

        data = dialog.get_data()
        category_id = data["category_id"]
        task_id = data["task_id"]
        instruction_title = data["title"]
        short_desc = data["short_desc"]
        sections = data["sections"]
        related_titles = data["related_titles"]

        if not task_id:
            QMessageBox.warning(self, "Ошибка", "Не указана задача для привязки инструкции.")
            return

        if not instruction_title:
            QMessageBox.warning(self, "Внимание", "Название инструкции не может быть пустым.")
            return

        if not short_desc:
            short_desc = instruction_title

        try:
            instruction_id, missing_related = self.db.add_instruction(
                category_id,
                task_id,
                instruction_title,
                short_desc,
                sections,
                related_titles
            )
        except Exception as exc:
            QMessageBox.warning(self, "Ошибка", f"Не удалось добавить инструкцию: {exc}")
            return

        self.reload_nav_tree()
        self.show_task(task_id)

        if missing_related:
            QMessageBox.information(
                self,
                "Внимание",
                "Не найдены связанные инструкции:\n" + "\n".join(missing_related)
            )

    def edit_tree_item(self):
        if not self.is_admin:
            return

        kind, entity_id = self.selected_tree_entity()
        if kind == "category":
            category = self.db.category_by_id(entity_id)
            if not category:
                return

            new_name, ok = QInputDialog.getText(
                self,
                "Редактировать категорию",
                "Новое название категории:",
                QLineEdit.Normal,
                category["name"]
            )
            if not ok:
                return

            new_name = new_name.strip()
            if not new_name:
                QMessageBox.warning(self, "Внимание", "Название категории не может быть пустым.")
                return

            try:
                self.db.rename_category(entity_id, new_name)
            except Exception as exc:
                QMessageBox.warning(self, "Ошибка", f"Не удалось изменить категорию: {exc}")
                return

        elif kind == "task":
            task = self.db.task_bundle(entity_id)
            if not task:
                return

            new_name, ok = QInputDialog.getText(
                self,
                "Редактировать задачу",
                "Новое название задачи:",
                QLineEdit.Normal,
                task["task_title"]
            )
            if not ok:
                return

            new_name = new_name.strip()
            if not new_name:
                QMessageBox.warning(self, "Внимание", "Название задачи не может быть пустым.")
                return

            try:
                self.db.rename_task(entity_id, new_name)
            except Exception as exc:
                QMessageBox.warning(self, "Ошибка", f"Не удалось изменить задачу: {exc}")
                return
        else:
            QMessageBox.warning(self, "Внимание", "Сначала выберите категорию или задачу.")
            return

        self.reload_nav_tree()

    def delete_tree_item(self):
        if not self.is_admin:
            return

        kind, entity_id = self.selected_tree_entity()
        if kind == "category":
            category = self.db.category_by_id(entity_id)
            if not category:
                return

            reply = QMessageBox.question(
                self,
                "Подтверждение",
                f"Удалить категорию «{category['name']}» вместе со всеми её задачами и инструкциями?"
            )
            if reply != QMessageBox.Yes:
                return

            try:
                self.db.delete_category(entity_id)
            except Exception as exc:
                QMessageBox.warning(self, "Ошибка", f"Не удалось удалить категорию: {exc}")
                return

        elif kind == "task":
            task = self.db.task_bundle(entity_id)
            if not task:
                return

            reply = QMessageBox.question(
                self,
                "Подтверждение",
                f"Удалить задачу «{task['task_title']}»?"
            )
            if reply != QMessageBox.Yes:
                return

            try:
                self.db.delete_task(entity_id)
            except Exception as exc:
                QMessageBox.warning(self, "Ошибка", f"Не удалось удалить задачу: {exc}")
                return
        else:
            QMessageBox.warning(self, "Внимание", "Сначала выберите категорию или задачу.")
            return

        self.reload_nav_tree()

    def save_tree_changes(self):
        if not self.is_admin:
            return

        self.reload_nav_tree()
        self.statusBar().showMessage("Дерево обновлено")

    # ================== Вкладка инструкции ==================

    def refresh_instruction_tab(self):
        clear_layout(self.instruction_container_layout)
        # Показываем кнопку редактирования, только если есть инструкция и пользователь админ
        if hasattr(self, 'edit_instruction_btn'):
            self.edit_instruction_btn.setVisible(
                self.current_instruction is not None and self.is_admin
            )
        self.current_section_titles = []

        if not self.current_instruction:
            self.instruction_title_label.setText("Инструкция")
            self.sticky_header_label.setVisible(False)
            if self.current_task:
                # Задача есть, но инструкция не привязана
                self.instruction_desc_label.setText("")
                placeholder = QLabel("Инструкция пока не добавлена.")
                placeholder.setWordWrap(True)
                placeholder.setStyleSheet("color: #5b6577; font-style: italic;")
                self.instruction_container_layout.addWidget(placeholder)
            else:
                # Вообще не выбрана задача
                self.instruction_desc_label.setText("Выберите задачу слева, чтобы открыть подробную инструкцию.")
                placeholder = QLabel("Здесь будут главы, блоки, картинки и внутренние ссылки.")
                placeholder.setWordWrap(True)
                self.instruction_container_layout.addWidget(placeholder)
            return

        self.instruction_title_label.setText(self.current_instruction["instruction_title"])
        self.instruction_desc_label.setText(self.current_instruction["short_desc"])

        self.current_section_titles = [s.get("title", "") for s in self.current_instruction["sections"]]

        self._section_cards = []  # сохраняем для sticky-заголовка
        for sec in self.current_instruction["sections"]:
            card = CollapsibleSection(
                sec.get("title", ""),
                sec.get("blocks", []),
                self.handle_link_activated
            )
            self.instruction_container_layout.addWidget(card)
            self._section_cards.append(card)

        if self.current_instruction["related_titles"]:
            related_group = QGroupBox("Связанные инструкции")
            related_layout = QVBoxLayout(related_group)
            related_layout.setSpacing(6)

            for title in self.current_instruction["related_titles"]:
                btn = QPushButton(title)
                btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
                btn.clicked.connect(lambda _, t=title: self.open_instruction_by_title(t))
                related_layout.addWidget(btn)

            self.instruction_container_layout.addWidget(related_group)
        self.instruction_container_layout.addStretch(1)

        # Обновляем список блоков для комментариев
        self.comment_anchor_box.blockSignals(True)
        self.comment_anchor_box.clear()
        self.comment_anchor_box.addItem("Вся инструкция", "")
        for title in self.current_section_titles:
            if title:
                self.comment_anchor_box.addItem(title, title)
        self.comment_anchor_box.setCurrentIndex(0)
        self.comment_anchor_box.blockSignals(False)

    def edit_current_instruction(self):
        """Открывает редактор для изменения текущей инструкции (только админ)."""
        if not self.is_admin or not self.current_instruction:
            return

        categories = self.db.conn.execute(
            "SELECT id, name FROM categories ORDER BY sort_order, name"
        ).fetchall()
        categories = [dict(row) for row in categories]

        # Подготавливаем данные для предзаполнения
        data = {
            "category_id": self.current_instruction["category_id"],
            "title": self.current_instruction["instruction_title"],
            "short_desc": self.current_instruction["short_desc"],
            "sections": self.current_instruction["sections"],
            "related_titles": self.current_instruction["related_titles"]
        }

        dialog = InstructionEditorDialog(
            db=self.db,
            categories=categories,
            instruction_data=data,
            parent=self
        )

        if dialog.exec() != QDialog.Accepted:
            return

        new_data = dialog.get_data()
        try:
            missing = self.db.update_instruction(
                self.current_instruction["instruction_id"],
                new_data["category_id"],
                new_data["title"],
                new_data["short_desc"],
                new_data["sections"],
                new_data["related_titles"]
            )
        except Exception as exc:
            QMessageBox.warning(self, "Ошибка", f"Не удалось обновить инструкцию: {exc}")
            return

        # Перезагружаем данные инструкции в текущем контексте
        self.current_instruction = self.db.instruction_by_id(self.current_instruction["instruction_id"])
        self.current_section_titles = [s.get("title", "") for s in self.current_instruction["sections"]]
        self.refresh_instruction_tab()
        self.refresh_feedback_tab()
        self.refresh_task_tab()  # обновить краткое описание на вкладке задачи

        self.statusBar().showMessage("Инструкция обновлена")
        if missing:
            QMessageBox.information(self, "Внимание",
                                    "Не найдены связанные инструкции:\n" + "\n".join(missing))

    def handle_link_activated(self, link):
        """Все клики по ссылкам – внутренние или через диалог."""
        if link.startswith("instruction://"):
            title = unquote(link.replace("instruction://", "", 1))
            self.open_instruction_by_title(title)
        else:
            self.show_link_dialog(link)

    def open_instruction_by_title(self, title):
        instruction = self.db.instruction_by_title(title)
        if not instruction:
            QMessageBox.warning(self, "Внимание", f"Инструкция не найдена: {title}")
            return

        task = self.db.task_by_instruction_id(instruction["instruction_id"])

        # Снимаем поисковый фильтр, чтобы целевая инструкция точно была видна в дереве
        self.search_timer.stop()
        self.search_edit.blockSignals(True)
        self.search_edit.clear()
        self.search_edit.blockSignals(False)

        if task:
            self.reload_nav_tree()
            self.show_task(task["task_id"])
            self.tabs.setCurrentIndex(1)
        else:
            self.show_instruction_only(instruction["instruction_id"])

    # ================== Справка ==================

    def refresh_help_tab(self):
        """Загружает и отображает содержимое справки."""
        if not hasattr(self, 'help_browser'):
            return
        content = self.db.get_help_content()
        html_content = render_markdown(content)
        self.help_browser.setHtml(html_content)

    def edit_help_content(self):
        """Открывает редактор справки (только админ)."""
        if not self.is_admin:
            return

        current = self.db.get_help_content()
        text, ok = QInputDialog.getMultiLineText(
            self,
            "Редактирование справки",
            "Текст справки (поддерживается Markdown):",
            current
        )
        if not ok:
            return

        text = text.strip()
        if not text:
            QMessageBox.warning(self, "Внимание", "Текст справки не может быть пустым.")
            return

        try:
            self.db.save_help_content(text)
            self.refresh_help_tab()
            self.statusBar().showMessage("Справка обновлена")
        except Exception as exc:
            QMessageBox.warning(self, "Ошибка", f"Не удалось сохранить справку: {exc}")

    def show_link_dialog(self, url: str):
        """Показывает диалог со ссылкой и кнопками 'Скопировать' и 'Закрыть'."""
        # Не показываем диалог для внутренних ссылок
        if url.startswith("instruction://"):
            self.open_instruction_by_title(unquote(url.replace("instruction://", "", 1)))
            return

        from PySide6.QtWidgets import QApplication

        dlg = QDialog(self)
        dlg.setWindowTitle("Ссылка")
        dlg.setMinimumWidth(480)

        layout = QVBoxLayout(dlg)

        hint = QLabel(
            "Если это ссылка на папку или файл на сервере — скопируй и вставь "
            "в строку поиска проводника.\n"
            "Если это ссылка на сайт — вставь в браузер."
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)

        url_edit = QLineEdit(url)
        url_edit.setReadOnly(True)
        layout.addWidget(url_edit)

        btn_layout = QHBoxLayout()
        copy_btn = QPushButton("Скопировать")
        close_btn = QPushButton("Закрыть")

        def copy():
            QApplication.clipboard().setText(url)
            dlg.accept()

        copy_btn.clicked.connect(copy)
        close_btn.clicked.connect(dlg.reject)

        btn_layout.addStretch()
        btn_layout.addWidget(copy_btn)
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)

        dlg.exec()

    # ================== Комментарии и оценка ==================

    def refresh_feedback_tab(self):
        """Обновляет всю вкладку: статистику, список блоков, комментарии."""
        if not hasattr(self, 'feedback_stats_label'):
            return

        if not self.current_instruction:
            self.feedback_stats_label.setText("Выберите инструкцию слева")
            self.submit_feedback_btn.setEnabled(False)
            self.feedback_comments_browser.setHtml("<i>Выберите инструкцию слева.</i>")
            return

        self.submit_feedback_btn.setEnabled(True)

        # Статистика
        avg_rating, cnt = self.db.rating_stats(self.current_instruction["instruction_id"])
        if cnt == 0:
            self.feedback_stats_label.setText("Пока нет оценок")
        else:
            self.feedback_stats_label.setText(
                f"Средняя оценка: {avg_rating:.1f} / 10 | Оценок: {cnt}"
            )

        # Обновляем список блоков для комментариев
        self.comment_anchor_box.blockSignals(True)
        self.comment_anchor_box.clear()
        self.comment_anchor_box.addItem("Вся инструкция", "")
        for title in self.current_section_titles:
            if title:
                self.comment_anchor_box.addItem(title, title)
        self.comment_anchor_box.setCurrentIndex(0)
        self.comment_anchor_box.blockSignals(False)

        # Комментарии
        comments = self.db.comments_for_instruction(self.current_instruction["instruction_id"])
        self.feedback_comments_browser.setHtml(self._render_comments_html(comments))

    def _render_comments_html(self, comments):
        if not comments:
            return "<p><i>Комментариев пока нет.</i></p>"

        blocks = []
        for c in comments:
            anchor = ""
            if c["anchor"]:
                anchor = f" <span style='color:#6b7280'>[блок: {escape_html(c['anchor'])}]</span>"

            blocks.append(
                "<div style='padding:10px 6px;border-bottom:1px solid #e5e7eb;'>"
                f"<div><b>{escape_html(c['author'])}</b>{anchor} "
                f"<span style='color:#6b7280'>{escape_html(c['created_at'])}</span></div>"
                f"<div style='margin-top:5px;line-height:1.45;'>{escape_html(c['text'])}</div>"
                "</div>"
            )

        return "".join(blocks)

    def submit_feedback(self):
        """Ставит оценку и/или добавляет комментарий."""
        if not self.current_instruction:
            QMessageBox.warning(self, "Внимание", "Сначала выберите инструкцию.")
            return

        rating = self.rating_spin.value()
        comment_text = self.comment_text_edit.toPlainText().strip()

        # Если оценка ниже 7 — требуем комментарий
        if rating < 7 and not comment_text:
            QMessageBox.warning(
                self, "Требуется комментарий",
                "Для оценки ниже 7 необходимо оставить комментарий.\n"
                "Опиши, что именно непонятно или требует доработки."
            )
            return

        # Сохраняем оценку
        if rating > 0:
            self.db.add_rating(self.current_instruction["instruction_id"], rating)

        # Сохраняем комментарий
        if comment_text:
            anchor = self.comment_anchor_box.currentData() or ""
            author = self.comment_author_edit.text().strip() or "Пользователь"
            self.db.add_comment(
                self.current_instruction["instruction_id"],
                anchor, author, False, comment_text
            )

        self.comment_text_edit.clear()
        self.rating_spin.setValue(10)
        self.refresh_feedback_tab()
        self.statusBar().showMessage("Оценка и комментарий добавлены")

    def open_feedback_manager(self):
        """Открывает диалог управления оценками и комментариями (только админ)."""
        if not self.is_admin or not self.current_instruction:
            return

        dialog = FeedbackManagerDialog(
            db=self.db,
            instruction_id=self.current_instruction["instruction_id"],
            parent=self
        )
        dialog.exec()
        self.refresh_feedback_tab()

    def export_feedback_to_excel(self):
        if not self.is_admin:
            return

        default_name = f"feedback_export_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Экспорт в Excel",
            default_name,
            "Excel files (*.xlsx)"
        )
        if not file_path:
            return

        try:
            self.db.export_feedback_xlsx(file_path)
        except Exception as exc:
            QMessageBox.warning(self, "Ошибка", f"Не удалось экспортировать данные: {exc}")
            return

        self.statusBar().showMessage(f"Экспорт выполнен: {file_path}")

    def closeEvent(self, event):
        try:
            self._backup_database()
            self.db.close()
        except Exception:
            pass
        super().closeEvent(event)

    def _backup_database(self):
        """Создаёт резервную копию базы данных в папке backups рядом с EXE."""
        backup_dir = DB_PATH.parent / "backups"
        backup_dir.mkdir(exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = backup_dir / f"knowledge_base_{timestamp}.db"
        try:
            import shutil
            # Убедимся, что все изменения записаны
            self.db.conn.commit()
            shutil.copy2(DB_PATH, backup_path)
            # Оставляем только последние 10 копий
            existing = sorted(backup_dir.glob("knowledge_base_*.db"))
            if len(existing) > 10:
                for old in existing[:-10]:
                    old.unlink()
        except Exception as e:
            print(f"Backup error: {e}")


# ================== СЛУЖЕБНОЕ ==================

def clear_layout(layout):
    while layout.count():
        item = layout.takeAt(0)
        widget = item.widget()
        child_layout = item.layout()

        if widget is not None:
            widget.deleteLater()
        elif child_layout is not None:
            clear_layout(child_layout)


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    apply_light_palette(app)
    app.setApplicationName("База знаний по задачам")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

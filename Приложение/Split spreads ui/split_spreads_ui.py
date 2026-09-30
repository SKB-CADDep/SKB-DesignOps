import copy
import os
import sys

import fitz  # PyMuPDF
from pypdf import PdfReader, PdfWriter
from pypdf.generic import RectangleObject
from PySide6.QtCore import QEvent, QPoint, QRect, QRectF, QSettings, Qt, QTimer, Signal
from PySide6.QtGui import QColor, QImage, QKeyEvent, QPainter, QPainterPath, QPen, QPixmap
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFrame,
    QGraphicsDropShadowEffect,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QSizePolicy,
    QSlider,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

STEP_PX = 5

MIN_PREVIEW_ZOOM = 0.10
MAX_PREVIEW_ZOOM = 4.00
DEFAULT_PREVIEW_ZOOM = 1.00
PREVIEW_FIT_MARGIN = 16

APP_VERSION = "1.0.0"
ONBOARDING_ORG = "FirstApp"
ONBOARDING_APP = "SplitSpreadsUI"
ONBOARDING_KEY = "onboarding_completed"


def _clamp(value: float, low: float, high: float) -> float:
    return max(low, min(high, value))


def _norm_rot(deg: int) -> int:
    deg = deg % 360
    return (deg // 90 * 90) % 360


# Режим выхода: линейный порядок половинок или спуск брошюры (два варианта зигзага).
OUTPUT_LEFT_RIGHT = "left_right"
OUTPUT_RIGHT_LEFT = "right_left"
OUTPUT_BROCHURE_ZIGZAG_A = "brochure_zigzag_a"  # …, (N,1), (2,N−1), (N−2,3), … в 1-based
OUTPUT_BROCHURE_ZIGZAG_B = "brochure_zigzag_b"  # …, (1,N), (N−1,2), (3,N−2), …


def _half_caption(spread_no: int, side: str) -> str:
    which = "левая" if side == "L" else "правая"
    return f"{which} половина разворота {spread_no}"


def brochure_sheet_pairs(plate_count: int) -> list[tuple[int, int]]:
    """Номера полос на каждом листе скана в режиме брошюры A (слева | справа)."""
    sheets = []
    sheet_count = plate_count // 2
    for i in range(sheet_count):
        if i % 2 == 0:
            sheets.append((plate_count - i, i + 1))
        else:
            sheets.append((i + 1, plate_count - i))
    return sheets


def build_output_plan(n_spreads: int, mode: str, skips: list[bool] | None = None) -> list[str]:
    """Описание каждой выходной страницы для выбранного режима."""
    if skips is None:
        skips = [False] * n_spreads
    if len(skips) != n_spreads:
        skips = [False] * n_spreads

    if mode in (OUTPUT_LEFT_RIGHT, OUTPUT_RIGHT_LEFT):
        first_is_left = mode == OUTPUT_LEFT_RIGHT
        plan = []
        for i in range(n_spreads):
            spread_no = i + 1
            if skips[i]:
                plan.append(f"разворот {spread_no} целиком (без разреза)")
                continue
            left = _half_caption(spread_no, "L")
            right = _half_caption(spread_no, "R")
            if first_is_left:
                plan.extend([left, right])
            else:
                plan.extend([right, left])
        return plan

    parts: list[tuple[str, int]] = []
    for i in range(n_spreads):
        parts.append(("one", i) if skips[i] else ("two", i))

    leading: list[int] = []
    trailing: list[int] = []
    core: list[int] = []
    seen_spread = False
    for kind, i in parts:
        if kind == "two":
            seen_spread = True
            core.append(i)
        elif not seen_spread:
            leading.append(i)
        else:
            trailing.append(i)

    plan = [f"разворот {i + 1} целиком (без разреза)" for i in leading]
    if core:
        plates = []
        for i in core:
            plates.append((i, "L"))
            plates.append((i, "R"))
        n_plates = len(plates)
        result: list = [None] * n_plates
        sheet_count = n_plates // 2
        for i in range(sheet_count):
            left_page = plates[2 * i]
            right_page = plates[2 * i + 1]
            if i % 2 == 0:
                page_high = n_plates - i
                page_low = i + 1
                result[page_high - 1] = left_page
                result[page_low - 1] = right_page
            else:
                page_low = i + 1
                page_high = n_plates - i
                result[page_low - 1] = left_page
                result[page_high - 1] = right_page
        if mode == OUTPUT_BROCHURE_ZIGZAG_B:
            result.reverse()
        for spread_i, side in result:
            plan.append(_half_caption(spread_i + 1, side))
    plan.extend(f"разворот {i + 1} целиком (без разреза)" for i in trailing)
    return plan


def format_page_numbers(count: int) -> str:
    if count <= 0:
        return "—"
    if count <= 24:
        return ", ".join(str(i) for i in range(1, count + 1))
    head = ", ".join(str(i) for i in range(1, 13))
    tail = ", ".join(str(i) for i in range(count - 5, count + 1))
    return f"{head}, …, {tail}"


class ModeHelpDialog(QDialog):
    def __init__(self, parent, title: str, html: str):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(560, 460)

        layout = QVBoxLayout(self)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        body = QLabel(html)
        body.setWordWrap(True)
        body.setTextFormat(Qt.TextFormat.RichText)
        body.setMargin(8)
        scroll.setWidget(body)
        layout.addWidget(scroll)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)


def split_spreads_per_page(
    input_path: str,
    output_path: str,
    output_mode: str,
    global_rotation: int,
    preview_zoom: float,
    per_page_offset_px: list[int],
    per_page_skip: list[bool],
    per_page_use_rotation: list[bool],
    per_page_rotation: list[int],
):
    """
    global_rotation: поворот, применяемый по умолчанию ко всем страницам
    per_page_use_rotation[idx]=True => для страницы используется per_page_rotation[idx]
    иначе используется global_rotation
    output_mode: OUTPUT_* — линейный порядок или спуск брошюры (два зигзага).
    Для брошюры полосы считаются по порядку разворотов во входном PDF: разворот i даёт
    полосы с индексами 2i и 2i+1 в читательском порядке (0 … 2·n−1), без привязки к номерам
    на полосах (−1, 0, … — это уже порядок страниц в файле).
    """
    if output_mode not in (
        OUTPUT_LEFT_RIGHT,
        OUTPUT_RIGHT_LEFT,
        OUTPUT_BROCHURE_ZIGZAG_A,
        OUTPUT_BROCHURE_ZIGZAG_B,
    ):
        raise ValueError(f"Неизвестный режим выхода: {output_mode!r}")

    first_is_left = output_mode == OUTPUT_LEFT_RIGHT
    if output_mode in (OUTPUT_BROCHURE_ZIGZAG_A, OUTPUT_BROCHURE_ZIGZAG_B):
        # Спуск считается для стандартного разворота «слева первая полоса, справа вторая».
        first_is_left = True

    reader = PdfReader(input_path)
    writer = PdfWriter()

    global_rotation = _norm_rot(global_rotation)

    n = len(reader.pages)
    if not (len(per_page_offset_px) == len(per_page_skip) == len(per_page_use_rotation) == len(per_page_rotation) == n):
        raise ValueError("Настройки по страницам не совпадают с количеством страниц PDF.")

    # PyMuPDF нужен только чтобы открыть файл (можно убрать, оставил для симметрии/возможных расширений)
    mu_doc = fitz.open(input_path)
    split_parts: list = []  # ("one", page) | ("two", out_a, out_b); out_a = первая в чтении, out_b = вторая

    try:
        for idx, page in enumerate(reader.pages):
            # Выбираем фактический поворот для этой страницы
            page_user_rot = _norm_rot(per_page_rotation[idx]) if per_page_use_rotation[idx] else global_rotation

            if per_page_skip[idx]:
                out = copy.copy(page)
                if page_user_rot:
                    out.rotation = _norm_rot((getattr(out, "rotation", 0) or 0) + page_user_rot)
                split_parts.append(("one", out))
                continue

            page_rot = _norm_rot(getattr(page, "rotation", 0) or 0)
            effective_rot = _norm_rot(page_rot + page_user_rot)

            box = page.cropbox if page.cropbox is not None else page.mediabox
            L = float(box.left); R = float(box.right); B = float(box.bottom); T = float(box.top)
            width_pts = R - L
            height_pts = T - B

            # offset_px -> points (быстро, без рендера)
            offset_pts = float(per_page_offset_px[idx]) / float(preview_zoom)

            # Инверсия направления смещения для “зеркальных” ориентаций
            if effective_rot in (180, 270):
                offset_pts = -offset_pts

            p1 = copy.copy(page)
            p2 = copy.copy(page)

            if effective_rot in (0, 180):
                mid = L + width_pts / 2.0 + offset_pts
                mid = max(L + 1.0, min(R - 1.0, mid))

                p1.cropbox = RectangleObject((L, B, mid, T))
                p2.cropbox = RectangleObject((mid, B, R, T))

                if first_is_left:
                    out_a, out_b = p1, p2
                else:
                    out_a, out_b = p2, p1

            else:
                mid = B + height_pts / 2.0 + offset_pts
                mid = max(B + 1.0, min(T - 1.0, mid))

                p1.cropbox = RectangleObject((L, B, R, mid))  # bottom
                p2.cropbox = RectangleObject((L, mid, R, T))  # top

                if effective_rot == 90:
                    visual_left = p1   # bottom
                    visual_right = p2  # top
                else:  # 270
                    visual_left = p2   # top
                    visual_right = p1  # bottom

                if first_is_left:
                    out_a, out_b = visual_left, visual_right
                else:
                    out_a, out_b = visual_right, visual_left

            if page_user_rot:
                out_a.rotation = _norm_rot((getattr(out_a, "rotation", 0) or 0) + page_user_rot)
                out_b.rotation = _norm_rot((getattr(out_b, "rotation", 0) or 0) + page_user_rot)

            split_parts.append(("two", out_a, out_b))
    finally:
        mu_doc.close()

    if output_mode in (OUTPUT_BROCHURE_ZIGZAG_A, OUTPUT_BROCHURE_ZIGZAG_B):
        leading = []
        trailing = []
        core = []
        seen_spread = False
        for part in split_parts:
            if part[0] == "two":
                seen_spread = True
                core.append(part)
            elif not seen_spread:
                leading.append(part[1])
            else:
                trailing.append(part[1])

        if not core:
            raise ValueError("Нет страниц для разрезания: все страницы помечены «Пропустить».")

        # Собираем страницы в порядке скана (только разрезаемые развороты)
        plates = []
        for part in core:
            plates.append(part[1])  # левая половина
            plates.append(part[2])  # правая половина

        N = len(plates)

        if N % 2 != 0:
            raise ValueError("Количество страниц должно быть чётным.")

        # ✅ Восстановление линейного порядка
        result = [None] * N
        sheet_count = N // 2

        for i in range(sheet_count):

            left_page = plates[2 * i]
            right_page = plates[2 * i + 1]

            if i % 2 == 0:
                page_high = N - i
                page_low = i + 1

                result[page_high - 1] = left_page
                result[page_low - 1] = right_page
            else:
                page_low = i + 1
                page_high = N - i

                result[page_low - 1] = left_page
                result[page_high - 1] = right_page

        # ✅ Если выбран режим "назад" — переворачиваем
        if output_mode == OUTPUT_BROCHURE_ZIGZAG_B:
            result.reverse()

        # ✅ Записываем в PDF: пропущенные страницы остаются как есть
        for page in leading:
            writer.add_page(page)
        for page in result:
            writer.add_page(page)
        for page in trailing:
            writer.add_page(page)
    else:
        for part in split_parts:
            if part[0] == "one":
                writer.add_page(part[1])
            else:
                writer.add_page(part[1])
                writer.add_page(part[2])

    with open(output_path, "wb") as f:
        writer.write(f)


def _fitz_pixmap_to_qpixmap(pix) -> QPixmap:
    fmt = QImage.Format.Format_RGBA8888 if pix.alpha else QImage.Format.Format_RGB888
    image = QImage(pix.samples, pix.width, pix.height, pix.stride, fmt)
    return QPixmap.fromImage(image.copy())


class PreviewCanvas(QWidget):
    offset_step = Signal(int)
    page_prev = Signal()
    page_next = Signal()
    rotate_page = Signal(int)
    resized = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFocusPolicy(Qt.StrongFocus)
        self.setMinimumSize(400, 300)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self._pixmap: QPixmap | None = None
        self._skip = False
        self._offset_px = 0
        self._placeholder = "Выберите PDF для предпросмотра."

    def set_placeholder(self, text: str):
        self._pixmap = None
        self._placeholder = text
        self.update()

    def set_preview(self, pixmap: QPixmap | None, skip: bool, offset_px: int):
        self._pixmap = pixmap
        self._skip = skip
        self._offset_px = offset_px
        self.update()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.resized.emit()

    def mousePressEvent(self, event):
        self.setFocus()
        super().mousePressEvent(event)

    def keyPressEvent(self, event: QKeyEvent):
        key = event.key()
        if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            if key == Qt.Key.Key_Left:
                self.rotate_page.emit(-90)
                return
            if key == Qt.Key.Key_Right:
                self.rotate_page.emit(+90)
                return
        if key == Qt.Key.Key_Left:
            self.offset_step.emit(-STEP_PX)
        elif key == Qt.Key.Key_Right:
            self.offset_step.emit(+STEP_PX)
        elif key == Qt.Key.Key_Up:
            self.page_prev.emit()
        elif key == Qt.Key.Key_Down:
            self.page_next.emit()
        else:
            super().keyPressEvent(event)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.fillRect(self.rect(), QColor("#1e1e1e"))
        if self._pixmap is None or self._pixmap.isNull():
            painter.setPen(QColor("white"))
            painter.drawText(10, 24, self._placeholder)
            return

        iw = self._pixmap.width()
        ih = self._pixmap.height()
        x0 = max(0, (self.width() - iw) // 2)
        y0 = max(0, (self.height() - ih) // 2)
        painter.drawPixmap(x0, y0, self._pixmap)

        if self._skip:
            painter.setPen(QColor("yellow"))
            painter.drawText(x0 + 10, y0 + 24, "Эта страница будет пропущена (без разреза).")
            return

        cut_x = x0 + iw / 2.0 + self._offset_px
        cut_x = max(x0 + 1, min(x0 + iw - 1, cut_x))
        painter.setPen(QPen(QColor("red"), 2))
        painter.drawLine(int(cut_x), y0, int(cut_x), y0 + ih)


APP_STYLESHEET = """
    QWidget#centralWidget {
        background-color: #f5f6f8;
    }
    QGroupBox {
        font-weight: 600;
        border: 1px solid #d7dde5;
        border-radius: 8px;
        margin-top: 10px;
        padding: 8px;
        background-color: #ffffff;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        left: 10px;
        padding: 0 4px;
    }
    QLineEdit {
        border: 1px solid #d7dde5;
        border-radius: 6px;
        padding: 6px 10px;
        background-color: #ffffff;
    }
    QPushButton {
        border: 1px solid #d7dde5;
        border-radius: 6px;
        padding: 6px 12px;
        background-color: #ffffff;
    }
    QPushButton:hover {
        background-color: #eef2f7;
    }
    QPushButton#runBtn {
        background-color: #0d6efd;
        color: white;
        font-weight: 600;
        border: none;
        padding: 10px 18px;
        min-height: 36px;
    }
    QPushButton#runBtn:hover {
        background-color: #0b5ed7;
    }
    QPushButton#helpBtn {
        background-color: #0d6efd;
        color: white;
        font-weight: 700;
        font-size: 16px;
        border: none;
        border-radius: 16px;
        padding: 0;
        min-width: 32px;
        max-width: 32px;
        min-height: 32px;
        max-height: 32px;
    }
    QPushButton#helpBtn:hover {
        background-color: #0b5ed7;
    }
    QTableWidget {
        border: 1px solid #d7dde5;
        gridline-color: #e5e7eb;
        background-color: #ffffff;
        selection-background-color: #dbeafe;
        selection-color: #111827;
    }
    QHeaderView::section {
        background-color: #f3f4f6;
        padding: 4px;
        border: none;
        border-right: 1px solid #e5e7eb;
        border-bottom: 1px solid #e5e7eb;
        font-weight: 600;
    }
    PreviewCanvas {
        border: 1px solid #444444;
    }
"""


def _onboarding_settings() -> QSettings:
    return QSettings(ONBOARDING_ORG, ONBOARDING_APP)


def is_onboarding_completed() -> bool:
    return str(_onboarding_settings().value(ONBOARDING_KEY, "") or "").strip() == APP_VERSION


def set_onboarding_completed(completed: bool = True) -> None:
    _onboarding_settings().setValue(ONBOARDING_KEY, APP_VERSION if completed else "")


class OnboardingOverlay(QWidget):
    """Полноэкранная подсветка элементов интерфейса с пошаговыми подсказками."""

    def __init__(self, main_window, steps, on_completed=None, parent=None):
        super().__init__(parent or main_window)
        self._main_window = main_window
        self._steps = steps
        self._on_completed = on_completed
        self._step_index = 0
        self._highlight_rect = QRect()

        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, False)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)

        self._card = QFrame(self)
        self._card.setObjectName("onboardingCard")
        self._card.setStyleSheet("""
            QFrame#onboardingCard {
                background: #ffffff;
                border: 1px solid #c6d8ff;
                border-radius: 10px;
            }
            QFrame#onboardingCard QLabel#onboardingTitle {
                font-size: 14px;
                font-weight: bold;
                color: #1f2937;
            }
            QFrame#onboardingCard QLabel#onboardingText {
                font-size: 12px;
                color: #374151;
            }
            QFrame#onboardingCard QLabel#onboardingCounter {
                font-size: 11px;
                color: #6b7280;
            }
            QFrame#onboardingCard QPushButton#onboardingNext {
                background-color: #0d6efd;
                color: white;
                font-weight: 600;
                border: none;
                border-radius: 6px;
                padding: 6px 14px;
            }
            QFrame#onboardingCard QPushButton#onboardingNext:hover {
                background-color: #0b5ed7;
            }
        """)
        shadow = QGraphicsDropShadowEffect(self._card)
        shadow.setBlurRadius(24)
        shadow.setOffset(0, 4)
        shadow.setColor(QColor(0, 0, 0, 60))
        self._card.setGraphicsEffect(shadow)

        card_layout = QVBoxLayout(self._card)
        card_layout.setContentsMargins(16, 14, 16, 14)
        card_layout.setSpacing(10)

        self._title_label = QLabel()
        self._title_label.setObjectName("onboardingTitle")
        self._title_label.setWordWrap(True)
        card_layout.addWidget(self._title_label)

        self._text_label = QLabel()
        self._text_label.setObjectName("onboardingText")
        self._text_label.setWordWrap(True)
        card_layout.addWidget(self._text_label)

        buttons_row = QHBoxLayout()
        buttons_row.setSpacing(8)

        self._counter_label = QLabel()
        self._counter_label.setObjectName("onboardingCounter")
        buttons_row.addWidget(self._counter_label, 1)

        self._skip_button = QPushButton("Пропустить")
        self._skip_button.setFlat(True)
        self._skip_button.clicked.connect(self._finish)
        buttons_row.addWidget(self._skip_button)

        self._back_button = QPushButton("Назад")
        self._back_button.clicked.connect(self._go_back)
        buttons_row.addWidget(self._back_button)

        self._next_button = QPushButton("Далее")
        self._next_button.setObjectName("onboardingNext")
        self._next_button.setDefault(True)
        self._next_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self._next_button.clicked.connect(self._go_next)
        buttons_row.addWidget(self._next_button)

        card_layout.addLayout(buttons_row)
        self._card.setFixedWidth(360)

    def start(self):
        self._fit_to_parent()
        self.show()
        self.raise_()
        self.setFocus()
        self._show_step(0)

    def refresh_current_step(self):
        if self.isVisible():
            self._apply_step_geometry()

    def _fit_to_parent(self):
        parent = self.parentWidget()
        if parent:
            self.setGeometry(parent.rect())

    def _rect_in_overlay(self, global_rect):
        if global_rect is None or global_rect.isNull():
            return QRect()
        top_left = self.mapFromGlobal(global_rect.topLeft())
        return QRect(top_left, global_rect.size())

    def _widget_highlight_rect(self, widget):
        if widget is None or not widget.isVisible():
            return QRect()
        global_rect = QRect(widget.mapToGlobal(QPoint(0, 0)), widget.size())
        return self._rect_in_overlay(global_rect)

    def _resolve_highlight_rect(self, step):
        if step.get("rect_getter"):
            rect = step["rect_getter"](self._main_window, self)
            if rect is not None and not rect.isNull():
                return rect
        widget = step.get("widget_getter")
        target = widget(self._main_window) if callable(widget) else widget
        if target is None:
            return QRect(self.width() // 4, self.height() // 4, self.width() // 2, self.height() // 2)
        return self._widget_highlight_rect(target)

    def _apply_step_geometry(self):
        if self._step_index < 0 or self._step_index >= len(self._steps):
            return

        step = self._steps[self._step_index]
        self._fit_to_parent()

        padding = step.get("padding", 10)
        rect = self._resolve_highlight_rect(step)
        if rect.isNull():
            rect = QRect(self.width() // 4, self.height() // 4, self.width() // 2, self.height() // 2)
        self._highlight_rect = rect.adjusted(-padding, -padding, padding, padding)

        self._card.adjustSize()
        self._position_card(step.get("placement", "auto"))
        self.update()

    def _show_step(self, index):
        if index < 0 or index >= len(self._steps):
            self._finish()
            return

        self._step_index = index
        step = self._steps[index]

        on_enter = step.get("on_enter")
        if on_enter:
            on_enter(self._main_window)

        self._title_label.setText(step.get("title", ""))
        self._text_label.setText(step.get("text", ""))
        self._counter_label.setText(f"Шаг {index + 1} из {len(self._steps)}")
        self._back_button.setEnabled(index > 0)

        is_last = index >= len(self._steps) - 1
        self._next_button.setText("Готово" if is_last else "Далее")

        QApplication.processEvents()
        delay = step.get("geometry_delay", 50)
        QTimer.singleShot(delay, self._apply_step_geometry)

    def _position_card(self, placement):
        margin = 14
        card_w = self._card.width()
        card_h = self._card.height()
        highlight = self._highlight_rect

        if placement == "right":
            candidates = ["right", "bottom", "left", "top"]
        elif placement == "left":
            candidates = ["left", "bottom", "right", "top"]
        elif placement == "bottom":
            candidates = ["bottom", "right", "left", "top"]
        elif placement == "top":
            candidates = ["top", "right", "bottom", "left"]
        else:
            candidates = ["right", "bottom", "left", "top"]

        def try_place(side):
            if side == "right":
                x = highlight.right() + margin
                y = highlight.center().y() - card_h // 2
            elif side == "left":
                x = highlight.left() - margin - card_w
                y = highlight.center().y() - card_h // 2
            elif side == "bottom":
                x = highlight.center().x() - card_w // 2
                y = highlight.bottom() + margin
            else:
                x = highlight.center().x() - card_w // 2
                y = highlight.top() - margin - card_h

            x = max(margin, min(x, self.width() - card_w - margin))
            y = max(margin, min(y, self.height() - card_h - margin))
            card_rect = QRect(x, y, card_w, card_h)
            if not card_rect.intersects(highlight.adjusted(-8, -8, 8, 8)):
                return card_rect
            return None

        placed = None
        for side in candidates:
            placed = try_place(side)
            if placed:
                break
        if not placed:
            placed = QRect(
                max(margin, (self.width() - card_w) // 2),
                max(margin, self.height() - card_h - margin - 8),
                card_w,
                card_h,
            )
        self._card.setGeometry(placed)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        if self._highlight_rect.isNull():
            painter.fillRect(self.rect(), QColor(0, 0, 0, 165))
            return

        path = QPainterPath()
        path.setFillRule(Qt.FillRule.OddEvenFill)
        path.addRect(QRectF(self.rect()))
        path.addRoundedRect(QRectF(self._highlight_rect), 8, 8)
        painter.fillPath(path, QColor(0, 0, 0, 165))

        pen = QPen(QColor("#3b82f6"), 2)
        painter.setPen(pen)
        painter.setBrush(Qt.BrushStyle.NoBrush)
        painter.drawRoundedRect(self._highlight_rect, 8, 8)

    def mousePressEvent(self, event):
        if self._card.geometry().contains(event.position().toPoint()):
            super().mousePressEvent(event)
            return
        self._go_next()

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter, Qt.Key.Key_Space):
            self._go_next()
            return
        if event.key() == Qt.Key.Key_Escape:
            self._finish()
            return
        if event.key() == Qt.Key.Key_Backspace and self._step_index > 0:
            self._go_back()
            return
        super().keyPressEvent(event)

    def _go_next(self):
        if self._step_index >= len(self._steps) - 1:
            self._finish()
        else:
            self._show_step(self._step_index + 1)

    def _go_back(self):
        if self._step_index > 0:
            self._show_step(self._step_index - 1)

    def _finish(self):
        self.hide()
        callback = self._on_completed
        self._on_completed = None
        self.deleteLater()
        if callback:
            callback()


class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF: разрезать развороты (предпросмотр + поворот + настройки по страницам)")
        self.resize(1200, 740)
        self.setMinimumSize(1200, 740)

        self._doc = None
        self._page_count = 0
        self._preview_pixmap: QPixmap | None = None
        self._page_index = 0
        self._offset = 0
        self._global_rotation = 0
        self._loading_page = False
        self._updating_table = False

        self._offsets: list[int] = []
        self._skips: list[bool] = []
        self._confirmed: list[bool] = []
        self._use_page_rotation: list[bool] = []
        self._page_rotations: list[int] = []  # 0/90/180/270
        self._auto_zoom_enabled = True
        self._last_zoom = DEFAULT_PREVIEW_ZOOM
        self._fit_page_w = 0.0
        self._fit_page_h = 0.0
        self._fit_canvas_w = 0
        self._fit_canvas_h = 0
        self._suspend_fit = False
        self._suspend_mode_info = False
        self._onboarding_overlay = None
        self._fit_timer = QTimer(self)
        self._fit_timer.setSingleShot(True)
        self._fit_timer.setInterval(50)
        self._fit_timer.timeout.connect(self._fit_if_auto)

        self._build_ui()
        self.setStyleSheet(APP_STYLESHEET)
        QTimer.singleShot(450, self._maybe_start_onboarding_on_first_run)

    def _build_ui(self):
        central = QWidget()
        central.setObjectName("centralWidget")
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(10, 8, 10, 10)
        root.setSpacing(8)

        self.input_row = QFrame()
        r1 = QHBoxLayout(self.input_row)
        r1.setContentsMargins(0, 0, 0, 0)
        r1.addWidget(QLabel("Входной PDF:"))
        self.input_edit = QLineEdit()
        r1.addWidget(self.input_edit, stretch=1)
        btn_choose_in = QPushButton("Выбрать…")
        btn_choose_in.clicked.connect(self.choose_input)
        r1.addWidget(btn_choose_in)
        btn_reset = QPushButton("Сбросить выбор")
        btn_reset.clicked.connect(self.reset_selection)
        r1.addWidget(btn_reset)
        self.help_btn = QPushButton("?")
        self.help_btn.setObjectName("helpBtn")
        self.help_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.help_btn.setToolTip("Обучение по программе")
        self.help_btn.clicked.connect(lambda: self.start_onboarding(mark_completed=False))
        r1.addWidget(self.help_btn)
        root.addWidget(self.input_row)

        self.output_row = QFrame()
        r2 = QHBoxLayout(self.output_row)
        r2.setContentsMargins(0, 0, 0, 0)
        r2.addWidget(QLabel("Сохранить как:"))
        self.output_edit = QLineEdit()
        r2.addWidget(self.output_edit, stretch=1)
        btn_choose_out = QPushButton("Сохранить куда")
        btn_choose_out.clicked.connect(self.choose_output)
        r2.addWidget(btn_choose_out)
        root.addWidget(self.output_row)

        self.mode_box = QGroupBox("Режим обработки (только один вариант)")
        mode_layout = QVBoxLayout(self.mode_box)
        r3a = QHBoxLayout()
        r3b = QHBoxLayout()
        self.mode_lr = QRadioButton("Слева → затем справа")
        self.mode_rl = QRadioButton("Справа → затем слева")
        self.mode_bza = QRadioButton("Брошюра: зигзаг по листам (N–1, 2–(N–1), …)")
        self.mode_bzb = QRadioButton("Брошюра: зигзаг по листам (1–N, (N–1)–2, …)")
        self.mode_lr.setChecked(True)
        for btn in (self.mode_lr, self.mode_rl, self.mode_bza):
            r3a.addWidget(btn)
        r3a.addStretch(1)
        r3b.addWidget(self.mode_bzb)
        r3b.addStretch(1)
        mode_layout.addLayout(r3a)
        mode_layout.addLayout(r3b)
        for btn in (self.mode_lr, self.mode_rl, self.mode_bza, self.mode_bzb):
            btn.toggled.connect(self._on_output_mode_changed)
        root.addWidget(self.mode_box)

        self.tools_row = QFrame()
        r4 = QHBoxLayout(self.tools_row)
        r4.setContentsMargins(0, 0, 0, 0)
        btn_off_left = QPushButton("<")
        btn_off_left.setFixedWidth(40)
        btn_off_left.clicked.connect(lambda: self.move_offset(-STEP_PX))
        r4.addWidget(btn_off_left)
        btn_off_right = QPushButton(">")
        btn_off_right.setFixedWidth(40)
        btn_off_right.clicked.connect(lambda: self.move_offset(+STEP_PX))
        r4.addWidget(btn_off_right)
        self.offset_label = QLabel("Смещение: 0 px (0 = центр)")
        r4.addWidget(self.offset_label)

        self.skip_cb = QCheckBox("Пропустить текущую страницу")
        self.skip_cb.toggled.connect(self._on_skip_toggle)
        r4.addWidget(self.skip_cb)

        r4.addWidget(QLabel("Поворот всех страниц:"))
        btn_rot_ccw = QPushButton("⟲ -90")
        btn_rot_ccw.clicked.connect(lambda: self.rotate_all_pages(-90))
        r4.addWidget(btn_rot_ccw)
        btn_rot_cw = QPushButton("⟳ +90")
        btn_rot_cw.clicked.connect(lambda: self.rotate_all_pages(+90))
        r4.addWidget(btn_rot_cw)
        self.rot_label = QLabel("0°")
        r4.addWidget(self.rot_label)
        r4.addStretch(1)
        root.addWidget(self.tools_row)

        self.zoom_row = QFrame()
        zoom_row = QHBoxLayout(self.zoom_row)
        zoom_row.setContentsMargins(0, 0, 0, 0)
        zoom_row.addWidget(QLabel("Масштаб предпросмотра:"))
        self.zoom_slider = QSlider(Qt.Orientation.Horizontal)
        self.zoom_slider.setRange(int(MIN_PREVIEW_ZOOM * 100), int(MAX_PREVIEW_ZOOM * 100))
        self.zoom_slider.setSingleStep(5)
        self.zoom_slider.setPageStep(5)
        self.zoom_slider.setTickInterval(5)
        self.zoom_slider.setTickPosition(QSlider.TickPosition.TicksBelow)
        self.zoom_slider.setValue(int(DEFAULT_PREVIEW_ZOOM * 100))
        self.zoom_slider.setFixedWidth(280)
        self.zoom_slider.valueChanged.connect(self._on_preview_zoom_change)
        zoom_row.addWidget(self.zoom_slider)
        btn_auto_zoom = QPushButton("Авто")
        btn_auto_zoom.clicked.connect(self._apply_auto_preview_zoom)
        zoom_row.addWidget(btn_auto_zoom)
        self.zoom_value_label = QLabel(f"{DEFAULT_PREVIEW_ZOOM:.2f}x")
        zoom_row.addWidget(self.zoom_value_label)
        zoom_row.addStretch(1)
        root.addWidget(self.zoom_row)

        mid = QHBoxLayout()
        left_panel = QFrame()
        left_panel.setFixedWidth(460)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)

        self.nav_box = QGroupBox("Предпросмотр (клавиатура)")
        nav_layout = QVBoxLayout(self.nav_box)
        nav_row = QHBoxLayout()
        btn_prev = QPushButton("←")
        btn_prev.setFixedWidth(40)
        btn_prev.clicked.connect(self.prev_page)
        nav_row.addWidget(btn_prev)
        self.page_label = QLabel("Стр.: - / -")
        self.page_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        nav_row.addWidget(self.page_label, stretch=1)
        btn_next = QPushButton("→")
        btn_next.setFixedWidth(40)
        btn_next.clicked.connect(self.next_page)
        nav_row.addWidget(btn_next)
        nav_row.addStretch(1)
        nav_layout.addLayout(nav_row)
        hint = QLabel(
            "Клик по предпросмотру или по настройкам страниц даёт фокус.\n"
            "←/→: смещение линии разреза, ↑/↓: смена страницы.\n"
            "Ctrl+←/→: поворот текущей страницы.\n"
            "Смена страницы подтверждает настройки."
        )
        hint.setWordWrap(True)
        nav_layout.addWidget(hint)
        left_layout.addWidget(self.nav_box)

        self.table_box = QGroupBox("Настройки по страницам")
        table_layout = QVBoxLayout(self.table_box)
        page_rot_row = QHBoxLayout()
        page_rot_row.addWidget(QLabel("Поворот страницы:"))
        btn_page_rot_ccw = QPushButton("⟲ -90")
        btn_page_rot_ccw.clicked.connect(lambda: self.rotate_current_page(-90))
        page_rot_row.addWidget(btn_page_rot_ccw)
        btn_page_rot_cw = QPushButton("⟳ +90")
        btn_page_rot_cw.clicked.connect(lambda: self.rotate_current_page(+90))
        page_rot_row.addWidget(btn_page_rot_cw)
        self.page_rot_label = QLabel("0°")
        page_rot_row.addWidget(self.page_rot_label)
        self.use_page_rotation_cb = QCheckBox("Свой поворот")
        self.use_page_rotation_cb.setToolTip(
            "Если выключено, страница использует «Поворот всех страниц».\n"
            "Двойной щелчок по столбцу «Поворот» тоже поворачивает эту страницу."
        )
        self.use_page_rotation_cb.toggled.connect(self._on_use_page_rotation_toggle)
        page_rot_row.addWidget(self.use_page_rotation_cb)
        page_rot_row.addStretch(1)
        table_layout.addLayout(page_rot_row)
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["Стр.", "Смещение (px)", "Пропустить", "Поворот", "Подтв."])
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        self.table.itemSelectionChanged.connect(self.on_table_select)
        self.table.cellDoubleClicked.connect(self.on_table_double_click)
        table_layout.addWidget(self.table)
        self.table_box.setFocusPolicy(Qt.FocusPolicy.ClickFocus)
        self.table_box.setFocusProxy(self.table)
        left_layout.addWidget(self.table_box, stretch=1)
        mid.addWidget(left_panel)
        self._install_preview_hotkeys(left_panel)

        self.canvas = PreviewCanvas()
        self.canvas.offset_step.connect(self.move_offset)
        self.canvas.page_prev.connect(self.prev_page)
        self.canvas.page_next.connect(self.next_page)
        self.canvas.rotate_page.connect(self.rotate_current_page)
        self.canvas.resized.connect(self._on_canvas_resized)
        mid.addWidget(self.canvas, stretch=1)
        root.addLayout(mid, stretch=1)

        bottom = QHBoxLayout()
        self.run_btn = QPushButton("Обработать и сохранить")
        self.run_btn.setObjectName("runBtn")
        self.run_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.run_btn.clicked.connect(self.run)
        bottom.addWidget(self.run_btn)
        bottom.addStretch(1)
        root.addLayout(bottom)

        self._update_labels()

    def _install_preview_hotkeys(self, root: QWidget):
        root.installEventFilter(self)
        for child in root.findChildren(QWidget):
            if not isinstance(child, QLineEdit):
                child.installEventFilter(self)

    def _handle_preview_hotkeys(self, event: QKeyEvent) -> bool:
        focus = QApplication.focusWidget()
        if isinstance(focus, (QLineEdit, QSlider)):
            return False
        key = event.key()
        mods = event.modifiers()
        if mods & Qt.KeyboardModifier.ControlModifier:
            if key == Qt.Key.Key_Left:
                self.rotate_current_page(-90)
                return True
            if key == Qt.Key.Key_Right:
                self.rotate_current_page(+90)
                return True
            return False
        if mods & Qt.KeyboardModifier.AltModifier or mods & Qt.KeyboardModifier.ShiftModifier:
            return False
        if key == Qt.Key.Key_Left:
            self.move_offset(-STEP_PX)
            return True
        if key == Qt.Key.Key_Right:
            self.move_offset(+STEP_PX)
            return True
        if key == Qt.Key.Key_Up:
            self.prev_page()
            return True
        if key == Qt.Key.Key_Down:
            self.next_page()
            return True
        return False

    def eventFilter(self, watched, event):
        if event.type() == QEvent.Type.KeyPress and self._handle_preview_hotkeys(event):
            return True
        return super().eventFilter(watched, event)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        overlay = getattr(self, "_onboarding_overlay", None)
        if overlay is not None and overlay.isVisible():
            overlay._fit_to_parent()
            overlay.refresh_current_step()

    def _is_onboarding_active(self) -> bool:
        overlay = getattr(self, "_onboarding_overlay", None)
        return overlay is not None and overlay.isVisible()

    def _maybe_start_onboarding_on_first_run(self):
        if not is_onboarding_completed():
            self.start_onboarding(mark_completed=True)

    def start_onboarding(self, mark_completed=False):
        overlay = getattr(self, "_onboarding_overlay", None)
        if overlay is not None:
            overlay.close()
            overlay.deleteLater()
            self._onboarding_overlay = None

        def on_completed():
            if mark_completed:
                set_onboarding_completed(True)
            self._onboarding_overlay = None

        self._onboarding_overlay = OnboardingOverlay(
            self,
            self._build_onboarding_steps(),
            on_completed=on_completed,
            parent=self.centralWidget(),
        )
        self._onboarding_overlay.start()

    def _build_onboarding_steps(self):
        return [
            {
                "title": "Входной PDF",
                "text": (
                    "Сначала выберите PDF с разворотами книги. Кнопка «Выбрать…» открывает файл, "
                    "«Сбросить выбор» очищает текущий документ."
                ),
                "widget_getter": lambda w: w.input_row,
                "placement": "bottom",
            },
            {
                "title": "Куда сохранить",
                "text": (
                    "Путь результата подставляется автоматически с суффиксом _split. "
                    "При необходимости укажите другое имя через «Сохранить куда»."
                ),
                "widget_getter": lambda w: w.output_row,
                "placement": "bottom",
            },
            {
                "title": "Режим обработки",
                "text": (
                    "Выберите, как складывать половинки: слева направо, справа налево "
                    "или зигзаг брошюры. При выборе режима появится окно с порядком страниц "
                    "на выходе — например 1, 2, 3, 4, 5…"
                ),
                "widget_getter": lambda w: w.mode_box,
                "placement": "bottom",
            },
            {
                "title": "Линия разреза и поворот",
                "text": (
                    "Стрелки «<» и «>» двигают красную линию разреза. "
                    "Можно пропустить разворот без разреза. "
                    "«Поворот всех страниц» поворачивает весь документ."
                ),
                "widget_getter": lambda w: w.tools_row,
                "placement": "bottom",
            },
            {
                "title": "Масштаб предпросмотра",
                "text": (
                    "Кнопка «Авто» один раз подгоняет разворот под окно. "
                    "Дальше масштаб меняется сам только если вы измените размер окна. "
                    "Ползунок можно двигать вручную."
                ),
                "widget_getter": lambda w: w.zoom_row,
                "placement": "bottom",
            },
            {
                "title": "Страницы и клавиатура",
                "text": (
                    "Листайте развороты стрелками. После клика по предпросмотру или "
                    "по таблице настроек: ←/→ двигают линию разреза, ↑/↓ меняют страницу, "
                    "Ctrl+←/→ поворачивают текущую страницу."
                ),
                "widget_getter": lambda w: w.nav_box,
                "placement": "right",
            },
            {
                "title": "Настройки по страницам",
                "text": (
                    "В таблице видно смещение, пропуск и поворот каждой страницы. "
                    "Здесь же можно повернуть выбранный разворот. "
                    "Двойной щелчок по «Пропустить» или «Поворот» меняет значение."
                ),
                "widget_getter": lambda w: w.table_box,
                "placement": "right",
            },
            {
                "title": "Предпросмотр",
                "text": (
                    "Справа виден текущий разворот и красная линия разреза. "
                    "Вся страница должна помещаться в окно после «Авто»."
                ),
                "widget_getter": lambda w: w.canvas,
                "placement": "left",
            },
            {
                "title": "Обработать и сохранить",
                "text": (
                    "Когда смещения и режим выбраны, нажмите эту кнопку — программа "
                    "разрежет развороты и запишет новый PDF."
                ),
                "widget_getter": lambda w: w.run_btn,
                "placement": "top",
            },
            {
                "title": "Повтор обучения",
                "text": (
                    "Это обучение показывается при первом запуске. "
                    "В любой момент его можно запустить снова кнопкой «?» в правом верхнем углу."
                ),
                "widget_getter": lambda w: w.help_btn,
                "placement": "bottom",
            },
        ]

    def _preview_zoom(self) -> float:
        return self.zoom_slider.value() / 100.0

    def _on_output_mode_changed(self, checked: bool):
        if not checked:
            return
        self.refresh_preview()
        if not self._suspend_mode_info and not self._is_onboarding_active():
            self._show_mode_help()

    def _show_mode_help(self):
        mode = self._get_output_mode()
        title, html = self._mode_help_content(mode)
        ModeHelpDialog(self, title, html).exec()

    def _mode_help_content(self, mode: str) -> tuple[str, str]:
        example_n = 4
        example_pages = example_n * 2
        sheets = brochure_sheet_pairs(example_pages)
        sheet_lines = "".join(
            f"<li>Разворот {i + 1}: <b>{left}</b> | <b>{right}</b></li>"
            for i, (left, right) in enumerate(sheets)
        )
        a_order = format_page_numbers(example_pages)
        b_order = ", ".join(str(i) for i in range(example_pages, 0, -1))

        if mode == OUTPUT_LEFT_RIGHT:
            title = "Режим: слева → затем справа"
            intro = (
                "<p>Каждый разворот режется пополам. В файл сначала попадает "
                "<b>левая</b> половина, затем <b>правая</b>.</p>"
                "<p>Пример для 3 разворотов [1|2], [3|4], [5|6]:</p>"
                f"<p style='font-size:16px'><b>1, 2, 3, 4, 5, 6</b></p>"
            )
        elif mode == OUTPUT_RIGHT_LEFT:
            title = "Режим: справа → затем слева"
            intro = (
                "<p>Каждый разворот режется пополам. В файл сначала попадает "
                "<b>правая</b> половина, затем <b>левая</b>.</p>"
                "<p>Пример для 3 разворотов [1|2], [3|4], [5|6]:</p>"
                f"<p style='font-size:16px'><b>2, 1, 4, 3, 6, 5</b></p>"
            )
        elif mode == OUTPUT_BROCHURE_ZIGZAG_A:
            title = "Режим: брошюра, зигзаг (N–1, 2–(N–1), …)"
            intro = (
                "<p>Скан считается спуском брошюры: на первом развороте слева последняя "
                "полоса, справа первая. Программа собирает книгу <b>с начала к концу</b>.</p>"
                f"<p>Пример на {example_n} разворотах ({example_pages} полос):</p>"
                f"<ul>{sheet_lines}</ul>"
                "<p>Порядок страниц в результате:</p>"
                f"<p style='font-size:16px'><b>{a_order}</b></p>"
            )
        else:
            title = "Режим: брошюра, зигзаг (1–N, (N–1)–2, …)"
            intro = (
                "<p>Тот же разбор спуска брошюры, но готовый файл идёт "
                "<b>с конца к началу</b>.</p>"
                f"<p>Пример на {example_n} разворотах ({example_pages} полос) после обработки:</p>"
                f"<p style='font-size:16px'><b>{b_order}</b></p>"
            )

        file_html = self._current_file_plan_html(mode)
        html = (
            "<div style='font-size:13px; color:#1f2937'>"
            f"{intro}"
            "<p>Пропущенные развороты не режутся и остаются целой страницей.</p>"
            f"{file_html}"
            "</div>"
        )
        return title, html

    def _current_file_plan_html(self, mode: str) -> str:
        if self._page_count <= 0:
            return (
                "<p style='color:#4b5563'>Откройте PDF — здесь появится порядок страниц "
                "для вашего файла.</p>"
            )
        plan = build_output_plan(self._page_count, mode, list(self._skips))
        total = len(plan)
        if total == 0:
            return "<p>Нет страниц для обработки.</p>"

        number_line = format_page_numbers(total)
        items = []
        if total <= 40:
            for idx, caption in enumerate(plan, start=1):
                items.append(f"<li><b>{idx}.</b> {caption}</li>")
        else:
            for idx, caption in enumerate(plan[:16], start=1):
                items.append(f"<li><b>{idx}.</b> {caption}</li>")
            hidden = total - 26
            items.append(f"<li>… ещё {hidden} стр. …</li>")
            for idx, caption in enumerate(plan[-10:], start=total - 9):
                items.append(f"<li><b>{idx}.</b> {caption}</li>")

        return (
            "<hr>"
            f"<p><b>Ваш файл:</b> {self._page_count} разворот(ов) → "
            f"<b>{total}</b> стр. на выходе.</p>"
            f"<p>Номера страниц результата: <b>{number_line}</b></p>"
            "<p>Что окажется на каждой странице:</p>"
            f"<ul style='margin-top:4px'>{''.join(items)}</ul>"
        )

    def _get_output_mode(self) -> str:
        if self.mode_bza.isChecked():
            return OUTPUT_BROCHURE_ZIGZAG_A
        if self.mode_bzb.isChecked():
            return OUTPUT_BROCHURE_ZIGZAG_B
        if self.mode_rl.isChecked():
            return OUTPUT_RIGHT_LEFT
        return OUTPUT_LEFT_RIGHT

    def _effective_user_rotation_for_page(self, i: int) -> int:
        """Какой user_rotation применяем к странице i: глобальный или индивидуальный"""
        if self._use_page_rotation[i]:
            return _norm_rot(self._page_rotations[i])
        return _norm_rot(self._global_rotation)

    def _page_preview_size_pts(self, idx: int) -> tuple[float, float]:
        page = self._doc.load_page(idx)
        rot = self._effective_user_rotation_for_page(idx)
        width = float(page.rect.width)
        height = float(page.rect.height)
        if rot in (90, 270):
            width, height = height, width
        return width, height

    def _calc_fit_zoom(self) -> float:
        canvas_w = self.canvas.width()
        canvas_h = self.canvas.height()
        page_w = self._fit_page_w
        page_h = self._fit_page_h
        if canvas_w < 50 or canvas_h < 50 or page_w <= 1 or page_h <= 1:
            return self._preview_zoom()
        avail_w = max(1.0, canvas_w - PREVIEW_FIT_MARGIN * 2)
        avail_h = max(1.0, canvas_h - PREVIEW_FIT_MARGIN * 2)
        zoom = min(avail_w / page_w, avail_h / page_h)
        return round(_clamp(zoom, MIN_PREVIEW_ZOOM, MAX_PREVIEW_ZOOM), 2)

    def _rescale_offsets_for_zoom(self, old_zoom: float, new_zoom: float):
        if old_zoom <= 1e-6 or not self._offsets:
            return
        if not any(self._offsets) and not self._offset:
            return
        factor = new_zoom / old_zoom
        if abs(factor - 1.0) < 1e-6:
            return
        self._offsets = [int(round(offset * factor)) for offset in self._offsets]
        self._offset = int(round(self._offset * factor))
        for i in range(len(self._offsets)):
            self._update_table_row(i)

    def _set_preview_zoom_value(self, zoom: float, *, force_render: bool = False):
        new_zoom = round(_clamp(float(zoom), MIN_PREVIEW_ZOOM, MAX_PREVIEW_ZOOM), 2)
        zoom_changed = abs(new_zoom - self._last_zoom) >= 1e-9
        if zoom_changed:
            self._rescale_offsets_for_zoom(self._last_zoom, new_zoom)
            self._last_zoom = new_zoom
            self.zoom_slider.blockSignals(True)
            self.zoom_slider.setValue(int(round(new_zoom * 100)))
            self.zoom_slider.blockSignals(False)
        if zoom_changed or force_render:
            self._update_labels()
            self.refresh_preview()

    def _apply_auto_preview_zoom(self):
        self._auto_zoom_enabled = True
        if self._doc is not None and self._page_count > 0:
            self._fit_page_w, self._fit_page_h = self._page_preview_size_pts(self._page_index)
        self._set_preview_zoom_value(self._calc_fit_zoom(), force_render=True)
        self._fit_canvas_w = self.canvas.width()
        self._fit_canvas_h = self.canvas.height()

    def _fit_if_auto(self):
        if self._suspend_fit or not self._auto_zoom_enabled or self._doc is None:
            return
        canvas_w = self.canvas.width()
        canvas_h = self.canvas.height()
        if canvas_w < 50 or canvas_h < 50:
            return
        if canvas_w == self._fit_canvas_w and canvas_h == self._fit_canvas_h:
            return
        self._set_preview_zoom_value(self._calc_fit_zoom())
        self._fit_canvas_w = canvas_w
        self._fit_canvas_h = canvas_h

    def _on_preview_zoom_change(self, value: int):
        self._auto_zoom_enabled = False
        new_zoom = value / 100.0
        self._rescale_offsets_for_zoom(self._last_zoom, new_zoom)
        self._last_zoom = new_zoom
        self._update_labels()
        self.refresh_preview()

    def _on_canvas_resized(self):
        if self._suspend_fit or not self._auto_zoom_enabled or self._doc is None:
            return
        self._fit_timer.start()

    def reset_selection(self):
        try:
            if self._doc is not None:
                self._doc.close()
        except Exception:
            pass

        self._doc = None
        self._page_count = 0
        self._preview_pixmap = None

        self.input_edit.clear()
        self.output_edit.clear()
        self._page_index = 0
        self._suspend_mode_info = True
        try:
            self.mode_lr.setChecked(True)
        finally:
            self._suspend_mode_info = False

        self._offset = 0
        self._global_rotation = 0
        self._loading_page = True
        try:
            self.skip_cb.setChecked(False)
            self.use_page_rotation_cb.setChecked(False)
        finally:
            self._loading_page = False

        self._offsets = []
        self._skips = []
        self._confirmed = []
        self._use_page_rotation = []
        self._page_rotations = []
        self._fit_page_w = 0.0
        self._fit_page_h = 0.0
        self._fit_canvas_w = 0
        self._fit_canvas_h = 0
        self._fit_timer.stop()

        self._updating_table = True
        try:
            self.table.setRowCount(0)
        finally:
            self._updating_table = False

        self._update_labels()
        self.canvas.set_placeholder("Выбор сброшен. Выберите PDF заново.")

    def _make_table_item(self, text: str) -> QTableWidgetItem:
        item = QTableWidgetItem(text)
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        return item

    def _init_per_page_settings(self, n_pages: int):
        self._offsets = [0] * n_pages
        self._skips = [False] * n_pages
        self._confirmed = [False] * n_pages
        self._use_page_rotation = [False] * n_pages
        self._page_rotations = [0] * n_pages

        self._updating_table = True
        try:
            self.table.setRowCount(n_pages)
            for i in range(n_pages):
                self._fill_table_row(i)
            if n_pages > 0:
                self.table.selectRow(0)
                self.table.scrollToItem(self.table.item(0, 0))
        finally:
            self._updating_table = False

    def _fill_table_row(self, i: int):
        rot = self._effective_user_rotation_for_page(i)
        values = (
            str(i + 1),
            str(self._offsets[i]),
            "Да" if self._skips[i] else "Нет",
            f"{rot}°",
            "✓" if self._confirmed[i] else "",
        )
        for col, text in enumerate(values):
            self.table.setItem(i, col, self._make_table_item(text))

    def _update_table_row(self, i: int):
        if i < 0 or i >= self.table.rowCount():
            return
        self._fill_table_row(i)

    def _load_settings_for_page(self, i: int):
        self._offset = int(self._offsets[i])
        self._loading_page = True
        try:
            self.skip_cb.setChecked(bool(self._skips[i]))
            self.use_page_rotation_cb.setChecked(bool(self._use_page_rotation[i]))
        finally:
            self._loading_page = False

    def _inherit_offset_from_previous(self, new_idx: int):
        if new_idx <= 0 or new_idx >= self._page_count:
            return
        if self._confirmed[new_idx]:
            return
        self._offsets[new_idx] = self._offsets[new_idx - 1]
        self._update_table_row(new_idx)

    def _select_table_row(self, i: int):
        self._updating_table = True
        try:
            self.table.selectRow(i)
            item = self.table.item(i, 0)
            if item is not None:
                self.table.scrollToItem(item)
        finally:
            self._updating_table = False

    def _update_labels(self):
        self.offset_label.setText(f"Смещение: {self._offset} px (0 = центр)")
        self.rot_label.setText(f"{_norm_rot(self._global_rotation)}°")

        i = self._page_index
        if 0 <= i < self._page_count:
            rot = self._effective_user_rotation_for_page(i)
            self.page_rot_label.setText(f"{rot}°")
        else:
            self.page_rot_label.setText("0°")

        if self._page_count > 0:
            self.page_label.setText(f"Стр.: {self._page_index + 1} / {self._page_count}")
        else:
            self.page_label.setText("Стр.: - / -")
        self.zoom_value_label.setText(
            f"{self._preview_zoom():.2f}x" + (" (авто)" if self._auto_zoom_enabled else "")
        )

    def choose_input(self):
        path, _ = QFileDialog.getOpenFileName(self, "Выберите PDF", "", "PDF files (*.pdf)")
        if not path:
            return

        self.input_edit.setText(path)
        base, ext = os.path.splitext(path)
        self.output_edit.setText(base + "_split" + ext)

        try:
            self._suspend_fit = True
            self._fit_timer.stop()
            if self._doc is not None:
                self._doc.close()
            self._doc = fitz.open(path)
            self._page_count = self._doc.page_count
            self._page_index = 0

            self._init_per_page_settings(self._page_count)
            self._load_settings_for_page(0)
            self._update_table_row(0)
            self._apply_auto_preview_zoom()
            self.canvas.setFocus()
        except Exception as e:
            self.reset_selection()
            QMessageBox.critical(self, "Ошибка", f"Не удалось открыть PDF:\n{e}")
        finally:
            self._suspend_fit = False

    def choose_output(self):
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить результат как",
            self.output_edit.text(),
            "PDF files (*.pdf)",
        )
        if path:
            self.output_edit.setText(path)

    def prev_page(self):
        if self._page_count <= 0:
            return

        idx = self._page_index
        if idx > 0:
            idx -= 1
            self._page_index = idx
            self._select_table_row(idx)
            self._load_settings_for_page(idx)
            self._update_labels()
            self.refresh_preview()

    def next_page(self):
        if self._page_count <= 0:
            return

        idx = self._page_index
        if idx < self._page_count - 1:
            idx += 1
            self._inherit_offset_from_previous(idx)

            self._page_index = idx
            self._select_table_row(idx)
            self._load_settings_for_page(idx)
            self._update_labels()
            self.refresh_preview()

    def on_table_select(self):
        if self._updating_table:
            return
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            return
        new_idx = rows[0].row()

        self._inherit_offset_from_previous(new_idx)
        self._page_index = new_idx
        self._load_settings_for_page(new_idx)
        self._update_labels()
        self.refresh_preview()

    def on_table_double_click(self, row: int, column: int):
        if row < 0 or row >= self._page_count:
            return
        if column == 3:
            self._rotate_page(row, +90)
            return
        self._skips[row] = not self._skips[row]
        self._confirmed[row] = True
        if row == self._page_index:
            self._loading_page = True
            try:
                self.skip_cb.setChecked(self._skips[row])
            finally:
                self._loading_page = False
        self._update_table_row(row)
        self.refresh_preview()

    def _on_skip_toggle(self):
        if self._loading_page or self._page_count <= 0:
            return
        i = self._page_index
        self._skips[i] = bool(self.skip_cb.isChecked())
        self._confirmed[i] = True
        self._update_table_row(i)
        self.refresh_preview()

    def _on_use_page_rotation_toggle(self):
        if self._loading_page or self._page_count <= 0:
            return
        i = self._page_index
        enabled = bool(self.use_page_rotation_cb.isChecked())
        if enabled and not self._use_page_rotation[i]:
            self._page_rotations[i] = _norm_rot(self._global_rotation)
        self._use_page_rotation[i] = enabled
        self._confirmed[i] = True
        self._update_table_row(i)
        self._update_labels()
        self.refresh_preview()

    def move_offset(self, delta):
        if self._page_count <= 0:
            return
        i = self._page_index
        self._offset = int(self._offset + delta)

        self._offsets[i] = int(self._offset)
        self._confirmed[i] = True
        self._update_table_row(i)

        self._update_labels()
        self.refresh_preview()

    def rotate_all_pages(self, delta):
        self._global_rotation = _norm_rot(self._global_rotation + delta)
        for i in range(self._page_count):
            if not self._use_page_rotation[i]:
                self._update_table_row(i)
        self._update_labels()
        self.refresh_preview()

    def rotate_current_page(self, delta):
        if self._page_count <= 0:
            return
        self._rotate_page(self._page_index, delta)

    def _rotate_page(self, i: int, delta: int):
        if i < 0 or i >= self._page_count:
            return

        if not self._use_page_rotation[i]:
            self._page_rotations[i] = _norm_rot(self._global_rotation)
        self._use_page_rotation[i] = True
        self._page_rotations[i] = _norm_rot(self._page_rotations[i] + delta)
        self._confirmed[i] = True

        if i == self._page_index:
            self._loading_page = True
            try:
                self.use_page_rotation_cb.setChecked(True)
            finally:
                self._loading_page = False

        self._update_table_row(i)
        self._update_labels()
        self.refresh_preview()

    def refresh_preview(self):
        self._update_labels()

        if self._doc is None or self._page_count == 0:
            self.canvas.set_placeholder("Выберите PDF для предпросмотра.")
            return

        idx = self._page_index
        page = self._doc.load_page(idx)

        rot_for_page = self._effective_user_rotation_for_page(idx)
        preview_zoom = self._preview_zoom()
        mat = fitz.Matrix(preview_zoom, preview_zoom).prerotate(rot_for_page)
        pix = page.get_pixmap(matrix=mat, alpha=False)

        self._preview_pixmap = _fitz_pixmap_to_qpixmap(pix)
        self.canvas.set_preview(self._preview_pixmap, bool(self._skips[idx]), self._offset)

    def run(self):
        in_path = self.input_edit.text().strip()
        out_path = self.output_edit.text().strip()

        if not in_path or not os.path.isfile(in_path):
            QMessageBox.critical(self, "Ошибка", "Выберите существующий входной PDF.")
            return
        if not out_path:
            QMessageBox.critical(self, "Ошибка", "Выберите путь для сохранения результата.")
            return
        if self._page_count <= 0:
            QMessageBox.critical(self, "Ошибка", "Нет открытого документа.")
            return

        mode = self._get_output_mode()
        if all(self._skips):
            QMessageBox.critical(
                self,
                "Ошибка",
                "Нет страниц для разрезания: все страницы помечены «Пропустить».",
            )
            return

        try:
            split_spreads_per_page(
                input_path=in_path,
                output_path=out_path,
                output_mode=mode,
                global_rotation=int(self._global_rotation),
                preview_zoom=self._preview_zoom(),
                per_page_offset_px=list(self._offsets),
                per_page_skip=list(self._skips),
                per_page_use_rotation=list(self._use_page_rotation),
                per_page_rotation=list(self._page_rotations),
            )
        except Exception as e:
            QMessageBox.critical(self, "Ошибка обработки", str(e))
            return

        QMessageBox.information(self, "Готово", f"Файл сохранён:\n{out_path}")

    def closeEvent(self, event):
        overlay = getattr(self, "_onboarding_overlay", None)
        if overlay is not None:
            overlay.close()
            self._onboarding_overlay = None
        try:
            if self._doc is not None:
                self._doc.close()
        except Exception:
            pass
        super().closeEvent(event)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = App()
    window.show()
    sys.exit(app.exec())

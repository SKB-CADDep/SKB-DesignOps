import os
import copy
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from pypdf import PdfReader, PdfWriter
from pypdf.generic import RectangleObject

import fitz  # PyMuPDF

STEP_PX = 5
PREVIEW_ZOOM = 1.2


def _norm_rot(deg: int) -> int:
    deg = deg % 360
    return (deg // 90 * 90) % 360


# Режим выхода: линейный порядок половинок или спуск брошюры (два варианта зигзага).
OUTPUT_LEFT_RIGHT = "left_right"
OUTPUT_RIGHT_LEFT = "right_left"
OUTPUT_BROCHURE_ZIGZAG_A = "brochure_zigzag_a"  # …, (N,1), (2,N−1), (N−2,3), … в 1-based
OUTPUT_BROCHURE_ZIGZAG_B = "brochure_zigzag_b"  # …, (1,N), (N−1,2), (3,N−2), …


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

        if any(p[0] == "one" for p in split_parts):
            raise ValueError("Режим восстановления книги несовместим с пропуском страниц.")

        # Собираем страницы в порядке скана
        plates = []
        for part in split_parts:
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

        # ✅ Записываем в PDF
        for page in result:
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


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF: разрезать развороты (предпросмотр + поворот + настройки по страницам)")
        self.geometry("1200x740")
        self.minsize(1200, 740)

        self.input_path_var = tk.StringVar()
        self.output_path_var = tk.StringVar()

        self._mode_lr = tk.BooleanVar(value=True)
        self._mode_rl = tk.BooleanVar(value=False)
        self._mode_bza = tk.BooleanVar(value=False)
        self._mode_bzb = tk.BooleanVar(value=False)

        self.offset_var = tk.IntVar(value=0)
        self.skip_var = tk.BooleanVar(value=False)

        # глобальный поворот
        self.global_rotation_var = tk.IntVar(value=0)

        # “поворот для текущей страницы”
        self.use_page_rotation_var = tk.BooleanVar(value=False)

        self.page_index_var = tk.IntVar(value=0)

        self._doc = None
        self._page_count = 0
        self._tk_img = None

        self._offsets: list[int] = []
        self._skips: list[bool] = []
        self._confirmed: list[bool] = []

        self._use_page_rotation: list[bool] = []
        self._page_rotations: list[int] = []  # 0/90/180/270

        self._build_ui()

    def _build_ui(self):
        top = tk.Frame(self)
        top.pack(side="top", fill="x", padx=10, pady=8)

        r1 = tk.Frame(top)
        r1.pack(fill="x")
        tk.Label(r1, text="Входной PDF:").pack(side="left")
        tk.Entry(r1, textvariable=self.input_path_var).pack(side="left", fill="x", expand=True, padx=8)
        tk.Button(r1, text="Выбрать…", command=self.choose_input).pack(side="left")
        tk.Button(r1, text="Сбросить выбор", command=self.reset_selection).pack(side="left", padx=(8, 0))

        r2 = tk.Frame(top)
        r2.pack(fill="x", pady=(6, 0))
        tk.Label(r2, text="Сохранить как:").pack(side="left")
        tk.Entry(r2, textvariable=self.output_path_var).pack(side="left", fill="x", expand=True, padx=8)
        tk.Button(r2, text="Сохранить куда", command=self.choose_output).pack(side="left")

        mode_frame = tk.LabelFrame(top, text="Режим обработки (только один вариант)")
        mode_frame.pack(fill="x", pady=(8, 0))
        r3a = tk.Frame(mode_frame)
        r3a.pack(fill="x", padx=8, pady=(6, 2))
        r3b = tk.Frame(mode_frame)
        r3b.pack(fill="x", padx=8, pady=(2, 6))

        tk.Checkbutton(
            r3a,
            text="Слева → затем справа",
            variable=self._mode_lr,
            command=lambda: self._exclusive_output_mode("lr"),
        ).pack(side="left", padx=(0, 12))
        tk.Checkbutton(
            r3a,
            text="Справа → затем слева",
            variable=self._mode_rl,
            command=lambda: self._exclusive_output_mode("rl"),
        ).pack(side="left", padx=(0, 12))
        tk.Checkbutton(
            r3a,
            text="Брошюра: зигзаг по листам (N–1, 2–(N–1), …)",
            variable=self._mode_bza,
            command=lambda: self._exclusive_output_mode("bza"),
        ).pack(side="left", padx=(0, 12))

        tk.Checkbutton(
            r3b,
            text="Брошюра: зигзаг по листам (1–N, (N–1)–2, …)",
            variable=self._mode_bzb,
            command=lambda: self._exclusive_output_mode("bzb"),
        ).pack(side="left")

        r4 = tk.Frame(top)
        r4.pack(fill="x", pady=(8, 0))

        tk.Button(r4, text="<", width=4, command=lambda: self.move_offset(-STEP_PX)).pack(side="left", padx=(0, 4))
        tk.Button(r4, text=">", width=4, command=lambda: self.move_offset(+STEP_PX)).pack(side="left", padx=4)
        self.offset_label = tk.Label(r4, text="Смещение: 0 px (0 = центр)")
        self.offset_label.pack(side="left", padx=10)

        tk.Checkbutton(
            r4, text="Пропустить текущую страницу",
            variable=self.skip_var,
            command=self._on_skip_toggle
        ).pack(side="left", padx=(18, 0))

        # Новый чекбокс
        tk.Checkbutton(
            r4, text="Поворот для текущей страницы",
            variable=self.use_page_rotation_var,
            command=self._on_use_page_rotation_toggle
        ).pack(side="left", padx=(10, 0))

        tk.Label(r4, text="Поворот:").pack(side="left", padx=(20, 6))
        tk.Button(r4, text="⟲ -90", width=7, command=lambda: self.rotate(-90)).pack(side="left", padx=4)
        tk.Button(r4, text="⟳ +90", width=7, command=lambda: self.rotate(+90)).pack(side="left", padx=4)
        self.rot_label = tk.Label(r4, text="0°")
        self.rot_label.pack(side="left", padx=10)

        mid = tk.Frame(self)
        mid.pack(side="top", fill="both", expand=True, padx=10, pady=8)

        left_panel = tk.Frame(mid, width=420)
        left_panel.pack(side="left", fill="y", padx=(0, 10))
        left_panel.pack_propagate(False)

        nav = tk.LabelFrame(left_panel, text="Предпросмотр (клавиатура)")
        nav.pack(fill="x")

        nav_row = tk.Frame(nav)
        nav_row.pack(fill="x", padx=8, pady=8)
        tk.Button(nav_row, text="←", width=4, command=self.prev_page).pack(side="left")
        self.page_label = tk.Label(nav_row, text="Стр.: - / -", width=16, anchor="center")
        self.page_label.pack(side="left", padx=8)
        tk.Button(nav_row, text="→", width=4, command=self.next_page).pack(side="left")

        tk.Label(
            nav,
            text="Клик по предпросмотру даёт фокус.\n"
                 "←/→: смещение линии, ↑/↓: смена страницы.\n"
                 "Смена страницы подтверждает настройки."
        ).pack(fill="x", padx=8, pady=(0, 8))

        table_frame = tk.LabelFrame(left_panel, text="Настройки по страницам")
        table_frame.pack(fill="both", expand=True, pady=(10, 0))

        self.tree = ttk.Treeview(
            table_frame,
            columns=("page", "offset", "skip", "rot", "ok"),
            show="headings",
            selectmode="browse",
            height=18
        )
        self.tree.heading("page", text="Стр.")
        self.tree.heading("offset", text="Смещение (px)")
        self.tree.heading("skip", text="Пропустить")
        self.tree.heading("rot", text="Поворот")
        self.tree.heading("ok", text="Подтв.")

        self.tree.column("page", width=50, anchor="center")
        self.tree.column("offset", width=120, anchor="center")
        self.tree.column("skip", width=90, anchor="center")
        self.tree.column("rot", width=70, anchor="center")
        self.tree.column("ok", width=70, anchor="center")

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.tree.bind("<<TreeviewSelect>>", self.on_table_select)
        self.tree.bind("<Double-1>", self.on_table_double_click)

        self.canvas = tk.Canvas(mid, bg="#1e1e1e", highlightthickness=1, highlightbackground="#444")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.canvas.configure(takefocus=True)
        self.canvas.bind("<Button-1>", lambda e: self.canvas.focus_set())
        self.canvas.bind("<Left>", lambda e: self.move_offset(-STEP_PX))
        self.canvas.bind("<Right>", lambda e: self.move_offset(+STEP_PX))
        self.canvas.bind("<Up>", lambda e: self.prev_page())
        self.canvas.bind("<Down>", lambda e: self.next_page())

        # Горячие клавиши поворота страницы
        self.canvas.bind("<Control-Left>", lambda e: self.rotate_current_page(-90))
        self.canvas.bind("<Control-Right>", lambda e: self.rotate_current_page(+90))

        bottom = tk.Frame(self)
        bottom.pack(side="bottom", fill="x", padx=10, pady=10)
        tk.Button(bottom, text="Обработать и сохранить", command=self.run, height=2).pack(side="left")

        self.canvas.bind("<Configure>", lambda e: self.refresh_preview())
        self._update_labels()

    def _exclusive_output_mode(self, key: str):
        """Взаимоисключающие чекбоксы режима выхода."""
        d = {"lr": self._mode_lr, "rl": self._mode_rl, "bza": self._mode_bza, "bzb": self._mode_bzb}
        if not d[key].get():
            # нельзя оставить все выключенными — по умолчанию «слева → справа»
            self._mode_lr.set(True)
            self._mode_rl.set(False)
            self._mode_bza.set(False)
            self._mode_bzb.set(False)
            self.refresh_preview()
            return
        for k, v in d.items():
            v.set(k == key)
        self.refresh_preview()

    def _get_output_mode(self) -> str:
        if self._mode_bza.get():
            return OUTPUT_BROCHURE_ZIGZAG_A
        if self._mode_bzb.get():
            return OUTPUT_BROCHURE_ZIGZAG_B
        if self._mode_rl.get():
            return OUTPUT_RIGHT_LEFT
        return OUTPUT_LEFT_RIGHT

    # -------- helpers --------
    def _effective_user_rotation_for_page(self, i: int) -> int:
        """Какой user_rotation применяем к странице i: глобальный или индивидуальный"""
        if self._use_page_rotation[i]:
            return _norm_rot(self._page_rotations[i])
        return _norm_rot(self.global_rotation_var.get())

    def reset_selection(self):
        try:
            if self._doc is not None:
                self._doc.close()
        except Exception:
            pass

        self._doc = None
        self._page_count = 0
        self._tk_img = None

        self.input_path_var.set("")
        self.output_path_var.set("")
        self.page_index_var.set(0)
        self._mode_lr.set(True)
        self._mode_rl.set(False)
        self._mode_bza.set(False)
        self._mode_bzb.set(False)

        self.offset_var.set(0)
        self.skip_var.set(False)
        self.global_rotation_var.set(0)
        self.use_page_rotation_var.set(False)

        self._offsets = []
        self._skips = []
        self._confirmed = []
        self._use_page_rotation = []
        self._page_rotations = []

        for item in self.tree.get_children():
            self.tree.delete(item)

        self.canvas.delete("all")
        self._update_labels()
        self.canvas.create_text(10, 10, anchor="nw", fill="white", text="Выбор сброшен. Выберите PDF заново.")

    def _init_per_page_settings(self, n_pages: int):
        self._offsets = [0] * n_pages
        self._skips = [False] * n_pages
        self._confirmed = [False] * n_pages

        self._use_page_rotation = [False] * n_pages
        self._page_rotations = [0] * n_pages

        for item in self.tree.get_children():
            self.tree.delete(item)

        for i in range(n_pages):
            self.tree.insert("", "end", iid=str(i), values=(i + 1, 0, "Нет", "0°", ""))

        if n_pages > 0:
            self.tree.selection_set("0")
            self.tree.see("0")

    def _update_table_row(self, i: int):
        rot = self._effective_user_rotation_for_page(i)
        self.tree.item(
            str(i),
            values=(
                i + 1,
                self._offsets[i],
                "Да" if self._skips[i] else "Нет",
                f"{rot}°",
                "✓" if self._confirmed[i] else ""
            )
        )

    def _load_settings_for_page(self, i: int):
        self.offset_var.set(int(self._offsets[i]))
        self.skip_var.set(bool(self._skips[i]))
        self.use_page_rotation_var.set(bool(self._use_page_rotation[i]))

    def _inherit_offset_from_previous(self, new_idx: int):
        if new_idx <= 0 or new_idx >= self._page_count:
            return
        if self._confirmed[new_idx]:
            return
        self._offsets[new_idx] = self._offsets[new_idx - 1]
        self._update_table_row(new_idx)

    def _update_labels(self):
        self.offset_label.config(text=f"Смещение: {self.offset_var.get()} px (0 = центр)")

        # показываем "какой поворот сейчас редактируем"
        i = self.page_index_var.get()
        if 0 <= i < self._page_count and self.use_page_rotation_var.get():
            rot = _norm_rot(self._page_rotations[i])
            self.rot_label.config(text=f"{rot}° (текущая стр.)")
        else:
            rot = _norm_rot(self.global_rotation_var.get())
            self.rot_label.config(text=f"{rot}° (все)")

        if self._page_count > 0:
            self.page_label.config(text=f"Стр.: {self.page_index_var.get() + 1} / {self._page_count}")
        else:
            self.page_label.config(text="Стр.: - / -")

    # -------- UI actions --------
    def choose_input(self):
        path = filedialog.askopenfilename(title="Выберите PDF", filetypes=[("PDF files", "*.pdf")])
        if not path:
            return

        self.input_path_var.set(path)
        base, ext = os.path.splitext(path)
        self.output_path_var.set(base + "_split" + ext)

        try:
            if self._doc is not None:
                self._doc.close()
            self._doc = fitz.open(path)
            self._page_count = self._doc.page_count
            self.page_index_var.set(0)

            self._init_per_page_settings(self._page_count)
            self._load_settings_for_page(0)
            self._update_table_row(0)
            self.refresh_preview()
        except Exception as e:
            self.reset_selection()
            messagebox.showerror("Ошибка", f"Не удалось открыть PDF:\n{e}")

    def choose_output(self):
        path = filedialog.asksaveasfilename(
            title="Сохранить результат как",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
        )
        if path:
            self.output_path_var.set(path)

    def prev_page(self):
        if self._page_count <= 0:
            return

        idx = self.page_index_var.get()
        if idx > 0:
            idx -= 1
            self.page_index_var.set(idx)
            self.tree.selection_set(str(idx))
            self.tree.see(str(idx))
            self._load_settings_for_page(idx)
            self._update_labels()
            self.refresh_preview()

    def next_page(self):
        if self._page_count <= 0:
            return

        idx = self.page_index_var.get()
        if idx < self._page_count - 1:
            idx += 1
            self._inherit_offset_from_previous(idx)

            self.page_index_var.set(idx)
            self.tree.selection_set(str(idx))
            self.tree.see(str(idx))
            self._load_settings_for_page(idx)
            self._update_labels()
            self.refresh_preview()

    def on_table_select(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        new_idx = int(sel[0])

        self._inherit_offset_from_previous(new_idx)
        self.page_index_var.set(new_idx)
        self._load_settings_for_page(new_idx)
        self._update_labels()
        self.refresh_preview()

    def on_table_double_click(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        i = int(sel[0])
        self._skips[i] = not self._skips[i]
        self._confirmed[i] = True
        if i == self.page_index_var.get():
            self.skip_var.set(self._skips[i])
        self._update_table_row(i)
        self.refresh_preview()

    def _on_skip_toggle(self):
        if self._page_count <= 0:
            return
        i = self.page_index_var.get()
        self._skips[i] = bool(self.skip_var.get())
        self._confirmed[i] = True
        self._update_table_row(i)
        self.refresh_preview()

    def _on_use_page_rotation_toggle(self):
        if self._page_count <= 0:
            return
        i = self.page_index_var.get()
        self._use_page_rotation[i] = bool(self.use_page_rotation_var.get())
        self._confirmed[i] = True
        self._update_table_row(i)
        self._update_labels()
        self.refresh_preview()

    def move_offset(self, delta):
        if self._page_count <= 0:
            return
        i = self.page_index_var.get()
        self.offset_var.set(self.offset_var.get() + delta)

        self._offsets[i] = int(self.offset_var.get())
        self._confirmed[i] = True
        self._update_table_row(i)

        self._update_labels()
        self.refresh_preview()

    def rotate(self, delta):
        if self._page_count <= 0:
            # можно разрешить менять глобальный без документа, но не критично
            self.global_rotation_var.set(_norm_rot(self.global_rotation_var.get() + delta))
            self._update_labels()
            return

        i = self.page_index_var.get()

        if self.use_page_rotation_var.get():
            self._page_rotations[i] = _norm_rot(self._page_rotations[i] + delta)
        else:
            self.global_rotation_var.set(_norm_rot(self.global_rotation_var.get() + delta))

        self._confirmed[i] = True
        self._update_table_row(i)
        self._update_labels()
        self.refresh_preview()

    def rotate_current_page(self, delta):
        if self._page_count <= 0:
            return

        i = self.page_index_var.get()

        # ✅ автоматически включаем режим "Поворот для текущей страницы"
        self._use_page_rotation[i] = True
        self.use_page_rotation_var.set(True)

        # ✅ меняем поворот
        self._page_rotations[i] = _norm_rot(self._page_rotations[i] + delta)

        self._confirmed[i] = True
        self._update_table_row(i)
        self._update_labels()
        self.refresh_preview()

    def refresh_preview(self):
        self._update_labels()
        self.canvas.delete("all")

        if self._doc is None or self._page_count == 0:
            self.canvas.create_text(10, 10, anchor="nw", fill="white", text="Выберите PDF для предпросмотра.")
            return

        idx = self.page_index_var.get()
        page = self._doc.load_page(idx)

        # Рендерим с эффективным user_rotation для этой страницы (глобальный или индивидуальный)
        rot_for_page = self._effective_user_rotation_for_page(idx)
        mat = fitz.Matrix(PREVIEW_ZOOM, PREVIEW_ZOOM).prerotate(rot_for_page)
        pix = page.get_pixmap(matrix=mat, alpha=False)

        self._tk_img = tk.PhotoImage(data=pix.tobytes("ppm"))

        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()
        iw = self._tk_img.width()
        ih = self._tk_img.height()

        x0 = max(0, (cw - iw) // 2)
        y0 = max(0, (ch - ih) // 2)

        self.canvas.create_image(x0, y0, anchor="nw", image=self._tk_img)

        if self.skip_var.get():
            self.canvas.create_text(
                x0 + 10, y0 + 10, anchor="nw",
                fill="yellow",
                text="Эта страница будет пропущена (без разреза)."
            )
            return

        cut_x = x0 + pix.width / 2.0 + self.offset_var.get()
        cut_x = max(x0 + 1, min(x0 + pix.width - 1, cut_x))
        self.canvas.create_line(cut_x, y0, cut_x, y0 + pix.height, fill="red", width=2)

    def run(self):
        in_path = self.input_path_var.get().strip()
        out_path = self.output_path_var.get().strip()

        if not in_path or not os.path.isfile(in_path):
            messagebox.showerror("Ошибка", "Выберите существующий входной PDF.")
            return
        if not out_path:
            messagebox.showerror("Ошибка", "Выберите путь для сохранения результата.")
            return
        if self._page_count <= 0:
            messagebox.showerror("Ошибка", "Нет открытого документа.")
            return

        mode = self._get_output_mode()
        if mode in (OUTPUT_BROCHURE_ZIGZAG_A, OUTPUT_BROCHURE_ZIGZAG_B) and any(self._skips):
            messagebox.showerror(
                "Ошибка",
                "Режимы «Брошюра: зигзаг по листам» требуют, чтобы ни одна страница не была помечена «Пропустить».",
            )
            return

        try:
            split_spreads_per_page(
                input_path=in_path,
                output_path=out_path,
                output_mode=mode,
                global_rotation=int(self.global_rotation_var.get()),
                preview_zoom=float(PREVIEW_ZOOM),
                per_page_offset_px=list(self._offsets),
                per_page_skip=list(self._skips),
                per_page_use_rotation=list(self._use_page_rotation),
                per_page_rotation=list(self._page_rotations),
            )
        except Exception as e:
            messagebox.showerror("Ошибка обработки", str(e))
            return

        messagebox.showinfo("Готово", f"Файл сохранён:\n{out_path}")


if __name__ == "__main__":
    App().mainloop()

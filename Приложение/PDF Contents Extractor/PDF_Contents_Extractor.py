import tkinter as tk
from tkinter import filedialog, messagebox
from pypdf import PdfReader

"""
Глобальные горячие клавиши и контекстное меню для стандартных виджетов.
Единая логика копирования/вырезания/вставки/отмены и простые снимки для корректного Undo.
Подключать один раз на окно верхнего уровня; обработчики распространяются на всех потомков.
"""
import tkinter as tk
from tkinter import ttk
import sys


def setup_keyboard_shortcuts(root):
    def _is_text_widget(w):
        return w.winfo_class() == 'Text'

    def _is_entry_like(w):
        return w.winfo_class() in ('Entry', 'TEntry', 'Spinbox', 'TSpinbox', 'TCombobox')

    def _is_editable(w):
        try:
            return str(w.cget('state')) == 'normal'
        except Exception:
            return True

    def _get_text(w):
        try:
            if _is_text_widget(w):
                return w.get('1.0', 'end-1c')
            return w.get()
        except Exception:
            return ""

    def _set_text(w, text):
        try:
            if _is_text_widget(w):
                w.delete('1.0', 'end')
                w.insert('1.0', text)
            else:
                w.delete(0, 'end')
                w.insert(0, text)
        except Exception:
            pass

    def _delete_selection(w):
        try:
            if _is_text_widget(w):
                w.delete('sel.first', 'sel.last')
            else:
                w.delete('sel.first', 'sel.last')
        except Exception:
            pass

    def _insert_at_cursor(w, text):
        try:
            if _is_text_widget(w):
                w.insert('insert', text)
            else:
                w.insert('insert', text)
        except Exception:
            pass

    def _get_selected_text(w):
        try:
            if _is_text_widget(w):
                return w.get('sel.first', 'sel.last')
            s = w.get()
            i1 = w.index('sel.first')
            i2 = w.index('sel.last')
            return s[i1:i2]
        except Exception:
            return ""

    if not hasattr(root, '_undo_snap'):
        root._undo_snap = {}

    def _snap_init(w):
        val = _get_text(w)
        root._undo_snap[w] = {'prev': val, 'curr': val}

    def _snap_update(w):
        val = _get_text(w)
        d = root._undo_snap.get(w)
        if d is None:
            _snap_init(w)
            return
        if val != d['curr']:
            d['prev'] = d['curr']
            d['curr'] = val

    def _on_copy(e):
        w = e.widget
        try:
            txt = _get_selected_text(w)
            if txt:
                root.clipboard_clear()
                root.clipboard_append(txt)
        except Exception:
            pass
        return 'break'

    def _on_cut(e):
        w = e.widget
        if not _is_editable(w):
            return 'break'
        try:
            txt = _get_selected_text(w)
            if txt:
                root.clipboard_clear()
                root.clipboard_append(txt)
                _delete_selection(w)
                _snap_update(w)
        except Exception:
            pass
        return 'break'

    def _on_paste(e):
        w = e.widget
        if not _is_editable(w):
            return 'break'
        try:
            txt = root.clipboard_get()
        except tk.TclError:
            txt = ""
        if not txt:
            return 'break'
        _delete_selection(w)
        _insert_at_cursor(w, txt)
        _snap_update(w)
        return 'break'

    def _on_undo(e):
        w = e.widget
        if _is_text_widget(w):
            try:
                w.event_generate('<<Undo>>')
                return 'break'
            except Exception:
                pass
        if _is_entry_like(w):
            d = root._undo_snap.get(w)
            if d:
                _set_text(w, d['prev'])
                d['curr'] = d['prev']
            return 'break'
        return 'break'

    def _on_select_all(e):
        w = e.widget
        try:
            if _is_text_widget(w):
                w.tag_add(tk.SEL, '1.0', 'end')
            else:
                w.selection_range(0, tk.END)
        except Exception:
            pass
        return 'break'

    def _on_key_with_ctrl(e):
        """
        Универсальный обработчик для Ctrl+Key.
        Работает независимо от раскладки.
        """
        kc = getattr(e, 'keycode', 0)

        if kc == 67:
            return _on_copy(e)
        elif kc == 88:
            return _on_cut(e)
        elif kc == 86:
            return _on_paste(e)
        elif kc == 90:
            return _on_undo(e)
        elif kc == 65:
            return _on_select_all(e)

        return None

    # Привязка горячих клавиш для всех виджетов
    for cls in ('Entry', 'TEntry', 'Text', 'TCombobox', 'Spinbox', 'TSpinbox'):
        # Английская раскладка (стандартные привязки)
        root.bind_class(cls, '<Control-c>', _on_copy, add='+')
        root.bind_class(cls, '<Control-C>', _on_copy, add='+')
        root.bind_class(cls, '<Control-x>', _on_cut, add='+')
        root.bind_class(cls, '<Control-X>', _on_cut, add='+')
        root.bind_class(cls, '<Control-v>', _on_paste, add='+')
        root.bind_class(cls, '<Control-V>', _on_paste, add='+')
        root.bind_class(cls, '<Control-z>', _on_undo, add='+')
        root.bind_class(cls, '<Control-Z>', _on_undo, add='+')

        # Универсальный обработчик для любых Ctrl+Key (включая русскую раскладку)
        root.bind_class(cls, '<Control-Key>', _on_key_with_ctrl, add='+')

        # Дополнительные привязки (Insert/Delete)
        root.bind_class(cls, '<Control-Insert>', _on_copy, add='+')
        root.bind_class(cls, '<Shift-Insert>', _on_paste, add='+')
        root.bind_class(cls, '<Shift-Delete>', _on_cut, add='+')

        # macOS (Command)
        if sys.platform == 'darwin':
            root.bind_class(cls, '<Command-c>', _on_copy, add='+')
            root.bind_class(cls, '<Command-x>', _on_cut, add='+')
            root.bind_class(cls, '<Command-v>', _on_paste, add='+')
            root.bind_class(cls, '<Command-z>', _on_undo, add='+')

        # Снимки для Undo
        root.bind_class(cls, '<FocusIn>', lambda e: _snap_init(e.widget), add='+')
        for seq in ('<KeyRelease>', '<<Paste>>', '<<Cut>>'):
            root.bind_class(cls, seq, lambda e: _snap_update(e.widget), add='+')


def install_context_menu(root):
    if getattr(root, '_global_cmenu_installed', False):
        return
    root._global_cmenu_installed = True

    cmenu = tk.Menu(root, tearoff=0)
    root._global_cmenu = cmenu
    root._cmenu_widget = None

    def _cls(w):
        try:
            return w.winfo_class()
        except Exception:
            return ''

    def _is_text(w):     return _cls(w) == 'Text'
    def _is_entry(w):    return _cls(w) in ('Entry','TEntry')
    def _is_combo(w):    return _cls(w) in ('TCombobox',)
    def _is_spin(w):     return _cls(w) in ('Spinbox','TSpinbox')
    def _is_listbox(w):  return _cls(w) == 'Listbox'
    def _is_tree(w):     return _cls(w) == 'Treeview'

    def _is_editable(w):
        try:
            return str(w.cget('state')) == 'normal'
        except Exception:
            return True

    def _has_selection(w):
        try:
            if _is_text(w):
                w.get('sel.first', 'sel.last'); return True
            if _is_entry(w) or _is_spin(w) or _is_combo(w):
                return bool(w.selection_present())
            if _is_listbox(w):
                return bool(w.curselection())
            if _is_tree(w):
                return bool(w.selection())
        except Exception:
            pass
        return False

    def _get_selection_text(w):
        try:
            if _is_text(w):
                return w.get('sel.first', 'sel.last')
            if _is_entry(w) or _is_spin(w) or _is_combo(w):
                s = w.get()
                try:
                    i1 = w.index('sel.first'); i2 = w.index('sel.last'); return s[i1:i2]
                except Exception:
                    return ''
            if _is_listbox(w):
                sel = w.curselection()
                if not sel: return ''
                return '\n'.join(w.get(i) for i in sel)
            if _is_tree(w):
                sel = w.selection()
                if not sel: return ''
                cols = w.cget('columns') or ()
                lines = []
                for iid in sel:
                    vals = w.item(iid, 'values') or ()
                    if vals:
                        lines.append('\t'.join(str(v) for v in vals))
                    else:
                        lines.append(str(w.item(iid, 'text') or ''))
                return '\n'.join(lines)
        except Exception:
            pass
        return ''

    def _target():
        w = getattr(root, '_cmenu_widget', None)
        if w and getattr(w, 'winfo_exists', lambda: False)():
            return w
        try:
            w = root.focus_get()
            if w and getattr(w, 'winfo_exists', lambda: False)():
                return w
        except Exception:
            pass
        return None

    def _on_undo():
        w = _target()
        if not w: return
        try:
            w.focus_set()
            w.event_generate('<Control-z>')
        except Exception:
            pass

    def _on_cut():
        w = _target()
        if not w: return
        try:
            can_edit = _is_text(w) or _is_entry(w) or _is_spin(w) or _is_combo(w)
            if not can_edit or not _is_editable(w): return
            if not _has_selection(w): return
            w.focus_set()
            w.event_generate('<<Cut>>')
        except Exception:
            pass

    def _on_copy():
        w = _target()
        if not w: return
        try:
            if _get_selection_text(w):
                w.focus_set()
                w.event_generate('<<Copy>>')
        except Exception:
            pass

    def _on_paste():
        w = _target()
        if not w: return
        try:
            can_edit = _is_text(w) or _is_entry(w) or _is_spin(w) or _is_combo(w)
            if not can_edit or not _is_editable(w): return
            w.focus_set()
            w.event_generate('<<Paste>>')
        except Exception:
            pass

    def _on_select_all():
        w = _target()
        if not w: return
        try:
            if _is_text(w):
                w.tag_add(tk.SEL, '1.0', 'end')
            elif _is_entry(w) or _is_spin(w) or _is_combo(w):
                w.selection_range(0, tk.END)
            elif _is_listbox(w):
                w.selection_set(0, tk.END)
            elif _is_tree(w):
                for iid in w.get_children(''):
                    w.selection_add(iid)
            w.focus_set()
        except Exception:
            pass

    cmenu.add_command(label="Отменить", command=_on_undo)
    cmenu.add_separator()
    cmenu.add_command(label="Вырезать", command=_on_cut)
    cmenu.add_command(label="Копировать", command=_on_copy)
    cmenu.add_command(label="Вставить", command=_on_paste)
    cmenu.add_separator()
    cmenu.add_command(label="Выделить всё", command=_on_select_all)

    def _show_menu(event):
        w = event.widget
        root._cmenu_widget = w
        allowed = any((_cls(w) in ('Text','Entry','TEntry','TCombobox','Spinbox','TSpinbox','Listbox','Treeview'),))
        if not allowed:
            return 'break'
        editable = _is_editable(w)
        has_sel = _has_selection(w)
        try:
            has_clip = bool(root.clipboard_get())
        except tk.TclError:
            has_clip = False
        can_edit = _cls(w) in ('Text','Entry','TEntry','TCombobox','Spinbox','TSpinbox')
        cmenu.entryconfigure("Отменить", state=(tk.NORMAL if (can_edit and editable) else tk.DISABLED))
        cmenu.entryconfigure("Вырезать", state=(tk.NORMAL if (can_edit and editable and has_sel) else tk.DISABLED))
        cmenu.entryconfigure("Копировать", state=(tk.NORMAL if has_sel else tk.DISABLED))
        cmenu.entryconfigure("Вставить", state=(tk.NORMAL if (can_edit and editable and has_clip) else tk.DISABLED))
        cmenu.entryconfigure("Выделить всё", state=tk.NORMAL)
        try:
            cmenu.tk_popup(event.x_root, event.y_root)
        finally:
            try: cmenu.grab_release()
            except Exception: pass
        return 'break'

    root.bind_all('<Button-3>', _show_menu, add='+')
    root.bind_all('<Button-2>', _show_menu, add='+')


def enable_undo_for_descendants(widget):
    try:
        cls = widget.winfo_class()
        if cls in ('Entry', 'TEntry'):
            try:
                widget.configure(undo=True, autoseparators=True, maxundo=-1)
            except Exception:
                try:
                    widget.configure(undo=True)
                except Exception:
                    pass
        elif cls == 'Text':
            try:
                widget.configure(undo=True, autoseparators=True, maxundo=-1)
            except Exception:
                try:
                    widget.configure(undo=True)
                except Exception:
                    pass
    except Exception:
        pass
    for child in getattr(widget, 'winfo_children', lambda: [])():
        enable_undo_for_descendants(child)

class PDFContentsExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Contents Extractor")
        self.root.geometry("800x600")

        self.file_path = None
        self.file_path_var = tk.StringVar()

        # Контейнер для кнопки и поля
        file_frame = tk.Frame(root)
        file_frame.pack(pady=10, fill="x", padx=10)

        # Поле отображения выбранного файла
        self.file_entry = tk.Entry(
            file_frame,
            textvariable=self.file_path_var
        )
        self.file_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))

        # Кнопка загрузки PDF
        self.load_button = tk.Button(
            file_frame,
            text="Загрузить PDF",
            command=self.load_pdf,
            width=20
        )
        self.load_button.pack(side="right")

        # Кнопка обработки
        self.process_button = tk.Button(
            root,
            text="Извлечь закладки",
            command=self.extract_bookmarks,
            width=20
        )
        self.process_button.pack(pady=5)

        # Текстовое поле для вывода (редактируемое)
        self.text_area = tk.Text(root, wrap="word")
        self.text_area.pack(expand=True, fill="both", padx=10, pady=10)

        # ✅ Подключаем глобальные горячие клавиши и контекстное меню
        setup_keyboard_shortcuts(self.root)
        install_context_menu(self.root)
        enable_undo_for_descendants(self.root)

    def load_pdf(self):
        self.file_path = filedialog.askopenfilename(
            filetypes=[("PDF files", "*.pdf")]
        )
        if self.file_path:
            self.file_path_var.set(self.file_path)
            messagebox.showinfo("Файл загружен", f"Выбран файл:\n{self.file_path}")

    def extract_bookmarks(self):
        if not self.file_path:
            messagebox.showwarning("Ошибка", "Сначала загрузите PDF файл.")
            return

        try:
            reader = PdfReader(self.file_path)
            outlines = reader.outline

            bookmarks = []
            self._extract_recursive(outlines, bookmarks)

            # Очистка текстового поля
            self.text_area.delete(1.0, tk.END)

            # Запись без пустых строк
            result = "\n".join(bookmarks)
            self.text_area.insert(tk.END, result)

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при обработке PDF:\n{str(e)}")

    def _extract_recursive(self, outlines, bookmarks_list):
        for item in outlines:
            if isinstance(item, list):
                self._extract_recursive(item, bookmarks_list)
            else:
                try:
                    title = item.title.strip()
                    if title:
                        bookmarks_list.append(title)
                except AttributeError:
                    pass


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFContentsExtractor(root)
    root.mainloop()

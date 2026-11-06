"""Окно обработки заявок: загрузка PDF, распознавание, правка полей, материалы, отчёт."""
import json
import os
import queue
import threading
import tkinter as tk
from dataclasses import asdict, dataclass, field, replace
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pymupdf

from materials import WORK_TYPES, Material, catalog, default_materials, format_qty
from report import ReportError, build_report
from request_parser import FIELDS, parse_request

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "данные"
PREVIEW_DPI = 110
DEFAULT_WORK_TYPE = "Установка"


@dataclass
class Request:
    pdf_path: Path
    data: dict[str, str] = field(default_factory=dict)
    work_type: str = DEFAULT_WORK_TYPE
    materials: list[Material] = field(default_factory=list)
    materials_edited: bool = False
    saved: bool = False

    @property
    def title(self) -> str:
        return self.pdf_path.stem


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Обработка заявок")
        self.geometry("1500x900")
        self.minsize(1000, 600)
        self.state("zoomed")  # на весь экран, чтобы нижняя панель не уходила под панель задач

        self.requests: list[Request] = []
        self.current: Request | None = None
        self.widgets: dict[str, tk.Widget] = {}
        self.preview_images: list[tk.PhotoImage] = []
        self.jobs: queue.Queue = queue.Queue()
        self.results: queue.Queue = queue.Queue()
        self._loading_form = False
        self.catalog: list[Material] = catalog()

        self._build_ui()
        threading.Thread(target=self._worker, daemon=True).start()
        self.after(200, self._poll_results)

    # ---------- интерфейс ----------
    def _build_ui(self) -> None:
        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Field.TLabel", font=("Segoe UI", 9, "bold"))
        style.configure("Small.TButton", padding=(2, 0), width=3)

        bar = ttk.Frame(self, padding=(8, 6))
        bar.pack(fill="x")
        ttk.Button(bar, text="Добавить заявку (PDF)…", command=self.add_files).pack(side="left")
        ttk.Button(bar, text="Сохранить", command=self.save_current).pack(side="left", padx=6)
        ttk.Button(bar, text="Распознать заново", command=self.reparse_current).pack(side="left")
        ttk.Button(bar, text="Сформировать отчёт…", command=self.make_report).pack(
            side="left", padx=6)
        self.status = tk.StringVar(value="Добавьте PDF-заявку")
        ttk.Label(bar, textvariable=self.status, foreground="#555").pack(side="left", padx=16)
        self.progress = ttk.Progressbar(bar, mode="indeterminate", length=160)

        panes = ttk.PanedWindow(self, orient="horizontal")
        panes.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        # список заявок
        left = ttk.Frame(panes)
        ttk.Label(left, text="Заявки", style="Field.TLabel").pack(anchor="w")
        self.listbox = tk.Listbox(left, activestyle="none", exportselection=False,
                                  font=("Segoe UI", 10))
        self.listbox.pack(fill="both", expand=True, pady=(4, 0))
        self.listbox.bind("<<ListboxSelect>>", self._on_select)
        panes.add(left, weight=1)

        # вкладки: данные заявки / работы и материалы
        self.tabs = ttk.Notebook(panes)
        self.form = self._scrollable(self.tabs, "Данные заявки")
        self._build_form()
        mat_tab = ttk.Frame(self.tabs, padding=(8, 8))
        self.tabs.add(mat_tab, text="Работы и материалы")
        self._build_materials_tab(mat_tab)
        panes.add(self.tabs, weight=3)

        # просмотр PDF
        right = ttk.Frame(panes)
        self.preview = tk.Canvas(right, background="#888", highlightthickness=0)
        pv_y = ttk.Scrollbar(right, orient="vertical", command=self.preview.yview)
        pv_x = ttk.Scrollbar(right, orient="horizontal", command=self.preview.xview)
        self.preview.configure(yscrollcommand=pv_y.set, xscrollcommand=pv_x.set)
        pv_y.pack(side="right", fill="y")
        pv_x.pack(side="bottom", fill="x")
        self.preview.pack(fill="both", expand=True)
        self._bind_wheel(self.preview, self.preview)
        panes.add(right, weight=3)

    def _scrollable(self, parent: tk.Widget, tab_title: str | None = None) -> ttk.Frame:
        """Прокручиваемая область; если указан tab_title — добавляется вкладкой."""
        outer = ttk.Frame(parent)
        canvas = tk.Canvas(outer, highlightthickness=0)
        scroll = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas, padding=(8, 0, 8, 8))
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        win = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(win, width=e.width))
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self._bind_wheel(canvas, inner)
        inner.canvas = canvas
        if tab_title:
            parent.add(outer, text=tab_title)
        else:
            outer.pack(fill="both", expand=True)
        return inner

    def _bind_wheel(self, canvas: tk.Canvas, area: tk.Widget) -> None:
        def wheel(e):
            canvas.yview_scroll(int(-e.delta / 120) * 3, "units")
        area.bind("<Enter>", lambda e: self.bind_all("<MouseWheel>", wheel))
        area.bind("<Leave>", lambda e: self.unbind_all("<MouseWheel>"))

    def _build_form(self) -> None:
        self.form.columnconfigure(0, weight=1)
        for row, f in enumerate(FIELDS):
            ttk.Label(self.form, text=f.title, style="Field.TLabel").grid(
                row=row * 2, column=0, sticky="w", pady=(8, 2))
            if f.multiline:
                w = tk.Text(self.form, height=4, wrap="word", font=("Segoe UI", 10),
                            relief="solid", borderwidth=1, undo=True)
                w.bind("<<Modified>>", lambda e, k=f.key: self._on_text_modified(e, k))
            else:
                var = tk.StringVar()
                var.trace_add("write", lambda *_, k=f.key, v=var: self._on_edit(k, v.get()))
                w = ttk.Entry(self.form, textvariable=var, font=("Segoe UI", 10))
                w.var = var
            w.grid(row=row * 2 + 1, column=0, sticky="ew")
            self.widgets[f.key] = w
        self._set_form_state("disabled")

    def _build_materials_tab(self, tab: ttk.Frame) -> None:
        top = ttk.Frame(tab)
        top.pack(fill="x")
        ttk.Label(top, text="Вид работ:", style="Field.TLabel").pack(side="left")
        self.work_type = tk.StringVar(value=DEFAULT_WORK_TYPE)
        self.work_buttons = []
        for wt in WORK_TYPES:
            b = ttk.Radiobutton(top, text=wt, value=wt, variable=self.work_type,
                                command=self._on_work_type)
            b.pack(side="left", padx=(10, 0))
            self.work_buttons.append(b)
        self.reset_btn = ttk.Button(top, text="Стандартный перечень",
                                    command=self._reset_materials)
        self.reset_btn.pack(side="right")

        self.mat_info = tk.StringVar()
        ttk.Label(tab, textvariable=self.mat_info, foreground="#a05000",
                  wraplength=600).pack(fill="x", pady=(6, 0))

        add = ttk.Frame(tab)
        add.pack(fill="x", side="bottom", pady=(8, 0))
        ttk.Label(add, text="Добавить материал:").pack(side="left")
        self.add_var = tk.StringVar()
        self.add_combo = ttk.Combobox(add, textvariable=self.add_var,
                                      values=[m.name for m in self.catalog])
        self.add_combo.pack(side="left", fill="x", expand=True, padx=6)
        self.add_btn = ttk.Button(add, text="Добавить", command=self._add_material)
        self.add_btn.pack(side="left")

        holder = ttk.Frame(tab)
        holder.pack(fill="both", expand=True, pady=(8, 0))
        self.mat_frame = self._scrollable(holder)
        self._render_materials()

    def _set_form_state(self, state: str) -> None:
        for w in self.widgets.values():
            w.configure(state=state)
        for w in getattr(self, "work_buttons", []) + [
                getattr(self, "reset_btn", None), getattr(self, "add_btn", None),
                getattr(self, "add_combo", None)]:
            if w is not None:
                w.configure(state=state)

    # ---------- редактирование полей ----------
    def _on_edit(self, key: str, value: str) -> None:
        if self._loading_form or not self.current:
            return
        self.current.data[key] = value
        self._mark_dirty()

    def _on_text_modified(self, event, key: str) -> None:
        w = event.widget
        if w.edit_modified():
            w.edit_modified(False)
            self._on_edit(key, w.get("1.0", "end-1c"))

    def _mark_dirty(self) -> None:
        if self.current and self.current.saved:
            self.current.saved = False
            self._refresh_list()

    def _fill_form(self, req: Request | None) -> None:
        self._loading_form = True
        self._set_form_state("normal")
        for f in FIELDS:
            w = self.widgets[f.key]
            value = req.data.get(f.key, "") if req else ""
            if isinstance(w, tk.Text):
                w.delete("1.0", "end")
                w.insert("1.0", value)
                w.edit_modified(False)
                w.edit_reset()
            else:
                w.var.set(value)
        self.work_type.set(req.work_type if req else DEFAULT_WORK_TYPE)
        self._render_materials()
        if not req:
            self._set_form_state("disabled")
        self._loading_form = False

    # ---------- материалы ----------
    def _on_work_type(self) -> None:
        req = self.current
        new = self.work_type.get()
        if not req or new == req.work_type:
            return
        if req.materials_edited and not messagebox.askyesno(
                "Вид работ", f"Заменить материалы стандартным перечнем «{new}»?\n"
                             "Ручные изменения количества будут потеряны."):
            self.work_type.set(req.work_type)
            return
        req.work_type = new
        req.materials = default_materials(new)
        req.materials_edited = False
        self._mark_dirty()
        self._render_materials()

    def _reset_materials(self) -> None:
        req = self.current
        if not req:
            return
        if req.materials_edited and not messagebox.askyesno(
                "Стандартный перечень", "Вернуть стандартный перечень и количества?"):
            return
        req.materials = default_materials(req.work_type)
        req.materials_edited = False
        self._mark_dirty()
        self._render_materials()

    def _materials_changed(self) -> None:
        if self.current:
            self.current.materials_edited = True
            self._mark_dirty()

    def _add_material(self) -> None:
        req = self.current
        name = self.add_var.get().strip()
        if not req or not name:
            return
        if any(m.name == name for m in req.materials):
            messagebox.showinfo("Материал", "Этот материал уже есть в списке.")
            return
        known = next((m for m in self.catalog if m.name == name), None)
        req.materials.append(replace(known) if known else Material(name, "", "шт", 1))
        self.add_var.set("")
        self._materials_changed()
        self._render_materials()
        self.after(50, lambda: self.mat_frame.canvas.yview_moveto(1.0))

    def _remove_material(self, m: Material) -> None:
        self.current.materials.remove(m)
        self._materials_changed()
        self._render_materials()

    def _change_qty(self, m: Material, var: tk.StringVar, delta: int) -> None:
        m.qty = max(0.0, round(m.qty + delta * m.step, 4))
        var.set(format_qty(m.qty))
        self._materials_changed()

    def _typed_qty(self, m: Material, var: tk.StringVar) -> None:
        try:
            q = float(var.get().replace(",", ".").strip() or 0)
        except ValueError:
            return
        if q >= 0 and q != m.qty:
            m.qty = q
            self._materials_changed()

    def _render_materials(self) -> None:
        frame = self.mat_frame
        for w in frame.winfo_children():
            w.destroy()
        req = self.current
        items = req.materials if req else []
        wt = req.work_type if req else self.work_type.get()

        if req and not items and not req.materials_edited:
            self.mat_info.set(f"Для вида работ «{wt}» нет стандартного перечня "
                              f"(файл шаблоны\\{WORK_TYPES[wt]}). Добавьте материалы вручную.")
        else:
            self.mat_info.set("Материалы с количеством 0 в отчёт не попадут." if items else "")

        frame.columnconfigure(1, weight=1)
        for c, h in enumerate(["№", "Наименование", "Код", "Ед.", "Количество", ""]):
            ttk.Label(frame, text=h, style="Field.TLabel").grid(
                row=0, column=c, sticky="w", padx=4, pady=(0, 4))
        for i, m in enumerate(items, 1):
            ttk.Label(frame, text=str(i)).grid(row=i, column=0, sticky="nw", padx=4, pady=3)
            ttk.Label(frame, text=m.name, wraplength=380, justify="left").grid(
                row=i, column=1, sticky="w", padx=4, pady=3)
            ttk.Label(frame, text=m.code).grid(row=i, column=2, sticky="w", padx=4)
            ttk.Label(frame, text=m.unit).grid(row=i, column=3, sticky="w", padx=4)

            qty = ttk.Frame(frame)
            qty.grid(row=i, column=4, sticky="w", padx=4)
            var = tk.StringVar(value=format_qty(m.qty))
            ttk.Button(qty, text="−", style="Small.TButton",
                       command=lambda m=m, v=var: self._change_qty(m, v, -1)).pack(side="left")
            e = ttk.Entry(qty, textvariable=var, width=7, justify="center")
            e.pack(side="left", padx=2)
            var.trace_add("write", lambda *_, m=m, v=var: self._typed_qty(m, v))
            ttk.Button(qty, text="+", style="Small.TButton",
                       command=lambda m=m, v=var: self._change_qty(m, v, +1)).pack(side="left")

            ttk.Button(frame, text="✕", style="Small.TButton",
                       command=lambda m=m: self._remove_material(m)).grid(
                row=i, column=5, sticky="w", padx=4)
        frame.canvas.yview_moveto(0)

    # ---------- список заявок ----------
    def _refresh_list(self) -> None:
        sel = self.listbox.curselection()
        self.listbox.delete(0, "end")
        for r in self.requests:
            mark = "" if r.saved else "  •"
            self.listbox.insert("end", r.title + mark)
        if sel:
            self.listbox.selection_set(sel[0])

    def _on_select(self, _event=None) -> None:
        sel = self.listbox.curselection()
        if not sel:
            return
        self.current = self.requests[sel[0]]
        self._fill_form(self.current)
        self._show_preview(self.current.pdf_path)

    def _select(self, req: Request) -> None:
        idx = self.requests.index(req)
        self.listbox.selection_clear(0, "end")
        self.listbox.selection_set(idx)
        self.listbox.see(idx)
        self._on_select()

    # ---------- просмотр ----------
    def _show_preview(self, path: Path) -> None:
        self.preview.delete("all")
        self.preview_images.clear()
        y = 10
        width = 0
        with pymupdf.open(path) as doc:
            for page in doc:
                pix = page.get_pixmap(dpi=PREVIEW_DPI)
                img = tk.PhotoImage(data=pix.tobytes("png"))
                self.preview_images.append(img)
                self.preview.create_image(10, y, image=img, anchor="nw")
                y += pix.height + 10
                width = max(width, pix.width)
        self.preview.configure(scrollregion=(0, 0, width + 20, y))
        self.preview.yview_moveto(0)

    # ---------- загрузка и распознавание ----------
    def add_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Выберите заявки", filetypes=[("PDF", "*.pdf"), ("Все файлы", "*.*")])
        for p in paths:
            path = Path(p)
            if any(r.pdf_path == path for r in self.requests):
                continue
            req = Request(path, materials=default_materials(DEFAULT_WORK_TYPE))
            self.requests.append(req)
            self._enqueue(req)
        self._refresh_list()

    def reparse_current(self) -> None:
        if not self.current:
            return
        if not messagebox.askyesno("Распознать заново",
                                   "Ручные правки полей этой заявки будут заменены. Продолжить?"):
            return
        self._enqueue(self.current)

    def _enqueue(self, req: Request) -> None:
        self.jobs.put(("parse", req))
        self._update_status()

    def _worker(self) -> None:
        while True:
            kind, payload = self.jobs.get()
            try:
                if kind == "parse":
                    self.results.put(("parse", payload, parse_request(str(payload.pdf_path)), None))
                else:  # отчёт
                    build_report(**payload)
                    self.results.put(("report", payload, None, None))
            except Exception as exc:  # показываем пользователю, не роняем поток
                self.results.put((kind, payload, None, exc))
            finally:
                self.jobs.task_done()

    def _poll_results(self) -> None:
        while not self.results.empty():
            kind, payload, data, error = self.results.get()
            if kind == "parse":
                self._on_parsed(payload, data, error)
            else:
                self._on_report(payload, error)
        self._update_status()
        self.after(200, self._poll_results)

    def _on_parsed(self, req: Request, data, error) -> None:
        if error:
            messagebox.showerror("Ошибка распознавания", f"{req.pdf_path.name}\n\n{error}")
            return
        req.data = data
        req.saved = False
        self._refresh_list()
        if self.current is req:
            self._fill_form(req)
        elif self.current is None:
            self._select(req)

    def _update_status(self) -> None:
        pending = self.jobs.unfinished_tasks
        if pending:
            self.status.set(f"Обработка… осталось задач: {pending}")
            if not self.progress.winfo_ismapped():
                self.progress.pack(side="left")
                self.progress.start(12)
        else:
            if self.progress.winfo_ismapped():
                self.progress.stop()
                self.progress.pack_forget()
            if not self.status.get().startswith(("Сохранено", "Отчёт")):
                self.status.set(f"Заявок: {len(self.requests)}" if self.requests
                                else "Добавьте PDF-заявку")

    # ---------- отчёт ----------
    def make_report(self) -> None:
        req = self.current
        if not req or not req.data:
            messagebox.showinfo("Отчёт", "Сначала добавьте и выберите заявку.")
            return
        used = [replace(m) for m in req.materials if m.qty > 0]
        if not used and not messagebox.askyesno(
                "Отчёт", "Список материалов пуст. Сформировать отчёт без материалов?"):
            return
        out = filedialog.asksaveasfilename(
            title="Сохранить отчёт", defaultextension=".doc",
            initialdir=str(req.pdf_path.parent),
            initialfile=f"отчет_о_выполненной_работе_{req.title}.doc",
            filetypes=[("Документ Word 97–2003", "*.doc")])
        if not out:
            return
        self.status.set("Отчёт: формирование…")
        self.jobs.put(("report", {"out_path": out, "reason": req.data.get("reason", ""),
                                  "materials": used}))
        self._update_status()

    def _on_report(self, payload: dict, error) -> None:
        if error:
            msg = str(error) if isinstance(error, ReportError) else repr(error)
            self.status.set("Отчёт: ошибка")
            messagebox.showerror("Ошибка формирования отчёта", msg)
            return
        out = payload["out_path"]
        self.status.set(f"Отчёт сохранён: {Path(out).name}")
        if messagebox.askyesno("Отчёт готов", f"Отчёт сохранён:\n{out}\n\nОткрыть его?"):
            os.startfile(out)

    # ---------- сохранение ----------
    def save_current(self) -> None:
        req = self.current
        if not req or not req.data:
            return
        DATA_DIR.mkdir(exist_ok=True)
        out = DATA_DIR / f"{req.title}.json"
        payload = {
            "source_pdf": str(req.pdf_path),
            "fields": req.data,
            "work_type": req.work_type,
            "materials": [asdict(m) for m in req.materials],
        }
        out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        req.saved = True
        self._refresh_list()
        self.status.set(f"Сохранено: {out.relative_to(APP_DIR)}")


if __name__ == "__main__":
    App().mainloop()

"""Окно обработки заявок: загрузка PDF, распознавание, редактирование полей."""
import json
import queue
import threading
import tkinter as tk
from dataclasses import dataclass, field
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pymupdf

from request_parser import FIELDS, parse_request

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "данные"
PREVIEW_DPI = 110


@dataclass
class Request:
    pdf_path: Path
    data: dict[str, str] = field(default_factory=dict)
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

        self.requests: list[Request] = []
        self.current: Request | None = None
        self.widgets: dict[str, tk.Widget] = {}
        self.preview_images: list[tk.PhotoImage] = []
        self.jobs: queue.Queue = queue.Queue()
        self.results: queue.Queue = queue.Queue()
        self._loading_form = False

        self._build_ui()
        threading.Thread(target=self._worker, daemon=True).start()
        self.after(200, self._poll_results)

    # ---------- интерфейс ----------
    def _build_ui(self) -> None:
        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Field.TLabel", font=("Segoe UI", 9, "bold"))

        bar = ttk.Frame(self, padding=(8, 6))
        bar.pack(fill="x")
        ttk.Button(bar, text="Добавить заявку (PDF)…", command=self.add_files).pack(side="left")
        ttk.Button(bar, text="Сохранить", command=self.save_current).pack(side="left", padx=6)
        ttk.Button(bar, text="Распознать заново", command=self.reparse_current).pack(side="left")
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

        # форма с полями
        mid = ttk.Frame(panes)
        canvas = tk.Canvas(mid, highlightthickness=0)
        scroll = ttk.Scrollbar(mid, orient="vertical", command=canvas.yview)
        self.form = ttk.Frame(canvas, padding=(8, 0, 8, 8))
        self.form.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        form_id = canvas.create_window((0, 0), window=self.form, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(form_id, width=e.width))
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self._bind_wheel(canvas, self.form)
        self._build_form()
        panes.add(mid, weight=3)

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

    def _set_form_state(self, state: str) -> None:
        for w in self.widgets.values():
            w.configure(state=state)

    # ---------- редактирование ----------
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
        if not req:
            self._set_form_state("disabled")
        self._loading_form = False

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
            req = Request(path)
            self.requests.append(req)
            self._enqueue(req)
        self._refresh_list()

    def reparse_current(self) -> None:
        if not self.current:
            return
        if not messagebox.askyesno("Распознать заново",
                                   "Ручные правки этой заявки будут заменены. Продолжить?"):
            return
        self._enqueue(self.current)

    def _enqueue(self, req: Request) -> None:
        self.jobs.put(req)
        self._update_status()

    def _worker(self) -> None:
        while True:
            req = self.jobs.get()
            try:
                self.results.put((req, parse_request(str(req.pdf_path)), None))
            except Exception as exc:  # показываем пользователю, не роняем поток
                self.results.put((req, None, exc))
            finally:
                self.jobs.task_done()

    def _poll_results(self) -> None:
        while not self.results.empty():
            req, data, error = self.results.get()
            if error:
                messagebox.showerror("Ошибка распознавания", f"{req.pdf_path.name}\n\n{error}")
            else:
                req.data = data
                req.saved = False
                self._refresh_list()
                if self.current is req or self.current is None:
                    self._select(req)
        self._update_status()
        self.after(200, self._poll_results)

    def _update_status(self) -> None:
        pending = self.jobs.unfinished_tasks
        if pending:
            self.status.set(f"Распознавание… осталось: {pending}")
            if not self.progress.winfo_ismapped():
                self.progress.pack(side="left")
                self.progress.start(12)
        else:
            if self.progress.winfo_ismapped():
                self.progress.stop()
                self.progress.pack_forget()
            self.status.set(f"Заявок: {len(self.requests)}" if self.requests
                            else "Добавьте PDF-заявку")

    # ---------- сохранение ----------
    def save_current(self) -> None:
        req = self.current
        if not req or not req.data:
            return
        DATA_DIR.mkdir(exist_ok=True)
        out = DATA_DIR / f"{req.title}.json"
        payload = {"source_pdf": str(req.pdf_path), "fields": req.data}
        out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        req.saved = True
        self._refresh_list()
        self.status.set(f"Сохранено: {out.relative_to(APP_DIR)}")


if __name__ == "__main__":
    App().mainloop()

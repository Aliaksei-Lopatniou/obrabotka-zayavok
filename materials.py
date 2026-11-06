"""Перечни материалов по видам работ (Excel-файлы в папке «шаблоны»)."""
from dataclasses import dataclass, replace
from pathlib import Path

import openpyxl

TEMPLATES_DIR = Path(__file__).resolve().parent / "шаблоны"

# Вид работ → файл стандартного перечня. Нет файла — перечень пустой.
WORK_TYPES: dict[str, str] = {
    "Установка": "материалы(установка).xlsx",
    "Переустановка": "материалы(переустановка).xlsx",
    "Ремонт": "материалы(ремонт).xlsx",
}


@dataclass
class Material:
    name: str
    code: str
    unit: str
    qty: float
    note: str = ""
    step: float = 0.0  # шаг кнопок +/−

    def __post_init__(self) -> None:
        # Дробные нормы (0,1 рулона, 0,04 бухты) меняются на свою же величину, остальные — на 1.
        if not self.step:
            self.step = self.qty if 0 < self.qty < 1 else 1.0


def format_qty(q: float) -> str:
    q = round(q, 4)
    return (f"{q:.4f}".rstrip("0").rstrip(".")).replace(".", ",") if q % 1 else str(int(q))


def _num(v) -> float | None:
    if v is None:
        return None
    try:
        return float(str(v).replace(",", "."))
    except ValueError:
        return None


def load_materials_xlsx(path: Path) -> list[Material]:
    """Читает таблицу формы «Приложение к отчёту»: строки, где в колонке A номер, в B название."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    items = []
    header_seen = False
    for row in ws.iter_rows(min_col=1, max_col=7, values_only=True):
        no, name, code, _unit_code, unit, qty, note = (list(row) + [None] * 7)[:7]
        if not header_seen:
            # первая строка данных идёт после строки нумерации колонок «1 2 3 …»
            header_seen = str(no).strip() == "1" and str(name).strip() == "2"
            continue
        if not name or _num(no) is None:
            continue
        code_s = "" if code is None else str(int(code) if isinstance(code, float) else code)
        items.append(Material(str(name).strip(), code_s, str(unit or "").strip(),
                              _num(qty) or 0.0, str(note or "").strip()))
    wb.close()
    return items


def default_materials(work_type: str) -> list[Material]:
    path = TEMPLATES_DIR / WORK_TYPES.get(work_type, "")
    if not path.is_file():
        return []
    return [replace(m) for m in load_materials_xlsx(path)]


def catalog() -> list[Material]:
    """Все материалы из всех перечней — для ручного добавления в отчёт."""
    seen: dict[str, Material] = {}
    for wt in WORK_TYPES:
        for m in default_materials(wt):
            seen.setdefault(m.name, m)
    return sorted(seen.values(), key=lambda m: m.name.lower())


if __name__ == "__main__":
    for wt in WORK_TYPES:
        ms = default_materials(wt)
        print(wt, len(ms))
        for m in ms:
            print(f"  {m.code:>6} {m.unit:4} {format_qty(m.qty):>5} step={format_qty(m.step):>5}  {m.name}")

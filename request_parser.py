"""Парсер заявки на выполнение ремонтных работ (скан PDF фиксированной формы)."""
import re
from dataclasses import dataclass
from difflib import SequenceMatcher

from ocr import OcrLine, OcrPage, ocr_pdf


@dataclass(frozen=True)
class Field:
    key: str
    title: str
    multiline: bool = False


# Порядок полей = порядок в бланке и в окне программы.
FIELDS: list[Field] = [
    Field("contract_date", "Договор: дата"),
    Field("contract_number", "Договор: номер"),
    Field("outgoing_number", "Исх. номер заявки"),
    Field("outgoing_date", "Исх. дата заявки"),
    Field("division", "Наименование подразделения"),
    Field("address", "Адрес подразделения, местонахождение оборудования", True),
    Field("equipment", "Наименование оборудования"),
    Field("model", "Марка и модель оборудования"),
    Field("inventory_number", "Инвентарный номер"),
    Field("serial_number", "Заводской №"),
    Field("reason", "Причина вызова (неисправность)", True),
    Field("applicant_position", "Должность лица, подавшего заявку"),
    Field("applicant_name", "Имя, отчество лица, подавшего заявку"),
    Field("applicant_surname", "Фамилия исполнителя письма"),
    Field("applicant_phone", "Контактный телефон заявителя"),
    Field("email", "E-mail"),
    Field("contact_name", "Контактное лицо в подразделении"),
    Field("contact_phone", "Телефон контактного лица"),
    Field("replacement", "Потребность в подменном оборудовании"),
    Field("request_date", "Дата подачи заявки"),
    Field("signer_position", "Должность подписанта", True),
    Field("signer_name", "Подписант"),
]

# Подписи слева от полей бланка. Для одинаковых подписей берётся n-е вхождение.
ANCHORS: list[tuple[str, str]] = [
    ("division", "наименование подразделения"),
    ("address", "адрес подразделения"),
    ("equipment", "наименование оборудования"),
    ("model", "марка и модель"),
    ("inventory_number", "инвентарный номер"),
    ("serial_number", "заводской"),
    ("reason", "причина вызова"),
    ("applicant_position", "должность лица"),
    ("applicant_name", "имя отчество лица"),
    ("applicant_phone", "контактный телефон"),
    ("email", "email"),
    ("contact_name", "контактное лицо"),
    ("contact_phone", "контактный телефон"),
    ("replacement", "потребность в"),
    ("request_date", "дата подачи"),
]


def _norm(s: str) -> str:
    return re.sub(r"[^a-zа-яё0-9 ]", "", s.lower().replace("ё", "е")).strip()


def _matches(text: str, keyword: str) -> bool:
    t = _norm(text).replace(" ", "")
    k = keyword.replace(" ", "")
    if k in t:
        return True
    head = t[: len(k)]
    return len(head) >= len(k) * 0.8 and SequenceMatcher(None, head, k).ratio() >= 0.8


def _fix_number_sign(s: str) -> str:
    # OCR читает «№» как «N2», «N9», «лг2», «Nе» и т.п.
    s = re.sub(r"(?:N[29°оo]|N[eе]|лг2|Nº)\s*(?=\d)", "№", s)
    return re.sub(r"\bN\s+(?=\d)", "№ ", s)


def _join(lines: list[OcrLine]) -> str:
    out = ""
    for ln in lines:
        t = ln.text.strip()
        if not out:
            out = t
        elif out.endswith("-") and t[:1].islower():
            out = out[:-1] + t
        else:
            out += " " + t
    return re.sub(r"\s+", " ", out).strip()


def _clean_code(s: str) -> str:
    """Марка, номера: без пробелов, заглавные, строчная l → 1."""
    s = s.replace(" ", "").replace("l", "1")
    return s.upper()


def _clean_phone(s: str) -> str:
    return re.sub(r"[^\d+\-() ]", "", s.replace("O", "0").replace("o", "0")).strip()


def parse_pages(pages: list[OcrPage]) -> dict[str, str]:
    data = {f.key: "" for f in FIELDS}
    page = pages[0]
    lines = sorted(page.lines, key=lambda l: (l.y, l.x))
    mid = page.width * 0.5
    tol = page.height * 0.006  # ~20 px при 300 dpi

    full_text = "\n".join(l.text for l in lines)

    # --- договор: строку перечитываем с увеличением, мелкие цифры так надёжнее ---
    contract_re = re.compile(
        r"от\s+(\d{1,2}[.,]\d{1,2}[.,]\d{4})\s*(.*?)\s*направляем", re.I | re.S
    )
    contract_line = next((l for l in lines if _matches(l.text, "на основании договора")), None)
    candidates = []
    if contract_line:
        candidates.append(page.ocr_region(
            contract_line.x - 10, contract_line.y - 15,
            contract_line.x + contract_line.w + 10, contract_line.y + contract_line.h + 15,
        ))
    candidates.append(full_text)
    for text in candidates:
        m = contract_re.search(_fix_number_sign(text))
        if m:
            data["contract_date"] = m.group(1).replace(",", ".")
            number = m.group(2).replace("№", "").strip()
            data["contract_number"] = re.sub(r"(?<=\d)\s+(?=[\d/])", "", number)
            break

    # --- исходящий номер/дата (рукописные, OCR обычно не справляется) ---
    for ln in lines:
        if ln.x < mid and ln.y < page.height * 0.25:
            m = re.search(r"(\d{2}[.,]\d{2}[.,]\d{4}).*?№\s*([\d\-/]+)", _fix_number_sign(ln.text))
            if m:
                data["outgoing_date"] = m.group(1).replace(",", ".")
                data["outgoing_number"] = m.group(2)

    # --- поля бланка по подписям слева ---
    labels = [l for l in lines if l.x + l.w < mid * 1.05 and l.x > page.width * 0.1]
    anchors: list[tuple[str, float]] = []
    used: set[int] = set()
    for key, kw in ANCHORS:
        for i, lb in enumerate(labels):
            if i not in used and _matches(lb.text, kw):
                used.add(i)
                anchors.append((key, lb.y))
                break
    anchors.sort(key=lambda a: a[1])

    values = [l for l in lines if l.x >= mid]
    for idx, (key, top) in enumerate(anchors):
        if idx + 1 < len(anchors):
            bottom = anchors[idx + 1][1] - tol
        else:
            bottom = top + page.height * 0.03
        chunk = [v for v in values if top - tol <= v.y < bottom]

        if key == "request_date":
            nums = [re.sub(r"\D", "", v.text) for v in sorted(chunk, key=lambda v: v.x)]
            nums = [n for n in nums if n]
            if len(nums) >= 3:
                data[key] = f"{int(nums[0]):02d}.{int(nums[1]):02d}.{nums[2]}"
            else:
                data[key] = " ".join(nums)
            continue

        data[key] = _fix_number_sign(_join(chunk))

    # --- подпись и исполнитель внизу ---
    date_y = next((y for k, y in anchors if k == "request_date"), page.height * 0.75)
    bottom_left = [l for l in lines if l.y > date_y + page.height * 0.03 and l.x < mid]
    bottom_right = [l for l in lines if l.y > date_y + page.height * 0.03 and l.x >= mid]
    signer_pos = [l for l in bottom_left if not re.search(r"\d", l.text)]
    data["signer_position"] = _join(signer_pos)
    data["signer_name"] = _join(bottom_right)
    for l in bottom_left:
        m = re.match(r"^\s*\d*\s*([А-ЯЁ][а-яё\-]+)\s+[\d\s\-]{5,}$", l.text)
        if m:
            data["applicant_surname"] = m.group(1)

    # --- нормализация ---
    for k in ("model", "inventory_number", "serial_number"):
        data[k] = _clean_code(data[k])
    for k in ("applicant_phone", "contact_phone"):
        data[k] = _clean_phone(data[k])
    data["email"] = data["email"].replace(" ", "")

    return data


def parse_request(pdf_path: str) -> dict[str, str]:
    return parse_pages(ocr_pdf(pdf_path))


if __name__ == "__main__":
    import sys

    result = parse_request(sys.argv[1])
    for f in FIELDS:
        print(f"{f.title:45} | {result[f.key]}")

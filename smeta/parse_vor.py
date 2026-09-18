# -*- coding: utf-8 -*-
"""Парсер ведомостей объёмов работ (ВОР) и сопоставление их со сметой.

ВОР — это то, против чего смету проверяют в первую очередь: в ней объёмы, взятые
из проекта, а в смете — объёмы, за которые платят. В выгрузке заказчика ВОР идёт
отдельным файлом с колонкой «№ в ЛСР», то есть ссылкой на позиции сметы:

    № п/п | № в ЛСР | Наименование работ        | Ед. изм. | Кол-во
      1   |    1     | Разборка деревянных...   |    м2    | 153,87
      2   |   2 3    | Разборка выравнивающих.. |    м2    | 320
      3   |   5-8    | Устройство покрытий...   |    м2    | 140

Одна строка ВОР может ссылаться на несколько позиций сметы — списком или диапазоном.
Единицы при этом разные: в смете «100 м2» и объём 1,5387, в ведомости «м2» и 153,87.
"""
import re
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from .parse_xlsx import _norm, is_numbering_row, to_number

VOR_COLUMNS = ["file", "sheet", "vor_no", "excel_row", "num", "lsr_ref", "name", "unit", "qty", "note"]

# «Ведомость объёмов работ № 1.1» -> «1.1»
VOR_NO_RE = re.compile(r"ведомость\s+объ[её]м\w*\s+работ\s*№?\s*([\d.]+)", re.I)
# «02-01-01д (2026)» -> базовый номер 1 и признак дополнительной сметы
LSR_NO_RE = re.compile(r"№\s*([\dА-Яа-яA-Za-z\-]+)")
# ссылки на позиции: «5-8», «2 3», «14 15», «1,2»
REF_RANGE_RE = re.compile(r"(\d+)\s*[-–]\s*(\d+)")
REF_NUM_RE = re.compile(r"\d+")

UNIT_ALIASES = {
    "м2": "м2", "м²": "м2", "кв.м": "м2", "кв. м": "м2", "m2": "м2",
    "м3": "м3", "м³": "м3", "куб.м": "м3", "куб. м": "м3",
    "м": "м", "пм": "м", "п.м": "м", "п. м": "м", "пог.м": "м",
    "т": "т", "тн": "т", "кг": "кг", "шт": "шт", "шт.": "шт", "компл": "компл",
    "компл.": "компл", "1 шт": "шт", "место": "шт",
}


def split_unit(unit):
    """«100 м2» -> (100.0, 'м2'). «т» -> (1.0, 'т'). Не разобрали — (1.0, как есть)."""
    s = _norm(unit).replace("(", " ").replace(")", " ").strip()
    if not s:
        return 1.0, ""
    m = re.match(r"^([\d\s.,]+)\s*(.*)$", s)
    mult, rest = 1.0, s
    if m and m.group(2):
        value = to_number(m.group(1))
        if value:
            mult, rest = value, m.group(2).strip()
    rest = rest.replace(" ", "")
    return mult, UNIT_ALIASES.get(rest, rest)


FULL_CODE_RE = re.compile(r"(\d{2}-\d{2}-\d{2})\s*([дd])?", re.I)


def lsr_key(title):
    """Номер ЛСР из заголовка: «02-01-01д (2026)» -> ('02-01-01', True).

    Ключом служит полный код сметы, а не последняя цифра: в комплекте бывают и
    02-01-01, и 02-01-14, и смета «СР-1д», у которой последняя цифра тоже единица.
    """
    m = FULL_CODE_RE.search(title or "")
    if not m:
        return None, False
    return m.group(1), bool(m.group(2))


def vor_key(title):
    """Номер ВОР: «№ 02-01-14» -> ('02-01-14', False), «№ 1.1» -> ('1', True).

    Дробная часть в коротком номере означает ведомость к дополнительной смете.
    """
    m = FULL_CODE_RE.search(title or "")
    if m:
        return m.group(1), bool(m.group(2))
    m = VOR_NO_RE.search(title or "")
    if not m:
        return None, False
    raw = m.group(1)
    return (raw.split(".")[0].lstrip("0") or "0"), "." in raw


def expand_refs(text):
    """«5-8» -> [5,6,7,8]; «2 3» -> [2,3]; «14 15» -> [14,15]."""
    if text is None:
        return []
    s = str(text)
    refs, used = [], []
    for m in REF_RANGE_RE.finditer(s):
        a, b = int(m.group(1)), int(m.group(2))
        if 0 < b - a < 200:
            refs.extend(range(a, b + 1))
        used.append((m.start(), m.end()))
    for m in REF_NUM_RE.finditer(s):
        if any(start <= m.start() < end for start, end in used):
            continue
        refs.append(int(m.group()))
    return sorted(set(refs))


def find_vor_header(rows, max_scan=30):
    """Строка шапки ВОР и разметка колонок. Опознаём по колонке «№ в ЛСР»."""
    for i, row in enumerate(rows[:max_scan]):
        cells = [_norm(c) for c in row]
        joined = " | ".join(cells)
        if "наименование работ" not in joined:
            continue
        if "в лср" not in joined and "№ в лср" not in joined:
            continue
        mapping = {}
        for idx, cell in enumerate(cells):
            if not cell:
                continue
            if "в лср" in cell:
                mapping.setdefault("lsr_ref", idx)
            elif cell.startswith("№") or "п/п" in cell:
                mapping.setdefault("num", idx)
            elif "наименование" in cell:
                mapping.setdefault("name", idx)
            elif "ед" in cell and "изм" in cell:
                mapping.setdefault("unit", idx)
            elif "кол-во" in cell or "количество" in cell:
                mapping.setdefault("qty", idx)
            elif "примечание" in cell:
                mapping.setdefault("note", idx)
        start = i + 1
        if start < len(rows) and is_numbering_row(rows[start]):
            start += 1
        if {"lsr_ref", "name", "qty"} <= set(mapping):
            return start, mapping
    return None, {}


def parse_vor_file(path):
    path = Path(path)
    wb = load_workbook(path, data_only=True, read_only=True)
    records = []
    try:
        for ws in wb.worksheets:
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
            start, mapping = find_vor_header(rows)
            if start is None:
                continue
            # Номер сметы бывает и в заголовке, и строкой «к локальной смете
            # № ЛСР-02-01-01», и только в имени файла — смотрим везде
            title = " ".join(str(c) for row in rows[:20] for c in row if isinstance(c, str))
            number, is_extra = vor_key(title)
            if number is None:
                number, is_extra = vor_key(ws.title)
            if number is None:
                number, is_extra = vor_key(path.name)
            for offset, cells in enumerate(rows[start:], start=start + 1):
                def get(field):
                    idx = mapping.get(field)
                    if idx is None or idx >= len(cells):
                        return None
                    v = cells[idx]
                    return v.strip() if isinstance(v, str) else v

                name = get("name")
                qty = to_number(get("qty"))
                refs = expand_refs(get("lsr_ref"))
                if not name or qty is None or not refs:
                    continue
                records.append(dict(
                    file=path.name, sheet=ws.title,
                    vor_no=(number + (".1" if is_extra else "")) if number else None,
                    excel_row=offset, num=to_number(get("num")),
                    lsr_ref=refs, name=re.sub(r"\s+", " ", str(name)).strip(),
                    unit=get("unit"), qty=qty, note=get("note"),
                ))
    finally:
        wb.close()
    return pd.DataFrame(records, columns=VOR_COLUMNS)


def parse_vor_dir(folder, pattern="*.xls*"):
    frames = []
    for f in sorted(Path(folder).glob(pattern)):
        if f.name.startswith("~$"):
            continue
        try:
            df = parse_vor_file(f)
        except Exception:
            continue
        if len(df):
            frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=VOR_COLUMNS)


def sheet_titles(path):
    """Лист -> заголовок ЛОКАЛЬНЫЙ СМЕТНЫЙ РАСЧЕТ, чтобы вытащить номер сметы."""
    out = {}
    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True, max_row=30):
                for c in row[:8]:
                    if isinstance(c, str) and "ЛОКАЛЬНЫЙ СМЕТНЫЙ" in c.upper():
                        out[ws.title] = re.sub(r"\s+", " ", c).strip()
                        break
                if ws.title in out:
                    break
    finally:
        wb.close()
    return out


def short_key(code, extra):
    """«02-01-01» -> «1»: ведомости часто нумеруются коротко, а сметы — полным кодом."""
    if not code:
        return None
    tail = re.findall(r"\d+", code)
    if not tail:
        return None
    base = tail[-1].lstrip("0") or "0"
    return base + (".1" if extra else "")


def build_sheet_index(folder, pattern="*.xls*"):
    """Номер ЛСР -> лист сметы. Регистрируем и полный код («02-01-01»), и короткий
    номер («1»): ведомость может ссылаться на смету любым из них."""
    index = {}
    for f in sorted(Path(folder).glob(pattern)):
        if f.name.startswith("~$"):
            continue
        try:
            titles = sheet_titles(f)
        except Exception:
            continue
        for sheet, title in titles.items():
            code, extra = lsr_key(title)
            if not code:
                continue
            full = code + ("д" if extra else "")
            index[full] = (f.name, sheet, title)
            index.setdefault(code + (".1" if extra else ""), (f.name, sheet, title))
            short = short_key(code, extra)
            if short:
                index.setdefault(short, (f.name, sheet, title))
    return index

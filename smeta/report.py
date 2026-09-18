# -*- coding: utf-8 -*-
"""Протокол замечаний в xlsx.

Формат выдачи важнее движка: сметчик открывает файл и должен сразу видеть, куда идти
и сколько это стоит. Поэтому сортировка по деньгам, адрес строки исходника в каждой
записи и лист со сводкой по видам замечаний.
"""
from pathlib import Path

import pandas as pd

SEVERITY_COLOR = {"высокая": "#FFC7CE", "средняя": "#FFEB9C", "низкая": "#DDEBF7"}

RULE_TITLES = {
    "arithmetic": "Арифметика позиции",
    "duplicate": "Дублирование позиций",
    "price_above_median": "Цена выше медианы по корпусу",
    "round_volume": "Круглый объём (прикидка)",
    "overhead_out_of_range": "НР/СП вне норматива",
    "coef_on_materials": "Коэффициент на материалы",
    "no_basis": "Позиция без обоснования",
    "subtotal_mismatch": "Итог раздела не сходится",
    "vor_mismatch": "Объём не бьётся с ведомостью",
    "position_total": "«Всего по позиции» не сходится",
    "overhead_inconsistent": "Разнобой НР/СП по виду работ",
    "market_price": "Цена по конъюнктурному анализу",
}

PROTOCOL_COLUMNS = [
    ("№", 5), ("Файл", 24), ("Лист (смета)", 26), ("Строка", 8), ("Раздел", 22),
    ("Шифр", 20), ("Наименование", 44), ("Вид замечания", 26), ("Критичность", 12),
    ("Замечание", 52), ("Основание для проверки", 42), ("Сумма влияния, руб.", 18),
]


def build(findings, parse_stats, out_path):
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    proto = pd.DataFrame({
        "№": range(1, len(findings) + 1),
        "Файл": findings.file,
        "Лист (смета)": findings.sheet,
        "Строка": findings.excel_row,
        "Раздел": findings.section,
        "Шифр": findings.code,
        "Наименование": findings.name,
        "Вид замечания": findings.rule.map(RULE_TITLES).fillna(findings.rule),
        "Критичность": findings.severity,
        "Замечание": findings.message,
        "Основание для проверки": findings.detail,
        "Сумма влияния, руб.": findings.impact_rub,
    }) if len(findings) else pd.DataFrame(columns=[c for c, _ in PROTOCOL_COLUMNS])

    summary = (findings.groupby("rule")
               .agg(Замечаний=("rule", "size"), Сумма=("impact_rub", "sum"))
               .reset_index()
               .assign(**{"Вид замечания": lambda d: d.rule.map(RULE_TITLES).fillna(d.rule)})
               .loc[:, ["Вид замечания", "Замечаний", "Сумма"]]
               .sort_values("Сумма", ascending=False)) if len(findings) else pd.DataFrame()

    by_file = (findings.groupby("file")
               .agg(Замечаний=("rule", "size"), Сумма=("impact_rub", "sum"))
               .reset_index().rename(columns={"file": "Файл"})
               .sort_values("Сумма", ascending=False)) if len(findings) else pd.DataFrame()

    stats = parse_stats.rename(columns={
        "file": "Файл", "sheet": "Лист (смета)", "positions": "Позиций",
        "overheads": "Строк НР/СП", "sections": "Разделов", "unknown": "Не распознано",
        "no_price": "Без цены", "clean": "Разобрано без ручной правки", "error": "Примечание",
    })

    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        book = writer.book
        proto.to_excel(writer, sheet_name="Протокол замечаний", index=False, startrow=1)
        ws = writer.sheets["Протокол замечаний"]

        title = book.add_format({"bold": True, "font_size": 13})
        head = book.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "white",
                                "text_wrap": True, "valign": "vcenter", "border": 1})
        wrap = book.add_format({"text_wrap": True, "valign": "top", "border": 1})
        money = book.add_format({"num_format": "# ##0,00", "valign": "top", "border": 1})
        ws.write(0, 0, "Протокол замечаний по сметной документации", title)

        for i, (name, width) in enumerate(PROTOCOL_COLUMNS):
            ws.set_column(i, i, width, money if "руб" in name else wrap)
            ws.write(1, i, name, head)
        ws.freeze_panes(2, 0)
        if len(proto):
            ws.autofilter(1, 0, 1 + len(proto), len(PROTOCOL_COLUMNS) - 1)
            for sev, color in SEVERITY_COLOR.items():
                fmt = book.add_format({"bg_color": color})
                ws.conditional_format(2, 0, 1 + len(proto), len(PROTOCOL_COLUMNS) - 1, {
                    "type": "formula",
                    "criteria": '=$I3="{}"'.format(sev),
                    "format": fmt,
                })
            total_fmt = book.add_format({"bold": True, "num_format": "# ##0,00", "top": 2})
            last_col = len(PROTOCOL_COLUMNS) - 1
            ws.write(len(proto) + 2, last_col - 1, "ИТОГО влияние", book.add_format({"bold": True, "top": 2}))
            ws.write_number(len(proto) + 2, last_col, float(findings.impact_rub.sum()), total_fmt)

        if len(summary):
            summary.to_excel(writer, sheet_name="Сводка", index=False, startrow=1)
            s = writer.sheets["Сводка"]
            s.write(0, 0, "Замечания по видам", title)
            s.set_column(0, 0, 34, wrap)
            s.set_column(1, 1, 12)
            s.set_column(2, 2, 20, money)
            by_file.to_excel(writer, sheet_name="Сводка", index=False, startrow=len(summary) + 5)
            s.write(len(summary) + 4, 0, "Замечания по файлам", title)

        stats.to_excel(writer, sheet_name="Качество разбора", index=False, startrow=1)
        q = writer.sheets["Качество разбора"]
        q.write(0, 0, "Как разобрались исходники", title)
        q.set_column(0, 0, 30, wrap)
        q.set_column(1, 1, 28, wrap)
        q.set_column(2, 8, 16)

    return out_path

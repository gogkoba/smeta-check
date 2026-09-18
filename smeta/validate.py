# -*- coding: utf-8 -*-
"""Движок проверок — ядро ценности пайплайна.

Каждое замечание имеет адрес (файл + лист + строка исходника), человекочитаемую
формулировку и, где это считается, оценку завышения в рублях: протокол сортируется
по деньгам, а не по алфавиту.

Важное про арифметику. В выгрузке ГРАНД-Сметы коэффициенты применяются к каждой
составляющей отдельно — ОЗП=0,8, ЭМ=0,8, МАТ=0 — поэтому проверка «объём × общая
цена × коэффициент» даёт ложные срабатывания пачками. Считаем по составляющим,
затем сверяем их сумму с итогом по расценке.
"""
import re

import pandas as pd

from .norms import TOLERANCE_PCT, norm_for
from .parse_xlsx import CODE_RE, ORDER_REF_RE
from .schema import KIND_OVERHEAD, KIND_POSITION, KIND_PROFIT

SEVERITY_ORDER = {"высокая": 0, "средняя": 1, "низкая": 2}

# шифр текущей цены из конъюнктурного анализа: ТЦ_<регион>_<ИНН>_<дата>
MARKET_PRICE_RE = re.compile(r"^(ТЦ|ПРАЙС|КП)[_\s-]", re.I)


def money(x):
    """1234567.8 -> «1 234 567,80» — как принято в сметах."""
    return "{:,.2f}".format(float(x or 0)).replace(",", " ").replace(".", ",").replace(" ", " ")


def _finding(row, rule, severity, message, impact=0.0, detail=""):
    return dict(
        file=row.get("file"),
        sheet=row.get("sheet"),
        excel_row=int(row["excel_row"]) if pd.notna(row.get("excel_row")) else None,
        section=row.get("section"),
        code=row.get("code"),
        name=row.get("name"),
        rule=rule,
        severity=severity,
        message=message,
        detail=detail,
        impact_rub=round(float(impact or 0.0), 2),
    )


def _positions(df):
    return df[df.kind == KIND_POSITION]


def _val(x):
    """None вместо NaN. Без этого «NaN or 0» даёт NaN, и любое сравнение с ним ложно —
    правило начинает срабатывать на каждой позиции подряд."""
    if x is None:
        return None
    try:
        return None if pd.isna(x) else x
    except (TypeError, ValueError):
        return x


def _close(actual, expected, rel_tol=0.005, abs_tol=1.0):
    return abs(actual - expected) <= max(abs_tol, abs(expected) * rel_tol)


# --- арифметика -------------------------------------------------------------

def rule_arithmetic(df, rel_tol=0.005, abs_tol=1.0):
    """Объём × цена × коэффициент по каждой составляющей, затем сумма против итога."""
    out = []
    for _, r in _positions(df).iterrows():
        if pd.isna(r.qty):
            continue
        comps = r.get("components") or []

        if comps:
            bad = []
            for c in comps:
                price, total, coef = c.get("price"), c.get("total"), c.get("coef")
                if price is None or total is None:
                    continue
                expected = r.qty * price * (1.0 if coef is None else coef)
                if not _close(total, expected, rel_tol, abs_tol):
                    bad.append((c.get("name") or c.get("kind"), total, expected))
            if bad:
                worst = max(bad, key=lambda b: abs(b[1] - b[2]))
                out.append(_finding(
                    r, "arithmetic", "высокая",
                    "Составляющая «{}» не сходится: в смете {}, по расчёту {}".format(
                        worst[0], money(worst[1]), money(worst[2])),
                    impact=max(sum(b[1] - b[2] for b in bad), 0),
                    detail="объём {:g} × цена × коэффициент; расхождений в позиции: {}".format(
                        r.qty, len(bad)),
                ))
                continue

            if r.total is not None:
                comp_sum = sum(c["total"] for c in comps if c.get("total") is not None)
                if comp_sum and not _close(r.total, comp_sum, rel_tol, abs_tol):
                    out.append(_finding(
                        r, "arithmetic", "высокая",
                        "Итого по расценке {} не равно сумме составляющих {}".format(
                            money(r.total), money(comp_sum)),
                        impact=max(r.total - comp_sum, 0),
                        detail="ОЗП + ЭМ + материалы",
                    ))
            continue

        # простая смета: одна строка на позицию
        price, total = _val(r.price), _val(r.total)
        if price is None or total is None:
            continue
        expected = r.qty * price * (r.coef or 1.0)
        # Для позиций из конъюнктурного анализа графа «на единицу» дана в текущем
        # уровне цен, а итог — в базисном: так прямо написано в шапке колонки
        # («в текущем уровне цен (гр. 8) для ресурсов, отсутствующих в СНБ»).
        # Без деления на индекс такие позиции все до одной выглядят ошибкой.
        index = _val(r.get("index"))
        if index and index > 1 and not CODE_RE.match(str(r.code or "")):
            expected = expected / index
        if not _close(total, expected, rel_tol, abs_tol):
            out.append(_finding(
                r, "arithmetic", "высокая",
                "Сумма не сходится с расчётом: в смете {}, по формуле {}".format(
                    money(total), money(expected)),
                impact=max(total - expected, 0),
                detail="объём {:g} × цена {} × коэф. {:g}{}".format(
                    r.qty, money(price), r.coef or 1.0,
                    ", с пересчётом по индексу {:g}".format(index) if index and index > 1
                    and not CODE_RE.match(str(r.code or "")) else ""),
            ))
    return out


def rule_position_total(df, rel_tol=0.005, abs_tol=1.0):
    """«Всего по позиции» должно быть равно итогу по расценке плюс НР и СП."""
    out = []
    for _, r in _positions(df).iterrows():
        whole = _val(r.get("total_with_overhead"))
        total = _val(r.total)
        if whole is None or total is None:
            continue
        nr, sp = _val(r.get("nr_sum")) or 0, _val(r.get("sp_sum")) or 0
        expected = total + nr + sp
        if not _close(whole, expected, rel_tol, abs_tol):
            out.append(_finding(
                r, "position_total", "высокая",
                "«Всего по позиции» {} не равно расценке с НР и СП {}".format(
                    money(whole), money(expected)),
                impact=max(whole - expected, 0),
                detail="итого по расценке {} + НР {} + СП {}".format(
                    money(total), money(nr), money(sp)),
            ))
    return out


# --- дубли и объёмы ---------------------------------------------------------

def rule_duplicates(df):
    """Одна и та же работа посчитана дважды внутри одной сметы (листа)."""
    out = []
    pos = _positions(df).copy()
    # Материалы и оборудование законно повторяются в разных системах и узлах;
    # подозрителен повтор работы по одной и той же расценке
    pos = pos[pos.code.fillna("").str.match(r"(?i)^(ФЕР|ТЕР|ГЭСН)")]
    if pos.empty:
        return out
    pos["_key"] = (pos.code.fillna("").str.lower().str.strip() + "|" +
                   pos.name.fillna("").str.lower().str.strip() + "|" +
                   pos.qty.round(4).astype(str))
    for (file, sheet, key), grp in pos.groupby(["file", "sheet", "_key"]):
        if len(grp) < 2 or not key.strip("|"):
            continue
        rows = grp.sort_values("excel_row")
        first = rows.iloc[0]
        for _, r in rows.iloc[1:].iterrows():
            same_section = r.section == first.section
            first_block, this_block = _val(first.subsection), _val(r.subsection)
            same_block = same_section and first_block == this_block
            if same_block:
                where = "тот же блок работ" if first_block else "тот же раздел"
            elif same_section:
                where = ("другой блок работ: «{}»".format(first_block) if first_block
                         else "начало раздела «{}»".format(first.section))
            else:
                where = "раздел «{}»".format(first.section)
            out.append(_finding(
                r, "duplicate", "высокая" if same_block else "средняя",
                "Позиция повторяет строку {} ({})".format(int(first.excel_row), where),
                impact=r.total or 0.0,
                detail="совпадают шифр, наименование и объём — проверить, не задвоена ли работа",
            ))
    return out


def rule_round_volume(df, min_value=100):
    """Круглые объёмы — признак прикидки, а не подсчёта по документации."""
    out = []
    pos = _positions(df)
    for _, r in pos[pos.qty.notna()].iterrows():
        q = float(r.qty)
        if q >= min_value and q == int(q) and int(q) % 50 == 0:
            out.append(_finding(
                r, "round_volume", "средняя",
                "Подозрительно круглый объём: {:g} {}".format(q, r.unit or ""),
                impact=0.0,
                detail="сверить с ведомостью объёмов работ",
            ))
    return out


# --- цены -------------------------------------------------------------------

def rule_price_above_median(df, threshold=0.25, min_samples=3):
    """Цена за единицу выше медианы по этому шифру в корпусе.

    Сравнивать можно только позиции с нормативным шифром. Шифр конъюнктурной цены
    (ТЦ_20.3.03.07_77_7722699172_05.07.2022_01) идентифицирует источник — регион,
    ИНН поставщика и дату, — а не товар: под одним таким шифром лежат пять разных
    светильников, и медиана по ним ничего не значит.
    """
    out = []
    pos = _positions(df)
    pos = pos[pos.price.notna() & pos.code.notna()].copy()
    if pos.empty:
        return out
    pos = pos[pos.code.str.match(CODE_RE)]
    if pos.empty:
        return out
    # Сравнивать можно только цены одного уровня и одной единицы измерения:
    # базисная цена из одной сметы и текущая из другой отличаются в разы законно
    pos["_code"] = (pos.code.str.upper().str.replace(r"\s+", "", regex=True) + "|" +
                    pos.price_level.fillna("?") + "|" +
                    pos.unit.fillna("").str.lower().str.replace(r"\s+", "", regex=True))
    medians = pos.groupby("_code").price.agg(["median", "count"])
    for _, r in pos.iterrows():
        stat = medians.loc[r._code]
        if stat["count"] < min_samples or not stat["median"]:
            continue
        dev = (r.price - stat["median"]) / stat["median"]
        if dev > threshold:
            overpay = (r.price - stat["median"]) * (r.qty or 0)
            out.append(_finding(
                r, "price_above_median", "высокая" if dev > 0.4 else "средняя",
                "Цена за единицу выше медианы по корпусу на {:.0f}%".format(dev * 100),
                impact=overpay,
                detail="в смете {}, медиана по {} позициям корпуса {} ({} уровень цен, {}){}".format(
                    money(r.price), int(stat["count"]), money(stat["median"]),
                    r.price_level or "?", r.unit or "",
                    "; применены коэффициенты: " + str(r.coef_text)[:80] if _val(r.coef_text) else ""),
            ))
    return out


def rule_no_basis(df):
    """Позиции без нормативной расценки: прайс, счёт, конъюнктурный анализ."""
    out = []
    for _, r in _positions(df).iterrows():
        code = str(r.code or "").strip()
        if code and CODE_RE.match(code):
            continue
        if MARKET_PRICE_RE.match(code) or ORDER_REF_RE.match(code):
            out.append(_finding(
                r, "market_price", "средняя",
                "Цена по конъюнктурному анализу, а не по нормативной базе",
                impact=0.0,
                detail="проверить обоснование: не менее трёх предложений поставщиков (421/пр), "
                       "актуальность и сопоставимость условий поставки",
            ))
            continue
        out.append(_finding(
            r, "no_basis", "высокая",
            "Позиция без обоснования расценкой: «{}»".format(code or "шифр не указан"),
            impact=(r.total or 0.0) * 0.15,
            detail="нет ссылки ни на нормативную базу, ни на конъюнктурный анализ",
        ))
    return out


# --- накладные расходы и сметная прибыль ------------------------------------

def rule_overhead_out_of_range(df):
    """НР и СП против норматива — только там, где вид работ есть в справочнике.

    Справочник в norms.py неполный, поэтому молчим, где не знаем, вместо того чтобы
    выдавать замечание на каждую вторую позицию. Разнобой внутри одного вида работ
    ловит соседнее правило, которому справочник не нужен.
    """
    out = []
    rows = df[df.kind.isin([KIND_OVERHEAD, KIND_PROFIT])]
    for _, r in rows.iterrows():
        if pd.isna(r.pct) or not r.pct:
            continue
        # вид работ из строки НР точнее раздела, но в простых сметах его нет
        key = r.get("work_type")
        norms = norm_for(key)
        if norms is None:
            key = r.get("section")
            norms = norm_for(key)
        if norms is None:
            continue
        nr_norm, sp_norm = norms
        norm = nr_norm if r.kind == KIND_OVERHEAD else sp_norm
        label = "Накладные расходы" if r.kind == KIND_OVERHEAD else "Сметная прибыль"
        if r.pct > norm + TOLERANCE_PCT:
            excess = (r.pct - norm) / r.pct * (r.total or 0.0)
            out.append(_finding(
                r, "overhead_out_of_range", "средняя",
                "{} {:g}% при нормативе {:g}% для «{}»".format(label, r.pct, norm, key),
                impact=excess,
                detail="норматив взят из встроенного справочника, он неполный — "
                       "обязательно сверить с приказом Минстроя {}".format(
                           "812/пр" if r.kind == KIND_OVERHEAD else "774/пр"),
            ))
    return out


def rule_overhead_inconsistent(df, min_samples=3):
    """Один вид работ — один процент НР и СП. Разнобой означает ошибку в смете.

    Норматив для этого знать не нужно: сравниваем смету саму с собой. Правило не
    зависит от нормативной базы и переживает её изменения.
    """
    out = []
    rows = df[df.kind.isin([KIND_OVERHEAD, KIND_PROFIT])]
    rows = rows[rows.pct.notna() & rows.work_type.notna() & (rows.pct > 0)]
    for (kind, work_type), grp in rows.groupby(["kind", "work_type"]):
        if len(grp) < min_samples:
            continue
        counts = grp.pct.value_counts()
        if len(counts) < 2:
            continue
        common = counts.index[0]
        label = "Накладные расходы" if kind == KIND_OVERHEAD else "Сметная прибыль"
        for _, r in grp[grp.pct != common].iterrows():
            excess = max((r.pct - common) / r.pct * (r.total or 0.0), 0) if r.pct else 0
            out.append(_finding(
                r, "overhead_inconsistent", "средняя",
                "{} {:g}% при {:g}% на остальных позициях того же вида работ".format(
                    label, r.pct, common),
                impact=excess,
                detail="вид работ «{}»: {:g}% встречается {} раз, {:g}% — {} раз".format(
                    work_type, common, int(counts.iloc[0]), r.pct, int(counts[r.pct])),
            ))
    return out


def rule_coef_on_materials(df, mat_share=0.85):
    """Повышающий коэффициент применён к материальной позиции."""
    out = []
    for _, r in _positions(df).iterrows():
        coef = r.coef or 1.0
        if coef <= 1.0:
            continue
        share = (r.mat or 0.0) / r.total if r.total else 0.0
        is_material_code = bool(re.match(r"^(ФССЦ|ТССЦ|ФСБЦ)", str(r.code or ""), re.I))
        if share >= mat_share or is_material_code:
            out.append(_finding(
                r, "coef_on_materials", "высокая",
                "Коэффициент {:g} применён к материальной позиции".format(coef),
                impact=(r.total or 0.0) * (coef - 1) / coef,
                detail=r.coef_text or "материалы {:.0f}% стоимости".format(share * 100),
            ))
    return out


def rule_subtotal_mismatch(df, rel_tol=0.01):
    """Итог раздела не сходится с суммой позиций этого раздела."""
    out = []
    scoped = df[df.section.notna()]
    for (file, sheet, section), grp in scoped.groupby(["file", "sheet", "section"]):
        body = grp[grp.kind == KIND_POSITION]
        sub = grp[grp.kind == "subtotal"]
        if body.empty or sub.empty:
            continue
        if body.total_with_overhead.notna().any():
            calc = body.total_with_overhead.fillna(body.total).fillna(0).sum()
        else:
            extra = grp[grp.kind.isin([KIND_OVERHEAD, KIND_PROFIT])].total.fillna(0).sum()
            calc = body.total.fillna(0).sum() + extra
        declared = sub.total.fillna(0).iloc[-1]
        if declared and calc and not _close(declared, calc, rel_tol, 1.0):
            row = sub.iloc[-1]
            out.append(_finding(
                row, "subtotal_mismatch", "средняя",
                "Итог по разделу: заявлено {}, сумма позиций {}".format(money(declared), money(calc)),
                impact=max(declared - calc, 0),
                detail="раздел «{}», позиций {}".format(section, len(body)),
            ))
    return out


RULES = [
    rule_arithmetic,
    rule_position_total,
    rule_duplicates,
    rule_price_above_median,
    rule_round_volume,
    rule_overhead_out_of_range,
    rule_overhead_inconsistent,
    rule_coef_on_materials,
    rule_no_basis,
    rule_subtotal_mismatch,
]


def validate(df, vor=None, sheet_index=None):
    """Все правила по таблице. Возвращает протокол, отсортированный по деньгам.

    vor и sheet_index не обязательны: без ведомостей объёмов работает всё
    остальное, просто одной проверкой меньше.
    """
    findings = []
    for rule in RULES:
        findings.extend(rule(df))
    if vor is not None and sheet_index:
        findings.extend(rule_vor_mismatch(df, vor, sheet_index))
    if not findings:
        return pd.DataFrame(columns=["file", "sheet", "excel_row", "section", "code", "name",
                                     "rule", "severity", "message", "detail", "impact_rub"])
    out = pd.DataFrame(findings)
    out["_sev"] = out.severity.map(SEVERITY_ORDER).fillna(3)
    out = out.sort_values(["impact_rub", "_sev"], ascending=[False, True]).drop(columns="_sev")
    return out.reset_index(drop=True)


# --- сверка с ведомостью объёмов работ --------------------------------------

def rule_vor_mismatch(df, vor, sheet_index, rel_tol=0.01, min_agreement=0.7):
    """Объёмы сметы против ведомости объёмов работ.

    Ведомость ссылается на позиции сметы колонкой «№ в ЛСР», иногда сразу на
    несколько («5-8»): тогда сравнивается суммарный объём. Единицы приводятся к
    общему виду — в смете «100 м2» и 1,5387, в ведомости «м2» и 153,87.

    Ссылке нельзя верить вслепую. В части комплектов ведомость нумерует позиции
    по-своему, и «№ в ЛСР» указывает на чужую строку: наименование при этом может
    совпасть, потому что одна и та же работа повторяется для разных помещений.
    Поэтому связка сначала калибруется целиком: если по паре «смета — ведомость»
    сходится меньше min_agreement строк, нумерация не совпадает и сверка не
    проводится вовсе. Лучше сказать «не могу проверить», чем выдать сотню
    выдуманных расхождений.
    """
    from rapidfuzz import fuzz

    from .parse_vor import split_unit

    out, skipped = [], []
    if vor is None or not len(vor) or not sheet_index:
        return out

    positions = _positions(df)
    for vor_no, block in vor.groupby("vor_no"):
        target = sheet_index.get(vor_no)
        if target is None:
            continue
        file_name, sheet_name, _title = target
        smeta_rows = positions[(positions.file == file_name) & (positions.sheet == sheet_name)]
        if smeta_rows.empty:
            continue
        by_num = {int(n): row for n, row in zip(smeta_rows.num, smeta_rows.to_dict("records"))
                  if pd.notna(n)}

        comparable = []
        for _, v in block.iterrows():
            refs = [n for n in (v.lsr_ref or []) if n in by_num]
            if not refs:
                continue
            similarity = max(fuzz.token_set_ratio(str(v["name"]).lower(),
                                                  str(by_num[n].get("name") or "").lower())
                             for n in refs)
            if similarity < 60:
                continue

            vor_mult, vor_unit = split_unit(v.unit)
            volumes, units = [], set()
            for n in refs:
                row = by_num[n]
                mult, unit = split_unit(row.get("unit"))
                qty = _val(row.get("qty"))
                if qty is None:
                    continue
                volumes.append(qty * mult)
                units.add(unit)
            if not volumes or len(units) != 1 or units.pop() != vor_unit:
                continue

            declared = v.qty * vor_mult
            # Одна строка ведомости часто ссылается сразу на работу и материал к ней,
            # и объём у них общий, а не складывается
            candidates = volumes + [sum(volumes)]
            agrees = any(_close(c, declared, rel_tol, 0.01) for c in candidates)
            comparable.append((v, refs, declared, candidates, agrees, similarity))

        if not comparable:
            continue
        agreement = sum(1 for c in comparable if c[4]) / len(comparable)
        if agreement < min_agreement:
            skipped.append((vor_no, len(comparable), agreement))
            continue

        for v, refs, declared, candidates, agrees, similarity in comparable:
            if agrees:
                continue
            closest = min(candidates, key=lambda c: abs(c - declared))
            diff = closest - declared
            first = by_num[refs[0]]
            _mult, vor_unit = split_unit(v.unit)
            out.append(_finding(
                first, "vor_mismatch",
                "высокая" if abs(diff) / max(declared, 1) > 0.05 else "средняя",
                "Объём в смете {:g} {}, в ведомости {:g} {} (расхождение {:+.1%})".format(
                    closest, vor_unit, declared, vor_unit, diff / declared if declared else 0),
                impact=abs(diff) * (_val(first.get("price")) or 0) / max(
                    split_unit(first.get("unit"))[0], 1),
                detail="{} строка {}, позиции сметы {}; совпадение наименований {:.0f}%".format(
                    v.sheet, v.excel_row, ", ".join(str(n) for n in refs), similarity),
            ))

    for vor_no, count, agreement in skipped:
        print("  ведомость {}: сверка не проводилась — по {} строкам сошлось только "
              "{:.0%}, нумерация «№ в ЛСР» не совпадает со сметой".format(
                  vor_no, count, agreement))
    return out

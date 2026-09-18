# -*- coding: utf-8 -*-
"""Выгрузка результатов прогона в JSON для веб-демо.

    python tools/export_web.py --input data/corpus --out data/reports/web_data.json

Страница демо статическая: весь расчёт делает пайплайн, сюда попадают только цифры.
"""
import argparse
import json
import re
import sys
import time
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from smeta.parse_vor import build_sheet_index, parse_vor_dir  # noqa: E402
from smeta.parse_xlsx import parse_dir  # noqa: E402
from smeta.report import RULE_TITLES  # noqa: E402
from smeta.validate import validate  # noqa: E402

RULE_EXPLAIN = {
    "arithmetic": "объём × цена × коэффициент по каждой составляющей, затем сумма против итога",
    "position_total": "«Всего по позиции» не равно расценке плюс НР и СП",
    "vor_mismatch": "объём в смете не совпал с ведомостью объёмов работ",
    "duplicate": "одна и та же работа посчитана дважды",
    "price_above_median": "цена за единицу выше медианы по корпусу",
    "overhead_out_of_range": "НР/СП вне норматива по виду работ",
    "overhead_inconsistent": "один вид работ — разные проценты НР/СП внутри сметы",
    "coef_on_materials": "повышающий коэффициент применён к материальной позиции",
    "market_price": "цена по конъюнктурному анализу — нужны три КП по 421/пр",
    "no_basis": "позиция без обоснования",
    "round_volume": "объём 500 / 1000 — признак прикидки, а не подсчёта",
    "subtotal_mismatch": "итог раздела не равен сумме позиций",
}


def _clean(obj):
    """NaN и NaT наружу не выпускаем: JSON.parse на них падает."""
    if isinstance(obj, dict):
        return {k: _clean(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_clean(v) for v in obj]
    if isinstance(obj, float) and obj != obj:
        return None
    if obj is pd.NaT or (obj is not None and not isinstance(obj, (str, bool, int, float)) and pd.isna(obj)):
        return None
    return obj


LOGISTICS_RE = re.compile(r"перевозк|погрузк|разгрузк|доставк|транспортиров", re.I)

# Строки НР/СП, где в графе «вид работ» стоит не вид работ, а название самой строки.
WORKTYPE_JUNK_RE = re.compile(r"^(всего|итого)?\s*(накладные расходы|сметная прибыль)\s*$", re.I)


def attach_work_type(df):
    """Вид работ печатается в строке НР/СП, а нужен он на позиции — переносим по parent_row.

    Норматив НР/СП задан по видам работ (812/пр, 774/пр), поэтому в выгрузке вид работ
    стоит на строке накладных, а не на самой позиции. Для аналитики он нужен на позиции.
    """
    ov = df[df.kind.isin(["overhead", "profit"])]
    link = {}
    for r in ov.itertuples():
        wt = (r.work_type or "").strip()
        if not wt or WORKTYPE_JUNK_RE.match(wt) or pd.isna(r.parent_row):
            continue
        link.setdefault((r.file, r.sheet, r.parent_row), wt)
    return [link.get((r.file, r.sheet, r.excel_row)) for r in df.itertuples()]


def smeta_analytics(pos, fnd, stat_row):
    """Аналитика по одной смете: деньги, виды работ, структура затрат, замечания."""
    rub = float(pos.total_current.sum(skipna=True))
    priced = pos[pos.total_current.notna()]

    by_work = (pos[pos.work_type.notna()]
               .groupby("work_type")
               .agg(n=("work_type", "size"), rub=("total_current", "sum"))
               .sort_values("rub", ascending=False).head(7))

    log = pos[pos.name.fillna("").str.contains(LOGISTICS_RE)]
    levels = pos.price_level.value_counts().to_dict()
    method = ("ресурсно-индексный" if levels.get("текущий", 0) >= levels.get("базисный", 0)
              else "базисно-индексный")

    direct = pos[pos.ozp.notna() | pos.emm.notna() | pos.mat.notna()]

    return {
        "positions": int(len(pos)),
        "rub": rub,
        "priced_positions": int(len(priced)),
        "sections": int(pos.section.nunique()),
        "resources": int(pos.resources.sum()),
        "method": method,
        "levels": {k: int(v) for k, v in levels.items()},
        "unknown": int(stat_row.unknown) if stat_row is not None else 0,
        "clean": bool(stat_row.clean) if stat_row is not None else False,
        "by_work": [
            {"work": w, "n": int(r.n), "rub": float(r.rub)}
            for w, r in by_work.iterrows()
        ],
        "work_attributed": int(pos.work_type.notna().sum()),
        "logistics": {
            "n": int(len(log)),
            "rub": float(log.total_current.sum(skipna=True)),
        },
        "direct": {
            "n": int(len(direct)),
            "ozp": float(pos.ozp.sum(skipna=True)),
            "emm": float(pos.emm.sum(skipna=True)),
            "mat": float(pos.mat.sum(skipna=True)),
            "nr": float(pos.nr_sum.sum(skipna=True)),
            "sp": float(pos.sp_sum.sum(skipna=True)),
        },
        "findings": {
            "n": int(len(fnd)),
            "rub": float(fnd.impact_rub.sum()) if len(fnd) else 0.0,
            "high": int((fnd.severity == "высокая").sum()) if len(fnd) else 0,
            "by_rule": [
                {"rule": k, "title": RULE_TITLES.get(k, k), "n": int(v)}
                for k, v in (fnd.rule.value_counts().items() if len(fnd) else [])
            ],
            "top": [
                {
                    "row": int(r.excel_row) if pd.notna(r.excel_row) else None,
                    "code": r.code, "name": r.name,
                    "title": RULE_TITLES.get(r.rule, r.rule),
                    "severity": r.severity, "message": r.message,
                    "impact": float(r.impact_rub),
                }
                for r in fnd.head(4).itertuples()
            ],
        },
    }


def main(argv=None):
    ap = argparse.ArgumentParser(description="JSON для веб-демо")
    ap.add_argument("--input", default="data/corpus")
    ap.add_argument("--out", default="data/reports/web_data.json")
    args = ap.parse_args(argv)

    t0 = time.time()
    df, stats = parse_dir(args.input)
    vor = parse_vor_dir(args.input)
    sheet_index = build_sheet_index(args.input)
    findings = validate(df, vor, sheet_index)
    elapsed = time.time() - t0

    df["work_type"] = attach_work_type(df)
    smety = stats[stats.sheet != "—"]
    positions = int((df.kind == "position").sum())

    all_pos = df[df.kind == "position"]
    per_smeta = []
    for row in smety.itertuples():
        pos = all_pos[(all_pos.file == row.file) & (all_pos.sheet == row.sheet)]
        if not len(pos):
            continue
        fnd = findings[(findings.file == row.file) & (findings.sheet == row.sheet)]
        block = smeta_analytics(pos, fnd, row)
        block["file"] = row.file
        block["sheet"] = row.sheet
        per_smeta.append(block)
    per_smeta.sort(key=lambda b: b["rub"], reverse=True)

    by_rule = (findings.groupby("rule")
               .agg(count=("rule", "size"), total=("impact_rub", "sum"))
               .reset_index().sort_values("total", ascending=False))

    by_file = (findings.groupby("file")
               .agg(count=("rule", "size"), total=("impact_rub", "sum"))
               .reset_index().sort_values("total", ascending=False))

    by_sev = findings.severity.value_counts().to_dict()

    top = findings.head(20)

    data = {
        "totals": {
            "smety_parsed": int((smety.positions > 0).sum()),
            "smety_total": int(len(smety)),
            "smety_clean": int(smety.clean.sum()),
            "positions": positions,
            "overhead_rows": int(df.kind.isin(["overhead", "profit"]).sum()),
            "unknown_rows": int(smety.unknown.sum()),
            "files": int(len(stats)),
            "vor_rows": int(len(vor)) if vor is not None else 0,
            "vor_matched": int(sum(1 for n in vor.vor_no if n in (sheet_index or {}))) if vor is not None and len(vor) else 0,
            "findings": int(len(findings)),
            "impact": float(findings.impact_rub.sum()),
            "seconds": round(elapsed, 1),
            "rules": len(RULE_TITLES),
        },
        "severity": {k: int(v) for k, v in by_sev.items()},
        "by_rule": [
            {
                "rule": r.rule,
                "title": RULE_TITLES.get(r.rule, r.rule),
                "explain": RULE_EXPLAIN.get(r.rule, ""),
                "count": int(r.count),
                "total": float(r.total),
            }
            for r in by_rule.itertuples()
        ],
        "by_file": [
            {"file": r.file, "count": int(r.count), "total": float(r.total)}
            for r in by_file.itertuples()
        ],
        "top": [
            {
                "file": r.file,
                "sheet": r.sheet,
                "row": int(r.excel_row) if pd.notna(r.excel_row) else None,
                "section": r.section,
                "code": r.code,
                "name": r.name,
                "rule": r.rule,
                "title": RULE_TITLES.get(r.rule, r.rule),
                "severity": r.severity,
                "message": r.message,
                "detail": r.detail,
                "impact": float(r.impact_rub),
            }
            for r in top.itertuples()
        ],
        "corpus": {
            "rub": float(all_pos.total_current.sum(skipna=True)),
            "priced_positions": int(all_pos.total_current.notna().sum()),
            "resources": int(all_pos.resources.sum()),
            "levels": {k: int(v) for k, v in all_pos.price_level.value_counts().items()},
            "logistics_rub": float(
                all_pos[all_pos.name.fillna("").str.contains(LOGISTICS_RE)]
                .total_current.sum(skipna=True)),
        },
        "smety": per_smeta,
        "parse_quality": [
            {
                "file": r.file,
                "sheet": r.sheet,
                "positions": int(r.positions),
                "unknown": int(r.unknown),
                "clean": bool(r.clean),
            }
            for r in stats.itertuples() if r.sheet != "—"
        ],
    }

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    # NaN — валидный литерал для питоновского json, но не для JSON.parse в браузере.
    out.write_text(json.dumps(_clean(data), ensure_ascii=False, indent=1), encoding="utf-8")
    print("Записано: {} ({} замечаний, {} позиций)".format(
        out, data["totals"]["findings"], data["totals"]["positions"]))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

# -*- coding: utf-8 -*-
"""Точка входа: папка со сметами -> протокол замечаний.

    python -m smeta.cli --input data/raw --out data/reports/protokol.xlsx

Ведомости объёмов работ, если они лежат в той же папке, подхватываются сами:
их видно по колонке «№ в ЛСР», и по ней объёмы сметы сверяются с ведомостью.
"""
import argparse
import sys
import time
from pathlib import Path

from .parse_vor import build_sheet_index, parse_vor_dir
from .parse_xlsx import parse_dir
from .report import build
from .validate import money, validate


def main(argv=None):
    ap = argparse.ArgumentParser(description="Проверка сметной документации")
    ap.add_argument("--input", default="data/corpus", help="папка со сметами (xlsx)")
    ap.add_argument("--out", default="data/reports/protokol.xlsx", help="куда сложить протокол")
    ap.add_argument("--interim", default="data/interim/normalized.csv",
                    help="нормализованная таблица позиций (для дашборда и разбора)")
    ap.add_argument("--no-vor", action="store_true", help="не сверять с ведомостями объёмов")
    args = ap.parse_args(argv)

    if sys.stdout.encoding and sys.stdout.encoding.lower() not in ("utf-8", "utf8"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    t0 = time.time()
    df, stats = parse_dir(args.input)
    if df.empty:
        print("В папке {} не нашлось смет, которые удалось разобрать".format(args.input))
        return 1

    vor, sheet_index = None, None
    if not args.no_vor:
        vor = parse_vor_dir(args.input)
        sheet_index = build_sheet_index(args.input)
        vor_files = set(vor.file) if len(vor) else set()
        stats.loc[stats.file.isin(vor_files) & (stats.sheet == "—"), "error"] = "ведомость объёмов работ"

    Path(args.interim).parent.mkdir(parents=True, exist_ok=True)
    df.drop(columns=["components"], errors="ignore").to_csv(
        args.interim, index=False, encoding="utf-8-sig")

    findings = validate(df, vor, sheet_index)
    out = build(findings, stats, args.out)
    elapsed = time.time() - t0

    smety = stats[stats.sheet != "—"]
    positions = int((df.kind == "position").sum())
    print("Локальных смет разобрано: {} из {}, без ручной правки {}".format(
        int((smety.positions > 0).sum()), len(smety), int(smety.clean.sum())))
    print("Позиций: {}, строк НР/СП: {}, нераспознанных строк: {}".format(
        positions, int(df.kind.isin(["overhead", "profit"]).sum()), int(smety.unknown.sum())))
    if vor is not None and len(vor):
        matched = int(sum(1 for n in vor.vor_no if n in (sheet_index or {})))
        print("Ведомости объёмов: строк {}, сопоставлено с позициями смет {}".format(len(vor), matched))
    print("Замечаний: {}, сумма влияния: {} руб.".format(len(findings), money(findings.impact_rub.sum())))
    if len(findings):
        top = findings.iloc[0]
        print("Крупнейшее: {} — {} ({}, строка {})".format(
            top.rule, money(top.impact_rub), top.sheet or top.file, top.excel_row))
        print("По видам: " + ", ".join(
            "{} {}".format(k, v) for k, v in findings.rule.value_counts().items()))
    print("Протокол: {}".format(out))
    print("Время: {:.1f} с на {} смет ({:.2f} с на смету)".format(
        elapsed, len(smety), elapsed / max(len(smety), 1)))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

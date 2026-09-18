# -*- coding: utf-8 -*-
"""Сверка протокола с заложенными дефектами: сколько поймали, что пропустили, что лишнего.

Качество должно называться числом. «Работает» — не метрика,
«9 из 10 заложенных дефектов, 3 замечания сверх плана, все три проверены вручную» — метрика.
"""
import io
import json
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from smeta.parse_xlsx import parse_dir  # noqa: E402
from smeta.validate import money, validate  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]


def main():
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    truth = pd.DataFrame(json.loads((ROOT / "data" / "fixtures_truth.json").read_text(encoding="utf-8")))
    df, stats = parse_dir(ROOT / "data" / "fixtures")
    found = validate(df)

    truth_keys = set(zip(truth.file, truth.row, truth.rule))
    found_keys = set(zip(found.file, found.excel_row, found.rule))

    caught = truth_keys & found_keys
    missed = truth_keys - found_keys
    extra = found_keys - truth_keys

    print("КАЧЕСТВО РАЗБОРА")
    print("  смет разобрано без ручной правки: {} из {}".format(int(stats.clean.sum()), len(stats)))
    print("  позиций извлечено: {}, нераспознанных строк: {}".format(
        int((df.kind == "position").sum()), int(stats.unknown.sum())))

    print("\nПОЛНОТА ПРОВЕРОК")
    print("  заложено дефектов: {}".format(len(truth_keys)))
    print("  поймано: {} ({:.0f}%)".format(len(caught), 100 * len(caught) / max(len(truth_keys), 1)))
    print("  пропущено: {}".format(len(missed)))
    print("  сверх плана: {}".format(len(extra)))
    print("  сумма влияния по протоколу: {} руб.".format(money(found.impact_rub.sum())))

    if missed:
        print("\nПРОПУЩЕНО (разбирать в первую очередь):")
        for file, row, rule in sorted(missed):
            note = truth[(truth.file == file) & (truth.row == row) & (truth.rule == rule)].note.iloc[0]
            print("  {} стр. {} — {}: {}".format(file, row, rule, note))

    if extra:
        print("\nСВЕРХ ПЛАНА (проверить руками: правило работает или шумит):")
        for file, row, rule in sorted(extra):
            m = found[(found.file == file) & (found.excel_row == row) & (found.rule == rule)].iloc[0]
            print("  {} стр. {} — {}: {} [{} руб.]".format(file, row, rule, m.message, money(m.impact_rub)))


if __name__ == "__main__":
    main()

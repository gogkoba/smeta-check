# -*- coding: utf-8 -*-
"""Регрессия на пайплайн. Запуск: python -m pytest tests -q  (или python tests/test_pipeline.py)

Тесты закрывают ровно те места, где я уже ошибался на реальных данных: разбор чисел,
границы коэффициентов, семантику колонок и NaN вместо None. Каждый из этих случаев
стоил пачки ложных замечаний, и без теста они вернутся при первой же правке.
"""
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from smeta.parse_vor import expand_refs, lsr_key, short_key, split_unit, vor_key  # noqa: E402
from smeta.parse_xlsx import (build_mapping, parse_coef, parse_dir, position_number,  # noqa: E402
                              to_number)
from smeta.validate import validate  # noqa: E402

FIXTURES = ROOT / "data" / "fixtures"


def test_to_number():
    assert to_number(1234.5) == 1234.5
    assert to_number("1 234,56") == 1234.56           # неразрывный пробел из выгрузки
    assert to_number("1 234,56") == 1234.56
    assert to_number("12 345,00 руб.") == 12345.0
    assert to_number("") is None
    assert to_number("—") is None
    assert to_number("100 м2") is None                # единица измерения не число


def test_parse_coef():
    assert parse_coef(None) == 1.0
    assert parse_coef("") == 1.0
    assert parse_coef(1) == 1.0
    assert parse_coef(0) == 0.0            # МАТ=0: материалы исключены при демонтаже
    assert parse_coef(52) == 52.0          # коэффициент из технической части сборника
    assert parse_coef("К=1,15 (стеснённые условия)") == 1.15


def test_position_number():
    """Номер позиции бывает с пометкой: «2 О» — оборудование, «13 М» — монтаж.
    Если такую ячейку не разобрать, позиция приклеится к предыдущей."""
    assert position_number("2 О") == 2.0
    assert position_number("13 М") == 13.0
    assert position_number(5) == 5.0
    assert position_number("Раздел 1. Демонтаж") is None
    assert position_number(None) is None


def test_subheader_priority():
    """«всего с учетом коэффициентов» — это объём, а не коэффициент."""
    header = ["№ пп", "Обоснование", "Наименование работ", None, None, "Единица измерения",
              "Количество", None, None, "Сметная стоимость в базисном уровне цен", None, None]
    sub = [[None] * 6 + ["на единицу", "коэффициенты", "всего с учетом коэффициентов",
                         "на единицу", "коэффициенты", "всего"]]
    m = build_mapping(header, sub)
    assert m["qty_unit"] == 6
    assert m["qty_coef"] == 7
    assert m["qty"] == 8, "объём должен браться из «всего с учётом коэффициентов»"
    assert m["price"] == 9
    assert m["total"] == 11


def test_base_level_not_confused_by_parenthetical():
    """В шапке базисной колонки есть слова «в текущем уровне цен» — в скобках."""
    header = ["Наименование работ", "Количество",
              "Сметная стоимость в базисном уровне цен "
              "(в текущем уровне цен (гр. 8) для ресурсов, отсутствующих в СНБ), руб."]
    m = build_mapping(header, [])
    assert "price" in m or "total" in m
    assert "price_current" not in m, "базисная колонка не должна стать текущей"


def test_vor_helpers():
    assert split_unit("100 м2") == (100.0, "м2")
    assert split_unit("1000 м3") == (1000.0, "м3")
    assert split_unit("т") == (1.0, "т")
    assert expand_refs("5-8") == [5, 6, 7, 8]
    assert expand_refs("2 3") == [2, 3]
    assert expand_refs("14 15") == [14, 15]
    assert expand_refs(None) == []
    assert vor_key("ВЕДОМОСТЬ ОБЪЕМОВ РАБОТ № 1.1") == ("1", True)
    assert vor_key("Ведомость объёмов работ № 2") == ("2", False)
    assert vor_key("ВЕДОМОСТЬ ОБЪЕМОВ РАБОТ № 02-01-14") == ("02-01-14", False)
    assert lsr_key("ЛОКАЛЬНЫЙ СМЕТНЫЙ РАСЧЕТ (СМЕТА) № 02-01-01д (2026)") == ("02-01-01", True)
    assert lsr_key("ЛОКАЛЬНЫЙ СМЕТНЫЙ РАСЧЕТ (СМЕТА) № 02-01-14") == ("02-01-14", False)
    # «СР-1д» не должен перехватывать ключ у настоящей ЛСР 02-01-01д
    assert lsr_key("ЛОКАЛЬНЫЙ СМЕТНЫЙ РАСЧЕТ (СМЕТА) № СР-1д (2026)") == (None, False)
    assert short_key("02-01-01", True) == "1.1"


def test_fixtures_parse_clean():
    df, stats = parse_dir(FIXTURES)
    assert len(stats) == 6, "должно разобраться шесть тестовых смет"
    assert stats.clean.all(), "все фикстуры разбираются без нераспознанных строк"
    assert int((df.kind == "position").sum()) == 45


def test_fixtures_recall():
    """Все заложенные дефекты должны находиться."""
    import json
    truth = pd.DataFrame(json.loads((ROOT / "data" / "fixtures_truth.json").read_text(encoding="utf-8")))
    df, _ = parse_dir(FIXTURES)
    found = validate(df)
    truth_keys = set(zip(truth.file, truth.row, truth.rule))
    found_keys = set(zip(found.file, found.excel_row, found.rule))
    missed = truth_keys - found_keys
    assert not missed, "пропущены дефекты: {}".format(sorted(missed))


def test_no_nan_leaks_into_findings():
    """NaN вместо None превращал сравнение в ложь и давал замечание на каждой строке."""
    df, _ = parse_dir(FIXTURES)
    found = validate(df)
    assert found.impact_rub.notna().all()
    assert not found.message.str.contains("nan", case=False).any()


if __name__ == "__main__":
    failures = 0
    for name, func in sorted(globals().items()):
        if name.startswith("test_") and callable(func):
            try:
                func()
                print("OK   {}".format(name))
            except AssertionError as e:
                failures += 1
                print("FAIL {}: {}".format(name, e))
    print("\n{} проверок, ошибок: {}".format(
        sum(1 for n in globals() if n.startswith("test_")), failures))
    raise SystemExit(1 if failures else 0)

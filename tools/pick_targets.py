# -*- coding: utf-8 -*-
"""Отбор закупок из выгрузки поиска ЕИС: где реально лежит сметная документация.

Выгрузка результатов поиска — это не сметы, а оглавление. Скрипт чистит его от
непрофильного и ранжирует так, чтобы качать сверху вниз и быстро набрать корпус
с пересекающимися шифрами расценок (иначе правило по медиане цен не на чем считать).
"""
import argparse
import csv
import io
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

# ОКПД2 строительных разделов: 41 — здания, 42 — сооружения, 43 — спецработы
BUILD_OKPD = re.compile(r"^(41|42|43)\.")

# однотипные работы ценнее разнородных: шифры расценок пересекаются между сметами
HOMOGENEOUS = [
    ("капитальный ремонт", 30), ("капремонт", 30), ("текущий ремонт", 20),
    ("ремонт кровли", 25), ("фасад", 20), ("благоустройство", 20),
    ("замена лифт", 15), ("инженерных сетей", 15), ("дорог", 15),
]
BUILD_WORDS = [
    "строительств", "реконструкц", "капитальн", "ремонт", "благоустройств",
    "монтаж", "кровл", "фасад", "сет", "дорог", "здани", "сооружен",
]
JUNK_WORDS = [
    "поставка", "охран", "медицин", "лицензи", "программн", "питани",
    "услуги связи", "страхован", "автомобил", "оборудован для",
]


def load(path):
    raw = Path(path).read_bytes()
    for enc in ("cp1251", "utf-8-sig", "utf-8"):
        try:
            return list(csv.DictReader(io.StringIO(raw.decode(enc)), delimiter=";"))
        except UnicodeDecodeError:
            continue
    raise SystemExit("не смог определить кодировку файла")


def num(s):
    try:
        return float(str(s or "").replace(" ", "").replace(",", "."))
    except ValueError:
        return 0.0


def score(row):
    """Чем выше балл, тем вероятнее внутри нормальный комплект ЛСР."""
    name = ((row.get("Наименование закупки") or "") + " " +
            (row.get("Наименование лота") or "")).lower()
    okpd = (row.get("Классификация по ОКПД2") or "").strip()
    law = (row.get("Закупки по") or "").strip()
    stage = (row.get("Этап закупки") or "").strip()
    price = num(row.get("Начальная (максимальная) цена контракта"))
    pp615 = (row.get("Предмет электронного аукциона (только для ПП РФ 615)") or "").strip()

    if any(w in name for w in JUNK_WORDS) and not any(w in name for w in BUILD_WORDS):
        return None
    if not (BUILD_OKPD.match(okpd) or any(w in name for w in BUILD_WORDS)):
        return None

    s = 0
    if BUILD_OKPD.match(okpd):
        s += 40
    if law == "44-ФЗ":          # документация обязательна и лежит открыто
        s += 25
    if stage == "Закупка завершена":  # ничего не переиграют, файлы на месте
        s += 15
    elif stage in ("Работа комиссии", "Подача заявок"):
        s += 5
    if pp615:                   # капремонт МКД — самый однородный корпус
        s += 25
    for word, w in HOMOGENEOUS:
        if word in name:
            s += w
            break
    if 10e6 <= price <= 1e9:    # ниже — мелочь без ЛСР, выше — стройка на годы
        s += 20
    elif 3e6 <= price < 10e6:
        s += 8
    return s


def links(row):
    reg = re.sub(r"\D", "", row.get("Реестровый номер закупки") or "")
    law = (row.get("Закупки по") or "").strip()
    if law == "44-ФЗ":
        card = "https://zakupki.gov.ru/epz/order/notice/ea44/view/documents.html?regNumber=" + reg
    else:
        card = "https://zakupki.gov.ru/223/purchase/public/purchase/info/documents.html?regNumber=" + reg
    search = "https://zakupki.gov.ru/epz/order/quicksearch/search.html?searchString=" + reg
    return reg, card, search


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("csv_path", help="выгрузка результатов поиска ЕИС")
    ap.add_argument("--top", type=int, default=25)
    ap.add_argument("--out", default=str(ROOT / "data" / "targets.csv"))
    args = ap.parse_args()
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

    rows = load(args.csv_path)
    scored = []
    for r in rows:
        s = score(r)
        if s is None:
            continue
        reg, card, search = links(r)
        scored.append(dict(
            score=s, reg=reg, law=(r.get("Закупки по") or "").strip(),
            stage=(r.get("Этап закупки") or "").strip(),
            price=num(r.get("Начальная (максимальная) цена контракта")),
            name=re.sub(r"\s+", " ", (r.get("Наименование закупки") or "")).strip(),
            customer=re.sub(r"\s+", " ", (r.get("Наименование Заказчика") or "")).strip(),
            okpd=(r.get("Классификация по ОКПД2") or "").split(":")[0],
            card=card, search=search,
        ))
    scored.sort(key=lambda d: (-d["score"], -d["price"]))

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    with out.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(scored[0].keys()) if scored else ["score"], delimiter=";")
        w.writeheader()
        w.writerows(scored)

    print("Строк в выгрузке: {} | профильных: {}".format(len(rows), len(scored)))
    print("Полный отобранный список: {}\n".format(out))
    print("ТОП-{} — качать сверху вниз:\n".format(args.top))
    for i, d in enumerate(scored[:args.top], 1):
        print("{:>2}. [{:>3}] {} · {} · {:,.0f} ₽".format(
            i, d["score"], d["law"], d["stage"], d["price"]).replace(",", " "))
        print("    {}".format(d["name"][:120]))
        print("    {}".format(d["card"]))
    print("\nЕсли прямая ссылка не открылась — искать по номеру:")
    print("https://zakupki.gov.ru/epz/order/quicksearch/search.html?searchString=<реестровый номер>")


if __name__ == "__main__":
    main()

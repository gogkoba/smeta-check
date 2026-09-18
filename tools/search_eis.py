# -*- coding: utf-8 -*-
"""Поиск строительных закупок в ЕИС и сбор ссылок на карточки.

Несколько запросов по разным видам работ вместо одного: корпус нужен разнородный.
Смета от школы, от кровельщиков и от дорожников свёрстана по-разному, и парсер,
подогнанный под один шаблон, ничего не докажет.

    python tools/search_eis.py --out data/urls.txt --per-query 2
"""
import argparse
import io
import re
import sys
import time
from pathlib import Path
from urllib.parse import quote

import requests

ROOT = Path(__file__).resolve().parents[1]
CA_BUNDLE = ROOT / "data" / "certs" / "russian_trusted_root.pem"

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"),
    "Accept-Language": "ru-RU,ru;q=0.9",
}

BASE = "https://zakupki.gov.ru/epz/order/extendedsearch/results.html"

# Разные виды работ дают разные сборники расценок и разных сметчиков
QUERIES = [
    "капитальный ремонт здания",
    "капитальный ремонт кровли",
    "ремонт фасада здания",
    "капитальный ремонт системы отопления",
    "благоустройство территории",
    "строительство здания",
    "капитальный ремонт помещений",
    "ремонт спортивного зала",
]

CARD_RE = re.compile(r"regNumber=(\d{11,})")


def search(session, query, page, price_from):
    url = (BASE + "?searchString=" + quote(query) +
           "&morphology=on&pageNumber={}&sortDirection=false&recordsPerPage=_50"
           "&sortBy=UPDATE_DATE&fz44=on&pc=on&currencyIdGeneral=-1"
           "&priceFromGeneral={}".format(page, price_from))
    r = session.get(url, timeout=60)
    if r.status_code != 200:
        return []
    return list(dict.fromkeys(CARD_RE.findall(r.text)))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=str(ROOT / "data" / "urls.txt"))
    ap.add_argument("--per-query", type=int, default=2, help="страниц выдачи на запрос")
    ap.add_argument("--price-from", type=int, default=15_000_000)
    ap.add_argument("--limit", type=int, default=60, help="сколько закупок оставить")
    ap.add_argument("--delay", type=float, default=1.5)
    args = ap.parse_args()
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

    session = requests.Session()
    session.headers.update(HEADERS)
    if CA_BUNDLE.exists():
        session.verify = str(CA_BUNDLE)

    by_query, seen = {}, set()
    for query in QUERIES:
        got = []
        for page in range(1, args.per_query + 1):
            got.extend(search(session, query, page, args.price_from))
            time.sleep(args.delay)
        fresh = [r for r in got if r not in seen]
        seen.update(fresh)
        by_query[query] = fresh
        print("{:<40} найдено {}  новых {}".format(query[:38], len(got), len(fresh)))

    # берём по кругу из каждого запроса: иначе весь корпус приедет от первого
    found, regs = {}, []
    cursors = {q: 0 for q in QUERIES}
    while len(regs) < args.limit and any(cursors[q] < len(by_query[q]) for q in QUERIES):
        for q in QUERIES:
            if cursors[q] < len(by_query[q]) and len(regs) < args.limit:
                reg = by_query[q][cursors[q]]
                cursors[q] += 1
                regs.append(reg)
                found[reg] = q
    lines = ["# Ссылки на карточки закупок. Собрано tools/search_eis.py",
             "# Запуск: python tools/fetch_docs.py --urls data/urls.txt --only-smeta", ""]
    for reg in regs:
        lines.append("# " + found[reg])
        lines.append("https://zakupki.gov.ru/epz/order/notice/ea44/view/documents.html?regNumber=" + reg)
    Path(args.out).write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("\nВсего закупок: {} -> {}".format(len(regs), args.out))


if __name__ == "__main__":
    raise SystemExit(main())

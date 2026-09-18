# -*- coding: utf-8 -*-
"""Из сохранённой карточки закупки — список файлов со сметами и прямыми ссылками.

Зачем: в строительной закупке к извещению приложено 50–200 файлов, и вручную
выцеплять из них ЛСР по одному — потерянный вечер. Скрипт разбирает сохранённый
HTML карточки, раскладывает вложения по типам и собирает страницу со ссылками,
где сметы лежат сверху: открыл, прокликал нужное, всё скачалось в браузер.

    python tools/extract_docs.py "C:/Users/Georgii/Downloads/Карточка закупки.html"
    python tools/extract_docs.py C:/Users/Georgii/Downloads --out data/doc_links.html
"""
import argparse
import csv
import html
import io
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

LINK_RE = re.compile(
    r'<a\s+href="(https://zakupki\.gov\.ru/[^"]*filestore[^"]*uid=[0-9A-Fa-f\-]+)"[^>]*>(.*?)</a>',
    re.S | re.I)
REG_RE = re.compile(r'(?:purchaseNoticeNumber|regNumber)=(\d{11,})')

# что ищем, в порядке убывания ценности для пайплайна.
# Короткие аббревиатуры (ЛСР, ССР, ВОР) ищем только как отдельные слова:
# иначе «проект догоВОРа» попадает в ведомости объёмов.
CATEGORIES = [
    # «смет» стемом, а не «смета»: иначе «Проект сметы контракта» уезжает в договоры
    ("смета", 0, ["локальн", "смет", "объектн", "нмцк"],
     [r"\bлср\b", r"\bсср\b"]),
    ("объёмы", 1, ["ведомость объем", "ведомость объём", "дефектн"], [r"\bвор\b"]),
    ("проектная документация", 3, ["проектн", "рабочая документ", "шифр"], []),
    ("договор", 8, ["проект договора", "проект контракта", "контракт"], []),
    ("прочее", 9, [], []),
]

SMETA_EXT = (".xls", ".xlsx", ".xlsm", ".gsfx", ".arp", ".xml")
ARCHIVE_EXT = (".zip", ".rar", ".7z")


def unwrap(name):
    """ЕИС заворачивает вложения в архив: «Обоснование НМЦК.xlsx.zip» -> («…НМЦК.xlsx», True).

    Классифицировать и оценивать пригодность нужно по имени внутри архива, иначе
    все сметы схлопываются в одну кучу «архив».
    """
    archived = False
    while Path(name).suffix.lower() in ARCHIVE_EXT:
        name = name[: -len(Path(name).suffix)]
        archived = True
    return name, archived


def categorize(name):
    inner, archived = unwrap(name)
    low = inner.lower()
    for label, order, keys, patterns in CATEGORIES:
        if any(k in low for k in keys) or any(re.search(p, low) for p in patterns):
            return (label + (" (в архиве)" if archived else "")), order
    return ("архив" if archived else "прочее"), (5 if archived else 9)


def parse_card(path):
    raw = Path(path).read_bytes()
    for enc in ("utf-8", "cp1251"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        return []
    reg = REG_RE.search(text)
    reg = reg.group(1) if reg else Path(path).stem
    out, seen = [], set()
    for url, inner in LINK_RE.findall(text):
        name = html.unescape(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", inner))).strip()
        if not name or url in seen:
            continue
        seen.add(url)
        label, order = categorize(name)
        inner, archived = unwrap(name)
        ext = Path(inner).suffix.lower()
        out.append(dict(reg=reg, name=name, url=url, category=label, order=order,
                        ext=ext, archived=archived, parsable=ext in SMETA_EXT,
                        source=Path(path).name))
    return out


def write_html(rows, out_path):
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    by_reg = {}
    for r in rows:
        by_reg.setdefault(r["reg"], []).append(r)

    parts = ["""<meta charset="utf-8"><title>Файлы закупок</title>
<style>
 body{font:15px/1.5 -apple-system,Segoe UI,sans-serif;margin:2rem auto;max-width:1000px;color:#1a1a1a}
 h2{margin:2rem 0 .5rem;font-size:1.05rem;border-bottom:2px solid #1F3864;padding-bottom:.3rem}
 table{border-collapse:collapse;width:100%}
 td{padding:.35rem .5rem;border-bottom:1px solid #eee;vertical-align:top}
 .cat{color:#666;font-size:.85em;white-space:nowrap}
 tr.hot{background:#FFF6D8}
 tr.hot td:first-child::before{content:"\\2605 ";color:#B8860B}
 a{color:#1F3864}
 .hint{color:#666;font-size:.9em;margin-bottom:1.5rem}
</style>
<h1>Файлы закупок</h1>
<p class=hint>Сметы и ведомости объёмов подняты наверх и отмечены звёздочкой. Кликать по ссылке — файл скачается браузером.
Класть скачанное в <code>data/raw/</code> (xlsx) или <code>data/raw_pdf/</code> (pdf).</p>"""]

    for reg, items in by_reg.items():
        items.sort(key=lambda r: (r["order"], r["name"].lower()))
        hot = sum(1 for r in items if r["order"] <= 1)
        parts.append('<h2>Закупка {} — файлов {}, из них сметных {}</h2><table>'.format(reg, len(items), hot))
        for r in items:
            cls = ' class="hot"' if r["order"] <= 1 else ""
            parts.append('<tr{}><td><a href="{}">{}</a></td><td class=cat>{}</td></tr>'.format(
                cls, r["url"], html.escape(r["name"]), r["category"]))
        parts.append("</table>")
    out_path.write_text("\n".join(parts), encoding="utf-8")
    return out_path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("path", help="сохранённая карточка .html или папка с ними")
    ap.add_argument("--out", default=str(ROOT / "data" / "doc_links.html"))
    ap.add_argument("--csv", default=str(ROOT / "data" / "doc_links.csv"))
    args = ap.parse_args()
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

    p = Path(args.path)
    files = sorted(p.glob("*.html")) if p.is_dir() else [p]
    rows = []
    for f in files:
        got = parse_card(f)
        rows.extend(got)
        print("{}: вложений {}".format(f.name, len(got)))

    if not rows:
        print("\nСсылок на файлы не нашлось. Сохранять карточку нужно целиком "
              "(Ctrl+S, тип «Веб-страница полностью» или «HTML только»), уже открыв вкладку «Документы».")
        return 1

    hot = [r for r in rows if r["order"] <= 1]
    print("\nВсего вложений: {} | сметных и ведомостей: {}".format(len(rows), len(hot)))
    for r in sorted(hot, key=lambda r: (r["order"], r["name"])):
        flag = "" if r["parsable"] else "  (не xlsx — пойдёт в PDF-ветку или руками)"
        print("  [{}] {}{}".format(r["category"], r["name"][:90], flag))

    with open(args.csv, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()), delimiter=";")
        w.writeheader()
        w.writerows(rows)
    out = write_html(rows, args.out)
    print("\nСтраница со ссылками: {}".format(out))
    print("Таблица: {}".format(args.csv))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

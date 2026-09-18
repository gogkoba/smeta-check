# -*- coding: utf-8 -*-
"""Качалка вложений ЕИС: список ссылок на закупки -> скачанные и разложенные файлы.

Два режима, потому что ЕИС может не пустить бота на страницы:

  1) по ссылкам   — сам открывает карточки, вытаскивает вложения, качает
     python tools/fetch_docs.py --urls data/urls.txt

  2) по сохранённым карточкам — если режим 1 упёрся в защиту: страницы сохраняешь
     браузером (Ctrl+S), скрипт качает только файлы
     python tools/fetch_docs.py --cards C:/Users/Georgii/Downloads

Скачанное распаковывается и раскладывается: сметные xlsx в data/raw, pdf в
data/raw_pdf, остальное в data/downloads/other. Повторный запуск не перекачивает
то, что уже лежит.
"""
import argparse
import csv
import io
import re
import sys
import time
import zipfile
from pathlib import Path
from urllib.parse import unquote

import requests

sys.path.insert(0, str(Path(__file__).resolve().parent))
from extract_docs import LINK_RE, REG_RE, SMETA_EXT, categorize, parse_card, unwrap  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
DOWNLOADS = ROOT / "data" / "downloads"

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "ru-RU,ru;q=0.9",
    "Referer": "https://zakupki.gov.ru/",
}


CA_BUNDLE = ROOT / "data" / "certs" / "russian_trusted_root.pem"


def make_session(insecure=False):
    """Сессия с доверием к системному хранилищу сертификатов.

    У zakupki.gov.ru сертификат российского УЦ: Windows его знает, а certifi внутри
    Python — нет, и запрос падает на проверке TLS. truststore подсовывает системное
    хранилище; если пакета нет, остаётся выключить проверку.
    """
    s = requests.Session()
    s.headers.update(HEADERS)
    if CA_BUNDLE.exists() and not insecure:
        s.verify = str(CA_BUNDLE)
        print("TLS: строгая проверка по закреплённому корню Минцифры ({})".format(CA_BUNDLE.name))
        return s
    try:
        import truststore
        truststore.inject_into_ssl()
        print("TLS: системное хранилище сертификатов (truststore)")
    except ImportError:
        if insecure:
            s.verify = False
            requests.packages.urllib3.disable_warnings()
            print("TLS: ПРОВЕРКА ВЫКЛЮЧЕНА (--insecure). Для разовой выгрузки открытых данных допустимо")
        else:
            print("TLS: certifi. Если упадёт на проверке сертификата — "
                  "поставь truststore (pip install truststore) или запусти с --insecure")
    return s


def to_documents_url(url):
    """Ссылку на любую вкладку закупки приводим к вкладке «Документы»."""
    url = url.strip()
    if not url or url.startswith("#"):
        return None
    if "documents.html" in url:
        return url
    if "/223/" in url:
        return re.sub(r"/(common-info|supplier-results|journal)\.html", "/documents.html", url)
    return re.sub(r"/(common-info|documents-history|supplier-results|journal|contract-info)\.html",
                  "/documents.html", url)


def fetch_page(session, url, delay, retries=3):
    for attempt in range(1, retries + 1):
        try:
            r = session.get(url, timeout=40)
            if r.status_code == 200 and "filestore" in r.text:
                return r.text
            if r.status_code == 200:
                return r.text  # страница открылась, но вложений нет — разберём выше
            print("    HTTP {} (попытка {}/{})".format(r.status_code, attempt, retries))
        except requests.RequestException as e:
            print("    сеть: {} (попытка {}/{})".format(type(e).__name__, attempt, retries))
        time.sleep(delay * attempt)
    return None


def links_from_html(text, fallback_reg=""):
    reg = REG_RE.search(text)
    reg = reg.group(1) if reg else fallback_reg
    out, seen = [], set()
    for url, inner in LINK_RE.findall(text):
        import html as htmlmod
        name = htmlmod.unescape(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", inner))).strip()
        if not name or url in seen:
            continue
        seen.add(url)
        label, order = categorize(name)
        base, archived = unwrap(name)
        out.append(dict(reg=reg, name=name, url=url, category=label, order=order,
                        ext=Path(base).suffix.lower(), archived=archived))
    return out


def filename_from_response(r, fallback):
    """Имя файла из заголовка Content-Disposition. ЕИС шлёт его то в utf-8, то в cp1251."""
    cd = r.headers.get("Content-Disposition", "")
    m = re.search(r"filename\*=(?:UTF-8'')?([^;]+)", cd, re.I)
    if m:
        return unquote(m.group(1).strip('"'))
    m = re.search(r'filename="?([^";]+)"?', cd, re.I)
    if m:
        raw = m.group(1)
        for enc in ("utf-8", "cp1251"):
            try:
                decoded = raw.encode("latin-1").decode(enc)
                if not decoded.count("\ufffd"):
                    return decoded
            except (UnicodeEncodeError, UnicodeDecodeError):
                continue
        return raw
    return fallback


def safe_name(name):
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip() or "file"


def download(session, item, dest_dir, delay, max_mb):
    dest_dir.mkdir(parents=True, exist_ok=True)
    uid = re.search(r"uid=([0-9A-Fa-f\-]+)", item["url"])
    uid = uid.group(1)[:12] if uid else "nouid"
    existing = list(dest_dir.glob(uid + "_*"))
    if existing:
        return existing[0], "уже скачан"
    try:
        r = session.get(item["url"], timeout=90, stream=True)
    except requests.RequestException as e:
        return None, "сеть: {}".format(type(e).__name__)
    if r.status_code != 200:
        return None, "HTTP {}".format(r.status_code)

    size_mb = int(r.headers.get("Content-Length", 0)) / 1e6
    if max_mb and size_mb > max_mb:
        return None, "пропущен: {:.0f} МБ > лимита {} МБ".format(size_mb, max_mb)

    name = safe_name(filename_from_response(r, item["name"]))
    path = dest_dir / "{}_{}".format(uid, name)
    written = 0
    with path.open("wb") as f:
        for chunk in r.iter_content(1 << 16):
            f.write(chunk)
            written += len(chunk)
            if max_mb and written > max_mb * 1e6:
                f.close()
                path.unlink(missing_ok=True)
                return None, "пропущен: больше лимита {} МБ".format(max_mb)
    time.sleep(delay)
    return path, "{:.1f} МБ".format(written / 1e6)


def unpack(path, out_dir):
    """Разворачиваем zip (в том числе вложенные). rar оставляем как есть — нужен 7-Zip."""
    produced = []
    if path.suffix.lower() != ".zip":
        return [path]
    out_dir.mkdir(parents=True, exist_ok=True)
    try:
        with zipfile.ZipFile(path) as z:
            for info in z.infolist():
                if info.is_dir():
                    continue
                raw = info.filename
                try:
                    inner_name = raw.encode("cp437").decode("cp866")
                except (UnicodeEncodeError, UnicodeDecodeError):
                    inner_name = raw
                target = out_dir / safe_name(Path(inner_name).name)
                with z.open(info) as src, target.open("wb") as dst:
                    dst.write(src.read())
                produced.extend(unpack(target, out_dir) if target.suffix.lower() == ".zip" else [target])
    except zipfile.BadZipFile:
        return [path]
    return produced


SMETA_DEST = ROOT / "data" / "raw"


def sort_file(path):
    """Раскладка по назначению. Возвращает (куда положили, почему)."""
    ext = path.suffix.lower()
    label, _ = categorize(path.name)
    if ext in SMETA_EXT:
        dest = SMETA_DEST
        why = "сметный формат"
    elif ext == ".pdf":
        dest = ROOT / "data" / "raw_pdf"
        why = "pdf, ветка M6"
    else:
        dest = DOWNLOADS / "other"
        why = "не для парсера"
    dest.mkdir(parents=True, exist_ok=True)
    target = dest / path.name
    if target.exists() and target.stat().st_size == path.stat().st_size:
        return target, why + ", уже был"
    path.replace(target)
    return target, why


def main():
    ap = argparse.ArgumentParser()
    src = ap.add_mutually_exclusive_group(required=True)
    src.add_argument("--urls", help="файл со ссылками на карточки закупок, по одной в строке")
    src.add_argument("--cards", help="папка с сохранёнными карточками .html")
    ap.add_argument("--delay", type=float, default=1.5, help="пауза между запросами, с")
    ap.add_argument("--max-mb", type=float, default=200, help="не качать файлы больше, МБ")
    ap.add_argument("--only-smeta", action="store_true", help="качать только сметы и ведомости объёмов")
    ap.add_argument("--dest", default=None, help="куда складывать разобранные сметы")
    ap.add_argument("--insecure", action="store_true", help="не проверять TLS-сертификат")
    args = ap.parse_args()
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

    global SMETA_DEST
    if args.dest:
        SMETA_DEST = Path(args.dest).resolve()
    session = make_session(args.insecure)
    items = []

    if args.urls:
        urls = [u.strip() for u in Path(args.urls).read_text(encoding="utf-8").splitlines()
                if u.strip() and not u.strip().startswith("#")]
        print("Ссылок на закупки: {}\n".format(len(urls)))
        for i, u in enumerate(urls, 1):
            page_url = to_documents_url(u)
            print("[{}/{}] {}".format(i, len(urls), page_url))
            text = fetch_page(session, page_url, args.delay)
            if not text:
                print("    не открылась — сохрани карточку браузером и используй --cards")
                continue
            found = links_from_html(text)
            print("    вложений: {}".format(len(found)))
            if not found:
                print("    вложений не видно: вероятно, страницу отдали без них (защита от ботов). "
                      "Через --cards сработает")
            items.extend(found)
            time.sleep(args.delay)
    else:
        folder = Path(args.cards)
        cards = sorted(folder.glob("*.html"))
        print("Сохранённых карточек: {}\n".format(len(cards)))
        for c in cards:
            found = parse_card(c)
            print("{}: вложений {}".format(c.name, len(found)))
            items.extend(found)

    if args.only_smeta:
        items = [i for i in items if i["order"] <= 1]
    if not items:
        print("\nНечего качать.")
        return 1

    print("\nК скачиванию: {} файлов".format(len(items)))
    manifest = []
    staging = DOWNLOADS / "archive"
    unpacked_dir = DOWNLOADS / "unpacked"

    for i, item in enumerate(items, 1):
        print("[{}/{}] {} [{}]".format(i, len(items), item["name"][:70], item["category"]))
        path, status = download(session, item, staging, args.delay, args.max_mb)
        print("    {}".format(status))
        if not path:
            manifest.append(dict(reg=item["reg"], name=item["name"], url=item["url"],
                                 status=status, dest=""))
            continue
        for produced in unpack(path, unpacked_dir):
            if item.get("reg") and not produced.name.startswith(str(item["reg"])[:8]):
                tagged = produced.with_name("{}_{}".format(str(item["reg"])[:8], produced.name))
                produced = produced.replace(tagged)
            dest, why = sort_file(produced)
            try:
                shown = dest.resolve().relative_to(ROOT)
            except ValueError:
                shown = dest
            print("      -> {} ({})".format(shown, why))
            manifest.append(dict(reg=item["reg"], name=produced.name, url=item["url"],
                                 status=status, dest=str(dest)))

    man_path = DOWNLOADS / "manifest.csv"
    man_path.parent.mkdir(parents=True, exist_ok=True)
    with man_path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["reg", "name", "url", "status", "dest"], delimiter=";")
        w.writeheader()
        w.writerows(manifest)

    raw_count = len(list((ROOT / "data" / "raw").glob("*.xls*")))
    print("\nВ data/raw теперь {} файлов Excel".format(raw_count))
    print("Что откуда взялось: {}".format(man_path.relative_to(ROOT)))
    print("Дальше: python -m smeta.cli")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

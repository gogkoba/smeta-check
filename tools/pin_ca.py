# -*- coding: utf-8 -*-
"""Закрепление корневого сертификата Минцифры для доступа к zakupki.gov.ru.

Зачем: сайт подписан цепочкой Russian Trusted Root CA, которой нет ни в certifi,
ни в хранилище Windows, поэтому обычный requests падает на проверке TLS. Вместо
того чтобы выключать проверку целиком (--insecure, MITM не заметишь), забираем
корень из живой цепочки, сохраняем и дальше проверяем строго по нему.

Оговорка: это доверие при первом подключении. Если первый коннект уже был
подменён — закрепится сертификат подменяющего. Отпечаток печатается: сверить
его с опубликованным на gosuslugi.ru/crt минута работы, и тогда сомнений нет.

    python tools/pin_ca.py
"""
import hashlib
import io
import socket
import ssl
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "data" / "certs" / "russian_trusted_root.pem"
HOST = "zakupki.gov.ru"


def describe(der):
    import os
    import tempfile
    pem = ssl.DER_cert_to_PEM_cert(der)
    f = tempfile.NamedTemporaryFile("w", suffix=".pem", delete=False, encoding="ascii")
    f.write(pem)
    f.close()
    try:
        info = ssl._ssl._test_decode_cert(f.name)
    finally:
        os.unlink(f.name)
    flat = lambda field: ", ".join("{}={}".format(k, v) for t in info.get(field, ()) for k, v in t)
    return pem, flat("subject"), flat("issuer"), info.get("notAfter")


def main():
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    with socket.create_connection((HOST, 443), timeout=20) as sock:
        with ctx.wrap_socket(sock, server_hostname=HOST) as ss:
            chain = ss.get_verified_chain()

    print("Цепочка сертификатов {}:".format(HOST))
    bundle = []
    for i, der in enumerate(chain):
        pem, subject, issuer, until = describe(der)
        role = "сертификат сайта" if i == 0 else ("промежуточный" if i < len(chain) - 1 else "КОРЕНЬ")
        print("  [{}] {}".format(i, role))
        print("      кому : {}".format(subject[:110]))
        print("      кем  : {}".format(issuer[:110]))
        print("      до   : {}".format(until))
        if i > 0:
            bundle.append(pem)
        if i == len(chain) - 1:
            fp = hashlib.sha256(der).hexdigest().upper()
            print("\n  Отпечаток корня SHA-256:")
            print("  " + " ".join(fp[j:j + 4] for j in range(0, len(fp), 4)))
            print("  Сверить с опубликованным на https://www.gosuslugi.ru/crt")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text("".join(bundle), encoding="ascii")
    print("\nСохранено: {}".format(OUT.relative_to(ROOT)))

    import requests
    r = requests.get("https://" + HOST + "/epz/main/public/home.html",
                     verify=str(OUT), timeout=30,
                     headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                                            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"})
    print("Проверка со строгой валидацией по этому корню: HTTP {}".format(r.status_code))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

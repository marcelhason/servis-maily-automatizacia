"""
Vyrobi .eml drafty pre Outlook podla NOVEHO vzoru `VZOR_2 - ...`.

Na rozdiel od `vyrob_maily.py` (stary vzor) ma novy vzor jednoduchsiu MIME
strukturu, ktoru Marcel vyrobil priamo v Outlooku ako koncept (draft):

    multipart/mixed (_004_)
    |-- multipart/alternative (_000_)
    |     |-- text/plain   (quoted-printable, windows-1250)
    |     +-- text/html    (quoted-printable, windows-1250)
    +-- priloha .xlsx       (base64)

Ziadny vnoreny obrazok podpisu (image001.png) ani multipart/related, co bol
pravdepodobne dovod, preco sa stary vzor v Outlooku ako draft nespraval dobre.

HTML tabulka je tu vyrobena novym Outlook web editorom (div.skipProofing +
inline rgb() styly), nie stara MsoNormalTable. Preto ma tento skript vlastny
generator riadkov `build_data_rows`, ktory presne replikuje format VZOR_2.

Pouzitie:
    python vyrob_maily_v2.py --test       # 1 nahodny servis
    python vyrob_maily_v2.py --limit 5    # prvych 5 servisov
    python vyrob_maily_v2.py              # vsetky servisy s mailom
"""
import argparse
import base64
import datetime as dt
import quopri
import random
import re
import sys
import uuid
from pathlib import Path

# Bezpecne spolocne funkcie zo stareho skriptu (nezavisia na MIME boundaries)
sys.path.insert(0, str(Path(__file__).parent))
from vyrob_maily import (  # noqa: E402
    ROOT, SERVIS_DIR, OUT_DIR, SUBJECT_PREFIX,
    sanitize_filename, fmt_date,
    load_unikaty, build_filename_map, read_xlsx_rows,
    build_plain_body,
    to_quoted_printable_cp1250, encode_header_q, q_encode_filename,
)

VZOR2 = ROOT / "VZOR_2 - Žiadosť o informácie k doriešeniu reklamácií – Candy Hoover ČR s.r.o..eml"

# Boundaries prevzate z VZOR_2 (lokalne v ramci jedneho .eml, mozu byt fixne).
BOUNDARY_004 = "_004_GVUPR03MB117227BB5E654EF827DD7C909C11A2GVUPR03MB11722eu_"
BOUNDARY_000 = "_000_GVUPR03MB117227BB5E654EF827DD7C909C11A2GVUPR03MB11722eu_"

# Sirky stlpcov tabulky (pt), v poradi: Hlaseni, Krat.text, Znacka, Servis, Datum, Poznamka
COL_WIDTHS = [63, 221, 94, 268, 79, 265]

# Spolocny style atribut pre vnutorny <div> bunky (font, atd.)
DIV_FONT = ("margin: 0px; font-family: Aptos, Aptos_EmbeddedFont, "
            "Aptos_MSFontService, Calibri, Helvetica, sans-serif; "
            "font-size: 12pt; color: rgb(0, 0, 0);")


# ---------- HTML escaping ----------

def esc(s: str) -> str:
    """Minimalny HTML escape pre obsah bunky."""
    return (s.replace("&", "&amp;")
             .replace("<", "&lt;")
             .replace(">", "&gt;"))


# ---------- generator datovych riadkov (format VZOR_2) ----------

def _cell(width_pt: int, content_html: str, *, align: str,
          first_col: bool, is_last_row: bool) -> str:
    """Postavi jednu <td> bunku presne v style VZOR_2.

    Borders:
      - prvy stlpec ma vzdy lavu hranicu (border-left)
      - posledny riadok ma vo vsetkych bunkach spodnu hranicu (border-bottom)
    """
    height = "15.75pt" if is_last_row else "15pt"
    parts = ["text-align: left;"]
    if is_last_row:
        parts.append("border-bottom: 1pt solid;")
    if first_col:
        parts.append("border-left: 1pt solid;")
    parts.append(f"padding: 0cm 3.5pt; vertical-align: top; "
                 f"width: {width_pt}pt; height: {height};")
    td_style = " ".join(parts)
    div = (f'<div class="skipProofing" style="text-align: {align}; {DIV_FONT}">\n'
           f'{content_html}</div>')
    return f'<td style="{td_style}">\n{div}\n</td>'


def build_data_rows(rows) -> str:
    """Vrati HTML string s datovymi <tr> riadkami (bez hlavickoveho riadku)."""
    out = []
    last_idx = len(rows) - 1
    for i, row in enumerate(rows):
        is_last = (i == last_idx)
        hlaseni, ktext, znacka, servis, datum, poznamka = row
        values = [
            str(hlaseni) if hlaseni is not None else "",
            str(ktext) if ktext else "",
            str(znacka) if znacka else "",
            str(servis) if servis else "",
            fmt_date(datum),
            str(poznamka) if poznamka else "",
        ]
        cells = []
        for j, (w, val) in enumerate(zip(COL_WIDTHS, values)):
            align = "right" if j == 4 else "left"  # datum vpravo
            content = esc(val) if val else "&nbsp;"
            cells.append(_cell(w, content, align=align,
                               first_col=(j == 0), is_last_row=is_last))
        out.append("<tr>\n" + "\n".join(cells) + "\n</tr>")
    return "\n".join(out)


# ---------- extrakcia HTML/plain sablony zo VZOR_2 ----------

def _decode_part(vzor_bytes: bytes, content_type_prefix: bytes) -> str:
    """Vrati dekodovane (cp1250) telo casti _000_ daneho Content-Type."""
    start = vzor_bytes.index(
        b"--" + BOUNDARY_000.encode() + b"\r\nContent-Type: " + content_type_prefix)
    body_start = vzor_bytes.index(b"\r\n\r\n", start) + 4
    body_end = vzor_bytes.index(b"\r\n--" + BOUNDARY_000.encode(), body_start)
    return quopri.decodestring(vzor_bytes[body_start:body_end]).decode("cp1250")


def build_html_body(vzor_bytes: bytes, rows) -> bytes:
    """Postavi text/html telo: zachova vsetko zo VZOR_2 okrem datovych riadkov,
    ktore nahradi novymi. Hlavickovy riadok tabulky (prvy <tr>) ostava."""
    html = _decode_part(vzor_bytes, b"text/html")
    trs = list(re.finditer(r"<tr[\s\S]*?</tr>", html))
    if len(trs) < 2:
        raise RuntimeError("V HTML VZOR_2 som nenasiel tabulku s aspon 2 riadkami.")
    header_end = trs[0].end()   # koniec hlavickoveho riadku
    last_end = trs[-1].end()    # koniec posledneho (povodneho) datoveho riadku
    pre = html[:header_end]
    post = html[last_end:]
    new_html = pre + "\n" + build_data_rows(rows) + "\n" + post
    return to_quoted_printable_cp1250(new_html)


def build_plain_body_bytes(rows) -> bytes:
    """Plain text telo cez spolocny generator zo stareho skriptu."""
    return to_quoted_printable_cp1250(build_plain_body(rows))


# ---------- attachment ----------

def build_attachment_block(filename: str, file_bytes: bytes, now: dt.datetime) -> bytes:
    """Attachment cast s --_004_ boundary, Q-encoded nazvom a base64 obsahom."""
    qname = q_encode_filename(filename)
    b64 = base64.encodebytes(file_bytes)
    b64 = b64.replace(b"\r\n", b"\n").replace(b"\n", b"\r\n")
    date_str = now.strftime("%a, %d %b %Y %H:%M:%S GMT")
    headers = (
        f"--{BOUNDARY_004}\r\n"
        f"Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;\r\n"
        f"\tname=\"{qname}\"\r\n"
        f"Content-Description: {qname}\r\n"
        f"Content-Disposition: attachment;\r\n"
        f"\tfilename=\"{qname}\"; size={len(file_bytes)};\r\n"
        f"\tcreation-date=\"{date_str}\";\r\n"
        f"\tmodification-date=\"{date_str}\"\r\n"
        f"Content-Transfer-Encoding: base64\r\n\r\n"
    ).encode("ascii")
    return headers + b64


# ---------- konstrukcia .eml ----------

def build_eml(vzor_bytes: bytes, *, subject: str, to_addr: str, rows,
              attachment_name: str, attachment_bytes: bytes) -> bytes:
    """Postavi cely .eml draft v style VZOR_2."""
    now = dt.datetime.now(dt.timezone.utc)
    subject_enc = encode_header_q(subject)
    msg_id = f"<{uuid.uuid4().hex.upper()}@GVUPR03MB11722.eurprd03.prod.outlook.com>"
    date_str = now.strftime("%a, %d %b %Y %H:%M:%S +0000")

    headers = (
        b"X-Unsent: 1\r\n"
        b"To: " + to_addr.encode("ascii") + b"\r\n"
        b"Subject: " + subject_enc + b"\r\n"
        b"Thread-Topic: " + subject_enc + b"\r\n"
        b"Date: " + date_str.encode("ascii") + b"\r\n"
        b"Message-ID:\r\n\t" + msg_id.encode("ascii") + b"\r\n"
        b"Content-Language: sk-SK\r\n"
        b"X-MS-Has-Attach: yes\r\n"
        b"X-MS-TNEF-Correlator:\r\n"
        b"X-MS-Exchange-Organization-RecordReviewCfmType: 0\r\n"
        b"msip_labels:\r\n"
        b"Content-Type: multipart/mixed;\r\n"
        b"\tboundary=\"" + BOUNDARY_004.encode() + b"\"\r\n"
        b"MIME-Version: 1.0\r\n"
    )

    alt_open = (
        f"--{BOUNDARY_004}\r\n"
        f"Content-Type: multipart/alternative;\r\n"
        f"\tboundary=\"{BOUNDARY_000}\"\r\n\r\n"
    ).encode("ascii")

    plain_headers = (
        f"--{BOUNDARY_000}\r\n"
        f"Content-Type: text/plain; charset=\"windows-1250\"\r\n"
        f"Content-Transfer-Encoding: quoted-printable\r\n\r\n"
    ).encode("ascii")
    plain_body = build_plain_body_bytes(rows)

    html_headers = (
        f"--{BOUNDARY_000}\r\n"
        f"Content-Type: text/html; charset=\"windows-1250\"\r\n"
        f"Content-Transfer-Encoding: quoted-printable\r\n\r\n"
    ).encode("ascii")
    html_body = build_html_body(vzor_bytes, rows)

    alt_close = f"--{BOUNDARY_000}--\r\n".encode("ascii")

    attach_block = build_attachment_block(attachment_name, attachment_bytes, now)
    mixed_close = f"--{BOUNDARY_004}--\r\n".encode("ascii")

    return (
        headers + b"\r\n" +
        alt_open +
        plain_headers + plain_body + b"\r\n" +
        html_headers + html_body + b"\r\n" +
        alt_close + b"\r\n" +
        attach_block + b"\r\n" +
        mixed_close
    )


# ---------- main ----------

def main():
    parser = argparse.ArgumentParser(description="Vyrobi .eml drafty podla VZOR_2.")
    parser.add_argument("--test", action="store_true", help="1 nahodny servis")
    parser.add_argument("--limit", type=int, default=None, help="Max pocet servisov")
    args = parser.parse_args()

    OUT_DIR.mkdir(parents=True, exist_ok=True)

    records = load_unikaty()
    fname_map = build_filename_map(records)
    vzor_bytes = VZOR2.read_bytes()

    if args.test:
        targets = random.sample(records, 1)
    elif args.limit is not None:
        targets = records[:args.limit]
    else:
        targets = records

    print(f"Spracovavam {len(targets)} servisov (mam emailov={len(records)})")
    skipped = []
    ok = 0
    for rec in targets:
        h = rec["h"]
        nazov = rec["nazov"]
        base_fname = fname_map[h]
        xlsx_path = SERVIS_DIR / f"{base_fname}.xlsx"
        if not xlsx_path.exists():
            print(f"  SKIP (chyba xlsx): {xlsx_path.name}")
            skipped.append(base_fname)
            continue

        rows = read_xlsx_rows(xlsx_path)
        attachment_bytes = xlsx_path.read_bytes()
        attachment_name = f"{base_fname}.xlsx"

        subject = SUBJECT_PREFIX + nazov
        to_parts = [rec["email1"]]
        if rec["email2"]:
            to_parts.append(rec["email2"])
        to_addr = ", ".join(to_parts)

        new_eml = build_eml(
            vzor_bytes,
            subject=subject,
            to_addr=to_addr,
            rows=rows,
            attachment_name=attachment_name,
            attachment_bytes=attachment_bytes,
        )

        out_name = sanitize_filename(f"{SUBJECT_PREFIX}{nazov}") + ".eml"
        out_path = OUT_DIR / out_name
        out_path.write_bytes(new_eml)
        ok += 1
        print(f"  OK {out_name}  ({len(rows)} riadkov, {len(attachment_bytes)} B priloha, to={to_addr})")

    print(f"\nHotovo: {ok} OK, {len(skipped)} preskocenych")
    print(f"Vystup: {OUT_DIR}")


if __name__ == "__main__":
    main()

"""
Vyrobi koncepty (drafts) v Luckinom firemnom Outlook Web (OWA) cez Playwright.

Tato verzia bezi na Marcelovom PC (kde sa da slobodne instalovat). Po prvom
spusteni sa otvori Chrome a Lucka sa fyzicky prihlasi (heslo + 2FA) do
outlook.office365.com. Session sa ulozi do priecinka `owa_session/`, takze
dalsie spustenia uz prihlasenie nevyzaduju.

Skript NIKDY neodosiela. Vsetky maily konci v priecinku Koncepty (Drafts).
Lucka ich v svojom korporatnom Outlooku skontroluje a klikne Send.

Pouzitie:
    pip install playwright openpyxl
    playwright install chromium

    python vyrob_maily_owa.py --test            # 1 nahodny servis
    python vyrob_maily_owa.py --limit 5         # prvych 5 servisov
    python vyrob_maily_owa.py                   # vsetky servisy s mailom
    python vyrob_maily_owa.py --reset-session   # vymaze ulozeny login
"""
import argparse
import quopri
import random
import re
import shutil
import sys
import time
from pathlib import Path

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# Reuse funkcie zo skriptu vyrob_maily.py
sys.path.insert(0, str(Path(__file__).parent))
from vyrob_maily import (  # noqa: E402
    ROOT, VZOR, SERVIS_DIR, SUBJECT_PREFIX,
    load_unikaty, build_filename_map, read_xlsx_rows,
    extract_html_template, build_html_data_rows,
)
# Kompaktne HTML telo (margin:0, skipProofing tabulka) berieme z VZOR_2 .eml
# ako sablonu — vyzera lepsie pri vlozeni do OWA nez stare MsoNormal HTML.
from vyrob_maily_v2 import (  # noqa: E402
    VZOR2, _decode_part as _v2_decode_part,
    build_data_rows as _v2_build_data_rows,
)
import base64 as _base64  # noqa: E402

LOGO_PATH = ROOT / "_logo.png"

USER_DATA_DIR = ROOT / "owa_session"
OWA_URL = "https://outlook.office365.com/mail/"

# Hostname-y, na ktorych moze bezat prihlaseny Outlook Web. Microsoft v 2026
# migroval OWA z outlook.office365.com / outlook.office.com na novu domenu
# outlook.cloud.microsoft, takze detekcia prihlasenia musi pokryt vsetky.
OUTLOOK_HOSTS = ("outlook.office", "outlook.cloud.microsoft", "outlook.live")


# ---------- HTML telo pre OWA (bez quoted-printable, iba obsah <body>) ----------

def _logo_data_uri() -> str:
    """Vrati _logo.png ako data:image/png;base64 URI (na vlozenie do <img src>)."""
    raw = LOGO_PATH.read_bytes()
    b64 = _base64.b64encode(raw).decode("ascii")
    return f"data:image/png;base64,{b64}"


def build_html_for_owa(vzor_bytes: bytes, rows) -> str:
    """Vrati obsah <body> kompaktneho HTML tela (UTF-8 string) na vlozenie do
    OWA compose editora cez clipboard.

    Telo berieme z VZOR_2 .eml ako sablonu — ma cisty 'skipProofing' format
    s margin:0 (ziadne velke medzery ako stare MsoNormal HTML). Datove riadky
    tabulky nahradime novymi a logo (povodne nefunkcny odkaz) vlozime ako
    data URI z _logo.png, aby sa zobrazilo aj po vlozeni cez schranku.
    Parameter `vzor_bytes` (stary vzor) sa ignoruje.
    """
    html = _v2_decode_part(VZOR2.read_bytes(), b"text/html")
    trs = list(re.finditer(r"<tr[\s\S]*?</tr>", html))
    if len(trs) < 2:
        raise RuntimeError("V HTML sablone (VZOR_2) som nenasiel tabulku.")
    header_end = trs[0].end()   # koniec hlavickoveho riadku tabulky
    last_end = trs[-1].end()    # koniec posledneho (povodneho) datoveho riadku
    html = html[:header_end] + "\n" + _v2_build_data_rows(rows) + "\n" + html[last_end:]

    # Logo: nahrad src vsetkych <img> za data URI z _logo.png.
    data_uri = _logo_data_uri()
    html = re.sub(
        r'(<img\b[^>]*\bsrc=")[^"]*(")',
        lambda m: m.group(1) + data_uri + m.group(2),
        html, flags=re.IGNORECASE,
    )

    body_m = re.search(r'<body[^>]*>([\s\S]*?)</body>', html, re.IGNORECASE)
    return body_m.group(1).strip() if body_m else html


# ---------- OWA selektory (viacjazycne) ----------

NEW_MAIL_BTN = ", ".join([
    'button[aria-label="New mail"]',
    'button[aria-label="Nová pošta"]',
    'button[aria-label="Nová správa"]',
    'button[aria-label="Nový e-mail"]',
    'button[aria-label="Nový email"]',
    'button[aria-label="New Email"]',
    'button[aria-label="Nová položka"]',
])

TO_FIELD = ", ".join([
    'div[aria-label="To"]',
    'div[aria-label="Komu"]',
    'input[aria-label="To"]',
    'input[aria-label="Komu"]',
    '[role="combobox"][aria-label="To"]',
    '[role="combobox"][aria-label="Komu"]',
])

SUBJECT_FIELD = ", ".join([
    'input[aria-label="Add a subject"]',
    'input[aria-label*="Predmet"]',
    'input[aria-label*="Pridať predmet"]',
    'input[aria-label*="Add subject"]',
])

BODY_FIELD = ", ".join([
    '[role="textbox"][aria-label*="Message body" i]',
    '[role="textbox"][aria-label*="Telo správy" i]',
    '[role="textbox"][aria-label*="Telo zprávy" i]',
    '[role="textbox"][aria-label*="message body" i]',
    'div[contenteditable="true"][aria-label*="body" i]',
    'div[contenteditable="true"][aria-label*="telo" i]',
])

ATTACH_BTN = ", ".join([
    'button[aria-label*="Priložiť" i]',
    'button[aria-label*="Prilozit" i]',
    'button[aria-label*="Attach" i]',
    'button[aria-label*="Pripojiť" i]',
    'button[aria-label*="Připojit" i]',
    'button[aria-label*="Insert" i]',
])

# Tlacidlo na ZATVORENIE compose, ktore PONECHA koncept (OWA ho auto-ulozi).
# Pozor: 'Zahodiť' (== Esc) koncept ZMAZE do kosa, preto ho nepouzivame na
# ulozenie, iba pri zotaveni z chyby (zrusenie nepodareneho pokusu).
# Pozor: v DOM su DVE tlacidla "Zavriet" — jedno skryte (display:none, z ineho
# panelu) a jedno viditelne v okne spravy. Preto `:visible`, inak `.first`
# chyti to skryte a click vyprsi (presne to zhodilo davku 5).
CLOSE_BTN = ", ".join([
    'button[aria-label="Zavrieť"]:visible',
    'button[aria-label="Zavřít"]:visible',
    'button[aria-label="Close"]:visible',
])
DISCARD_BTN = ", ".join([
    'button[aria-label="Zahodiť"]:visible',
    'button[aria-label="Zahodit"]:visible',
    'button[aria-label="Discard"]:visible',
])

BROWSE_LOCAL_TEXTS = [
    "Prehľadávať tento počítač",
    "Prehľadávať tento počítač…",
    "Tento počítač",
    "Browse this computer",
    "Procházet tento počítač",
    "Browse computer",
    "Z tohto zariadenia",
]


# ---------- Playwright kroky ----------

def _find_outlook_page(context):
    """Najde otvorenu page, ktora je na outlook.office.com / .office365.com domene.
    Vracia None ak ziadna nie je."""
    for p in context.pages:
        try:
            url = p.url
        except Exception:
            continue
        if any(h in url for h in OUTLOOK_HOSTS) and "/mail" in url:
            return p
    return None


def wait_for_login(context, page, timeout_ms: int = 10 * 60_000):
    """Pocka, kym je Lucka prihlasena a vidi inbox.

    Sleduje VSETKY otvorene taby (OWA pri SSO presmerovanie obcas otvori novy
    tab). Po nastavenom intervale skenuje context.pages a hlada page na
    outlook.office.com s viditelnym New mail buttonom.

    Vracia tu stranku, ktora je pripravena (moze byt ina nez vstupna `page`).
    """
    print("Cakam na prihlasenie do Outlook Web (max 10 min)...")
    print("V otvorenom Chrome okne sa prihlas, vratane 2FA na mobile.")
    deadline = time.time() + timeout_ms / 1000.0
    last_url = None
    while time.time() < deadline:
        # Snazi sa najst outlook page
        ow_page = _find_outlook_page(context)
        if ow_page is not None:
            try:
                btn = ow_page.locator(NEW_MAIL_BTN).first
                if btn.is_visible(timeout=2000):
                    print(f"OK, som prihlaseny. (page: {ow_page.url})")
                    return ow_page
            except Exception:
                pass
        # Diagnostika: ak sa URL zmenila, vypisi ju
        try:
            cur_url = page.url
            if cur_url != last_url:
                print(f"  ... aktualna URL: {cur_url}")
                last_url = cur_url
        except Exception:
            pass
        time.sleep(2)
    raise TimeoutError("Login sa nepodaril v limite 10 min.")


def open_new_compose(page):
    """Klikne New mail a pocka, kym sa compose PLNE nacita (telo aj pole Komu).

    Pole Komu sa najma pri prvom otvoreni renderuje pomalsie nez telo, preto
    explicitne cakame aj na nu — inak fill_to vyprsi (stalo sa pri 1. maile).
    """
    page.locator(NEW_MAIL_BTN).first.click()
    page.wait_for_selector(BODY_FIELD, timeout=20_000)
    try:
        page.locator(TO_FIELD).first.wait_for(state="visible", timeout=20_000)
    except Exception:
        pass
    time.sleep(1.0)


def fill_to(page, to_addr: str):
    """Vyplni pole Komu. to_addr moze obsahovat viac adries oddelenych ; alebo ,."""
    to_locator = page.locator(TO_FIELD).first
    to_locator.wait_for(state="visible", timeout=20_000)
    to_locator.click()
    time.sleep(0.2)
    emails = [e.strip() for e in re.split(r"[;,]", to_addr) if e.strip()]
    for email in emails:
        page.keyboard.type(email, delay=10)
        time.sleep(0.3)
        page.keyboard.press("Enter")
        time.sleep(0.3)


def fill_subject(page, subject: str):
    subj = page.locator(SUBJECT_FIELD).first
    subj.click()
    time.sleep(0.1)
    try:
        subj.fill(subject)
    except Exception:
        page.keyboard.type(subject, delay=5)


def fill_body_html(page, html: str):
    """Vlozi HTML do compose body cez clipboard paste (zachova formatovanie)."""
    body = page.locator(BODY_FIELD).first
    body.click()
    time.sleep(0.3)
    # Polož HTML do clipboardu (text/html) cez JS
    page.evaluate(
        """async (html) => {
            const blob = new Blob([html], {type: 'text/html'});
            const item = new ClipboardItem({'text/html': blob});
            await navigator.clipboard.write([item]);
        }""",
        html,
    )
    page.keyboard.press("Control+V")
    time.sleep(1.0)


def _attachment_appeared(page, filename: str, timeout_s: float = 8.0) -> bool:
    """Vrati True, ked sa v compose objavi priloha s danym nazvom suboru."""
    # OWA nazov prilohy v UI obcas skrati, preto hladame prefix nazvu suboru.
    needle = Path(filename).stem[:10]
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        try:
            cnt = page.evaluate(
                """(name) => {
                    const t = document.body.innerText || '';
                    return t.includes(name) ? 1 : 0;
                }""",
                needle,
            )
            if cnt:
                return True
        except Exception:
            pass
        time.sleep(0.4)
    return False


def attach_file(page, file_path: Path):
    """Pripne lokalny subor. Primarne nahra priamo do skryteho
    <input type=file> (nove OWA na outlook.cloud.microsoft), co je
    spolahlivejsie nez preklikavanie menu. Ako fallback pouzije tlacidlo
    'Prilozit subor' -> 'Prehladavat tento pocitac' s file chooserom."""
    # 1) Priamy upload do univerzalneho file inputu (preskoc input pre obrazky).
    inputs = page.locator('input[type="file"]')
    try:
        n = inputs.count()
    except Exception:
        n = 0
    for i in range(n):
        inp = inputs.nth(i)
        try:
            accept = (inp.get_attribute("accept") or "").lower()
        except Exception:
            accept = ""
        if "image/" in accept:
            continue  # input urceny len pre obrazky preskocime
        try:
            inp.set_input_files(str(file_path))
        except Exception:
            continue
        if _attachment_appeared(page, file_path.name):
            return

    # 2) Fallback: tlacidlo Prilozit -> menu -> Prehladavat tento pocitac.
    page.locator(ATTACH_BTN).first.click()
    time.sleep(0.5)
    browse_locator = None
    for txt in BROWSE_LOCAL_TEXTS:
        cand = page.locator(
            f'[role="menuitem"]:has-text("{txt}"), button:has-text("{txt}")')
        if cand.count() > 0:
            browse_locator = cand.first
            break
    if browse_locator is None:
        raise RuntimeError(
            "Prilohu sa nepodarilo pripnut: ani priamy input, ani menu polozka "
            f"'Prehladavat tento pocitac' (skusane: {BROWSE_LOCAL_TEXTS})"
        )
    with page.expect_file_chooser() as fc_info:
        browse_locator.click()
    fc_info.value.set_files(str(file_path))
    if not _attachment_appeared(page, file_path.name, timeout_s=10.0):
        raise RuntimeError("Priloha sa po nahrati neobjavila v compose okne.")


def save_and_close_compose(page):
    """Zatvori compose cez tlacidlo 'Zavriet', co PONECHA koncept (OWA ho
    automaticky ulozi do priecinka Koncepty).

    NEPOUZIVAME Escape — v novom OWA Esc znamena 'Zahodit' a koncept by
    skoncil v kosi (presne to sa stalo pri prvych testoch).
    """
    time.sleep(2.0)  # necha OWA auto-ulozit rozpracovany koncept
    page.locator(CLOSE_BTN).first.click(timeout=15_000)
    # Pocka, kym sa compose zatvori (telo spravy zmizne z DOM).
    try:
        page.wait_for_selector(BODY_FIELD, state="detached", timeout=8_000)
    except Exception:
        pass
    time.sleep(0.8)


def compose_one_draft(page, *, to_addr, subject, html_body, attachment_path: Path):
    open_new_compose(page)
    fill_to(page, to_addr)
    fill_subject(page, subject)
    fill_body_html(page, html_body)
    attach_file(page, attachment_path)
    save_and_close_compose(page)


# ---------- main ----------

def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[1])
    parser.add_argument("--test", action="store_true", help="1 nahodny servis")
    parser.add_argument("--sample", type=int, default=None, help="N nahodnych servisov")
    parser.add_argument("--limit", type=int, default=None, help="Prvych N servisov")
    parser.add_argument(
        "--only", type=str, default=None,
        help="Len servisy s tymito nazvami (oddelene bodkociarkou ;)")
    parser.add_argument(
        "--reset-session", action="store_true",
        help="Vymaze owa_session/ a vynuti novy login",
    )
    args = parser.parse_args()

    if args.reset_session and USER_DATA_DIR.exists():
        shutil.rmtree(USER_DATA_DIR)
        print(f"Vymazany session priecinok: {USER_DATA_DIR}")

    USER_DATA_DIR.mkdir(parents=True, exist_ok=True)

    records = load_unikaty()
    fname_map = build_filename_map(records)
    vzor_bytes = VZOR.read_bytes()

    if args.test:
        targets = random.sample(records, 1)
    elif args.only:
        wanted = [s.strip() for s in args.only.split(";") if s.strip()]
        targets = [r for r in records if r["nazov"] in wanted]
        missing = [w for w in wanted if w not in {r["nazov"] for r in targets}]
        if missing:
            print(f"POZOR: tieto nazvy som v datach nenasiel: {missing}")
    elif args.sample is not None:
        targets = random.sample(records, min(args.sample, len(records)))
    elif args.limit is not None:
        targets = records[:args.limit]
    else:
        targets = records

    print(f"Spracovavam {len(targets)} servisov (mam emailov={len(records)})")

    with sync_playwright() as p:
        # channel="chrome" -> pouzije systemovy Chrome Stable (nie Chrome for Testing).
        # ignore_default_args vypina --enable-automation flag, ktory Microsoft
        # detekuje a blokuje. add_init_script skryva navigator.webdriver.
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(USER_DATA_DIR),
            channel="chrome",
            headless=False,
            viewport={"width": 1400, "height": 900},
            permissions=["clipboard-read", "clipboard-write"],
            args=["--disable-blink-features=AutomationControlled"],
            ignore_default_args=["--enable-automation"],
        )
        context.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', "
            "{get: () => undefined});"
        )
        page = context.pages[0] if context.pages else context.new_page()
        page.goto(OWA_URL)
        page = wait_for_login(context, page)

        ok = 0
        skipped = []
        failed = []
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
            html_body = build_html_for_owa(vzor_bytes, rows)
            subject = SUBJECT_PREFIX + nazov
            to_parts = [rec["email1"]]
            if rec["email2"]:
                to_parts.append(rec["email2"])
            to_addr = "; ".join(to_parts)

            try:
                compose_one_draft(
                    page,
                    to_addr=to_addr,
                    subject=subject,
                    html_body=html_body,
                    attachment_path=xlsx_path,
                )
                ok += 1
                print(f"  OK {nazov}  ({len(rows)} riadkov, to={to_addr})")
            except (PWTimeout, Exception) as e:
                print(f"  FAIL {nazov}: {type(e).__name__}: {e}")
                failed.append(nazov)
                # Zotavenie: nepodareny pokus ZAHODIME (Zahodit/Esc), aby
                # nezostal otvoreny a neblokoval dalsi servis.
                try:
                    disc = page.locator(DISCARD_BTN).first
                    if disc.is_visible(timeout=2000):
                        disc.click()
                        time.sleep(0.4)
                        # potvrdenie zahodenia, ak sa objavi
                        try:
                            page.keyboard.press("Enter")
                        except Exception:
                            pass
                    else:
                        page.keyboard.press("Escape")
                    time.sleep(0.6)
                except Exception:
                    try:
                        page.keyboard.press("Escape")
                        time.sleep(0.5)
                    except Exception:
                        pass

        print(f"\nHotovo: {ok} OK, {len(skipped)} preskocenych, {len(failed)} chyb")
        if skipped:
            print("Preskoceni (chybajuci xlsx):")
            for s in skipped:
                print(f"  - {s}")
        if failed:
            print("Chyby pri OWA orchestracii:")
            for f in failed:
                print(f"  - {f}")
        print("\nKoncepty su v Outlook Web priecinku 'Koncepty' (Drafts).")
        print("Stlac Enter pre zatvorenie prehliadaca...")
        try:
            input()
        except EOFError:
            pass
        context.close()


if __name__ == "__main__":
    main()

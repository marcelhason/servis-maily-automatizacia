"""
Stiahne odpovede servisov z Luckinho Outlook Web (OWA) cez Playwright.

Odpovede chodia (pravidlo v schránke) do priečinka
Doručená pošta / 1.5.25 / DBPSERVIS — skript ide priamo doň.

BEZPEČNOSŤ (overené testom 2026-06-11): schránka sa NEMENÍ. Maily sa
otvárajú pod filtrom „Neprečítané" — Eva má zapnuté nastavenie „Vo filtri
neprečítaných položiek: položky vždy ponechať ako neprečítané", takže
otvorenie mail neoznačí. Po každom maile sa overí, že položka zostala
neprečítaná; ak nie, skript OKAMŽITE skončí. Filter je len zobrazovací
a pri novej relácii sa resetuje sám. Nič sa nemaže ani neoznačuje.

Iba SŤAHUJE surové artefakty (parsovanie rieši spracuj_odpovede.py)
do `odpovede/<mail_id>/`:

    meta.json   — odosielateľ (s emailom), predmet, label, kedy stiahnuté
    telo.txt    — viditeľné telá správ konverzácie (naša pôvodná žiadosť
                  je v konverzácii zbalená, do textu sa nedostane)
    *.xlsx, ... — prílohy (cez menu „Ďalšie akcie" -> „Stiahnuť")

Existujúci priečinok `odpovede/<mail_id>/` sa preskočí => idempotentné.
Pozn.: mail_id je odvodený z ID konverzácie — ak do konverzácie neskôr
pribudne ďalšia odpoveď, treba priečinok zmazať, aby sa stiahla nanovo.

Použitie:
    python stiahni_odpovede_owa.py              # všetky neprečítané odpovede
    python stiahni_odpovede_owa.py --limit 3    # prvé 3 (na test)
"""
import argparse
import hashlib
import json
import re
import sys
import time
from datetime import date
from pathlib import Path

from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
from vyrob_maily import ROOT  # noqa: E402
from vyrob_maily_owa import (  # noqa: E402
    OWA_URL, launch_owa_context, wait_for_login,
)

ODPOVEDE_DIR = ROOT / "odpovede"
FOLDER = "DBPSERVIS"

# ---------- selektory (overené na outlook.cloud.microsoft, 2026-06) ----------

# Treeitemy priečinkov nemajú aria-label -> hľadáme podľa textu.
FOLDER_ITEM = f'[role="treeitem"]:has-text("{FOLDER}")'
FILTER_BTN = "#mailListFilterMenu"
FILTER_NEPRECITANE = '[role="menuitemradio"]:has-text("Neprečítané")'

# Položky zoznamu správ: pod filtrom Neprečítané začína ich aria-label
# „Neprečítané..." — zároveň tým NEchytíme karty príloh v reading pane
# (tie sú tiež [role=option], ale label majú „subor.xlsx Otvoriť 9 kB").
MSG_ITEM = '[role="option"][aria-label^="Neprečítan"]'

# Telo správy: BEZ :visible by sa chytali staré skryté panely z predtým
# otvorených mailov (OWA ich necháva v DOM-e).
VIDITELNE_TELO = '[role="document"]:visible'

PRILOHA_OPT = ('[role="listbox"][aria-label*="súbory" i] '
               '[role="option"]:visible')
DALSIE_AKCIE = 'button[aria-label="Ďalšie akcie"]'
MENU_STIAHNUT = ('[role="menuitem"]:has-text("Stiahnuť"), '
                 '[role="menuitemradio"]:has-text("Stiahnuť"), '
                 'button:has-text("Stiahnuť")')
# Fallback pre obrázky (ich menu „Ďalšie akcie" položku Stiahnuť nemá):
# tlačidlo v hlavičke zoznamu príloh stiahne všetko ako jeden zip —
# priprav_klasifikaciu.py ho potom rozbalí do _rozbalene/.
# Selektor je best-effort odhad, overí sa pri ďalšom ostrom behu.
STIAHNUT_VSETKO = ('button[aria-label*="Stiahnuť všetk" i]:visible, '
                   'button:has-text("Stiahnuť všetko"):visible')

# Dátum prijatia správy — v konverzácii jeden per správa („Štv 11. 6. 2026
# 13:10"), odpoveď servisu je ten najnovší.
DATUM_EL = '[data-testid="SentReceivedSavedTime"]'
_DATUM_RE = re.compile(r"(\d{1,2})\.\s*(\d{1,2})\.\s*(\d{4})"
                       r"(?:\s+(\d{1,2}):(\d{2}))?")

# Po otvorení mailu má URL tvar .../mail/<id priečinka>/id/<convid
# url-encoded> — slúži ako deep link späť na konverzáciu (klik naň mimo
# filtra Neprečítané mail označí za prečítaný, čo je pri vybavovaní žiaduce).


# ---------- kroky ----------

def otvor_priecinok_s_filtrom(page):
    page.locator(FOLDER_ITEM).first.click()
    page.wait_for_selector('[role="option"]', timeout=20_000)
    time.sleep(2)
    page.locator(FILTER_BTN).click()
    time.sleep(1.0)
    page.locator(FILTER_NEPRECITANE).first.click()
    time.sleep(2.0)


def precitaj_datum(page) -> tuple[str, str]:
    """(ISO dátum, surový text) najnovšej správy konverzácie — t.j. odpovede
    servisu (naša žiadosť je staršia). Prázdne, ak sa nenašiel."""
    texty = page.eval_on_selector_all(
        DATUM_EL, 'els => els.map(e => (e.textContent || "").trim())')
    najlepsi = ("", "")
    for t in texty:
        m = _DATUM_RE.search(t)
        if not m:
            continue
        den, mesiac, rok, hod, minuta = m.groups()
        iso = f"{rok}-{int(mesiac):02d}-{int(den):02d}"
        kluc = iso + (f" {int(hod):02d}:{minuta}" if hod else "")
        if kluc > najlepsi[0]:
            najlepsi = (kluc, t)
    return (najlepsi[0][:10], najlepsi[1])


def rozdel_odosielatela(s: str) -> tuple[str, str]:
    """„Meno<email>" -> (meno, email)."""
    m = re.match(r"(.*?)<([^<>@\s]+@[^<>\s]+)>", s or "")
    return (m.group(1).strip(), m.group(2).strip()) if m else (s or "", "")


def precitaj_mail(page) -> dict:
    """Z otvoreného reading panelu vráti viditeľné telá, predmet,
    odosielateľa (headingy obsahujú „Meno<email>"), dátum prijatia
    a deep link URL."""
    tela = page.locator(VIDITELNE_TELO)
    texty = [tela.nth(i).inner_text() for i in range(tela.count())]
    headings = page.eval_on_selector_all(
        '[role="main"] [role="heading"]',
        """els => els.filter(e => e.offsetParent)
                .map(e => (e.innerText || "").replace(/\\s+/g, " ").trim())
                .filter(t => t)""")
    odosielatel = next((h for h in headings if "@" in h), "")
    meno, email = rozdel_odosielatela(odosielatel)
    datum_iso, datum_text = precitaj_datum(page)
    return {
        "predmet": headings[0] if headings else "",
        "odosielatel": odosielatel,
        "odosielatel_meno": meno,
        "odosielatel_email": email,
        "datum_prijatia": datum_iso,
        "prijate_text": datum_text,
        "owa_url": page.url if "/id/" in page.url else "",
        "telo": "\n\n--- ďalšia správa v konverzácii ---\n\n".join(texty),
    }


def stiahni_prilohy(page, out_dir: Path) -> tuple[list[str], list[str]]:
    """Stiahne prílohy otvoreného mailu cez menu karty prílohy
    „Ďalšie akcie" -> „Stiahnuť". Kliká VÝHRADNE na položku Stiahnuť.
    Pri zlyhaní karty (typicky obrázky — ich menu Stiahnuť nemá) skúsi
    na záver „Stiahnuť všetko" (zip). Vráti (stiahnuté, nestiahnuté
    labely) — nestiahnuté sa zapisujú do meta.json, klasifikátor podľa
    nich vie, že mailu chýbajú obrázky."""
    nazvy, nestiahnute = [], []
    karty = page.locator(PRILOHA_OPT)
    for i in range(karty.count()):
        karta = karty.nth(i)
        label = karta.get_attribute("aria-label") or ""
        try:
            karta.locator(DALSIE_AKCIE).first.click(timeout=8_000)
            time.sleep(0.8)
            with page.expect_download(timeout=20_000) as dl_info:
                page.locator(MENU_STIAHNUT).first.click(timeout=8_000)
            dl = dl_info.value
            dl.save_as(str(out_dir / dl.suggested_filename))
            nazvy.append(dl.suggested_filename)
            time.sleep(0.5)
        except Exception as e:
            menu_texty = []
            try:
                menu_texty = page.eval_on_selector_all(
                    '[role="menuitem"], [role="menuitemradio"]',
                    'els => els.map(e => (e.innerText||"").trim())'
                    '.filter(t => t)')
            except Exception:
                pass
            print(f"    priloha '{label[:50]}': nepodarilo sa stiahnuť "
                  f"({type(e).__name__}; menu: {menu_texty[:6]})")
            nestiahnute.append(label[:120])
            # menu zatvor opätovným klikom na tlačidlo — Escape by zrušil
            # výber položky v zozname a rozbil navigáciu
            try:
                karta.locator(DALSIE_AKCIE).first.click(timeout=3_000)
            except Exception:
                pass

    if nestiahnute:
        try:
            with page.expect_download(timeout=30_000) as dl_info:
                page.locator(STIAHNUT_VSETKO).first.click(timeout=8_000)
            dl = dl_info.value
            meno = dl.suggested_filename
            if not meno.lower().endswith(".zip"):
                meno += ".zip"
            dl.save_as(str(out_dir / meno))
            nazvy.append(meno)
            print(f"    fallback Stiahnuť všetko: {meno} "
                  f"(pokrýva {len(nestiahnute)} nestiahnutých)")
            nestiahnute = []
            time.sleep(0.5)
        except Exception as e:
            print(f"    fallback Stiahnuť všetko zlyhal "
                  f"({type(e).__name__}) — {len(nestiahnute)} príloh "
                  "zostáva nestiahnutých")
    return nazvy, nestiahnute


def dopln_meta(page, out_dir: Path) -> bool:
    """Backfill: do meta.json už stiahnutého mailu doplní datum_prijatia,
    owa_url a rozdeleného odosielateľa, ak chýbajú. Mail je v tej chvíli
    otvorený (klikli sme naň kvôli posunu zoznamu). True = niečo doplnené."""
    meta_path = out_dir / "meta.json"
    meta = json.loads(meta_path.read_text(encoding="utf-8"))
    if meta.get("datum_prijatia") and meta.get("owa_url"):
        return False
    page.locator(VIDITELNE_TELO).first.wait_for(state="visible",
                                                timeout=20_000)
    time.sleep(1.5)
    datum_iso, datum_text = precitaj_datum(page)
    meno, email = rozdel_odosielatela(meta.get("odosielatel", ""))
    meta.update({
        "datum_prijatia": datum_iso,
        "prijate_text": datum_text,
        "owa_url": page.url if "/id/" in page.url else "",
        "odosielatel_meno": meno,
        "odosielatel_email": email,
    })
    meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=1),
                         encoding="utf-8")
    return True


def zostal_neprecitany(page, convid: str) -> bool:
    try:
        lab = page.locator(
            f'[role="option"][data-convid="{convid}"]'
        ).first.get_attribute("aria-label", timeout=5_000) or ""
        return lab.startswith("Neprečítan")
    except Exception:
        # položka mohla vypadnúť z virtualizovaného zoznamu — neblokuj,
        # ale povedz to
        print("    (položku sa nepodarilo znova nájsť na kontrolu)")
        return True


# ---------- main ----------

def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[1])
    parser.add_argument("--limit", type=int, default=None,
                        help="Spracuje len prvých N mailov")
    args = parser.parse_args()

    ODPOVEDE_DIR.mkdir(exist_ok=True)
    dnes = date.today().isoformat()

    with sync_playwright() as p:
        context = launch_owa_context(p)
        page = context.pages[0] if context.pages else context.new_page()
        page.goto(OWA_URL)
        page = wait_for_login(context, page)
        time.sleep(2)

        otvor_priecinok_s_filtrom(page)
        print(f"Priečinok {FOLDER}, filter Neprečítané aktívny.")

        nove, preskocene, doplnene, chyby = 0, 0, 0, 0
        videne = set()
        stagnacia = 0

        while stagnacia < 4:
            if args.limit and nove >= args.limit:
                break

            # nájdi prvú nespracovanú NAČÍTANÚ položku
            items = page.locator(MSG_ITEM)
            target, convid, label = None, None, None
            for i in range(items.count()):
                it = items.nth(i)
                c = it.get_attribute("data-convid") or ""
                if c and c not in videne:
                    target, convid = it, c
                    label = it.get_attribute("aria-label") or ""
                    break

            if target is None:
                # nič nenačítané — posuň viewport: klikni na poslednú
                # položku a šípkou dole donačítaj (klik = len výber,
                # pod filtrom Neprečítané bezpečný)
                stagnacia += 1
                try:
                    items.last.click(timeout=5_000)
                    time.sleep(0.5)
                    for _ in range(5):
                        page.keyboard.press("ArrowDown")
                        time.sleep(0.6)
                except Exception:
                    break
                continue
            stagnacia = 0
            videne.add(convid)

            # klikni VŽDY (aj keď bude skip) — posúva viewport zoznamu
            target.click()
            time.sleep(1.0)

            mid = hashlib.sha1(convid.encode()).hexdigest()[:16]
            out_dir = ODPOVEDE_DIR / mid
            if out_dir.exists():
                try:
                    if dopln_meta(page, out_dir):
                        doplnene += 1
                        print(f"  META [{mid}] doplnený dátum prijatia "
                              "+ odkaz")
                    else:
                        preskocene += 1
                except Exception as e:
                    chyby += 1
                    print(f"  FAIL META [{mid}] {type(e).__name__}: {e}")
                time.sleep(1.0)
                if not zostal_neprecitany(page, convid):
                    print("\nSTOP! Mail už nie je neprečítaný — okamžite "
                          "končím. Over nastavenie filtra!")
                    break
                continue

            try:
                page.locator(VIDITELNE_TELO).first.wait_for(
                    state="visible", timeout=20_000)
                time.sleep(2.0)
                obsah = precitaj_mail(page)
                out_dir.mkdir()
                (out_dir / "telo.txt").write_text(
                    obsah["telo"], encoding="utf-8")
                prilohy, nestiahnute = (
                    stiahni_prilohy(page, out_dir)
                    if "Obsahuje prílohy" in label else ([], []))
                (out_dir / "meta.json").write_text(json.dumps({
                    "mail_id": mid,
                    "convid": convid,
                    "predmet": obsah["predmet"],
                    "odosielatel": obsah["odosielatel"],
                    "odosielatel_meno": obsah["odosielatel_meno"],
                    "odosielatel_email": obsah["odosielatel_email"],
                    "datum_prijatia": obsah["datum_prijatia"],
                    "prijate_text": obsah["prijate_text"],
                    "owa_url": obsah["owa_url"],
                    "aria_label": label,
                    "prilohy": prilohy,
                    "prilohy_nestiahnute": nestiahnute,
                    "stiahnute": dnes,
                }, ensure_ascii=False, indent=1), encoding="utf-8")
                nove += 1
                print(f"  OK [{mid}] {label[:80]}  (priloh: {len(prilohy)})")
            except Exception as e:
                chyby += 1
                print(f"  FAIL [{mid}] {type(e).__name__}: {e}")
                if out_dir.exists() and not any(out_dir.iterdir()):
                    out_dir.rmdir()

            # bezpečnostná kontrola: mail musí zostať neprečítaný
            time.sleep(1.5)
            if not zostal_neprecitany(page, convid):
                print("\nSTOP! Mail už nie je neprečítaný — okamžite končím, "
                      "ďalšie maily neotváram. Over nastavenie filtra!")
                break

        print(f"\nHotovo: {nove} nových, {doplnene} doplnených meta, "
              f"{preskocene} už kompletných, {chyby} chýb. "
              f"Artefakty: {ODPOVEDE_DIR}")
        context.close()


if __name__ == "__main__":
    main()

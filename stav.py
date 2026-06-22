"""
Stavový sklad reklamácií — perzistentný JSON `stav_reklamacii.json`.

Pravdou o stave komunikácie so servismi je tento súbor, NIE excel (ten sa
generuje ako pohľad cez spracuj_odpovede.py). Komunikácia býva viackolová,
preto sa poznámky iba PRIDÁVAJÚ (append-only, s dátumom), nikdy neprepisujú.

Stavový automat per hlásenie (pole `stav` + odvodené `on_turn`):

    odoslane   (on_turn=servis)  — žiadosť odoslaná, čakáme na servis
    odpovedane (on_turn=my)      — prišla spárovaná odpoveď, loptička u nás
    uzavrete   (on_turn=nikto)   — uzavreté ručne (--uzavri)

Flag „ručná kontrola" je na úrovni SERVISU (nespárované telo mailu),
nie stav hlásenia.

Použitie:
    python stav.py --init             # prvotné naplnenie zo servis_subory/
    python stav.py --init --force     # prepíše existujúci stavový súbor
    python stav.py --prehlad          # súhrn stavu na konzolu
    python stav.py --uzavri "8803000002;8803000001"
"""
import argparse
import json
import sys
from pathlib import Path

STAV_XLSX = Path(__file__).resolve().parent / "dpb servis 180-360 - stav.xlsx"
_ZELENA_RGB = "FF92D050"

sys.path.insert(0, str(Path(__file__).parent))
from vyrob_maily import (  # noqa: E402
    ROOT, SERVIS_DIR, load_unikaty, build_filename_map, read_xlsx_rows,
)

STAV_PATH = ROOT / "stav_reklamacii.json"
DATUM_ZIADOSTI = "2026-06-10"  # deň, keď Eva rozoslala žiadosti

STAVY = ("odoslane", "odpovedane", "uzavrete")
ON_TURN = {"odoslane": "servis", "odpovedane": "my", "uzavrete": "nikto"}


# ---------- načítanie / uloženie ----------

def load_stav() -> dict:
    if not STAV_PATH.exists():
        raise FileNotFoundError(
            f"Chýba {STAV_PATH.name} — najprv spusti: python stav.py --init")
    return json.loads(STAV_PATH.read_text(encoding="utf-8"))


def save_stav(data: dict):
    """Atomický zápis (tmp súbor + replace), aby pád neponechal poškodený JSON."""
    tmp = STAV_PATH.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=1),
                   encoding="utf-8")
    tmp.replace(STAV_PATH)


# ---------- prvotné naplnenie ----------

def init_z_dat() -> dict:
    """Seed: pre každý servis z unikátov načíta jeho servis_subory/*.xlsx
    a založí hlásenia v stave `odoslane` s úvodnou poznámkou."""
    records = load_unikaty()
    fname_map = build_filename_map(records)
    data = {"verzia": 1, "hlasenia": {}, "servisy": {}}

    for rec in records:
        h = rec["h"]
        xlsx_path = SERVIS_DIR / f"{fname_map[h]}.xlsx"
        if not xlsx_path.exists():
            print(f"  SKIP (chýba xlsx): {xlsx_path.name}")
            continue
        data["servisy"][h] = {
            "nazov": rec["nazov"],
            "email1": rec["email1"],
            "email2": rec["email2"],
            "odpovedal": [],          # [{datum, mail_id}]
            "rucna_kontrola": [],     # [{datum, mail_id, popis}]
        }
        for row in read_xlsx_rows(xlsx_path):
            if row[0] is None:
                continue
            hlasenie = str(row[0]).strip()
            data["hlasenia"][hlasenie] = {
                "servis_h": h,
                "stav": "odoslane",
                "on_turn": ON_TURN["odoslane"],
                "poznamky": [{
                    "datum": DATUM_ZIADOSTI,
                    "zdroj": "ziadost",
                    "text": "Žiadosť o stav reklamácie odoslaná servisu",
                }],
            }
    return data


# ---------- mutácie (append-only, idempotentné podľa mail_id) ----------

def pridaj_odpoved(data: dict, id_hlasenia: str, *, text: str,
                   confidence: str, datum: str, zdroj: str,
                   mail_id: str) -> bool:
    """Pripojí poznámku z odpovede servisu k hláseniu a posunie stavový
    automat (odoslane -> odpovedane). Vráti False, ak už poznámka z tohto
    mailu a zdroja existuje (dedup -> opakované spustenie nič neduplikuje)."""
    rec = data["hlasenia"].get(id_hlasenia)
    if rec is None:
        return False
    for p in rec["poznamky"]:
        if p.get("mail_id") == mail_id and p.get("zdroj") == zdroj:
            return False
    rec["poznamky"].append({
        "datum": datum,
        "zdroj": zdroj,            # "priloha" | "telo"
        "text": text,
        "confidence": confidence,  # "high" | "low"
        "mail_id": mail_id,
    })
    if rec["stav"] == "odoslane":
        rec["stav"] = "odpovedane"
        rec["on_turn"] = ON_TURN["odpovedane"]
    return True


def zaznamenaj_odpoved_servisu(data: dict, servis_h: str, *,
                               datum: str, mail_id: str) -> bool:
    """Eviduje, že od servisu prišiel mail (nezávisle od spárovania)."""
    srv = data["servisy"].get(servis_h)
    if srv is None:
        return False
    if any(z.get("mail_id") == mail_id for z in srv["odpovedal"]):
        return False
    srv["odpovedal"].append({"datum": datum, "mail_id": mail_id})
    return True


def pridaj_rucnu_kontrolu(data: dict, servis_h: str, *, datum: str,
                          mail_id: str, popis: str) -> bool:
    """Mail, z ktorého sa nedalo nič spárovať — celý ide na ručnú kontrolu
    k servisu. Žiadne hádanie."""
    srv = data["servisy"].get(servis_h)
    if srv is None:
        return False
    if any(r.get("mail_id") == mail_id for r in srv["rucna_kontrola"]):
        return False
    srv["rucna_kontrola"].append(
        {"datum": datum, "mail_id": mail_id, "popis": popis})
    return True


def uzavri_hlasenie(data: dict, id_hlasenia: str, datum: str) -> bool:
    rec = data["hlasenia"].get(id_hlasenia)
    if rec is None or rec["stav"] == "uzavrete":
        return False
    rec["stav"] = "uzavrete"
    rec["on_turn"] = ON_TURN["uzavrete"]
    rec["poznamky"].append(
        {"datum": datum, "zdroj": "manual", "text": "Uzavreté ručne"})
    return True


# ---------- prehľad ----------

def vypis_prehlad(data: dict):
    hlasenia = data["hlasenia"]
    servisy = data["servisy"]
    podla_stavu = {s: 0 for s in STAVY}
    for rec in hlasenia.values():
        podla_stavu[rec["stav"]] += 1
    odpovedali = [h for h, s in servisy.items() if s["odpovedal"]]
    mlcia = [h for h, s in servisy.items() if not s["odpovedal"]]
    rucne = [(h, len(s["rucna_kontrola"]))
             for h, s in servisy.items() if s["rucna_kontrola"]]

    print(f"Hlásení: {len(hlasenia)}  "
          + "  ".join(f"{s}={n}" for s, n in podla_stavu.items()))
    print(f"Servisov: {len(servisy)}, odpovedalo {len(odpovedali)}, "
          f"mlčí {len(mlcia)}")
    if rucne:
        print("Ručná kontrola:")
        for h, n in rucne:
            print(f"  - {h} ({n} mail/y)")
    if mlcia:
        print("Zatiaľ neodpovedali:")
        for h in mlcia:
            print(f"  - {h}")


def sync_zelene(data: dict) -> tuple[int, int]:
    """Označí ako uzavreté všetky hlásenia, ktoré má Eva zelené v Sheet1.

    Vracia (označených, preskočených).
    """
    import openpyxl
    from datetime import date
    wb = openpyxl.load_workbook(STAV_XLSX, data_only=True)
    ws = wb["Sheet1"]
    dnes = date.today().isoformat()
    oznacenych = preskoc = 0
    for row in ws.iter_rows(min_row=2):
        fill = row[0].fill
        if not (fill and fill.fgColor and fill.fgColor.type == "rgb"
                and fill.fgColor.rgb == _ZELENA_RGB):
            continue
        hid_val = row[0].value
        if hid_val is None:
            continue
        hid = str(int(hid_val)).strip()
        ok = uzavri_hlasenie(data, hid, dnes)
        if ok:
            oznacenych += 1
        else:
            preskoc += 1
    return oznacenych, preskoc


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[1])
    parser.add_argument("--init", action="store_true",
                        help="Prvotné naplnenie zo servis_subory/")
    parser.add_argument("--force", action="store_true",
                        help="S --init prepíše existujúci stavový súbor")
    parser.add_argument("--prehlad", action="store_true",
                        help="Vypíše súhrn stavu")
    parser.add_argument("--uzavri", type=str, default=None,
                        help="Uzavrie hlásenia (čísla oddelené ;)")
    parser.add_argument("--sync-zelene", action="store_true",
                        help="Označí zelené riadky zo Sheet1 ako uzavreté v JSON")
    args = parser.parse_args()

    if args.init:
        if STAV_PATH.exists() and not args.force:
            print(f"{STAV_PATH.name} už existuje — použi --force na prepis.")
            return
        data = init_z_dat()
        save_stav(data)
        print(f"Založené: {len(data['hlasenia'])} hlásení, "
              f"{len(data['servisy'])} servisov -> {STAV_PATH.name}")
    elif args.uzavri:
        from datetime import date
        data = load_stav()
        dnes = date.today().isoformat()
        for hid in [x.strip() for x in args.uzavri.split(";") if x.strip()]:
            ok = uzavri_hlasenie(data, hid, dnes)
            print(f"  {'OK' if ok else 'SKIP (neznáme/uzavreté)'} {hid}")
        save_stav(data)
    elif args.sync_zelene:
        data = load_stav()
        n, skip = sync_zelene(data)
        save_stav(data)
        print(f"Sync zelených: {n} nových uzavretých, {skip} preskočených "
              f"(už uzavreté alebo neznáme hlásenie)")
    elif args.prehlad:
        vypis_prehlad(load_stav())
    else:
        parser.print_help()


if __name__ == "__main__":
    main()

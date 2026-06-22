"""
Príprava cache odpovede/<mail_id>/ na klasifikáciu Claude agentom
(.claude/agents/klasifikator-odpovedi.md, spúšťa sa cez /klasifikuj-odpovede).

Pre každý stiahnutý mail deterministicky:
  1. rozbalí *.zip prílohy do _rozbalene/ (servisy v zipoch posielajú PDF
     servisné listy aj fotky potvrdeniek),
  2. prepíše .xlsx a .eml prílohy (aj rozbalené) na text do _prepisy/ —
     agent má len Read/Glob/Write a Read binárne xlsx nečíta; PDF
     a obrázky číta natívne, tie sa neprepisujú,
  3. zapíše kontext.json — servis, dátum a hlásenia patriace servisu,
     aby agent nemusel robiť párovanie (reuse najdi_servis cez
     zostav_mapu_mailov zo spracuj_odpovede.py),
  4. NDR (nedoručenky) a nespárované maily označí "preskocit" —
     tie agent neklasifikuje, riešia sa inde.

Na záver vypíše zoznam mailov čakajúcich na klasifikáciu (tie bez
klasifikacia.json). Idempotentné — bezpečné spúšťať opakovane.

Použitie:
    python priprav_klasifikaciu.py            # priprav + zoznam čakajúcich
    python priprav_klasifikaciu.py --zoznam   # len vypíš čakajúce (nič nemení)
"""
import argparse
import json
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from spracuj_odpovede import (  # noqa: E402
    ODPOVEDE_DIR, zostav_mapu_mailov, _bezpecny_nazov, _unikatna,
)
from rozdel_servisy import SRC  # noqa: E402
from parsers import eml_na_text, xlsx_na_text  # noqa: E402
import stav as stavmod  # noqa: E402


# ---------- rozbalenie zip príloh ----------

def _oprav_meno_zo_zipu(zinfo: zipfile.ZipInfo) -> str:
    """Zipy bez UTF-8 flagu majú mená dekódované ako cp437 — slovenská
    diakritika z nich vyjde rozsypaná. Skús ju vrátiť späť."""
    if zinfo.flag_bits & 0x800:
        return zinfo.filename
    try:
        return zinfo.filename.encode("cp437").decode("utf-8")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return zinfo.filename


def rozbal_zipy(mdir: Path) -> int:
    """Rozbalí všetky .zip v priečinku mailu do _rozbalene/ naplocho
    (podpriečinky zo zipu sa zahodia, kolízie rieši _unikatna).
    Idempotentné — ak _rozbalene/ už existuje, nechá ho tak.
    Vráti počet súborov v _rozbalene/ (0 = mail nemá zipy)."""
    zipy = sorted(mdir.glob("*.zip"))
    if not zipy:
        return 0
    ciel = mdir / "_rozbalene"
    if ciel.exists():
        return sum(1 for _ in ciel.iterdir())
    ciel.mkdir()
    n = 0
    for zp in zipy:
        try:
            with zipfile.ZipFile(zp) as zf:
                for zinfo in zf.infolist():
                    if zinfo.is_dir():
                        continue
                    meno = Path(_oprav_meno_zo_zipu(zinfo).replace("\\", "/")).name
                    surovy = Path(meno)
                    von = _unikatna(ciel / (
                        _bezpecny_nazov(surovy.stem, 70) + surovy.suffix.lower()))
                    von.write_bytes(zf.read(zinfo))
                    n += 1
        except (zipfile.BadZipFile, OSError) as e:
            print(f"  !! [{mdir.name}] zip {zp.name}: {type(e).__name__}: {e}")
    return n


# ---------- textové prepisy xlsx/eml príloh ----------

def prepis_prilohy(mdir: Path) -> int:
    """Prepíše .xlsx a .eml prílohy (v koreni mailu aj v _rozbalene/)
    na .txt do _prepisy/, aby ich agent prečítal nástrojom Read.
    Idempotentné — ak _prepisy/ existuje, nechá ho tak."""
    zdroje = sorted(mdir.glob("*.xlsx")) + sorted(mdir.glob("*.eml"))
    if (mdir / "_rozbalene").exists():
        zdroje += sorted((mdir / "_rozbalene").glob("*.xlsx"))
        zdroje += sorted((mdir / "_rozbalene").glob("*.eml"))
    if not zdroje:
        return 0
    ciel = mdir / "_prepisy"
    if ciel.exists():
        return sum(1 for _ in ciel.iterdir())
    ciel.mkdir()
    n = 0
    for subor in zdroje:
        try:
            text = (xlsx_na_text(subor) if subor.suffix.lower() == ".xlsx"
                    else eml_na_text(subor))
        except Exception as e:
            print(f"  !! [{mdir.name}] prepis {subor.name}: "
                  f"{type(e).__name__}: {e}")
            continue
        povod = subor.relative_to(mdir).as_posix()
        von = _unikatna(ciel / (_bezpecny_nazov(subor.stem, 70) + ".txt"))
        von.write_text(f"[textový prepis prílohy: {povod}]\n\n" + text,
                       encoding="utf-8")
        n += 1
    return n


# ---------- kontext pre agenta ----------

def nacitaj_druhy() -> dict[str, str]:
    """Hlásenie -> druh (stĺpec B „Druh hlášení" zdrojového excelu:
    H1 záručná, H2 odstúpenie od KS, H3 predpredajná; HX/H9 neznáme).
    Agent ho dostáva v kontext.json ako pomôcku na kontrolu verdiktu."""
    import openpyxl
    druhy = {}
    wb = openpyxl.load_workbook(SRC, read_only=True)
    for r in wb.active.iter_rows(min_row=2, max_col=2, values_only=True):
        if r[0] is not None and r[1] is not None:
            druhy[str(r[0]).strip()] = str(r[1]).strip()
    wb.close()
    return druhy


def zapis_kontext(z: dict, zname: set[str], servis_nazov: str | None,
                  druhy: dict[str, str]):
    kontext = {
        "mail_id": z["meta"]["mail_id"],
        "servis_h": z["servis_h"],
        "servis_nazov": servis_nazov,
        "datum_prijatia": z["datum"],
        "predmet": z["meta"].get("predmet") or "",
        "zname_hlasenia": sorted(zname),
        "druh_hlasenia": {h: druhy[h] for h in sorted(zname) if h in druhy},
    }
    if z["ndr"]:
        kontext["preskocit"] = "ndr"
    elif z["servis_h"] is None:
        kontext["preskocit"] = "nezaradeny"
    (z["dir"] / "kontext.json").write_text(
        json.dumps(kontext, ensure_ascii=False, indent=1), encoding="utf-8")


# ---------- hlavný beh ----------

def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[1])
    parser.add_argument("--zoznam", action="store_true",
                        help="Len vypíše čakajúce maily, nič nemení")
    args = parser.parse_args()

    data = stavmod.load_stav()
    mapa = zostav_mapu_mailov(data)
    if not mapa:
        print(f"Priečinok {ODPOVEDE_DIR.name}/ je prázdny — najprv spusti "
              "stiahni_odpovede_owa.py")
        return

    hlasenia_servisu: dict[str, set[str]] = {}
    for hid, rec in data["hlasenia"].items():
        hlasenia_servisu.setdefault(rec["servis_h"], set()).add(hid)
    druhy = nacitaj_druhy() if not args.zoznam else {}

    cakajuce, hotove, preskocene = [], [], []
    for mid, z in sorted(mapa.items(), key=lambda kv: (kv[1]["datum"], kv[0])):
        servis_nazov = (data["servisy"][z["servis_h"]]["nazov"]
                        if z["servis_h"] else None)
        zname = hlasenia_servisu.get(z["servis_h"], set())

        if not args.zoznam:
            rozbalenych = rozbal_zipy(z["dir"])
            if rozbalenych:
                print(f"  zip [{mid}] {servis_nazov or '?'}: "
                      f"{rozbalenych} súborov v _rozbalene/")
            prepis_prilohy(z["dir"])
            zapis_kontext(z, zname, servis_nazov, druhy)

        if z["ndr"] or z["servis_h"] is None:
            preskocene.append(mid)
        elif (z["dir"] / "klasifikacia.json").exists():
            hotove.append(mid)
        else:
            cakajuce.append((mid, servis_nazov, z["datum"]))

    print(f"\nMailov: {len(mapa)}  klasifikovaných: {len(hotove)}  "
          f"preskočených (NDR/nezaradené): {len(preskocene)}  "
          f"čaká: {len(cakajuce)}")
    if cakajuce:
        print("\nČakajú na klasifikáciu (odpovede/<mail_id>/):")
        for mid, nazov, datum in cakajuce:
            print(f"  {mid}  {datum}  {nazov}")


if __name__ == "__main__":
    main()

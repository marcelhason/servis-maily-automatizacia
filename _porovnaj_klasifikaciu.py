"""
Diag: porovná pokrytie hlásení medzi vrstvami per mail —
klasifikacia.json (agent) vs ai_vysledok.json (API parser) vs heuristika
(poznámky so zdrojom telo/priloha v stave).

Regres = agent našiel MENEJ hlásení než AI vrstva. Vypíše rozdielové
množiny; zhodné maily len spočíta.

Použitie: python _porovnaj_klasifikaciu.py [--vsetko]  (--vsetko = aj zhody)
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from spracuj_odpovede import zostav_mapu_mailov  # noqa: E402
import stav as stavmod  # noqa: E402


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[1])
    parser.add_argument("--vsetko", action="store_true",
                        help="Vypíše aj maily bez rozdielu")
    args = parser.parse_args()

    data = stavmod.load_stav()
    mapa = zostav_mapu_mailov(data)

    # mail_id -> množina hlásení podľa heuristiky (zo stavu)
    heur_podla_mailu: dict[str, set[str]] = {}
    for hid, rec in data["hlasenia"].items():
        for p in rec["poznamky"]:
            if p["zdroj"] in ("telo", "priloha", "ai+heur") and p.get("mail_id"):
                heur_podla_mailu.setdefault(p["mail_id"], set()).add(hid)

    klasifikovanych, zhody, regresy = 0, 0, 0
    for mid, z in sorted(mapa.items(), key=lambda kv: (kv[1]["datum"], kv[0])):
        kl = z.get("klasifikacia")
        if not kl:
            continue
        klasifikovanych += 1
        nazov = (data["servisy"][z["servis_h"]]["nazov"]
                 if z["servis_h"] else "?")

        agent = {str(h.get("id_hlasenia", "")).strip()
                 for h in kl.get("hlasenia", [])}
        ai = set()
        ai_cache = z["dir"] / "ai_vysledok.json"
        if ai_cache.exists():
            ulozene = json.loads(ai_cache.read_text(encoding="utf-8"))
            ai = {p["id_hlasenia"] for p in ulozene["polozky"]}
        heur = heur_podla_mailu.get(mid, set())

        len_agent = agent - ai - heur
        chyba_agentovi = (ai | heur) - agent
        if not len_agent and not chyba_agentovi:
            zhody += 1
            if args.vsetko:
                print(f"  == [{mid}] {nazov}: zhoda "
                      f"(agent={len(agent)}, ai={len(ai)}, heur={len(heur)})")
            continue

        print(f"  [{mid}] {nazov} ({kl.get('mail_typ')}): "
              f"agent={len(agent)}, ai={len(ai)}, heur={len(heur)}")
        if len_agent:
            print(f"      navyše u agenta (PDF/fotky?): "
                  + ", ".join(sorted(len_agent)))
        if chyba_agentovi:
            regresy += 1
            print(f"      REGRES — agentovi chýba: "
                  + ", ".join(sorted(chyba_agentovi)))

    print(f"\nKlasifikovaných mailov: {klasifikovanych}, "
          f"úplná zhoda: {zhody}, regresov: {regresy}")


if __name__ == "__main__":
    main()

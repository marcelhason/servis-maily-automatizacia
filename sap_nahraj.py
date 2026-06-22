"""
SAP nahrávanie príloh a poznámok k hláseniu.

Číta stav_reklamacii.json a prilohy/<cislo_hlasenia>/ a pre každé vybrané
hlásenie nahrá prílohy + vloží pokec do poznámky v SAP cez SAP GUI Scripting.

Spustenie:
    python sap_nahraj.py          # otvorí GUI okno
    python sap_nahraj.py --help

Build → exe:
    build_sap.bat                 # PyInstaller → dist/sap_nahraj.exe
"""
import json
import re
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import ttk, scrolledtext, messagebox

# ROOT kompatibilný s PyInstaller (exe má sys.frozen=True)
if getattr(sys, "frozen", False):
    ROOT = Path(sys.executable).resolve().parent
else:
    ROOT = Path(__file__).resolve().parent

STAV_PATH = ROOT / "stav_reklamacii.json"
PRILOHY_DIR = ROOT / "prilohy"
LOG_PATH = ROOT / "sap_nahraj_log.txt"   # perzistentný log behov (append)

AKCIE_SAP = ("zapisat-do-sap", "ziadost-o-dobropis-v-sap")

KATEGORIA_POPIS = {
    "uzavrete-dobropisovane": "Dobropisované",
    "mrtve": "Mŕtve",
    "oprava": "Opravené",
    "nerentabilna-oprava": "Nerentabilná oprava",
    "nedostupnost-nd": "Nedost. ND",
    "zamietnutie": "Zamietnuté",
    "ping-pong": "Ping-pong",
}


# ============================================================
# === SAP SEKCIA (SAP GUI Scripting cez Windows COM)      ===
# ============================================================

# ID cesty v QM02 — z nahrávky Script3.vbs
_T09 = r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB09"
_T01 = r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01"
# TAB09 subscreen keď je zoznam príloh prázdny (číslo 7217)
_T09_PRVY = (
    _T09 + "/ssubSUB_GROUP_10:SAPLIQS0:7217"
    "/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:8014"
)
# TAB09 subscreen keď už aspoň jedna príloha existuje (7235/7212)
_T09_DALSI = (
    _T09 + "/ssubSUB_GROUP_10:SAPLIQS0:7235"
    "/subCUSTOM_SCREEN:SAPLIQS0:7212"
    "/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:8014"
)
# Základ pre 7235 bez SUBSCREEN_1 — pre čítanie tabuľky príloh
_T09_BASE = (
    _T09 + "/ssubSUB_GROUP_10:SAPLIQS0:7235"
    "/subCUSTOM_SCREEN:SAPLIQS0:7212"
)
# Pole internej poznámky v TAB01 — DVE varianty podľa typu hlásenia:
#   8002 = typ s TAB09 (Script18), 8302 = typ H2/H3 bez TAB09 (Script22)
# sap_zadaj_poznamku skúša obe.
_T01_POZN = (
    _T01 + "/ssubSUB_GROUP_10:SAPLIQS0:7235"
    "/subCUSTOM_SCREEN:SAPLIQS0:7212"
    "/subSUBSCREEN_3:SAPLIQS0:7900/subUSER0001:SAPLXQQM:8002"
    "/cntlINTERNI_POZNAMKA/shellcont/shell"
)
_T01_POZN_8302 = (
    _T01 + "/ssubSUB_GROUP_10:SAPLIQS0:7235"
    "/subCUSTOM_SCREEN:SAPLIQS0:7212"
    "/subSUBSCREEN_3:SAPLIQS0:7900/subUSER0001:SAPLXQQM:8302"
    "/cntlINTERNI_POZNAMKA/shellcont/shell"
)
# GOS toolbox — generický prístup k prílohám (funguje pre VŠETKY typy hlásení,
# nutný pre H2/H3 ktoré nemajú TAB09). Script22.vbs.
_GOS_TITL = "wnd[0]/titl/shellcont/shell"
# Tabuľka partnerov v TAB01 — treba scrollovať aby sa INTERNI_POZNAMKA objavila v strome
_T01_PARTNER_TBL = (
    _T01 + "/ssubSUB_GROUP_10:SAPLIQS0:7235"
    "/subCUSTOM_SCREEN:SAPLIQS0:7212"
    "/subSUBSCREEN_1:SAPLIQS0:7517"
    "/subPARTNER:SAPLIPAR:0201/tblSAPLIPARTCTRL_0200"
)


def sap_spoj():
    """Pripojí sa na bežiaci SAP GUI cez Windows COM. Vracia session objekt."""
    import win32com.client
    gui = win32com.client.GetObject("SAPGUI")
    app = gui.GetScriptingEngine
    conn = app.Children(0)
    return conn.Children(0)


def _najdi_hodnotu_v_strome(obj, hladaj_upper: str, depth: int = 0):
    """Rekurzívne hľadá v GUI strome objekt, ktorého .Text obsahuje hľadaný
    reťazec (case-insensitive). Vracia Id nájdeného poľa alebo None."""
    if depth > 9:
        return None
    try:
        t = obj.Text
        if t and hladaj_upper in t.upper():
            return obj.Id
    except Exception:
        pass
    try:
        ch = obj.Children
        n = ch.Count
    except Exception:
        return None
    for i in range(n):
        try:
            c = ch(i)
        except Exception:
            continue
        r = _najdi_hodnotu_v_strome(c, hladaj_upper, depth + 1)
        if r:
            return r
    return None


def sap_je_dbpservis(session, log=None) -> bool:
    """True ak hlásenie patrí nám — pole "Sklad" (Data zboží) obsahuje 'DBPSERVIS'.

    Marcelovo kritérium: ak je v Sklade '045 124 DBPSERVIS', hlásenie NIE je
    uzavreté a treba ho spracovať. Hľadáme hodnotu v strome (pole je read-only,
    presné ID sa cez recording nedá získať)."""
    try:
        usr = session.findById("wnd[0]/usr")
    except Exception:
        return False
    nid = _najdi_hodnotu_v_strome(usr, "DBPSERVIS")
    if log:
        if nid:
            log(f"  DBPSERVIS OK (pole ...{nid.split('/')[-1]})")
        else:
            log("  DBPSERVIS nenájdené v aktuálnom pohľade hlásenia")
    return nid is not None


def _diag_obrazovka(session, log):
    """Vypíše taby a textové polia hlásenia — pre ladenie keď DBPSERVIS detekcia
    alebo edit-mode zlyhá (nech vidíme štruktúru bez ďalšieho hádania)."""
    try:
        usr = session.findById("wnd[0]/usr")
    except Exception:
        log("  diag: wnd[0]/usr neprístupné")
        return
    log("  --- DIAGNOSTIKA OBRAZOVKY (taby + polia) ---")

    def _rek(obj, depth):
        if depth > 7:
            return
        try:
            t = obj.Type
            oid = obj.Id.split("/")[-1]
            txt = ""
            try:
                txt = (obj.Text or "")[:45]
            except Exception:
                pass
            if t in ("GuiTextField", "GuiCTextField", "GuiTab",
                     "GuiButton", "GuiComboBox") and (txt or t == "GuiTab"):
                log(f"    [{t}] {oid} = '{txt}'")
        except Exception:
            pass
        try:
            ch = obj.Children
            n = ch.Count
        except Exception:
            return
        for i in range(n):
            try:
                c = ch(i)
            except Exception:
                continue
            _rek(c, depth + 1)

    _rek(usr, 0)
    log("  --- koniec diagnostiky ---")


def sap_otvor_hlasenie(session, cislo: str, log=None) -> str:
    """Otvorí hlásenie v QM02. Vracia stav:
    "tab09" — DBPSERVIS hlásenie so záložkou príloh TAB09 (bežný typ),
    "gos"   — DBPSERVIS hlásenie BEZ TAB09 (typ H2/H3) → prílohy cez GOS,
    "skip"  — nie je DBPSERVIS (cudzie / uzavreté hlásenie, preskočiť).
    """
    import time
    # Zatvor prípadné modálne okná (file dialog, GOS popup) pred navigáciou —
    # inak wnd[0]/tbar[0]/okcd nie je prístupné a dostaneme 619
    for w in ("wnd[3]", "wnd[2]", "wnd[1]"):
        try:
            session.findById(w).sendVKey(12)  # ESC
            time.sleep(0.1)
        except Exception:
            pass
    # /n zaistí čistú navigáciu aj keby predchádzajúci beh zanechal SAP v zlom stave
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nqm02"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").text = cislo
    session.findById("wnd[0]").sendVKey(0)
    # KRITÉRIUM: pole Sklad obsahuje 'DBPSERVIS' → naše hlásenie na spracovanie
    if not sap_je_dbpservis(session, log):
        return "skip"
    # Má hlásenie záložku príloh TAB09? (bežný typ áno, H2/H3 nie)
    try:
        session.findById(_T09).select()
        return "tab09"
    except Exception:
        # Typ H2/H3 bez TAB09 — prílohy pôjdu cez GOS toolbox
        if log:
            log("  (bez TAB09 — typ H2/H3, prílohy cez GOS)")
        return "gos"


def sap_ma_prilohy(session) -> bool:
    """True ak TAB09 je v stave 7235 (má aspoň jednu prílohu)."""
    try:
        session.findById(_T09_DALSI + "/btnPB_CREATE")
        return True
    except Exception:
        return False


def _najdi_grid(obj, log=None, depth=0):
    """Rekurzívne hľadá v GUI strome objekt s RowCount (ALV grid).

    Vracia prvý objekt, na ktorom RowCount funguje. Cestou logguje typy
    shell/grid objektov — aby sme videli reálnu štruktúru zoznamu príloh.
    """
    if depth > 7:
        return None
    # Má tento objekt RowCount? Potom je to hľadaný grid.
    try:
        _ = obj.RowCount
        if log:
            try:
                log(f"    GRID nájdený: {obj.Id} Type={obj.Type}")
            except Exception:
                pass
        return obj
    except Exception:
        pass
    # Rekurzia do detí
    try:
        ch = obj.Children
        n = ch.Count
    except Exception:
        return None
    for i in range(n):
        try:
            c = ch(i)
        except Exception:
            continue
        if log:
            try:
                t = c.Type
                if "Shell" in t or "Grid" in t or "Cont" in t:
                    log(f"    strom[{depth}]: {c.Id.split('/')[-1]} Type={t}")
            except Exception:
                pass
        g = _najdi_grid(c, log, depth + 1)
        if g is not None:
            return g
    return None


def sap_zoznam_tab09_priloh(session, log=None) -> tuple[set, bool]:
    """Vráti (množina názvov TAB09 príloh, ok) cez btnPB_READ.

    DÔLEŽITÉ: TAB09 sa pri OTVORENÍ hlásenia VŽDY ukáže v stave 7217 (prázdny
    formulár na pridanie prílohy) — aj keď prílohy existujú (Script16.vbs to
    potvrdzuje). Preto NEPODMIEŇUJEME čítanie cez sap_ma_prilohy() — btnPB_READ
    existuje v oboch stavoch a otvorí zoznam existujúcich príloh.
    Grid hľadáme rekurzívne (jeho presná cesta nie je istá).
    ok=False ak sa grid nenájde — radšej preskočíme než duplikovať.
    """
    import time
    # btnPB_READ skús v oboch stavoch — pri otvorení sme v 7217
    read_btn = None
    for path in (_T09_PRVY + "/btnPB_READ", _T09_DALSI + "/btnPB_READ"):
        try:
            session.findById(path)
            read_btn = path
            break
        except Exception:
            pass
    if read_btn is None:
        if log:
            log("  btnPB_READ nenájdený — preskakujem dedup (prázdne)")
        return set(), True
    try:
        session.findById(read_btn).press()
        time.sleep(0.5)
        # Zoznam príloh sa mohol otvoriť na wnd[0] (shellcont) alebo wnd[1]
        grid = None
        for wid in ("wnd[0]/shellcont[1]", "wnd[1]", "wnd[0]"):
            try:
                kontajner = session.findById(wid)
            except Exception:
                continue
            if log:
                log(f"  hľadám grid v {wid}...")
            grid = _najdi_grid(kontajner, log)
            if grid is not None:
                break
        if grid is None:
            if log:
                log("  grid s RowCount nenájdený — preskakujem hlásenie")
            for c in ("wnd[0]/shellcont[1]", "wnd[1]"):
                try:
                    session.findById(c).close()
                except Exception:
                    pass
            return set(), False
        # Stĺpce gridu
        try:
            cols = list(grid.ColumnOrder)
        except Exception:
            cols = []
        if log:
            log(f"  grid riadkov={grid.RowCount}, stĺpce={cols}")
        hodnoty = set()
        zoznam_cols = cols + [c for c in ("BITM_DESCR", "FILE_NAME", "FILENAME", "DESCRIPTION")
                              if c not in cols]
        for i in range(grid.RowCount):
            for col in zoznam_cols:
                try:
                    val = grid.GetCellValue(i, col)
                    if val and val.strip():
                        hodnoty.add(val.strip())
                except Exception:
                    pass
        if log:
            log(f"  prečítané hodnoty: {sorted(hodnoty)}")
        for c in ("wnd[0]/shellcont[1]", "wnd[1]"):
            try:
                session.findById(c).close()
            except Exception:
                pass
        return hodnoty, True
    except Exception as e:
        if log:
            log(f"  čítanie príloh chyba: {e}")
        for c in ("wnd[0]/shellcont[1]", "wnd[1]"):
            try:
                session.findById(c).close()
            except Exception:
                pass
        return set(), False


def _je_duplicita(p: "Path", sap_set: set) -> bool:
    """True ak súbor p je už v SAP (podľa GOS titulek).

    GOS obsahuje aj systémové záznamy (Hlášení jakosti, Opatření…) —
    filtrujeme len tie čo vyzerajú ako naše prílohy:
    titulek obsahuje 'SERVIS_' alebo začína číslicou (dátum/číslo hlásenia).
    """
    nase = {v for v in sap_set
            if v and ("SERVIS_" in v or v[0].isdigit())}
    if not nase:
        return False
    if p.name in nase or p.stem in nase:
        return True
    stem_prefix = p.stem[:30].lower()
    return any(stem_prefix in val.lower() for val in nase)


def _gos_zavri(session):
    """Zavrie GOS zoznam príloh (ESC na modálnych oknách wnd[2]/wnd[1])."""
    import time
    for _ in range(4):
        zatvorene = False
        for w in ("wnd[2]", "wnd[1]"):
            try:
                session.findById(w).sendVKey(12)  # ESC
                zatvorene = True
                time.sleep(0.2)
            except Exception:
                pass
        if not zatvorene:
            break


def sap_gos_spracuj_prilohy(session, prilohy, log=None):
    """GOS vetva (typ H2/H3 bez TAB09) — Script22.vbs.

    Otvorí GOS zoznam príloh (toolbox → F2 → VIEW_ATTA), prečíta existujúce
    (dedup), pridá nové cez %ATTA_CREATE → %GOS_PCATTA_CREA → file dialóg
    (ctxtDY_PATH/ctxtDY_FILENAME). Vracia počet nahraných alebo None pri chybe.
    """
    import time
    GRID = "wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell"
    # 1. Otvor GOS toolbox
    try:
        session.findById(_GOS_TITL).pressButton("%GOS_TOOLBOX")
    except Exception as e:
        if log:
            log(f"  GOS toolbox sa nepodarilo otvoriť: {e}")
        return None
    # GOS toolbox popup (wnd[1]) sa otvára s oneskorením — počkaj naň pred F2,
    # inak sendVKey padne na "control could not be found by id".
    if not _wait_for_wnd(session, "wnd[1]", 3.0):
        if log:
            log("  GOS toolbox popup (wnd[1]) sa neotvoril")
        _gos_zavri(session)
        return None
    # 2. F2 na prvej položke (Seznam příloh) — otvorí zoznam
    try:
        session.findById(
            "wnd[1]/usr/tblSAPLSWUGOBJECT_CONTROL/txtSWLOBJTDYN-DESCRIPT[0,0]"
        ).caretPosition = 0
    except Exception:
        pass
    try:
        session.findById("wnd[1]").sendVKey(2)  # F2 = zobraziť zoznam
        time.sleep(0.5)
    except Exception as e:
        if log:
            log(f"  GOS F2 (zobraziť zoznam) zlyhalo: {e}")
        _gos_zavri(session)
        return None
    # 3. VIEW_ATTA — rozbalí plný zoznam príloh (cntlCONTAINER_0100)
    try:
        session.findById("wnd[0]/shellcont[1]/shell").pressButton("VIEW_ATTA")
        time.sleep(0.5)
    except Exception:
        pass   # nemusí byť vždy potrebné
    # 4. Nájdi grid a prečítaj existujúce prílohy (dedup)
    grid = None
    try:
        grid = session.findById(GRID)
        _ = grid.RowCount
    except Exception:
        for wid in ("wnd[1]", "wnd[0]/shellcont[1]", "wnd[2]"):
            try:
                grid = _najdi_grid(session.findById(wid), log)
            except Exception:
                grid = None
            if grid is not None:
                break
    # Grid nenájdený NIE JE chyba — znamená, že hlásenie nemá žiadnu prílohu
    # (niet čo zobraziť). Vtedy sap_prilohy = prázdne a ideme rovno vytvárať.
    sap_prilohy = set()
    if grid is None:
        if log:
            log("  GOS grid nenájdený — hlásenie nemá prílohy, idem vytvárať")
    else:
        try:
            cols = list(grid.ColumnOrder)
        except Exception:
            cols = []
        if log:
            log(f"  GOS grid riadkov={grid.RowCount}, stĺpce={cols}")
        # Čítaj len stĺpce s názvom/popisom prílohy (nie CREATOR, CREADATE, ICON)
        zoznam_cols = [c for c in cols
                       if any(k in c.upper() for k in ("DESCR", "NAME", "FILE"))]
        if not zoznam_cols:
            zoznam_cols = cols
        for i in range(grid.RowCount):
            for col in zoznam_cols:
                try:
                    val = grid.GetCellValue(i, col)
                    if val and val.strip():
                        sap_prilohy.add(val.strip())
                except Exception:
                    pass
        if log:
            log(f"  existujúce prílohy (GOS, {len(sap_prilohy)}): {sorted(sap_prilohy)}")
    # 5. Pridaj nové prílohy (dedup) — cesta zo Script27.vbs
    nahratych = 0
    for p in prilohy:
        if _je_duplicita(p, sap_prilohy):
            if log:
                log(f"  preskočená (duplicita v SAP): {p.name}")
            continue
        if log:
            log(f"  nahrávam (GOS) {p.name}...")
        try:
            # shellcont[1]/shell sa po F2 môže objaviť až po krátkom čakaní
            shell = None
            for _ in range(6):
                try:
                    shell = session.findById("wnd[0]/shellcont[1]/shell")
                    break
                except Exception:
                    time.sleep(0.3)
            if shell is None:
                shell = session.findById("wnd[0]/shellcont[1]/shell")  # vyhodí chybu
            shell.pressContextButton("CREATE_ATTA")
            time.sleep(0.3)
            shell.selectContextMenuItem("PCATTA_CREA")
            _wait_for_wnd(session, "wnd[1]", 3.0)
            # Bezpečnostný dialóg pred file dialógom?
            try:
                session.findById("wnd[1]/usr/ctxtDY_PATH")
            except Exception:
                try:
                    session.findById("wnd[1]/tbar[0]/btn[0]").press()
                    _wait_for_wnd(session, "wnd[1]", 3.0)
                except Exception:
                    pass
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = str(p.parent) + "\\"
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = p.name
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            time.sleep(0.3)
            _potvrd_security_ak_je(session)
            sap_prilohy.add(p.name)
            nahratych += 1
        except Exception as e:
            if log:
                log(f"  GOS pridanie {p.name} zlyhalo: {e}")
            _gos_zavri(session)
            return None
    # 6. Zavri GOS zoznam
    _gos_zavri(session)
    return nahratych


def _dismiss_security_dialog(session) -> bool:
    """Klikne 'Povolenie' v SAP GUI bezpečnostnom dialógu (ak sa objaví po upload).

    Dialóg 'Zabezpečenie SAP GUI' má fokus na 'Zamietnutí' → sendVKey(0)/Enter
    by ho zamietol. Preto tlačidlo hľadáme priamo podľa ID alebo textu.
    3 pokusy × 0.5s na prípad oneskorenia dialógu.
    """
    import time
    time.sleep(0.5)
    for attempt in range(3):
        for wid in ("wnd[1]", "wnd[2]"):
            try:
                wnd = session.findById(wid)
                # Zaškrtni "Uchování mého rozhodnutí" — SAP si zapamätá
                for chk_id in (wid + "/usr/chkSCRIPTING_SAVE_FLAG",
                                wid + "/usr/chkSAVE_FLAG"):
                    try:
                        session.findById(chk_id).selected = True
                        break
                    except Exception:
                        pass
                # Klikni "Povolenie" (Allow) priamo — Enter by klikol Zamietnutí
                clicked = False
                for btn_id in (wid + "/usr/btnALLOW",
                                wid + "/usr/btnPOVOLENIE",
                                wid + "/tbar[0]/btn[0]"):
                    try:
                        session.findById(btn_id).press()
                        clicked = True
                        break
                    except Exception:
                        pass
                if not clicked:
                    # Fallback: prehľadaj wnd[1]/usr aj wnd[1] priamo
                    for cont_path in (wid + "/usr", wid):
                        try:
                            cont = session.findById(cont_path)
                            for i in range(cont.Children.Count):
                                ch = cont.Children(i)
                                txt = getattr(ch, "Text", "").lower()
                                if getattr(ch, "Type", "") == "GuiButton" and \
                                   ("ovolen" in txt or "allow" in txt):
                                    ch.press()
                                    clicked = True
                                    break
                        except Exception:
                            pass
                        if clicked:
                            break
                if not clicked:
                    wnd.sendVKey(0)
                time.sleep(0.3)
                return True
            except Exception:
                pass
        if attempt < 2:
            time.sleep(0.5)
    return False


def _wait_for_wnd(session, wnd_id: str, timeout: float = 3.0) -> bool:
    """Čaká kým SAP otvorí okno wnd_id (max timeout sekúnd). Vracia True ak úspech."""
    import time
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            session.findById(wnd_id)
            return True
        except Exception:
            time.sleep(0.15)
    return False


def sap_nahraj_prilohu(session, cesta: Path, prvy: bool = False):
    """Nahrá jeden súbor ako prílohu otvoreného hlásenia (TAB09).

    Parameter `prvy` sa ignoruje — stav TAB09 (7217 vs 7235) zisťujeme priamo
    z existencie btnPB_CREATE. POZOR: po btnPB_READ + close je TAB09 VŽDY v
    stave 7235 (Script17.vbs riadky 25-27), preto skúšame 7235 ako prvé.
    """
    import time
    # Auto-detekcia stavu: nájdi existujúci btnPB_CREATE (7235 alebo 7217)
    create_base = None
    for base in (_T09_DALSI, _T09_PRVY):
        try:
            session.findById(base + "/btnPB_CREATE")
            create_base = base
            break
        except Exception:
            pass
    if create_base is None:
        raise RuntimeError("btnPB_CREATE v TAB09 nenájdený")
    # Nastav typ prílohy SERVIS, ak combobox v danom stave existuje
    try:
        session.findById(create_base + "/cmbLS_ATTACHTYPE-ATTACH_TYPE").key = "SERVIS"
    except Exception:
        pass
    btn_id = create_base + "/btnPB_CREATE"
    session.findById(btn_id).press()
    # Čakaj kým SAP otvorí wnd[1] (security dialóg ALEBO file dialóg)
    _wait_for_wnd(session, "wnd[1]", timeout=3.0)
    # Script15.vbs: po PB_CREATE prichádza security dialóg ("Zabezpečenie SAP GUI")
    # ako wnd[1]. "Povolenie" je wnd[1]/tbar[0]/btn[0]. File dialóg príde až po ňom.
    # Ak txtDY_PATH neexistuje → sme na security dialógu → klikneme Povolenie.
    try:
        session.findById("wnd[1]/usr/txtDY_PATH")
    except Exception:
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        # Čakaj kým security dialóg zatvori a objaví sa file dialóg
        _wait_for_wnd(session, "wnd[1]", timeout=3.0)
    session.findById("wnd[1]/usr/txtDY_PATH").text = str(cesta.parent) + "\\"
    session.findById("wnd[1]/usr/txtDY_FILENAME").text = cesta.name
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    # Po potvrdení súboru SAP reálne číta súbor → môže prísť ďalší security
    # dialóg ("Systém sa pokúša o prístup k súboru"). Odklikni Povolenie.
    _potvrd_security_ak_je(session)


def _potvrd_security_ak_je(session, timeout: float = 1.5):
    """Ak sa po prístupe k súboru objaví security dialóg (wnd[1] bez file polí),
    klikne Povolenie (btn[0]). File dialóg (má txtDY_PATH) nechá tak."""
    import time
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            session.findById("wnd[1]")
        except Exception:
            time.sleep(0.1)
            continue
        # wnd[1] existuje — je to file dialóg alebo security?
        try:
            session.findById("wnd[1]/usr/txtDY_PATH")
            return   # file dialóg, nie security
        except Exception:
            pass
        try:
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            time.sleep(0.2)
        except Exception:
            pass
        return
    return


def _najdi_pole_v_strome(obj, substr_id: str, depth: int = 0):
    """Rekurzívne hľadá objekt, ktorého Id obsahuje substr_id. Vracia objekt/None."""
    if depth > 11:
        return None
    try:
        if substr_id in obj.Id:
            return obj
    except Exception:
        pass
    try:
        ch = obj.Children
        n = ch.Count
    except Exception:
        return None
    for i in range(n):
        try:
            c = ch(i)
        except Exception:
            continue
        r = _najdi_pole_v_strome(c, substr_id, depth + 1)
        if r:
            return r
    return None


def sap_zadaj_poznamku(session, text: str, log=None) -> bool:
    """Prepend-uje poznámku (s prefixom LH:) pred existujúci obsah INTERNI_POZNAMKA.

    Pole poznámky má rôzne čísla podobrazovky podľa typu hlásenia (8002, 8302, …).
    Skúšame známe fixné cesty, inak hľadáme cntlINTERNI_POZNAMKA rekurzívne.
    Vracia False ak rovnaký text tam už je (duplicita — nezapíše).
    """
    try:
        session.findById(_T01).select()
    except Exception:
        pass
    # Scroll partner tabuľky — aby sa INTERNI_POZNAMKA objavila v DOM
    # (Script3/23: scroll na 12). Pri niektorých typoch netreba (try/except).
    try:
        tbl = session.findById(_T01_PARTNER_TBL)
        tbl.verticalScrollbar.position = 12
    except Exception:
        pass
    # Nájdi pole poznámky — najprv známe fixné cesty (8002 TAB09 typ, 8302 H3 typ)
    pole = None
    for pid in (_T01_POZN, _T01_POZN_8302):
        try:
            pole = session.findById(pid)
            break
        except Exception:
            pass
    # Fallback: nájdi cntlINTERNI_POZNAMKA kdekoľvek v TAB01 (iné typy, napr. H2)
    if pole is None:
        try:
            t01 = session.findById(_T01)
            cnt = _najdi_pole_v_strome(t01, "cntlINTERNI_POZNAMKA")
            if cnt is not None:
                pole = session.findById(cnt.Id + "/shellcont/shell")
                if log:
                    log(f"  poznámka: pole nájdené rekurzívne (...{cnt.Id.split('/')[-1]})")
        except Exception:
            pole = None
    if pole is None:
        raise RuntimeError("INTERNI_POZNAMKA pole sa nenašlo (8002/8302 ani rekurzívne)")
    # SAP vracia \r; stačí skontrolovať či text začína "LH: " — vždy prepend-ujeme
    stary = (pole.text or "").replace("\r\n", "\n").replace("\r", "\n")
    # Kontrola: ak začiatok poznámky zodpovedá nášmu textu (prvých 20 znakov),
    # je to duplicita — ak Eva napíše inú správu, prefix sa líši → zapíše sa
    if stary.lstrip().startswith("LH: " + text[:20]):
        return False
    novy = "LH: " + text + ("\n\n" + stary if stary.strip() else "")
    pole.text = novy
    return True


def sap_uloz(session):
    """Uloží hlásenie (F11 / disketa)."""
    session.findById("wnd[0]/tbar[0]/btn[11]").press()


# ============================================================
# === DÁTOVÁ VRSTVA                                       ===
# ============================================================

def nacitaj_stav() -> dict:
    if not STAV_PATH.exists():
        raise FileNotFoundError(f"Chýba {STAV_PATH.name}")
    return json.loads(STAV_PATH.read_text(encoding="utf-8"))


def uloz_stav(data: dict):
    STAV_PATH.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def zisti_hlasenia(data: dict, len_sap: bool = True) -> list[dict]:
    """Vráti zoznam hlásení vhodných na spracovanie.

    Každý záznam: {hid, servis, kategoria, akcia, pokec, cislo_dobropisu,
                   prilohy (list Path), prilohy_count}
    """
    servisy = data.get("servisy", {})
    vysledok = []
    for hid, rec in data.get("hlasenia", {}).items():
        if rec.get("stav") == "uzavrete":
            continue
        kl = rec.get("klasifikacia")
        if not kl:
            continue
        akcia = kl.get("akcia", "")
        if len_sap and akcia not in AKCIE_SAP:
            continue
        adresar = PRILOHY_DIR / hid
        subory = sorted(adresar.iterdir()) if adresar.exists() else []

        servis_h = rec.get("servis_h", "")
        servis_nazov = servisy.get(servis_h, {}).get("nazov", servis_h) if servis_h else ""

        vysledok.append({
            "hid": hid,
            "servis": servis_nazov,
            "kategoria": kl.get("kategoria", ""),
            "akcia": akcia,
            "pokec": kl.get("pokec", ""),
            "cislo_dobropisu": kl.get("cislo_dobropisu"),
            "prilohy": subory,
            "prilohy_count": len(subory),
        })
    vysledok.sort(key=lambda r: (r["akcia"], r["hid"]))
    return vysledok


# marker výsledku v logu: "  [8803000005] OK ✓ ..." / "CHYBA" / "STOP" / "SKIP"
_LOG_VYSLEDOK = re.compile(r"\[(\d+)\]\s+(OK|CHYBA|STOP|SKIP)\b")


def nacitaj_stav_z_logu() -> dict:
    """Prečíta sap_nahraj_log.txt → {hid: 'ok'|'chyba'|'skip'} podľa
    POSLEDNÉHO výsledku každého hlásenia (log je append, hlásenie mohlo
    byť skúšané viackrát — platí posledný zápis)."""
    if not LOG_PATH.exists():
        return {}
    mapovanie = {"OK": "ok", "CHYBA": "chyba", "STOP": "chyba", "SKIP": "skip"}
    vysledok: dict[str, str] = {}
    try:
        for riadok in LOG_PATH.read_text(encoding="utf-8").splitlines():
            m = _LOG_VYSLEDOK.search(riadok)
            if m:
                vysledok[m.group(1)] = mapovanie[m.group(2)]
    except Exception:
        pass
    return vysledok


# ============================================================
# === SPRACOVANIE JEDNÉHO HLÁSENIA                        ===
# ============================================================

def spracuj_hlasenie(h: dict, session, dry_run: bool, log,
                     data: dict | None = None) -> str:
    """Spracuje jedno hlásenie — nahrá prílohy + vloží poznámku.

    Vracia stav: "ok" (spracované), "skip" (legitímne preskočené — uzavreté
    hlásenie / chýba pokec) alebo "chyba" (reálne zlyhanie).
    dry_run=True: len logguje, nič nespúšťa v SAP.
    """
    hid = h["hid"]
    log(f"  [{hid}] {h['servis'][:40]} | {KATEGORIA_POPIS.get(h['kategoria'], h['kategoria'])}")

    if not h["pokec"]:
        log(f"  [{hid}] SKIP: chýba pokec (poznámka)")
        return "skip"

    if dry_run:
        log(f"  [{hid}] DRY: otvorím QM02 {hid}")
        for p in h["prilohy"]:
            log(f"  [{hid}] DRY:   {p.name} → nahrám (GOS check v reálnom behu)")
        log(f"  [{hid}] DRY:   poznámka → zapíšem ({len(h['pokec'])} znakov)")
        log(f"  [{hid}] DRY: uložím — OK ✓")
        return "ok"

    try:
        stav_otvor = sap_otvor_hlasenie(session, hid, log)
        if stav_otvor == "skip":
            log(f"  [{hid}] SKIP: nie je DBPSERVIS (cudzie/uzavreté hlásenie)")
            return "skip"

        prilohy_zlyhali = False
        if stav_otvor == "gos":
            # Typ H2/H3 bez TAB09 — prílohy cez GOS toolbox (Script27)
            nahratych = sap_gos_spracuj_prilohy(session, h["prilohy"], log)
            if nahratych is None:
                # Prílohy sa nepodarili, ALE poznámku skúsime zapísať aj tak
                # (má vlastný "LH:" dedup → opakovanie nič nezduplikuje).
                log(f"  [{hid}] GOS prílohy zlyhali — skúsim aspoň poznámku")
                prilohy_zlyhali = True
                nahratych = 0
        else:
            # Bežný typ s TAB09 — čítanie cez btnPB_READ, upload cez btnPB_CREATE
            sap_prilohy, sap_ok = sap_zoznam_tab09_priloh(session, log)
            if not sap_ok:
                log(f"  [{hid}] STOP: nepodarilo sa prečítať zoznam príloh — riziko duplicít, preskačujem")
                return "chyba"
            log(f"  [{hid}] existujúce prílohy ({len(sap_prilohy)}): "
                + (", ".join(sorted(sap_prilohy)) or "prázdne"))
            # Po btnPB_READ + close je TAB09 v stave 7235 — sap_nahraj_prilohu si
            # stav detekuje sám (žiadny re-select ani `prvy` — Script17.vbs).
            nahratych = 0
            for p in h["prilohy"]:
                if _je_duplicita(p, sap_prilohy):
                    log(f"  [{hid}] preskočená (duplicita v SAP): {p.name}")
                    continue
                log(f"  [{hid}] nahrávam {p.name}...")
                sap_nahraj_prilohu(session, p)
                sap_prilohy.add(p.name)
                nahratych += 1

        log(f"  [{hid}] vkladám poznámku ({len(h['pokec'])} znakov)...")
        zapisana = sap_zadaj_poznamku(session, h["pokec"], log)
        sap_uloz(session)
        if prilohy_zlyhali:
            log(f"  [{hid}] CHYBA: GOS prílohy zlyhali (poznámka "
                f"{'zapísaná' if zapisana else 'už tam je'}) — skús znova")
            return "chyba"
        log(f"  [{hid}] OK ✓  ({nahratych} nových príloh, "
            f"poznámka {'zapísaná' if zapisana else 'preskočená — už tam je'})")
        return "ok"
    except Exception as e:
        log(f"  [{hid}] CHYBA: {type(e).__name__}: {e}")
        return "chyba"


# ============================================================
# === GUI                                                 ===
# ============================================================

class SapNahrajApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SAP nahrávanie príloh a poznámok")
        self.geometry("760x560")
        self.resizable(True, True)
        self._hlasenia: list[dict] = []
        self._checkboxy: dict[str, tk.BooleanVar] = {}
        self._stav_log: dict[str, str] = {}
        self._buduj_gui()
        self.after(100, self._nacitaj)

    # ---- stavba okna ----

    # poradie a popisy stavového filtra (kľúč → text v selekte)
    _FILTER_KLUCE = ("vsetky", "nespracovane", "hotove", "preskocene", "chybne")
    _FILTER_POPIS = {
        "vsetky": "Všetky", "nespracovane": "Ešte nespracované",
        "hotove": "Hotové", "preskocene": "Preskočené", "chybne": "Chybné",
    }
    # stav z logu → kľúč filtra
    _STAV_KLUC = {"ok": "hotove", "skip": "preskocene", "chyba": "chybne"}

    def _buduj_gui(self):
        top = tk.Frame(self, padx=8, pady=6)
        top.pack(fill=tk.X)

        tk.Label(top, text="Zobraz:").pack(side=tk.LEFT)
        self._filter_combo = ttk.Combobox(top, state="readonly", width=26)
        self._filter_combo.pack(side=tk.LEFT, padx=4)
        self._filter_combo.bind("<<ComboboxSelected>>",
                                lambda e: self._napln_tabulku())

        # tabuľka hlásení
        stlpce = ("hid", "servis", "kategoria", "prilohy", "akcia")
        nadpisy = ("Hlásenie", "Servis", "Kategória", "Príloh", "Akcia")
        sirky = (90, 260, 120, 55, 80)

        frame_tree = tk.Frame(self)
        frame_tree.pack(fill=tk.BOTH, expand=True, padx=8)

        vsb = ttk.Scrollbar(frame_tree, orient=tk.VERTICAL)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self._tree = ttk.Treeview(
            frame_tree, columns=stlpce, show="headings",
            selectmode="none", yscrollcommand=vsb.set)
        vsb.config(command=self._tree.yview)

        for col, nadpis, sirka in zip(stlpce, nadpisy, sirky):
            self._tree.heading(col, text=nadpis)
            self._tree.column(col, width=sirka, minwidth=40)

        # vlastný stĺpec pre checkbox (obrázok simulujeme textom)
        self._tree.column("hid", width=90)
        # farby podľa stavu z logu (sap_nahraj_log.txt)
        self._tree.tag_configure("ok", background="#c6efce")     # zelená — prešlo
        self._tree.tag_configure("chyba", background="#ffc7ce")  # červená — spadlo
        self._tree.tag_configure("skip", background="#e8e8e8")   # sivá — preskočené
        self._tree.pack(fill=tk.BOTH, expand=True)
        self._tree.bind("<Button-1>", self._preklik_riadok)

        # --- tlačidlá ---
        btn_frame = tk.Frame(self, padx=8, pady=4)
        btn_frame.pack(fill=tk.X)
        tk.Button(btn_frame, text="Vybrať všetky",
                  command=self._vybrat_vsetky).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_frame, text="Odznačiť",
                  command=self._odznacit).pack(side=tk.LEFT, padx=2)
        tk.Label(btn_frame, text="").pack(side=tk.LEFT, expand=True)
        tk.Label(btn_frame, text="dávka:").pack(side=tk.LEFT)
        self._davka_var = tk.StringVar(value="50")
        tk.Spinbox(btn_frame, from_=1, to=999, width=5,
                   textvariable=self._davka_var).pack(side=tk.LEFT, padx=(2, 6))
        tk.Button(btn_frame, text="▶  Spustiť ďalšiu dávku", bg="#e6f4ea",
                  command=self._nova_davka).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_frame, text="▶  Spustiť vybrané", bg="#dcedf7",
                  command=self._spusti_sap).pack(side=tk.LEFT, padx=2)

        # --- progres ---
        prog_frame = tk.Frame(self, padx=8, pady=2)
        prog_frame.pack(fill=tk.X)
        self._progress = ttk.Progressbar(prog_frame, mode="determinate")
        self._progress.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._progress_var = tk.StringVar(value="pripravené")
        tk.Label(prog_frame, textvariable=self._progress_var, width=22,
                 anchor="w").pack(side=tk.LEFT, padx=8)

        # log
        tk.Label(self, text="Priebeh:", anchor="w",
                 padx=8).pack(fill=tk.X)
        self._log = scrolledtext.ScrolledText(
            self, height=8, state=tk.DISABLED, font=("Consolas", 9))
        self._log.pack(fill=tk.BOTH, padx=8, pady=(0, 8))

    def _progress_set(self, hotovo: int, total: int, popis: str = ""):
        """Aktualizuje progress-bar a textový stav (volať cez self.after)."""
        self._progress["maximum"] = max(1, total)
        self._progress["value"] = hotovo
        if popis:
            self._progress_var.set(popis)
        else:
            pct = int(hotovo / total * 100) if total else 0
            self._progress_var.set(f"{hotovo}/{total}  ({pct} %)")

    # ---- načítanie dát ----

    def _riadok_values(self, h: dict) -> tuple:
        """Hodnoty riadku tabuľky pre hlásenie h (checkbox podľa _checkboxy)."""
        var = self._checkboxy.get(h["hid"])
        znak = "☑ " if (var and var.get()) else "☐ "
        return (
            znak + h["hid"],
            h["servis"][:40],
            KATEGORIA_POPIS.get(h["kategoria"], h["kategoria"]),
            h["prilohy_count"],
            h["akcia"].replace("-", " "),
        )

    def _kluc_pre_hid(self, hid: str) -> str:
        """Kľúč filtra pre hlásenie podľa stavu z logu."""
        return self._STAV_KLUC.get(self._stav_log.get(hid), "nespracovane")

    def _aktualny_filter(self) -> str:
        idx = self._filter_combo.current()
        return self._FILTER_KLUCE[idx] if idx >= 0 else "vsetky"

    def _nacitaj(self):
        try:
            data = nacitaj_stav()
        except FileNotFoundError as e:
            messagebox.showerror("Chyba", str(e))
            return

        # vždy všetky hlásenia s klasifikáciou; selekt nižšie filtruje podľa stavu
        self._hlasenia = zisti_hlasenia(data, len_sap=False)
        self._stav_log = nacitaj_stav_z_logu()
        self._checkboxy.clear()
        for h in self._hlasenia:
            # nespracované default zaškrtnuté, spracované (ok/skip/chyba) nie
            self._checkboxy[h["hid"]] = tk.BooleanVar(
                value=self._stav_log.get(h["hid"]) is None)

        self._aktualizuj_filter_combo()
        self._napln_tabulku()
        self._log_pis(f"Načítané: {len(self._hlasenia)} hlásení s klasifikáciou\n")

    def _aktualizuj_filter_combo(self):
        """Prebuduje hodnoty selektu s počtami; zachová aktuálny výber."""
        pocty = {k: 0 for k in self._FILTER_KLUCE}
        pocty["vsetky"] = len(self._hlasenia)
        for h in self._hlasenia:
            pocty[self._kluc_pre_hid(h["hid"])] += 1
        values = [f"{self._FILTER_POPIS[k]} ({pocty[k]})"
                  for k in self._FILTER_KLUCE]
        idx = self._filter_combo.current()
        if idx < 0:
            idx = 0
        self._filter_combo["values"] = values
        self._filter_combo.current(idx)

    def _napln_tabulku(self):
        """Naplní tabuľku hláseniami vyhovujúcimi aktuálnemu stavovému filtru."""
        for item in self._tree.get_children():
            self._tree.delete(item)
        f = self._aktualny_filter()
        for h in self._hlasenia:
            stav = self._stav_log.get(h["hid"])
            if f != "vsetky" and f != self._kluc_pre_hid(h["hid"]):
                continue
            tag = (stav,) if stav in ("ok", "chyba", "skip") else ()
            self._tree.insert("", tk.END, iid=h["hid"],
                              values=self._riadok_values(h), tags=tag)

    def _obnov_farby(self):
        """Po behu znovu načíta log, odznačí spracované a prekreslí tabuľku."""
        self._stav_log = nacitaj_stav_z_logu()
        for hid, var in self._checkboxy.items():
            var.set(self._stav_log.get(hid) is None)
        self._aktualizuj_filter_combo()
        self._napln_tabulku()

    # ---- interakcia s tabuľkou ----

    def _preklik_riadok(self, event):
        item = self._tree.identify_row(event.y)
        if not item or item not in self._checkboxy:
            return
        var = self._checkboxy[item]
        var.set(not var.get())
        h = next((x for x in self._hlasenia if x["hid"] == item), None)
        if h:
            stav = self._stav_log.get(item)
            tag = (stav,) if stav in ("ok", "chyba", "skip") else ()
            self._tree.item(item, values=self._riadok_values(h), tags=tag)

    def _nastav_vsetky_zobrazene(self, hodnota: bool):
        """Zaškrtne/odškrtne len práve zobrazené riadky (podľa filtra)."""
        for hid in self._tree.get_children():
            if hid in self._checkboxy:
                self._checkboxy[hid].set(hodnota)
                h = next((x for x in self._hlasenia if x["hid"] == hid), None)
                if h:
                    stav = self._stav_log.get(hid)
                    tag = (stav,) if stav in ("ok", "chyba", "skip") else ()
                    self._tree.item(hid, values=self._riadok_values(h), tags=tag)

    def _vybrat_vsetky(self):
        self._nastav_vsetky_zobrazene(True)

    def _odznacit(self):
        self._nastav_vsetky_zobrazene(False)

    # ---- log ----

    def _log_pis(self, text: str):
        self._log.config(state=tk.NORMAL)
        self._log.insert(tk.END, text)
        self._log.see(tk.END)
        self._log.config(state=tk.DISABLED)
        # Zapíš aj do perzistentného log súboru (pre náhodnú kontrolu)
        try:
            with open(LOG_PATH, "a", encoding="utf-8") as f:
                f.write(text)
        except Exception:
            pass   # log nesmie zhodiť beh

    def _log_clear(self):
        self._log.config(state=tk.NORMAL)
        self._log.delete("1.0", tk.END)
        self._log.config(state=tk.DISABLED)

    # ---- spustenie ----

    def _vybrane(self) -> list[dict]:
        """Zaškrtnuté hlásenia, ktoré sú zároveň zobrazené v aktuálnom filtri."""
        zobrazene = set(self._tree.get_children())
        return [h for h in self._hlasenia
                if h["hid"] in zobrazene
                and self._checkboxy.get(h["hid"], tk.BooleanVar()).get()]

    def _spusti_sap(self):
        vybrane = self._vybrane()
        if not vybrane:
            messagebox.showinfo("Info", "Žiadne zobrazené hlásenie nie je zaškrtnuté.")
            return
        self._spusti_zoznam(vybrane, f"vybrané ({len(vybrane)})")

    def _nova_davka(self):
        """Vezme prvých N ešte nespracovaných hlásení (bez záznamu v logu)."""
        try:
            n = max(1, int(self._davka_var.get()))
        except (ValueError, TypeError):
            n = 50
        self._stav_log = nacitaj_stav_z_logu()
        nespracovane = [h for h in self._hlasenia
                        if self._stav_log.get(h["hid"]) is None]
        davka = nespracovane[:n]
        if not davka:
            messagebox.showinfo(
                "Info", "Žiadne nespracované hlásenia — všetko je hotové.\n"
                "(Chybné spustíš cez selekt Chybné → Vybrať všetky → "
                "Spustiť vybrané.)")
            return
        self._spusti_zoznam(
            davka, f"nová dávka ({len(davka)} z {len(nespracovane)} ešte "
                   "nespracovaných)")

    def _spusti_zoznam(self, zoznam: list[dict], popis: str):
        """Spoločné spustenie reálneho behu pre daný zoznam hlásení."""
        if not messagebox.askyesno(
                "Potvrdiť",
                f"Spustiť reálne nahrávanie — {popis}?\n"
                "(SAP musí byť otvorený a prihlásený)"):
            return
        self._log_clear()
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._log_pis(f"\n=== SAP NAHRÁVANIE — {popis} — {ts} ===\n")
        threading.Thread(
            target=self._beh, args=(zoznam, None, False), daemon=True
        ).start()

    def _beh(self, hlasenia: list[dict], _unused, dry_run: bool):
        session = None
        ok_count = 0
        skip_count = 0
        err_count = 0
        total = len(hlasenia)
        self.after(0, self._progress_set, 0, total, "pripájam sa…")

        if not dry_run:
            try:
                self.after(0, self._log_pis, "Pripájam sa na SAP GUI...\n")
                session = sap_spoj()
                self.after(0, self._log_pis, "SAP pripojený.\n")
            except NotImplementedError as e:
                self.after(0, self._log_pis, f"CHYBA: {e}\n")
                self.after(0, messagebox.showerror, "SAP chyba", str(e))
                return
            except Exception as e:
                msg = f"Chyba pripojenia na SAP: {type(e).__name__}: {e}"
                self.after(0, self._log_pis, msg + "\n")
                self.after(0, messagebox.showerror, "SAP chyba", msg)
                return

        for i, h in enumerate(hlasenia, start=1):
            def _log(text, _h=h):
                self.after(0, self._log_pis, text + "\n")

            self.after(0, self._progress_set, i - 1, total,
                       f"{i}/{total}: {h['hid']}")
            stav = spracuj_hlasenie(h, session, dry_run, _log)
            if stav == "ok":
                ok_count += 1
            elif stav == "skip":
                skip_count += 1
            else:
                err_count += 1
            self.after(0, self._progress_set, i, total)

        sufix = "(DRY RUN)" if dry_run else ""
        self.after(0, self._progress_set, total, total,
                   f"✓ HOTOVO ({total})")
        self.after(0, self._log_pis,
                   f"\n=== Hotovo {sufix}: {ok_count} OK, "
                   f"{skip_count} preskočených, {err_count} chýb ===\n")
        # prefarbi tabuľku podľa nového logu (len reálny beh, dry log nezapisuje výsledky)
        if not dry_run:
            self.after(0, self._obnov_farby)


# ============================================================
# === MAIN                                                ===
# ============================================================

def main():
    app = SapNahrajApp()
    app.mainloop()


if __name__ == "__main__":
    main()

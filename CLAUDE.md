# Projekt: servis-maily-automatizacia

Automatizácia opakovanej reklamačnej agendy: z veľkého Excelu reklamácií pripraviť
maily servisom, zozbierať a klasifikovať ich odpovede a (čiastočne) preniesť výsledok
do SAP. Pomocník pre operátorku podpory (v dokumentácii vystupuje ako „Eva", firma ako
„ACME") — človek rozhoduje, skript šetrí mechanickú prácu.

## Štyri fázy a kľúčové súbory

1. **Príprava mailov** — `rozdel_servisy.py` (rozdelí zdrojový Excel podľa servisu),
   `vyrob_maily_owa.py` (cez Playwright zakladá koncepty v Outlook Web). Staršie
   `.eml`/`.msg` varianty (`vyrob_maily.py`, `vyrob_maily_v2.py`, `_archiv/`) sú slepé
   uličky, ponechané ako ilustrácia.
2. **Zber odpovedí** — `stav.py` (stavový sklad), `stiahni_odpovede_owa.py` (Playwright
   sťahuje maily), `parsers.py` (heuristika + AI parser cez Claude API).
   `spracuj_odpovede.py` páruje a generuje pohľady (strom priečinkov + Excel).
3. **Klasifikácia AI agentom** — `priprav_klasifikaciu.py` (príprava cache),
   `.claude/agents/klasifikator-odpovedi.md` + `.claude/skills/` (subagent číta telo aj
   prílohy vrátane PDF/fotiek, zapíše `klasifikacia.json`). Skill `/klasifikuj-odpovede`.
4. **Zápis do SAP** — `sap_nahraj.py` (tkinter GUI, SAP GUI Scripting cez COM),
   `build_sap.bat` → `sap_nahraj.exe`.

## Aktuálny stav

- **Fázy 1-3 funkčné.** `stav_reklamacii.json` je jediná pravda o stave; Excel a strom
  priečinkov sa z neho len generujú.
- **Fáza 4 (SAP) zastavená** — závisí od serverového parametra `sapgui/user_scripting`,
  ktorý je vo firme po reštarte vypnutý a nedá sa eskalovať. Bez neho sa `sap_nahraj.exe`
  k SAP nepripojí. Kód je hotový (dedup, dva typy hlásení TAB09/GOS), len nemá prostredie.

## Konvencie a pravidlá

- **Komunikácia po slovensky** (viď globálny `~/.claude/CLAUDE.md`).
- **Dáta NIKDY na git.** `.gitignore` chráni Excely (`*.xlsx`), maily (`*.eml`, `maily/`,
  `odpovede/`), prílohy (`prilohy/`), stav (`stav_reklamacii.json`), log
  (`sap_nahraj_log.txt` — obsahuje mená firiem!), build artefakty a `owa_session/`.
- **Anonymizácia pred commitom.** Reálne mená firiem/osôb a čísla hlásení (`880…`),
  dobropisov (`Z…`) a dodávateľov (`3000…`) sa pred zverejnením nahrádzajú fiktívnymi,
  formátovo platnými hodnotami. Repozitár je verejný (portfólio) — žiadne reálne údaje.
- **Lokálna konfigurácia** je v `config.py` (ignorovaný); vzor je `config_example.py`.
  Obsahuje API kľúč pre Claude a údaje do podpisu.
- **Build SAP exe**: `build_sap.bat` (PyInstaller). Bat končí `pause` — pri spúšťaní
  neinteraktívne treba presmerovať vstup (napr. `< nul`).

## Gotchas

- **SAP GUI Scripting je krehké** — ID prvkov odčítané z nahrávok (`Script*.vbs`),
  správanie závisí od stavu obrazovky a časovania okien. Obranný štýl kódu (čakanie na
  okná, rekurzívne hľadanie prvkov, diagnostické logy) je zámerný.
- **Dedup je kritický** — príloha aj poznámka (prefix „LH: ") sa pred zápisom kontrolujú,
  aby opakovaný beh nič nezduplikoval.
- **Idempotencia naprieč fázami** — sťahovanie aj klasifikácia sa dajú spúšťať opakovane;
  hotové položky sa preskakujú (dedup podľa mail_id / existencie `klasifikacia.json`).
- Pri práci na Windows preferuj **forward-slash** cesty (Git Bash zje backslashe).

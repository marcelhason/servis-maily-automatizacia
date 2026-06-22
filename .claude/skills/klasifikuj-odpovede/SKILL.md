---
name: klasifikuj-odpovede
description: Spustí klasifikáciu stiahnutých odpovedí servisov — pripraví cache (priprav_klasifikaciu.py), na každý čakajúci mail pustí subagenta klasifikator-odpovedi a na záver zmerguje výsledky do stavu a excelu (spracuj_odpovede.py). Argumenty - voliteľne "--limit N" alebo konkrétne mail_id oddelené medzerou.
---

# Klasifikácia odpovedí servisov

Orchestruješ klasifikáciu mailov v `odpovede/<mail_id>/` pomocou subagenta
`klasifikator-odpovedi`. Výstupom behu je `klasifikacia.json` per mail
+ aktualizovaný stav a excel.

## Argumenty

- bez argumentov — klasifikuj všetky čakajúce maily
- `--limit N` — len prvých N čakajúcich (testovacie behy)
- zoznam mail_id oddelených medzerou — len tieto konkrétne maily

## Postup

1. Spusti `python priprav_klasifikaciu.py` (PYTHONIOENCODING=utf-8).
   Vypíše zoznam čakajúcich mailov (mail_id, dátum, servis). Ak nečaká nič,
   ohlás to a pokračuj rovno krokom 4.
2. Zostav zoznam na spracovanie podľa argumentov (všetky / --limit N /
   konkrétne id).
3. Pre KAŽDÝ mail spusti subagenta `klasifikator-odpovedi` s promptom:
   „Klasifikuj mail v priečinku odpovede/<mail_id>/" — **1 mail = 1 subagent**
   (izolácia kontextu; mail s mnohými PDF má desiatky tisíc tokenov).
   Spúšťaj 3–4 paralelne, po dobehnutí dávky ďalšiu. Súhrny subagentov
   priebežne hlás používateľovi (servis, mail_typ, kategórie, pochybnosti).
4. Po dobehnutí spusti znova `python priprav_klasifikaciu.py --zoznam`
   a over, že nič nečaká (ak áno, zopakuj krok 3 pre zvyšok; po druhom
   neúspechu sa zastav a vypíš, ktoré maily zlyhali).
5. Spusti `python spracuj_odpovede.py` — zmerguje klasifikácie do
   `stav_reklamacii.json` a pregeneruje excel (stĺpce Kategória / Číslo
   dobropisu / Akcia / Pokec pre SAP + hárok „Triáž").
6. Záverečný súhrn: počty podľa kategórií a akcií, koľko mailov skončilo
   `pozriet-rucne`, a pripomeň, že excel + strom priečinkov sa Eve
   kopírujú SPOLU.

## Poznámky

- Idempotencia: maily s hotovou `klasifikacia.json` sa preskakujú.
  Preklasifikovanie = zmazať `klasifikacia.json` daného mailu a spustiť znova.
- NDR a nezaradené maily majú v `kontext.json` pole `preskocit` —
  subagent ich neklasifikuje, to je v poriadku.
- Headless variant (opakované behy bez interakcie):
  `claude -p "/klasifikuj-odpovede" --permission-mode acceptEdits`

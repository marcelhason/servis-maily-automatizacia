# Hromadná príprava reklamačných mailov pre servisy

Automatizácia opakovanej kancelárskej úlohy: z jedného veľkého Excelu s reklamáciami
pripraviť pre **každý servis (dodávateľa)** samostatný e-mail so zhrnutím jeho prípadov
a priloženou tabuľkou — a založiť ich ako **koncepty priamo v Outlooku**, pripravené na
odoslanie ľudskou rukou. Skript **nikdy nič neodosiela sám**.

## Ako to funguje

Pipeline má dva kroky:

1. **`rozdel_servisy.py`** — načíta zdrojový Excel a rozdelí ho podľa servisu (stĺpec
   *Servis*) na samostatné `.xlsx` súbory, jeden na servis (vrátane prázdneho stĺpca
   *Poznámka servis*, ktorý dodávateľ vyplní).

2. **`vyrob_maily_owa.py`** — cez [Playwright](https://playwright.dev/python/) otvorí
   systémový Chrome, používateľ sa prihlási do svojho **Outlook Web**, a skript pre
   každý servis založí **koncept**: vyplní príjemcu, predmet, telo (s tabuľkou prípadov),
   priloží príslušný `.xlsx` a koncept uloží. Prihlásená relácia sa cachuje, takže ďalšie
   spustenia už prihlásenie nevyžadujú.

## Spracovanie odpovedí (druhá fáza)

Keď servisy začnú odpovedať (mix: vyplnená tabuľka v prílohe aj voľný text v tele),
nadväzuje pipeline odpovedí. Pravdou o stave komunikácie je **`stav_reklamacii.json`**
(append-only história poznámok per hlásenie + stavový automat *odoslané → odpovedané →
uzavreté* s poľom „na ťahu"); Excel sa z neho len generuje ako pohľad.

1. **`stav.py --init`** — jednorazovo založí stavový súbor zo zoznamu odoslaných hlásení.
2. **`stiahni_odpovede_owa.py`** — cez Playwright prejde priečinok odpovedí v Outlook Web
   (pod filtrom *Neprečítané*, schránku nemení) a každý mail uloží surový do
   `odpovede/<mail_id>/` (meta s dátumom prijatia a deep linkom, telo, prílohy).
   Iba sťahuje — parsovanie je oddelené. Opakované spustenie nič neduplikuje.
3. **`spracuj_odpovede.py`** — spáruje maily so servismi a prečíta ich dvoma vrstvami:
   deterministická heuristika v **`parsers.py`** (čísla hlásení v texte/prílohe; citácia
   pôvodného mailu sa orezáva) + **AI parser** (Claude API, číta aj odpovede vpísané do
   citovanej tabuľky a preposlané `.eml`). Výsledky sa zlúčia: zhoda oboch = vysoká
   istota; len-AI nález pod prahom `AI_MIN_CONFIDENCE` sa nezapisuje a mail ide na
   *ručnú kontrolu* aj s AI návrhmi — nič sa tichо neháda. AI odpovede sa cachujú
   (`ai_vysledok.json`), opakovaný beh API nevolá.

Výstupy sú stavané na prácu človeka (a kopírujú sa Eve **spolu**, odkazy sú relatívne):

- **`odpovede_podla_servisov/<SERVIS>/`** — čitateľné maily (`.txt` s hlavičkou
  Od / Prijaté / Predmet / deep link do OWA) a prílohy s dátumom v názve; na zálohu
  do SAPu stačí skopírovať priečinok servisu.
- **`... - stav.xlsx`** — zdroj + stĺpce Stav / Odpoveď servisu / Dátum / Od koho /
  Predmet mailu / klikací odkaz na mail v Outlooku / odkaz na priečinok / História
  s pôvodom každej poznámky; hárok *Prehľad servisov* (kto odpovedal, kto mlčí).

```bash
python stav.py --init                  # raz na začiatku
python stiahni_odpovede_owa.py         # stiahne odpovede (--limit 3 na test)
python spracuj_odpovede.py             # spracuje (heuristika+AI) + strom + excel
python spracuj_odpovede.py --len-excel # iba pregeneruje pohľady (bez AI)
python stav.py --prehlad               # rýchly súhrn do konzoly
python parsers.py                      # samotest heuristiky a zlučovania
```

## Klasifikácia odpovedí agentom (tretia fáza)

Vedieť, *že* servis odpovedal, nestačí — človek potrebuje vedieť, **ako reklamácia
dopadla a čo s ňou má v SAP urobiť**. Verdikt je pritom často len v PDF servisnom
liste alebo na fotke, kam textové parsery nedosiahnu. Túto vrstvu rieši
**Claude Code subagent** (`.claude/agents/klasifikator-odpovedi.md`) s doménovým
skillom (`.claude/skills/reklamacie-biznis/SKILL.md` — biznis proces, presné kritériá
kategórií, schéma výstupu, hraničné prípady):

1. **`priprav_klasifikaciu.py`** — deterministická príprava: rozbalí `.zip` prílohy
   (`_rozbalene/`), prepíše `.xlsx`/`.eml` na text (`_prepisy/`) a zapíše
   `kontext.json` (servis + jeho hlásenia), aby agent nemusel nič párovať.
2. **`/klasifikuj-odpovede`** (skill v Claude Code) — na každý mail pustí samostatného
   subagenta; ten prečíta telo aj **všetky prílohy vrátane PDF a obrázkov** a zapíše
   `odpovede/<id>/klasifikacia.json`: kategóriu per hlásenie (*dobropisované / mŕtve /
   oprava / nerentabilná oprava / nedostupný diel / zamietnuté / ping-pong*), číslo
   dobropisu, odporúčanú akciu a hotový **„pokec"** — slovenské zhrnutie na copy-paste
   do SAP poznámky. Žiadne hádanie: bez dôkazu v maile sa kategória neprideľuje.
3. **`spracuj_odpovede.py`** výsledky zmerguje do stavu (s validáciou enumov
   a krížovou kontrolou voči heuristike/AI vrstve) a do excelu pridá stĺpce
   *Kategória / Číslo dobropisu / Akcia / Pokec pre SAP* + hárok **„Triáž"** —
   pracovný zoznam zoskupený podľa akcií (hore „Žiadosť o dobropis v SAP").

```bash
python priprav_klasifikaciu.py            # príprava + zoznam čakajúcich mailov
# potom v Claude Code:  /klasifikuj-odpovede   (alebo s --limit N na test)
python spracuj_odpovede.py --len-klasifikacie # merge + excel bez AI vrstvy
python _porovnaj_klasifikaciu.py          # diag: agent vs AI vs heuristika
```

## Zápis do SAP (štvrtá fáza)

Keď je odpoveď klasifikovaná, treba ju preniesť do SAP reklamácie — pripojiť prílohy
od servisu a vložiť poznámku. Pri stovkách reklamácií je to časovo náročná, mechanická
robota. **`sap_nahraj.py`** ju zautomatizuje *čiastočne*: ku každému hláseniu, ku ktorému
máme podklady, nahrá prílohy a skrátenú poznámku — ale **samotné rozhodnutie, ako prípad
uzavrieť, ostáva na človeku**. Skript len ušetrí klikanie; verdikt nesie operátorka.

Je to desktopové GUI (tkinter) ovládajúce SAP GUI cez **SAP GUI Scripting (Windows COM)** —
SAP nemá použiteľné API, takže sa doň „kliká" programovo. Beží ako `.exe` (PyInstaller,
`build_sap.bat`) priamo na stanici, kde je SAP prihlásený.

Kľúčové vlastnosti:

- **Dva typy hlásení.** Bežné majú záložku príloh (TAB09); typy H2/H3 ju nemajú a prílohy
  sa k nim pripájajú cez **GOS toolbox** (Generic Object Services). Skript typ rozpozná
  a vetví sa — ID prvkov boli odčítané z nahrávok SAP GUI.
- **Žiadne duplicity.** Pred zápisom prečíta, čo už v SAP je, a prílohu ani poznámku
  nepridá druhýkrát — takže opakované spustenie (po páde, po doplnení) nič nezduplikuje.
- **Dávkovanie + farebný stav.** Položky sa spracúvajú po dávkach; stav (hotové / chybné /
  preskočené / ešte nespracované) sa číta z behového logu a v tabuľke je farebne odlíšený,
  takže je vidno, čo zostáva a kde nahrávanie zlyhalo.
- **Bezpečné preskočenie.** Hlásenia, ktoré operátorka už uzavrela ručne (označí ich
  zeleno v exceli), sa do SAPu vôbec neposielajú.

```bash
python sap_nahraj.py     # otvorí GUI (SAP musí byť spustený a prihlásený)
build_sap.bat            # zostaví sap_nahraj.exe (PyInstaller)
```

> Poznámka: SAP GUI Scripting je krehké prostredie — správanie závisí od stavu obrazovky
> (koľko príloh hlásenie má, časovanie otvárania okien). Najťažšou časťou bolo spoľahlivé
> čítanie existujúcich príloh a vetvenie podľa typu hlásenia; tomu zodpovedá obranný štýl
> kódu (čakanie na okná, rekurzívne hľadanie prvkov, diagnostické logy).

## Prečo cez Outlook Web (a nie `.eml`/`.msg`)

Cestou k riešeniu boli viaceré slepé uličky, ktoré dobre ilustrujú, prečo padla voľba na
ovládanie webového klienta:

- **`.eml` súbory** (skripty `vyrob_maily.py`, `vyrob_maily_v2.py`) — Outlook ich otváral
  ako *doručenú* správu, nie ako editovateľný koncept, a vnorené obrázky podpisu zlobili.
- **`.msg` cez Outlook COM** a **priame SMTP odosielanie** — ďalšie pokusy (v `_archiv/`),
  ktoré nevyhovovali požiadavke „pripravené koncepty, odoslanie potvrdí človek".
- **Outlook Web cez Playwright** — vytvára koncepty priamo v schránke, takže sa správajú
  presne ako ručne rozpísaný mail. Toto je výsledné, funkčné riešenie.

> Poznámka k webovej automatizácii: Microsoft v priebehu vývoja presunul Outlook Web na
> doménu `outlook.cloud.microsoft` a zmenil UI (napr. prílohy cez skrytý `input[type=file]`,
> ukladanie konceptu cez tlačidlo *Zavrieť* — *Esc* znamená *Zahodiť*). Skript tieto
> zvláštnosti rieši a je pripravený na ďalšie zmeny UI.

## Spustenie

```bash
pip install playwright openpyxl
playwright install chromium

python rozdel_servisy.py                 # rozdelí zdrojový Excel
python vyrob_maily_owa.py --test         # 1 náhodný servis (skúška)
python vyrob_maily_owa.py --sample 5     # 5 náhodných
python vyrob_maily_owa.py                # všetky servisy
python vyrob_maily_owa.py --only "Firma A s.r.o.;Firma B a.s."   # cielene
```

## Konfigurácia

Osobné údaje do podpisu sú mimo kódu. Skopíruj vzor a vyplň vlastné:

```bash
cp config_example.py config.py    # config.py je v .gitignore
```

## Dáta nie sú v repozitári

Zdrojové Excely, vygenerované súbory servisov, e-mailové vzory aj prihlásený profil
prehliadača sú v `.gitignore` — obsahujú reálne mená, e-maily a firmy. Repozitár obsahuje
**iba kód**. Na reálne spustenie treba dodať vlastné dáta a vzor tela mailu.

## Technológie

Python · Playwright (Chromium) · openpyxl · MIME/quoted-printable (pri `.eml` variante)

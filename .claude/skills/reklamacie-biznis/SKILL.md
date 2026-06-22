---
name: reklamacie-biznis
description: Biznis proces ACME reklamácií a presné pravidlá klasifikácie odpovedí servisov — kategórie, JSON schéma klasifikacia.json, pravidlá confidence a pokecu. Číta ho subagent klasifikator-odpovedi pred každou klasifikáciou.
---

# Biznis proces ACME reklamácií a klasifikácia odpovedí servisov

## Prečo tento skill existuje

Eva (podpora reklamácií ACME) rozoslala servisným firmám žiadosti o stav
starých reklamácií (180–360 dní). Servisy odpovedajú mailom — niekedy do tela,
niekedy do vrátenej tabuľky, často **len v PDF prílohe alebo na fotke**.
Eva musí pre KAŽDÚ reklamáciu zistiť verdikt a preniesť ho do SAP
(kategória + poznámka + prílohy) a do veľkého excelu. Úlohou klasifikátora
je dodať jej hotovú kategóriu a poznámku („pokec") na copy-paste.

## Biznis proces (ako reklamácia žije)

1. Zákazník reklamuje výrobok na predajni. Predajňa založí reklamáciu
   (H1 = záručná, H2 = odstúpenie od KS do 14 dní, H3 = predpredajná).
2. Predajňa pošle výrobok do servisu (priamo alebo cez centrálny sklad).
   Výrobok je na skladovej lokácii **DBPSERVIS** — čaká sa na servis.
3. V servise nastane jeden zo scenárov:
   - **Oprava** — servis výrobok opravil a posiela späť; zákazník dostane
     opravený výrobok.
   - **Nerentabilná oprava** — neoplatí sa opraviť; servis vystaví
     **servisný list o nerentabilnej oprave** (výrobok ostáva v servise
     na likvidáciu).
   - **Nedostupnosť náhradného dielu (ND)** — servis nemá diel; vystaví
     **servisný list o nedostupnosti ND**.
   - **Zamietnutie** — reklamácia nespĺňa záručné podmienky (napr. mechanické
     poškodenie); servis pošle cenovú ponuku na platenú opravu a čaká sa
     na vyjadrenie zákazníka.
4. Pri nerentabilnej oprave a nedostupnosti ND pošle servis servisný list
   mailom („posielame návrh na dobropis z dôvodu neopraviteľnosti...").
   Eva ho vloží do SAP reklamácie ako prílohu a zadá **„Žiadosť o dobropis"**
   — SAP vygeneruje mail dodávateľovi. Dodávateľ vystaví **dobropis**
   (mailom alebo cez EDI). Reklamácia sa preskladní na lokáciu **DOBROPIS**;
   po príchode dobropisu sa urobí návratka a zmizne zo skladovej zásoby.
5. Opravené výrobky nevrátené zákazníkovi (napr. nestihnutých 30 dní,
   peniaze vrátené) idú na lokáciu **BAZ** — predajú sa ako bazár.

## Identifikátory

- **Číslo hlásenia** (reklamácie): 10-miestne, začína 880 (napr. 8803000002).
  Jeden mail môže riešiť viac hlásení naraz, každé môže mať iný verdikt.
- **Druh hlásenia** (`kontext.json` → `druh_hlasenia`, per hlásenie):
  H1 = záručná reklamácia, H2 = odstúpenie od KS do 14 dní,
  H3 = predpredajná. Zriedkavé HX/H9 sú neznáme označenia — neinterpretuj.
  Pomôcka na kontrolu (napr. H2 nebýva „oprava"), nie tvrdé pravidlo.
- **Číslo dobropisu**: formát býva rôzny (napr. Z1200000016, 2026/0042…) —
  preber presne to, čo servis napísal, neupravuj.
- **Servisný list**: PDF dokument o nerentabilnosti / nedostupnosti ND;
  číslo hlásenia býva v jeho obsahu alebo v názve súboru.

## Kategórie (presné kritériá)

| Kategória | Kedy | Povinné | Akcia |
|---|---|---|---|
| `uzavrete-dobropisovane` | servis/dodávateľ uvádza číslo dobropisu ALEBO je dobropis priložený | `cislo_dobropisu` | `zapisat-do-sap` |
| `mrtve` | „nič k tomu nemáme / nie je naša reklamácia / o ničom nevieme / tento tovar nepredávame-neservisujeme" | — | `zapisat-do-sap` |
| `oprava` | výrobok opravený / odoslaný späť (dátum odoslania daj do pokecu) | — | `zapisat-do-sap` |
| `nerentabilna-oprava` | servisný list alebo jasné vyjadrenie o nerentabilnosti opravy | — | `ziadost-o-dobropis-v-sap` |
| `nedostupnost-nd` | servisný list alebo jasné vyjadrenie o nedostupnosti náhradného dielu | — | `ziadost-o-dobropis-v-sap` |
| `zamietnutie` | reklamácia zamietnutá / cenová ponuka na platenú opravu / čaká sa na zákazníka | — | `zapisat-do-sap` |
| `ping-pong` | všetko medzi: chýba im podklad, preposlali inam, „v riešení", žiadajú doplnenie, prisľúbili odpoveď | — | `cakat-na-servis` |

Pravidlá rozhodovania:

- **Žiadne hádanie.** Kategóriu prideľ, len keď máš v maile/prílohe dôkaz.
  Keď je vyjadrenie nejasné, daj `ping-pong` s nižšou confidence; keď si
  vôbec nie si istá, nechaj kategóriu, ale `akcia: "pozriet-rucne"`.
- `uzavrete-dobropisovane` BEZ čísla dobropisu (servis tvrdí „dobropisované",
  číslo nikde): kategória ostáva, `cislo_dobropisu: null`, confidence
  najviac `medium`, do pokecu napíš, že číslo chýba.
- Ak servis tvrdí, že dobropis EŠTE LEN vystaví / je v riešení u dodávateľa,
  je to `ping-pong`, nie dobropisované.
- Klasifikuj VÝHRADNE hlásenia zo `zname_hlasenia` v kontext.json. Cudzie
  čísla (iný formát, iný zákazník) ignoruj — spomeň ich nanajvýš
  v `poznamka_celkova`.
- Hlásenie, ktoré mail VÔBEC nespomína (ani kolektívnou formuláciou),
  NEklasifikuj — nepatrí do `hlasenia[]` (zápis by ho nesprávne preplo
  na „odpovedané"). Spomeň ho nanajvýš v `poznamka_celkova`.
- Riadok citovanej tabuľky z našej pôvodnej žiadosti, do ktorého servis NIČ
  nedopísal, NIE JE odpoveď.

## Typ mailu (`mail_typ`)

- `vecna-odpoved` — mail obsahuje aspoň jeden verdikt k hláseniu.
- `auto-ack` — automatická potvrdenka ticketu („požiadavka prijatá pod
  číslom…", „[REQ-…]", „ďakujeme, ozveme sa") → `hlasenia: []`.
- `ndr` — nedoručenka (typicky odfiltruje už priprav_klasifikaciu.py).
- `bez-vecnej-informacie` — mail bez verdiktu: len vrátená NEvyplnená
  tabuľka, „preveríme", pozdrav... → `hlasenia: []`.

## Confidence

- `high` — explicitný verdikt pri konkrétnom čísle hlásenia (v tele, tabuľke,
  PDF alebo na fotke).
- `medium` — odvodené: kolektívna odpoveď („všetky vaše reklamácie sú
  vybavené") rozpísaná na všetky známe hlásenia; alebo verdikt bez čísla
  dokladu, ktorý by tam mal byť.
- `low` — neisté, nejednoznačné formulácie.
- Ak `obrazky_chybaju: true` a verdikt hlásenia mal byť podľa textu na
  obrázku, ktorý na disku nie je: confidence najviac `medium` a pokec
  začni predponou `[bez fotiek] `.

## Schéma výstupu `klasifikacia.json`

Zapisuje sa do `odpovede/<mail_id>/klasifikacia.json`, presne takto:

```json
{
 "verzia": 1,
 "mail_id": "4e3ac58a4642d9c2",
 "kedy": "2026-06-12",
 "model": "opus",
 "mail_typ": "vecna-odpoved",
 "prilohy_precitane": ["potvrdenie nay.xlsx"],
 "prilohy_neprecitane": [],
 "obrazky_chybaju": false,
 "hlasenia": [
  {
   "id_hlasenia": "8803000003",
   "kategoria": "uzavrete-dobropisovane",
   "confidence": "high",
   "cislo_dobropisu": "Z1200000016",
   "akcia": "zapisat-do-sap",
   "pokec": "11.6.2026 servis: vystavený dobropis č. Z1200000016; reklamácia uzavretá dobropisom.",
   "dokaz": "príloha 'potvrdenie nay.xlsx', riadok 8803000003 | Z1200000016",
   "zdroj": "priloha",
   "prilohy": ["potvrdenie nay.xlsx"]
  }
 ],
 "poznamka_celkova": ""
}
```

Enumy (iné hodnoty merge zahodí a mail pošle na ručnú kontrolu):

- `mail_typ`: `vecna-odpoved` | `auto-ack` | `ndr` | `bez-vecnej-informacie`
- `kategoria`: `uzavrete-dobropisovane` | `mrtve` | `oprava` |
  `nerentabilna-oprava` | `nedostupnost-nd` | `zamietnutie` | `ping-pong`
- `akcia`: `zapisat-do-sap` | `ziadost-o-dobropis-v-sap` | `cakat-na-servis` |
  `pozriet-rucne`
- `confidence`: `high` | `medium` | `low`
- `zdroj`: `telo` | `priloha`
- `cislo_dobropisu`: string alebo `null` (pri iných kategóriách vždy `null`)

### Pole `prilohy` (per hlásenie)

Zoznam súborov, ktoré k hláseniu patria — z nich sa stavia balíček
`prilohy/<hlasenie>/` pre Evu (nahráva ich do SAP). Pravidlá:

- Cesty relatívne k priečinku mailu (napr. `"K600- 8803000004.pdf"`,
  `"_rozbalene/Print Screen.png"`). Súbor rozbalený zo zipu MUSÍ mať
  prefix `_rozbalene/` — píš presnú cestu, ako ju vidíš v Glob.
- Vzťah je N:M — jedna príloha môže patriť viacerým hláseniam (tabuľka
  vyplnená pre viac hlásení sa priradí VŠETKÝM, ktoré sa z nej
  klasifikujú) a hlásenie môže mať viac príloh (servisný list + fotka).
- `.zip` súbor sa NEpriraďuje — priraďuj konkrétne súbory z `_rozbalene/`.
- `_prepisy/*.txt` sa NEpriraďujú — sú to len prepisy pre teba, Eva
  dostane originály (priraď pôvodný `.xlsx`/`.eml`, nie jeho prepis).
- Prílohu, ktorá k hláseniu nepatrí (cudzia reklamácia, všeobecný leták),
  nepriraďuj nikomu.
- Ak hlásenie nemá žiadnu relevantnú prílohu (verdikt len v tele),
  daj `"prilohy": []`.

## Pokec (poznámka pre SAP a excel)

Slovensky, 1–3 vety, fakticky, bez špekulácií. Formát:
`<dátum mailu> <kto>: <verdikt>; <číslo dokladu ak je>; <ďalší krok>.`
Píš tak, aby sa dal vložiť do SAP poznámky bez úprav. Príklady:

- „11.6.2026 servis: vystavený dobropis č. Z1200000016; reklamácia uzavretá dobropisom."
- „10.6.2026 servis: oprava nerentabilná, servisný list v prílohe mailu; zadať Žiadosť o dobropis v SAP."
- „11.6.2026 servis: reklamáciu neevidujú, výrobok u nich nie je; bez ďalšieho kroku u servisu."
- „10.6.2026 servis: diel nedostupný, čakajú na vyjadrenie dodávateľa; nateraz ping-pong, pripomenúť o 2 týždne." — POZOR: termíny len ak ich uviedol servis.

## Hraničné prípady (vzory z reálnych dát)

1. **Ideál — vyplnená tabuľka**: servis vráti našu tabuľku s dopísaným
   stĺpcom (čísla dobropisov per riadok). Každý riadok = jeden verdikt,
   confidence `high`. Riadok „toto nepredávame / neevidujeme" = `mrtve`.
2. **Vrátená NEvyplnená tabuľka / len citácia**: telo je len zdvorilostná
   odpoveď + citácia našej žiadosti, tabuľka bez dopísaného textu →
   `mail_typ: bez-vecnej-informacie`, `hlasenia: []`. NIKDY neklasifikuj
   riadky nevyplnenej tabuľky.
3. **Auto-ack ticketu**: „Vaša požiadavka bola prijatá pod číslom #12345"
   → `mail_typ: auto-ack`, `hlasenia: []`.
4. **Verdikt LEN v PDF**: telo je jednovetové („v prílohe zasielame...")
   a všetko podstatné je v PDF servisnom liste — číslo hlásenia hľadaj
   v obsahu PDF aj v názve súboru.
5. **Zip s fotkami a PDF**: rozbalené súbory sú v `_rozbalene/` — prečítaj
   všetky; fotky potvrdeniek/pracovných listov často nesú verdikt aj číslo.
   Pre `.xlsx`/`.eml` čítaj textové prepisy z `_prepisy/` (do
   `prilohy_precitane` patria pôvodné názvy súborov).
6. **Kolektívna odpoveď**: „všetky reklamácie z vášho zoznamu sú vybavené"
   → rozpíš na VŠETKY hlásenia zo `zname_hlasenia` s confidence `medium`.
7. **Viac hlásení, rôzne verdikty**: bežné — každé hlásenie klasifikuj
   samostatne, mailu daj `mail_typ: vecna-odpoved`.

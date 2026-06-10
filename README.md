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

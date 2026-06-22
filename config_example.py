"""
Vzor lokalnej konfiguracie. Skopiruj tento subor na `config.py` a vypln
vlastne udaje:

    copy config_example.py config.py     (Windows)
    cp   config_example.py config.py     (Linux/Mac)

`config.py` je v .gitignore a NEcommituje sa (obsahuje osobne udaje do podpisu).
Skripty importuju `config` a ak `config.py` neexistuje, pouziju tieto vzorove
(placeholder) hodnoty.
"""

# Meno a pozicia do podpisu plain-textovej casti mailu.
SIGNATURE_NAME = "Meno Priezvisko"
SIGNATURE_ROLE = "Pozícia / oddelenie"

# Riadky paticky firmy (kazdy je samostatny riadok podpisu). Format odkazu
# "text<url>" je konvencia plain-text mailu.
SIGNATURE_COMPANY_LINES = [
    "Firma, a. s. | Ulica 1 | 000 00 Mesto |",
    "www.firma.sk<http://www.firma.sk/>",
    "facebook<https://www.facebook.com/firma/>",
]

# --- AI parser odpovedí (parsers.parse_reply_ai) ---

# API kluc z https://console.anthropic.com/ — bez neho sa AI parser preskoci
# a bezi len heuristika.
ANTHROPIC_API_KEY = ""

# Model na citanie odpovedi servisov.
AI_MODEL = "claude-sonnet-4-6"

# Minimalna confidence, pri ktorej sa len-AI vysledok (heuristika ho nenasla)
# zapise ako odpoved. Pod prahom ide mail na rucnu kontrolu aj s AI navrhmi.
# Hodnoty: "low" | "medium" | "high"
AI_MIN_CONFIDENCE = "medium"

# Deep link na konverzaciu v OWA ({id} = url-encoded convid). Tvar zisti
# z URL prehliadaca po otvoreni lubovolneho mailu v danom priecinku:
# .../mail/<id priecinka>/id/<id mailu> — sem patri vsetko PRED "/id/".
OWA_LINK_VZOR = ""

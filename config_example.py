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

# Mamut naživo

Statistiky koncertů kapely [Pozor Mamut](https://pozormamut.cz/) (punk, Jihlava) od roku 2016.

Stránka `index.html` je vygenerovaná skriptem `make.py` z kapelní Google tabulky s evidencí koncertů
(stáhne ji přes `rclone`). Úpravy vzhledu patří do `template.html`, ne do `index.html`.

```
python3 make.py            # stáhne tabulku a sestaví index.html
```

Souřadnice obcí: [GeoNames](https://www.geonames.org/) (`data/CZ.txt`, CC BY 4.0, v repu není –
stáhnout z https://download.geonames.org/export/dump/CZ.zip). Obrys ČR: Natural Earth (public domain).

## Odkud se berou materiály u koncertů

Skript páruje soubory na kapelním Disku podle data `RRRR-MM-DD` na začátku názvu:

- `Koncerty/RRRR/RRRR-MM-DD Místo, akce – plakát (autor).jpg` → zmenšenina v `obrazky/` (veřejná),
- `Písničky/Playlisty/RRRR-MM-DD Místo, akce – playlist.txt` → pořadí písniček (jedna na řádek,
  čísla na začátku se odstraní, řádek začínající `#` je poznámka); fotka papíru `….jpg` → `obrazky/`,
- `Fotky a videa/RRRR/RRRR-MM-DD Místo, akce/` → odkaz na složku na Disku (jen pro kapelu).
- `Písničky/Texty/Název písničky.docx` → odkaz „text“ u písničky (jen pro kapelu). Text může být
  dokument (`.docx`, otevře se v Dokumentech), `.md`, `.txt`, `.pdf` i obrázek (`.jpg`, `.png`, třeba
  fotka papíru s textem, otevře se v náhledu Disku). Slovo „text“ za názvem se ignoruje
  (`Pozitivní myšlení text.jpg`). Datovaná verze `RRRR-MM-DD Název…` má přednost, mezi nedatovanými
  dokument před ostatními. Podsložka `Koncepty z přepisu (k opravě)/` → odkaz „koncept textu“.

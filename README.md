# Mamut naživo

Statistiky koncertů kapely [Pozor Mamut](https://pozormamut.cz/) (punk, Jihlava) od roku 2016.

Stránka `index.html` je vygenerovaná skriptem `make.py` z kapelní Google tabulky s evidencí koncertů
(stáhne ji přes `rclone`). Úpravy vzhledu patří do `template.html`, ne do `index.html`.

```
python3 make.py            # stáhne tabulku a sestaví index.html
```

Souřadnice obcí: [GeoNames](https://www.geonames.org/) (`data/CZ.txt`, CC BY 4.0, v repu není –
stáhnout z https://download.geonames.org/export/dump/CZ.zip). Obrys ČR: Natural Earth (public domain).

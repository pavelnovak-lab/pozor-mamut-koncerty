#!/usr/bin/env python3
"""Sestaví index.html se statistikami koncertů Pozor Mamut.

Zdroje:
  - Google tabulka „Koncerty/Pozor Mamut – evidence koncertů“ (stáhne se přes rclone jako xlsx),
  - souhrny cest z ~/Dropbox/Kapela/pozor-mamut-cesty.tsv (soukromé; na stránku jdou
    jen součty a rekordy, žádné časy ani místa jednotlivých cest),
  - GeoNames CZ.txt (souřadnice obcí), data/cz_obrys.json (obrys ČR, Natural Earth).
Použití:  python3 make.py [--xlsx soubor.xlsx] [--geonames CZ.txt]
"""
import argparse, base64, collections, csv, datetime as dt, hashlib, io, json, math, os, re, subprocess, sys
import openpyxl

TU = os.path.dirname(os.path.abspath(__file__))
HOME = os.path.expanduser('~')
REMOTE = 'gdrive,root_folder_id=1Ns75usanWOJoy2fHk7KD8EFYv58sNub_:Koncerty/Pozor Mamut – evidence koncertů.xlsx'
CESTY = HOME + '/Dropbox/Kapela/pozor-mamut-cesty.tsv'
LOGO = TU + '/data/logo_1barva.svg'  # 2b bílá plná v barvě písma (currentColor); výstražný trojúhelník z 2a v --logo-troj s vnitřním lemem, okno, oči a kly v barvě písma, obrysy v --band

ap = argparse.ArgumentParser()
ap.add_argument('--xlsx'); ap.add_argument('--geonames', default=TU + '/data/CZ.txt')
a = ap.parse_args()

raw = open(a.xlsx, 'rb').read() if a.xlsx else subprocess.run(['rclone', 'cat', REMOTE], check=True, capture_output=True).stdout
wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True)

def datum(v):
    if isinstance(v, dt.datetime): return v.date()
    if isinstance(v, dt.date): return v
    s = str(v or '').strip()
    m = re.match(r'(\d{4})-(\d{2})-(\d{2})', s)
    return dt.date(*map(int, m.groups())) if m else s  # neúplné datum zůstane textem

# --- koncerty ---
ws = wb['Koncerty']; H = [c.value for c in ws[1]]
kon, plan = [], []
for r in ws.iter_rows(min_row=2, values_only=True):
    x = dict(zip(H, r))
    if not x.get('Datum'): continue
    d = datum(x['Datum']); stav = (x.get('Stav') or '').strip()
    rec = dict(d=d, mesto=(x.get('Město') or '').strip(), akce=(x.get('Klub / akce') or '').strip(),
               kapely=[k.strip() for k in re.split(r',\s*', x.get('Hráli s námi') or '') if k.strip()],
               sestava=(x.get('Sestava / kdo chyběl') or '').strip(),
               odkazy=[dict(t=a.strip(), url=b.strip()) for a, b in (o.split('|', 1) for o in re.split(r';\s*', x.get('Odkazy') or '') if '|' in o)])
    if stav == 'odehráno': kon.append(rec)
    elif stav == 'plánováno': plan.append(rec)

def rok(r): return r['d'].year if isinstance(r['d'], dt.date) else int(str(r['d'])[:4])

# --- obce: souřadnice z GeoNames ---
# Jmenovců je hodně (Švihov 5×, Mostiště 2×). Přednost má obec v kraji Vysočina (kód 80),
# kde kapela hraje nejčastěji; výjimky mimo Vysočinu určuje okres (kód GeoNames admin2).
OKRES = {'Švihov': '0531',   # Švihov u Miřetic, okres Chrudim
         'Opatov': '0632',   # Opatov u Jihlavy (podle časové osy)
         'Nemojov': '0525',  # (Dolní) Nemojov u Dvora Králové – časová osa: ze zkušebny přes Brod, Chrudim, Hradec, přespání, návrat do zkušebny
         'Chválkov': '0633'}  # Chválkov u Kamenice nad Lipou (Kravín fest)
obce = {}
def lepsi(n, nove, stare):
    if n in OKRES: return nove[4] == OKRES[n] and stare[4] != OKRES[n]
    if (nove[3] == '80') != (stare[3] == '80'): return nove[3] == '80'
    return nove[2] > stare[2]
for r in csv.reader(open(a.geonames, encoding='utf-8'), delimiter='\t', quoting=csv.QUOTE_NONE):
    if r[6] != 'P': continue
    x = (float(r[4]), float(r[5]), int(r[14] or 0), r[10], r[11])
    if r[1] not in obce or lepsi(r[1], x, obce[r[1]]): obce[r[1]] = x
obce.setdefault('Praha', (50.0880, 14.4208, 1300000))
obce.setdefault('Rounek', (49.4275, 15.4790, 0))  # část Vyskytné nad Jihlavou, v GeoNames chybí
def cisti_mesto(m):
    m = re.sub(r'\s*\(.*?\)', '', m).strip()
    return {'Škrovád': 'Slatiňany', 'Horní Pohleď': 'Pohleď', 'Vyskytná nad Jihlavou': 'Vyskytná nad Jihlavou'}.get(m, m)

mista = collections.OrderedDict()
for r in kon:
    m = cisti_mesto(r['mesto'])
    if not m: continue
    mista.setdefault(m, 0); mista[m] += 1
bez_souradnic = [m for m in mista if m not in obce]
if bez_souradnic: print('bez souřadnic:', bez_souradnic, file=sys.stderr)
mapa = [dict(m=m, n=n, la=obce[m][0], lo=obce[m][1]) for m, n in mista.items() if m in obce]

# --- kapely ---
ALIAS = {'bermudskéj kvádr': 'Bermudskej kvádr', 'bermudský kvádr': 'Bermudskej kvádr', 'bermudskej kvádr': 'Bermudskej kvádr',
         'hak mak': 'Hakmak', 'hakmak': 'Hakmak', 'exploze jádra': 'Exploze jádra',
         'paní kolombová': 'Mrs Columbo', 'paní colombová': 'Mrs Columbo', 'mrs columbo': 'Mrs Columbo', 'mrs. columbo': 'Mrs Columbo'}
kap = collections.Counter(); kap_jmeno = {}
for r in kon:
    for k in r['kapely']:
        k0 = re.sub(r'\s*\(.*?\)', '', k).strip()
        key = k0.lower()
        jm = ALIAS.get(key, k0)
        kap[jm.lower()] += 1; kap_jmeno.setdefault(jm.lower(), jm)
kapely = [dict(k=kap_jmeno[k], n=n) for k, n in kap.most_common() if n >= 2]

# --- roky, dny, měsíce ---
roky = collections.Counter(rok(r) for r in kon)
r0, r1 = min(roky), max(roky)
po_letech = [dict(r=y, n=roky.get(y, 0)) for y in range(r0, r1 + 1)]
DNY = ['po', 'út', 'st', 'čt', 'pá', 'so', 'ne']
dny = collections.Counter(r['d'].weekday() for r in kon if isinstance(r['d'], dt.date))
po_dnech = [dict(d=DNY[i], n=dny.get(i, 0)) for i in range(7)]

# --- materiály na Disku: plakáty, playlisty, fotky, videa (párují se podle data RRRR-MM-DD v názvu) ---
ROOT = 'gdrive,root_folder_id=1Ns75usanWOJoy2fHk7KD8EFYv58sNub_:'
def lsjson(cesta):
    out = subprocess.run(['rclone', 'lsjson', '-R', '--no-modtime', ROOT + cesta], check=True, capture_output=True, text=True).stdout
    return json.loads(out)
DATUM = re.compile(r'(\d{4}-\d{2}-\d{2})')
OBR = re.compile(r'\.(jpe?g|png|webp|heic|pdf|gif)$', re.I); VID = re.compile(r'\.(mp4|mov|avi|mkv|m4v)$', re.I)
mat = collections.defaultdict(lambda: dict(plakaty=[], playlist=[], fotky=None, videa=None, pozvanky=[], pisne=[], zdroj={}))
os.makedirs(TU + '/obrazky', exist_ok=True)
def nahled(cesta, cil, sirka):
    """Zmenšenina souboru z Disku do repa (veřejná); stáhne se jen jednou."""
    if os.path.exists(TU + '/' + cil): return cil
    data = subprocess.run(['rclone', 'cat', ROOT + cesta], check=True, capture_output=True).stdout
    if cesta.lower().endswith('.gif'):  # animovaný plakát → animované WebP
        subprocess.run(['convert', 'gif:-', '-coalesce', '-resize', f'{sirka}x{sirka}>', '-quality', '70', TU + '/' + cil], input=data, check=True)
        return cil
    vstup = '-[0]' if cesta.lower().endswith(('.pdf', '.heic')) else '-'
    subprocess.run(['convert', '-density', '110', vstup, '-auto-orient', '-resize', f'{sirka}x{sirka}>', '-quality', '82', TU + '/' + cil],
                   input=data, check=True)
    return cil
def otisk(f):  # cache náhledu podle celé cesty a velikosti: přejmenování nebo nová verze = nový náhled
    return hashlib.md5(f"{f['Path']}|{f['Size']}".encode()).hexdigest()[:8]
def bezpecne(s): return re.sub(r'[^a-z0-9]+', '-', s.lower().encode('ascii', 'ignore').decode()).strip('-')[:60]

for f in lsjson('Koncerty'):
    if f['IsDir'] or 'Pracovní' in f['Path'] or 'Videopozvánky' in f['Path'] or not OBR.search(f['Name']): continue
    m = DATUM.match(f['Name']) or DATUM.search(f['Path'].split('/')[1] if '/' in f['Path'] else '')
    if not m or not re.search(r'plakát|program|line-up|plakat', f['Path'], re.I): continue
    d = m.group(1); jm = f'{d}-plakat-{bezpecne(f["Name"])[:40]}-{otisk(f)}'
    mat[d]['plakaty'].append(dict(mal=nahled('Koncerty/' + f['Path'], f'obrazky/{jm}-m.webp', 360),
                                  vel=nahled('Koncerty/' + f['Path'], f'obrazky/{jm}.webp', 1400), gif=f['Name'].lower().endswith('.gif'), src=f['Path']))
for m in mat.values():  # když existuje animovaný plakát, ukáže se jen on (ostatní jsou jeho statické verze)
    if any(p['gif'] for p in m['plakaty']): m['plakaty'] = [p for p in m['plakaty'] if p['gif']]
    for p in m['plakaty']: p.pop('gif')
# videopozvánky: Koncerty/Videopozvánky/RRRR-MM-DD Místo, akce – videopozvánka (autor) NN.mp4
for f in sorted(lsjson('Koncerty/Videopozvánky'), key=lambda f: f['Name']):
    m = DATUM.match(f['Name'])
    if f['IsDir'] or not m or not VID.search(f['Name']): continue
    mat[m.group(1)]['pozvanky'].append(dict(url='https://drive.google.com/file/d/{}/view'.format(f['ID']), src=f['Name']))
for f in lsjson('Písničky/Playlisty'):
    if f['IsDir']: continue
    for d in DATUM.findall(f['Name']):
        if f['Name'].lower().endswith('.txt'):
            text = subprocess.run(['rclone', 'cat', ROOT + 'Písničky/Playlisty/' + f['Path']], check=True, capture_output=True, text=True).stdout
            mat[d]['pisne'] = [re.sub(r'^\s*\d+\s*[.)/:-]?\s*', '', l).strip() for l in text.splitlines()
                               if l.strip() and not l.lstrip().startswith('#')]
        elif re.search(r'\.(jpe?g|png|webp)$', f['Name'], re.I):
            jm = f'{d}-playlist-{bezpecne(f["Name"])[:40]}-{otisk(f)}'
            mat[d]['playlist'].append(dict(vel=nahled('Písničky/Playlisty/' + f['Path'], f'obrazky/{jm}.webp', 1400), src=f['Name']))
DRIVE_F = 'https://drive.google.com/file/d/{}/view'; DRIVE_D = 'https://drive.google.com/drive/folders/{}'
fv = lsjson('Fotky a videa')
for f in fv:
    casti = f['Path'].split('/')
    m = DATUM.match(casti[-1])
    if not m: continue
    d = m.group(1)
    if f['IsDir'] and len(casti) == 2:  # složka akce: RRRR/RRRR-MM-DD Místo, akce/
        uvnitr = [g['Name'] for g in fv if g['Path'].startswith(f['Path'] + '/')]
        if any(OBR.search(x) for x in uvnitr): mat[d]['fotky'] = DRIVE_D.format(f['ID']); mat[d]['zdroj']['fotky'] = casti[-1]
        if any(VID.search(x) for x in uvnitr): mat[d]['videa'] = DRIVE_D.format(f['ID']); mat[d]['zdroj']['videa'] = casti[-1]
    elif not f['IsDir'] and len(casti) == 2:  # jednotlivý soubor přímo v roce
        klic = 'videa' if VID.search(f['Name']) else 'fotky' if OBR.search(f['Name']) else None
        if klic and not mat[d][klic]: mat[d][klic] = DRIVE_F.format(f['ID']); mat[d]['zdroj'][klic] = casti[-1]
# co se na web nedostane: v roční složce bez celého data na začátku názvu
# (záměrně nezařazené „RRRR-MM …“ a Zkušebna a studio se nehlásí)
nezarazene = sorted({'/'.join(f['Path'].split('/')[:2]) for f in fv
                     if re.fullmatch(r'\d{4}', f['Path'].split('/')[0]) and len(f['Path'].split('/')) >= 2
                     and not DATUM.match(f['Path'].split('/')[1]) and not re.match(r'\d{4}-\d{2} ', f['Path'].split('/')[1])})

# --- playlisty (pořadí písniček na koncertě): list v tabulce + TXT soubory v Písničky/Playlisty ---
sl = collections.Counter(); sl_koncerty = set(); z_tabulky = collections.defaultdict(list)
listy = [n for n in ('Playlisty', 'Setlisty') if n in wb.sheetnames]
if listy:
    w2 = wb[listy[0]]; H2 = [c.value for c in w2[1]]
    for r in w2.iter_rows(min_row=2, values_only=True):
        x = dict(zip(H2, r))
        if x.get('Písnička'):
            z_tabulky[str(datum(x.get('Datum')))].append((x.get('Pořadí') or 0, x['Písnička'].strip()))
for d, p in z_tabulky.items():
    if not mat[d]['pisne']: mat[d]['pisne'] = [s for _, s in sorted(p)]
for d, m in mat.items():
    if m['pisne']:
        sl_koncerty.add(d)
        for s in m['pisne']: sl[s] += 1
pisne = [dict(p=p, n=n) for p, n in sl.most_common()]

# --- cesty: jen souhrny ---
cesty = dict(km=None)
if os.path.exists(CESTY):
    c = list(csv.DictReader(open(CESTY, encoding='utf-8'), delimiter='\t'))
    km = lambda x: (int(x['km_tam'] or 0), int(x['km_zpet'] or 0))
    tot = sum(sum(km(x)) for x in c)
    nej = max(c, key=lambda x: km(x)[0])
    benz = collections.Counter()
    for x in c:
        for z in re.findall(r'⛽ ([^,]+),', x['zastavky_tam'] + ';' + x['zastavky_zpet']): benz[z.strip()] += 1
    def odjezd_h(x):  # odjezd z koncertu v hodinách po poledni dne koncertu
        m = re.match(r'(\d+)\. (\d+):(\d+)', x['odjezd_z_koncertu'])
        if not m: return None
        den, h, mi = map(int, m.groups())
        return (h + mi / 60) + (24 if den != int(x['datum'][8:10]) else 0)
    pozde = max((x for x in c if odjezd_h(x) and odjezd_h(x) < 31), key=odjezd_h)
    def jm(x): return x['koncert'].split(' – ')[0] + ', ' + dt.date.fromisoformat(x['datum']).strftime('%-d. %-m. %Y')
    h = odjezd_h(pozde) % 24
    cesty = dict(km=round(tot, -2), n=len(c), nej_km=km(nej)[0] + km(nej)[1], nej=jm(nej),
                 pozde=f'{int(h)}:{int(round(h % 1 * 60)):02d}', pozde_kde=jm(pozde),
                 benz=benz.most_common(1)[0][0] if benz else None, benz_n=sum(benz.values()))
    # pro legraci: do kterého velkého města je z Jihlavy vzdušnou čarou podobně daleko
    MESTA = [('Lisabonu', 38.72, -9.14), ('Reykjavíku', 64.15, -21.94), ('Moskvy', 55.76, 37.62), ('Istanbulu', 41.01, 28.98),
             ('Káhiry', 30.04, 31.24), ('Teheránu', 35.69, 51.39), ('Dubaje', 25.20, 55.27), ('Almaty', 43.24, 76.89),
             ('Lagosu', 6.52, 3.38), ('Dillí', 28.61, 77.21), ('Nairobi', -1.29, 36.82), ('Bombaje', 19.08, 72.88),
             ('New Yorku', 40.71, -74.01), ('Montrealu', 45.50, -73.57), ('Pekingu', 39.90, 116.40), ('Bangkoku', 13.76, 100.50),
             ('Kapského Města', -33.92, 18.42), ('Tokia', 35.68, 139.69), ('Rio de Janeira', -22.91, -43.17), ('Los Angeles', 34.05, -118.24)]
    def vzd(la, lo, la2=49.3961, lo2=15.5912):
        f1, f2, dl = map(math.radians, (la, la2, lo2 - lo))
        return 6371 * math.acos(min(1, math.sin(f1) * math.sin(f2) + math.cos(f1) * math.cos(f2) * math.cos(dl)))
    mesto, km_m = min(((m, vzd(la, lo)) for m, la, lo in MESTA), key=lambda t: abs(t[1] - tot))
    mesto2, km_m2 = min(((m, vzd(la, lo)) for m, la, lo in MESTA), key=lambda t: abs(2 * t[1] - tot))
    cesty.update(srov=mesto, srov_km=int(round(km_m, -2)), srov2=mesto2, srov2_km=int(round(km_m2, -2)))

def sestava_web(t):
    # na web jen odchylky od běžné sestavy
    t = re.sub(r'^s Ivkou(\s*\(asi\))?[;,]?\s*', '', t)
    t = t.replace('bez Ivky', 'nehrála Ivka').replace('místo Pavla bubnoval Vorel', 'nehrál Pavel, bubnoval Vorel')
    t = t.replace('bez Zdeňka J. (sax)', 'nehrál Zdeněk J. (saxofon)').replace('(basa Vóďa)', '').replace('bez saxofonu (Zdeněk J.)', 'nehrál Zdeněk J. (saxofon)')
    t = t.strip()
    return t[:1].upper() + t[1:] if t else ''
def fmt(d):
    return d.strftime('%-d. %-m. %Y') if isinstance(d, dt.date) else str(d)
from PIL import Image
def dhash(cesta):
    g = Image.open(TU + '/' + cesta).convert('L').resize((9, 8)); px = list(g.getdata())
    return sum(1 << i for i in range(64) if px[(i // 8) * 9 + i % 8] < px[(i // 8) * 9 + i % 8 + 1])
def bez_hacku(t): return re.sub(r'[^a-z0-9]', '', __import__('unicodedata').normalize('NFKD', t).encode('ascii', 'ignore').decode().lower())
def klice(r):  # podle čeho se pozná, že soubor patří k tomuto koncertu: město a začátek názvu akce
    k = [cisti_mesto(r['mesto'])] + [re.split(r'[–(,]', r['akce'])[0]]
    return [bez_hacku(x) for x in k if len(bez_hacku(x)) >= 3]
def materialy(r):
    d = str(r['d']); m = mat.get(d)
    if not m: return {}
    souse = [k for k in kon if str(k['d']) == d]
    def patri(src):  # víc koncertů v jeden den: soubor patří tomu, jehož místo nebo akce je v názvu
        if len(souse) < 2: return True
        s = bez_hacku(src); kdo = [k for k in souse if any(x in s for x in klice(k))]
        return not kdo or r in kdo
    plak, otisky = [], []
    for p in sorted((p for p in m['plakaty'] if patri(p['src'])), key=lambda p: -os.path.getsize(TU + '/' + p['vel'])):
        h = dhash(p['mal'])
        if any(bin(h ^ o).count('1') <= 6 for o in otisky): continue  # stejný plakát v jiné verzi: nechá se největší
        otisky.append(h); plak.append(dict(mal=p['mal'], vel=p['vel']))
    out = dict(plakaty=plak, playlist=[dict(vel=p['vel']) for p in m['playlist'] if patri(p['src'])],
               pisne=m['pisne'] if patri('') or len(souse) < 2 else m['pisne'],
               fotky=m['fotky'] if patri(m['zdroj'].get('fotky', '')) else None,
               videa=m['videa'] if patri(m['zdroj'].get('videa', '')) else None,
               pozvanky=[p['url'] for p in m['pozvanky'] if patri(p['src'])])
    return {k: v for k, v in out.items() if v}
seznam = [dict(d=fmt(r['d']), r=rok(r), s=str(r['d']), mesto=r['mesto'], akce=r['akce'], kapely=', '.join(r['kapely']), sestava=sestava_web(r['sestava']), odkazy=r['odkazy'], m=materialy(r))
          for r in sorted(kon, key=lambda r: str(r['d']), reverse=True)]
budouci = [dict(d=fmt(r['d']), mesto=r['mesto'], akce=r['akce']) for r in sorted(plan, key=lambda r: str(r['d']))
           if isinstance(r['d'], dt.date) and r['d'] >= dt.date.today()]

SESTAVA = [  # kdo kdy přišel a odešel; první koncerty podle evidence; bývalí členové na konci
    dict(j='Koks', n='kytara', od='od začátku'),
    dict(j='Zenk', n='zpěv', od='od začátku'),
    dict(j='Pavel', n='bicí', od='od začátku'),
    dict(j='Olda', n='basa', od='od 3. 3. 2018', pozn='první koncert v Pelhřimově'),
    dict(j='Zdeněk J.', n='saxofon', od='od prosince 2022', pozn='poprvé se o jeho přidání mluvilo na koncertě ve Zborné 10. 12. 2022'),
    dict(j='Ivka', n='pozoun', od='od 7. 6. 2025', pozn='první koncert na Pikniku na Keťásku v Jihlavě'),
    dict(j='Vóďa', n='basa', od='6. 8. 2016 – 5. 5. 2018', pozn='první basák, na rozlučce v Bezvědomí hrál půl koncertu; ještě jednou na svých narozeninách v Jamném 17. 5. 2025'),
]
prvni = min((r for r in kon if isinstance(r['d'], dt.date)), key=lambda r: r['d'])
DATA = dict(sestava=SESTAVA, prvni=dict(akce=re.split(r' – ', prvni['akce'])[0], mesto=cisti_mesto(prvni['mesto'])), n=len(kon), mest=len(mista), kapel=len(kap), od=r0, roky=po_letech, dny=po_dnech, mapa=mapa,
            kapely=kapely[:15], pisne=pisne, playlistu=len(sl_koncerty), cesty=cesty, seznam=seznam,
            budouci=budouci, aktualizace=dt.date.today().strftime('%-d. %-m. %Y'),
            nejvic=dict(n=max(x['n'] for x in po_letech), r=[x['r'] for x in po_letech if x['n'] == max(y['n'] for y in po_letech)]))

obrys = json.load(open(TU + '/data/cz_obrys.json'))
logo = open(LOGO, encoding='utf-8').read()
html = open(TU + '/template.html', encoding='utf-8').read()
html = html.replace('__DATA__', json.dumps(DATA, ensure_ascii=False)).replace('__OBRYS__', json.dumps(obrys)).replace('__LOGO__', logo)
open(TU + '/nahled.html', 'w', encoding='utf-8').write(html)  # fragment pro náhled v Artifactu
hlava, telo = html.split('<header', 1)
open(TU + '/index.html', 'w', encoding='utf-8').write(
    '<!doctype html>\n<html lang="cs">\n<head>\n<meta charset="utf-8">\n'
    '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
    '<link rel="icon" href="data:image/webp;base64,' + base64.b64encode(open(TU + '/data/logo.webp', 'rb').read()).decode() + '">\n' + hlava + '</head>\n<body>\n<header' + telo + '\n</body>\n</html>\n')
print(f'index.html: {len(kon)} koncertů, {len(mista)} míst, {len(kap)} kapel, km {cesty.get("km")}')
bez_koncertu = sorted({m['zdroj'][k] for d, m in mat.items() for k in ('fotky', 'videa')
                       if m['zdroj'].get(k) and d not in {str(r['d']) for r in kon}
                       and not re.search(r'zkušebn|nahrávání|studio', m['zdroj'][k], re.I)})
if nezarazene or bez_koncertu:
    print('\nNa web se nedostalo (Fotky a videa):')
    for p in nezarazene: print('  název nezačíná datem RRRR-MM-DD:', p)
    for p in bez_koncertu: print('  datum není v tabulce koncertů:', p)

#!/usr/bin/env python3
"""Sestaví index.html se statistikami koncertů Pozor Mamut.

Zdroje:
  - Google tabulka „Koncerty/Pozor Mamut – evidence koncertů“ (stáhne se přes rclone jako xlsx),
  - souhrny cest z ~/Dropbox/org-mode/Poznamky/pozor-mamut-cesty.tsv (soukromé; na stránku jdou
    jen součty a rekordy, žádné časy ani místa jednotlivých cest),
  - GeoNames CZ.txt (souřadnice obcí), data/cz_obrys.json (obrys ČR, Natural Earth).
Použití:  python3 make.py [--xlsx soubor.xlsx] [--geonames CZ.txt]
"""
import argparse, base64, collections, csv, datetime as dt, io, json, math, os, re, subprocess, sys
import openpyxl

TU = os.path.dirname(os.path.abspath(__file__))
HOME = os.path.expanduser('~')
REMOTE = 'gdrive,root_folder_id=1Ns75usanWOJoy2fHk7KD8EFYv58sNub_:Koncerty/Pozor Mamut – evidence koncertů.xlsx'
CESTY = HOME + '/Dropbox/org-mode/Poznamky/pozor-mamut-cesty.tsv'
LOGO = TU + '/data/logo.webp'

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
               kapely=[k.strip() for k in re.split(r',\s*', x.get('Hráli s námi') or '') if k.strip()])
    if stav == 'odehráno': kon.append(rec)
    elif stav == 'plánováno': plan.append(rec)

def rok(r): return r['d'].year if isinstance(r['d'], dt.date) else int(str(r['d'])[:4])

# --- obce: souřadnice z GeoNames ---
obce = {}
for r in csv.reader(open(a.geonames, encoding='utf-8'), delimiter='\t', quoting=csv.QUOTE_NONE):
    if r[6] != 'P': continue
    n, la, lo, pop = r[1], float(r[4]), float(r[5]), int(r[14] or 0)
    if n not in obce or pop > obce[n][2]: obce[n] = (la, lo, pop)
obce.setdefault('Praha', (50.0880, 14.4208, 1300000))
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
         'hak mak': 'Hakmak', 'hakmak': 'Hakmak', 'exploze jádra': 'Exploze jádra'}
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
OBR = re.compile(r'\.(jpe?g|png|webp|heic|pdf)$', re.I); VID = re.compile(r'\.(mp4|mov|avi|mkv|m4v)$', re.I)
mat = collections.defaultdict(lambda: dict(plakaty=[], playlist=[], fotky=None, videa=None, pisne=[]))
os.makedirs(TU + '/obrazky', exist_ok=True)
def nahled(cesta, cil, sirka):
    """Zmenšenina souboru z Disku do repa (veřejná); stáhne se jen jednou."""
    if os.path.exists(TU + '/' + cil): return cil
    data = subprocess.run(['rclone', 'cat', ROOT + cesta], check=True, capture_output=True).stdout
    vstup = '-[0]' if cesta.lower().endswith(('.pdf', '.heic')) else '-'
    subprocess.run(['convert', '-density', '110', vstup, '-auto-orient', '-resize', f'{sirka}x{sirka}>', '-quality', '82', TU + '/' + cil],
                   input=data, check=True)
    return cil
def bezpecne(s): return re.sub(r'[^a-z0-9]+', '-', s.lower().encode('ascii', 'ignore').decode()).strip('-')[:60]

for f in lsjson('Koncerty'):
    if f['IsDir'] or 'Pracovní' in f['Path'] or 'Videopozvánky' in f['Path'] or not OBR.search(f['Name']): continue
    m = DATUM.match(f['Name']) or DATUM.search(f['Path'].split('/')[1] if '/' in f['Path'] else '')
    if not m or not re.search(r'plakát|program|line-up|plakat', f['Path'], re.I): continue
    d = m.group(1); n = len(mat[d]['plakaty']) + 1; jm = f'{d}-plakat-{n}-{bezpecne(f["Name"])[:20]}'
    mat[d]['plakaty'].append(dict(mal=nahled('Koncerty/' + f['Path'], f'obrazky/{jm}-m.webp', 360),
                                  vel=nahled('Koncerty/' + f['Path'], f'obrazky/{jm}.webp', 1400)))
for f in lsjson('Písničky/Playlisty'):
    if f['IsDir']: continue
    for d in DATUM.findall(f['Name']):
        if f['Name'].lower().endswith('.txt'):
            text = subprocess.run(['rclone', 'cat', ROOT + 'Písničky/Playlisty/' + f['Path']], check=True, capture_output=True, text=True).stdout
            mat[d]['pisne'] = [re.sub(r'^\s*\d+\s*[.)/:-]?\s*', '', l).strip() for l in text.splitlines()
                               if l.strip() and not l.lstrip().startswith('#')]
        elif re.search(r'\.(jpe?g|png|webp)$', f['Name'], re.I):
            jm = f'{d}-playlist-{bezpecne(f["Name"])[:20]}'
            mat[d]['playlist'].append(dict(vel=nahled('Písničky/Playlisty/' + f['Path'], f'obrazky/{jm}.webp', 1400)))
DRIVE_F = 'https://drive.google.com/file/d/{}/view'; DRIVE_D = 'https://drive.google.com/drive/folders/{}'
fv = lsjson('Fotky a videa')
for f in fv:
    casti = f['Path'].split('/')
    m = DATUM.match(casti[-1])
    if not m: continue
    d = m.group(1)
    if f['IsDir'] and len(casti) == 2:  # složka akce: RRRR/RRRR-MM-DD Místo, akce/
        uvnitr = [g['Name'] for g in fv if g['Path'].startswith(f['Path'] + '/')]
        if any(OBR.search(x) for x in uvnitr): mat[d]['fotky'] = DRIVE_D.format(f['ID'])
        if any(VID.search(x) for x in uvnitr): mat[d]['videa'] = DRIVE_D.format(f['ID'])
    elif not f['IsDir'] and len(casti) == 2:  # jednotlivý soubor přímo v roce
        klic = 'videa' if VID.search(f['Name']) else 'fotky' if OBR.search(f['Name']) else None
        if klic and not mat[d][klic]: mat[d][klic] = DRIVE_F.format(f['ID'])

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

def fmt(d):
    return d.strftime('%-d. %-m. %Y') if isinstance(d, dt.date) else str(d)
def materialy(d):
    m = mat.get(str(d))
    if not m: return {}
    return {k: v for k, v in dict(plakaty=m['plakaty'], playlist=m['playlist'], pisne=m['pisne'], fotky=m['fotky'], videa=m['videa']).items() if v}
seznam = [dict(d=fmt(r['d']), r=rok(r), s=str(r['d']), mesto=r['mesto'], akce=r['akce'], kapely=', '.join(r['kapely']), m=materialy(r['d']))
          for r in sorted(kon, key=lambda r: str(r['d']), reverse=True)]
budouci = [dict(d=fmt(r['d']), mesto=r['mesto'], akce=r['akce']) for r in sorted(plan, key=lambda r: str(r['d']))
           if isinstance(r['d'], dt.date) and r['d'] >= dt.date.today()]

DATA = dict(n=len(kon), mest=len(mista), kapel=len(kap), od=r0, roky=po_letech, dny=po_dnech, mapa=mapa,
            kapely=kapely[:15], pisne=pisne, playlistu=len(sl_koncerty), cesty=cesty, seznam=seznam,
            budouci=budouci, aktualizace=dt.date.today().strftime('%-d. %-m. %Y'),
            nejvic=dict(n=max(x['n'] for x in po_letech), r=[x['r'] for x in po_letech if x['n'] == max(y['n'] for y in po_letech)]))

obrys = json.load(open(TU + '/data/cz_obrys.json'))
logo = 'data:image/webp;base64,' + base64.b64encode(open(LOGO, 'rb').read()).decode()
html = open(TU + '/template.html', encoding='utf-8').read()
html = html.replace('__DATA__', json.dumps(DATA, ensure_ascii=False)).replace('__OBRYS__', json.dumps(obrys)).replace('__LOGO__', logo)
open(TU + '/nahled.html', 'w', encoding='utf-8').write(html)  # fragment pro náhled v Artifactu
hlava, telo = html.split('<header', 1)
open(TU + '/index.html', 'w', encoding='utf-8').write(
    '<!doctype html>\n<html lang="cs">\n<head>\n<meta charset="utf-8">\n'
    '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
    '<link rel="icon" href="' + logo + '">\n' + hlava + '</head>\n<body>\n<header' + telo + '\n</body>\n</html>\n')
print(f'index.html: {len(kon)} koncertů, {len(mista)} míst, {len(kap)} kapel, km {cesty.get("km")}')

"""Bring this repo's CDC data up to date, keeping what it replaces.

Source: CDC's by-country data feed, the JSON the Yellow Book page
(Yellow Fever Vaccine & Malaria Prevention Information, by Country) renders
from. ExtractMalariaCDC9.ipynb fetched it on 2025-04-26 (CDC's 23 Apr 2025
edition). This script reads the feed from a saved file, or fetches it, and:

  1. Rewrites yf_malaria_by_country.csv (same columns as the notebook).
  2. Updates the CDC columns in place in the Tableau workbooks'
     source tables (yf_malaria_for_tableau_countrymatch*.xlsx), keeping every
     derived column (categories, country matching, codes), cell formatting
     included.
  3. Appends every replaced value whose text or links changed (markup-only
     differences are applied but not logged) to history/changes.csv (country, column,
     CDC's date for the new text, the date applied, previous and new HTML) and
     keeps a copy of each CSV it replaces under history/.

Never applied, only reported: Malaria_Has_Transmission. CDC's feed now flags
Syria as having transmission from a hidden "last case reported" field while
its page shows N/A. Categories (Entry Category, Resistance Type, Map
Category) are hand-derived; CATEGORY sets the ones a text change moves.

  python3 update_from_cdc.py [saved-feed-file]     (no file: fetch live)

The Tableau packaged workbook (.twbx) holds its own copy of the data: open it
in Tableau and refresh the extract after running this.
"""
import csv, html, json, os, re, shutil, sys
from datetime import date, datetime
from urllib.request import urlopen

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ENDPOINT = 'https://wwwnc.cdc.gov/travel/Services/xmlservices.asmx/YellowFeverInformationJson'
CSV = os.path.join(HERE, 'yf_malaria_by_country.csv')
XLSX = ['yf_malaria_for_tableau_countrymatch.xlsx', 'yf_malaria_for_tableau_countrymatch_vkp.xlsx',
        'yf_malaria_for_tableau_countrymatch_vkp_season.xlsx']
HISTORY = os.path.join(HERE, 'history')
TODAY = date.today().isoformat()
FIRST_EDITION = '2025-04-23'
COLUMNS = ['YF_Recommendations', 'YF_Entry_Requirements', 'Malaria_Has_Transmission', 'Malaria_Areas_of_Risk',
           'Malaria_Chloroquine_Resistance', 'Malaria_Species', 'Malaria_Recommended_Chemoprophylaxis',
           'Other_Vaccines_To_Consider', 'Map_Links', 'Map_HTML']
NEVER_APPLY = {'Malaria_Has_Transmission'}
CATEGORY = {'Colombia': {'Entry Category': 'Vaccine is required for all arriving travelers'}}


def load_feed(path):
    text = open(path, encoding='utf-8').read() if path else urlopen(ENDPOINT, timeout=60).read().decode('utf-8')
    m = re.search(r'<string[^>]*>(.*?)</string>', text, re.S)
    return json.loads(html.unescape(m.group(1)) if '&quot;' in m.group(1) else m.group(1))


def records(feed):
    out = {}
    for c in feed:
        yf, m = c.get('YellowFever', {}), c.get('Malaria', {})
        out[c['LongName']] = {
            'Country': c['LongName'],
            'YF_Recommendations': yf.get('Recommendations'), 'YF_Entry_Requirements': yf.get('Requirements'),
            'Malaria_Has_Transmission': m.get('HasTransmission'), 'Malaria_Areas_of_Risk': m.get('AreaOfRisk'),
            'Malaria_Chloroquine_Resistance': m.get('ChloroquineResistance'), 'Malaria_Species': m.get('Species'),
            'Malaria_Recommended_Chemoprophylaxis': m.get('RecommendedProphylaxis'),
            'Other_Vaccines_To_Consider': c.get('OtherVaccinesToConsider'),
            'Map_Links': c.get('MapLinks'), 'Map_HTML': c.get('MapHtml')}
    return out


def text(v):
    """What a reader sees: tags, footnote marks, CDC's "Updated" stamp,
    whitespace and full stops removed. Only real tags go ("<2,300 m" is
    text)."""
    v = '' if v is None else str(v)
    v = re.sub(r'<sup>.*?</sup>', '', v, flags=re.S)
    v = re.sub(r'<div class="updatedate">.*?</div>', '', v, flags=re.S)
    v = html.unescape(re.sub(r'</?[A-Za-z][^>]*>', ' ', v))
    v = re.sub(r'Updated [A-Z][a-z]+ \d{1,2}, \d{4}', '', v)
    return re.sub(r'[\s*•.]+', '', v)


def links(v):
    return sorted(set(re.findall(r'(?:href|src)="([^"]+)"', '' if v is None else str(v))))


def same(col, a, b):
    """Equal for a reader: the same text, and for the map columns the same
    link and image addresses. Markup-only differences don't count."""
    if col in ('Map_Links', 'Map_HTML'):
        return links(a) == links(b) and text(a) == text(b)
    if col == 'Malaria_Has_Transmission':
        return str(a) == str(b)
    return text(a) == text(b)


def cdc_date(rec, col):
    src = rec['YF_Entry_Requirements'] if col.startswith('YF') else rec['Malaria_Recommended_Chemoprophylaxis']
    m = re.search(r'Updated ([A-Z][a-z]+ \d{1,2}, \d{4})', src or '')
    return datetime.strptime(m.group(1), '%B %d, %Y').date().isoformat() if m else ''


def main():
    new = records(load_feed(sys.argv[1] if len(sys.argv) > 1 else None))
    os.makedirs(HISTORY, exist_ok=True)
    changes, flags, markup = [], set(), set()

    # 1. CSV
    old_rows = list(csv.DictReader(open(CSV, encoding='utf-8', newline='')))
    old = {r['Country']: r for r in old_rows}
    rows = []
    for name, rec in new.items():
        prev = old.get(name, {})
        row = dict(rec)
        for col in COLUMNS:
            if same(col, prev.get(col), rec[col]):
                if prev.get(col) != ('' if rec[col] is None else str(rec[col])):
                    markup.add((name, col))   # CDC's current markup, not logged
                continue
            if col in NEVER_APPLY:
                flags.add(name); row[col] = prev.get(col); continue
            changes.append({'country': name, 'column': col, 'cdc_date': cdc_date(rec, col), 'applied': TODAY,
                            'previous': prev.get(col, ''), 'new': rec[col] or ''})
        rows.append(row)
    if changes or markup:
        backup = os.path.join(HISTORY, f'yf_malaria_by_country_until_{TODAY}.csv')
        if not os.path.exists(backup):   # a second run the same day keeps the first copy
            shutil.copy(CSV, backup)
        with open(CSV, 'w', encoding='utf-8', newline='') as fh:
            w = csv.DictWriter(fh, fieldnames=['Country'] + COLUMNS)
            w.writeheader(); w.writerows(rows)

    # 2. Spreadsheets: CDC columns in place, categories from CATEGORY
    applied = {(c['country'], c['column']): c['new'] for c in changes}
    applied.update({k: new[k[0]][k[1]] for k in markup})
    for f in XLSX:
        path = os.path.join(HERE, f)
        if not os.path.exists(path):
            continue
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        head = {c.value: c.column for c in ws[1]}
        n = 0
        for r in range(2, ws.max_row + 1):
            name = ws.cell(r, head['Country']).value
            for col in COLUMNS:
                if (name, col) in applied and col in head:
                    ws.cell(r, head[col]).value = applied[(name, col)]; n += 1
            for col, value in CATEGORY.get(name, {}).items():
                cell = ws.cell(r, head[col])
                if col in head and cell.value != value:
                    if f == XLSX[0]:
                        changes.append({'country': name, 'column': col, 'cdc_date': '', 'applied': TODAY,
                                        'previous': cell.value, 'new': value})
                    cell.value = value; n += 1
        wb.save(path)
        print(f'{f}: {n} cell(s) updated')

    # 3. History
    log = os.path.join(HISTORY, 'changes.csv')
    exists = os.path.exists(log)
    with open(log, 'a', encoding='utf-8', newline='') as fh:
        w = csv.DictWriter(fh, fieldnames=['country', 'column', 'cdc_date', 'applied', 'previous', 'new'])
        if not exists:
            w.writeheader()
        w.writerows(changes)
    for c in changes:
        print('changed', c['country'], c['column'], c['cdc_date'] or '-')
    print(f'{len(changes)} change(s), logged with previous values in history/changes.csv; '
          f'{len(markup)} markup-only update(s) applied without logging')
    if flags:
        print('Malaria_Has_Transmission differs, not applied (check the CDC page):', sorted(flags))


if __name__ == '__main__':
    main()

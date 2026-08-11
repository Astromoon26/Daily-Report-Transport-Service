#!/usr/bin/env python3
"""
Parse the published-to-web CSV export of the "Daily VM" Google Sheet's
"Today" tab into the same JSON shape the original openpyxl-based parser
(run inside Cowork against a downloaded .xlsx) used to produce.

Why this exists: the original parser used openpyxl against a downloaded
.xlsx plus explicit <mergeCell> ranges read from the file's XML to know
which blank cells were merge continuations. Publish-to-web CSV carries no
merge metadata at all -- merged cells just show blank in every
continuation row. This script reconstructs the same effective grid with a
manual forward-fill step (ffill_col below) instead of merge-range lookups.

Column indices below are unchanged from the original openpyxl script --
CSV export preserves true column positions (unlike the Drive markdown
export used for the Delivery Monitoring sheet, which added a ragged
leading offset). If you're diffing this against the openpyxl version, the
row[N] indices should match 1:1.

KNOWN FRAGILITY (carried over, still true here):
- "Execution OTA by Site"'s per-type-armada columns (CDDL/FUSO/WINGBOX/
  Cont-20/Cont-40) have flipped between a single direct-percent column and
  a Hit|Total pair before, in the same day, without notice. Auto-detected
  from the header row same as before -- don't hardcode either format.
- If a brand-new table appears after "OTA by Top Vendor Sea" that isn't
  "Reason Keterlambatan Kedatangan", the generic issues-table detector at
  the bottom should still pick it up (header-row auto-detect, no hardcoded
  column names) -- but double check if issues comes back suspiciously
  empty.
"""
import csv
import io
import re
import sys


def clean(v):
    if v is None:
        return ''
    return str(v).strip()


def load_grid(csv_text):
    reader = csv.reader(io.StringIO(csv_text))
    return [[clean(c) for c in row] for row in reader]


def ffill_col(rows, start, end, col):
    last = ''
    for i in range(start, end):
        if col < len(rows[i]):
            if rows[i][col]:
                last = rows[i][col]
            else:
                rows[i][col] = last


def get(row, idx):
    return row[idx] if idx < len(row) else ''


def to_num(v):
    v = clean(v)
    if v == '':
        return None
    v = v.replace(',', '')
    try:
        if '.' in v:
            return float(v)
        return int(v)
    except ValueError:
        return v


def pct_to_num(v):
    v = clean(v).replace('%', '').strip()
    if v == '':
        return None
    try:
        return round(float(v), 2)
    except ValueError:
        return None


def find_marker(rows, text, start=0):
    for i in range(start, len(rows)):
        for cell in rows[i]:
            if text in cell:
                return i
    return None


def parse_vm(csv_text):
    rows = load_grid(csv_text)
    result = {}

    m_plan = find_marker(rows, 'Plan Order Today')
    m_ota = find_marker(rows, 'Execution OTA by Site')
    m_load = find_marker(rows, 'Execution Loading Time')
    m_land = find_marker(rows, 'OTA by Top Vendor Land')
    m_sea = find_marker(rows, 'OTA by Top Vendor Sea')

    if None in (m_plan, m_ota, m_load, m_land, m_sea):
        missing = [name for name, v in [('Plan Order Today', m_plan), ('Execution OTA by Site', m_ota),
                                         ('Execution Loading Time', m_load), ('OTA by Top Vendor Land', m_land),
                                         ('OTA by Top Vendor Sea', m_sea)] if v is None]
        raise ValueError(f"Could not locate section marker(s) in VM CSV: {missing}. "
                          f"Sheet structure may have changed -- check section titles.")

    # Forward-fill the vertically-merged Site (col0) and Site Code (col1)
    # columns across each data block before extracting rows.
    ffill_col(rows, m_plan + 3, m_ota, 0)
    ffill_col(rows, m_plan + 3, m_ota, 1)
    ffill_col(rows, m_ota + 3, m_load, 0)
    ffill_col(rows, m_ota + 3, m_load, 1)
    # Loading Time has TWO site blocks side by side (left cols 0-5, right
    # cols 6-11) -- forward-fill both independently.
    ffill_col(rows, m_load + 3, m_land, 0)
    ffill_col(rows, m_load + 3, m_land, 1)
    ffill_col(rows, m_load + 3, m_land, 6)
    ffill_col(rows, m_load + 3, m_land, 7)

    # ---------------- Section A: Plan Order Today ----------------
    plan_rows = []
    i = m_plan + 3
    while i < m_ota:
        row = rows[i]
        # Guard on moda (col2), not site/siteCode (col0/1) -- those two are
        # forward-filled across the whole block (see ffill_col above), so
        # they're truthy even on the trailing blank separator row before the
        # next section title. moda is never forward-filled and is always
        # present on genuine data rows ('Land'/'Sea'/'TOTAL'), so it reliably
        # tells real rows apart from the ffill-contaminated blank row.
        if get(row, 2):
            code = get(row, 1)
            # site (col0) was blanket-forward-filled above, which leaks the
            # last real site's name into the "Total Order" row that follows
            # it (that row's own col0 is genuinely blank on the sheet -- it's
            # not actually merged with the site above it). Null it out for
            # total rows so it matches the original openpyxl parser's
            # behavior (site=None on totals; siteCode='Total Order' is what
            # the template actually renders).
            site_val = None if code == 'Total Order' else (get(row, 0) or None)
            plan_rows.append({
                'site': site_val, 'siteCode': code or None, 'moda': get(row, 2),
                'totalOrder': to_num(get(row, 3)), 'confirmQty': to_num(get(row, 4)), 'confirmPct': pct_to_num(get(row, 5)),
                'lcH': to_num(get(row, 6)), 'lcH1': to_num(get(row, 7)), 'lcH2After12': to_num(get(row, 8)), 'lcH2Before12': to_num(get(row, 9)),
                'typeCDDL': to_num(get(row, 10)), 'typeFUSO': to_num(get(row, 11)), 'typeWINGBOX': to_num(get(row, 12)),
                'typeCont20': to_num(get(row, 13)), 'typeCont40': to_num(get(row, 14)),
                'reasonActionPlan': get(row, 15) or None,
                'isTotal': code == 'Total Order',
            })
        i += 1
    result['planOrder'] = plan_rows
    result['planOrderDate'] = get(rows[m_plan], 2) or None

    # ---------------- Section B: Execution OTA by Site ----------------
    ota_header = rows[m_ota + 2]
    ota_new_format = clean(get(ota_header, 7)).lower() == 'hit'

    def type_rate_new(row, hit_idx, total_idx):
        hit_v = to_num(get(row, hit_idx))
        tot_v = to_num(get(row, total_idx))
        denom = tot_v or 0
        p = round(hit_v / denom * 100, 2) if denom and hit_v is not None else None
        return hit_v, tot_v, p

    ota_rows = []
    i = m_ota + 3
    while i < m_load:
        row = rows[i]
        # Same reasoning as the Plan Order guard above -- gate on moda, not
        # on the forward-filled site columns.
        if get(row, 2):
            sudah_datang, ota_mis = to_num(get(row, 5)), to_num(get(row, 6))
            denom = (sudah_datang or 0) + (ota_mis or 0)
            ota_rate = round(sudah_datang / denom * 100, 2) if denom else None
            if ota_new_format:
                cddl_hit, cddl_tot, type_cddl = type_rate_new(row, 7, 8)
                fuso_hit, fuso_tot, type_fuso = type_rate_new(row, 9, 10)
                wb_hit, wb_tot, type_wingbox = type_rate_new(row, 11, 12)
                c20_hit, c20_tot, type_cont20 = type_rate_new(row, 13, 14)
                c40_hit, c40_tot, type_cont40 = type_rate_new(row, 15, 16)
            else:
                cddl_hit = cddl_tot = fuso_hit = fuso_tot = wb_hit = wb_tot = c20_hit = c20_tot = c40_hit = c40_tot = None
                type_cddl = pct_to_num(get(row, 7)); type_fuso = pct_to_num(get(row, 8)); type_wingbox = pct_to_num(get(row, 9))
                type_cont20 = pct_to_num(get(row, 10)); type_cont40 = pct_to_num(get(row, 11))
            code = get(row, 1)
            # Same site-leak fix as Plan Order above.
            site_val = None if code == 'Total Order' else (get(row, 0) or None)
            ota_rows.append({
                'site': site_val, 'siteCode': code or None, 'moda': get(row, 2),
                'qtyConfirm': to_num(get(row, 3)), 'belumDatang': to_num(get(row, 4)), 'sudahDatang': sudah_datang,
                'otaMis': ota_mis, 'otaRate': ota_rate,
                'typeCDDL': type_cddl, 'typeFUSO': type_fuso, 'typeWINGBOX': type_wingbox,
                'typeCont20': type_cont20, 'typeCont40': type_cont40,
                'typeCDDLHit': cddl_hit, 'typeCDDLTotal': cddl_tot,
                'typeFUSOHit': fuso_hit, 'typeFUSOTotal': fuso_tot,
                'typeWINGBOXHit': wb_hit, 'typeWINGBOXTotal': wb_tot,
                'typeCont20Hit': c20_hit, 'typeCont20Total': c20_tot,
                'typeCont40Hit': c40_hit, 'typeCont40Total': c40_tot,
                'isTotal': code == 'Total Order',
            })
        i += 1
    result['executionOTA'] = ota_rows
    result['executionOtaDate'] = get(rows[m_ota], 2) or None
    result['otaFormatDetected'] = 'new_hit_total' if ota_new_format else 'old_percent'

    # ---------------- Section C: Execution Loading Time by Site ----------------
    left_entries, right_entries = [], []
    i = m_load + 3
    while i < m_land:
        row = rows[i]
        # Guard on armada (col2/col8), not site/siteCode -- same
        # ffill-contamination reasoning as Plan Order / Execution OTA above.
        if get(row, 2):
            left_entries.append({
                'site': get(row, 0) or None, 'siteCode': get(row, 1) or None, 'armada': get(row, 2),
                'ciOpen': get(row, 3) or None, 'openClose': get(row, 4) or None, 'closeCO': get(row, 5) or None,
            })
        if get(row, 8):
            right_entries.append({
                'site': get(row, 6) or None, 'siteCode': get(row, 7) or None, 'armada': get(row, 8),
                'ciOpen': get(row, 9) or None, 'openClose': get(row, 10) or None, 'closeCO': get(row, 11) or None,
            })
        i += 1
    result['loadingTime'] = {'rows': left_entries + right_entries}
    result['loadingTimeDate'] = get(rows[m_load], 2) or None

    # ---------------- Section D: OTA by Top Vendor Land ----------------
    vendor_land_header = rows[m_land + 1]
    vendors_land = []
    ci = 2
    while ci < len(vendor_land_header):
        v = get(vendor_land_header, ci)
        if v:
            vendors_land.append((v, ci))
        ci += 2
    land_rows = []
    i = m_land + 3
    while i < m_sea:
        row = rows[i]
        if get(row, 0) or get(row, 1):
            vendors = {vname: {'qty': to_num(get(row, vidx)), 'pctHit': pct_to_num(get(row, vidx + 1))}
                       for vname, vidx in vendors_land}
            land_rows.append({'site': get(row, 0) or None, 'siteCode': get(row, 1) or None,
                               'vendors': vendors, 'isTotal': get(row, 1) == 'Total'})
        i += 1
    result['otaVendorLand'] = {'vendorNames': [v for v, _ in vendors_land], 'rows': land_rows}

    # ---------------- Section E: OTA by Top Vendor Sea ----------------
    vendor_sea_header = rows[m_sea + 1]
    vendors_sea = []
    ci = 2
    while ci < len(vendor_sea_header):
        v = get(vendor_sea_header, ci)
        if v:
            vendors_sea.append((v, ci))
        ci += 2
    sea_end = m_sea + 3
    while sea_end < len(rows) and any(clean(c) for c in rows[sea_end]):
        sea_end += 1
    sea_rows = []
    i = m_sea + 3
    while i < sea_end:
        row = rows[i]
        if get(row, 0) or get(row, 1):
            vendors = {vname: {'qty': to_num(get(row, vidx)), 'pctHit': pct_to_num(get(row, vidx + 1))}
                       for vname, vidx in vendors_sea}
            sea_rows.append({'site': get(row, 0) or None, 'siteCode': get(row, 1) or None,
                              'vendors': vendors, 'isTotal': get(row, 1) == 'Total'})
        i += 1
    result['otaVendorSea'] = {'vendorNames': [v for v, _ in vendors_sea], 'rows': sea_rows}

    # ---------------- Section F: Issues (generic header auto-detect) ----------------
    issues = []
    j = sea_end
    while j < len(rows) and not any(clean(c) for c in rows[j]):
        j += 1
    if j < len(rows):
        filled_count = sum(1 for c in rows[j] if clean(c))
        if filled_count <= 1:
            j += 1
    if j < len(rows) and any(clean(c) for c in rows[j]):
        header_labels = [clean(h) for h in rows[j]]
        k = j + 1
        while k < len(rows):
            row = rows[k]
            if not any(clean(c) for c in row):
                break
            record = {}
            for ci2, val in enumerate(row):
                if ci2 < len(header_labels) and header_labels[ci2] and clean(val):
                    record[header_labels[ci2]] = clean(val)
            if record:
                issues.append(record)
            k += 1
    result['issues'] = issues

    return result


if __name__ == '__main__':
    import json
    path = sys.argv[1] if len(sys.argv) > 1 else '/dev/stdin'
    with open(path, encoding='utf-8') as f:
        text = f.read()
    result = parse_vm(text)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    totals_plan = sum(1 for r in result['planOrder'] if r['isTotal'])
    totals_ota = sum(1 for r in result['executionOTA'] if r['isTotal'])
    print(f"[parse_vm] planOrder={len(result['planOrder'])} (totals={totals_plan}) "
          f"executionOTA={len(result['executionOTA'])} (totals={totals_ota}) "
          f"otaFormat={result['otaFormatDetected']} "
          f"loadingTime={len(result['loadingTime']['rows'])} "
          f"vendorLand={len(result['otaVendorLand']['rows'])} vendorSea={len(result['otaVendorSea']['rows'])} "
          f"issues={len(result['issues'])} "
          f"planOrderDate={result['planOrderDate']!r} executionOtaDate={result['executionOtaDate']!r} "
          f"loadingTimeDate={result['loadingTimeDate']!r}", file=sys.stderr)

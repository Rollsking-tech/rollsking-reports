import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from difflib import SequenceMatcher
from datetime import datetime
import io

# ── PAGE CONFIG ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="RollsKing Reports",
    page_icon="🍱",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# ── STYLING ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:wght@300;400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

.stApp {
    background: #0f0f0f;
    color: #f0f0f0;
}

.main-title {
    font-family: 'Syne', sans-serif;
    font-size: 2.6rem;
    font-weight: 800;
    color: #ffffff;
    letter-spacing: -1px;
    line-height: 1.1;
    margin-bottom: 0.2rem;
}

.main-subtitle {
    font-family: 'DM Sans', sans-serif;
    font-size: 1rem;
    color: #888;
    font-weight: 300;
    margin-bottom: 2.5rem;
}

.section-label {
    font-family: 'Syne', sans-serif;
    font-size: 0.75rem;
    font-weight: 700;
    letter-spacing: 2px;
    text-transform: uppercase;
    color: #e8a020;
    margin-bottom: 0.5rem;
}

.card {
    background: #1a1a1a;
    border: 1px solid #2a2a2a;
    border-radius: 12px;
    padding: 1.5rem;
    margin-bottom: 1.2rem;
}

.status-ok {
    color: #4ade80;
    font-size: 0.85rem;
    font-weight: 500;
}

.status-warn {
    color: #fbbf24;
    font-size: 0.85rem;
    font-weight: 500;
}

.status-err {
    color: #f87171;
    font-size: 0.85rem;
    font-weight: 500;
}

.divider {
    border: none;
    border-top: 1px solid #2a2a2a;
    margin: 1.5rem 0;
}

/* Streamlit overrides */
.stFileUploader > div {
    background: #1a1a1a !important;
    border: 1.5px dashed #333 !important;
    border-radius: 10px !important;
}
.stFileUploader > div:hover {
    border-color: #e8a020 !important;
}
.stNumberInput > div > div > input,
.stTextInput > div > div > input,
.stSelectbox > div > div {
    background: #1a1a1a !important;
    border: 1px solid #333 !important;
    color: #f0f0f0 !important;
    border-radius: 8px !important;
}
.stButton > button {
    background: #e8a020 !important;
    color: #0f0f0f !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important;
    font-size: 1rem !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 0.6rem 2rem !important;
    letter-spacing: 0.5px !important;
    width: 100% !important;
    transition: all 0.2s !important;
}
.stButton > button:hover {
    background: #f5b535 !important;
    transform: translateY(-1px) !important;
}
.stPasswordInput > div > div > input {
    background: #1a1a1a !important;
    border: 1px solid #333 !important;
    color: #f0f0f0 !important;
    border-radius: 8px !important;
}
.stSuccess {
    background: #0f2a1a !important;
    border: 1px solid #1a5c33 !important;
    border-radius: 8px !important;
}
.stError {
    background: #2a0f0f !important;
    border: 1px solid #5c1a1a !important;
    border-radius: 8px !important;
}
.stWarning {
    background: #2a200f !important;
    border: 1px solid #5c3d0a !important;
    border-radius: 8px !important;
}
div[data-testid="stExpander"] {
    background: #1a1a1a !important;
    border: 1px solid #2a2a2a !important;
    border-radius: 10px !important;
}
</style>
""", unsafe_allow_html=True)

# ── PASSWORD ──────────────────────────────────────────────────────────────────
PASSWORD = "rollsking2025"

# ── CALCULATOR HELPERS ────────────────────────────────────────────────────────
def safe_id(v):
    try:
        s = str(v).strip()
        return None if s in ('#N/A', '', 'None', 'nan') else int(float(s))
    except: return None

def safe_f(v, d=0.0):
    try:
        f = float(v if v is not None else d)
        return d if f != f else f
    except: return d

def parse_pct(v):
    if v is None: return None
    try:
        f = float(str(v).strip().replace('%', ''))
        return f if f > 1 else f * 100
    except: return None

def parse_min(v):
    if v is None: return None
    try: return float(str(v).strip().replace(' min', '').replace('min', ''))
    except: return None

def fuzzy(name, candidates, threshold=0.45):
    best, score = None, 0
    for c in candidates:
        s = SequenceMatcher(None, name.lower().strip(), c.lower().strip()).ratio()
        if s > score: score, best = s, c
    return best if score >= threshold else None

def score_c(pct):
    if pct <= 1: return 4
    if pct <= 2: return 3
    if pct <= 3: return 1
    return 0

def score_tier(avg):
    for name, lo, hi in [("Bronze", 0, 3), ("Silver", 3, 6), ("Gold", 6, 8), ("Platinum", 8, 100)]:
        if lo <= avg < hi: return name
    return "Platinum" if avg >= 8 else "Bronze"

# ── DATA LOADER ───────────────────────────────────────────────────────────────
def load_all_data(raw_bytes, biz_bytes, hygiene_scores):
    wb = openpyxl.load_workbook(io.BytesIO(raw_bytes))
    wb2 = openpyxl.load_workbook(io.BytesIO(biz_bytes)) if biz_bytes else None

    # Mapping
    mapping = {}
    for row in wb['Manager to Res ID '].iter_rows(min_row=2, values_only=True):
        if not row[0]: continue
        tl = str(row[0]).strip()
        if tl not in mapping: mapping[tl] = []
        mapping[tl].append({
            'outlet': str(row[1]).strip(), 'zmt_rk': safe_id(row[2]),
            'swg_rk': safe_id(row[3]), 'zmt_rf': safe_id(row[4]),
            'swg_rf': safe_id(row[5]), 'pos': safe_id(row[6])
        })

    # Zomato
    zmt = {}
    for row in wb['Zomato Report'].iter_rows(min_row=2, values_only=True):
        if row[0]:
            try: zmt[int(row[0])] = {'orders': safe_f(row[4]), 'complaints': safe_f(row[5])}
            except: pass

    # Swiggy
    swg = {}
    for row in wb['Swiggy Report'].iter_rows(min_row=2, values_only=True):
        if row[0]:
            try: swg[int(row[0])] = {'accepted': safe_f(row[12]), 'igcc': safe_f(row[25]),
                                      'kpt': safe_f(row[9], 999), 'rating': safe_f(row[5])}
            except: pass

    # Biz metrics
    biz = {}
    if wb2:
        for row in wb2['Business Metrics Report'].iter_rows(min_row=2, values_only=True):
            if not row[0]: continue
            rid = safe_id(row[0])
            if rid:
                if rid not in biz: biz[rid] = {}
                biz[rid][row[5]] = row[10]

    # Ratings
    ratings = {}
    for row in wb['Aggregaor Rating Link'].iter_rows(min_row=2, values_only=True):
        if not row[1]: continue
        pos = safe_id(row[1])
        if pos:
            rs = [safe_f(r) for r in [row[6], row[7], row[8], row[9]] if r and safe_f(r) > 0]
            ratings[pos] = round(sum(rs) / len(rs), 2) if rs else 0.0

    # Stock
    stock = {}
    for row in wb['OpeningClosing stock'].iter_rows(min_row=2, values_only=True):
        if row[0]: stock[str(row[0]).strip().lower()] = {'opening': safe_f(row[1]), 'closing': safe_f(row[3])}

    # Sale
    sale = {}
    for row in wb['Sale Report'].iter_rows(min_row=6, values_only=True):
        if row[7]:
            pos = safe_id(row[7])
            if pos: sale[pos] = safe_f(row[16])

    # Hyperpure
    hyp, hyp_names = {}, {}
    for row in wb['Hyperpure Phcs Rpt'].iter_rows(min_row=15, values_only=True):
        if row[0] and row[15]:
            oid = safe_id(row[0])
            if oid:
                hyp[oid] = hyp.get(oid, 0) + safe_f(row[15])
                if oid not in hyp_names and row[1]: hyp_names[oid] = str(row[1]).strip()

    # Store purchases
    store_pur = []
    for row in wb['Store Purchases Rep'].iter_rows(min_row=2, values_only=True):
        if row[1] and row[2]:
            key = str(row[1]).upper().replace('PADDY HOSPITALITY PRIVATE LIMITED', '').strip()
            store_pur.append({'key': key, 'amount': safe_f(row[2])})

    return dict(mapping=mapping, zmt=zmt, swg=swg, biz=biz, ratings=ratings,
                stock=stock, sale=sale, hyp=hyp, hyp_names=hyp_names,
                store_pur=store_pur, hygiene=hygiene_scores)


# ── CALCULATOR ────────────────────────────────────────────────────────────────
def calculate(data):
    mapping = data['mapping']
    zmt = data['zmt']; swg = data['swg']; biz = data['biz']
    ratings = data['ratings']; stock = data['stock']
    sale = data['sale']; hyp = data['hyp']; hyp_names = data['hyp_names']
    store_pur = data['store_pur']; hygiene = data['hygiene']

    results, disclaimers, flags = [], [], []

    for tl, outlets in mapping.items():
        n = len(outlets)
        tl_c = tl_k = tl_r = tl_fc = tl_av = 0
        outlet_rows = []

        for o in outlets:
            outlet = o['outlet']; pos = o['pos']
            zids = [x for x in [o['zmt_rk'], o['zmt_rf']] if x]
            sids = [x for x in [o['swg_rk'], o['swg_rf']] if x]
            notes = []

            # COMPLAINTS
            biz_cmp = None
            for sid in sids:
                if sid in biz and biz[sid].get('% Orders with Complaints'):
                    biz_cmp = parse_pct(biz[sid]['% Orders with Complaints']); break

            if biz_cmp is not None:
                s_pts = score_c(biz_cmp); cmp_src = "Swiggy BizMetrics"
                z_ord = sum(zmt[r]['orders'] for r in zids if r in zmt)
                z_cmp_v = sum(zmt[r]['complaints'] for r in zids if r in zmt)
                z_pct = z_cmp_v / z_ord * 100 if z_ord > 0 else None
                z_pts = score_c(z_pct) if z_pct is not None else 0
                cmp_pts = z_pts + s_pts; cmp_pct_display = round(biz_cmp, 2)
            else:
                s_ord = sum(swg[r]['accepted'] for r in sids if r in swg)
                s_cmp_v = sum(swg[r]['igcc'] for r in sids if r in swg)
                s_pct = s_cmp_v / s_ord * 100 if s_ord > 0 else None
                s_pts = score_c(s_pct) if s_pct is not None else 0
                z_ord = sum(zmt[r]['orders'] for r in zids if r in zmt)
                z_cmp_v = sum(zmt[r]['complaints'] for r in zids if r in zmt)
                z_pct = z_cmp_v / z_ord * 100 if z_ord > 0 else None
                z_pts = score_c(z_pct) if z_pct is not None else 0
                cmp_pts = z_pts + s_pts
                cmp_pct_display = round(s_pct, 2) if s_pct else 0
                cmp_src = "Swiggy+Zomato raw"
            tl_c += cmp_pts
            if cmp_pct_display and cmp_pct_display > 3:
                flags.append((tl, outlet, "High Complaints", f"{cmp_pct_display:.1f}%", ">3%", cmp_pts))

            # KPT
            kpt_vals, kpt_src = [], None
            for sid in sids:
                if sid in biz and biz[sid].get('Kitchen Prep Time'):
                    v = parse_min(biz[sid]['Kitchen Prep Time'])
                    if v: kpt_vals.append(v); kpt_src = "BizMetrics"
            if not kpt_vals:
                for sid in sids:
                    if sid in swg and swg[sid]['kpt'] < 999:
                        kpt_vals.append(swg[sid]['kpt']); kpt_src = "Swiggy O2MFR"
            if kpt_vals:
                avg_kpt = round(sum(kpt_vals) / len(kpt_vals), 2)
                kpt_pts = 1 if avg_kpt < 12 else 0
            else:
                avg_kpt = None; kpt_pts = 0
                notes.append("No KPT data")
                disclaimers.append(f"{tl} | {outlet}: KPT unavailable — scored 0")
            tl_k += kpt_pts
            if avg_kpt and avg_kpt >= 12:
                flags.append((tl, outlet, "KPT Exceeded", f"{avg_kpt:.1f} min", "≥12 min", kpt_pts))

            # RATINGS
            avg_rat = ratings.get(pos, 0)
            rat_pts = 1 if avg_rat >= 4.0 else 0
            tl_r += rat_pts
            if 0 < avg_rat < 4.0:
                flags.append((tl, outlet, "Low Rating", f"{avg_rat:.2f}", "<4.0", rat_pts))

            # AVAILABILITY
            avail_pct = None
            for sid in sids:
                if sid in biz and biz[sid].get('Online Availability %'):
                    avail_pct = parse_pct(biz[sid]['Online Availability %']); break
            avail_pts = 1 if (avail_pct and avail_pct >= 98) else 0
            tl_av += avail_pts
            if avail_pct and avail_pct < 98:
                flags.append((tl, outlet, "Low Availability", f"{avail_pct:.1f}%", "<98%", avail_pts))
            if avail_pct is None:
                disclaimers.append(f"{tl} | {outlet}: Availability missing — scored 0")

            # FOOD COST
            stk_key = outlet.lower().strip()
            stk_data = stock.get(stk_key) or stock.get(fuzzy(outlet, list(stock.keys())))
            hyp_amt = 0
            for hid, hname in hyp_names.items():
                if fuzzy(outlet, [hname]): hyp_amt = hyp.get(hid, 0); break
            cands = [s['key'] for s in store_pur]
            bk = fuzzy(outlet, cands, threshold=0.38)
            store_amt = next((s['amount'] for s in store_pur if s['key'] == bk), 0) if bk else 0
            net_sale = sale.get(pos, 0)
            if stk_data and net_sale > 0:
                cogs = stk_data['opening'] + (store_amt + hyp_amt) - stk_data['closing']
                fc_pct = round(cogs / net_sale * 100, 2)
                fc_pts = 1 if fc_pct < 40 else 0
            else:
                fc_pct = None; fc_pts = 0
                notes.append("FC data incomplete")
                disclaimers.append(f"{tl} | {outlet}: Food cost data incomplete — scored 0")
            tl_fc += fc_pts
            if fc_pct and fc_pct >= 40:
                flags.append((tl, outlet, "High Food Cost", f"{fc_pct:.1f}%", "≥40%", fc_pts))

            outlet_rows.append({
                'outlet': outlet, 'pos': pos,
                'cmp_pct': cmp_pct_display, 'cmp_pts': cmp_pts, 'cmp_src': cmp_src,
                'kpt_avg': avg_kpt, 'kpt_pts': kpt_pts, 'kpt_src': kpt_src or "N/A",
                'rat_avg': avg_rat, 'rat_pts': rat_pts,
                'avail_pct': avail_pct, 'avail_pts': avail_pts,
                'fc_pct': fc_pct, 'fc_pts': fc_pts,
                'notes': "; ".join(notes) if notes else "OK"
            })

        hyg_val = hygiene.get(tl, 0)
        total_pts = tl_c + tl_k + tl_r + tl_fc + hyg_val + tl_av
        avg_score = round(total_pts / n, 1) if n > 0 else 0
        tier = score_tier(avg_score)
        results.append({
            'tl': tl, 'outlets': n, 'sales_pts': 0,
            'fc_pts': tl_fc, 'cmp_pts': tl_c, 'kpt_pts': tl_k,
            'rat_pts': tl_r, 'hyg_pts': hyg_val, 'avail_pts': tl_av,
            'total_pts': total_pts, 'avg_score': avg_score,
            'tier': tier, 'outlet_detail': outlet_rows
        })

    return results, disclaimers, flags


# ── EXCEL BUILDER ─────────────────────────────────────────────────────────────
TIER_CLR = {
    "Platinum": ("1F1F1F", "FFD700"), "Gold": ("1F1F1F", "FFA500"),
    "Silver":   ("1F1F1F", "C0C0C0"), "Bronze": ("FFFFFF", "8B4513")
}
CLR = {"hd": "1F2D3D", "hm": "2E4057", "wh": "FFFFFF", "lg": "F2F2F2",
       "mg": "D9D9D9", "gn": "C6EFCE", "rd": "FFC7CE", "yw": "FFF2CC"}

def bdr():
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)

def hrow(ws, r, cols, bg, fg="FFFFFF", sz=9):
    for c, t in enumerate(cols, 1):
        cell = ws.cell(row=r, column=c, value=t)
        cell.font = Font(bold=True, color=fg, size=sz, name="Arial")
        cell.fill = PatternFill("solid", start_color=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = bdr()

def build_excel(results, disclaimers, flags, month):
    wb = openpyxl.Workbook()

    # ── SHEET 1: TL SUMMARY ──────────────────────────────────────────────────
    ws1 = wb.active; ws1.title = "TL Performance Summary"
    ws1.merge_cells("A1:K1")
    c = ws1["A1"]
    c.value = f"RollsKing — Monthly Performance Report | {month}"
    c.font = Font(bold=True, size=14, color="FFFFFF", name="Arial")
    c.fill = PatternFill("solid", start_color=CLR["hd"])
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 32

    ws1.merge_cells("A2:K2")
    c = ws1["A2"]
    c.value = f"Generated: {datetime.now().strftime('%d %b %Y, %H:%M')}  |  Hygiene & Sales Break Point require monthly manual input"
    c.font = Font(size=9, color="FFFFFF", italic=True, name="Arial")
    c.fill = PatternFill("solid", start_color=CLR["hm"])
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[2].height = 18

    headers = ["Team Leader", "Outlets", "Sales Pts", "Food Cost Pts", "Complaint Pts",
               "KPT Pts", "Rating Pts", "Hygiene Pts", "Avail Pts", "Total Avg", "Tier"]
    hrow(ws1, 3, headers, CLR["hd"])
    ws1.row_dimensions[3].height = 32

    sorted_r = sorted(results, key=lambda x: x['avg_score'], reverse=True)
    for i, r in enumerate(sorted_r, start=4):
        row = [r['tl'], r['outlets'], r['sales_pts'], r['fc_pts'], r['cmp_pts'],
               r['kpt_pts'], r['rat_pts'], r['hyg_pts'], r['avail_pts'], r['avg_score'], r['tier']]
        bg = CLR["lg"] if i % 2 == 0 else CLR["wh"]
        for col, val in enumerate(row, 1):
            c = ws1.cell(row=i, column=col, value=val)
            c.font = Font(size=9, name="Arial", bold=(col == 1))
            c.fill = PatternFill("solid", start_color=bg)
            c.alignment = Alignment(horizontal="center" if col > 1 else "left", vertical="center")
            c.border = bdr()
        fg_t, bg_t = TIER_CLR.get(r['tier'], ("000000", "FFFFFF"))
        tc = ws1.cell(row=i, column=11)
        tc.font = Font(bold=True, color=fg_t, size=9, name="Arial")
        tc.fill = PatternFill("solid", start_color=bg_t)
        tc.alignment = Alignment(horizontal="center", vertical="center")
        ws1.row_dimensions[i].height = 18

    # Grand Total
    tr = len(sorted_r) + 4
    col_keys = {3: 'sales_pts', 4: 'fc_pts', 5: 'cmp_pts', 6: 'kpt_pts',
                7: 'rat_pts', 8: 'hyg_pts', 9: 'avail_pts'}
    for col in range(1, 12):
        c = ws1.cell(row=tr, column=col)
        c.font = Font(bold=True, size=9, name="Arial")
        c.fill = PatternFill("solid", start_color=CLR["mg"])
        c.border = bdr()
        c.alignment = Alignment(horizontal="center", vertical="center")
        if col == 1:
            c.value = "GRAND TOTAL"
            c.alignment = Alignment(horizontal="left", vertical="center")
        elif col in col_keys:
            c.value = sum(r[col_keys[col]] for r in results)
    ws1.row_dimensions[tr].height = 20

    # Conditions
    cr = tr + 2
    conds = [
        ("SCORING CONDITIONS", "", "", ""),
        ("Sale Break Point", "Incremental Sale MOM", "", "PetPooja — prev month needed"),
        ("Food Cost < 40%", "1pt / 0pt", "", "Manual + Hyperpure + Store Purchases"),
        ("Complaint", "0-1%=4pts | 1-2%=3pts | 2-3%=1pt | >3%=0pt", "", "Swiggy BizMetrics + Zomato"),
        ("KPT", "< 12 min = 1pt | ≥ 12 min = 0pt", "", "Swiggy BizMetrics / O2MFR"),
        ("Hygiene", "Manual input required each month", "", "Surprise visit scores"),
        ("Rating", "≥ 4.0 = 1pt | < 4.0 = 0pt", "", "Aggregator Rating Link"),
        ("Availability", "≥ 98% = 1pt | < 98% = 0pt", "", "Swiggy BizMetrics"),
        ("Grade", "Bronze 0–3 | Silver 3–6", "Gold 6–8", "Platinum 8–10"),
    ]
    for j, (a, b, cv, d) in enumerate(conds):
        bold = (j == 0)
        bg = CLR["hm"] if j == 0 else (CLR["lg"] if j % 2 else CLR["wh"])
        fg = "FFFFFF" if j == 0 else "000000"
        for ci, val in enumerate([a, b, cv, d], 1):
            cell = ws1.cell(row=cr + j, column=ci, value=val)
            cell.font = Font(bold=bold, size=8, color=fg, name="Arial")
            cell.fill = PatternFill("solid", start_color=bg)
            cell.alignment = Alignment(vertical="center")
            cell.border = bdr()
        ws1.row_dimensions[cr + j].height = 16

    for i, w in enumerate([28, 9, 10, 14, 14, 9, 11, 12, 10, 11, 12], 1):
        ws1.column_dimensions[get_column_letter(i)].width = w

    # ── SHEET 2: OUTLET DETAIL ───────────────────────────────────────────────
    ws2 = wb.create_sheet("Outlet Detail")
    ws2.merge_cells("A1:O1")
    c = ws2["A1"]; c.value = f"Outlet-Level Detail | {month}"
    c.font = Font(bold=True, size=13, color="FFFFFF", name="Arial")
    c.fill = PatternFill("solid", start_color=CLR["hd"])
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 28

    h2 = ["Team Leader", "Outlet", "Compl %", "Compl Pts", "Compl Source",
          "KPT (min)", "KPT Pts", "KPT Source", "Avg Rating", "Rating Pts",
          "Avail %", "Avail Pts", "Food Cost %", "FC Pts", "Notes"]
    hrow(ws2, 2, h2, CLR["hm"])
    ws2.row_dimensions[2].height = 28

    rn = 3
    for r in sorted(results, key=lambda x: x['avg_score'], reverse=True):
        for o in r['outlet_detail']:
            bg = CLR["lg"] if rn % 2 == 0 else CLR["wh"]
            vals = [r['tl'], o['outlet'],
                    f"{o['cmp_pct']:.2f}%" if o['cmp_pct'] is not None else "N/A",
                    o['cmp_pts'], o['cmp_src'],
                    f"{o['kpt_avg']:.2f}" if o['kpt_avg'] is not None else "N/A",
                    o['kpt_pts'], o['kpt_src'],
                    f"{o['rat_avg']:.2f}" if o['rat_avg'] else "N/A",
                    o['rat_pts'],
                    f"{o['avail_pct']:.1f}%" if o['avail_pct'] is not None else "N/A",
                    o['avail_pts'],
                    f"{o['fc_pct']:.1f}%" if o['fc_pct'] is not None else "N/A",
                    o['fc_pts'], o['notes']]
            for col, val in enumerate(vals, 1):
                c = ws2.cell(row=rn, column=col, value=val)
                c.font = Font(size=8, name="Arial"); c.border = bdr()
                c.fill = PatternFill("solid", start_color=bg)
                c.alignment = Alignment(horizontal="left", vertical="center")
                if col in (4, 7, 10, 12, 14):
                    if val == 0: c.fill = PatternFill("solid", start_color=CLR["rd"])
                    elif isinstance(val, int) and val >= 1: c.fill = PatternFill("solid", start_color=CLR["gn"])
            ws2.row_dimensions[rn].height = 15; rn += 1

    for i, w in enumerate([24, 26, 10, 10, 20, 10, 9, 20, 11, 10, 10, 10, 12, 8, 35], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    # ── SHEET 3: FLAGGED OUTLETS ─────────────────────────────────────────────
    ws3 = wb.create_sheet("Flagged Outlets")
    ws3.merge_cells("A1:G1")
    c = ws3["A1"]; c.value = f"Flagged Outlets — Threshold Breaches | {month}"
    c.font = Font(bold=True, size=13, color="FFFFFF", name="Arial")
    c.fill = PatternFill("solid", start_color="C00000")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 28
    hrow(ws3, 2, ["Team Leader", "Outlet", "Issue", "Value", "Threshold", "Score", "Action Needed"], "C00000")
    ws3.row_dimensions[2].height = 24
    actions = {"High Complaints": "Investigate complaint types",
               "KPT Exceeded": "Review kitchen workflow",
               "Low Rating": "Review customer feedback",
               "Low Availability": "Check platform uptime",
               "High Food Cost": "Review stock & wastage"}
    rn = 3
    for f in flags:
        row = list(f) + [actions.get(f[2], "Review required")]
        for col, val in enumerate(row, 1):
            c = ws3.cell(row=rn, column=col, value=val)
            c.font = Font(size=8, name="Arial"); c.border = bdr()
            c.fill = PatternFill("solid", start_color=CLR["rd"])
            c.alignment = Alignment(horizontal="left", vertical="center")
        ws3.row_dimensions[rn].height = 15; rn += 1
    for i, w in enumerate([24, 28, 18, 14, 14, 8, 30], 1):
        ws3.column_dimensions[get_column_letter(i)].width = w

    # ── SHEET 4: DATA NOTES ──────────────────────────────────────────────────
    ws4 = wb.create_sheet("Data Notes")
    ws4.merge_cells("A1:C1")
    c = ws4["A1"]; c.value = f"Data Notes & Disclaimers | {month}"
    c.font = Font(bold=True, size=12, name="Arial")
    c.fill = PatternFill("solid", start_color=CLR["yw"])
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws4.row_dimensions[1].height = 25
    hrow(ws4, 2, ["Outlet / Note", "Detail", "Impact"], CLR["hm"])
    for i, d in enumerate(disclaimers, start=3):
        parts = d.split(": ", 1) if ": " in d else [d, ""]
        for col, val in enumerate([parts[0], parts[1] if len(parts) > 1 else "", "0 pts awarded"], 1):
            c = ws4.cell(row=i, column=col, value=val)
            c.font = Font(size=8, name="Arial"); c.border = bdr()
            c.fill = PatternFill("solid", start_color=CLR["yw"] if i % 2 == 0 else CLR["wh"])
            c.alignment = Alignment(vertical="center")
        ws4.row_dimensions[i].height = 15
    ws4.column_dimensions["A"].width = 50
    ws4.column_dimensions["B"].width = 45
    ws4.column_dimensions["C"].width = 25

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ── SESSION STATE ─────────────────────────────────────────────────────────────
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'report_bytes' not in st.session_state: st.session_state.report_bytes = None
if 'report_name' not in st.session_state: st.session_state.report_name = None

# ── LOGIN ─────────────────────────────────────────────────────────────────────
if not st.session_state.logged_in:
    st.markdown('<div class="main-title">RollsKing<br>Reports</div>', unsafe_allow_html=True)
    st.markdown('<div class="main-subtitle">Monthly performance reporting — automated</div>', unsafe_allow_html=True)
    st.markdown('<hr class="divider">', unsafe_allow_html=True)
    pwd = st.text_input("Password", type="password", placeholder="Enter access password")
    if st.button("Enter"):
        if pwd == PASSWORD:
            st.session_state.logged_in = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    st.stop()

# ── MAIN APP ──────────────────────────────────────────────────────────────────
st.markdown('<div class="main-title">RollsKing<br>Reports</div>', unsafe_allow_html=True)
st.markdown('<div class="main-subtitle">Upload your files, enter manual scores, generate the report.</div>', unsafe_allow_html=True)
st.markdown('<hr class="divider">', unsafe_allow_html=True)

# ── STEP 1: FILES ─────────────────────────────────────────────────────────────
st.markdown('<div class="section-label">Step 1 — Upload Files</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    st.markdown("**Main Raw Data File**")
    st.caption("Ops_Review_Raw_Data file with all sheets")
    raw_file = st.file_uploader("Upload main file", type=["xlsx"], key="raw", label_visibility="collapsed")
    if raw_file:
        st.markdown('<span class="status-ok">✓ Uploaded</span>', unsafe_allow_html=True)

with col2:
    st.markdown("**Swiggy Business Metrics File**")
    st.caption("business_metrics_report download from Swiggy")
    biz_file = st.file_uploader("Upload business metrics", type=["xlsx"], key="biz", label_visibility="collapsed")
    if biz_file:
        st.markdown('<span class="status-ok">✓ Uploaded</span>', unsafe_allow_html=True)
    else:
        st.markdown('<span class="status-warn">⚠ Optional — KPT & Availability may be limited without it</span>', unsafe_allow_html=True)

st.markdown('<hr class="divider">', unsafe_allow_html=True)

# ── STEP 2: MONTH ─────────────────────────────────────────────────────────────
st.markdown('<div class="section-label">Step 2 — Report Month</div>', unsafe_allow_html=True)
months = ["January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]
years = [2024, 2025, 2026]
col1, col2 = st.columns(2)
with col1:
    sel_month = st.selectbox("Month", months, index=11)
with col2:
    sel_year = st.selectbox("Year", years, index=1)
month_label = f"{sel_month} {sel_year}"

st.markdown('<hr class="divider">', unsafe_allow_html=True)

# ── STEP 3: HYGIENE ───────────────────────────────────────────────────────────
st.markdown('<div class="section-label">Step 3 — Hygiene Scores (Manual)</div>', unsafe_allow_html=True)
st.caption("Enter surprise visit scores for each Team Leader. Leave as 0 if not conducted this month.")

if raw_file:
    try:
        wb_tmp = openpyxl.load_workbook(io.BytesIO(raw_file.read()))
        raw_file.seek(0)
        tl_names = []
        for row in wb_tmp['Manager to Res ID '].iter_rows(min_row=2, values_only=True):
            if row[0] and str(row[0]).strip() not in tl_names:
                tl_names.append(str(row[0]).strip())
    except:
        tl_names = []

    hygiene_scores = {}
    if tl_names:
        cols = st.columns(3)
        for i, tl in enumerate(tl_names):
            with cols[i % 3]:
                short = tl.split('(')[0].strip()
                hygiene_scores[tl] = st.number_input(short, min_value=0, max_value=20, value=0, key=f"hyg_{tl}")
    else:
        st.warning("Upload the main file first to see TL names.")
        hygiene_scores = {}
else:
    st.info("Upload the main file first to enter hygiene scores.")
    hygiene_scores = {}

st.markdown('<hr class="divider">', unsafe_allow_html=True)

# ── STEP 4: GENERATE ─────────────────────────────────────────────────────────
st.markdown('<div class="section-label">Step 4 — Generate Report</div>', unsafe_allow_html=True)

if not raw_file:
    st.warning("Please upload the main raw data file to continue.")
elif st.button("Generate Report"):
    with st.spinner("Reading data and calculating scores..."):
        try:
            raw_bytes = raw_file.read()
            biz_bytes = biz_file.read() if biz_file else None

            data = load_all_data(raw_bytes, biz_bytes, hygiene_scores)
            results, disclaimers, flags = calculate(data)

            st.success(f"Calculated scores for {sum(r['outlets'] for r in results)} outlets across {len(results)} Team Leaders.")

            with st.spinner("Building Excel report..."):
                excel_bytes = build_excel(results, disclaimers, flags, month_label)
                st.session_state.report_bytes = excel_bytes
                st.session_state.report_name = f"RollsKing_Performance_{sel_month}_{sel_year}.xlsx"

        except Exception as e:
            st.error(f"Error: {str(e)}")
            st.exception(e)

# ── DOWNLOAD ──────────────────────────────────────────────────────────────────
if st.session_state.report_bytes:
    st.markdown('<hr class="divider">', unsafe_allow_html=True)
    st.markdown('<div class="section-label">Report Ready</div>', unsafe_allow_html=True)

    # Quick preview
    try:
        wb_prev = openpyxl.load_workbook(io.BytesIO(st.session_state.report_bytes))
        ws_prev = wb_prev["TL Performance Summary"]
        with st.expander("Preview — TL Scores", expanded=True):
            rows = []
            for row in ws_prev.iter_rows(min_row=4, values_only=True):
                if row[0] and row[0] != "GRAND TOTAL" and row[9]:
                    rows.append({"Team Leader": row[0], "Avg Score": row[9], "Tier": row[10]})
                if row[0] == "GRAND TOTAL": break
            if rows:
                import pandas as pd
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    except: pass

    st.download_button(
        label="⬇ Download Excel Report",
        data=st.session_state.report_bytes,
        file_name=st.session_state.report_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ── FOOTER ────────────────────────────────────────────────────────────────────
st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.markdown(
    '<p style="color:#444;font-size:0.75rem;text-align:center;">RollsKing Internal Tool · Files are not stored · Session only</p>',
    unsafe_allow_html=True
)

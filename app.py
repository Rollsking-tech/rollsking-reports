import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from difflib import SequenceMatcher
from datetime import datetime
import io
import json

# ── PAGE CONFIG ───────────────────────────────────────────────────────────────
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

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.stApp { background: #0f0f0f; color: #f0f0f0; }

.main-title {
    font-family: 'Syne', sans-serif;
    font-size: 2.4rem; font-weight: 800;
    color: #ffffff; letter-spacing: -1px;
    line-height: 1.1; margin-bottom: 0.2rem;
}
.main-subtitle {
    font-family: 'DM Sans', sans-serif;
    font-size: 0.95rem; color: #888;
    font-weight: 300; margin-bottom: 2rem;
}
.section-label {
    font-family: 'Syne', sans-serif;
    font-size: 0.7rem; font-weight: 700;
    letter-spacing: 2.5px; text-transform: uppercase;
    color: #e8a020; margin-bottom: 0.4rem;
}
.card {
    background: #1a1a1a; border: 1px solid #2a2a2a;
    border-radius: 14px; padding: 1.5rem; margin-bottom: 1.2rem;
}
.card-gold {
    background: #1a1a1a; border: 1px solid #e8a020;
    border-radius: 14px; padding: 1.5rem; margin-bottom: 1.2rem;
}
.status-ok  { color: #4ade80; font-size: 0.82rem; font-weight: 500; }
.status-warn{ color: #fbbf24; font-size: 0.82rem; font-weight: 500; }
.status-err { color: #f87171; font-size: 0.82rem; font-weight: 500; }
.chip-ok  { display:inline-block; background:#052e16; color:#4ade80; border-radius:20px;
             padding:2px 10px; font-size:0.75rem; font-weight:600; margin:2px; }
.chip-warn{ display:inline-block; background:#2a1f00; color:#fbbf24; border-radius:20px;
             padding:2px 10px; font-size:0.75rem; font-weight:600; margin:2px; }
.chip-miss{ display:inline-block; background:#2a0f0f; color:#f87171; border-radius:20px;
             padding:2px 10px; font-size:0.75rem; font-weight:600; margin:2px; }
.step-badge {
    display:inline-block; background:#e8a020; color:#0f0f0f;
    border-radius:50%; width:26px; height:26px; text-align:center;
    line-height:26px; font-weight:800; font-size:0.85rem;
    font-family:'Syne',sans-serif; margin-right:8px;
}
.divider { border:none; border-top:1px solid #2a2a2a; margin:1.5rem 0; }

/* Streamlit overrides */
.stFileUploader > div {
    background: #1a1a1a !important;
    border: 1.5px dashed #3a3a3a !important;
    border-radius: 12px !important;
}
.stFileUploader > div:hover { border-color: #e8a020 !important; }
.stNumberInput > div > div > input,
.stTextInput > div > div > input,
.stSelectbox > div > div {
    background: #1a1a1a !important; border: 1px solid #333 !important;
    color: #f0f0f0 !important; border-radius: 8px !important;
}
.stButton > button {
    background: #e8a020 !important; color: #0f0f0f !important;
    font-family: 'Syne', sans-serif !important; font-weight: 700 !important;
    font-size: 1rem !important; border: none !important;
    border-radius: 8px !important; padding: 0.6rem 2rem !important;
    letter-spacing: 0.5px !important; width: 100% !important;
    transition: all 0.2s !important;
}
.stButton > button:hover { background: #f5b535 !important; transform: translateY(-1px) !important; }
.stPasswordInput > div > div > input {
    background: #1a1a1a !important; border: 1px solid #333 !important;
    color: #f0f0f0 !important; border-radius: 8px !important;
}
div[data-testid="stExpander"] {
    background: #1a1a1a !important; border: 1px solid #2a2a2a !important;
    border-radius: 10px !important;
}
.stTabs [data-baseweb="tab-list"] { background: #1a1a1a !important; border-radius: 10px !important; }
.stTabs [data-baseweb="tab"] { color: #888 !important; }
.stTabs [aria-selected="true"] { color: #e8a020 !important; }
[data-testid="stMetricValue"] { color: #e8a020 !important; font-family: 'Syne', sans-serif !important; }
</style>
""", unsafe_allow_html=True)

# ── PASSWORD ──────────────────────────────────────────────────────────────────
APP_PASSWORD = "rollsking2025"

# ── HELPERS ───────────────────────────────────────────────────────────────────
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
        f = float(str(v).strip().replace('%','').replace('₹',''))
        return f if f > 1 else f * 100
    except: return None

def parse_min(v):
    if v is None: return None
    try: return float(str(v).strip().replace(' min','').replace('min',''))
    except: return None

def fuzzy(name, candidates, threshold=0.45):
    best, score = None, 0
    for c in candidates:
        s = SequenceMatcher(None, str(name).lower().strip(), str(c).lower().strip()).ratio()
        if s > score: score, best = s, c
    return best if score >= threshold else None

def score_c(pct):
    if pct is None: return 0
    if pct <= 1: return 4
    if pct <= 2: return 3
    if pct <= 3: return 1
    return 0

def score_tier(avg):
    for name, lo, hi in [("Bronze", 0, 3), ("Silver", 3, 6), ("Gold", 6, 8), ("Platinum", 8, 100)]:
        if lo <= avg < hi: return name
    return "Platinum" if avg >= 8 else "Bronze"

# ── DETECT FILE TYPE ──────────────────────────────────────────────────────────
def detect_file_type(file_bytes, filename):
    """Auto-detect what kind of file was uploaded based on sheet names and content."""
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        sheets = [s.lower() for s in wb.sheetnames]

        # New monthly format: has zomato raw data + swiggy raw data + food cost
        has_zomato_raw = any('zomato' in s and ('raw' in s or 'row' in s) for s in sheets)
        has_swiggy_raw = any('swiggy' in s and 'raw' in s for s in sheets)
        has_food_cost  = any('food cost' in s for s in sheets)
        has_sale_raw   = any('sale' in s for s in sheets)
        has_mapping    = any(s in ('sheet1',) or ('mapping' in s) for s in sheets)

        if has_zomato_raw and has_food_cost and has_sale_raw:
            return 'monthly_raw', wb
        if has_mapping and not has_zomato_raw:
            return 'mapping', wb
        return 'unknown', wb
    except:
        return 'unknown', None

# ── LOAD MAPPING FROM STORED JSON ─────────────────────────────────────────────
def parse_mapping_from_wb(wb):
    """Parse mapping — supports both Nov Sheet1 format and old Manager to Res ID format."""
    mapping = {}
    sheet = None
    sheet_format = None

    for name in wb.sheetnames:
        n = name.lower().strip()
        if n == 'sheet1' or 'mapping' in n:
            sheet = wb[name]; sheet_format = 'new'; break
        if 'manager' in n and 'res' in n:
            sheet = wb[name]; sheet_format = 'old'; break

    if not sheet:
        available = ", ".join(f"\'{s}\'" for s in wb.sheetnames)
        return None, (
            f"Could not find a mapping sheet. "
            f"Sheets found in this file: {available}. "
            f"Please upload Nov_Month_Data.xlsx or the Manager to Res ID file."
        )

    for row in sheet.iter_rows(min_row=2, values_only=True):
        if not row[0]: continue

        if sheet_format == 'old':
            # Old format: Manager name, Outlet name, Zomato RK, Swiggy RK, Zomato RF, Swiggy RF, PetPooja ID
            asm     = str(row[0]).strip() if row[0] else ''
            subzone = str(row[1]).strip() if row[1] else ''
            zmt_rk  = safe_id(row[2])
            swg_rk  = safe_id(row[3])
            zmt_rf  = safe_id(row[4])
            swg_rf  = safe_id(row[5])
            pos_id  = safe_id(row[6])
        else:
            # New format (Sheet1): Subzone, POS ID, Zone, ASM, Zomato RK, Zomato RF, Swiggy RK, Swiggy RF
            subzone = str(row[0]).strip() if row[0] else ''
            pos_id  = safe_id(row[1])
            asm     = str(row[3]).strip() if row[3] else ''
            zmt_rk  = safe_id(row[4])
            zmt_rf  = safe_id(row[5])
            swg_rk  = safe_id(row[6])
            swg_rf  = safe_id(row[7])

        if not asm or not subzone: continue
        if asm not in mapping: mapping[asm] = []
        mapping[asm].append({
            'outlet': subzone, 'pos': pos_id,
            'zmt_rk': zmt_rk, 'zmt_rf': zmt_rf,
            'swg_rk': swg_rk, 'swg_rf': swg_rf,
        })
    return mapping, None

# ── LOAD MONTHLY RAW DATA ─────────────────────────────────────────────────────
def load_monthly_raw(wb):
    """Load Zomato, Swiggy, Food Cost from new monthly file format."""

    # ── Zomato ────────────────────────────────────────────────────────────────
    zmt_sheet = None
    for name in wb.sheetnames:
        if 'zomato' in name.lower(): zmt_sheet = wb[name]; break

    zmt = {}  # {res_id: {orders, complaints, kpt, rating, online_pct}}
    if zmt_sheet:
        for row in zmt_sheet.iter_rows(min_row=2, values_only=True):
            res_id = safe_id(row[0])
            if not res_id: continue
            metric = str(row[5]).strip() if row[5] else ''
            value  = row[6]
            if res_id not in zmt:
                zmt[res_id] = {'orders': 0, 'complaints': 0, 'kpt': None,
                               'rating': None, 'online_pct': None}
            if metric == 'Delivered orders':
                zmt[res_id]['orders'] = safe_f(value)
            elif metric == 'Total complaints':
                zmt[res_id]['complaints'] = safe_f(value)
            elif metric == 'KPT (in minutes)':
                zmt[res_id]['kpt'] = safe_f(value)
            elif metric == 'Average rating':
                zmt[res_id]['rating'] = safe_f(value)
            elif metric == 'Online %':
                zmt[res_id]['online_pct'] = parse_pct(value)

    # ── Swiggy ────────────────────────────────────────────────────────────────
    swg_sheet = None
    for name in wb.sheetnames:
        if 'swiggy' in name.lower(): swg_sheet = wb[name]; break

    swg = {}  # {res_id: {kpt, avail, cmp_pct, orders}}
    if swg_sheet:
        for row in swg_sheet.iter_rows(min_row=2, values_only=True):
            res_id = safe_id(row[0])
            if not res_id: continue
            metric = str(row[5]).strip() if row[5] else ''
            value  = row[6]
            if res_id not in swg:
                swg[res_id] = {'kpt': None, 'avail': None, 'cmp_pct': None, 'orders': 0}
            if metric == 'Kitchen Prep Time':
                swg[res_id]['kpt'] = parse_min(str(value).replace(' mins','').replace(' min','')) if value else None
            elif metric == 'Online Availability %':
                swg[res_id]['avail'] = parse_pct(value)
            elif metric == '% Orders with Complaints':
                swg[res_id]['cmp_pct'] = parse_pct(value)
            elif metric in ('Delivered Orders', 'Orders'):
                swg[res_id]['orders'] = safe_f(value)

    # ── Food Cost ─────────────────────────────────────────────────────────────
    fc_sheet = None
    for name in wb.sheetnames:
        if 'food cost' in name.lower(): fc_sheet = wb[name]; break

    food_cost = {}  # {pos_id: {fc_pct, net_sale}}
    if fc_sheet:
        for row in fc_sheet.iter_rows(min_row=2, values_only=True):
            pos = safe_id(row[1])
            if not pos: continue
            # Columns: Subzone, POS ID, Zone, ASM, Net Sale+PC, Opening, Closing, Local/Hyperpure, Store, Food Cost
            net_sale = safe_f(row[4])
            opening  = safe_f(row[5])
            closing  = safe_f(row[6])
            hyperpure= safe_f(row[7])
            store    = safe_f(row[8])
            fc_val   = row[9]  # pre-calculated food cost %
            if fc_val is not None:
                fc_pct = safe_f(fc_val) * 100 if safe_f(fc_val) < 2 else safe_f(fc_val)
            elif net_sale > 0:
                cogs   = opening + hyperpure + store - closing
                fc_pct = round(cogs / net_sale * 100, 2)
            else:
                fc_pct = None
            food_cost[pos] = {'fc_pct': fc_pct, 'net_sale': net_sale}

    # ── Ratings from Zomato ───────────────────────────────────────────────────
    # We'll use Zomato average rating keyed by res_id; match to outlet via mapping

    return zmt, swg, food_cost

# ── CALCULATOR (new format) ───────────────────────────────────────────────────
def calculate_new(mapping, zmt, swg, food_cost, hygiene_scores):
    results, disclaimers, flags = [], [], []

    for tl, outlets in mapping.items():
        n = len(outlets)
        tl_c = tl_k = tl_r = tl_fc = tl_av = 0
        outlet_rows = []

        for o in outlets:
            outlet = o['outlet']
            pos    = o['pos']
            zids   = [x for x in [o['zmt_rk'], o['zmt_rf']] if x]
            sids   = [x for x in [o['swg_rk'], o['swg_rf']] if x]
            notes  = []

            # ── COMPLAINTS ────────────────────────────────────────────────────
            # Prefer Swiggy % orders with complaints, fallback to Zomato raw calc
            s_cmp_pct = None
            for sid in sids:
                if sid in swg and swg[sid].get('cmp_pct') is not None:
                    s_cmp_pct = swg[sid]['cmp_pct']; break

            z_ord   = sum(zmt[r]['orders']     for r in zids if r in zmt)
            z_cmp_v = sum(zmt[r]['complaints'] for r in zids if r in zmt)
            z_pct   = round(z_cmp_v / z_ord * 100, 2) if z_ord > 0 else None
            z_pts   = score_c(z_pct)

            if s_cmp_pct is not None:
                s_pts   = score_c(s_cmp_pct)
                cmp_pts = z_pts + s_pts
                cmp_display = round(s_cmp_pct, 2)
                cmp_src = "Swiggy+Zomato"
            elif z_pct is not None:
                cmp_pts = z_pts
                cmp_display = z_pct
                cmp_src = "Zomato only"
            else:
                cmp_pts = 0; cmp_display = 0; cmp_src = "No data"
                notes.append("No complaint data")
                disclaimers.append(f"{tl} | {outlet}: Complaint data missing — scored 0")

            tl_c += cmp_pts
            if cmp_display and cmp_display > 3:
                flags.append((tl, outlet, "High Complaints", f"{cmp_display:.1f}%", ">3%", cmp_pts))

            # ── KPT ──────────────────────────────────────────────────────────
            kpt_vals = []
            for sid in sids:
                if sid in swg and swg[sid].get('kpt') is not None:
                    kpt_vals.append(swg[sid]['kpt'])
            # Fallback: Zomato KPT
            if not kpt_vals:
                for rid in zids:
                    if rid in zmt and zmt[rid].get('kpt') is not None:
                        kpt_vals.append(zmt[rid]['kpt'])

            if kpt_vals:
                avg_kpt = round(sum(kpt_vals) / len(kpt_vals), 2)
                kpt_pts = 1 if avg_kpt < 12 else 0
                kpt_src = "Swiggy" if any(swg.get(s, {}).get('kpt') for s in sids) else "Zomato"
            else:
                avg_kpt = None; kpt_pts = 0; kpt_src = "N/A"
                notes.append("No KPT data")
                disclaimers.append(f"{tl} | {outlet}: KPT unavailable — scored 0")

            tl_k += kpt_pts
            if avg_kpt and avg_kpt >= 12:
                flags.append((tl, outlet, "KPT Exceeded", f"{avg_kpt:.1f} min", "≥12 min", kpt_pts))

            # ── RATING ───────────────────────────────────────────────────────
            rat_vals = []
            for rid in zids:
                if rid in zmt and zmt[rid].get('rating'):
                    rat_vals.append(zmt[rid]['rating'])
            avg_rat = round(sum(rat_vals) / len(rat_vals), 2) if rat_vals else 0
            rat_pts = 1 if avg_rat >= 4.0 else 0
            tl_r += rat_pts
            if 0 < avg_rat < 4.0:
                flags.append((tl, outlet, "Low Rating", f"{avg_rat:.2f}", "<4.0", rat_pts))
            if not rat_vals:
                notes.append("No rating data")

            # ── AVAILABILITY ─────────────────────────────────────────────────
            avail_pct = None
            for sid in sids:
                if sid in swg and swg[sid].get('avail') is not None:
                    avail_pct = swg[sid]['avail']; break
            avail_pts = 1 if (avail_pct is not None and avail_pct >= 98) else 0
            tl_av += avail_pts
            if avail_pct is not None and avail_pct < 98:
                flags.append((tl, outlet, "Low Availability", f"{avail_pct:.1f}%", "<98%", avail_pts))
            if avail_pct is None:
                notes.append("No availability data")
                disclaimers.append(f"{tl} | {outlet}: Availability missing — scored 0")

            # ── FOOD COST ────────────────────────────────────────────────────
            fc_data = food_cost.get(pos)
            if fc_data and fc_data['fc_pct'] is not None:
                fc_pct = round(fc_data['fc_pct'], 2)
                fc_pts = 1 if fc_pct < 40 else 0
            else:
                fc_pct = None; fc_pts = 0
                notes.append("FC data missing")
                disclaimers.append(f"{tl} | {outlet}: Food cost data missing — scored 0")

            tl_fc += fc_pts
            if fc_pct is not None and fc_pct >= 40:
                flags.append((tl, outlet, "High Food Cost", f"{fc_pct:.1f}%", "≥40%", fc_pts))

            outlet_rows.append({
                'outlet': outlet, 'pos': pos,
                'cmp_pct': cmp_display, 'cmp_pts': cmp_pts, 'cmp_src': cmp_src,
                'kpt_avg': avg_kpt, 'kpt_pts': kpt_pts, 'kpt_src': kpt_src,
                'rat_avg': avg_rat, 'rat_pts': rat_pts,
                'avail_pct': avail_pct, 'avail_pts': avail_pts,
                'fc_pct': fc_pct, 'fc_pts': fc_pts,
                'notes': "; ".join(notes) if notes else "OK"
            })

        hyg_val    = hygiene_scores.get(tl, 0)
        total_pts  = tl_c + tl_k + tl_r + tl_fc + hyg_val + tl_av
        avg_score  = round(total_pts / n, 1) if n > 0 else 0
        tier       = score_tier(avg_score)

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
    c.value = f"Generated: {datetime.now().strftime('%d %b %Y, %H:%M')}  |  Hygiene requires manual input each month"
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

    cr = tr + 2
    conds = [
        ("SCORING CONDITIONS", "", "", ""),
        ("Food Cost < 40%", "1pt / 0pt", "", "Pre-calculated in Food Cost sheet"),
        ("Complaint", "0-1%=4pts | 1-2%=3pts | 2-3%=1pt | >3%=0pt", "", "Swiggy + Zomato"),
        ("KPT", "< 12 min = 1pt | ≥ 12 min = 0pt", "", "Swiggy / Zomato"),
        ("Hygiene", "Manual input required each month", "", "Surprise visit scores"),
        ("Rating", "≥ 4.0 = 1pt | < 4.0 = 0pt", "", "Zomato Average Rating"),
        ("Availability", "≥ 98% = 1pt | < 98% = 0pt", "", "Swiggy Online Availability"),
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
                    f"{o['cmp_pct']:.2f}%" if o['cmp_pct'] else "N/A",
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

    for i, w in enumerate([24, 26, 10, 10, 18, 10, 9, 15, 11, 10, 10, 10, 12, 8, 35], 1):
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
    wb.save(buf); buf.seek(0)
    return buf.read()

# ── PDF BUILDER ───────────────────────────────────────────────────────────────
def build_pdf_report(results, flags, disclaimers, month):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                     TableStyle, HRFlowable, PageBreak)
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

    C_BG=colors.HexColor('#0f0f0f'); C_GOLD=colors.HexColor('#e8a020')
    C_WHITE=colors.white; C_GREY=colors.HexColor('#888888')
    C_RED=colors.HexColor('#ef4444')
    T_PLATINUM=colors.HexColor('#FFD700'); T_GOLD=colors.HexColor('#FFA500')
    T_SILVER=colors.HexColor('#C0C0C0'); T_BRONZE=colors.HexColor('#CD7F32')
    TIER_BG={'Platinum':T_PLATINUM,'Gold':T_GOLD,'Silver':T_SILVER,'Bronze':T_BRONZE}
    TIER_FG={'Platinum':colors.black,'Gold':colors.black,'Silver':colors.black,'Bronze':colors.white}

    def sty(name,**kw):
        s=ParagraphStyle(name)
        for k,v in kw.items(): setattr(s,k,v)
        return s

    buf=io.BytesIO()
    doc=SimpleDocTemplate(buf,pagesize=A4,leftMargin=15*mm,rightMargin=15*mm,
                          topMargin=12*mm,bottomMargin=12*mm)
    story=[]; sorted_r=sorted(results,key=lambda x:x['avg_score'],reverse=True)

    hd=[[Paragraph('RollsKing',sty('h1',fontSize=32,leading=36,textColor=C_WHITE,
                   fontName='Helvetica-Bold')),
         Paragraph(f'Generated<br/>{datetime.now().strftime("%d %b %Y")}',
                   sty('hr',fontSize=8,leading=12,textColor=C_GREY,
                       fontName='Helvetica',alignment=TA_RIGHT))]]
    ht=Table(hd,colWidths=[120*mm,60*mm])
    ht.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),C_BG),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(0,0),(-1,-1),6*mm),
        ('RIGHTPADDING',(0,0),(-1,-1),6*mm),('TOPPADDING',(0,0),(-1,-1),6*mm),
        ('BOTTOMPADDING',(0,0),(-1,-1),4*mm)]))
    story.append(ht)
    story.append(HRFlowable(width='100%',thickness=2,color=C_GOLD,spaceAfter=4))

    sd=[[Paragraph(f'Monthly Performance Report — {month}',
                   sty('ms',fontSize=11,leading=14,textColor=C_GOLD,fontName='Helvetica-Bold')),
         Paragraph(f'{len(results)} Team Leaders  ·  {sum(r["outlets"] for r in results)} Outlets',
                   sty('ms2',fontSize=9,leading=12,textColor=C_GREY,
                       fontName='Helvetica',alignment=TA_RIGHT))]]
    st2=Table(sd,colWidths=[120*mm,60*mm])
    st2.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),C_BG),
        ('LEFTPADDING',(0,0),(-1,-1),6*mm),('RIGHTPADDING',(0,0),(-1,-1),6*mm),
        ('TOPPADDING',(0,0),(-1,-1),3*mm),('BOTTOMPADDING',(0,0),(-1,-1),4*mm)]))
    story.append(st2); story.append(Spacer(1,6*mm))

    pc=sum(1 for r in results if r['tier']=='Platinum')
    gc=sum(1 for r in results if r['tier']=='Gold')
    sc=sum(1 for r in results if r['tier']=='Silver')
    bc=sum(1 for r in results if r['tier']=='Bronze')
    top=sorted_r[0]
    def scell(val,lbl):
        return [Paragraph(f'<b>{val}</b>',sty('sv',fontSize=22,leading=26,
                textColor=colors.black,fontName='Helvetica-Bold',alignment=TA_CENTER)),
                Paragraph(lbl,sty('sl',fontSize=7,leading=9,textColor=colors.black,
                fontName='Helvetica',alignment=TA_CENTER))]
    stats=[[scell(f'{top["avg_score"]}',f'Top Score')[0],scell(str(pc),'Platinum')[0],
            scell(str(gc),'Gold')[0],scell(str(sc),'Silver')[0],
            scell(str(bc),'Bronze')[0],
            Paragraph(f'<b>{len(flags)}</b>',sty('sv2',fontSize=22,leading=26,
            textColor=colors.white,fontName='Helvetica-Bold',alignment=TA_CENTER))],
           [scell(f'{top["avg_score"]}',f'Top\n{top["tl"].split("(")[0].strip()}')[1],
            scell(str(pc),'Platinum')[1],scell(str(gc),'Gold')[1],
            scell(str(sc),'Silver')[1],scell(str(bc),'Bronze')[1],
            Paragraph('Flags',sty('sl2',fontSize=7,leading=9,textColor=colors.white,
            fontName='Helvetica',alignment=TA_CENTER))]]
    sbgs=[C_GOLD,T_PLATINUM,T_GOLD,T_SILVER,T_BRONZE,C_RED]
    stbl=Table(stats,colWidths=[30*mm]*6)
    sstyle=[('ALIGN',(0,0),(-1,-1),'CENTER'),('VALIGN',(0,0),(-1,-1),'MIDDLE'),
            ('TOPPADDING',(0,0),(-1,-1),3*mm),('BOTTOMPADDING',(0,0),(-1,-1),3*mm)]
    for i,bg in enumerate(sbgs): sstyle+=[('BACKGROUND',(i,0),(i,1),bg)]
    stbl.setStyle(TableStyle(sstyle))
    story.append(stbl); story.append(Spacer(1,6*mm))

    story.append(Paragraph('TEAM LEADER PERFORMANCE',
                            sty('sec',fontSize=8,leading=10,textColor=C_GOLD,
                                fontName='Helvetica-Bold',spaceAfter=4)))
    hdrs=[['#','Team Leader','Outlets','Complaints','KPT','Rating','Food Cost','Hygiene','Avail','Avg','Tier']]
    rows=hdrs+[[str(i),r['tl'].split('(')[0].strip(),str(r['outlets']),
                str(r['cmp_pts']),str(r['kpt_pts']),str(r['rat_pts']),
                str(r['fc_pts']),str(r['hyg_pts']),str(r['avail_pts']),
                str(r['avg_score']),r['tier']] for i,r in enumerate(sorted_r,1)]
    cw=[8*mm,42*mm,14*mm,20*mm,10*mm,13*mm,17*mm,14*mm,12*mm,12*mm,18*mm]
    ttbl=Table(rows,colWidths=cw,repeatRows=1)
    ts=[('BACKGROUND',(0,0),(-1,0),C_BG),('TEXTCOLOR',(0,0),(-1,0),C_GOLD),
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,0),7.5),
        ('ALIGN',(0,0),(-1,-1),'CENTER'),('ALIGN',(1,0),(1,-1),'LEFT'),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),('FONTNAME',(0,1),(-1,-1),'Helvetica'),
        ('FONTSIZE',(0,1),(-1,-1),8),('TOPPADDING',(0,0),(-1,-1),2.5*mm),
        ('BOTTOMPADDING',(0,0),(-1,-1),2.5*mm),('LEFTPADDING',(0,0),(-1,-1),2*mm),
        ('GRID',(0,0),(-1,-1),0.3,colors.HexColor('#dddddd')),
        ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#f5f5f5')])]
    for i,r in enumerate(sorted_r,1):
        ts+=[('BACKGROUND',(10,i),(10,i),TIER_BG.get(r['tier'],colors.white)),
             ('TEXTCOLOR',(10,i),(10,i),TIER_FG.get(r['tier'],colors.black)),
             ('FONTNAME',(10,i),(10,i),'Helvetica-Bold')]
        if i==1: ts.append(('BACKGROUND',(0,1),(9,1),colors.HexColor('#fffbeb')))
    ttbl.setStyle(TableStyle(ts)); story.append(ttbl); story.append(Spacer(1,6*mm))

    story.append(Paragraph('SCORING KEY',sty('sec2',fontSize=8,leading=10,
                            textColor=C_GOLD,fontName='Helvetica-Bold',spaceAfter=4)))
    kd=[['Metric','Rule','Source'],
        ['Complaints','0-1%=4pts | 1-2%=3pts | 2-3%=1pt | >3%=0pt','Swiggy + Zomato'],
        ['KPT','Under 12 min = 1pt | 12 min or above = 0pt','Swiggy / Zomato'],
        ['Rating','4.0 or above = 1pt | Below 4.0 = 0pt','Zomato Average Rating'],
        ['Food Cost','Under 40% = 1pt | 40% or above = 0pt','Food Cost Compile Sheet'],
        ['Availability','98% or above = 1pt | Below 98% = 0pt','Swiggy Online Availability'],
        ['Hygiene','Manual input each month','Surprise visit scores'],
        ['Grade','Bronze 0-3 | Silver 3-6 | Gold 6-8 | Platinum 8+','Total Avg / Outlets']]
    ktbl=Table(kd,colWidths=[28*mm,82*mm,70*mm])
    ktbl.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),C_BG),('TEXTCOLOR',(0,0),(-1,0),C_GOLD),
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,0),7.5),
        ('FONTNAME',(0,1),(-1,-1),'Helvetica'),('FONTSIZE',(0,1),(-1,-1),7.5),
        ('FONTNAME',(0,1),(0,-1),'Helvetica-Bold'),
        ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#f5f5f5')]),
        ('ALIGN',(0,0),(-1,-1),'LEFT'),('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('TOPPADDING',(0,0),(-1,-1),2*mm),('BOTTOMPADDING',(0,0),(-1,-1),2*mm),
        ('LEFTPADDING',(0,0),(-1,-1),2*mm),
        ('GRID',(0,0),(-1,-1),0.3,colors.HexColor('#dddddd'))]))
    story.append(ktbl)

    if flags:
        story.append(PageBreak())
        story.append(Paragraph('FLAGGED OUTLETS',sty('sec3',fontSize=8,leading=10,
                                textColor=C_GOLD,fontName='Helvetica-Bold',spaceAfter=4)))
        story.append(Paragraph(f'{len(flags)} outlets breached thresholds this month.',
                                sty('fi',fontSize=8,leading=11,textColor=C_GREY,
                                    fontName='Helvetica',spaceAfter=4)))
        issue_types={}
        for f in flags:
            if f[2] not in issue_types: issue_types[f[2]]=[]
            issue_types[f[2]].append(f)
        ibgs={'High Complaints':colors.HexColor('#7f1d1d'),'KPT Exceeded':colors.HexColor('#7c2d12'),
              'Low Rating':colors.HexColor('#1e3a5f'),'Low Availability':colors.HexColor('#14532d'),
              'High Food Cost':colors.HexColor('#4a1d96')}
        for issue,iflags in issue_types.items():
            story.append(Spacer(1,3*mm))
            ih=Table([[Paragraph(f'<b>{issue}</b>',sty('ih',fontSize=8,leading=10,
                        textColor=colors.white,fontName='Helvetica-Bold')),
                       Paragraph(f'{len(iflags)} outlets',sty('ic',fontSize=8,leading=10,
                        textColor=colors.white,fontName='Helvetica',alignment=TA_RIGHT))]],
                     colWidths=[140*mm,40*mm])
            ih.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),ibgs.get(issue,C_BG)),
                ('LEFTPADDING',(0,0),(-1,-1),3*mm),('RIGHTPADDING',(0,0),(-1,-1),3*mm),
                ('TOPPADDING',(0,0),(-1,-1),2*mm),('BOTTOMPADDING',(0,0),(-1,-1),2*mm)]))
            story.append(ih)
            fr=[['Team Leader','Outlet','Value','Threshold']]+\
               [[f[0].split('(')[0].strip(),f[1],f[3],f[4]] for f in iflags]
            ft=Table(fr,colWidths=[45*mm,65*mm,25*mm,45*mm])
            ft.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f3f4f6')),
                ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),7.5),
                ('FONTNAME',(0,1),(-1,-1),'Helvetica'),
                ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#f5f5f5')]),
                ('ALIGN',(0,0),(-1,-1),'LEFT'),('VALIGN',(0,0),(-1,-1),'MIDDLE'),
                ('TOPPADDING',(0,0),(-1,-1),1.8*mm),('BOTTOMPADDING',(0,0),(-1,-1),1.8*mm),
                ('LEFTPADDING',(0,0),(-1,-1),2*mm),
                ('GRID',(0,0),(-1,-1),0.3,colors.HexColor('#e5e7eb'))]))
            story.append(ft)

    story.append(Spacer(1,6*mm))
    story.append(HRFlowable(width='100%',thickness=0.5,color=C_GREY))
    story.append(Spacer(1,2*mm))
    story.append(Paragraph(
        f'RollsKing Internal Report  ·  {month}  ·  Auto-generated  ·  Hygiene requires manual input',
        sty('foot',fontSize=7,leading=9,textColor=C_GREY,fontName='Helvetica',alignment=TA_CENTER)))
    doc.build(story)
    buf.seek(0)
    return buf.read()

# ── SESSION STATE ─────────────────────────────────────────────────────────────
defaults = {
    'logged_in': False, 'mapping': None, 'report_bytes': None,
    'report_name': None, 'pdf_bytes': None, 'pdf_name': None,
}
for k, v in defaults.items():
    if k not in st.session_state: st.session_state[k] = v

# ── LOGIN ─────────────────────────────────────────────────────────────────────
if not st.session_state.logged_in:
    st.markdown('<div class="main-title">RollsKing</div>', unsafe_allow_html=True)
    st.markdown('<div class="main-subtitle">Operations Report Generator</div>', unsafe_allow_html=True)
    pw = st.text_input("Password", type="password", placeholder="Enter access password")
    if st.button("Sign In"):
        if pw == APP_PASSWORD:
            st.session_state.logged_in = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    st.stop()

# ── MAIN APP ──────────────────────────────────────────────────────────────────
st.markdown('<div class="main-title">RollsKing Reports</div>', unsafe_allow_html=True)
st.markdown('<div class="main-subtitle">Monthly Performance Report Generator</div>', unsafe_allow_html=True)

tab_report, tab_mapping = st.tabs(["📊  Generate Report", "🗂️  Manage Mapping"])

# ════════════════════════════════════════════════════════════════════════════════
# ════════════════════════════════════════════════════════════════════════════════
# TAB 1 — GENERATE REPORT
# ════════════════════════════════════════════════════════════════════════════════
with tab_report:

    mapping_loaded = bool(st.session_state.mapping)

    if not mapping_loaded:
        st.markdown("""
        <div class="card-gold">
            <div class="section-label">One-Time Setup Required</div>
            <p style="color:#f0f0f0; font-size:0.92rem; margin:0.5rem 0 0.3rem;">
                Before generating reports, the system needs to know which outlets belong
                to which Team Leader. This is a
                <strong style="color:#e8a020;">one-time setup</strong> — done once,
                updated only when outlets or managers change.
            </p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("""
        <div class="card">
            <div class="section-label">How to Get Started</div>
            <div style="color:#ccc; font-size:0.88rem; line-height:2.2rem;">
                <div><span style="background:#e8a020;color:#000;border-radius:50%;width:22px;height:22px;
                display:inline-block;text-align:center;line-height:22px;font-weight:800;
                font-size:0.8rem;margin-right:10px;">1</span>
                Click the <strong style="color:#e8a020;">Manage Mapping</strong> tab above</div>
                <div><span style="background:#e8a020;color:#000;border-radius:50%;width:22px;height:22px;
                display:inline-block;text-align:center;line-height:22px;font-weight:800;
                font-size:0.8rem;margin-right:10px;">2</span>
                Upload <strong style="color:#e8a020;">Nov_Month_Data.xlsx</strong>
                (or any monthly file containing the mapping sheet)</div>
                <div><span style="background:#e8a020;color:#000;border-radius:50%;width:22px;height:22px;
                display:inline-block;text-align:center;line-height:22px;font-weight:800;
                font-size:0.8rem;margin-right:10px;">3</span>
                Return here and follow the 4 steps to generate your report</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("""
        <div class="card">
            <div class="section-label">What You Will Need Every Month</div>
            <div style="color:#ccc; font-size:0.88rem; line-height:2rem;">
                <div>📁 &nbsp;<strong style="color:#fff;">Monthly Data File</strong>
                &nbsp;—&nbsp; One .xlsx file with Zomato, Swiggy, Food Cost and Sale sheets</div>
                <div>✏️ &nbsp;<strong style="color:#fff;">Hygiene Scores</strong>
                &nbsp;—&nbsp; Manually enter surprise visit scores for each TL</div>
                <div>⬇️ &nbsp;<strong style="color:#fff;">Download Reports</strong>
                &nbsp;—&nbsp; Excel (detailed) + PDF (management summary) generated instantly</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    else:
        mapping = st.session_state.mapping
        tl_names = sorted(mapping.keys())
        total_outlets = sum(len(v) for v in mapping.values())

        st.markdown(f"""
        <div class="card">
            <div class="section-label">Status</div>
            <span class="status-ok">✓ Mapping Active &nbsp;·&nbsp; {len(tl_names)} Team Leaders &nbsp;·&nbsp; {total_outlets} Outlets</span>
            <div style="color:#555; font-size:0.78rem; margin-top:0.4rem;">
                Need to add or change an outlet? Go to the <strong>Manage Mapping</strong> tab.
            </div>
        </div>
        """, unsafe_allow_html=True)

        # STEP 1
        st.markdown("""<div style="margin-bottom:0.4rem;">
            <span class="step-badge">1</span>
            <span style="font-family:'Syne',sans-serif;font-size:0.85rem;font-weight:700;color:#fff;">
            Upload This Month's Data File</span>
            <span style="color:#555;font-size:0.8rem;margin-left:8px;">
            Must contain: Zomato, Swiggy, Food Cost and Sale sheets</span>
        </div>""", unsafe_allow_html=True)

        uploaded_files = st.file_uploader(
            "Drop monthly .xlsx file(s) here",
            type=["xlsx"],
            accept_multiple_files=True
        )

        detected = []
        if uploaded_files:
            for f in uploaded_files:
                fbytes = f.read()
                ftype, wb = detect_file_type(fbytes, f.name)
                detected.append({"name": f.name, "type": ftype, "bytes": fbytes, "wb": wb})
            for d in detected:
                icon  = "✓" if d["type"] == "monthly_raw" else "⚠"
                color = "chip-ok" if d["type"] == "monthly_raw" else "chip-warn"
                label = "Detected: Monthly Data File ✓" if d["type"] == "monthly_raw" else "Format not recognised — check sheet names"
                st.markdown(f'<span class="{color}">{icon} {d["name"]} — {label}</span>', unsafe_allow_html=True)

        st.markdown("<div style='margin:1.2rem 0;'></div>", unsafe_allow_html=True)

        # STEP 2
        st.markdown("""<div style="margin-bottom:0.4rem;">
            <span class="step-badge">2</span>
            <span style="font-family:'Syne',sans-serif;font-size:0.85rem;font-weight:700;color:#fff;">
            Select Report Month</span>
        </div>""", unsafe_allow_html=True)

        months = ["December 2025","November 2025","January 2026","February 2026",
                  "March 2026","April 2026","May 2026","June 2026",
                  "July 2026","August 2026","September 2026","October 2026"]
        sel_month = st.selectbox("Month", months, label_visibility="collapsed")

        st.markdown("<div style='margin:1.2rem 0;'></div>", unsafe_allow_html=True)

        # STEP 3
        st.markdown("""<div style="margin-bottom:0.4rem;">
            <span class="step-badge">3</span>
            <span style="font-family:'Syne',sans-serif;font-size:0.85rem;font-weight:700;color:#fff;">
            Enter Hygiene Scores</span>
            <span style="color:#555;font-size:0.8rem;margin-left:8px;">0–5 pts · Based on surprise visit this month</span>
        </div>""", unsafe_allow_html=True)

        hygiene_scores = {}
        cols = st.columns(2)
        for i, tl in enumerate(tl_names):
            with cols[i % 2]:
                short = tl.split("(")[0].strip()
                hygiene_scores[tl] = st.number_input(short, min_value=0, max_value=5, value=0, step=1, key=f"hyg_{tl}")

        st.markdown("<div style='margin:1.5rem 0;'></div>", unsafe_allow_html=True)

        # STEP 4
        st.markdown("""<div style="margin-bottom:0.5rem;">
            <span class="step-badge">4</span>
            <span style="font-family:'Syne',sans-serif;font-size:0.85rem;font-weight:700;color:#fff;">
            Generate Report</span>
        </div>""", unsafe_allow_html=True)

        valid_files = [d for d in detected if d["type"] == "monthly_raw"] if detected else []

        if not valid_files:
            st.markdown("""<div style="background:#1a1a1a;border:1px dashed #333;border-radius:10px;
            padding:1rem;color:#555;font-size:0.85rem;text-align:center;">
                Upload a monthly data file in Step 1 to enable report generation
            </div>""", unsafe_allow_html=True)
        else:
            if st.button("⚡ Generate Report"):
                with st.spinner("Processing data and building report..."):
                    try:
                        all_zmt, all_swg, all_fc = {}, {}, {}
                        for d in valid_files:
                            zmt, swg, fc = load_monthly_raw(d["wb"])
                            all_zmt.update(zmt)
                            all_swg.update(swg)
                            all_fc.update(fc)

                        results, disclaimers, flags = calculate_new(
                            mapping, all_zmt, all_swg, all_fc, hygiene_scores
                        )

                        excel_bytes = build_excel(results, disclaimers, flags, sel_month)
                        pdf_bytes   = build_pdf_report(results, flags, disclaimers, sel_month)
                        month_slug  = sel_month.replace(" ", "_")

                        st.session_state.report_bytes = excel_bytes
                        st.session_state.report_name  = f"RollsKing_Report_{month_slug}.xlsx"
                        st.session_state.pdf_bytes    = pdf_bytes
                        st.session_state.pdf_name     = f"RollsKing_Report_{month_slug}.pdf"

                        st.success(
                            f"✓ Report ready — {len(results)} Team Leaders · "
                            f"{sum(r['outlets'] for r in results)} Outlets · "
                            f"{len(flags)} Flags"
                        )
                    except Exception as e:
                        st.error(f"Error: {e}")
                        import traceback; st.code(traceback.format_exc())

        if st.session_state.report_bytes:
            st.markdown("<hr class='divider'>", unsafe_allow_html=True)
            st.markdown('<div class="section-label">Download Reports</div>', unsafe_allow_html=True)
            c1, c2 = st.columns(2)
            with c1:
                st.download_button("📥 Download Excel Report",
                    data=st.session_state.report_bytes,
                    file_name=st.session_state.report_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with c2:
                st.download_button("📄 Download PDF Summary",
                    data=st.session_state.pdf_bytes,
                    file_name=st.session_state.pdf_name,
                    mime="application/pdf")

# ════════════════════════════════════════════════════════════════════════════════
# TAB 2 — MANAGE MAPPING
# ════════════════════════════════════════════════════════════════════════════════
with tab_mapping:

    st.markdown("""
    <div class="card">
        <div class="section-label">What is Mapping?</div>
        <p style="color:#ccc;font-size:0.88rem;line-height:1.7rem;margin:0.3rem 0 0;">
            Mapping tells the system which outlets belong to which Team Leader,
            and links each outlet to its Zomato and Swiggy restaurant IDs.
            <strong style="color:#e8a020;">Upload this once</strong> — it stays saved
            until you update it. Re-upload only when outlets open, close, or TLs change.
        </p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""<div style="margin-bottom:0.4rem;">
        <span class="step-badge">1</span>
        <span style="font-family:'Syne',sans-serif;font-size:0.85rem;font-weight:700;color:#fff;">
        Upload Mapping File</span>
        <span style="color:#555;font-size:0.8rem;margin-left:8px;">
        Use Nov_Month_Data.xlsx · Sheet must be named 'Sheet1'</span>
    </div>""", unsafe_allow_html=True)

    map_file = st.file_uploader("Upload mapping file (.xlsx)", type=["xlsx"], key="map_uploader", label_visibility="collapsed")
    if map_file:
        mb = map_file.read()
        ftype, wb = detect_file_type(mb, map_file.name)
        if wb:
            mapping_data, err = parse_mapping_from_wb(wb)
            if err:
                st.error(f"Could not parse mapping: {err}")
            else:
                st.session_state.mapping = mapping_data
                total = sum(len(v) for v in mapping_data.values())
                st.success(f"✓ Mapping saved — {len(mapping_data)} Team Leaders · {total} Outlets · Go to Generate Report tab to continue")

    st.markdown("<div style='margin:1.2rem 0;'></div>", unsafe_allow_html=True)

    if st.session_state.mapping:
        tl_count  = len(st.session_state.mapping)
        out_count = sum(len(v) for v in st.session_state.mapping.values())

        st.markdown(f"""
        <div class="card">
            <div class="section-label">Current Mapping</div>
            <span class="status-ok">✓ {tl_count} Team Leaders · {out_count} Outlets loaded</span>
        </div>""", unsafe_allow_html=True)

        for tl, outlets in sorted(st.session_state.mapping.items()):
            short_tl = tl.split("(")[0].strip()
            with st.expander(f"{short_tl}  —  {len(outlets)} outlets"):
                for o in outlets:
                    st.markdown(
                        f"**{o['outlet']}** &nbsp;·&nbsp; POS: `{o['pos']}` "
                        f"&nbsp;·&nbsp; Zomato: `{o['zmt_rk']}` / `{o['zmt_rf']}` "
                        f"&nbsp;·&nbsp; Swiggy: `{o['swg_rk']}` / `{o['swg_rf']}`"
                    )

        st.markdown("<div style='margin:1.2rem 0;'></div>", unsafe_allow_html=True)

        st.markdown("""<div style="margin-bottom:0.4rem;">
            <span class="step-badge">2</span>
            <span style="font-family:'Syne',sans-serif;font-size:0.85rem;font-weight:700;color:#fff;">
            Add a New Outlet</span>
            <span style="color:#555;font-size:0.8rem;margin-left:8px;">
            Only needed when a new outlet opens or changes TL</span>
        </div>""", unsafe_allow_html=True)

        tl_options = sorted(st.session_state.mapping.keys())
        sel_tl = st.selectbox("Assign to Team Leader", tl_options, key="add_tl")
        c1, c2 = st.columns(2)
        with c1:
            new_outlet = st.text_input("Outlet Name", placeholder="e.g. Sector 62, Noida", key="new_outlet")
            new_pos    = st.text_input("POS ID", placeholder="e.g. 23687", key="new_pos")
            new_zrk    = st.text_input("Zomato ID (RollsKing)", placeholder="e.g. 19476740", key="new_zrk")
            new_zrf    = st.text_input("Zomato ID (Rolling Fresh)", placeholder="e.g. 20884624", key="new_zrf")
        with c2:
            new_srk    = st.text_input("Swiggy ID (RollsKing)", placeholder="e.g. 313666", key="new_srk")
            new_srf    = st.text_input("Swiggy ID (Rolling Fresh)", placeholder="e.g. 783919", key="new_srf")
            st.caption("Outlet Name and POS ID are required. IDs can be found on Zomato/Swiggy partner portals.")

        if st.button("➕ Add Outlet to Mapping"):
            if new_outlet and new_pos:
                st.session_state.mapping[sel_tl].append({
                    "outlet": new_outlet.strip(), "pos": safe_id(new_pos),
                    "zmt_rk": safe_id(new_zrk) if new_zrk else None,
                    "zmt_rf": safe_id(new_zrf) if new_zrf else None,
                    "swg_rk": safe_id(new_srk) if new_srk else None,
                    "swg_rf": safe_id(new_srf) if new_srf else None,
                })
                st.success(f"✓ {new_outlet} added under {sel_tl.split('(')[0].strip()}")
                st.rerun()
            else:
                st.warning("Outlet Name and POS ID are required.")
    else:
        st.markdown("""
        <div class="card">
            <div class="section-label">No Mapping Loaded Yet</div>
            <p style="color:#888;font-size:0.88rem;margin:0.3rem 0 0;">
                Upload the mapping file above to get started. Once loaded, it will remain
                active for all reports until you update it.
            </p>
        </div>""", unsafe_allow_html=True)

# ── FOOTER ────────────────────────────────────────────────────────────────────
st.markdown("""
<div style='text-align:center; color:#444; font-size:0.75rem; padding:2rem 0 1rem;'>
    RollsKing Internal Tools · Built for Operations
</div>
""", unsafe_allow_html=True)

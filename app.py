import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import io

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
.main-subtitle { color: #666; font-size: 0.95rem; margin-bottom: 1.5rem; }
.section-label {
    font-family: 'Syne', sans-serif;
    font-size: 0.78rem; font-weight: 700;
    color: #e8a020; letter-spacing: 1.5px;
    text-transform: uppercase; margin-bottom: 0.4rem;
}
.card {
    background: #1a1a1a; border: 1px solid #2a2a2a;
    border-radius: 12px; padding: 1.1rem 1.3rem;
    margin-bottom: 1rem;
}
.card-gold {
    background: #1a1500; border: 1px solid #e8a020;
    border-radius: 12px; padding: 1.1rem 1.3rem;
    margin-bottom: 1rem;
}
.status-ok  { color: #4ade80; font-weight: 600; font-size: 0.9rem; }
.status-warn{ color: #fbbf24; font-weight: 600; font-size: 0.9rem; }
.chip-ok   { background:#14532d; color:#4ade80; border-radius:6px; padding:3px 10px; font-size:0.82rem; display:inline-block; margin:2px 0; }
.chip-warn { background:#7c2d12; color:#fca5a5; border-radius:6px; padding:3px 10px; font-size:0.82rem; display:inline-block; margin:2px 0; }
.step-badge {
    background: #e8a020; color: #000; border-radius: 50%;
    width: 22px; height: 22px; display: inline-block;
    text-align: center; line-height: 22px;
    font-weight: 800; font-size: 0.8rem; margin-right: 8px;
}
hr.divider { border: none; border-top: 1px solid #2a2a2a; margin: 1.5rem 0; }
.stButton > button {
    background: #e8a020 !important; color: #000 !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important; border-radius: 8px !important;
    border: none !important; padding: 0.55rem 1.5rem !important;
    font-size: 0.9rem !important;
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
</style>
""", unsafe_allow_html=True)

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
        s = str(v).replace('%', '').strip()
        f = float(s)
        return f if f > 1 else f * 100
    except: return None

def parse_min(v):
    if v is None: return None
    try: return float(str(v).replace(' mins','').replace(' min','').strip())
    except: return None

def score_c(pct):
    if pct is None: return 0
    if pct <= 1:  return 4
    if pct <= 2:  return 3
    if pct <= 3:  return 1
    return 0

def score_tier(avg):
    if avg >= 8: return "Platinum"
    if avg >= 6: return "Gold"
    if avg >= 3: return "Silver"
    return "Bronze"

# ── HARDCODED MAPPING ─────────────────────────────────────────────────────────
DEFAULT_MAPPING = {
  "Navneet Singh": [
    {"outlet":"Sec 104","pos":28039,"zmt_rk":19476740,"zmt_rf":20884624,"swg_rk":313666,"swg_rf":783919},
    {"outlet":"Sector-141 Noida","pos":26592,"zmt_rk":18734595,"zmt_rf":21198454,"swg_rk":63465,"swg_rf":882814},
    {"outlet":"Sector-132 Noida","pos":34303,"zmt_rk":18750756,"zmt_rf":21563311,"swg_rk":68184,"swg_rf":998424},
    {"outlet":"Sector 125, Noida","pos":373602,"zmt_rk":21824279,"zmt_rf":21819721,"swg_rk":1069514,"swg_rf":1069500},
    {"outlet":"Sector-73,Noida","pos":97074,"zmt_rk":20508934,"zmt_rf":20511693,"swg_rk":635149,"swg_rf":641252},
    {"outlet":"sector-44 Noida.","pos":39966,"zmt_rk":18575970,"zmt_rf":21570341,"swg_rk":42813,"swg_rf":1035832}
  ],
  "Ajay Halder": [
    {"outlet":"Sector 4 Noida","pos":21787,"zmt_rk":19364731,"zmt_rf":20884617,"swg_rk":54622,"swg_rf":784582},
    {"outlet":"Sector-62","pos":23687,"zmt_rk":302308,"zmt_rf":21292616,"swg_rk":42808,"swg_rf":907194},
    {"outlet":"sector-35","pos":32851,"zmt_rk":20374787,"zmt_rf":20884635,"swg_rk":583789,"swg_rf":798506},
    {"outlet":"Sector-18","pos":26952,"zmt_rk":304612,"zmt_rf":None,"swg_rk":42807,"swg_rf":1238358},
    {"outlet":"Gaur City, G.Noida","pos":112772,"zmt_rk":20589872,"zmt_rf":21087471,"swg_rk":879431,"swg_rf":940381},
    {"outlet":"Eco Loft","pos":74178,"zmt_rk":20264919,"zmt_rf":22245955,"swg_rk":531120,"swg_rf":1227942}
  ],
  "Sunil Sharma": [
    {"outlet":"RDC Raj Nagar, Gzb","pos":363143,"zmt_rk":18962941,"zmt_rf":21184529,"swg_rk":879460,"swg_rf":883371},
    {"outlet":"GNB Mall","pos":113953,"zmt_rk":21341669,"zmt_rf":21340893,"swg_rk":1082869,"swg_rf":1089374},
    {"outlet":"Shipra Mall","pos":408910,"zmt_rk":22103426,"zmt_rf":None,"swg_rk":1238359,"swg_rf":None}
  ],
  "Vishwanath Rao": [
    {"outlet":"Indirapuram","pos":38041,"zmt_rk":18633334,"zmt_rf":20884637,"swg_rk":46674,"swg_rf":784485},
    {"outlet":"Rajendra Nagar Gzb","pos":37055,"zmt_rk":19283683,"zmt_rf":21563308,"swg_rk":241917,"swg_rf":998786},
    {"outlet":"Vasundhra","pos":122466,"zmt_rk":20711593,"zmt_rf":21780276,"swg_rk":731841,"swg_rf":1069774}
  ],
  "Sanjay Morya": [
    {"outlet":"Kalkaji","pos":31247,"zmt_rk":18869459,"zmt_rf":20884610,"swg_rk":90719,"swg_rf":783920},
    {"outlet":"Tilak Nagar","pos":63819,"zmt_rk":18942689,"zmt_rf":20884619,"swg_rk":123197,"swg_rf":786729},
    {"outlet":"Vasant Kunj","pos":25924,"zmt_rk":19030978,"zmt_rf":21198457,"swg_rk":131217,"swg_rf":883284},
    {"outlet":"Chattarpur","pos":26423,"zmt_rk":19052007,"zmt_rf":21198467,"swg_rk":140433,"swg_rf":883277},
    {"outlet":"Paschim Vihar","pos":43412,"zmt_rk":20256577,"zmt_rf":21751304,"swg_rk":531480,"swg_rf":1065481},
    {"outlet":"Gtb Nagar","pos":79050,"zmt_rk":20323930,"zmt_rf":21929075,"swg_rk":569414,"swg_rf":1101863},
    {"outlet":"Nathupur,Gurugram","pos":108782,"zmt_rk":20582763,"zmt_rf":20884615,"swg_rk":668475,"swg_rf":783922},
    {"outlet":"Old DLF Sec-14,Gurugram","pos":108777,"zmt_rk":20582827,"zmt_rf":20884599,"swg_rk":668470,"swg_rf":854941},
    {"outlet":"Sector-57 Gurugram","pos":74068,"zmt_rk":20463325,"zmt_rf":20964530,"swg_rk":624165,"swg_rf":803919},
    {"outlet":"Wazirabad,Gurugram","pos":108779,"zmt_rk":20582847,"zmt_rf":None,"swg_rk":668467,"swg_rf":None},
    {"outlet":"Gurugram Sec-82","pos":30407,"zmt_rk":19513923,"zmt_rf":21087476,"swg_rk":327106,"swg_rf":850542},
    {"outlet":"Sector 90, Gurugram","pos":380769,"zmt_rk":21929020,"zmt_rf":21929049,"swg_rk":1102249,"swg_rf":1101896},
    {"outlet":"Rohini","pos":93493,"zmt_rk":22100897,"zmt_rf":22101011,"swg_rk":622353,"swg_rf":1167156},
    {"outlet":"Vikashpuri","pos":404584,"zmt_rk":22227640,"zmt_rf":22227658,"swg_rk":1224350,"swg_rf":1224348},
    {"outlet":"Subhash Nagar","pos":398993,"zmt_rk":22165860,"zmt_rf":22165921,"swg_rk":1196325,"swg_rf":1196321}
  ],
  "Zeeshan Ali": [
    {"outlet":"Shaheen Bagh","pos":118685,"zmt_rk":20666436,"zmt_rf":21190578,"swg_rk":704360,"swg_rf":879433},
    {"outlet":"NIT,Faridabad","pos":96843,"zmt_rk":20480333,"zmt_rf":21087481,"swg_rk":632083,"swg_rf":852302},
    {"outlet":"Sec-15 Faridabad","pos":54369,"zmt_rk":18567324,"zmt_rf":21702217,"swg_rk":42815,"swg_rf":1037447},
    {"outlet":"Lakkarpur,Faridabad","pos":143500,"zmt_rk":20873208,"zmt_rf":21087485,"swg_rk":775707,"swg_rf":855113},
    {"outlet":"Greenfield, Faridabad","pos":154254,"zmt_rk":21446399,"zmt_rf":21446783,"swg_rk":983943,"swg_rf":991036}
  ],
  "Badir Alam": [
    {"outlet":"Bhopal","pos":338959,"zmt_rk":21340655,"zmt_rf":21340565,"swg_rk":934354,"swg_rf":937374},
    {"outlet":"Indore","pos":109589,"zmt_rk":20566161,"zmt_rf":21304975,"swg_rk":673809,"swg_rf":920802},
    {"outlet":"Siddharth Nagar,Indore","pos":156653,"zmt_rk":21022031,"zmt_rf":21643899,"swg_rk":690867,"swg_rf":1027884}
  ],
  "Abhishek Kumar": [
    {"outlet":"Whitefield Bangalore","pos":89397,"zmt_rk":20410563,"zmt_rf":21075165,"swg_rk":606509,"swg_rf":850483},
    {"outlet":"Mahadevpura (Bangalore)","pos":72269,"zmt_rk":20201048,"zmt_rf":20790266,"swg_rk":515199,"swg_rf":700101},
    {"outlet":"Koramangala","pos":83769,"zmt_rk":20359621,"zmt_rf":20790279,"swg_rk":580691,"swg_rf":709590},
    {"outlet":"Electronic City Bangalore","pos":72413,"zmt_rk":20213913,"zmt_rf":21087516,"swg_rk":515053,"swg_rf":848482},
    {"outlet":"Bangalore,Sarjapur","pos":68691,"zmt_rk":20163232,"zmt_rf":20790275,"swg_rk":494751,"swg_rf":649361},
    {"outlet":"Kalyan Nagar Bangalore","pos":75899,"zmt_rk":20265149,"zmt_rf":21087530,"swg_rk":544214,"swg_rf":848481},
    {"outlet":"Bel Road Banglore","pos":75897,"zmt_rk":20263151,"zmt_rf":21037571,"swg_rk":536015,"swg_rf":728281},
    {"outlet":"Habble Banglore","pos":95682,"zmt_rk":20471662,"zmt_rf":None,"swg_rk":625912,"swg_rf":None},
    {"outlet":"Bangalore (Indira Nagar)","pos":403199,"zmt_rk":22179137,"zmt_rf":22179218,"swg_rk":1203098,"swg_rf":1203101}
  ],
  "Virendra Pratap": [
    {"outlet":"Mohanram Nagar","pos":84743,"zmt_rk":20410863,"zmt_rf":20994725,"swg_rk":588878,"swg_rf":808838},
    {"outlet":"Madipakkam","pos":84742,"zmt_rk":20410826,"zmt_rf":21627457,"swg_rk":588790,"swg_rf":1021132},
    {"outlet":"Parengudi,Chennai","pos":97078,"zmt_rk":20486896,"zmt_rf":21087508,"swg_rk":631195,"swg_rf":848486}
  ],
  "Atul Kumar": [
    {"outlet":"Apple Ghar , Pune","pos":129883,"zmt_rk":20748035,"zmt_rf":21044940,"swg_rk":733937,"swg_rf":741458},
    {"outlet":"Hinjewadi Phase 1, Pune","pos":141096,"zmt_rk":20855724,"zmt_rf":21049119,"swg_rk":756772,"swg_rf":734618},
    {"outlet":"Millennium Mall Pune","pos":137998,"zmt_rk":21067154,"zmt_rf":None,"swg_rk":833916,"swg_rf":None},
    {"outlet":"Shivaji Nagar Pune","pos":346318,"zmt_rk":21435196,"zmt_rf":21604740,"swg_rk":354312,"swg_rf":1009928}
  ],
  "Bhupesh Bhatt": [
    {"outlet":"Madhapur","pos":24485,"zmt_rk":18953624,"zmt_rf":21049126,"swg_rk":120196,"swg_rf":698521},
    {"outlet":"Gachibowli","pos":24487,"zmt_rk":19271816,"zmt_rf":21044950,"swg_rk":214621,"swg_rf":773418},
    {"outlet":"Banjara Hills","pos":129436,"zmt_rk":21028217,"zmt_rf":21080628,"swg_rk":711834,"swg_rf":844883},
    {"outlet":"Taranagar,Hyderabad","pos":141099,"zmt_rk":20855101,"zmt_rf":21080636,"swg_rk":766665,"swg_rf":844889},
    {"outlet":"R K Puram Hyderabad","pos":44718,"zmt_rk":19714313,"zmt_rf":None,"swg_rk":375980,"swg_rf":None},
    {"outlet":"Lulu Mall, Hyderabad","pos":141214,"zmt_rk":21154081,"zmt_rf":None,"swg_rk":866698,"swg_rf":None},
    {"outlet":"Miyapur","pos":24489,"zmt_rk":21779883,"zmt_rf":21779942,"swg_rk":1063096,"swg_rf":1061876},
    {"outlet":"Goa (Anjuna)","pos":339817,"zmt_rk":21365117,"zmt_rf":22213849,"swg_rk":946493,"swg_rf":1203077}
  ],
  "Milan": [
    {"outlet":"G Corp","pos":367105,"zmt_rk":21734559,"zmt_rf":21865596,"swg_rk":1063584,"swg_rf":1079375},
    {"outlet":"Mumbai Pawai","pos":353027,"zmt_rk":21522379,"zmt_rf":21618090,"swg_rk":985771,"swg_rf":1005642},
    {"outlet":"Raymond","pos":369670,"zmt_rk":21794492,"zmt_rf":21794546,"swg_rk":1066710,"swg_rf":1066711},
    {"outlet":"Mumbai BKC","pos":375018,"zmt_rk":21824216,"zmt_rf":21824173,"swg_rk":1076952,"swg_rf":1102348},
    {"outlet":"Mumbai Chembur","pos":383109,"zmt_rk":21966993,"zmt_rf":22030585,"swg_rk":1104174,"swg_rf":1140316},
    {"outlet":"Airoli Navi Mumbai City","pos":386396,"zmt_rk":21982077,"zmt_rf":22030515,"swg_rk":1123441,"swg_rf":1140313},
    {"outlet":"Mira Road MumbaiCity","pos":386734,"zmt_rk":22044961,"zmt_rf":22150856,"swg_rk":1142360,"swg_rf":1179699}
  ]
}

# ── FILE DETECTION ────────────────────────────────────────────────────────────
def detect_file_type(file_bytes, filename):
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        sheets = [s.lower() for s in wb.sheetnames]
        has_zomato   = any('zomato' in s for s in sheets)
        has_swiggy   = any('swiggy' in s for s in sheets)
        has_foodcost = any('food cost' in s for s in sheets)
        has_sale     = any('sale' in s for s in sheets)
        if has_zomato and has_swiggy and has_foodcost and has_sale:
            return 'monthly_raw', wb
        return 'unknown', wb
    except:
        return 'unknown', None

# ── LOAD MONTHLY DATA ─────────────────────────────────────────────────────────
def load_monthly_raw(wb):
    # Zomato
    zmt = {}
    for name in wb.sheetnames:
        if 'zomato' in name.lower():
            for row in wb[name].iter_rows(min_row=2, values_only=True):
                rid = safe_id(row[0])
                if not rid: continue
                metric = str(row[5]).strip() if row[5] else ''
                val = row[6]
                if rid not in zmt:
                    zmt[rid] = {'orders':0,'complaints':0,'kpt':None,'rating':None}
                if   metric == 'Delivered orders':    zmt[rid]['orders']     = safe_f(val)
                elif metric == 'Total complaints':    zmt[rid]['complaints'] = safe_f(val)
                elif metric == 'KPT (in minutes)':   zmt[rid]['kpt']        = safe_f(val)
                elif metric == 'Average rating':      zmt[rid]['rating']     = safe_f(val)
            break

    # Swiggy
    swg = {}
    for name in wb.sheetnames:
        if 'swiggy' in name.lower():
            for row in wb[name].iter_rows(min_row=2, values_only=True):
                rid = safe_id(row[0])
                if not rid: continue
                metric = str(row[5]).strip() if row[5] else ''
                val = row[6]
                if rid not in swg:
                    swg[rid] = {'kpt':None,'avail':None,'cmp_pct':None,'orders':0}
                if   metric == 'Kitchen Prep Time':
                    swg[rid]['kpt'] = parse_min(str(val).replace(' mins','').replace(' min','')) if val else None
                elif metric == 'Online Availability %':       swg[rid]['avail']   = parse_pct(val)
                elif metric == '% Orders with Complaints':    swg[rid]['cmp_pct'] = parse_pct(val)
                elif metric in ('Delivered Orders','Orders'): swg[rid]['orders']  = safe_f(val)
            break

    # Food Cost
    fc = {}
    for name in wb.sheetnames:
        if 'food cost' in name.lower():
            for row in wb[name].iter_rows(min_row=2, values_only=True):
                pos = safe_id(row[1])
                if not pos: continue
                net_sale = safe_f(row[4]); fc_val = row[9]
                if fc_val is not None:
                    fc_pct = safe_f(fc_val)*100 if safe_f(fc_val)<2 else safe_f(fc_val)
                elif net_sale > 0:
                    cogs = safe_f(row[5]) + safe_f(row[7]) + safe_f(row[8]) - safe_f(row[6])
                    fc_pct = round(cogs/net_sale*100, 2)
                else:
                    fc_pct = None
                fc[pos] = {'fc_pct': fc_pct, 'net_sale': net_sale}
            break

    # Delivered orders = Zomato + Swiggy combined (for sales scoring)
    delivered = {}
    for rid, d in zmt.items():
        delivered[rid] = delivered.get(rid, 0) + d.get('orders', 0)
    for rid, d in swg.items():
        delivered[rid] = delivered.get(rid, 0) + d.get('orders', 0)

    return zmt, swg, fc, delivered

# ── CALCULATOR ────────────────────────────────────────────────────────────────
def calculate(mapping, zmt, swg, fc, hygiene_scores, prev_delivered=None):
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

            # Complaints
            s_cmp = next((swg[s]['cmp_pct'] for s in sids if s in swg and swg[s].get('cmp_pct') is not None), None)
            z_ord = sum(zmt[r]['orders']     for r in zids if r in zmt)
            z_cmp = sum(zmt[r]['complaints'] for r in zids if r in zmt)
            z_pct = round(z_cmp/z_ord*100, 2) if z_ord > 0 else None
            if s_cmp is not None:
                cmp_pts = score_c(s_cmp) + score_c(z_pct)
                cmp_disp = round(s_cmp, 2); cmp_src = "Swiggy+Zomato"
            elif z_pct is not None:
                cmp_pts = score_c(z_pct); cmp_disp = z_pct; cmp_src = "Zomato only"
            else:
                cmp_pts = 0; cmp_disp = 0; cmp_src = "No data"
                notes.append("No complaint data")
                disclaimers.append(f"{tl} | {outlet}: complaint data missing")
            tl_c += cmp_pts
            if cmp_disp and cmp_disp > 3:
                flags.append((tl, outlet, "High Complaints", f"{cmp_disp:.1f}%", ">3%", cmp_pts))

            # KPT
            kpt_vals = [swg[s]['kpt'] for s in sids if s in swg and swg[s].get('kpt') is not None]
            if not kpt_vals:
                kpt_vals = [zmt[r]['kpt'] for r in zids if r in zmt and zmt[r].get('kpt') is not None]
            if kpt_vals:
                avg_kpt = round(sum(kpt_vals)/len(kpt_vals), 2)
                kpt_pts = 1 if avg_kpt < 12 else 0
                kpt_src = "Swiggy" if any(swg.get(s,{}).get('kpt') for s in sids) else "Zomato"
            else:
                avg_kpt = None; kpt_pts = 0; kpt_src = "N/A"
                notes.append("No KPT data")
                disclaimers.append(f"{tl} | {outlet}: KPT missing")
            tl_k += kpt_pts
            if avg_kpt and avg_kpt >= 12:
                flags.append((tl, outlet, "KPT Exceeded", f"{avg_kpt:.1f} min", "≥12 min", kpt_pts))

            # Rating
            rat_vals = [zmt[r]['rating'] for r in zids if r in zmt and zmt[r].get('rating')]
            avg_rat = round(sum(rat_vals)/len(rat_vals), 2) if rat_vals else 0
            rat_pts = 1 if avg_rat >= 4.0 else 0
            tl_r += rat_pts
            if 0 < avg_rat < 4.0:
                flags.append((tl, outlet, "Low Rating", f"{avg_rat:.2f}", "<4.0", rat_pts))

            # Availability
            avail = next((swg[s]['avail'] for s in sids if s in swg and swg[s].get('avail') is not None), None)
            avail_pts = 1 if (avail is not None and avail >= 98) else 0
            tl_av += avail_pts
            if avail is not None and avail < 98:
                flags.append((tl, outlet, "Low Availability", f"{avail:.1f}%", "<98%", avail_pts))
            if avail is None:
                notes.append("No availability data")
                disclaimers.append(f"{tl} | {outlet}: availability missing")

            # Food Cost
            fc_data = fc.get(pos)
            if fc_data and fc_data['fc_pct'] is not None:
                fc_pct = round(fc_data['fc_pct'], 2)
                fc_pts = 1 if fc_pct < 40 else 0
            else:
                fc_pct = None; fc_pts = 0
                notes.append("FC data missing")
                disclaimers.append(f"{tl} | {outlet}: food cost missing")
            tl_fc += fc_pts
            if fc_pct is not None and fc_pct >= 40:
                flags.append((tl, outlet, "High Food Cost", f"{fc_pct:.1f}%", "≥40%", fc_pts))

            outlet_rows.append({
                'outlet': outlet, 'pos': pos,
                'cmp_pct': cmp_disp, 'cmp_pts': cmp_pts, 'cmp_src': cmp_src,
                'kpt_avg': avg_kpt,  'kpt_pts': kpt_pts, 'kpt_src': kpt_src,
                'rat_avg': avg_rat,  'rat_pts': rat_pts,
                'avail_pct': avail,  'avail_pts': avail_pts,
                'fc_pct': fc_pct,    'fc_pts': fc_pts,
                'notes': "; ".join(notes) if notes else "OK"
            })

        # Sales (month-on-month)
        if prev_delivered:
            ids = [x for o in outlets for x in [o['zmt_rk'],o['zmt_rf'],o['swg_rk'],o['swg_rf']] if x]
            curr_total = sum(zmt.get(i,{}).get('orders',0) for i in ids) + \
                         sum(swg.get(i,{}).get('orders',0) for i in ids)
            prev_total = sum(prev_delivered.get(i,0) for i in ids)
            sales_pts = 1 if (prev_total > 0 and curr_total > prev_total) else 0
        else:
            sales_pts = 0

        hyg_val   = hygiene_scores.get(tl, 0)
        total_pts = tl_c + tl_k + tl_r + tl_fc + hyg_val + tl_av + sales_pts
        avg_score = round(total_pts / n, 1) if n > 0 else 0
        tier      = score_tier(avg_score)

        results.append({
            'tl': tl, 'outlets': n, 'sales_pts': sales_pts,
            'fc_pts': tl_fc, 'cmp_pts': tl_c, 'kpt_pts': tl_k,
            'rat_pts': tl_r, 'hyg_pts': hyg_val, 'avail_pts': tl_av,
            'total_pts': total_pts, 'avg_score': avg_score,
            'tier': tier, 'outlet_detail': outlet_rows
        })

    return results, disclaimers, flags


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
        ['Sales','Current month orders > prev month = 1pt | Decline = 0pt','Zomato + Swiggy Delivered Orders'],
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
if 'logged_in'    not in st.session_state: st.session_state.logged_in    = False
if 'report_bytes' not in st.session_state: st.session_state.report_bytes = None
if 'report_name'  not in st.session_state: st.session_state.report_name  = None
if 'pdf_bytes'    not in st.session_state: st.session_state.pdf_bytes    = None
if 'pdf_name'     not in st.session_state: st.session_state.pdf_name     = None
if 'mapping'      not in st.session_state:
    st.session_state.mapping = {k: list(v) for k, v in DEFAULT_MAPPING.items()}

# ── LOGIN ─────────────────────────────────────────────────────────────────────
APP_PASSWORD = "rollsking2025"
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

mapping = st.session_state.mapping
tl_names = sorted(mapping.keys())
total_outlets = sum(len(v) for v in mapping.values())

tab_report, tab_mapping = st.tabs(["📊  Generate Report", "🗂️  Manage Mapping"])

# ════════════════════════════════════════════════════════════════════════════════
# TAB 1 — GENERATE REPORT
# ════════════════════════════════════════════════════════════════════════════════
with tab_report:

    st.markdown(f"""
    <div class="card">
        <div class="section-label">Status</div>
        <span class="status-ok">✓ {len(tl_names)} Team Leaders &nbsp;·&nbsp; {total_outlets} Outlets active</span>
        <div style="color:#555;font-size:0.78rem;margin-top:0.3rem;">
            Mapping is built-in. Use Manage Mapping tab to add new outlets or TLs.
        </div>
    </div>
    """, unsafe_allow_html=True)

    # STEP 1 — UPLOAD
    st.markdown("""<div style="margin-bottom:0.5rem;">
        <span class="step-badge">1</span>
        <span style="font-family:'Syne',sans-serif;font-size:0.9rem;font-weight:700;color:#fff;">
        Upload Data Files</span>
        <span style="color:#555;font-size:0.8rem;margin-left:8px;">
        For Sales scoring: upload current month + previous month together</span>
    </div>""", unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "Drop .xlsx files here (current month required, previous month optional for Sales scoring)",
        type=["xlsx"],
        accept_multiple_files=True,
        label_visibility="collapsed"
    )

    detected = []
    if uploaded_files:
        for f in uploaded_files:
            fbytes = f.read()
            ftype, wb = detect_file_type(fbytes, f.name)
            detected.append({"name": f.name, "type": ftype, "bytes": fbytes, "wb": wb})
        for d in detected:
            ok = d["type"] == "monthly_raw"
            icon  = "✓" if ok else "⚠"
            color = "chip-ok" if ok else "chip-warn"
            label = "Monthly Data ✓" if ok else "Unrecognised format — needs Zomato, Swiggy, Food Cost, Sale sheets"
            st.markdown(f'<span class="{color}">{icon} {d["name"]} — {label}</span>', unsafe_allow_html=True)

    st.markdown("<div style='margin:1rem 0;'></div>", unsafe_allow_html=True)

    # STEP 2 — MONTH
    st.markdown("""<div style="margin-bottom:0.5rem;">
        <span class="step-badge">2</span>
        <span style="font-family:'Syne',sans-serif;font-size:0.9rem;font-weight:700;color:#fff;">
        Select Report Month</span>
    </div>""", unsafe_allow_html=True)
    months = ["January 2026","February 2026","March 2026","April 2026","May 2026",
              "June 2026","July 2026","August 2026","September 2026","October 2026",
              "November 2025","December 2025"]
    sel_month = st.selectbox("Month", months, label_visibility="collapsed")

    st.markdown("<div style='margin:1rem 0;'></div>", unsafe_allow_html=True)

    # STEP 3 — HYGIENE
    st.markdown("""<div style="margin-bottom:0.5rem;">
        <span class="step-badge">3</span>
        <span style="font-family:'Syne',sans-serif;font-size:0.9rem;font-weight:700;color:#fff;">
        Hygiene Scores</span>
        <span style="color:#555;font-size:0.8rem;margin-left:8px;">0–5 pts · Based on surprise visit this month</span>
    </div>""", unsafe_allow_html=True)

    hygiene_scores = {}
    cols = st.columns(2)
    for i, tl in enumerate(tl_names):
        with cols[i % 2]:
            hygiene_scores[tl] = st.number_input(
                tl.split("(")[0].strip(), min_value=0, max_value=5, value=0, step=1, key=f"hyg_{tl}"
            )

    st.markdown("<div style='margin:1.2rem 0;'></div>", unsafe_allow_html=True)

    # STEP 4 — GENERATE
    st.markdown("""<div style="margin-bottom:0.5rem;">
        <span class="step-badge">4</span>
        <span style="font-family:'Syne',sans-serif;font-size:0.9rem;font-weight:700;color:#fff;">
        Generate Report</span>
    </div>""", unsafe_allow_html=True)

    valid = [d for d in detected if d["type"] == "monthly_raw"] if detected else []

    if not valid:
        st.markdown("""<div style="background:#1a1a1a;border:1px dashed #333;border-radius:10px;
        padding:1rem;color:#555;font-size:0.85rem;text-align:center;">
            Upload at least one monthly data file above to enable report generation
        </div>""", unsafe_allow_html=True)
    else:
        if st.button("⚡ Generate Report"):
            with st.spinner("Processing data... please wait"):
                try:
                    # Split files: current month vs previous month
                    curr_key = sel_month[:3].lower()  # "jan", "dec" etc
                    curr_files = [d for d in valid if curr_key in d["name"].lower()]
                    prev_files = [d for d in valid if curr_key not in d["name"].lower()]
                    if not curr_files:
                        curr_files = valid  # fallback: treat all as current

                    # Load current month
                    all_zmt, all_swg, all_fc, all_del = {}, {}, {}, {}
                    for d in curr_files:
                        z, s, f, dl = load_monthly_raw(d["wb"])
                        all_zmt.update(z); all_swg.update(s)
                        all_fc.update(f);  all_del.update(dl)

                    # Load previous month (for sales scoring)
                    prev_del = {}
                    for d in prev_files:
                        _, _, _, dl = load_monthly_raw(d["wb"])
                        prev_del.update(dl)

                    results, disclaimers, flags = calculate(
                        mapping, all_zmt, all_swg, all_fc, hygiene_scores,
                        prev_delivered=prev_del if prev_del else None
                    )

                    # Count how many TLs got sales pts
                    sales_scored = sum(1 for r in results if r['sales_pts'] == 1)
                    sales_note = f" · Sales: {sales_scored}/{len(results)} TLs grew" if prev_del else " · Sales: upload prev month for scoring"

                    excel_bytes = build_excel(results, disclaimers, flags, sel_month)
                    month_slug  = sel_month.replace(" ", "_")
                    st.session_state.report_bytes = excel_bytes
                    st.session_state.report_name  = f"RollsKing_Report_{month_slug}.xlsx"

                    try:
                        pdf_bytes = build_pdf_report(results, flags, disclaimers, sel_month)
                        st.session_state.pdf_bytes = pdf_bytes
                        st.session_state.pdf_name  = f"RollsKing_Report_{month_slug}.pdf"
                        pdf_ok = True
                    except ModuleNotFoundError:
                        st.session_state.pdf_bytes = None
                        pdf_ok = False

                    n_tls     = len(results)
                    n_outlets = sum(r["outlets"] for r in results)
                    n_flags   = len(flags)
                    st.success(f"✓ Report ready — {n_tls} TLs · {n_outlets} Outlets · {n_flags} Flags{sales_note}")
                    if not pdf_ok:
                        st.warning("PDF unavailable — reportlab missing on server. Excel is ready.")

                except Exception as e:
                    if "module" in str(e).lower():
                        st.error(f"Missing library: {e}. Check requirements.txt.")
                    elif "sheet" in str(e).lower():
                        st.error("Could not read file. Check sheet names: needs Zomato, Swiggy, Food Cost, Sale.")
                    else:
                        st.error(f"Error: {e}")
                    import traceback; st.code(traceback.format_exc())

    # DOWNLOADS
    if st.session_state.report_bytes:
        st.markdown("<hr class='divider'>", unsafe_allow_html=True)
        st.markdown('<div class="section-label">Download Reports</div>', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("📥 Download Excel",
                data=st.session_state.report_bytes,
                file_name=st.session_state.report_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with c2:
            if st.session_state.pdf_bytes:
                st.download_button("📄 Download PDF",
                    data=st.session_state.pdf_bytes,
                    file_name=st.session_state.pdf_name,
                    mime="application/pdf")
            else:
                st.markdown("""<div style="background:#1a1a1a;border:1px dashed #333;
                border-radius:8px;padding:0.6rem 1rem;color:#555;font-size:0.82rem;
                text-align:center;margin-top:4px;">
                    📄 PDF unavailable — reportlab missing on server
                </div>""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════════
# TAB 2 — MANAGE MAPPING
# ════════════════════════════════════════════════════════════════════════════════
with tab_mapping:

    st.markdown(f"""
    <div class="card">
        <div class="section-label">Current Mapping</div>
        <span class="status-ok">✓ {len(tl_names)} Team Leaders · {total_outlets} Outlets</span>
        <div style="color:#555;font-size:0.78rem;margin-top:0.3rem;">
            Mapping is built into the app. Changes here last this session only.
            For permanent changes, send outlet details to your developer.
        </div>
    </div>
    """, unsafe_allow_html=True)

    for tl, outlets in sorted(mapping.items()):
        with st.expander(f"{tl}  —  {len(outlets)} outlets"):
            for o in outlets:
                st.markdown(
                    f"**{o['outlet']}** &nbsp;·&nbsp; POS: `{o['pos']}` "
                    f"&nbsp;·&nbsp; Z: `{o['zmt_rk']}` / `{o['zmt_rf']}` "
                    f"&nbsp;·&nbsp; S: `{o['swg_rk']}` / `{o['swg_rf']}`"
                )

    st.markdown("<div style='margin:1.2rem 0;'></div>", unsafe_allow_html=True)

    # Add new TL
    st.markdown("""<div style="margin-bottom:0.4rem;">
        <span class="step-badge">+</span>
        <span style="font-family:'Syne',sans-serif;font-size:0.85rem;font-weight:700;color:#fff;">
        Add New Team Leader</span>
    </div>""", unsafe_allow_html=True)
    new_tl = st.text_input("Team Leader Name", placeholder="e.g. Rajesh Sharma", key="new_tl")
    if st.button("➕ Add Team Leader"):
        if new_tl.strip() and new_tl.strip() not in mapping:
            st.session_state.mapping[new_tl.strip()] = []
            st.success(f"✓ Added {new_tl.strip()}. Now add outlets to them below.")
            st.rerun()
        elif new_tl.strip() in mapping:
            st.warning("That Team Leader already exists.")
        else:
            st.warning("Enter a name first.")

    st.markdown("<div style='margin:1rem 0;'></div>", unsafe_allow_html=True)

    # Add new outlet
    st.markdown("""<div style="margin-bottom:0.4rem;">
        <span class="step-badge">+</span>
        <span style="font-family:'Syne',sans-serif;font-size:0.85rem;font-weight:700;color:#fff;">
        Add New Outlet</span>
    </div>""", unsafe_allow_html=True)

    sel_tl = st.selectbox("Assign to Team Leader", sorted(mapping.keys()), key="add_tl")
    c1, c2 = st.columns(2)
    with c1:
        new_outlet = st.text_input("Outlet Name", placeholder="e.g. Sector 62, Noida", key="new_outlet")
        new_pos    = st.text_input("POS ID", placeholder="e.g. 23687", key="new_pos")
        new_zrk    = st.text_input("Zomato ID (RollsKing)", placeholder="e.g. 19476740", key="new_zrk")
        new_zrf    = st.text_input("Zomato ID (Rolling Fresh)", placeholder="e.g. 20884624", key="new_zrf")
    with c2:
        new_srk    = st.text_input("Swiggy ID (RollsKing)", placeholder="e.g. 313666", key="new_srk")
        new_srf    = st.text_input("Swiggy ID (Rolling Fresh)", placeholder="e.g. 783919", key="new_srf")
        st.caption("Outlet Name and POS ID are required.")

    if st.button("➕ Add Outlet"):
        if new_outlet and new_pos:
            st.session_state.mapping[sel_tl].append({
                "outlet": new_outlet.strip(), "pos": safe_id(new_pos),
                "zmt_rk": safe_id(new_zrk) if new_zrk else None,
                "zmt_rf": safe_id(new_zrf) if new_zrf else None,
                "swg_rk": safe_id(new_srk) if new_srk else None,
                "swg_rf": safe_id(new_srf) if new_srf else None,
            })
            st.success(f"✓ {new_outlet} added under {sel_tl}")
            st.rerun()
        else:
            st.warning("Outlet Name and POS ID are required.")

# ── FOOTER ────────────────────────────────────────────────────────────────────
st.markdown("""
<div style='text-align:center;color:#333;font-size:0.75rem;padding:2rem 0 1rem;'>
    RollsKing Internal Tools · Built for Operations
</div>
""", unsafe_allow_html=True)

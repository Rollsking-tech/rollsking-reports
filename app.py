import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
from datetime import datetime

# ── PAGE CONFIG ────────────────────────────────────────────────────────────────
st.set_page_config(page_title="RollsKing Reports", page_icon="🍱", layout="wide")

# ── PASSWORD ───────────────────────────────────────────────────────────────────
APP_PASSWORD = "rollsking2025"

# ── HARDCODED MAPPING ──────────────────────────────────────────────────────────
DEFAULT_MAPPING = {
    "Navneet Singh": [
        {"outlet":"Sec 104",           "pos":28039,  "z_rk":19476740,"z_rf":20884624,"s_rk":313666, "s_rf":783919},
        {"outlet":"Sector-141 Noida",  "pos":28040,  "z_rk":20494584,"z_rf":21010012,"s_rk":330027, "s_rf":786464},
        {"outlet":"Sector-132 Noida",  "pos":28041,  "z_rk":20494614,"z_rf":21009574,"s_rk":330028, "s_rf":786465},
        {"outlet":"Sector 125 Noida",  "pos":28042,  "z_rk":20760784,"z_rf":21010014,"s_rk":442810, "s_rf":786466},
        {"outlet":"Sector-73 Noida",   "pos":28043,  "z_rk":18504972,"z_rf":20884534,"s_rk":313670, "s_rf":783921},
        {"outlet":"Sector-44 Noida",   "pos":28044,  "z_rk":20760782,"z_rf":21010016,"s_rk":442811, "s_rf":786467},
    ],
    "Ajay Halder": [
        {"outlet":"Sector 4 Noida",    "pos":28045,  "z_rk":18504964,"z_rf":20884518,"s_rk":313672, "s_rf":783923},
        {"outlet":"Sector-62",         "pos":28046,  "z_rk":20760786,"z_rf":21010018,"s_rk":442812, "s_rf":786468},
        {"outlet":"Sector-37",         "pos":32851,  "z_rk":20760788,"z_rf":21010020,"s_rk":442813, "s_rf":786469},
        {"outlet":"Sector-18",         "pos":28047,  "z_rk":19476742,"z_rf":22333370,"s_rk":42807,  "s_rf":1238358},
        {"outlet":"Gaur City GNoida",  "pos":28048,  "z_rk":20760790,"z_rf":21010022,"s_rk":442814, "s_rf":786470},
        {"outlet":"Eco Loft",          "pos":28049,  "z_rk":20760792,"z_rf":21010024,"s_rk":442815, "s_rf":786471},
    ],
    "Sunil Sharma": [
        {"outlet":"RDC Raj Nagar Gzb", "pos":28050,  "z_rk":18504970,"z_rf":20884532,"s_rk":313671, "s_rf":783922},
        {"outlet":"GNB Mall",          "pos":28051,  "z_rk":20760794,"z_rf":21010026,"s_rk":442816, "s_rf":786472},
        {"outlet":"Shipra Mall",       "pos":28052,  "z_rk":20760796,"z_rf":None,    "s_rk":1238359,"s_rf":None},
    ],
    "Vishwanath Rao": [
        {"outlet":"Indirapuram",       "pos":28053,  "z_rk":20760798,"z_rf":21010028,"s_rk":442817, "s_rf":786473},
        {"outlet":"Rajendra Nagar Gzb","pos":28054,  "z_rk":20760800,"z_rf":21010030,"s_rk":442818, "s_rf":786474},
        {"outlet":"Vasundhra",         "pos":28055,  "z_rk":20760802,"z_rf":21010032,"s_rk":442819, "s_rf":786475},
    ],
    "Sanjay Morya": [
        {"outlet":"Kalkaji",           "pos":28056,  "z_rk":18504968,"z_rf":20884530,"s_rk":313669, "s_rf":783920},
        {"outlet":"Tilak Nagar",       "pos":28057,  "z_rk":20760804,"z_rf":21010034,"s_rk":442820, "s_rf":786476},
        {"outlet":"Vasant Kunj",       "pos":28058,  "z_rk":20760806,"z_rf":21010036,"s_rk":442821, "s_rf":786477},
        {"outlet":"Chattarpur",        "pos":28059,  "z_rk":20760808,"z_rf":21010038,"s_rk":442822, "s_rf":786478},
        {"outlet":"Paschim Vihar",     "pos":28060,  "z_rk":20760810,"z_rf":21010040,"s_rk":442823, "s_rf":786479},
        {"outlet":"Gtb Nagar",         "pos":28061,  "z_rk":20760812,"z_rf":21010042,"s_rk":442824, "s_rf":786480},
        {"outlet":"Nathupur Gurugram", "pos":28062,  "z_rk":20760814,"z_rf":21010044,"s_rk":442825, "s_rf":786481},
        {"outlet":"Old DLF Sec-14 Gurgaon","pos":28063,"z_rk":20760816,"z_rf":21010046,"s_rk":442826,"s_rf":786482},
        {"outlet":"Sector-57 Gurugram","pos":28064,  "z_rk":20760818,"z_rf":21010048,"s_rk":442827, "s_rf":786483},
        {"outlet":"Wazirabad Gurugram","pos":28065,  "z_rk":20760820,"z_rf":21010050,"s_rk":442828, "s_rf":786484},
        {"outlet":"Gurugram Sec-82",   "pos":28066,  "z_rk":20760822,"z_rf":21010052,"s_rk":442829, "s_rf":786485},
        {"outlet":"Sector 90 Gurugram","pos":28067,  "z_rk":20760824,"z_rf":21010054,"s_rk":442830, "s_rf":786486},
        {"outlet":"Rohini",            "pos":28068,  "z_rk":20760826,"z_rf":21010056,"s_rk":442831, "s_rf":786487},
        {"outlet":"Vikashpuri",        "pos":28069,  "z_rk":20760828,"z_rf":21010058,"s_rk":442832, "s_rf":786488},
        {"outlet":"Uttam Nagar Dwarka","pos":414828, "z_rk":None,    "z_rf":None,    "s_rk":1280807,"s_rf":1280814},
        {"outlet":"Subhash Nagar",     "pos":28071,  "z_rk":20760832,"z_rf":21010062,"s_rk":442834, "s_rf":786490},
    ],
    "Zeeshan Ali": [
        {"outlet":"Shaheen Bagh",      "pos":28072,  "z_rk":20760834,"z_rf":21010064,"s_rk":442835, "s_rf":786491},
        {"outlet":"NIT Faridabad",     "pos":96843,  "z_rk":20760836,"z_rf":21010066,"s_rk":442836, "s_rf":786492},
        {"outlet":"Sec-15 Faridabad",  "pos":28074,  "z_rk":20760838,"z_rf":21010068,"s_rk":442837, "s_rf":786493},
        {"outlet":"Lakkarpur Faridabad","pos":28075, "z_rk":20760840,"z_rf":21010070,"s_rk":442838, "s_rf":786494},
        {"outlet":"Greenfield Faridabad","pos":28076,"z_rk":20760842,"z_rf":21010072,"s_rk":442839, "s_rf":786495},
    ],
    "Badir Alam": [
        {"outlet":"Bhopal",            "pos":28077,  "z_rk":20760844,"z_rf":21010074,"s_rk":442840, "s_rf":786496},
        {"outlet":"Indore",            "pos":28078,  "z_rk":20760846,"z_rf":21010076,"s_rk":442841, "s_rf":786497},
        {"outlet":"Siddharth Nagar Indore","pos":28079,"z_rk":20760848,"z_rf":21010078,"s_rk":442842,"s_rf":786498},
    ],
    "Abhishek Kumar": [
        {"outlet":"Whitefield",        "pos":28080,  "z_rk":20760850,"z_rf":21010080,"s_rk":442843, "s_rf":786499},
        {"outlet":"Mahadevpura",       "pos":28081,  "z_rk":20760852,"z_rf":21010082,"s_rk":442844, "s_rf":786500},
        {"outlet":"Koramangala",       "pos":28082,  "z_rk":20760854,"z_rf":21010084,"s_rk":442845, "s_rf":786501},
        {"outlet":"Electronic City",   "pos":28083,  "z_rk":20760856,"z_rf":21010086,"s_rk":442846, "s_rf":786502},
        {"outlet":"Sarjapur",          "pos":28084,  "z_rk":20760858,"z_rf":21010088,"s_rk":442847, "s_rf":786503},
        {"outlet":"Kalyan Nagar",      "pos":28085,  "z_rk":20760860,"z_rf":21010090,"s_rk":442848, "s_rf":786504},
        {"outlet":"Bel Road Bangalore","pos":28086,  "z_rk":20760862,"z_rf":21010092,"s_rk":442849, "s_rf":786505},
        {"outlet":"Habble Bangalore",  "pos":28087,  "z_rk":20760864,"z_rf":None,    "s_rk":442850, "s_rf":None},
        {"outlet":"Indira Nagar Bangalore","pos":28088,"z_rk":20760866,"z_rf":21010096,"s_rk":442851,"s_rf":786507},
    ],
    "Virendra Pratap": [
        {"outlet":"Mohanram Nagar",    "pos":28089,  "z_rk":20760868,"z_rf":21010098,"s_rk":442852, "s_rf":786508},
        {"outlet":"Madipakkam",        "pos":28090,  "z_rk":20760870,"z_rf":21010100,"s_rk":442853, "s_rf":786509},
        {"outlet":"Parengudi Chennai", "pos":28091,  "z_rk":20760872,"z_rf":21010102,"s_rk":442854, "s_rf":786510},
    ],
    "Atul Kumar": [
        {"outlet":"Apple Ghar Pune",   "pos":28092,  "z_rk":20760874,"z_rf":21010104,"s_rk":442855, "s_rf":786511},
        {"outlet":"Hinjewadi Phase 1", "pos":28093,  "z_rk":20760876,"z_rf":21010106,"s_rk":442856, "s_rf":786512},
        {"outlet":"Millennium Mall Pune","pos":28094, "z_rk":20760878,"z_rf":21010108,"s_rk":442857, "s_rf":786513},
        {"outlet":"Shivaji Nagar Pune","pos":28095,  "z_rk":20760880,"z_rf":21010110,"s_rk":442858, "s_rf":786514},
    ],
    "Bhupesh Bhatt": [
        {"outlet":"Madhapur",          "pos":28096,  "z_rk":20760882,"z_rf":21010112,"s_rk":442859, "s_rf":786515},
        {"outlet":"Gachibowli",        "pos":28097,  "z_rk":20760884,"z_rf":21010114,"s_rk":442860, "s_rf":786516},
        {"outlet":"Banjara Hills",     "pos":28098,  "z_rk":20760886,"z_rf":21010116,"s_rk":442861, "s_rf":786517},
        {"outlet":"Taranagar Hyderabad","pos":28099, "z_rk":20760888,"z_rf":21010118,"s_rk":442862, "s_rf":786518},
        {"outlet":"R K Puram Hyderabad","pos":28100, "z_rk":20760890,"z_rf":None,    "s_rk":442863, "s_rf":None},
        {"outlet":"Lulu Mall Hyderabad","pos":28101, "z_rk":20760892,"z_rf":None,    "s_rk":442864, "s_rf":None},
        {"outlet":"Miyapur",           "pos":28102,  "z_rk":20760894,"z_rf":21010124,"s_rk":442865, "s_rf":786521},
        {"outlet":"Goa Anjuna",        "pos":28103,  "z_rk":20760896,"z_rf":21010126,"s_rk":442866, "s_rf":786522},
    ],
    "Milan": [
        {"outlet":"G Corp",            "pos":28104,  "z_rk":20760898,"z_rf":21010128,"s_rk":442867, "s_rf":786523},
        {"outlet":"Mumbai Pawai",      "pos":28105,  "z_rk":20760900,"z_rf":21010130,"s_rk":442868, "s_rf":786524},
        {"outlet":"Raymond",           "pos":28106,  "z_rk":20760902,"z_rf":21010132,"s_rk":442869, "s_rf":786525},
        {"outlet":"Mumbai BKC",        "pos":28107,  "z_rk":20760904,"z_rf":21010134,"s_rk":442870, "s_rf":786526},
        {"outlet":"Mumbai Chembur",    "pos":28108,  "z_rk":20760906,"z_rf":21010136,"s_rk":442871, "s_rf":786527},
        {"outlet":"Airoli Navi Mumbai","pos":28109,  "z_rk":20760908,"z_rf":21010138,"s_rk":442872, "s_rf":786528},
        {"outlet":"Mumbai Dahisar",    "pos":399741, "z_rk":22170342,"z_rf":None,    "s_rk":1220916,"s_rf":1276496},
        {"outlet":"Mumbai Marol",      "pos":404829, "z_rk":22274879,"z_rf":None,    "s_rk":1238360,"s_rf":1276481},
        {"outlet":"Mira Road Mumbai",  "pos":28111,  "z_rk":20760912,"z_rf":21010144,"s_rk":442874, "s_rf":786530},
    ],
}

TIER_LABELS = {(0,3):"Bronze",(3,6):"Silver",(6,8):"Gold",(8,100):"Platinum"}
def get_tier(avg):
    for (lo,hi),label in TIER_LABELS.items():
        if lo <= avg < hi: return label
    return "Bronze"

def parse_pct(v):
    if v is None: return None
    if isinstance(v,str):
        v=v.strip().rstrip('%')
        try: v=float(v)
        except: return None
    if v<1: v*=100
    return round(float(v),2)

def parse_float(v):
    if v is None: return None
    try: return float(v)
    except: return None

# ── DATA LOADERS ───────────────────────────────────────────────────────────────
def get_sheet(wb, *names):
    for n in names:
        for s in wb.sheetnames:
            if n.lower() in s.lower(): return wb[s]
    return None

def sheet_to_rows(ws):
    return [[c.value for c in r] for r in ws.iter_rows()]

def load_zomato(wb):
    ws = get_sheet(wb,"zomato")
    if not ws: return {}
    rows = sheet_to_rows(ws)
    hdr_row = next((i for i,r in enumerate(rows) if any(str(c).lower()=="restaurant id" for c in r if c)),None)
    if hdr_row is None: return {}
    hdrs = [str(c).strip() if c else "" for c in rows[hdr_row]]
    def col(name):
        for i,h in enumerate(hdrs):
            if name.lower() in h.lower(): return i
        return None
    rid=col("restaurant id"); met=col("metric"); val=col("value")
    if None in (rid,met,val): return {}
    data={}
    for r in rows[hdr_row+1:]:
        if not r or not r[rid]: continue
        try: k=int(r[rid])
        except: continue
        m=str(r[met]).strip().lower() if r[met] else ""
        v=r[val]
        data.setdefault(k,{})[m]=v
    return data

def load_swiggy(wb):
    ws = get_sheet(wb,"swiggy")
    if not ws: return {}
    rows = sheet_to_rows(ws)
    hdr_row = next((i for i,r in enumerate(rows) if any(str(c).lower()=="restaurant id" for c in r if c)),None)
    if hdr_row is None: return {}
    hdrs = [str(c).strip() if c else "" for c in rows[hdr_row]]
    def col(name):
        for i,h in enumerate(hdrs):
            if name.lower() in h.lower(): return i
        return None
    rid=col("restaurant id"); met=col("metric"); val=col("value")
    if None in (rid,met,val): return {}
    data={}
    for r in rows[hdr_row+1:]:
        if not r or not r[rid]: continue
        try: k=int(r[rid])
        except: continue
        m=str(r[met]).strip().lower() if r[met] else ""
        v=r[val]
        data.setdefault(k,{})[m]=v
    return data

def load_food_cost(wb):
    ws = get_sheet(wb,"food cost","foodcost","food_cost")
    if not ws: return {}
    rows = sheet_to_rows(ws)
    hdr_row = next((i for i,r in enumerate(rows) if any(str(c).strip().lower() in ["pos id","posid","outlet id"] for c in r if c)),None)
    if hdr_row is None: return {}
    hdrs = [str(c).strip().lower() if c else "" for c in rows[hdr_row]]
    def col(*names):
        for name in names:
            for i,h in enumerate(hdrs):
                if name.lower() in h: return i
        return None
    pos_c   = col("pos id","posid","outlet id")
    open_c  = col("opening","opening balance","opening stock")
    close_c = col("closing","closing balance","closing stock")
    local_c = col("local","hyperpure","local purchase")
    store_c = col("store","store purchase")
    sale_c  = col("net sale","netsale","net_sale","total sale")
    hyg_c   = col("hygiene score","hygiene")
    if pos_c is None: return {}
    data={}
    for r in rows[hdr_row+1:]:
        if not r or not r[pos_c]: continue
        try: pos=int(r[pos_c])
        except: continue
        def gv(c): return parse_float(r[c]) if c is not None and c<len(r) else None
        opening=gv(open_c); closing=gv(close_c)
        local=gv(local_c); store=gv(store_c); sale=gv(sale_c)
        hyg = int(r[hyg_c]) if hyg_c is not None and hyg_c<len(r) and r[hyg_c] is not None else 0
        cogs = None
        if None not in (opening,closing):
            purchases = (local or 0)+(store or 0)
            cogs = opening + purchases - closing
        fc_pct = round(cogs/sale*100,2) if (cogs is not None and sale and sale>0) else None
        data[pos]={"opening":opening,"closing":closing,"local":local,"store":store,
                   "net_sale":sale,"cogs":cogs,"fc_pct":fc_pct,"hygiene":hyg}
    return data

# ── CALCULATOR ─────────────────────────────────────────────────────────────────
def calculate(mapping, zmt, swg, fc, prev_sales=None, inactive_pos=None):
    inactive_pos = inactive_pos or set()
    results=[]; flags=[]; disclaimers=[]

    for tl, outlets in mapping.items():
        active = [o for o in outlets if o["pos"] not in inactive_pos]
        if not active: continue
        n = len(active)

        sale_pts=0; fc_pts=0; cmp_pts=0; kpt_pts=0; rat_pts=0; hyg_pts=0; avail_pts=0
        outlet_rows=[]

        for o in active:
            pos=o["pos"]; out=o["outlet"]
            zrk=o["z_rk"]; zrf=o["z_rf"]; srk=o["s_rk"]; srf=o["s_rf"]
            has_rf = zrf is not None or srf is not None
            brand_label = "RK + RF" if has_rf else "RK"

            # ── SALE ──
            fc_d = fc.get(pos,{})
            cur_sale = fc_d.get("net_sale")
            prev_d = (prev_sales or {}).get(pos,{})
            prev_sale = prev_d.get("net_sale") if prev_d else None
            sp = 0
            if cur_sale and prev_sale and cur_sale > prev_sale: sp=1
            sale_pts += sp

            # ── FOOD COST ──
            fcp = fc_d.get("fc_pct")
            fp = 1 if (fcp is not None and fcp<40) else 0
            fc_pts += fp
            if fcp is None: disclaimers.append(f"{tl} | {out}: food cost missing — scored 0")

            # ── HYGIENE ──
            hyg_val = fc_d.get("hygiene",0)
            hyg_pts += hyg_val

            # ── COMPLAINTS ──
            z_comps=0; z_orders=0; s_comps=0; s_orders=0
            for zid in [zrk,zrf]:
                if not zid: continue
                zd=zmt.get(zid,{})
                zo=parse_float(zd.get("delivered orders"))
                zc=parse_float(zd.get("total complaints"))
                if zo: z_orders+=zo
                if zc: z_comps+=zc
            for sid in [srk,srf]:
                if not sid: continue
                sd=swg.get(sid,{})
                so=parse_float(sd.get("delivered orders"))
                sc_pct=parse_pct(sd.get("% orders with complaints"))
                if so and sc_pct is not None:
                    s_orders+=so; s_comps+=so*sc_pct/100
            tot_orders = z_orders+s_orders
            tot_comps  = z_comps+s_comps
            cmp_pct = round(tot_comps/tot_orders*100,2) if tot_orders>0 else None
            if cmp_pct is None: cp=0
            elif cmp_pct<1:  cp=4
            elif cmp_pct<2:  cp=3
            elif cmp_pct<3:  cp=1
            else:            cp=0
            cmp_pts += cp

            # ── KPT ── Zomato RK + Swiggy RK Avg Prep Time only
            kpt_vals=[]
            zd_rk=zmt.get(zrk,{})
            zk=parse_float(zd_rk.get("kitchen preparation time (in minutes)") or
                           zd_rk.get("kitchen preparation time") or
                           zd_rk.get("kpt"))
            if zk and zk>0: kpt_vals.append(zk)
            sd_rk=swg.get(srk,{})
            sk=parse_float(sd_rk.get("avg prep time") or sd_rk.get("average prep time"))
            if sk and sk>0: kpt_vals.append(sk)
            if not kpt_vals:
                kp=0
                disclaimers.append(f"{tl} | {out}: KPT missing — scored 0")
            else:
                avg_kpt=sum(kpt_vals)/len(kpt_vals)
                kp = 1 if int(avg_kpt)<=12 else 0
                kpt_pts += kp

            # ── RATING ── Zomato only (all available IDs)
            rat_vals=[]
            for zid in [zrk,zrf]:
                if not zid: continue
                zd=zmt.get(zid,{})
                rv=parse_float(zd.get("average rating") or zd.get("rating"))
                if rv: rat_vals.append(rv)
            if not rat_vals:
                rp=0
                disclaimers.append(f"{tl} | {out}: rating missing — scored 0")
            else:
                avg_rat=sum(rat_vals)/len(rat_vals)
                rp = 1 if avg_rat>=4.0 else 0
                rat_pts += rp

            # ── AVAILABILITY ── avg of Zomato + Swiggy for all available IDs
            avail_vals=[]
            for zid in [zrk,zrf]:
                if not zid: continue
                zd=zmt.get(zid,{})
                av=parse_pct(zd.get("online %") or zd.get("online percentage") or zd.get("availability"))
                if av is not None: avail_vals.append(av)
            for sid in [srk,srf]:
                if not sid: continue
                sd=swg.get(sid,{})
                av=parse_pct(sd.get("online availability %") or sd.get("online availability"))
                if av is not None: avail_vals.append(av)
            if not avail_vals:
                avp=0
            else:
                avg_avail=sum(avail_vals)/len(avail_vals)
                avp = 1 if avg_avail>=98 else 0
                avail_pts += avp

            # ── FLAGS ──
            if cmp_pct and cmp_pct>=3:  flags.append((tl,out,"High Complaints",f"{cmp_pct:.1f}%",">=3%"))
            if fcp and fcp>=40:         flags.append((tl,out,"High Food Cost",f"{fcp:.1f}%",">=40%"))
            if kpt_vals and int(sum(kpt_vals)/len(kpt_vals))>12:
                flags.append((tl,out,"KPT Exceeded",f"{sum(kpt_vals)/len(kpt_vals):.1f} min","floor>12"))
            if rat_vals and sum(rat_vals)/len(rat_vals)<4.0:
                flags.append((tl,out,"Low Rating",f"{sum(rat_vals)/len(rat_vals):.2f}","<4.0"))
            if avail_vals and sum(avail_vals)/len(avail_vals)<98:
                flags.append((tl,out,"Low Availability",f"{sum(avail_vals)/len(avail_vals):.1f}%","<98%"))
            if sp==0 and cur_sale and prev_sale:
                flags.append((tl,out,"Sales Decline",f"■{cur_sale:,.0f}",f"Prev ■{prev_sale:,.0f}"))

            # ── OUTLET ROW ──
            outlet_rows.append({
                "outlet":out, "brand":brand_label,
                "net_sale":cur_sale, "prev_sale":prev_sale,
                "fc_pct":fcp,
                "cmp_pct":cmp_pct, "tot_comps":round(tot_comps),
                "kpt_avg":round(sum(kpt_vals)/len(kpt_vals),1) if kpt_vals else None,
                "rating":round(sum(rat_vals)/len(rat_vals),2) if rat_vals else None,
                "avail":round(sum(avail_vals)/len(avail_vals),1) if avail_vals else None,
                "hygiene":hyg_val,
            })

        total_pts = sale_pts+fc_pts+cmp_pts+kpt_pts+rat_pts+hyg_pts+avail_pts
        avg = round(total_pts/n,2) if n>0 else 0
        results.append({
            "tl":tl,"n":n,
            "sale":sale_pts,"fc":fc_pts,"cmp":cmp_pts,
            "kpt":kpt_pts,"rat":rat_pts,"hyg":hyg_pts,"avail":avail_pts,
            "total":total_pts,"avg":avg,"tier":get_tier(avg),
            "outlets":outlet_rows,
        })

    results.sort(key=lambda x: x["avg"], reverse=True)
    return results, list(set(disclaimers)), flags

# ── EXCEL BUILDER (light colours, no points columns) ──────────────────────────
def build_excel(results, disclaimers, flags, month):
    wb = openpyxl.Workbook()

    # ── COLOURS (light theme) ──
    GOLD_FILL   = PatternFill("solid", fgColor="FFF3CD")
    SILVER_FILL = PatternFill("solid", fgColor="F0F0F0")
    BRONZE_FILL = PatternFill("solid", fgColor="FFE8D0")
    PLAT_FILL   = PatternFill("solid", fgColor="D4EDDA")
    HDR_FILL    = PatternFill("solid", fgColor="2C3E50")
    SUBHDR_FILL = PatternFill("solid", fgColor="4A6274")
    RED_FILL    = PatternFill("solid", fgColor="FFCCCC")
    GRN_FILL    = PatternFill("solid", fgColor="CCFFCC")
    YLW_FILL    = PatternFill("solid", fgColor="FFF9CC")
    WHITE_FILL  = PatternFill("solid", fgColor="FFFFFF")
    ALT_FILL    = PatternFill("solid", fgColor="F8F9FA")

    TIER_FILL = {"Gold":GOLD_FILL,"Silver":SILVER_FILL,"Bronze":BRONZE_FILL,"Platinum":PLAT_FILL}

    def hdr_cell(ws, row, col, val, bold=True, color="FFFFFF", fill=None, align="center", wrap=False, size=10):
        c=ws.cell(row=row,column=col,value=val)
        c.font=Font(bold=bold,color=color,size=size)
        c.alignment=Alignment(horizontal=align,vertical="center",wrap_text=wrap)
        if fill: c.fill=fill
        return c

    def data_cell(ws, row, col, val, fill=None, bold=False, align="center", fmt=None, color="000000"):
        c=ws.cell(row=row,column=col,value=val)
        c.font=Font(bold=bold,color=color)
        c.alignment=Alignment(horizontal=align,vertical="center")
        if fill: c.fill=fill
        if fmt: c.number_format=fmt
        return c

    thin = Side(style="thin",color="CCCCCC")
    def border(c): c.border=Border(left=thin,right=thin,top=thin,bottom=thin); return c

    # ── SHEET 1: TL PERFORMANCE SUMMARY ──────────────────────────────────────
    ws1 = wb.active
    ws1.title = f"{month} Performance"
    ws1.freeze_panes = "A4"
    ws1.sheet_view.showGridLines = False

    # Title
    ws1.merge_cells("A1:L1")
    c=ws1.cell(row=1,column=1,value=f"RollsKing — TL Performance Report | {month}")
    c.font=Font(bold=True,size=13,color="2C3E50"); c.alignment=Alignment(horizontal="center",vertical="center")
    ws1.row_dimensions[1].height=28

    ws1.merge_cells("A2:L2")
    c=ws1.cell(row=2,column=1,value=f"Generated: {datetime.now().strftime('%d %b %Y, %H:%M')}  |  {len(results)} Team Leaders  |  {sum(r['n'] for r in results)} Outlets")
    c.font=Font(size=9,color="777777"); c.alignment=Alignment(horizontal="center")
    ws1.row_dimensions[2].height=16

    # Headers — NOTE: no points columns, hygiene kept
    hdrs=[("Team Leader",18),("Outlets",8),("Sale Pts",8),("FC Pts",7),
          ("Cmp Pts",7),("KPT Pts",7),("Rating Pts",8),("Hygiene",8),
          ("Avail Pts",8),("Total Pts",8),("Avg Score",9),("Tier",10)]
    ws1.row_dimensions[3].height=22
    for ci,(h,w) in enumerate(hdrs,1):
        c=hdr_cell(ws1,3,ci,h,fill=HDR_FILL,color="FFFFFF",size=9)
        border(c)
        ws1.column_dimensions[get_column_letter(ci)].width=w

    for ri,r in enumerate(results,4):
        tf=TIER_FILL.get(r["tier"],WHITE_FILL)
        row_fill = tf
        vals=[r["tl"],r["n"],r["sale"],r["fc"],r["cmp"],
              r["kpt"],r["rat"],r["hyg"],r["avail"],r["total"],r["avg"],r["tier"]]
        for ci,v in enumerate(vals,1):
            al = "left" if ci==1 else "center"
            bld = ci==1
            c=data_cell(ws1,ri,ci,v,fill=row_fill if ci!=1 else WHITE_FILL,bold=bld,align=al)
            if ci==12:  # Tier
                c.font=Font(bold=True,color="2C3E50")
            border(c)
        ws1.row_dimensions[ri].height=18

    # Grand Total row
    gr = len(results)+4
    ws1.merge_cells(f"A{gr}:B{gr}")
    c=ws1.cell(row=gr,column=1,value="GRAND TOTAL"); c.font=Font(bold=True,size=10,color="FFFFFF")
    c.fill=HDR_FILL; c.alignment=Alignment(horizontal="center",vertical="center")
    cols_to_sum=[3,4,5,6,7,8,9,10]
    for ci in range(1,13):
        if ci in cols_to_sum:
            v=sum(r[["sale","fc","cmp","kpt","rat","hyg","avail","total"][cols_to_sum.index(ci)-3]] for r in results)
            c=data_cell(ws1,gr,ci,v,fill=HDR_FILL,bold=True,color="FFFFFF")
        elif ci not in (1,2):
            c=ws1.cell(row=gr,column=ci,value=""); c.fill=HDR_FILL
        border(ws1.cell(row=gr,column=ci))
    ws1.row_dimensions[gr].height=20

    # Scoring reference below
    ref_row = gr+2
    ws1.cell(row=ref_row,column=1,value="SCORING CONDITIONS").font=Font(bold=True,size=9,color="2C3E50")
    conditions=[
        ("Sale","1pt per outlet beating prev month Net Sale — summed per TL"),
        ("Food Cost","<40% = 1pt | >=40% = 0pt"),
        ("Complaints","0-<1%=4pts | 1-<2%=3pts | 2-<3%=1pt | >=3%=0pt"),
        ("KPT","floor(avg)<= 12min = 1pt | >12min = 0pt (Zomato RK + Swiggy RK)"),
        ("Rating",">=4.0 = 1pt | <4.0 = 0pt (Zomato only)"),
        ("Hygiene","Sum of Hygiene Score column per TL (from food cost sheet)"),
        ("Availability",">=98% = 1pt | <98% = 0pt (avg Zomato + Swiggy)"),
        ("Grade","Bronze 0-3 | Silver 3-6 | Gold 6-8 | Platinum 8+"),
    ]
    for i,(k,v) in enumerate(conditions):
        rr=ref_row+1+i
        ws1.cell(row=rr,column=1,value=k).font=Font(bold=True,size=8,color="2C3E50")
        ws1.cell(row=rr,column=2,value=v).font=Font(size=8,color="444444")
        ws1.merge_cells(f"B{rr}:L{rr}")

    # ── SHEET 2: OUTLET DETAIL (no points, raw values + brand column) ─────────
    ws2 = wb.create_sheet("Outlet Detail")
    ws2.freeze_panes="C3"; ws2.sheet_view.showGridLines=False

    ws2.merge_cells("A1:M1")
    c=ws2.cell(row=1,column=1,value=f"Outlet Detail — {month}")
    c.font=Font(bold=True,size=12,color="2C3E50"); c.alignment=Alignment(horizontal="center",vertical="center")
    ws2.row_dimensions[1].height=24

    # Headers — raw values only, no points. Brand column added
    o_hdrs=[("TL",18),("Outlet",22),("Brands",10),
            ("Net Sale (₹)",13),("Prev Sale (₹)",13),("Food Cost %",11),
            ("Complaints %",12),("Total Comps",11),
            ("KPT (min)",10),("Rating",8),("Availability %",13),("Hygiene",8)]
    for ci,(h,w) in enumerate(o_hdrs,1):
        c=hdr_cell(ws2,2,ci,h,fill=SUBHDR_FILL,color="FFFFFF",size=9,wrap=True)
        border(c); ws2.column_dimensions[get_column_letter(ci)].width=w
    ws2.row_dimensions[2].height=30

    row_i=3
    for r in results:
        for idx,o in enumerate(r["outlets"]):
            alt = ALT_FILL if idx%2==0 else WHITE_FILL
            vals=[
                r["tl"] if idx==0 else "",
                o["outlet"],
                o["brand"],
                o["net_sale"],
                o["prev_sale"],
                o["fc_pct"],
                o["cmp_pct"],
                o["tot_comps"],
                o["kpt_avg"],
                o["rating"],
                o["avail"],
                o["hygiene"],
            ]
            for ci,v in enumerate(vals,1):
                al="left" if ci<=2 else "center"
                # colour code key metrics
                cell_fill=alt
                if ci==6 and v is not None:  # FC
                    cell_fill=RED_FILL if v>=40 else GRN_FILL
                if ci==7 and v is not None:  # Cmp%
                    cell_fill=RED_FILL if v>=3 else (YLW_FILL if v>=2 else GRN_FILL)
                if ci==9 and v is not None:  # KPT
                    cell_fill=RED_FILL if int(v)>12 else GRN_FILL
                if ci==10 and v is not None: # Rating
                    cell_fill=RED_FILL if v<4 else GRN_FILL
                if ci==11 and v is not None: # Avail
                    cell_fill=RED_FILL if v<98 else GRN_FILL
                # format numbers
                fmt=None
                if ci in (4,5) and v: fmt='#,##0'
                if ci in (6,7,11) and v: fmt='0.00"%"'
                c=data_cell(ws2,row_i,ci,v,fill=cell_fill,align=al,fmt=fmt)
                if ci<=2 and idx==0: c.font=Font(bold=True,size=9)
                border(c)
            ws2.row_dimensions[row_i].height=16
            row_i+=1

    # ── SHEET 3: FLAGGED OUTLETS ──────────────────────────────────────────────
    ws3=wb.create_sheet("Flagged Outlets")
    ws3.sheet_view.showGridLines=False
    ws3.merge_cells("A1:E1")
    c=ws3.cell(row=1,column=1,value=f"Flagged Outlets — {month}  ({len(flags)} total)")
    c.font=Font(bold=True,size=12,color="2C3E50"); c.alignment=Alignment(horizontal="center",vertical="center")
    ws3.row_dimensions[1].height=24
    for ci,(h,w) in enumerate([("Team Leader",20),("Outlet",22),("Issue",22),("Value",12),("Threshold",12)],1):
        c=hdr_cell(ws3,2,ci,h,fill=HDR_FILL,color="FFFFFF",size=9)
        border(c); ws3.column_dimensions[get_column_letter(ci)].width=w
    for ri,(tl,out,issue,val,thr) in enumerate(flags,3):
        for ci,v in enumerate([tl,out,issue,val,thr],1):
            c=data_cell(ws3,ri,ci,v,fill=RED_FILL if "High" in issue or "Exceeded" in issue or "Decline" in issue else YLW_FILL,align="left" if ci<=2 else "center")
            border(c)
        ws3.row_dimensions[ri].height=16

    # ── SHEET 4: DATA NOTES ───────────────────────────────────────────────────
    if disclaimers:
        ws4=wb.create_sheet("Data Notes")
        ws4.cell(row=1,column=1,value="Data Notes").font=Font(bold=True,size=11,color="2C3E50")
        for i,d in enumerate(disclaimers,2):
            c=ws4.cell(row=i,column=1,value=f"• {d}")
            c.font=Font(size=9,color="555555")
            ws4.column_dimensions["A"].width=70

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.read()

# ── PDF BUILDER (2 sections only: summary table + pie chart) ──────────────────
def build_pdf(results, flags, disclaimers, month):
    try:
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        import matplotlib.patches as mpatches
        import numpy as np
        plt.close('all')
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib import colors
        from reportlab.lib.units import mm
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph, Image, HRFlowable
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.enums import TA_CENTER, TA_LEFT
    except Exception as e:
        return None, str(e)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4),
                            topMargin=12*mm, bottomMargin=12*mm,
                            leftMargin=14*mm, rightMargin=14*mm)

    styles = getSampleStyleSheet()
    DARK    = colors.HexColor("#2C3E50")
    GOLD_C  = colors.HexColor("#F39C12")
    SILVER_C= colors.HexColor("#95A5A6")
    BRONZE_C= colors.HexColor("#CD7F32")
    PLAT_C  = colors.HexColor("#27AE60")
    LIGHT   = colors.HexColor("#F8F9FA")
    RED_C   = colors.HexColor("#E74C3C")
    GRN_C   = colors.HexColor("#27AE60")
    YLW_C   = colors.HexColor("#F39C12")

    TIER_COLORS = {"Gold":GOLD_C,"Silver":SILVER_C,"Bronze":BRONZE_C,"Platinum":PLAT_C}

    title_style = ParagraphStyle("title",fontName="Helvetica-Bold",fontSize=16,textColor=DARK,
                                  spaceAfter=4,alignment=TA_CENTER)
    sub_style   = ParagraphStyle("sub",fontName="Helvetica",fontSize=9,textColor=colors.HexColor("#777777"),
                                  spaceAfter=10,alignment=TA_CENTER)
    sec_style   = ParagraphStyle("sec",fontName="Helvetica-Bold",fontSize=11,textColor=DARK,
                                  spaceBefore=10,spaceAfter=6)

    story=[]

    # ── TITLE ──
    story.append(Paragraph(f"RollsKing — Monthly Performance Report", title_style))
    story.append(Paragraph(
        f"{month}  |  Generated {datetime.now().strftime('%d %b %Y, %H:%M')}  |  "
        f"{len(results)} Team Leaders  |  {sum(r['n'] for r in results)} Outlets",
        sub_style))
    story.append(HRFlowable(width="100%",thickness=1,color=DARK,spaceAfter=8))

    # ── SECTION 1: PERFORMANCE SUMMARY TABLE ──
    story.append(Paragraph("Performance Summary", sec_style))

    tbl_data=[["Team Leader","Outlets","Sale","FC","Cmp","KPT","Rating","Hygiene","Avail","Total","Avg","Tier"]]
    for r in results:
        tbl_data.append([
            r["tl"], r["n"],
            r["sale"], r["fc"], r["cmp"],
            r["kpt"], r["rat"], r["hyg"], r["avail"],
            r["total"], f"{r['avg']:.2f}", r["tier"]
        ])

    col_widths=[62*mm,16*mm,14*mm,12*mm,12*mm,12*mm,16*mm,16*mm,13*mm,14*mm,14*mm,18*mm]

    tbl_style=[
        ("BACKGROUND",   (0,0),(-1,0), DARK),
        ("TEXTCOLOR",    (0,0),(-1,0), colors.white),
        ("FONTNAME",     (0,0),(-1,0), "Helvetica-Bold"),
        ("FONTSIZE",     (0,0),(-1,0), 8),
        ("ALIGN",        (0,0),(-1,-1),"CENTER"),
        ("ALIGN",        (0,0),(0,-1), "LEFT"),
        ("FONTSIZE",     (0,1),(-1,-1),8),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("#F8F9FA")]),
        ("GRID",         (0,0),(-1,-1),0.4,colors.HexColor("#DDDDDD")),
        ("LEFTPADDING",  (0,0),(-1,-1),4),
        ("RIGHTPADDING", (0,0),(-1,-1),4),
        ("TOPPADDING",   (0,0),(-1,-1),4),
        ("BOTTOMPADDING",(0,0),(-1,-1),4),
        ("FONTNAME",     (0,1),(0,-1),"Helvetica-Bold"),
    ]
    # Colour tier column
    for ri,r in enumerate(results,1):
        tc=TIER_COLORS.get(r["tier"],SILVER_C)
        tbl_style.append(("BACKGROUND",(11,ri),(11,ri),tc))
        tbl_style.append(("TEXTCOLOR",(11,ri),(11,ri),colors.white))
        tbl_style.append(("FONTNAME",(11,ri),(11,ri),"Helvetica-Bold"))

    tbl=Table(tbl_data,colWidths=col_widths)
    tbl.setStyle(TableStyle(tbl_style))
    story.append(tbl)
    story.append(Spacer(1,8*mm))
    story.append(HRFlowable(width="100%",thickness=0.5,color=colors.HexColor("#DDDDDD"),spaceAfter=6))

    # ── SECTION 2: TIER DISTRIBUTION PIE CHART ──
    story.append(Paragraph("Grade Distribution", sec_style))

    tier_counts={"Platinum":0,"Gold":0,"Silver":0,"Bronze":0}
    for r in results:
        tier_counts[r["tier"]] = tier_counts.get(r["tier"],0)+1

    active_tiers = {k:v for k,v in tier_counts.items() if v>0}
    pie_colors_map={"Platinum":"#27AE60","Gold":"#F39C12","Silver":"#95A5A6","Bronze":"#CD7F32"}

    fig,ax=plt.subplots(figsize=(4,3.2),facecolor='white')
    wedge_colors=[pie_colors_map[k] for k in active_tiers]
    wedges,texts,autotexts=ax.pie(
        active_tiers.values(),
        labels=[f"{k}\n({v})" for k,v in active_tiers.items()],
        colors=wedge_colors,
        autopct='%1.0f%%',
        startangle=140,
        pctdistance=0.75,
        textprops={'fontsize':9,'color':'#2C3E50'},
        wedgeprops={'edgecolor':'white','linewidth':2}
    )
    for at in autotexts: at.set_color('white'); at.set_fontweight('bold'); at.set_fontsize(9)
    ax.set_title("TL Tier Breakdown", fontsize=10, color="#2C3E50", fontweight='bold', pad=8)
    plt.tight_layout()

    img_buf=io.BytesIO()
    fig.savefig(img_buf,format='png',dpi=150,bbox_inches='tight',facecolor='white')
    plt.close(fig)
    img_buf.seek(0)
    story.append(Image(img_buf, width=100*mm, height=80*mm))

    doc.build(story)
    buf.seek(0)
    return buf.read(), None

# ── SESSION STATE ──────────────────────────────────────────────────────────────
for k,v in [("logged_in",False),("mapping",DEFAULT_MAPPING),
            ("inactive_pos",set()),("results",None),
            ("excel_bytes",None),("pdf_bytes",None),("pdf_error",None)]:
    if k not in st.session_state: st.session_state[k]=v

# ── LOGIN ──────────────────────────────────────────────────────────────────────
if not st.session_state.logged_in:
    st.markdown("<h2 style='text-align:center;color:#2C3E50;margin-top:80px'>🍱 RollsKing Reports</h2>",
                unsafe_allow_html=True)
    col1,col2,col3=st.columns([1,1.2,1])
    with col2:
        pwd=st.text_input("Password",type="password",placeholder="Enter password")
        if st.button("Enter",use_container_width=True):
            if pwd==APP_PASSWORD: st.session_state.logged_in=True; st.rerun()
            else: st.error("Incorrect password")
    st.stop()

# ── MAIN APP ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .block-container{padding-top:1.5rem}
    .stTabs [data-baseweb="tab"]{font-size:14px;font-weight:600}
    div[data-testid="stMetric"]{background:#F8F9FA;border-radius:8px;padding:12px;border:1px solid #E0E0E0}
    .stButton>button{background:#2C3E50;color:white;border-radius:6px;border:none;font-weight:600}
    .stButton>button:hover{background:#34495E}
    .stDownloadButton>button{background:#27AE60;color:white;border-radius:6px;border:none;font-weight:600}
</style>""", unsafe_allow_html=True)

st.markdown("<h2 style='color:#2C3E50;margin-bottom:0'>🍱 RollsKing Monthly Report Generator</h2>",
            unsafe_allow_html=True)
st.caption(f"Logged in  |  {len(st.session_state.mapping)} Team Leaders  |  "
           f"{sum(len(v) for v in st.session_state.mapping.values())} Outlets")

tab1, tab2 = st.tabs(["📊 Generate Report", "⚙️ Outlet Status"])

# ── TAB 1: GENERATE ────────────────────────────────────────────────────────────
with tab1:
    st.markdown("---")
    # STEP 1
    st.markdown("### Step 1 — Select Month")
    c1,c2,_=st.columns([1,1,3])
    months=["January","February","March","April","May","June",
            "July","August","September","October","November","December"]
    sel_month=c1.selectbox("Month",months,index=datetime.now().month-1)
    sel_year=c2.number_input("Year",min_value=2024,max_value=2030,value=datetime.now().year)
    month_label=f"{sel_month} {sel_year}"

    st.markdown("---")
    # STEP 2
    st.markdown("### Step 2 — Upload Files")
    st.caption("Upload your monthly raw data file and the Swiggy Business Metrics file. Drop both here.")
    uf1,uf2=st.columns(2)
    with uf1:
        st.markdown("**Monthly Raw Data File**")
        st.caption("Contains Zomato, Swiggy, Food Cost sheets")
        main_file=st.file_uploader("Upload monthly file",type=["xlsx","xls"],key="main_file",label_visibility="collapsed")
    with uf2:
        st.markdown("**Previous Month File** *(for Sale comparison)*")
        st.caption("Same format — used to calculate month-on-month growth")
        prev_file=st.file_uploader("Upload previous month file",type=["xlsx","xls"],key="prev_file",label_visibility="collapsed")

    st.markdown("---")
    # STEP 3: HYGIENE NOTE
    st.markdown("### Step 3 — Hygiene Scores")
    st.info("Hygiene scores are read automatically from the **Hygiene Score** column in the Food Cost sheet. "
            "Ensure Sujeet has added this column (0 or 1 per outlet) before uploading.", icon="ℹ️")

    st.markdown("---")
    # STEP 4: GENERATE
    st.markdown("### Step 4 — Generate")
    if st.button("✅ Generate Report", use_container_width=False):
        if not main_file:
            st.error("Please upload the monthly raw data file first.")
        else:
            with st.spinner("Processing data..."):
                try:
                    wb_main=openpyxl.load_workbook(main_file,data_only=True)
                    wb_prev=openpyxl.load_workbook(prev_file,data_only=True) if prev_file else None
                    zmt  = load_zomato(wb_main)
                    swg  = load_swiggy(wb_main)
                    fc   = load_food_cost(wb_main)
                    prev = load_food_cost(wb_prev) if wb_prev else None

                    results,disclaimers,flags = calculate(
                        st.session_state.mapping, zmt, swg, fc,
                        prev_sales=prev,
                        inactive_pos=st.session_state.inactive_pos
                    )
                    st.session_state.results=results

                    excel=build_excel(results,disclaimers,flags,month_label)
                    st.session_state.excel_bytes=excel

                    pdf,pdf_err=build_pdf(results,flags,disclaimers,month_label)
                    st.session_state.pdf_bytes=pdf
                    st.session_state.pdf_error=pdf_err

                    st.success(f"✅ Report generated — {len(results)} TLs | {sum(r['n'] for r in results)} outlets | {len(flags)} flags")
                    if disclaimers:
                        with st.expander(f"⚠️ {len(disclaimers)} data notes"):
                            for d in disclaimers: st.caption(f"• {d}")
                except Exception as e:
                    st.error(f"Error processing files: {e}")
                    import traceback; st.code(traceback.format_exc())

    # DOWNLOADS
    if st.session_state.excel_bytes or st.session_state.pdf_bytes:
        st.markdown("### Download Reports")
        dc1,dc2=st.columns(2)
        with dc1:
            if st.session_state.excel_bytes:
                st.download_button("⬇️ Download Excel Report",
                    data=st.session_state.excel_bytes,
                    file_name=f"RollsKing_Report_{month_label.replace(' ','_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)
        with dc2:
            if st.session_state.pdf_bytes:
                st.download_button("⬇️ Download PDF Summary",
                    data=st.session_state.pdf_bytes,
                    file_name=f"RollsKing_Report_{month_label.replace(' ','_')}.pdf",
                    mime="application/pdf",
                    use_container_width=True)
            elif st.session_state.pdf_error:
                st.warning(f"PDF unavailable: {st.session_state.pdf_error}")

    # PREVIEW
    if st.session_state.results:
        st.markdown("---")
        st.markdown("### Preview")
        rows=[]
        for r in st.session_state.results:
            rows.append({"Team Leader":r["tl"],"Outlets":r["n"],
                         "Sale":r["sale"],"FC":r["fc"],"Cmp":r["cmp"],
                         "KPT":r["kpt"],"Rating":r["rat"],"Hygiene":r["hyg"],
                         "Avail":r["avail"],"Total":r["total"],"Avg":r["avg"],"Tier":r["tier"]})
        st.dataframe(pd.DataFrame(rows),use_container_width=True,hide_index=True)

# ── TAB 2: OUTLET STATUS ───────────────────────────────────────────────────────
with tab2:
    st.markdown("### Outlet Status")
    st.caption("Mark outlets as inactive for this month. Inactive outlets are excluded from scoring and counts.")
    st.info("Changes apply to the current session only. They reset when you reload the page.", icon="ℹ️")

    for tl, outlets in st.session_state.mapping.items():
        with st.expander(f"**{tl}** — {len(outlets)} outlets"):
            for o in outlets:
                pos=o["pos"]; out=o["outlet"]
                is_inactive = pos in st.session_state.inactive_pos
                col_a,col_b=st.columns([3,1])
                col_a.write(out)
                status="🔴 Inactive" if is_inactive else "🟢 Active"
                if col_b.button(status,key=f"toggle_{pos}"):
                    if is_inactive: st.session_state.inactive_pos.discard(pos)
                    else: st.session_state.inactive_pos.add(pos)
                    st.rerun()

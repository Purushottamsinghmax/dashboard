#!/usr/bin/env python3
"""
Net Metering Intelligence dashboard generator  (DVVNL)
------------------------------------------------------
Reads   DV_NET_METER_TAGGED.csv  +  net_meter_billing_UPDATED.xlsx
        (sheets MD03_MASTER, MATCHED_SMART)
Writes  MCO_NET_METER_DASHBOARD_STANDALONE.html  (single offline file, no internet needed)

Smart  = bill cycle MD03, everything else = non-smart. Wrong tagging uses the REMARK column only.
Tabs   1. Billing & Exceptions   (smart net meter billing status, exception breakdown)
       2. Wrong Tagging          (Net Meter / Net Bill configured vs eligible)
Smart  dropdowns (search + counts + cascading), click-to-filter charts, drill-down
       Zone > Circle > Division > SDO > Consumer, consumer drawer, CSV export.

Needs: DVVNL_SubstationWise_HKEY-2.xlsx (hierarchy) in the same folder.
Run:   python net_meter_dashboard.py [input_folder] [output_folder]   (both default to the script folder)
"""
import base64, json, sys
from pathlib import Path
import numpy as np
import pandas as pd

BASE = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parent
CSV, XLSX = BASE / "DV_NET_METER_TAGGED.csv", BASE / "net_meter_billing_UPDATED.xlsx"
OUT = (Path(sys.argv[2]) if len(sys.argv) > 2 else BASE) / "MCO_NET_METER_DASHBOARD_STANDALONE.html"
DETAIL_SCOPE = "attention"   # "attention" = full consumer details only for unbilled / exception / wrong-tagged rows
                             # "all"       = details for every row (HTML becomes ~25 MB)
GMR = {"AGRA I", "AGRA II", "ALIGARH", "ETAH", "FIROZABAD"}      # everything else = Genus
STATUS = {"G": "G - Multiple SA", "O": "O - Off-cycle bill", "NET METER": "Net meter",
          "METER BADGE/SERIAL NO MIS-MATCH": "Meter badge mismatch", "ERROR": "Error",
          "LESS READ": "Less read", "PENDING": "Pending", "COMPLETED": "Completed"}

# ---------------------------------------------------------------- load & derive
d = pd.read_csv(CSV, low_memory=False)
matched = pd.read_excel(XLSX, sheet_name="MATCHED_SMART")
HX = BASE / "DVVNL_SubstationWise_HKEY-2.xlsx"      # official Zone > Circle > Division > SDO hierarchy
nm = lambda v: " ".join(str(v).split()).upper()
ALIAS = {"EDSD-1 LALITPUR": "EDSD I LALITPUR", "EDSD-IV TALBEHAT": "2336318-EDSD IV TALBEHAT", "EDSD-BANSI": "2336317-EDSD BANSI"}
h = pd.read_excel(HX, dtype=str)
for c in ("ZONE_NAME", "CIRCLE_NAME", "DIV_NAME", "SDO_NAME"):
    h[c] = h[c].map(nm)
hs = h.drop_duplicates("SDO_CODE").set_index("SDO_NAME")
d["SDO_NAME"] = d["SDO_NAME"].map(nm).replace(ALIAS)
print("SDO names not found in hierarchy file:", int((~d["SDO_NAME"].isin(hs.index)).sum()))
for c in ("ZONE_NAME", "CIRCLE_NAME", "DIV_NAME"):      # parents always come from the Excel hierarchy
    d[c] = d["SDO_NAME"].map(hs[c]).fillna(d[c].map(nm))
d["SUBSTATION"] = d["SUBSTATION"].fillna("-").astype(str).str.strip().str.upper()
tree = [[r.ZONE_NAME, r.CIRCLE_NAME, r.DIV_NAME, r.SDO_NAME, r.ZONE_CODE, r.CIRCLE_CODE, r.DIV_CODE, r.SDO_CODE]
        for r in h.drop_duplicates("SDO_CODE").sort_values(["ZONE_NAME", "CIRCLE_NAME", "DIV_NAME", "SDO_NAME"]).itertuples()]

stat = matched.drop_duplicates("ACCT_ID").set_index("ACCT_ID")["BILL_STATUS_FLG"].map(STATUS)
smart = d["BILL_CYC_CD"].astype(str).str.strip().str.upper().eq("MD03")      # MD03 = smart, every other cycle = non-smart
billed = d["FLG"].astype(str).str.strip().str.upper().eq("BILLED")
d["_stat"] = d["ACCT_ID"].map(stat)
d["SMART"] = np.where(smart, "Smart (MD03)", "Non-smart")
d["BSTAT"] = np.select([~smart & billed, ~smart, billed, d["_stat"].notna()],
                       ["Non-smart billed", "Non-smart unbilled", "Billed", d["_stat"]], default="Unbilled - no status")
d["GRP"] = np.where(d["ZONE_NAME"].isin(GMR), "GMR", "Genus")

DIMS = [("ZONE", "ZONE_NAME"), ("CIRCLE", "CIRCLE_NAME"), ("DIV", "DIV_NAME"), ("SDO", "SDO_NAME"), ("SS", "SUBSTATION"),
        ("GRP", "GRP"), ("SMT", "SMART"), ("CYC", "BILL_CYC_CD"), ("TAR", "TARIFF_TYPE"), ("CFG", "MTR_CONFIG_TY_CD"),
        ("CTAG", "CONFIGURED_TAG"), ("ETAG", "ELIGIBLE_TAG"), ("REM", "REMARK"), ("BST", "BSTAT"),
        ("CON", "CON_STATUS"), ("MAKE", "MTR_MAKE")]

def b64(a): return base64.b64encode(np.ascontiguousarray(a).tobytes()).decode()

cols = {}
for key, col in DIMS:
    cat = pd.Categorical(d[col].fillna("-").astype(str))
    cols[key] = {"v": [str(x) for x in cat.categories], "d": b64(cat.codes.astype("<u2"))}

attention = (d["REMARK"] != "Correctly Tagged") | (~d["BSTAT"].isin(["Billed", "Non-smart billed"]))
sel = d.index if DETAIL_SCOPE == "all" else d.index[attention]
clean = lambda v: "" if pd.isna(v) else str(v)
det = {}
for i in sel:
    r = d.loc[i]
    det[int(i)] = [clean(r.NAME), "" if pd.isna(r.MOBILE_NO) else str(int(r.MOBILE_NO)), clean(r.SERIAL_NBR),
                   clean(r.FEEDER), clean(r.SUBSTATION), clean(r.ADDRESS), clean(r.LAST_BILL_DATE),
                   clean(r.PAY_DATE), clean(r.CURRENT_ASSESSMENT), clean(r.LOAD) + " " + clean(r.LOAD_UNIT)]

data = {"n": int(len(d)), "c": cols, "acct": b64(d["ACCT_ID"].astype("<f8")),
        "out": b64(d["TOTAL_OUTSTANDING"].fillna(0).astype("<f4")), "det": det, "gmr": sorted(GMR), "tree": tree,
        "stamp": pd.Timestamp.now().strftime("%d-%m-%Y %H:%M")}

# ---------------------------------------------------------------- html
HTML = r"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>Net Metering Intelligence</title><link rel="icon" href="__LOGO__">
<style>
:root{--nv:#1b2d4b;--nv2:#243a5e;--bg:#f3f5f9;--tx:#1c2433;--mu:#6b7690;--ln:#e2e7f0;--tl:#3d8f95;--or:#e8812a;--rd:#c8453b;--gn:#2f9e6b}
*{box-sizing:border-box}body{margin:0;font:14px/1.45 "Segoe UI",Roboto,Arial,sans-serif;background:var(--bg);color:var(--tx)}
button{font:inherit;cursor:pointer}:focus-visible{outline:2px solid var(--tl);outline-offset:1px}
header{background:var(--nv);color:#fff;padding:14px 24px;display:flex;justify-content:space-between;align-items:center}
header h1{margin:0;font-size:20px}header small{color:#aebbd3}
.app{display:grid;grid-template-columns:300px 1fr;min-height:calc(100vh - 66px)}
aside{background:var(--nv);padding:16px;color:#dce5f4;overflow:auto;max-height:calc(100vh - 66px);position:sticky;top:0}
aside h3{margin:14px 0 8px;font-size:12px;color:#8ea3c6;font-weight:600}
.sh{display:flex;justify-content:space-between;align-items:center}
.ghost{background:none;border:1px solid #4a6089;color:#dce5f4;border-radius:8px;padding:5px 10px}
.dd{position:relative;margin-bottom:8px}
.dd>button{width:100%;display:flex;justify-content:space-between;align-items:center;background:var(--nv2);border:1px solid #38507a;color:#fff;border-radius:10px;padding:10px 12px;text-align:left}
.dd>button b{background:#e9eef7;color:var(--nv);border-radius:6px;padding:1px 9px;font-size:12px;max-width:130px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.dd>button b.on{background:var(--or);color:#fff}
.pn{display:none;position:absolute;z-index:20;left:0;right:0;top:calc(100% + 4px);background:#fff;color:var(--tx);border-radius:10px;box-shadow:0 12px 30px #0006;padding:8px}
.dd.open .pn{display:block}.pn input{width:100%;padding:8px;border:1px solid var(--ln);border-radius:8px;margin-bottom:6px}
.pn .ac{display:flex;gap:6px;margin-bottom:6px}.pn .ac button{flex:1;background:var(--bg);border:1px solid var(--ln);border-radius:6px;padding:4px;font-size:12px}
.ol{max-height:230px;overflow:auto}.ol label{display:flex;gap:8px;align-items:center;padding:5px 6px;border-radius:6px;cursor:pointer}
.ol label:hover{background:var(--bg)}.ol span{flex:1}.ol i{font-style:normal;color:var(--mu);font-size:12px}.ol .z{opacity:.4}
main{padding:20px 24px;min-width:0}
.hero{background:linear-gradient(120deg,var(--nv),#3a7f86);color:#fff;border-radius:16px;padding:20px 26px;display:flex;justify-content:space-between;align-items:center;gap:20px}
.hero h2{margin:0 0 4px;font-size:26px}.hero p{margin:0;color:#d5e3ee;max-width:640px}.hero .big{font-size:38px;font-weight:700;text-align:right}.hero .big small{display:block;font-size:12px;font-weight:400;color:#c4d6e4}
.tabs{display:flex;gap:10px;margin:18px 0 12px}.tabs button{background:#fff;border:1px solid var(--ln);border-radius:10px;padding:10px 18px;font-weight:600;color:var(--nv)}
.tabs button.on{background:var(--nv);color:#fff}
.chips{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:12px;min-height:6px}
.chip{background:#fff;border:1px solid var(--or);color:#8a4a10;border-radius:20px;padding:2px 4px 2px 12px;display:flex;gap:6px;align-items:center;font-size:13px}
.chip button{border:0;background:#fdeedd;border-radius:50%;width:20px;height:20px;line-height:18px;padding:0}
.kp{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px;margin-bottom:14px}
.card{background:#fff;border:1px solid var(--ln);border-radius:14px;padding:16px 18px;min-width:0}
.kp .card b{display:block;font-size:28px;margin:2px 0}.kp .card span{color:var(--mu);font-size:13px}.kp .card em{font-style:normal;font-size:12px;color:var(--mu)}
.kp .bad b{color:var(--rd)}.kp .good b{color:var(--gn)}
.g2{display:grid;grid-template-columns:repeat(auto-fit,minmax(380px,1fr));gap:14px;margin-bottom:14px}
.card h4{margin:0 0 2px;font-size:16px}.card .s{color:var(--mu);font-size:13px;margin-bottom:10px}
.br{display:grid;grid-template-columns:minmax(120px,190px) 1fr 70px;gap:10px;align-items:center;padding:4px 6px;border-radius:8px;cursor:pointer;font-size:13px}
.br:hover{background:var(--bg)}.br.on{background:#fdeedd}.br .t{background:#eef1f7;border-radius:6px;height:12px;overflow:hidden}.br .t i{display:block;height:100%;background:linear-gradient(90deg,var(--nv2),var(--tl));min-width:2px}
.br.w .t i{background:linear-gradient(90deg,var(--or),var(--rd))}.br b{text-align:right}
.bc{display:flex;flex-wrap:wrap;gap:4px;align-items:center;margin-bottom:10px;font-size:13px}.bc button{background:var(--bg);border:1px solid var(--ln);border-radius:6px;padding:2px 8px}.bc button:last-child{background:var(--nv);color:#fff}
.tw{overflow:auto;max-height:480px}table{border-collapse:collapse;width:100%;font-size:13px}
th{position:sticky;top:0;background:#f0f3f9;text-align:right;padding:8px 10px;white-space:nowrap;cursor:pointer;border-bottom:1px solid var(--ln)}th:first-child,td:first-child{text-align:left}
td{padding:7px 10px;text-align:right;border-bottom:1px solid var(--ln);white-space:nowrap}tbody tr.d{cursor:pointer}tbody tr.d:hover{background:#eef6f6}
tr.tot td{font-weight:700;background:#f6f8fc}tr.sub td{font-weight:600;background:#fafbfd;color:var(--nv)}
td.a{text-align:left}.lnk{color:var(--tl);font-weight:600}.pill{padding:2px 8px;border-radius:12px;font-size:12px;background:#eef1f7}.pill.x{background:#fdeedd;color:#8a4a10}.pill.r{background:#fbe3e0;color:var(--rd)}
.bar{display:inline-block;width:60px;height:8px;background:#eef1f7;border-radius:4px;vertical-align:middle;margin-left:6px}.bar i{display:block;height:100%;border-radius:4px;background:var(--gn)}
.tb{display:flex;gap:8px;margin:0 0 10px;flex-wrap:wrap;align-items:center}.tb input{flex:1;min-width:200px;padding:8px 10px;border:1px solid var(--ln);border-radius:8px}
.tb button,.pg button{background:#fff;border:1px solid var(--ln);border-radius:8px;padding:6px 12px}.pg{display:flex;gap:8px;align-items:center;justify-content:flex-end;margin-top:8px;color:var(--mu)}
#dr{position:fixed;top:0;right:-440px;width:420px;max-width:100%;height:100%;background:#fff;box-shadow:-10px 0 30px #0003;padding:20px;transition:right .2s;z-index:50;overflow:auto}#dr.o{right:0}
#dr dl{display:grid;grid-template-columns:130px 1fr;gap:6px 10px;font-size:13px}#dr dt{color:var(--mu)}#dr dd{margin:0;word-break:break-word}
@media(max-width:900px){.app{grid-template-columns:1fr}aside{position:static;max-height:none}.hero{flex-direction:column;align-items:flex-start}}
@media(prefers-reduced-motion:reduce){#dr{transition:none}}

body{--cd:#fff}body.dark{--bg:#0f1624;--cd:#182236;--tx:#e6ebf5;--ln:#2a3650;--mu:#93a1bd}
header{background:var(--cd);color:var(--tx);border-bottom:1px solid var(--ln);gap:14px;padding:12px 24px}header small{color:var(--mu)}
.lg0{background:var(--nv);color:#fff;font-weight:700;border-radius:10px;padding:9px 11px}.dt{background:var(--bg);border-radius:20px;padding:3px 12px;font-size:12px;color:var(--mu)}
#dk{background:var(--cd);color:var(--tx);border:1px solid var(--ln);border-radius:8px;padding:8px 16px}
.app{display:block}main{padding:16px 24px}.hero{display:none}
.fb{position:sticky;top:0;z-index:30;background:var(--cd);border-bottom:1px solid var(--ln);padding:12px 24px;display:flex;flex-wrap:wrap;gap:8px;align-items:center}
.fb #f1,.fb #f2{display:contents}.dd{margin:0;width:200px}.dd>button{background:var(--cd);color:var(--tx);border:1px solid var(--ln)}.dd>button b{background:var(--bg);color:var(--tx)}.dd>button b.on{background:var(--or);color:#fff}
.pn{background:var(--cd);color:var(--tx);min-width:260px}.pn input{background:var(--bg);color:var(--tx)}.pn .ac button{background:var(--bg);color:var(--tx)}
#rst{background:#b0463f;color:#fff;border:0;border-radius:8px;padding:10px 18px;font-weight:600}
.card,.tabs button,.chip,#dr{background:var(--cd);color:var(--tx)}.tabs button.on{background:var(--nv);color:#fff}.tb input,.tb button,.pg button{background:var(--cd);color:var(--tx)}
th{background:var(--bg);color:var(--tx)}tr.tot td,tr.sub td{background:var(--bg);color:var(--tx)}.br:hover{background:var(--bg)}.br .t,.bar{background:var(--bg)}
.kp .card{border-top:4px solid var(--nv)}.kp .card:nth-child(3){border-top-color:#5b8fb9}.kp .card:nth-child(4){border-top-color:#b04a45}.kp .card:nth-child(5){border-top-color:var(--or)}.kp .card:nth-child(6){border-top-color:var(--tl)}
.chh{display:flex;justify-content:space-between;gap:8px}.chh button{background:#5b8fb9;color:#fff;border:0;border-radius:8px;padding:7px 12px;margin-left:6px}.chh button[data-ex]{background:var(--bg);color:var(--tx);border:1px solid var(--ln)}
.cc{display:flex;gap:8px;align-items:flex-end;height:250px;padding-top:10px}.cl{flex:1;display:flex;flex-direction:column;align-items:center;height:100%;cursor:pointer;min-width:0}.cl em{font-style:normal;font-size:11px;color:var(--mu);text-align:center}
.cb{flex:1;width:100%;display:flex;align-items:flex-end;justify-content:center}.cb i{width:70%;border-radius:4px 4px 0 0;min-height:2px}.cl span{font-size:11px;color:var(--mu);text-align:center;margin-top:4px;word-break:break-word}.cl.on span{color:var(--or);font-weight:700}.cl:hover .cb i{opacity:.8}
.pie{display:flex;flex-wrap:wrap;align-items:center;justify-content:center;gap:10px}.pie svg{width:230px;max-width:100%}.lg{flex:1;min-width:190px}.lr{display:flex;gap:8px;align-items:center;padding:3px 6px;border-radius:6px;cursor:pointer;font-size:13px}.lr:hover,.lr.on{background:var(--bg)}.lr i{width:12px;height:12px;border-radius:3px}.lr b{margin-left:auto}
.ch.full{position:fixed;inset:5%;z-index:60;overflow:auto;box-shadow:0 0 0 100vmax #0008}.ch.full .cc{height:55vh}

.dd{width:180px}.dd>button{padding:8px 10px}.fb.col #f2{display:none}#mf{background:var(--cd);color:var(--tx);border:1px solid var(--ln);border-radius:8px;padding:9px 14px}
.pn{min-width:320px}#f2 .dd .pn,#f1 .dd:nth-child(n+5) .pn{left:auto;right:0}
.ol input{width:16px;height:16px;margin:0;flex:none}.ol label{align-items:center}.ol span{overflow-wrap:anywhere}
.pn.tp{min-width:440px}#tree{max-height:380px}.tn{display:flex;align-items:center;gap:6px;padding:3px 6px;border-radius:6px;font-size:13px}.tn:hover{background:var(--bg)}.tn.z{opacity:.45}
.tg{width:20px;border:0;background:none;color:var(--mu);padding:0}.tl{flex:1;cursor:pointer}.tl small{color:var(--mu);margin-left:4px}.tl.on{color:var(--or);font-weight:700}.tn i{font-style:normal;color:var(--mu);font-size:12px}
.pre{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:10px}.pre button{background:var(--cd);color:var(--tx);border:1px solid var(--ln);border-radius:20px;padding:5px 14px;font-size:13px}.pre button:hover{border-color:var(--tl);color:var(--tl)}

/* left sidebar layout */
.app{display:grid;grid-template-columns:310px minmax(0,1fr);align-items:start}
.fb{position:sticky;top:0;height:100vh;overflow:auto;flex-direction:column;flex-wrap:nowrap;align-items:stretch;gap:8px;padding:16px;border-bottom:0;border-right:1px solid var(--ln)}
.fb #f1,.fb #f2{display:block}.fb.col #f2{display:block}#mf{display:none}
#f1::before,#f2::before{display:block;font-size:12px;font-weight:600;color:var(--mu);margin:6px 2px 8px}#f1::before{content:"Geography hierarchy"}#f2::before{content:"Meter and billing filters";margin-top:16px}
#rst{order:-1}.dd{width:100%;margin-bottom:8px}
.dd .pn{position:static;min-width:0;margin-top:6px;box-shadow:none;border:1px solid var(--ln)}.dd .pn.tp{min-width:0}
@media(max-width:900px){.app{grid-template-columns:1fr}.fb{position:static;height:auto}}
.pr{padding:10px 12px;border-radius:10px;cursor:pointer}.pr:hover,.pr.on{background:var(--bg)}.ph{display:flex;gap:10px;align-items:baseline}.ph span{flex:1;color:var(--mu);font-size:13px}.ph strong{font-size:18px}
.pt{height:12px;background:var(--bg);border-radius:6px;overflow:hidden;margin:6px 0 2px;border:1px solid var(--ln)}.pt i{display:block;height:100%;background:var(--gn);min-width:2px}.pr small{color:var(--mu)}
.lgi{width:46px;height:46px;object-fit:contain;border-radius:50%}
</style></head><body>
<header><img class="lgi" src="__LOGO__" alt="DVVNL logo"><div style="flex:1"><h1>Net Metering Intelligence</h1><small>DVVNL · Smart net meter billing and tagging</small></div><span class="dt" id="stamp"></span><button id="dk">Dark</button></header>
<div class="app"><section class="fb col"><div id="f1"></div><div id="f2"></div><button id="mf">More filters</button><button id="rst">Clear all</button></section>
<main><div class="hero"><div><h2 id="ht"></h2><p id="hp"></p></div><div class="big"><span id="hn"></span><small id="hl"></small></div></div>
<div class="tabs"><button data-t="bill">Net meter billing and exceptions</button><button data-t="wrong">Wrong tagging</button></div>
<div class="pre" id="pre"></div><div class="chips" id="chips"></div><div id="view"></div></main></div>
<div id="dr" aria-live="polite"></div>
<script>
const D=__DATA__,N=D.n,$=s=>document.querySelector(s);
const u8=s=>{const x=atob(s),u=new Uint8Array(x.length);for(let i=0;i<x.length;i++)u[i]=x.charCodeAt(i);return u.buffer};
const C={};for(const k in D.c)C[k]={v:D.c[k].v,a:new Uint16Array(u8(D.c[k].d))};
const ACCT=new Float64Array(u8(D.acct)),OUT=new Float32Array(u8(D.out)),GMR=new Set(D.gmr);
const LV=[["ZONE","Zone"],["CIRCLE","Circle"],["DIV","Division"],["SDO","SDO"],["SS","Substation"]];
const FL=[...LV,["GRP","Franchisee"],["SMT","Meter type"],["CYC","Billing cycle"],["TAR","Tariff"],["REM","Tagging result"],["BST","Bill status"],["CON","Connection status"],["MAKE","Meter make"]];
const NAME=Object.fromEntries(FL),sel={};FL.forEach(f=>sel[f[0]]=new Set());
const BV=C.BST.v,bk=BV.map(v=>v==="Non-smart billed"?5:v==="Non-smart unbilled"?6:v==="Billed"?1:v==="Unbilled - no status"?2:v==="Completed"?3:4);
const ECODES=BV.map((v,i)=>bk[i]===4?i:-1).filter(i=>i>=0).sort((a,b)=>BV[a].localeCompare(BV[b]));
const RV=C.REM.v,OKR=RV.indexOf("Correctly Tagged"),R_NM=RV.indexOf("Eligible for Net Meter"),R_NB=RV.indexOf("Eligible for Net Bill");
let TB,M,tab="bill",open=null,pg=0,q="",sk="t",sd=-1;
const f=n=>Math.round(n).toLocaleString("en-IN"),pc=(a,b)=>b?(100*a/b).toFixed(1)+"%":"-",
esc=s=>String(s).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
function mask(skip){const SK=skip==="GEO"?["ZONE","CIRCLE","DIV","SDO","SS"]:[skip],act=FL.filter(x=>!SK.includes(x[0])&&sel[x[0]].size).map(x=>[C[x[0]].a,sel[x[0]]]),m=new Uint8Array(N);
 o:for(let i=0;i<N;i++){for(const[a,s]of act)if(!s.has(a[i]))continue o;m[i]=1}return m}
const level=()=>LV.find(l=>sel[l[0]].size!==1);
function toggle(k,code){const s=sel[k];s.has(code)?s.delete(code):s.add(code);go()}
function only(k,code){sel[k]=new Set([code]);pg=0;go()}
// ---- sidebar dropdowns
function buildSide(){FL.forEach(([k,l],i)=>{const el=document.createElement("div");el.className="dd";el.id="dd"+k;
 el.innerHTML=`<button type="button"><span>${l}</span><b>All</b></button><div class="pn"><input placeholder="Search ${l.toLowerCase()}..."><div class="ac"><button type="button" data-a="v">Select visible</button><button type="button" data-a="c">Clear</button></div><div class="ol"></div></div>`;
 $(i<6?"#f1":"#f2").appendChild(el);
 el.firstChild.onclick=e=>{open=open===k?null:k;el.querySelector("input").value="";side();if(open===k)el.querySelector("input").focus()};
 el.querySelector("input").oninput=()=>fill(k);
 el.querySelectorAll(".ac button").forEach(b=>b.onclick=()=>{if(b.dataset.a==="c")sel[k].clear();else vis(k).forEach(c=>sel[k].add(c));go()});
 el.querySelector(".ol").onclick=e=>{const c=e.target.closest("input");if(c)toggle(k,+c.value)}});}
function vis(k){const t=$("#dd"+k+" input").value.toLowerCase(),m=mask(k),a=C[k].a,n=C[k].v.length,ct=new Array(n).fill(0);
 for(let i=0;i<N;i++)if(m[i])ct[a[i]]++;vis.ct=ct;return C[k].v.map((v,i)=>i).filter(i=>(ct[i]||sel[k].has(i))&&C[k].v[i].toLowerCase().includes(t))}
function fill(k){const ids=vis(k),ct=vis.ct;ids.sort((a,b)=>ct[b]-ct[a]);
 $("#dd"+k+" .ol").innerHTML=ids.slice(0,300).map(i=>`<label class="${ct[i]?"":"z"}"><input type="checkbox" value="${i}" ${sel[k].has(i)?"checked":""}><span>${esc(C[k].v[i])}</span><i>${f(ct[i])}</i></label>`).join("")||"<i>No match</i>"}
function side(){FL.forEach(([k])=>{const el=$("#dd"+k),s=sel[k],b=el.querySelector("b");
 b.textContent=s.size===0?"All":s.size===1?C[k].v[[...s][0]]:s.size+" selected";b.className=s.size?"on":"";el.classList.toggle("open",open===k)});if(open&&open!=="TREE")fill(open);$("#ddTREE").classList.toggle("open",open==="TREE");if(open==="TREE")fillTree();treeLbl()}
document.addEventListener("click",e=>{if(open&&!e.composedPath().some(n=>n.classList&&n.classList.contains("dd"))){open=null;side()}});
document.addEventListener("keydown",e=>{if(e.key==="Escape"){$("#dr").classList.remove("o");document.querySelectorAll(".ch.full").forEach(x=>x.classList.remove("full"));if(open){open=null;side()}}});
$("#rst").onclick=()=>{FL.forEach(x=>sel[x[0]].clear());pg=0;go()};
document.querySelectorAll(".tabs button").forEach(b=>b.onclick=()=>{tab=b.dataset.t;pg=0;q="";go()});
// ---- hierarchy tree (from Excel)
const TREE=(()=>{const Z=new Map,mk=(m,k,n,c)=>{let x=m.get(n);if(!x)m.set(n,x={k,n,c,ch:new Map});return x};
 D.tree.forEach(([zn,cn,dn,sn,zc,cc,dc,sc])=>{const z=mk(Z,"ZONE",zn,zc),c=mk(z.ch,"CIRCLE",cn,cc),d=mk(c.ch,"DIV",dn,dc);mk(d.ch,"SDO",sn,sc)});return[...Z.values()]})();
let tq="",tEx=new Set;
function fillTree(){const m=mask("GEO"),a=C.SDO.a,ct=new Array(C.SDO.v.length).fill(0);for(let i=0;i<N;i++)if(m[i])ct[a[i]]++;const t=tq.toLowerCase();
 function walk(n,path,dep){const kids=[...n.ch.values()].map(x=>walk(x,[...path,x],dep+1)),own=n.k==="SDO"?ct[C.SDO.v.indexOf(n.n)]||0:0,cnt=own+kids.reduce((s,x)=>s+x.cnt,0),
  any=!t||n.n.toLowerCase().includes(t)||n.c.toLowerCase().includes(t)||kids.some(x=>x.show),id=n.k+n.n,exp=t?kids.some(x=>x.show):tEx.has(id),ix=C[n.k].v.indexOf(n.n),on=sel[n.k].size===1&&sel[n.k].has(ix);
  return{cnt,show:any,h:any?`<div class="tn ${cnt?"":"z"}" style="padding-left:${dep*18}px"><button type="button" class="tg" data-t="${esc(id)}">${n.ch.size?(exp?"▾":"▸"):""}</button><span class="tl ${on?"on":""}" data-p="${esc(path.map(x=>x.k+"|"+x.n).join("~"))}">${esc(n.n)}<small>${n.c}</small></span><i>${f(cnt)}</i></div>${exp?kids.map(x=>x.h).join(""):""}`:""}}
 $("#tree").innerHTML=TREE.map(z=>walk(z,[z],0).h).join("")||"<i>No match</i>"}
function treeLbl(){const p=LV.filter(l=>sel[l[0]].size===1).map(l=>C[l[0]].v[[...sel[l[0]]][0]]);$("#ddTREE b").textContent=p.length?p[p.length-1]:"Browse";$("#ddTREE b").className=p.length?"on":""}
function treeInit(){const el=document.createElement("div");el.className="dd";el.id="ddTREE";
 el.innerHTML='<button type="button"><span>Hierarchy</span><b>Browse</b></button><div class="pn tp"><input id="tq" placeholder="Search zone, circle, division, SDO or code..."><div class="ac"><button type="button" data-a="x">Expand all</button><button type="button" data-a="c">Collapse all</button></div><div class="ol" id="tree"></div></div>';$("#f1").prepend(el);
 el.firstChild.onclick=()=>{open=open==="TREE"?null:"TREE";side()};$("#tq").oninput=e=>{tq=e.target.value;fillTree()};
 el.querySelectorAll(".ac button").forEach(b=>b.onclick=()=>{tEx=new Set;if(b.dataset.a==="x")(function r(ns){ns.forEach(n=>{tEx.add(n.k+n.n);r([...n.ch.values()])})})(TREE);fillTree()});
 $("#tree").onclick=e=>{const g=e.target.closest(".tg"),l=e.target.closest(".tl");if(g){tEx.has(g.dataset.t)?tEx.delete(g.dataset.t):tEx.add(g.dataset.t);fillTree()}
  else if(l){const P=l.dataset.p.split("~").map(x=>x.split("|"));if(P.some(([k,n])=>C[k].v.indexOf(n)<0))return;LV.forEach(([k])=>sel[k].clear());P.forEach(([k,n])=>sel[k]=new Set([C[k].v.indexOf(n)]));pg=0;go()}}}
const PRE=[["Smart (MD03)","",{SMT:()=>[C.SMT.v.indexOf("Smart (MD03)")]}],["Non-smart","",{SMT:()=>[C.SMT.v.indexOf("Non-smart")]}],["Unbilled with exception","bill",{BST:()=>ECODES}],["Unbilled, no status","bill",{BST:()=>[BV.indexOf("Unbilled - no status")]}],["Should be Net Meter","wrong",{REM:()=>[R_NM]}],["Should be Net Bill","wrong",{REM:()=>[R_NB]}],["GMR","",{GRP:()=>[C.GRP.v.indexOf("GMR")]}],["Genus","",{GRP:()=>[C.GRP.v.indexOf("Genus")]}]];
$("#pre").innerHTML=PRE.map((p,i)=>`<button data-i="${i}">${p[0]}</button>`).join("");
$("#pre").onclick=e=>{const b=e.target.closest("button");if(!b)return;const p=PRE[+b.dataset.i];FL.forEach(x=>sel[x[0]].clear());for(const k in p[2])sel[k]=new Set(p[2][k]());if(p[1])tab=p[1];pg=0;go()};
$("#mf").onclick=()=>{const c=$(".fb").classList.toggle("col");$("#mf").textContent=c?"More filters":"Fewer filters"};
function top10(t,k){const r=agg("DIV").map(x=>({...x,v:k==="u"?x.u+x.nu:x.w})).sort((a,b)=>b.v-a.v).slice(0,10).filter(x=>x.v);
 return `<div class="card"><h4>${t}</h4><div class="s">Click a division to filter.</div>${bars(r.map(x=>({n:x.v,c:x.i,k:"DIV",l:x.name,on:sel.DIV.has(x.i)})),0,{w:1})}</div>`}
// ---- aggregation
function agg(col){const a=C[col].a,bs=C.BST.a,rm=C.REM.a,n=C[col].v.length,R=Array.from({length:n},()=>({a:0,sa:0,na:0,t:0,b:0,u:0,c:0,e:0,x:{},ok:0,w:0,nm:0,nb:0,sw:0,nw:0,nbd:0,nu:0,o:0}));
 for(let i=0;i<N;i++){if(!M[i])continue;const r=R[a[i]],k=bk[bs[i]],sm=k<5;r.a++;r.o+=OUT[i];sm?r.sa++:r.na++;
  if(rm[i]===OKR)r.ok++;else{r.w++;sm?r.sw++:r.nw++;if(rm[i]===R_NM)r.nm++;else if(rm[i]===R_NB)r.nb++}
  if(!sm){if(k===5)r.nbd++;else r.nu++;continue}
  r.t++;if(k===1)r.b++;else{r.u++;if(k===3)r.c++;if(k===4){r.e++;r.x[bs[i]]=(r.x[bs[i]]||0)+1}}}
 return R.map((r,i)=>({...r,i,name:C[col].v[i]})).filter(r=>r.a)}
const sum=rs=>rs.reduce((s,r)=>{for(const k in r)if(typeof r[k]==="number"&&k!=="i")s[k]=(s[k]||0)+r[k];for(const c in r.x)s.x[c]=(s.x[c]||0)+r.x[c];return s},{x:{}});
function bars(items,total,{filt,w}={}){const mx=Math.max(1,...items.map(x=>x.n));
 return items.map(x=>`<div class="br ${w?"w":""} ${x.on?"on":""}" data-k="${x.k}" data-c="${x.c}" title="Click to filter"><span>${esc(x.l)}</span><div class="t"><i style="width:${100*x.n/mx}%"></i></div><b>${f(x.n)}</b></div>`).join("")}
function dist(col,mk){const a=C[col].a,ct=new Array(C[col].v.length).fill(0);for(let i=0;i<N;i++)if(M[i]&&(!mk||mk(i)))ct[a[i]]++;
 return ct.map((n,c)=>({n,c,k:col,l:C[col].v[c],on:sel[col].has(c)})).filter(x=>x.n).sort((a,b)=>b.n-a.n)}
// ---- drill table
function drill(rows,cols,total,col){const L=level(),sub=L&&L[0]==="ZONE";
 const s=rows.slice().sort((a,b)=>sd*((typeof a[sk]==="string"?a[sk].localeCompare(b[sk]):a[sk]-b[sk])||0));
 const cell=(r,c)=>c.h?c.h(r):c.k==="name"?esc(r.name):f(r[c.k]),mxs={};cols.forEach(c=>{if(/^(x|e$|u$|w$|nm$|nb$|sw$|nw$)/.test(c.k))mxs[c.k]=Math.max(1,...rows.map(r=>r[c.k]||0))});
 const hs=(r,c)=>mxs[c.k]&&r[c.k]?` style="background:rgba(232,129,42,${(.6*r[c.k]/mxs[c.k]).toFixed(2)})"`:"";TB={rows:s,cols,total};
 const row=(r,cls)=>`<tr class="${cls}">${cols.map((c,j)=>`<td>${j?cell(r,c):r.name}</td>`).join("")}</tr>`;
 let body=s.map(r=>`<tr class="d" data-i="${r.i}">${cols.map((c,j)=>`<td class="${j?"":"lnk"}"${hs(r,c)}>${cell(r,c)}</td>`).join("")}</tr>`).join("");
 let ft="";if(sub){const g=n=>{const x=sum(rows.filter(r=>(GMR.has(r.name))===(n==="GMR")));x.name=n+" total";return x};
  ft=row(g("GMR"),"sub")+row(g("Genus"),"sub")}
 return `<div class="tb"><button id="tcsv">Export table CSV</button></div><div class="tw"><table><thead><tr>${cols.map(c=>`<th data-k="${c.k}">${c.l}${sk===c.k?(sd>0?" ▲":" ▼"):""}</th>`).join("")}</tr></thead><tbody>${body}${ft}${row({...total,name:"DVVNL total"},"tot")}</tbody></table></div>`}
function crumbs(){let h=`<button data-n="0">All zones</button>`;LV.forEach(([k],i)=>{if(sel[k].size===1)h+=` › <button data-n="${i+1}">${esc(C[k].v[[...sel[k]][0]])}</button>`});return `<div class="bc">${h}</div>`}
// ---- consumer list
function list(pred,extra,title){const L=[];for(let i=0;i<N;i++)if(M[i]&&pred(i))L.push(i);
 const t=q.toLowerCase(),F=t?L.filter(i=>String(ACCT[i]).includes(t)||(D.det[i]&&D.det[i].slice(0,3).join(" ").toLowerCase().includes(t))):L,P=25,pages=Math.max(1,Math.ceil(F.length/P));pg=Math.min(pg,pages-1);
 const rows=F.slice(pg*P,pg*P+P).map(i=>`<tr class="d" data-r="${i}"><td class="a lnk">${String(ACCT[i])}</td><td class="a">${esc(D.det[i]?D.det[i][0]:"")}</td><td class="a">${esc(C.DIV.v[C.DIV.a[i]])}</td><td class="a">${esc(C.SDO.v[C.SDO.a[i]])}</td>${extra.map(e=>`<td class="a">${e[1](i)}</td>`).join("")}<td>${f(OUT[i])}</td></tr>`).join("");
 return `<div class="card"><h4>${title} (${f(F.length)})</h4><div class="s">Click an account for the full consumer view.</div><div class="tb"><input id="q" placeholder="Search account id, name or meter serial" value="${esc(q)}"><button id="csv">Export CSV</button></div>
 <div class="tw"><table><thead><tr><th>Account ID</th><th>Name</th><th>Division</th><th>SDO</th>${extra.map(e=>`<th style="text-align:left">${e[0]}</th>`).join("")}<th>Outstanding</th></tr></thead><tbody>${rows||'<tr><td colspan="9">No accounts match the current filters.</td></tr>'}</tbody></table></div>
 <div class="pg"><button id="pv">Prev</button> Page ${pg+1} of ${pages} <button id="nx">Next</button></div></div>`}
function csv(pred){const H=["ACCT_ID",...FL.map(x=>x[1]),"Outstanding","Name","Mobile","Meter serial","Feeder","Substation","Address","Last bill date","Pay date","Assessment","Load"],o=[H];
 for(let i=0;i<N;i++)if(M[i]&&pred(i))o.push([String(ACCT[i]),...FL.map(x=>C[x[0]].v[C[x[0]].a[i]]),OUT[i],...(D.det[i]||Array(10).fill(""))]);
 const b=new Blob(["\ufeff"+o.map(r=>r.map(v=>'"'+String(v).replace(/"/g,'""')+'"').join(",")).join("\n")],{type:"text/csv"}),a=document.createElement("a");
 a.href=URL.createObjectURL(b);a.download="net_meter_"+tab+"_export.csv";a.click()}
function consumer(i){const d=D.det[i]||[],g=k=>C[k].v[C[k].a[i]],rows=[["Account ID",String(ACCT[i])],["Name",d[0]||"(details not embedded for correctly tagged, billed accounts)"],["Mobile",d[1]],["Address",d[5]],
 ["Zone / Circle",g("ZONE")+" / "+g("CIRCLE")],["Division",g("DIV")],["SDO",g("SDO")],["Substation / Feeder",(d[4]||"")+" / "+(d[3]||"")],["Meter serial",d[2]],["Meter make",g("MAKE")],["Meter config",g("CFG")],["Configured tag",g("CTAG")],["Eligible tag",g("ETAG")],["Tagging result",g("REM")],["Bill status",g("BST")],["Billing cycle",g("CYC")],["Tariff",g("TAR")],["Last bill / Pay date",(d[6]||"")+" / "+(d[7]||"")],["Assessment",d[8]],["Outstanding",f(OUT[i])],["Load",d[9]]];
 $("#dr").innerHTML=`<button class="ghost" style="color:#1b2d4b;border-color:#c8d0e0" id="cl">Close</button><h3>Consumer 360</h3><dl>${rows.map(r=>`<dt>${r[0]}</dt><dd>${esc(r[1]||"-")}</dd>`).join("")}</dl>`;$("#dr").classList.add("o");$("#cl").onclick=()=>$("#dr").classList.remove("o")}
// ---- tabs
const PAL=["#2b3558","#b04a45","#4f9d5c","#d98b45","#e0b84a","#5b8fb9","#7b5ea7","#8a8f9c"],CH={};
function chartCard(id,t,sub,type,items){CH[id]=items;const tot=items.reduce((a,x)=>a+x.n,0)||1,pp=n=>f(n)+" ("+(100*n/tot).toFixed(1)+"%)";let b;
 if(type==="col"){const mx=Math.max(1,...items.map(x=>x.n));b=`<div class="cc">${items.map((x,i)=>`<div class="cl ${x.on?"on":""}" data-k="${x.k}" data-c="${x.c}"><em>${pp(x.n)}</em><div class="cb"><i style="height:${100*x.n/mx}%;background:${PAL[i%8]}"></i></div><span>${esc(x.l)}</span></div>`).join("")}</div>`}
 else{const R=70,K=2*Math.PI*R;let o=0;const sg=items.map((x,i)=>{const l=K*x.n/tot,e=`<circle cx="100" cy="100" r="${R}" fill="none" stroke="${PAL[i%8]}" stroke-width="${x.on?38:28}" stroke-dasharray="${l} ${K-l}" stroke-dashoffset="${-o}" transform="rotate(-90 100 100)" data-k="${x.k}" data-c="${x.c}" style="cursor:pointer"><title>${esc(x.l)}: ${pp(x.n)}</title></circle>`;o+=l;return e}).join("");
  b=`<div class="pie"><svg viewBox="0 0 200 200">${sg}<text x="100" y="98" text-anchor="middle" font-size="22" font-weight="700" fill="currentColor">${f(tot)}</text><text x="100" y="116" text-anchor="middle" font-size="10" fill="#8a93a8">accounts</text></svg><div class="lg">${items.map((x,i)=>`<div class="lr ${x.on?"on":""}" data-k="${x.k}" data-c="${x.c}"><i style="background:${PAL[i%8]}"></i>${esc(x.l)}<b>${pp(x.n)}</b></div>`).join("")}</div></div>`}
 return `<div class="card ch" id="ch-${id}"><div class="chh"><div><h4>${t}</h4><div class="s">${sub}</div></div><div><button data-ex="${id}" title="Expand">⤢</button><button data-csv="${id}">CSV</button></div></div>${b}</div>`}
const kpi=(l,v,s,c="")=>`<div class="card ${c}"><span>${l}</span><b>${v}</b><em>${s}</em></div>`;
const sumz=rs=>{const T=sum(rs);["a","sa","na","t","b","u","c","e","nbd","nu","w","ok","nm","nb","sw","nw"].forEach(k=>T[k]=T[k]||0);return T};
const prog=(rows,wrong)=>rows.filter(x=>x[1]).map(x=>{const c=C.SMT.v.indexOf(x[0]);return `<div class="pr ${sel.SMT.has(c)?"on":""}" data-k="SMT" data-c="${c}"><div class="ph"><b>${x[0]}</b><span>${f(x[2])} ${wrong?"wrong":"billed"} of ${f(x[1])}</span><strong>${pc(x[2],x[1])}</strong></div><div class="pt"><i style="width:${Math.min(100,wrong?500*x[2]/x[1]:100*x[2]/x[1])}%;${wrong?"background:var(--or)":""}"></i></div>${wrong?"":`<small>${f(x[1]-x[2])} unbilled</small>`}</div>`}).join("")||"<div class='s'>No accounts in scope.</div>";
function viewBill(){const L=level(),col=L?L[0]:"SDO",rs=agg(col),T=sumz(rs),X=ECODES.map(c=>({n:T.x[c]||0,c,k:"BST",l:BV[c],on:sel.BST.has(c)}));
 const cols=[{k:"name",l:L?L[1]:"SDO"},{k:"t",l:"Smart accounts"},{k:"b",l:"Smart billed"},{k:"p",l:"Smart billed %",h:r=>pc(r.b,r.t)+`<span class="bar"><i style="width:${r.t?100*r.b/r.t:0}%"></i></span>`},{k:"u",l:"Smart unbilled"},
  ...ECODES.map(c=>({k:"x"+c,l:BV[c],h:r=>f(r.x[c]||0)})),{k:"e",l:"Exceptions"},{k:"r",l:"Read to be sent by AMISP",h:r=>f(r.u-r.e)},
  {k:"na",l:"Non-smart accounts"},{k:"nbd",l:"Non-smart billed"},{k:"np",l:"Non-smart billed %",h:r=>pc(r.nbd,r.na)+`<span class="bar"><i style="width:${r.na?100*r.nbd/r.na:0}%"></i></span>`}];
 const fx=r=>{r.p=r.t?r.b/r.t:0;r.r=r.u-r.e;r.np=r.na?r.nbd/r.na:0;ECODES.forEach(c=>r["x"+c]=r.x[c]||0)};rs.forEach(fx);fx(T);
 const un=i=>[2,3,4,6].includes(bk[C.BST.a[i]]);
 return `<div class="kp">${kpi("Total accounts",f(T.a),"smart "+f(T.sa)+" · non-smart "+f(T.na))}${kpi("Smart billed %",pc(T.b,T.t),f(T.b)+" of "+f(T.t)+" MD03 accounts","good")}${kpi("Non-smart billed %",pc(T.nbd,T.na),f(T.nbd)+" of "+f(T.na)+" accounts","good")}
 ${kpi("Smart unbilled",f(T.u),"MD03 accounts not billed","bad")}${kpi("Smart exceptions",f(T.e),pc(T.e,T.u)+" of smart unbilled","bad")}${kpi("Read to be sent by AMISP",f(T.r),"smart unbilled minus exceptions")}</div>
 <div class="g2"><div class="card"><h4>Smart vs non-smart billing progress</h4><div class="s">Smart = bill cycle MD03; every other cycle is non-smart. Click a row to filter.</div>${prog([["Smart (MD03)",T.t,T.b],["Non-smart",T.na,T.nbd]])}</div>
 ${chartCard("ex","Smart exception breakdown","Click a column to filter. Completed is for reference, not an exception.","col",X.concat([{n:T.c,c:BV.indexOf("Completed"),k:"BST",l:"Completed",on:sel.BST.has(BV.indexOf("Completed"))}]))}${top10("Top 10 divisions by unbilled accounts","u")}</div>
 <div class="card" style="margin-bottom:14px"><h4>Billing status by ${L?L[1].toLowerCase():"consumer"}</h4><div class="s">Click a row to drill down. Smart and non-smart are shown side by side.</div>${crumbs()}${L?drill(rs,cols,T,col):"<div class='s'>Showing individual accounts below.</div>"}</div>
 ${list(un,[["Bill status",i=>`<span class="pill ${bk[C.BST.a[i]]===4?"x":""}">${esc(C.BST.v[C.BST.a[i]])}</span>`]],"Unbilled accounts")}`}
function viewWrong(){const L=level(),col=L?L[0]:"SDO",rs=agg(col),T=sumz(rs),wr=i=>C.REM.a[i]!==OKR;
 const fx=r=>{r.p=r.a?r.w/r.a:0};rs.forEach(fx);fx(T);
 const cols=[{k:"name",l:L?L[1]:"SDO"},{k:"a",l:"Accounts"},{k:"ok",l:"Correctly tagged"},{k:"w",l:"Wrongly tagged"},{k:"p",l:"Wrong %",h:r=>pc(r.w,r.a)+`<span class="bar"><i style="width:${Math.min(100,r.a?500*r.w/r.a:0)}%;background:var(--or)"></i></span>`},{k:"nm",l:"Should be Net Meter"},{k:"nb",l:"Should be Net Bill"},{k:"sw",l:"Smart wrong"},{k:"nw",l:"Non-smart wrong"}];
 const mx={};for(let i=0;i<N;i++)if(M[i]&&wr(i)){const r=C.REM.v[C.REM.a[i]],q=bk[C.BST.a[i]]<5?0:1;(mx[r]=mx[r]||[0,0])[q]++}
 const mt=Object.entries(mx).map(([r,v])=>`<tr><td class="a">${esc(r)}</td><td>${f(v[0])}</td><td>${f(v[1])}</td><td><b>${f(v[0]+v[1])}</b></td></tr>`).join("");
 return `<div class="kp">${kpi("Accounts in scope",f(T.a),"smart "+f(T.sa)+" · non-smart "+f(T.na))}${kpi("Correctly tagged",f(T.ok),pc(T.ok,T.a),"good")}${kpi("Wrongly tagged",f(T.w),pc(T.w,T.a)+" per remarks","bad")}
 ${kpi("Smart wrongly tagged",f(T.sw),pc(T.sw,T.sa)+" of smart accounts","bad")}${kpi("Non-smart wrongly tagged",f(T.nw),pc(T.nw,T.na)+" of non-smart accounts","bad")}${kpi("Should be Net Meter",f(T.nm),"remark: Eligible for Net Meter","bad")}${kpi("Should be Net Bill",f(T.nb),"remark: Eligible for Net Bill","bad")}</div>
 <div class="g2">${chartCard("to","Tagging outcome","Based on the remarks column only.","pie",dist("REM"))}
 <div class="card"><h4>Smart vs non-smart wrong tagging</h4><div class="s">Click a row to filter by meter type.</div>${prog([["Smart (MD03)",T.sa,T.sw],["Non-smart",T.na,T.nw]],1)}
 <table style="margin-top:10px"><thead><tr><th>Remark</th><th>Smart</th><th>Non-smart</th><th>Total</th></tr></thead><tbody>${mt||'<tr><td colspan="4">No wrong tagging in scope.</td></tr>'}</tbody></table></div>${top10("Top 10 divisions by wrong tagging","w")}</div>
 <div class="card" style="margin-bottom:14px"><h4>Wrong tagging by ${L?L[1].toLowerCase():"consumer"}</h4><div class="s">Click a row to drill down.</div>${crumbs()}${L?drill(rs,cols,T,col):"<div class='s'>Showing individual accounts below.</div>"}</div>
 ${list(wr,[["Remark",i=>`<span class="pill r">${esc(C.REM.v[C.REM.a[i]])}</span>`],["Meter type",i=>C.SMT.v[C.SMT.a[i]]]],"Wrongly tagged accounts")}`}
// ---- render
let pred=()=>true;
function go(){M=mask(null);side();const B=tab==="bill";
 document.querySelectorAll(".tabs button").forEach(b=>b.classList.toggle("on",b.dataset.t===tab));
 $("#chips").innerHTML=FL.flatMap(([k,l])=>[...sel[k]].map(c=>`<span class="chip">${l}: ${esc(C[k].v[c])}<button data-k="${k}" data-c="${c}" aria-label="Remove">×</button></span>`)).join("");
 let n=0;for(let i=0;i<N;i++)n+=M[i];
 $("#ht").textContent=B?"Net meter billing and exceptions":"Wrong tagging";$("#hp").textContent=B?"Smart net meter billing status with exception drill-down from zone to consumer.":"Net Meter and Net Bill accounts where the configured tag differs from the eligible tag.";
 $("#hn").textContent=f(n);$("#hl").textContent="accounts in current filter";
 $("#view").innerHTML=B?viewBill():viewWrong();pred=B?i=>[2,3,4,6].includes(bk[C.BST.a[i]]):i=>C.REM.a[i]!==OKR;wire()}
function wire(){const V=$("#view");
 V.querySelectorAll("[data-k][data-c]").forEach(e=>e.onclick=()=>{pg=0;toggle(e.dataset.k,+e.dataset.c)});
 V.querySelectorAll("[data-ex]").forEach(e=>e.onclick=()=>$("#ch-"+e.dataset.ex).classList.toggle("full"));
 V.querySelectorAll("[data-csv]").forEach(e=>e.onclick=()=>{const o=[["Item","Count"],...CH[e.dataset.csv].map(x=>[x.l,x.n])],b=new Blob(["\ufeff"+o.map(r=>r.map(v=>'"'+v+'"').join(",")).join("\n")],{type:"text/csv"}),a=document.createElement("a");a.href=URL.createObjectURL(b);a.download=e.dataset.csv+"_chart.csv";a.click()});
 V.querySelectorAll("th[data-k]").forEach(e=>e.onclick=()=>{const k=e.dataset.k;sd=sk===k?-sd:-1;sk=k;go()});
 const L=level();V.querySelectorAll("tr.d[data-i]").forEach(e=>e.onclick=()=>only(L[0],+e.dataset.i));
 V.querySelectorAll(".bc button").forEach(e=>e.onclick=()=>{const n=+e.dataset.n;LV.forEach(([k],i)=>{if(i>=n)sel[k].clear()});pg=0;go()});
 V.querySelectorAll("tr.d[data-r]").forEach(e=>e.onclick=()=>consumer(+e.dataset.r));
 const tc=$("#tcsv");if(tc)tc.onclick=()=>{const o=[TB.cols.map(c=>c.l),...[...TB.rows,{...TB.total,name:"DVVNL total"}].map(r=>TB.cols.map(c=>c.k==="name"?r.name:c.k==="p"?(100*r.p).toFixed(1)+"%":r[c.k]||0))],b=new Blob(["\ufeff"+o.map(r=>r.map(v=>'"'+v+'"').join(",")).join("\n")],{type:"text/csv"}),a=document.createElement("a");a.href=URL.createObjectURL(b);a.download="net_meter_"+tab+"_table.csv";a.click()};
 const qi=$("#q");if(qi){qi.oninput=()=>{q=qi.value;pg=0;const p=qi.selectionStart;go();const n=$("#q");n.focus();n.setSelectionRange(p,p)};$("#pv").onclick=()=>{pg=Math.max(0,pg-1);go()};$("#nx").onclick=()=>{pg++;go()};$("#csv").onclick=()=>csv(pred)}}
$("#chips").onclick=e=>{const b=e.target.closest("button");if(b)toggle(b.dataset.k,+b.dataset.c)};
$("#stamp").textContent=N.toLocaleString("en-IN")+" source rows · "+D.stamp;$("#dk").onclick=()=>{document.body.classList.toggle("dark");$("#dk").textContent=document.body.classList.contains("dark")?"Light":"Dark"};
buildSide();treeInit();go();
</script></body></html>"""

def logo_uri():                       # DVVNL_Logo.png (same folder) is cropped to the circle and embedded
    f = BASE / "DVVNL_Logo.png"
    try:
        from PIL import Image, ImageChops
        import io
        im = Image.open(f).convert("RGBA")
        bg = Image.new("RGBA", im.size, (255, 255, 255, 255))
        box = ImageChops.difference(im.convert("RGB"), bg.convert("RGB")).getbbox()
        im = im.crop(box).resize((96, 96), Image.LANCZOS) if box else im.resize((96, 96))
        buf = io.BytesIO(); im.save(buf, "PNG")
        return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()
    except Exception as e:
        print("Logo not embedded:", e)
        return "data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'><rect width='10' height='10' rx='2' fill='%231b2d4b'/></svg>"

payload = json.dumps(data, separators=(",", ":"), ensure_ascii=False).replace("</", "<\\/")
OUT.write_text(HTML.replace("__LOGO__", logo_uri()).replace("__DATA__", payload), encoding="utf-8")
print(f"Done: {OUT}  ({OUT.stat().st_size/1e6:.1f} MB)  rows={len(d):,}  detail rows={len(det):,}")

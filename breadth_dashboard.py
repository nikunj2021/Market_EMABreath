#!/usr/bin/env python3
"""Build docs/index.html: Nifty 50 / Nifty 500 performance + EMA-breadth screener.

Inputs : Google Sheet EMASTATE (tabs Nifty50_EMA, Nifty500_EMA: Date, % >20/50/200 EMA),
         read via service account (GOOGLE_SERVICE_ACCOUNT_JSON + GSHEET_ID) or, if the sheet is
         shared "anyone with the link can view", via GSHEET_ID alone. Local EMASTATE.xlsx works for tests.
         index prices via yfinance (^NSEI, ^CRSLDX)
Output : docs/index.html (self-contained, static)
"""
import json, sys, os, datetime as dt
import numpy as np, pandas as pd

XLSX = os.environ.get("EMASTATE_XLSX")          # local file override (testing)
GSHEET_ID = os.environ.get("GSHEET_ID")          # the long id in the sheet URL
SA_JSON = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
OUT = os.environ.get("DASH_OUT", "docs/index.html")
INDICES = {  # name: (breadth sheet, yahoo ticker)
    "Nifty 50": ("Nifty50_EMA", "^NSEI"),
    "Nifty 500": ("Nifty500_EMA", "^CRSLDX"),
}
PRICE_FILE = os.environ.get("PRICE_CSV")  # optional offline override: Date,Nifty 50,Nifty 500


STATE_FILE = os.environ.get("STATE_FILE", "docs/state.json")
DASH_URL = os.environ.get("DASH_URL", "")
ICON = {"Extreme oversold": "🔴", "Oversold": "🟠", "Neutral": "⚪", "Overbought": "🟢"}


def send_telegram(text):
    token, chat = os.environ.get("TELEGRAM_BOT_TOKEN"), os.environ.get("TELEGRAM_CHAT_ID")
    if os.environ.get("TELEGRAM_DRY_RUN"):
        print("--- telegram (dry run) ---\n" + text)
        return True
    if not (token and chat):
        print("Telegram secrets not set, skipping alert", file=sys.stderr)
        return None
    import urllib.request, urllib.parse
    body = urllib.parse.urlencode({"chat_id": chat, "text": text, "disable_web_page_preview": "true"}).encode()
    try:
        urllib.request.urlopen(urllib.request.Request(
            f"https://api.telegram.org/bot{token}/sendMessage", data=body), timeout=30).read()
        return True
    except Exception as e:
        print(f"telegram send failed: {e}", file=sys.stderr)
        return False


def alerts(data):
    """Send a Telegram message every run: a daily summary, flagged when an index status changed."""
    try:
        old = json.load(open(STATE_FILE))
    except Exception:
        old = None
    new, lines, changed = {}, [], False
    for name, d in data["idx"].items():
        new[name] = d["status"]
        prev = (old or {}).get(name)
        if prev != d["status"]:
            changed = True
        c, r = d["cur"], d["ret"]
        arrow = f"{prev} -> {d['status']}" if prev and prev != d["status"] else d["status"]
        px = f"{d['last']:,.2f} ({r.get('1D', 0):+.2f}% 1D, {r.get('1M', 0):+.2f}% 1M)" if d["last"] else "price n/a"
        lines.append(f"{ICON[d['status']]} {name}: {arrow}\n   Close {px}\n"
                     f"   % above EMA  20: {c['e20']}  50: {c['e50']}  200: {c['e200']}\n"
                     f"   50 EMA breadth percentile {d['pct']['e50']} (oversold at {d['zones']['oversold']} or lower)")
    daily = os.environ.get("TELEGRAM_DAILY", "1") != "0"
    if changed or daily:
        day = data["idx"][next(iter(data["idx"]))]["breadth_date"]
        head = (f"NSE breadth ALERT: status changed, {day}" if changed and old is not None
                else f"NSE breadth daily summary, {day}")
        msg = head + "\n\n" + "\n\n".join(lines) + (f"\n\n{DASH_URL}" if DASH_URL else "")
        if send_telegram(msg) is False:
            return                      # keep old state so a status-change alert is retried next run
    os.makedirs(os.path.dirname(STATE_FILE) or ".", exist_ok=True)
    json.dump(new, open(STATE_FILE, "w"))


_cache = {}


def _to_dates(col):
    if pd.api.types.is_datetime64_any_dtype(col):
        return col.dt.normalize()
    num = pd.to_numeric(col, errors="coerce")
    num = num.where((num > 20000) & (num < 80000))              # plausible Sheets/Excel serials only
    d = pd.to_datetime(num, unit="D", origin="1899-12-30")
    txt = pd.to_datetime(col.where(num.isna()).astype(str), errors="coerce")
    return d.fillna(txt).dt.normalize()


def _raw(sheet):
    """Return the tab as a DataFrame with the header row as columns."""
    if sheet in _cache:
        return _cache[sheet]
    if XLSX:
        df = pd.read_excel(XLSX, sheet_name=sheet)
    elif GSHEET_ID and SA_JSON:
        import gspread
        from google.oauth2.service_account import Credentials
        creds = Credentials.from_service_account_info(
            json.loads(SA_JSON), scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
        ws = gspread.authorize(creds).open_by_key(GSHEET_ID).worksheet(sheet)
        rows = ws.get_all_values(value_render_option="UNFORMATTED_VALUE")
        df = pd.DataFrame(rows[1:], columns=rows[0])
    elif GSHEET_ID:
        from urllib.parse import quote
        url = f"https://docs.google.com/spreadsheets/d/{GSHEET_ID}/gviz/tq?tqx=out:csv&sheet={quote(sheet)}"
        df = pd.read_csv(url)
    else:
        sys.exit("Set GSHEET_ID (and GOOGLE_SERVICE_ACCOUNT_JSON if the sheet is private) or EMASTATE_XLSX")
    _cache[sheet] = df
    return df


def load_breadth(sheet):
    df = _raw(sheet).rename(columns=lambda c: str(c).strip())
    df = df.rename(columns={"% >20 EMA": "e20", "% >50 EMA": "e50", "% >200 EMA": "e200"})
    df["Date"] = _to_dates(df["Date"])
    for c in ("e20", "e50", "e200"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df = df.dropna(subset=["Date", "e20", "e50", "e200"])
    return (df.drop_duplicates("Date", keep="last").sort_values("Date")
              .set_index("Date")[["e20", "e50", "e200"]].astype(float))


def load_prices():
    if PRICE_FILE:
        p = pd.read_csv(PRICE_FILE, parse_dates=["Date"]).set_index("Date")
        return {k: p[k].dropna() for k in INDICES}
    import yfinance as yf
    out = {}
    for name, (_, tk) in INDICES.items():
        try:
            h = yf.download(tk, period="3y", interval="1d", auto_adjust=False, progress=False)
            s = h["Close"]
            s = s.iloc[:, 0] if hasattr(s, "columns") else s
            s.index = pd.to_datetime(s.index).tz_localize(None).normalize()
            out[name] = s.dropna()
        except Exception as e:
            print(f"price fetch failed for {name}: {e}", file=sys.stderr)
            out[name] = pd.Series(dtype=float)
    return out


def returns(px):
    if len(px) < 3:
        return {}
    last = px.iloc[-1]
    def r(n): return round((last / px.iloc[-1 - n] - 1) * 100, 2) if len(px) > n else None
    ytd_base = px[px.index < pd.Timestamp(px.index[-1].year, 1, 1)]
    return {"1D": r(1), "1W": r(5), "1M": r(21), "3M": r(63), "6M": r(126), "1Y": r(252),
            "YTD": round((last / ytd_base.iloc[-1] - 1) * 100, 2) if len(ytd_base) else None,
            "From 52W high": round((last / px.iloc[-252:].max() - 1) * 100, 2)}


def forward_stats(b, thr):
    m = (b.e50 <= thr).values
    idx = np.where(m)[0]
    rows = []
    for n in (5, 10, 20):
        v = [b.e50.iloc[i + n] - b.e50.iloc[i] for i in idx if i + n < len(b)]
        rows.append({"days": n, "n": len(v),
                     "median": round(float(np.median(v)), 1) if v else None,
                     "up": round(float(np.mean(np.array(v) > 0) * 100)) if v else None})
    return rows


def build():
    prices = load_prices()
    data = {"updated": dt.datetime.now().strftime("%d %b %Y %H:%M"), "idx": {}}
    for name, (sheet, _) in INDICES.items():
        b = load_breadth(sheet)
        px = prices.get(name, pd.Series(dtype=float))
        z = {"oversold": round(float(b.e50.quantile(.10)), 1),
             "extreme": round(float(b.e50.quantile(.03)), 1),
             "overbought": round(float(b.e50.quantile(.90)), 1),
             "e200_oversold": round(float(b.e200.quantile(.10)), 1)}
        cur = b.iloc[-1]
        pct = {k: round(float((b[k] <= cur[k]).mean() * 100), 1) for k in ("e20", "e50", "e200")}
        status = ("Extreme oversold" if cur.e50 <= z["extreme"] else "Oversold" if cur.e50 <= z["oversold"]
                  else "Overbought" if cur.e50 >= z["overbought"] else "Neutral")
        pxs = px.iloc[-520:]
        d20 = px.ewm(span=20, adjust=False).mean().iloc[-520:]
        d50 = px.ewm(span=50, adjust=False).mean().iloc[-520:]
        d200 = px.ewm(span=200, adjust=False).mean().iloc[-520:]
        data["idx"][name] = {
            "status": status, "zones": z, "pct": pct,
            "breadth_date": b.index[-1].strftime("%Y-%m-%d"),
            "cur": {k: round(float(cur[k]), 1) for k in ("e20", "e50", "e200")},
            "b": {"d": [d.strftime("%Y-%m-%d") for d in b.index],
                  **{k: b[k].round(1).tolist() for k in ("e20", "e50", "e200")}},
            "p": {"d": [d.strftime("%Y-%m-%d") for d in pxs.index], "c": pxs.round(2).tolist(),
                  "m20": [None if np.isnan(v) else round(float(v), 2) for v in d20],
                  "m50": [None if np.isnan(v) else round(float(v), 2) for v in d50],
                  "m200": [None if np.isnan(v) else round(float(v), 2) for v in d200]},
            "last": round(float(px.iloc[-1]), 2) if len(px) else None,
            "last_date": px.index[-1].strftime("%Y-%m-%d") if len(px) else None,
            "ret": returns(px), "fwd": forward_stats(b, z["oversold"]),
        }
    os.makedirs(os.path.dirname(OUT) or ".", exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(TEMPLATE.replace("__DATA__", json.dumps(data, separators=(",", ":"))))
    print("wrote", OUT, {k: v["status"] for k, v in data["idx"].items()})
    alerts(data)


TEMPLATE = r'''<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>NSE index screener</title>
<style>
:root{--bg:#f3f5f7;--ink:#14202b;--mut:#5c6b78;--grid:#d9e0e6;--card:#fff;--c20:#0f8a7e;--c50:#d9822b;--c200:#4b4fb8;--px:#14202b;--ov:rgba(214,69,65,.10);--ex:rgba(214,69,65,.20);--sh:rgba(214,69,65,.09);--red:#c93a36;--grn:#1b8a4b;--amb:#b7791f}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--bg:#10171d;--ink:#e6edf3;--mut:#93a1ad;--grid:#27333d;--card:#17212a;--c20:#3ec9b9;--c50:#f0a050;--c200:#8f94ff;--px:#e6edf3;--ov:rgba(255,107,100,.12);--ex:rgba(255,107,100,.24);--sh:rgba(255,107,100,.12);--red:#ff6b64;--grn:#4cd68a;--amb:#f0b84a}}
*{box-sizing:border-box}
html{scroll-padding-top:env(safe-area-inset-top,0px)}
body{margin:0;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px);background:var(--bg);color:var(--ink);font:15px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif}
main{max-width:1040px;margin:0 auto;padding:20px 16px 32px}
h1{font:600 24px/1.2 Georgia,"Times New Roman",serif;margin:0 0 4px}
h2{font:600 16px/1.3 system-ui,sans-serif;margin:0 0 8px}
.sub{color:var(--mut);margin:0 0 16px}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:12px;margin-bottom:16px}
.card{background:var(--card);border:1px solid var(--grid);border-radius:8px;padding:14px}
.head{display:flex;justify-content:space-between;align-items:baseline;gap:8px;flex-wrap:wrap}
.px{font:600 26px/1.1 Georgia,serif;font-variant-numeric:tabular-nums}
.badge{font-weight:600;font-size:13px;padding:2px 9px;border-radius:99px;border:1.5px solid currentColor}
.rets{display:grid;grid-template-columns:repeat(4,1fr);gap:6px;margin-top:12px}
.rets div{border:1px solid var(--grid);border-radius:6px;padding:4px 8px;font-size:12px;color:var(--mut)}
.rets b{display:block;font-size:15px;font-variant-numeric:tabular-nums}
.up{color:var(--grn)}.dn{color:var(--red)}
.meter{display:grid;grid-template-columns:auto 1fr auto;gap:4px 10px;align-items:center;margin-top:12px;font-size:13px;color:var(--mut)}
.track{position:relative;height:8px;border-radius:4px;background:linear-gradient(90deg,var(--ex),var(--ov) 20%,var(--grid) 20%,var(--grid) 80%,rgba(27,138,75,.18) 80%)}
.dot{position:absolute;top:-3px;width:14px;height:14px;border-radius:50%;border:2px solid var(--card);transform:translateX(-50%)}
.bar{display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap}
button{font:inherit;color:var(--ink);background:var(--card);border:1px solid var(--grid);border-radius:6px;padding:6px 12px;cursor:pointer}
button[aria-pressed="true"]{border-color:var(--ink);font-weight:600}
button:focus-visible{outline:2px solid var(--c200);outline-offset:2px}
.sw{display:inline-block;width:10px;height:10px;border-radius:2px;margin-right:6px}
svg{width:100%;height:auto;display:block;touch-action:none}
.read{min-height:22px;color:var(--mut);font-size:13px;font-variant-numeric:tabular-nums;margin-top:4px}
table{border-collapse:collapse;width:100%;font-size:13px;font-variant-numeric:tabular-nums}
th,td{text-align:right;padding:5px 8px;border-bottom:1px solid var(--grid)}th:first-child,td:first-child{text-align:left}
.scroll{overflow-x:auto}
.note{color:var(--mut);font-size:13px;margin-top:14px}
</style></head><body><main>
<h1>NSE index screener</h1>
<p class="sub" id="upd"></p>
<div class="cards" id="cards"></div>
<div class="bar" id="tabs"></div>
<div class="card"><h2 id="t1"></h2><div class="bar" id="pser"></div><svg id="s1" viewBox="0 0 900 300" role="img" aria-label="Index price with 20, 50 and 200 EMA"></svg><div class="read" id="r1"></div></div>
<div class="card" style="margin-top:12px"><h2>Breadth: share of stocks above EMA</h2><div class="bar" id="ser"></div><svg id="s2" viewBox="0 0 900 300" role="img" aria-label="Breadth lines with oversold bands"></svg><div class="read" id="r2"></div></div>
<div class="card" style="margin-top:12px"><h2 id="t3"></h2><div class="scroll"><table id="fwd"></table></div></div>
<p class="note">Status uses this index's own history: oversold is the bottom 10% of % above 50 EMA readings, extreme is the bottom 3%, overbought is the top 10%. Red shading on the price chart marks days breadth was oversold. Breadth history is short, so zones describe this sample only. Not investment advice.</p>
</main>
<script>
const DATA=__DATA__;
const NS='http://www.w3.org/2000/svg',css=v=>getComputedStyle(document.documentElement).getPropertyValue(v).trim();
const S=[['e20','% above 20 EMA','--c20'],['e50','% above 50 EMA','--c50'],['e200','% above 200 EMA','--c200']];
let cur=Object.keys(DATA.idx)[0],on={e20:true,e50:true,e200:true},pon={c:true,m20:true,m50:true,m200:true};
const P=[['c','Close','--px',1.8],['m20','20 EMA','--c20',1.2],['m50','50 EMA','--c50',1.2],['m200','200 EMA','--c200',1.2]];
const f=(v,d=2)=>v==null?'n/a':v.toLocaleString('en-IN',{minimumFractionDigits:d,maximumFractionDigits:d});
const cls=v=>v==null?'':v>=0?'up':'dn';
const sc=s=>s.startsWith('Extreme')?'--red':s==='Oversold'?'--amb':s==='Overbought'?'--grn':'--mut';
function el(svg,n,a){const e=document.createElementNS(NS,n);for(const k in a)e.setAttribute(k,a[k]);svg.appendChild(e);return e}
function chart(svg,rd,o){
 svg.innerHTML='';const W=900,H=300,L=46,R=10,T=10,B=24,t0=o.t0,t1=o.t1;
 const x=t=>L+(W-L-R)*(t-t0)/(t1-t0),y=v=>T+(H-T-B)*(1-(v-o.min)/(o.max-o.min));
 (o.bands||[]).forEach(b=>el(svg,'rect',{x:L,y:y(b[1]),width:W-L-R,height:y(b[0])-y(b[1]),fill:css(b[2])}));
 (o.shade||[]).forEach(s=>el(svg,'rect',{x:x(s[0]),y:T,width:Math.max(1.5,x(s[1])-x(s[0])),height:H-T-B,fill:css('--sh')}));
 for(let i=0;i<=4;i++){const v=o.min+(o.max-o.min)*i/4;el(svg,'line',{x1:L,x2:W-R,y1:y(v),y2:y(v),stroke:css('--grid')});
  const t=el(svg,'text',{x:L-6,y:y(v)+4,'text-anchor':'end','font-size':11,fill:css('--mut')});t.textContent=o.fmt(v)}
 (o.hl||[]).forEach(v=>el(svg,'line',{x1:L,x2:W-R,y1:y(v),y2:y(v),stroke:css('--c200'),'stroke-dasharray':'5 4',opacity:.8}));
 let last='';for(let t=t0;t<=t1;t+=864e5*30){const d=new Date(t),k=d.getFullYear()*12+d.getMonth();if(d.getMonth()%3===0&&k!==last){last=k;
  const e=el(svg,'text',{x:x(t),y:H-6,'font-size':11,fill:css('--mut')});e.textContent=d.toLocaleString('en',{month:'short',year:'2-digit'})}}
 o.series.forEach(s=>{let p='',g=true;s.v.forEach((v,i)=>{if(v==null){g=true;return}p+=(g?'M':'L')+x(s.t[i]).toFixed(1)+' '+y(v).toFixed(1);g=false});
  el(svg,'path',{d:p,fill:'none',stroke:css(s.c),'stroke-width':s.w||1.6,'stroke-linejoin':'round'})});
 const cl=el(svg,'line',{x1:0,x2:0,y1:T,y2:H-B,stroke:css('--ink'),opacity:0});
 const mv=ev=>{const r=svg.getBoundingClientRect(),px=((ev.touches?ev.touches[0].clientX:ev.clientX)-r.left)/r.width*W;
  const t=t0+(px-L)/(W-L-R)*(t1-t0);cl.setAttribute('x1',x(t));cl.setAttribute('x2',x(t));cl.setAttribute('opacity',.5);
  const s0=o.series[0];let bi=0,bd=1e18;s0.t.forEach((tt,i)=>{const d=Math.abs(tt-t);if(d<bd){bd=d;bi=i}});
  rd.textContent=new Date(s0.t[bi]).toISOString().slice(0,10)+'   '+o.series.map(s=>{let j=0,m=1e18;s.t.forEach((tt,i)=>{const d=Math.abs(tt-t);if(d<m){m=d;j=i}});return s.n+' '+o.fmt(s.v[j])}).join('   ')};
 svg.onmousemove=mv;svg.ontouchmove=mv;svg.ontouchstart=mv}
const T=a=>a.map(d=>Date.parse(d));
function render(){
 const _d=DATA.idx[cur];document.getElementById('upd').textContent='Updated '+DATA.updated+' (breadth through '+_d.breadth_date+(_d.last_date?', price through '+_d.last_date+(_d.breadth_date<_d.last_date?'. Breadth is a day behind price, the sheet may not have updated yet':''):', price data unavailable')+')';
 const cards=document.getElementById('cards');cards.innerHTML='';
 for(const [name,d] of Object.entries(DATA.idx)){
  const c=document.createElement('div');c.className='card';
  const rets=['1D','1W','1M','3M','6M','1Y','YTD','From 52W high'].map(k=>{const v=d.ret[k];return '<div>'+k+'<b class="'+cls(v)+'">'+(v==null?'n/a':(v>0?'+':'')+v.toFixed(2)+'%')+'</b></div>'}).join('');
  const m=[['e20','20 EMA'],['e50','50 EMA'],['e200','200 EMA']].map(([k,l])=>'<span>'+l+'</span><div class="track"><div class="dot" style="left:'+Math.max(2,Math.min(98,d.pct[k]))+'%;background:var('+sc(d.status)+')"></div></div><span>'+d.cur[k]+'% ('+d.pct[k]+'th pct)</span>').join('');
  c.innerHTML='<div class="head"><h2>'+name+'</h2><span class="badge" style="color:var('+sc(d.status)+')">'+d.status+'</span></div><div class="px">'+f(d.last)+'</div><div class="rets">'+(Object.keys(d.ret).length?rets:'<div>Price data unavailable</div>')+'</div><div class="meter">'+m+'</div>';
  cards.appendChild(c)}
 const tabs=document.getElementById('tabs');tabs.innerHTML='';
 Object.keys(DATA.idx).forEach(k=>{const b=document.createElement('button');b.textContent=k;b.setAttribute('aria-pressed',k===cur);b.onclick=()=>{cur=k;render()};tabs.appendChild(b)});
 const ser=document.getElementById('ser');ser.innerHTML='';
 S.forEach(([k,l,c])=>{const b=document.createElement('button');b.setAttribute('aria-pressed',on[k]);b.innerHTML='<span class="sw" style="background:var('+c+')"></span>'+l;b.onclick=()=>{on[k]=!on[k];render()};ser.appendChild(b)});
 const ps=document.getElementById('pser');ps.innerHTML='';
 P.forEach(([k,l,c])=>{const b=document.createElement('button');b.setAttribute('aria-pressed',pon[k]);b.innerHTML='<span class="sw" style="background:var('+c+')"></span>'+l;b.onclick=()=>{pon[k]=!pon[k];render()};ps.appendChild(b)});
 const d=DATA.idx[cur],bt=T(d.b.d),pt=T(d.p.d);
 document.getElementById('t1').textContent=cur+' price with 20, 50 and 200 EMA';
 document.getElementById('t3').textContent='After breadth got this oversold ('+d.zones.oversold+'% above 50 EMA or lower): change in % above 50 EMA';
 const t0=pt.length?pt[0]:bt[0],t1=Math.max(bt[bt.length-1],pt.length?pt[pt.length-1]:0);
 // shade price chart where breadth oversold
 const sh=[];let st=null;d.b.e50.forEach((v,i)=>{if(v<=d.zones.oversold){if(st==null)st=bt[i]}else if(st!=null){sh.push([st,bt[i-1]+864e5]);st=null}});if(st!=null)sh.push([st,bt[bt.length-1]+864e5]);
 const r1=document.getElementById('r1'),s1=document.getElementById('s1');
 if(pt.length){const act=P.filter(p=>pon[p[0]]);const all=[].concat(...(act.length?act:P.slice(0,1)).map(p=>d.p[p[0]].filter(v=>v!=null))),mn=Math.min(...all),mx=Math.max(...all),pad=(mx-mn)*.05;
  chart(s1,r1,{t0,t1,min:mn-pad,max:mx+pad,fmt:v=>f(v,0),shade:sh,series:act.length?act.map(p=>({n:p[1],t:pt,v:d.p[p[0]],c:p[2],w:p[3]})):[{n:'Close',t:pt,v:d.p.c,c:'--px',w:0}]})}
 else{s1.innerHTML='';r1.textContent='Price data unavailable for this run.'}
 const r2=document.getElementById('r2');
 chart(document.getElementById('s2'),r2,{t0:bt[0],t1:bt[bt.length-1],min:0,max:100,fmt:v=>Math.round(v)+'%',
  bands:[[0,d.zones.extreme,'--ex'],[d.zones.extreme,d.zones.oversold,'--ov']],hl:[d.zones.e200_oversold],
  series:S.filter(s=>on[s[0]]).map(s=>({n:s[1].replace('% above ','').replace(' EMA',''),t:bt,v:d.b[s[0]],c:s[2]}))});
 document.getElementById('fwd').innerHTML='<tr><th>Days later</th><th>Median change</th><th>Higher than at signal</th><th>Signal days</th></tr>'+d.fwd.map(r=>'<tr><td>'+r.days+'</td><td>'+(r.median==null?'n/a':(r.median>0?'+':'')+r.median+' pts')+'</td><td>'+(r.up==null?'n/a':r.up+'%')+'</td><td>'+r.n+'</td></tr>').join('');
}
render();matchMedia('(prefers-color-scheme: dark)').addEventListener('change',render);
</script></body></html>
'''

if __name__ == "__main__":
    build()

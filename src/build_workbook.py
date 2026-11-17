import random, math, os, datetime as dt
PUBLIC = os.environ.get('PUBLIC')=='1'
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule

random.seed(30)
SEGS = ["Large Ag", "Small Ag & Turf", "Construction & Forestry", "Parts & Service"]
REGIONS = ["Midwest", "Plains", "South", "West"]
FYS = ["FY2024", "FY2025", "FY2026"]

# ---------- dimensions ----------
products = [
 ("P101","Row-Crop Tractor 250HP","Tractors","Large Ag",310000,.78,5),
 ("P102","Row-Crop Tractor 400HP","Tractors","Large Ag",520000,.78,4),
 ("P103","4WD Tractor 600HP","Tractors","Large Ag",720000,.78,2),
 ("P104","Combine Harvester","Harvesting","Large Ag",780000,.78,3),
 ("P105","Planter 24-Row","Planting","Large Ag",330000,.78,3),
 ("P106","Self-Propelled Sprayer","Application","Large Ag",560000,.78,3),
 ("P201","Utility Tractor 75HP","Tractors","Small Ag & Turf",58000,.76,5),
 ("P202","Utility Tractor 120HP","Tractors","Small Ag & Turf",105000,.76,3),
 ("P203","Compact Tractor 25HP","Compact","Small Ag & Turf",19000,.76,8),
 ("P204","Compact Tractor 40HP","Compact","Small Ag & Turf",32000,.76,6),
 ("P205","Zero-Turn Mower","Turf","Small Ag & Turf",9500,.76,9),
 ("P206","Utility Vehicle","Turf","Small Ag & Turf",17000,.76,6),
 ("P301","Excavator 20-Ton","Earthmoving","Construction & Forestry",245000,.79,4),
 ("P302","Wheel Loader","Earthmoving","Construction & Forestry",285000,.79,3),
 ("P303","Skid Steer","Compact Construction","Construction & Forestry",62000,.79,7),
 ("P304","Compact Track Loader","Compact Construction","Construction & Forestry",78000,.79,6),
 ("P305","Backhoe Loader","Earthmoving","Construction & Forestry",128000,.79,3),
 ("P306","Motor Grader","Roadbuilding","Construction & Forestry",390000,.79,1.5),
 ("P401","Parts Order","Parts","Parts & Service",3800,.58,10),
 ("P402","Service & Maintenance Plan","Service","Parts & Service",6500,.55,5),
 ("P403","Precision Ag Technology Upgrade","Technology","Parts & Service",12000,.45,2),
]
dealers = [  # id, name, state, region, size, mix(LA,SA,CF,PS), extra discount, late prob
 ("D01","Des Moines","IA","Midwest",1.3,(1.2,1,.8,1),0,.10),
 ("D02","Peoria","IL","Midwest",1.1,(1.2,1,.9,1),0,.12),
 ("D03","Lafayette","IN","Midwest",.9,(1.1,1.1,.8,1),0,.11),
 ("D04","Grand Island","NE","Plains",1.1,(1.5,.7,.6,1),0,.13),
 ("D05","Salina","KS","Plains",.9,(1.5,.7,.6,1),.06,.14),
 ("D06","Fargo","ND","Plains",1.0,(1.6,.6,.5,1),0,.12),
 ("D07","Lubbock","TX","South",.9,(1.0,1,1.1,1),0,.13),
 ("D08","Conroe","TX","South",1.1,(.4,1.3,1.7,1),0,.15),
 ("D09","Tifton","GA","South",.8,(.8,1.3,1.0,1),.06,.12),
 ("D10","Fresno","CA","West",1.0,(.7,1.1,1.3,1),0,.36),
 ("D11","Yakima","WA","West",.7,(.6,1.2,1.1,1),0,.38),
 ("D12","Greeley","CO","West",.8,(.9,1,1.3,1),0,.33),
]
season = {  # by calendar month 1..12
 "Large Ag":[1.0,1.25,1.4,1.2,.9,.8,.9,1.15,1.2,.95,.8,1.1],
 "Small Ag & Turf":[.7,.9,1.3,1.5,1.5,1.3,1.0,.9,.9,.8,.6,.6],
 "Construction & Forestry":[.7,.8,1.1,1.2,1.25,1.2,1.1,1.1,1.05,1.0,.8,.7],
 "Parts & Service":[.85,.9,1.05,1.2,1.2,1.0,.95,1.0,1.2,1.15,.9,.8],
}
levels = {"Large Ag":[1.0,.90,.72,.66],"Small Ag & Turf":[1.0,.97,.93,1.05],
          "Construction & Forestry":[1.0,.98,.97,1.18],"Parts & Service":[1.0,1.04,1.09,1.16]}
base_orders = {"Large Ag":28,"Small Ag & Turf":45,"Construction & Forestry":25,"Parts & Service":60}
disc_base = {"Large Ag":[.06,.085,.11],"Small Ag & Turf":[.05,.055,.05],
             "Construction & Forestry":[.06,.065,.06],"Parts & Service":[.02,.02,.02]}
cost_creep = [1.0,1.02,1.045]
cust = {"Large Ag":[("Large Farm Operation",.7),("Family Farm",.25),("Custom Operator",.05)],
        "Small Ag & Turf":[("Family Farm",.35),("Homeowner / Hobby",.35),("Landscaper / Municipal",.3)],
        "Construction & Forestry":[("Contractor",.75),("Landscaper / Municipal",.25)],
        "Parts & Service":[("Large Farm Operation",.3),("Family Farm",.3),("Contractor",.25),("Landscaper / Municipal",.1),("Homeowner / Hobby",.05)]}
lead = {"Large Ag":(30,60),"Small Ag & Turf":(5,21),"Construction & Forestry":(14,45),"Parts & Service":(1,7)}
fin_p = {"Large Ag":.72,"Small Ag & Turf":.5,"Construction & Forestry":.6,"Parts & Service":.05}

def poisson(lam):
    Lm = math.exp(-lam) if lam < 30 else None
    if Lm is None: return max(0, int(round(random.gauss(lam, math.sqrt(lam)))))
    k, p = 0, 1.0
    while True:
        p *= random.random()
        if p <= Lm: return k
        k += 1
def pick(items, weights): return random.choices(items, weights=weights)[0]

rows = []
segidx = {s:i for i,s in enumerate(SEGS)}
for t in range(36):
    y, m = 2023 + (10 + t)//12, (10 + t) % 12 + 1
    fy = t // 12; frac = (t % 12 + .5)/12
    for s in SEGS:
        lv = levels[s][fy] + (levels[s][fy+1]-levels[s][fy])*frac
        n = max(1, round(base_orders[s]*season[s][m-1]*lv*random.gauss(1,.05)))
        prods = [p for p in products if p[3]==s]
        for _ in range(n):
            d = pick(dealers, [x[4]*x[5][segidx[s]] for x in dealers])
            w = [p[6] for p in prods]
            if s=="Parts & Service": w = [10,5,2*(1+t/14)]
            p = pick(prods, w)
            day = random.randint(1, 28 if m==2 else 30)
            od = dt.date(y,m,day)
            if s=="Parts & Service": units = random.randint(1,8) if p[0]=="P401" else random.randint(1,3)
            elif p[4] < 40000: units = random.choices([1,2,3],[.75,.18,.07])[0]
            else: units = random.choices([1,2],[.92,.08])[0]
            disc = disc_base[s][fy] + (d[6] if s!="Parts & Service" else 0) + random.gauss(0,.028)
            disc = round(min(.25,max(0,disc))*200)/200
            ucost = round(p[4]*p[5]*cost_creep[fy]*random.uniform(.985,1.015), 0)
            prom = od + dt.timedelta(days=random.randint(*lead[s]))
            lp = d[7] + (.08 if s=="Construction & Forestry" else 0)
            late = random.randint(1,21) if random.random()<lp else 0
            c = pick([c[0] for c in cust[s]],[c[1] for c in cust[s]])
            fin = "Y" if random.random()<fin_p[s] else "N"
            rows.append((od,d[0],p[0],c,units,disc,ucost,prom,prom+dt.timedelta(days=late),fin,s,p))
rows.sort(key=lambda r:r[0])
N = len(rows); LAST = N+1
print("rows", N)

# python-side numbers used only to set sample inputs (inventory snapshot, targets)
def fyi(d): return d.year + (1 if d.month>=11 else 0) - 2024
units26 = {p[0]:0 for p in products}; net = {(s,f):0 for s in SEGS for f in range(3)}
for r in rows:
    f = fyi(r[0]); net[(r[10],f)] += r[4]*r[11][4]*(1-r[5])
    if f==2: units26[r[2]] += r[4]
mos_plant = {"P101":4.5,"P102":7.2,"P103":8.0,"P104":9.5,"P105":5.0,"P106":6.8,"P201":3.2,"P202":3.8,"P203":1.1,
             "P204":1.3,"P205":2.6,"P206":3.0,"P301":3.4,"P302":4.0,"P303":1.2,"P304":1.6,"P305":3.1,"P306":4.8}
onhand = {k: max(1, round(v*units26[k]/12)) for k,v in mos_plant.items()}
tg = {"Large Ag":[1.02,.95,.95],"Small Ag & Turf":[1.02,1.03,1.05],"Construction & Forestry":[1.02,1.03,1.08],"Parts & Service":[1.02,1.03,1.06]}
targets = {s:[round(net[(s,0)]*tg[s][0],-5), round(net[(s,0)]*tg[s][1],-5), round(net[(s,1)]*tg[s][2],-5)] for s in SEGS}

# ---------- styles ----------
F = "Arial"
NAVY, TEAL, LIGHT, GREY = "1F3A4D", "2E7D7A", "E8F1F2", "F3F4F6"
def font(**k): return Font(name=F, **{**dict(size=10), **k})
thin = Side(style="thin", color="C9CED4"); BORDER = Border(left=thin,right=thin,top=thin,bottom=thin)
HFILL = PatternFill("solid", fgColor=NAVY); SFILL = PatternFill("solid", fgColor=LIGHT)
YFILL = PatternFill("solid", fgColor="FFFF00"); GFILL = PatternFill("solid", fgColor=GREY)
MONEY, MM, PCT, INT, DATE = '$#,##0', '$#,##0.0,,"M"', '0.0%', '#,##0', 'mmm d, yyyy'
wb = Workbook()
def sheet(name, tab=None, widths=None):
    ws = wb.create_sheet(name); ws.sheet_view.showGridLines = False
    if tab: ws.sheet_properties.tabColor = tab
    for i,w in enumerate(widths or [], 1): ws.column_dimensions[L(i)].width = w
    return ws
def title(ws, text, sub=None):
    ws["A1"] = text; ws["A1"].font = font(size=16, bold=True, color=NAVY)
    if sub: ws["A2"] = sub; ws["A2"].font = font(italic=True, color="5B6770")
def section(ws, row, text, ncol=8):
    for c in range(1, ncol+1): ws.cell(row, c).fill = PatternFill("solid", fgColor=TEAL)
    ws.cell(row,1,text).font = font(size=11, bold=True, color="FFFFFF")
def header(ws, row, labels, col=1):
    for i,l in enumerate(labels):
        c = ws.cell(row, col+i, l); c.font = font(bold=True, color="FFFFFF"); c.fill = HFILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = BORDER
    ws.row_dimensions[row].height = 30
def put(ws, row, col, v, fmt=None, bold=False, color=None, fill=None, wrap=False, align=None):
    c = ws.cell(row, col, v); c.font = font(bold=bold, color=color); c.border = BORDER
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    if wrap or align: c.alignment = Alignment(wrap_text=wrap, vertical="top", horizontal=align)
    return c
def inp(ws, row, col, v, fmt=None, key=False):
    c = put(ws,row,col,v,fmt,color="0000FF"); 
    if key: c.fill = YFILL
    return c
def note(ws, row, text, col=1, color="5B6770"):
    c = ws.cell(row, col, text); c.font = font(italic=True, color=color, size=9); return c
def chart_fix(ch, title_, w=16, h=8):
    ch.title = title_; ch.roundedCorners = False; ch.width = w; ch.height = h
    if ch.legend: ch.legend.overlay = False; ch.legend.position = "b"
    ch.x_axis.delete = False; ch.y_axis.delete = False
    return ch

wb.remove(wb.active)
ws0 = sheet("Start_Here", NAVY, [3,30,95])
dash = sheet("Dashboard", NAVY, [2]+[14]*14)
fin = sheet("Findings", NAVY, [4,62,26,48,18,20,52])
sa = sheet("Sales_Analysis", TEAL, [26,24,16,15,15,15,13,13,13,13,13,13,15,13])
ops = sheet("Ops_Analysis", TEAL, [10,30,24,11,13,15,12,12,12,16,11,15,12,16])
st = sheet("Statistics", TEAL, [46,18,18,18,18,40])
fc = sheet("Forecast", TEAL, [26,16,16,13,13,14,16,16,16])
wi = sheet("What_If", TEAL, [34,18,18,18,18,18,18])
ig = sheet("Interview_Guide", "B45F06", [3,28,110])
mg = sheet("Macro_Guide", "B45F06", [3,30,34,90])
fm = sheet("Forecast_Model", "7F7F7F")
sd = sheet("Sales_Data", "7F7F7F", [11,13,9,10,10,21,24,24,7,12,10,13,13,11,13,13,9,13,13,9,9,10,12,9])
dp = sheet("Dim_Product", "7F7F7F", [11,32,22,24,13,13])
dd = sheet("Dim_Dealer", "7F7F7F", [10,34,8,12])
tgs = sheet("Targets", "7F7F7F", [26,16,16,16])

# ---------- Dim sheets ----------
title(dp,"Product master","Sample catalogue for a fictional equipment dealer network. Blue = hardcoded sample input.")
header(dp,4,["Product ID","Product","Category","Segment","List Price ($)","Std Unit Cost ($)"])
for i,p in enumerate(products):
    r=5+i; put(dp,r,1,p[0]); put(dp,r,2,p[1]); put(dp,r,3,p[2]); put(dp,r,4,p[3])
    inp(dp,r,5,p[4],MONEY); inp(dp,r,6,round(p[4]*p[5]),MONEY)
PR0,PR1 = 5, 4+len(products)
title(dd,"Dealer master","12 fictional dealer locations in 4 regions.")
header(dd,4,["Dealer ID","Dealer","State","Region"])
for i,d in enumerate(dealers):
    r=5+i; put(dd,r,1,d[0]); put(dd,r,2,f"Heartland Equipment - {d[1]}"); put(dd,r,3,d[2]); put(dd,r,4,d[3])
DR0,DR1 = 5, 4+len(dealers)
title(tgs,"Annual net sales targets by segment ($)","Sample plan numbers (hardcoded inputs, set at network level). Edit to test attainment.")
header(tgs,4,["Segment"]+FYS)
for i,s in enumerate(SEGS):
    put(tgs,5+i,1,s)
    for j in range(3): inp(tgs,5+i,2+j,targets[s][j],MONEY)
put(tgs,9,1,"Total",bold=True)
for j in range(3): put(tgs,9,2+j,f"=SUM({L(2+j)}5:{L(2+j)}8)",MONEY,bold=True)

# ---------- Sales_Data ----------
cols = ["Order ID","Order Date","Dealer ID","Region","Product ID","Category","Segment","Customer Type","Units","List Price",
        "Discount %","List Value","Net Sales","Unit Cost","COGS","Gross Margin","GM %","Promised Date","Delivered Date",
        "Days Late","On Time","Fiscal Year","Month","Financed"]
for i,c in enumerate(cols,1):
    x=sd.cell(1,i,c); x.font=font(bold=True,color="FFFFFF"); x.fill=HFILL; x.alignment=Alignment(horizontal="center",wrap_text=True)
dl = f"Dim_Dealer!$A${DR0}:$A${DR1}"; pl = f"Dim_Product!$A${PR0}:$A${PR1}"
def pidx(col,r): return f"=INDEX(Dim_Product!${col}${PR0}:${col}${PR1},MATCH(E{r},{pl},0))"
bf = font()
for i,rw in enumerate(rows):
    r=i+2
    vals = [f"SO-{100000+i}", rw[0], rw[1], f"=INDEX(Dim_Dealer!$D${DR0}:$D${DR1},MATCH(C{r},{dl},0))", rw[2],
            pidx("C",r), pidx("D",r), rw[3], rw[4], pidx("E",r), rw[5], f"=I{r}*J{r}", f"=L{r}*(1-K{r})", rw[6],
            f"=I{r}*N{r}", f"=M{r}-O{r}", f"=IF(M{r}=0,0,P{r}/M{r})", rw[7], rw[8], f"=MAX(0,S{r}-R{r})", f"=IF(T{r}=0,1,0)",
            f'="FY"&(YEAR(B{r})+IF(MONTH(B{r})>=11,1,0))', f"=DATE(YEAR(B{r}),MONTH(B{r}),1)", rw[9]]
    for j,v in enumerate(vals,1):
        c=sd.cell(r,j,v); c.font=bf
    for j in (2,18,19,23): sd.cell(r,j).number_format = DATE if j!=23 else 'mmm yyyy'
    for j in (10,12,13,14,15,16): sd.cell(r,j).number_format = MONEY
    for j in (11,17): sd.cell(r,j).number_format = PCT
sd.freeze_panes="A2"
tab = Table(displayName="tblSales", ref=f"A1:X{LAST}"); tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
sd.add_table(tab)
names = {"s_Date":"B","s_Dealer":"C","s_Region":"D","s_Product":"E","s_Segment":"G","s_Cust":"H","s_Units":"I","s_Disc":"K",
         "s_List":"L","s_Net":"M","s_COGS":"O","s_GM":"P","s_GMpct":"Q","s_DaysLate":"T","s_OnTime":"U","s_FY":"V","s_Month":"W"}
for n,c in names.items(): wb.defined_names[n] = DefinedName(n, attr_text=f"Sales_Data!${c}$2:${c}${LAST}")

# ---------- Forecast_Model ----------
fm.sheet_view.showGridLines = True
title(fm,"Forecast engine (calculation sheet)","Seasonal index x linear trend, per segment. Full fit uses 36 months; backtest (bt) fits on FY2024-25 and predicts FY2026.")
R0,R36,R48 = 6,41,53
fm["A5"]="t"; fm["B5"]="Month"; fm["C5"]="Fiscal Year"; fm["D5"]="Fiscal Month #"
fm["C2"]="Trend slope"; fm["C3"]="Trend intercept"
tA=f"$A${R0}:$A${R36}"; tBT=f"$A${R0}:$A${R0+23}"
for k in range(48):
    r=R0+k; fm.cell(r,1,k+1)
    fm.cell(r,2, dt.date(2023,11,1) if k==0 else f"=EDATE(B{r-1},1)").number_format='mmm yyyy'
    fm.cell(r,3,f'="FY"&(YEAR(B{r})+IF(MONTH(B{r})>=11,1,0))'); fm.cell(r,4,f"=MOD(MONTH(B{r})-11,12)+1")
blk = {}
sub = ["Actual","Seasonal idx","Deseasonalized","Model","Abs % err","Seasonal idx (bt)","Deseason (bt)","Model (bt)","Abs % err (bt)"]
for si,s in enumerate(SEGS):
    c0 = 5+si*9; cl=[L(c0+j) for j in range(9)]; blk[s]=cl
    fm.cell(4,c0,s).font=font(bold=True)
    for j,h in enumerate(sub): fm.cell(5,c0+j,h)
    a,ix,de,mo,ape,ixb,deb,mob,apeb = cl
    fm[f"{mo}2"]=f"=SLOPE({de}{R0}:{de}{R36},{tA})"; fm[f"{mo}3"]=f"=INTERCEPT({de}{R0}:{de}{R36},{tA})"
    fm[f"{mob}2"]=f"=SLOPE({deb}{R0}:{deb}{R0+23},{tBT})"; fm[f"{mob}3"]=f"=INTERCEPT({deb}{R0}:{deb}{R0+23},{tBT})"
    for k in range(48):
        r=R0+k
        fm[f"{ix}{r}"]=f"=AVERAGEIFS({a}${R0}:{a}${R36},$D${R0}:$D${R36},$D{r})/AVERAGE({a}${R0}:{a}${R36})"
        fm[f"{mo}{r}"]=f"=({mo}$3+{mo}$2*$A{r})*{ix}{r}"
        if k<36:
            fm[f"{a}{r}"]=f'=SUMIFS(s_Net,s_Month,$B{r},s_Segment,"{s}")'
            fm[f"{de}{r}"]=f"={a}{r}/{ix}{r}"; fm[f"{ape}{r}"]=f"=ABS({a}{r}-{mo}{r})/{a}{r}"
            fm[f"{ixb}{r}"]=f"=AVERAGEIFS({a}${R0}:{a}${R0+23},$D${R0}:$D${R0+23},$D{r})/AVERAGE({a}${R0}:{a}${R0+23})"
        if k<24: fm[f"{deb}{r}"]=f"={a}{r}/{ixb}{r}"
        if 24<=k<36:
            fm[f"{mob}{r}"]=f"=({mob}$3+{mob}$2*$A{r})*{ixb}{r}"; fm[f"{apeb}{r}"]=f"=ABS({a}{r}-{mob}{r})/{a}{r}"
tc = 5+36; TA,TM,TAPE,TMB,TAPEB,TZ,LD,LU = [L(tc+j) for j in range(8)]
fm.cell(4,tc,"Total network").font=font(bold=True)
for j,h in enumerate(["Actual total","Model total","Abs % err","Model total (bt)","Abs % err (bt)","Z-score of actual","Large Ag avg discount","Large Ag units"]): fm.cell(5,tc+j,h)
for k in range(48):
    r=R0+k
    fm[f"{TM}{r}"]="="+"+".join(f"{blk[s][3]}{r}" for s in SEGS)
    if k<36:
        fm[f"{TA}{r}"]="="+"+".join(f"{blk[s][0]}{r}" for s in SEGS)
        fm[f"{TAPE}{r}"]=f"=ABS({TA}{r}-{TM}{r})/{TA}{r}"
        fm[f"{TZ}{r}"]=f"=({TA}{r}-AVERAGE({TA}${R0}:{TA}${R36}))/STDEV({TA}${R0}:{TA}${R36})"
        fm[f"{LD}{r}"]=f'=AVERAGEIFS(s_Disc,s_Month,$B{r},s_Segment,"Large Ag")'
        fm[f"{LU}{r}"]=f'=SUMIFS(s_Units,s_Month,$B{r},s_Segment,"Large Ag")'
    if 24<=k<36:
        fm[f"{TMB}{r}"]="="+"+".join(f"{blk[s][7]}{r}" for s in SEGS); fm[f"{TAPEB}{r}"]=f"=ABS({TA}{r}-{TMB}{r})/{TA}{r}"
for row in fm.iter_rows(min_row=2,max_row=R48):
    for c in row:
        c.font = font(bold=(c.row==5))
        if c.column>=5 and c.row>=R0:
            h = fm.cell(5,c.column).value or ""
            c.number_format = PCT if "%" in h or "discount" in h else ('0.000' if "idx" in h else ('0.00' if "Z-" in h else INT))
for c in range(1,tc+8): fm.column_dimensions[L(c)].width = 14
fm.freeze_panes = "E6"

# ---------- Sales_Analysis ----------
title(sa,"Sales analysis","All figures are live formulas on Sales_Data. Fiscal year runs November to October.")
section(sa,4,"A. Net sales by segment and fiscal year",14)
header(sa,5,["Segment","FY2024","FY2025","FY2026","YoY FY2025","YoY FY2026","GM % FY2024","GM % FY2025","GM % FY2026","Avg disc FY2024","Avg disc FY2025","Avg disc FY2026","Target FY2026","Attainment FY2026"])
for i,s in enumerate(SEGS):
    r=6+i; put(sa,r,1,s)
    for j,f in enumerate(FYS):
        put(sa,r,2+j,f'=SUMIFS(s_Net,s_Segment,$A{r},s_FY,"{f}")',MM)
        put(sa,r,7+j,f'=SUMIFS(s_GM,s_Segment,$A{r},s_FY,"{f}")/{L(2+j)}{r}',PCT)
        put(sa,r,10+j,f'=1-{L(2+j)}{r}/SUMIFS(s_List,s_Segment,$A{r},s_FY,"{f}")',PCT)
    put(sa,r,5,f"=C{r}/B{r}-1",PCT); put(sa,r,6,f"=D{r}/C{r}-1",PCT)
    put(sa,r,13,f"=Targets!D{5+i}",MM,color="008000"); put(sa,r,14,f"=D{r}/M{r}",PCT)
r=10; put(sa,r,1,"Total",bold=True)
for j,f in enumerate(FYS):
    put(sa,r,2+j,f"=SUM({L(2+j)}6:{L(2+j)}9)",MM,bold=True)
    put(sa,r,7+j,f'=SUMIFS(s_GM,s_FY,"{f}")/{L(2+j)}10',PCT,bold=True)
    put(sa,r,10+j,f'=1-{L(2+j)}10/SUMIFS(s_List,s_FY,"{f}")',PCT,bold=True)
put(sa,r,5,"=C10/B10-1",PCT,bold=True); put(sa,r,6,"=D10/C10-1",PCT,bold=True)
put(sa,r,13,"=SUM(M6:M9)",MM,bold=True); put(sa,r,14,"=D10/M10",PCT,bold=True)

section(sa,13,"B. Net sales by region",14)
header(sa,14,["Region","FY2024","FY2025","FY2026","YoY FY2025","YoY FY2026","GM % FY2026","Share FY2026"])
for i,g in enumerate(REGIONS):
    r=15+i; put(sa,r,1,g)
    for j,f in enumerate(FYS): put(sa,r,2+j,f'=SUMIFS(s_Net,s_Region,$A{r},s_FY,"{f}")',MM)
    put(sa,r,5,f"=C{r}/B{r}-1",PCT); put(sa,r,6,f"=D{r}/C{r}-1",PCT)
    put(sa,r,7,f'=SUMIFS(s_GM,s_Region,$A{r},s_FY,"FY2026")/D{r}',PCT); put(sa,r,8,f"=D{r}/$D$19",PCT)
put(sa,19,1,"Total",bold=True)
for j in range(3): put(sa,19,2+j,f"=SUM({L(2+j)}15:{L(2+j)}18)",MM,bold=True)

section(sa,22,"C. Discount bands, FY2026 (does deeper discounting pay?)",14)
header(sa,23,["Discount band","From (>=)","To (<)","Orders","Net Sales","Share of sales","GM %","Avg order value"])
bands=[("0% to <5%",0,.05),("5% to <10%",.05,.10),("10% to <15%",.10,.15),("15% and above",.15,1)]
for i,(n,lo,hi) in enumerate(bands):
    r=24+i; put(sa,r,1,n); inp(sa,r,2,lo,PCT); inp(sa,r,3,hi,PCT)
    crit=f's_FY,"FY2026",s_Disc,">="&$B{r},s_Disc,"<"&$C{r}'
    put(sa,r,4,f"=COUNTIFS({crit})",INT); put(sa,r,5,f"=SUMIFS(s_Net,{crit})",MM); put(sa,r,6,f"=E{r}/SUM($E$24:$E$27)",PCT)
    put(sa,r,7,f"=IF(E{r}=0,0,SUMIFS(s_GM,{crit})/E{r})",PCT); put(sa,r,8,f"=IF(D{r}=0,0,E{r}/D{r})",MONEY)

section(sa,30,"D. Customer type, FY2026",14)
header(sa,31,["Customer type","Net Sales","Share","GM %","Avg discount"])
ctypes=["Large Farm Operation","Family Farm","Custom Operator","Contractor","Landscaper / Municipal","Homeowner / Hobby"]
for i,c in enumerate(ctypes):
    r=32+i; put(sa,r,1,c); put(sa,r,2,f'=SUMIFS(s_Net,s_Cust,$A{r},s_FY,"FY2026")',MM); put(sa,r,3,f"=B{r}/SUM($B$32:$B$37)",PCT)
    put(sa,r,4,f'=SUMIFS(s_GM,s_Cust,$A{r},s_FY,"FY2026")/B{r}',PCT); put(sa,r,5,f'=1-B{r}/SUMIFS(s_List,s_Cust,$A{r},s_FY,"FY2026")',PCT)

section(sa,40,"E. Dealer scorecard",14)
header(sa,41,["Dealer ID","Dealer","Region","FY2025 Net Sales","FY2026 Net Sales","YoY","GM % FY2026","Avg discount FY2026","On-time % FY2026","Rank (sales)","Flag"])
DS0=42; DS1=DS0+11
for i,d in enumerate(dealers):
    r=DS0+i; put(sa,r,1,d[0]); put(sa,r,2,f"=INDEX(Dim_Dealer!$B${DR0}:$B${DR1},MATCH(A{r},{dl},0))"); put(sa,r,3,f"=INDEX(Dim_Dealer!$D${DR0}:$D${DR1},MATCH(A{r},{dl},0))")
    put(sa,r,4,f'=SUMIFS(s_Net,s_Dealer,$A{r},s_FY,"FY2025")',MM); put(sa,r,5,f'=SUMIFS(s_Net,s_Dealer,$A{r},s_FY,"FY2026")',MM)
    put(sa,r,6,f"=E{r}/D{r}-1",PCT); put(sa,r,7,f'=SUMIFS(s_GM,s_Dealer,$A{r},s_FY,"FY2026")/E{r}',PCT)
    put(sa,r,8,f'=1-E{r}/SUMIFS(s_List,s_Dealer,$A{r},s_FY,"FY2026")',PCT)
    put(sa,r,9,f'=COUNTIFS(s_Dealer,$A{r},s_FY,"FY2026",s_OnTime,1)/COUNTIFS(s_Dealer,$A{r},s_FY,"FY2026")',PCT)
    put(sa,r,10,f"=RANK(E{r},$E${DS0}:$E${DS1})",INT)
    put(sa,r,11,f'=IF(AND(H{r}>$L$10+0.03,G{r}<$I$10),"High discount, low margin",IF(I{r}<0.8,"Delivery issue","OK"))')
sa.conditional_formatting.add(f"K{DS0}:K{DS1}", FormulaRule(formula=[f'K{DS0}<>"OK"'], fill=PatternFill("solid",fgColor="FCE4D6")))

section(sa,57,"F. Product Pareto, FY2026 (ABC by cumulative share of net sales)",14)
header(sa,58,["Product ID","Product","Segment","FY2026 Net Sales","Share","Cumulative share","ABC class","Units FY2026","YoY net sales"])
PP0=59; PP1=PP0+len(products)-1
for i,p in enumerate(products):
    r=PP0+i; put(sa,r,1,p[0]); put(sa,r,2,p[1]); put(sa,r,3,p[3])
    put(sa,r,4,f'=SUMIFS(s_Net,s_Product,$A{r},s_FY,"FY2026")',MM); put(sa,r,5,f"=D{r}/SUM($D${PP0}:$D${PP1})",PCT)
    put(sa,r,6,f'=SUMIF($D${PP0}:$D${PP1},">="&D{r})/SUM($D${PP0}:$D${PP1})',PCT)
    put(sa,r,7,f'=IF(F{r}<=0.8,"A",IF(F{r}<=0.95,"B","C"))',align="center")
    put(sa,r,8,f'=SUMIFS(s_Units,s_Product,$A{r},s_FY,"FY2026")',INT)
    put(sa,r,9,f'=D{r}/SUMIFS(s_Net,s_Product,$A{r},s_FY,"FY2025")-1',PCT)
note(sa,PP1+1,"A = products that make up the first 80% of net sales, B = next 15%, C = last 5%. Thresholds are a common convention, not a fixed rule.")
ch=BarChart(); ch.type="col"; ch.add_data(Reference(sa,min_col=2,max_col=4,min_row=5,max_row=9),titles_from_data=True)
ch.set_categories(Reference(sa,min_col=1,min_row=6,max_row=9)); chart_fix(ch,"Net sales by segment ($)",18,8.5); ch.y_axis.number_format='$#,##0,,"M"'
sa.add_chart(ch,"J13")
ch=BarChart(); ch.type="bar"; ch.add_data(Reference(sa,min_col=7,min_row=23,max_row=27),titles_from_data=True)
ch.set_categories(Reference(sa,min_col=1,min_row=24,max_row=27)); chart_fix(ch,"Gross margin % by discount band, FY2026",18,7.5); ch.legend=None; ch.y_axis.number_format='0%'
sa.add_chart(ch,"J28")

# ---------- Ops_Analysis ----------
title(ops,"Operations analysis: inventory and delivery","Inventory on hand is a sample end-of-FY2026 snapshot (blue inputs). Demand and delivery metrics are formulas on Sales_Data.")
section(ops,4,"Assumptions (edit the yellow cells)",14)
ass=[("Overstock threshold (months of supply)",6,'0.0'),("Stockout-risk threshold (months of supply)",1.5,'0.0'),("Target months of supply for reorders",3,'0.0'),("Annual inventory carrying cost rate",.08,PCT)]
for i,(l,v,f_) in enumerate(ass):
    put(ops,5+i,1,l); ops.merge_cells(start_row=5+i,start_column=1,end_row=5+i,end_column=3); inp(ops,5+i,4,v,f_,key=True)
note(ops,5,"Assumption: generic rule-of-thumb thresholds for this sample; replace with company policy.",5)
note(ops,8,"Assumption: illustrative floorplan interest + storage + insurance rate; not a John Deere figure.",5)
section(ops,10,"A. Inventory health by product",14)
header(ops,11,["Product ID","Product","Segment","On hand (units)","Unit cost","Inventory value","Units sold FY2026","Avg monthly demand","Months of supply","Status","Excess units","Excess value","Suggested order (units)","Sales value of suggested order"])
eq=[p for p in products if p[3]!="Parts & Service"]; IV0=12; IV1=IV0+len(eq)-1
for i,p in enumerate(eq):
    r=IV0+i; put(ops,r,1,p[0]); put(ops,r,2,p[1]); put(ops,r,3,p[3]); inp(ops,r,4,onhand[p[0]],INT)
    put(ops,r,5,f"=INDEX(Dim_Product!$F${PR0}:$F${PR1},MATCH(A{r},{pl},0))",MONEY,color="008000"); put(ops,r,6,f"=D{r}*E{r}",MONEY)
    put(ops,r,7,f'=SUMIFS(s_Units,s_Product,$A{r},s_FY,"FY2026")',INT); put(ops,r,8,f"=G{r}/12",'0.0')
    put(ops,r,9,f"=IF(H{r}=0,99,D{r}/H{r})",'0.0')
    put(ops,r,10,f'=IF(I{r}>$D$5,"Overstock",IF(I{r}<$D$6,"Stockout risk","Healthy"))')
    put(ops,r,11,f"=MAX(0,ROUND(D{r}-$D$5*H{r},0))",INT); put(ops,r,12,f"=K{r}*E{r}",MONEY)
    put(ops,r,13,f'=IF(J{r}="Stockout risk",ROUNDUP($D$7*H{r}-D{r},0),0)',INT)
    put(ops,r,14,f"=M{r}*INDEX(Dim_Product!$E${PR0}:$E${PR1},MATCH(A{r},{pl},0))",MONEY)
IT=IV1+1; put(ops,IT,1,"Total",bold=True)
for c in (4,6,7,11,12,13,14): put(ops,IT,c,f"=SUM({L(c)}{IV0}:{L(c)}{IV1})",MONEY if c in(6,12,14) else INT,bold=True)
ops.conditional_formatting.add(f"J{IV0}:J{IV1}", CellIsRule(operator="equal",formula=['"Overstock"'],fill=PatternFill("solid",fgColor="FCE4D6")))
ops.conditional_formatting.add(f"J{IV0}:J{IV1}", CellIsRule(operator="equal",formula=['"Stockout risk"'],fill=PatternFill("solid",fgColor="FFF2CC")))
DL=IT+3; section(ops,DL,"B. Delivery performance, FY2026",14)
header(ops,DL+1,["Group","","","Orders","On-time %","Late orders","Avg days late (late orders)"])
def deliv(r,label,rng):
    put(ops,r,1,label); ops.merge_cells(start_row=r,start_column=1,end_row=r,end_column=3)
    crit = f'{rng},$A{r},s_FY,"FY2026"' if rng else 's_FY,"FY2026"'
    put(ops,r,4,f"=COUNTIFS({crit})",INT); put(ops,r,5,f"=COUNTIFS({crit},s_OnTime,1)/D{r}",PCT)
    put(ops,r,6,f"=D{r}-COUNTIFS({crit},s_OnTime,1)",INT); put(ops,r,7,f"=IF(F{r}=0,0,SUMIFS(s_DaysLate,{crit})/F{r})",'0.0')
DRG=DL+2
for i,g in enumerate(REGIONS): deliv(DRG+i,g,"s_Region")
deliv(DRG+4,"All regions",None)
DSG=DRG+6
for i,s in enumerate(SEGS): deliv(DSG+i,s,"s_Segment")
ch=BarChart(); ch.type="bar"; ch.add_data(Reference(ops,min_col=9,min_row=11,max_row=IV1),titles_from_data=True)
ch.set_categories(Reference(ops,min_col=2,min_row=IV0,max_row=IV1)); chart_fix(ch,"Months of supply by product",18,11); ch.legend=None
ops.add_chart(ch,"I"+str(DL+1))

# ---------- Statistics ----------
title(st,"Statistics","Descriptive statistics, correlation and trend regression. All formulas; all three fiscal years unless stated.")
def interp(c): return f'=IF(ABS({c})>=0.7,"Strong",IF(ABS({c})>=0.4,"Moderate",IF(ABS({c})>=0.2,"Weak","Negligible")))&IF({c}<0," negative"," positive")&" relationship"'
section(st,4,"A. Order value distribution (net sales per order)",6)
header(st,5,["Statistic","Value"])
ds=[("Orders","=COUNT(s_Net)",INT),("Mean","=AVERAGE(s_Net)",MONEY),("Median","=MEDIAN(s_Net)",MONEY),("Standard deviation","=STDEV(s_Net)",MONEY),
    ("Minimum","=MIN(s_Net)",MONEY),("1st quartile","=QUARTILE(s_Net,1)",MONEY),("3rd quartile","=QUARTILE(s_Net,3)",MONEY),("Maximum","=MAX(s_Net)",MONEY),
    ("Skewness","=SKEW(s_Net)",'0.00'),("Coefficient of variation","=B9/B7",'0.00')]
for i,(l,f_,nf) in enumerate(ds): put(st,6+i,1,l); put(st,6+i,2,f_,nf)
note(st,6,"Mean far above median + high skew = a few very large equipment orders dominate; report medians and segment cuts, not one average.",3)
section(st,18,"B. By segment",6)
header(st,19,["Segment","Orders","Avg order value","Avg discount (per order)","Avg GM % (per order)"])
for i,s in enumerate(SEGS):
    r=20+i; put(st,r,1,s); put(st,r,2,f"=COUNTIFS(s_Segment,A{r})",INT); put(st,r,3,f"=AVERAGEIFS(s_Net,s_Segment,A{r})",MONEY)
    put(st,r,4,f"=AVERAGEIFS(s_Disc,s_Segment,A{r})",PCT); put(st,r,5,f"=AVERAGEIFS(s_GMpct,s_Segment,A{r})",PCT)
section(st,26,"C. Correlation (Pearson r)",6)
header(st,27,["Question","r","","","","Reading"])
cq=[("Order level: discount % vs gross margin %","=CORREL(s_Disc,s_GMpct)"),
    ("Large Ag monthly: avg discount vs units sold (does discounting lift volume?)",f"=CORREL(Forecast_Model!{LD}{R0}:{LD}{R36},Forecast_Model!{LU}{R0}:{LU}{R36})"),
    ("Dealer level: avg discount FY2026 vs YoY sales growth (12 dealers)",f"=CORREL(Sales_Analysis!H{DS0}:H{DS1},Sales_Analysis!F{DS0}:F{DS1})")]
for i,(q,f_) in enumerate(cq):
    r=28+i; put(st,r,1,q,wrap=True); put(st,r,2,f_,'0.00'); put(st,r,6,interp(f"B{r}"),wrap=True); st.row_dimensions[r].height=28
note(st,31,"Correlation is not causation. The dealer-level test has only 12 points, so treat it as a prompt for a conversation, not proof.")
section(st,33,"D. Trend regression on deseasonalized monthly net sales (36 months)",6)
header(st,34,["Segment","Slope ($ per month)","Intercept ($)","R-squared","Direction"])
for i,s in enumerate(SEGS):
    r=35+i; de=blk[s][2]; rng=f"Forecast_Model!{de}{R0}:{de}{R36},Forecast_Model!{tA}"
    put(st,r,1,s); put(st,r,2,f"=SLOPE({rng})",MONEY); put(st,r,3,f"=INTERCEPT({rng})",MONEY); put(st,r,4,f"=RSQ({rng})",'0.00')
    put(st,r,5,f'=IF(B{r}>0,"Growing","Declining")')
section(st,41,"E. Outlier months (total network net sales)",6)
zr=f"Forecast_Model!{TZ}{R0}:{TZ}{R36}"
put(st,42,1,"Months with |z-score| above 2"); put(st,42,2,f'=COUNTIF({zr},">2")+COUNTIF({zr},"<-2")',INT)
put(st,43,1,"Highest month"); put(st,43,2,f"=INDEX(Forecast_Model!$B${R0}:$B${R36},MATCH(MAX({zr}),{zr},0))",'mmm yyyy')
put(st,44,1,"Lowest month"); put(st,44,2,f"=INDEX(Forecast_Model!$B${R0}:$B${R36},MATCH(MIN({zr}),{zr},0))",'mmm yyyy')

# ---------- Forecast ----------
title(fc,"FY2027 sales forecast","Method: seasonal index x linear trend per segment (see Forecast_Model). Accuracy shown two ways: in-sample fit and an FY2026 holdout backtest.")
section(fc,4,"A. Forecast summary",9)
header(fc,5,["Segment","FY2026 actual","FY2027 forecast","Change %","In-sample MAPE","Backtest MAPE (FY2026 holdout)","Low case","High case","Trend ($ per month)"])
for i,s in enumerate(SEGS):
    r=6+i; a,ix,de,mo,ape,ixb,deb,mob,apeb=blk[s]; M="Forecast_Model!"
    put(fc,r,1,s); put(fc,r,2,f"=SUM({M}{a}{R36-11}:{a}{R36})",MM); put(fc,r,3,f"=SUM({M}{mo}{R36+1}:{mo}{R48})",MM)
    put(fc,r,4,f"=C{r}/B{r}-1",PCT); put(fc,r,5,f"=AVERAGE({M}{ape}{R0}:{ape}{R36})",PCT); put(fc,r,6,f"=AVERAGE({M}{apeb}{R36-11}:{apeb}{R36})",PCT)
    put(fc,r,7,f"=C{r}*(1-F{r})",MM); put(fc,r,8,f"=C{r}*(1+F{r})",MM); put(fc,r,9,f"={M}{mo}2",MONEY)
r=10; M="Forecast_Model!"; put(fc,r,1,"Total network",bold=True); put(fc,r,2,"=SUM(B6:B9)",MM,bold=True); put(fc,r,3,"=SUM(C6:C9)",MM,bold=True); put(fc,r,4,"=C10/B10-1",PCT,bold=True)
put(fc,r,5,f"=AVERAGE({M}{TAPE}{R0}:{TAPE}{R36})",PCT,bold=True); put(fc,r,6,f"=AVERAGE({M}{TAPEB}{R36-11}:{TAPEB}{R36})",PCT,bold=True)
put(fc,r,7,"=C10*(1-F10)",MM,bold=True); put(fc,r,8,"=C10*(1+F10)",MM,bold=True); put(fc,r,9,"=SUM(I6:I9)",MONEY,bold=True)
note(fc,11,"Low/high case = forecast -/+ the backtest MAPE. This is a practical error band, not a statistical confidence interval.")
note(fc,12,"Limit: a linear trend extends the past. It cannot see a cycle turning (e.g. Large Ag bottoming). Use What_If to overlay a management view.")
section(fc,14,"B. FY2027 forecast by month",9)
header(fc,15,["Month"]+SEGS+["Total"])
for k in range(12):
    r=16+k; fr=R36+1+k; put(fc,r,1,f"=Forecast_Model!B{fr}",'mmm yyyy')
    for i,s in enumerate(SEGS): put(fc,r,2+i,f"=Forecast_Model!{blk[s][3]}{fr}",MM)
    put(fc,r,6,f"=SUM(B{r}:E{r})",MM,bold=True)
put(fc,28,1,"FY2027 total",bold=True)
for c in range(2,7): put(fc,28,c,f"=SUM({L(c)}16:{L(c)}27)",MM,bold=True)
def trend_chart(w=24,h=9):
    ch=LineChart(); col=tc
    ch.add_data(Reference(fm,min_col=col,max_col=col+1,min_row=5,max_row=R48),titles_from_data=True)
    ch.set_categories(Reference(fm,min_col=2,min_row=R0,max_row=R48)); chart_fix(ch,"Monthly net sales: actual vs model, with FY2027 forecast",w,h)
    ch.y_axis.number_format='$#,##0,,"M"'; ch.x_axis.number_format='mmm yy'; ch.x_axis.tickLblSkip=3
    ch.series[1].graphicalProperties.line.dashStyle="dash"; return ch
fc.add_chart(trend_chart(),"A31")

# ---------- What_If ----------
title(wi,"What-if scenarios for FY2027","Pick a scenario in the yellow cell. Scenario 4 is yours to edit. Base = statistical forecast with FY2026 margin and discount levels.")
put(wi,4,1,"Active scenario (1-4)",bold=True); inp(wi,4,2,2,INT,key=True)
dv=DataValidation(type="list",formula1='"1,2,3,4"',allow_blank=False); wi.add_data_validation(dv); dv.add("B4")
put(wi,4,3,'=INDEX(B7:E7,1,B4)',bold=True)
section(wi,6,"A. Scenario levers (changes versus the statistical baseline)",7)
header(wi,7,["Lever","1. Statistical baseline","2. Large Ag cycle bottoms","3. Downside: weak demand + tariffs","4. Custom (edit me)","Active"])
levers=[("Large Ag volume change",[0,.15,-.10,0]),("Small Ag & Turf volume change",[0,.03,-.05,0]),("Construction & Forestry volume change",[0,.05,-.10,0]),
        ("Parts & Service volume change",[0,.02,0,0]),("Discount change (percentage points)",[0,-.01,.02,0]),("Unit cost inflation",[0,.01,.04,0])]
for i,(l,v) in enumerate(levers):
    r=8+i; put(wi,r,1,l)
    for j in range(4): inp(wi,r,2+j,v[j],PCT,key=(j==3))
    put(wi,r,6,f"=INDEX(B{r}:E{r},1,$B$4)",PCT,bold=True)
note(wi,14,"Scenario values are illustrative assumptions written for this sample, not company guidance.")
section(wi,16,"B. Result by segment",7)
header(wi,17,["Segment","Base net sales","Scenario net sales","Base gross margin","Scenario gross margin","Scenario GM %","Change in gross margin"])
for i,s in enumerate(SEGS):
    r=18+i; sr=6+i; put(wi,r,1,s); put(wi,r,2,f"=Forecast!C{sr}",MM,color="008000")
    put(wi,r,3,f"=B{r}*(1+F{8+i})*(1-(Sales_Analysis!L{sr}+$F$12))/(1-Sales_Analysis!L{sr})",MM)
    put(wi,r,4,f"=B{r}*Sales_Analysis!I{sr}",MM)
    put(wi,r,5,f"=C{r}-B{r}*(1-Sales_Analysis!I{sr})*(1+F{8+i})*(1+$F$13)",MM)
    put(wi,r,6,f"=IF(C{r}=0,0,E{r}/C{r})",PCT); put(wi,r,7,f"=E{r}-D{r}",MM)
put(wi,22,1,"Total",bold=True)
for c in (2,3,4,5,7): put(wi,22,c,f"=SUM({L(c)}18:{L(c)}21)",MM,bold=True)
put(wi,22,6,"=E22/C22",PCT,bold=True)
put(wi,24,1,"Scenario net sales vs FY2026 actual"); put(wi,24,2,"=C22/Forecast!B10-1",PCT,bold=True)
put(wi,25,1,"Scenario gross margin vs baseline"); put(wi,25,2,"=G22",MM,bold=True)
note(wi,27,"Logic: sales scale with volume and the change in discount; cost of goods scales with volume and unit cost inflation.")

# ---------- Findings ----------
title(fin,"Findings and recommended actions","Every number in the text is a live formula. Change the data or an assumption and the story updates.")
put(fin,3,2,"Assumption: discount ceiling used for finding 2"); inp(fin,3,3,.12,PCT,key=True)
note(fin,3,"Illustrative policy level; the analysis assumes the capped orders would still close (see Statistics C).",4)
header(fin,5,["#","Finding","Value at stake","Recommended action","$ amount","Type of value","How it is calculated"])
SA="Sales_Analysis!"; OP="Ops_Analysis!"
hi=f's_FY,"FY2026",s_Disc,">="&0.15'
T=lambda x,f_="0.0": f'TEXT({x},"{f_}")'
m=lambda x: f'TEXT(({x})/1000000,"$#,##0.0")&"M"'
p=lambda x: f'TEXT({x},"0.0%")'
west_r=DRG+3; all_r=DRG+4
F_=[
 (f'="Large Ag net sales fell "&{p(f"-({SA}D6/{SA}B6-1)")}&" from FY2024 to FY2026 ("&{m(f"{SA}B6")}&" to "&{m(f"{SA}D6")}&"), and its share of the mix dropped from "&{p(f"{SA}B6/{SA}B10")}&" to "&{p(f"{SA}D6/{SA}D10")}&". Construction & Forestry grew "&{p(f"{SA}F8")}&" in FY2026."',
  "Sales lost in the down cycle","Shift sales effort, stocking and marketing toward Construction, Small Ag and aftermarket while Large Ag is weak; plan Large Ag inventory for a recovery rather than chasing volume with price.",
  f"={SA}B6-{SA}D6","Revenue decline (context)","FY2024 minus FY2026 Large Ag net sales (Sales_Analysis A)."),
 (f'="Orders discounted 15%+ were "&{p(f"{SA}F27")}&" of FY2026 sales but earned only "&{p(f"{SA}G27")}&" gross margin, versus "&{p(f"{SA}G24")}&" on orders under 5%. Large Ag average discount rose from "&{p(f"{SA}J6")}&" to "&{p(f"{SA}L6")}&" in two years."',
  "Margin recovered by a discount ceiling","Set a discount approval ceiling; review the dealers flagged in the scorecard; replace blanket price cuts with targeted incentives (financing, trade-in, service bundles).",
  f"=SUMIFS(s_List,{hi})*(1-C3)-SUMIFS(s_Net,{hi})","Margin recovery (annual)","List value of FY2026 orders discounted 15%+ x (1 - ceiling) minus their actual net sales. Assumes no lost orders."),
 (f'=COUNTIF({OP}J{IV0}:J{IV1},"Overstock")&" products sit above "&{T(f"{OP}D5")}&" months of supply. Excess stock is worth "&{m(f"{OP}L{IT}")}&" at cost, out of "&{m(f"{OP}F{IT}")}&" total inventory."',
  "Cash tied up in excess stock","Slow Large Ag replenishment, transfer units between dealers before ordering new, and run a targeted sell-down on the overstocked lines.",
  f"={OP}L{IT}","Cash released (one-time)","Units above the overstock threshold x unit cost (Ops_Analysis A). Carrying cost saved per year = this x the carrying rate."),
 (f'=COUNTIF({OP}J{IV0}:J{IV1},"Stockout risk")&" fast-moving products are under "&{T(f"{OP}D6")}&" months of supply. "&TEXT({OP}M{IT},"#,##0")&" more units are needed to get back to "&{T(f"{OP}D7")}&" months of cover."',
  "Sales protected by reordering","Reorder to the target months of supply now; add a weekly months-of-supply alert so growth lines never drop below threshold.",
  f"={OP}N{IT}","Revenue protected","Suggested order units x list price (Ops_Analysis A)."),
 (f'="West region delivered only "&{p(f"{OP}E{west_r}")}&" of FY2026 orders on time versus "&{p(f"{OP}E{all_r}")}&" for the network; late orders there averaged "&{T(f"{OP}G{west_r}")}&" days late."',
  "Revenue exposed to late delivery","Root-cause the West freight and prep process; set a network on-time target and track it on the dashboard by dealer.",
  f'=SUMIFS(s_Net,s_Region,"West",s_FY,"FY2026",s_OnTime,0)',"Revenue at risk (customer experience)","Net sales of FY2026 West orders delivered late."),
 (f'="Parts & Service is "&{p(f"{SA}D9/{SA}D10")}&" of FY2026 sales but "&{p(f"SUMIFS(s_GM,s_Segment,{SA}A9,s_FY,{chr(34)*2}FY2026{chr(34)*2})/SUMIFS(s_GM,s_FY,{chr(34)*2}FY2026{chr(34)*2})")}&" of gross margin, at "&{p(f"{SA}I9")}&" margin versus "&{p(f"(SUMIFS(s_GM,s_FY,{chr(34)*2}FY2026{chr(34)*2})-SUMIFS(s_GM,s_Segment,{SA}A9,s_FY,{chr(34)*2}FY2026{chr(34)*2}))/({SA}D10-{SA}D9)")}&" on equipment, and it grew "&{p(f"{SA}F9")}&" while Large Ag shrank."',
  "Aftermarket gross margin","Protect and grow the aftermarket: attach service plans and technology upgrades to every equipment sale and to the existing fleet.",
  f'=SUMIFS(s_GM,s_Segment,{SA}A9,s_FY,"FY2026")',"Margin base to grow","FY2026 Parts & Service gross margin."),
 (f'="The model forecasts FY2027 net sales of "&{m("Forecast!C10")}&" ("&{p("Forecast!D10")}&" vs FY2026), range "&{m("Forecast!G10")}&" to "&{m("Forecast!H10")}&". Backtested on FY2026 it was off by "&{p("Forecast!F10")}&" on average per month."',
  "Planning baseline","Use the forecast as the baseline for stocking and targets, then overlay scenarios in What_If; refresh monthly and track forecast error.",
  "=Forecast!C10","Forecast (planning input)","Sum of the four segment forecasts for Nov 2026 to Oct 2027."),
]
for i,(txt,stake,act,amt,typ,how) in enumerate(F_):
    r=6+i; fin.row_dimensions[r].height=92
    put(fin,r,1,i+1,align="center"); put(fin,r,2,txt.replace('""','"'),wrap=True); put(fin,r,3,stake,wrap=True); put(fin,r,4,act,wrap=True)
    put(fin,r,5,amt,MM,bold=True,align="right"); put(fin,r,6,typ,wrap=True); put(fin,r,7,how,wrap=True)
put(fin,14,2,"Carrying cost saved per year if excess stock is cleared",bold=True); put(fin,14,5,f"=E8*{OP}D8",MM,bold=True)
put(fin,15,2,"Recurring annual profit opportunity (findings 2 + carrying cost)",bold=True); put(fin,15,5,"=E7+E14",MM,bold=True,fill=SFILL)
put(fin,16,2,"...as a share of FY2026 gross margin",bold=True); put(fin,16,5,'=E15/SUMIFS(s_GM,s_FY,"FY2026")',PCT,bold=True)
note(fin,18,"These amounts describe the SAMPLE data only. With real data, each would be validated with finance, inventory planning and dealer-facing teams before any recommendation." if PUBLIC else "These amounts describe the SAMPLE data. In an interview, present them as what the model shows on synthetic data, and explain how you would validate each with the business.",2)

# ---------- Dashboard ----------
dash["B1"]="Sales & Operations Command Center"; dash["B1"].font=font(size=18,bold=True,color=NAVY)
dash["B2"]="Fictional equipment dealer network (synthetic sample data). Choose a fiscal year and region in the yellow cells."; dash["B2"].font=font(italic=True,color="5B6770")
put(dash,4,2,"Fiscal year",bold=True); inp(dash,4,3,"FY2026",key=True); put(dash,4,5,"Region",bold=True); inp(dash,4,6,"All",key=True)
dv1=DataValidation(type="list",formula1='"FY2024,FY2025,FY2026"'); dv2=DataValidation(type="list",formula1='"All,Midwest,Plains,South,West"')
dash.add_data_validation(dv1); dash.add_data_validation(dv2); dv1.add("C4"); dv2.add("F4")
dash["H4"]="Prior FY"; dash["I4"]='="FY"&(VALUE(RIGHT(C4,4))-1)'; dash["K4"]="Region filter"; dash["L4"]='=IF(F4="All","*",F4)'
for a in ("H4","I4","K4","L4"): dash[a].font=font(color="9AA0A6",size=9)
cr='s_FY,$C$4,s_Region,$L$4'; crp='s_FY,$I$4,s_Region,$L$4'
cards=[("Net sales",f"=SUMIFS(s_Net,{cr})",MM,'="Orders: "&TEXT(COUNTIFS(%s),"#,##0")'%cr),
       ("YoY growth",f'=IF(SUMIFS(s_Net,{crp})=0,"n/a",SUMIFS(s_Net,{cr})/SUMIFS(s_Net,{crp})-1)',PCT,'="vs "&I4'),
       ("Gross margin %",f"=SUMIFS(s_GM,{cr})/SUMIFS(s_Net,{cr})",PCT,f'="Gross margin: "&TEXT(SUMIFS(s_GM,{cr})/1000000,"$#,##0.0")&"M"'),
       ("Avg discount",f"=1-SUMIFS(s_Net,{cr})/SUMIFS(s_List,{cr})",PCT,"Off list price"),
       ("On-time delivery",f"=COUNTIFS({cr},s_OnTime,1)/COUNTIFS({cr})",PCT,"Share of orders"),
       ("Excess inventory",f"=Ops_Analysis!L{IT}",MM,"Network, end FY2026"),
       ("FY2027 forecast","=Forecast!C10",MM,'="Backtest error: "&TEXT(Forecast!F10,"0.0%")')]
for i,(lab,f_,nf,subf) in enumerate(cards):
    c=2+i*2
    for rr,(v,fo,h) in enumerate([(lab,font(bold=True,color="FFFFFF",size=10),18),(f_,font(bold=True,size=16,color=NAVY),30),(subf if subf.startswith("=") else subf,font(size=9,color="5B6770"),16)]):
        r=6+rr; dash.merge_cells(start_row=r,start_column=c,end_row=r,end_column=c+1)
        x=dash.cell(r,c,v); x.font=fo; x.alignment=Alignment(horizontal="center",vertical="center"); dash.row_dimensions[r].height=h
        x.fill = HFILL if rr==0 else SFILL
        if rr==1: x.number_format=nf
header(dash,11,["Segment","","Selected FY","Prior FY","YoY","GM %","Target (network)","Attainment"],2)
for i,s in enumerate(SEGS):
    r=12+i; put(dash,r,2,s); dash.merge_cells(start_row=r,start_column=2,end_row=r,end_column=3)
    put(dash,r,4,f"=SUMIFS(s_Net,{cr},s_Segment,$B{r})",MM); put(dash,r,5,f"=SUMIFS(s_Net,{crp},s_Segment,$B{r})",MM)
    put(dash,r,6,f'=IF(E{r}=0,"n/a",D{r}/E{r}-1)',PCT); put(dash,r,7,f"=IF(D{r}=0,0,SUMIFS(s_GM,{cr},s_Segment,$B{r})/D{r})",PCT)
    put(dash,r,8,f"=INDEX(Targets!$B$5:$D$8,MATCH($B{r},Targets!$A$5:$A$8,0),MATCH($C$4,Targets!$B$4:$D$4,0))",MM)
    put(dash,r,9,f'=IF($F$4="All",D{r}/H{r},"n/a")',PCT)
dash["D11"]="Selected FY"
header(dash,11,["Region","Net sales (selected FY)","On-time %"],11)
for i,g in enumerate(REGIONS):
    r=12+i; put(dash,r,11,g); put(dash,r,12,f"=SUMIFS(s_Net,s_FY,$C$4,s_Region,K{r})",MM); put(dash,r,13,f"=COUNTIFS(s_FY,$C$4,s_Region,K{r},s_OnTime,1)/COUNTIFS(s_FY,$C$4,s_Region,K{r})",PCT)
note(dash,16,"Targets are set at network level, so attainment shows only when Region = All. Region table always shows all regions.",2)
dash["B18"]="What needs attention"; dash["B18"].font=font(size=12,bold=True,color=NAVY)
for i in range(5):
    r=19+i; dash.merge_cells(start_row=r,start_column=2,end_row=r,end_column=15)
    x=dash.cell(r,2,f'="{i+1}. "&Findings!B{6+i}'); x.font=font(); x.alignment=Alignment(wrap_text=True,vertical="center"); dash.row_dimensions[r].height=30
    x.fill=GFILL if i%2==0 else PatternFill()
ch=BarChart(); ch.type="col"; ch.add_data(Reference(dash,min_col=4,max_col=5,min_row=11,max_row=15),titles_from_data=True)
ch.set_categories(Reference(dash,min_col=2,min_row=12,max_row=15)); chart_fix(ch,"Net sales by segment: selected vs prior FY",17,8); ch.y_axis.number_format='$#,##0,,"M"'
dash.add_chart(ch,"B25")
dash.add_chart(trend_chart(19,8),"I25")
ch=BarChart(); ch.type="col"; ch.add_data(Reference(dash,min_col=12,min_row=11,max_row=15),titles_from_data=True)
ch.set_categories(Reference(dash,min_col=11,min_row=12,max_row=15)); chart_fix(ch,"Net sales by region (selected FY)",17,8); ch.legend=None; ch.y_axis.number_format='$#,##0,,"M"'
dash.add_chart(ch,"B42")
ch=BarChart(); ch.type="bar"; ch.add_data(Reference(ops,min_col=9,min_row=11,max_row=IV1),titles_from_data=True)
ch.set_categories(Reference(ops,min_col=2,min_row=IV0,max_row=IV1)); chart_fix(ch,"Inventory: months of supply by product",19,11); ch.legend=None
dash.add_chart(ch,"I42")

# ---------- Start_Here ----------
ws0["B1"]="Sales & Operations Command Center"; ws0["B1"].font=font(size=18,bold=True,color=NAVY)
ws0["B2"]="An end-to-end Excel analytics project: equipment dealer sales, operations, statistics and forecasting"; ws0["B2"].font=font(italic=True,color="5B6770")
def block(ws,r,head,lines):
    ws.cell(r,2,head).font=font(size=12,bold=True,color="FFFFFF"); ws.cell(r,2).fill=PatternFill("solid",fgColor=TEAL); ws.cell(r,3).fill=PatternFill("solid",fgColor=TEAL); r+=1
    for a,b in lines:
        x=ws.cell(r,2,a); x.font=font(bold=True); x.alignment=Alignment(vertical="top",wrap_text=True)
        y=ws.cell(r,3,b); y.font=font(); y.alignment=Alignment(vertical="top",wrap_text=True)
        ws.row_dimensions[r].height=max(15, 14*(1+len(str(b))//(100 if ws is ws0 else 115))); r+=1
    return r+1
r=block(ws0,4,"Read this first",[
 ("Data","100% synthetic. 'Heartland Equipment' is a fictional 12-dealer network. This is NOT John Deere data and no number here describes John Deere."),
 ("Rows",f'=TEXT(COUNTA(s_Dealer),"#,##0")&" order lines, November 2023 to October 2026 (three fiscal years)."'),
 ("Fiscal year","November to October, labelled by the year it ends in (simplified from the 52/53-week fiscal calendar John Deere uses)."),
 ("Segments","Large Ag, Small Ag & Turf, Construction & Forestry, Parts & Service. Modelled loosely on how John Deere reports (Production & Precision Ag, Small Ag & Turf, Construction & Forestry) plus dealer aftermarket."),
 ("Colour code","Blue text = hardcoded sample input. Yellow fill = cell you are meant to change. Black = formula. Green = link from another sheet."),
])
r=block(ws0,r,"Industry context the sample is modelled on (public John Deere earnings information)" if PUBLIC else "Why this project fits John Deere right now (public information)",[
 ("Large ag is in a down cycle","John Deere's FY2026 outlook called for US & Canada large ag industry sales down 15 to 20%, while management described 2026 as the bottom of the cycle. Source: Deere Q1 FY2026 release, via finviz.com/news/315392 and agnavigator.com (Feb 19, 2026)."),
 ("Other segments are growing","Small Ag & Turf and Construction & Forestry net sales each grew over 20% in Q1 FY2026; Construction & Forestry was up 18% in Q3 FY2026. Sources: Q1 FY2026 earnings call transcript (fintool.com); Q3 FY2026 release (Aug 2026)."),
 ("Inventory discipline","Management highlighted cutting North American new field inventory of high-horsepower tractors and combines by over 50% from the 2024 peak. Source: Q2 FY2026 earnings call summary (finance.yahoo.com)."),
 ("Margin pressure","Tariffs and production costs were cited as margin pressures in the FY2025 results release (Nov 26, 2025)."),
 ("So the sample data was built to show","A shrinking Large Ag line, rising discounts, overstocked big machines, understocked compact machines, a growing aftermarket, and a region with delivery problems. Those are the business questions the workbook answers."),
])
r=block(ws0,r,"Sheet map",[
 ("Dashboard","One page for a manager: KPI cards, segment and region tables, four charts, top issues. Driven by two dropdowns."),
 ("Findings","Seven insights written in plain language with recommended actions and dollar value. Text updates from formulas."),
 ("Sales_Analysis","Segment, region, discount bands, customer type, dealer scorecard, product Pareto (ABC)."),
 ("Ops_Analysis","Inventory health (months of supply, overstock, stockout risk, suggested orders) and on-time delivery."),
 ("Statistics","Descriptive stats, correlations, trend regression, outlier months."),
 ("Forecast","FY2027 forecast by segment with error measured by an FY2026 holdout backtest."),
 ("What_If","Four scenarios (baseline, recovery, downside, custom) for volume, discount and cost."),
 ("Interview_Guide","How to present the project: pitch, what/why/impact, questions to expect, honest limits."),
 ("Macro_Guide","How to enable the macros, what each one does, and a fallback if Excel blocks them."),
 ("Pivot_* and Control_Panel","Created by the BuildPivotReports macro the first time you run it: six pivot sheets with shared slicers, and a button panel."),
 ("Forecast_Model, Sales_Data, Dim_*, Targets","Calculation engine, raw data, lookup tables and plan numbers."),
])
r=block(ws0,r,"Try this",[
 ("1","On Dashboard, switch Region to West and watch on-time delivery drop."),
 ("2","On Ops_Analysis, change the overstock threshold from 6 to 5 months and see Findings re-price the excess inventory."),
 ("3","On What_If, switch the scenario from 2 to 3 and read the change in gross margin."),
 ("4","Enable macros and answer Yes to the setup prompt (or run BuildPivotReports from Alt+F8). It builds 7 pivot tables, 3 pivot charts and 4 shared slicers, plus a Control_Panel sheet with a button for every macro. See Macro_Guide."),
])

# ---------- Interview guide ----------
ig["B1"]="Interview guide: how to present this project"; ig["B1"].font=font(size=16,bold=True,color=NAVY)
ig["B2"]="Present it as a self-built portfolio project on synthetic data. Numbers below update from the workbook."; ig["B2"].font=font(italic=True,color="5B6770")
r=block(ig,4,"60-second pitch",[
 ("Say this",'="I built an end-to-end sales and operations analytics workbook in Excel for a fictional 12-dealer equipment network, using about "&TEXT(COUNTA(s_Dealer),"#,##0")&" synthetic orders over three fiscal years. I modelled it on the situation equipment makers face today: large ag in a down cycle, pressure on margin, and inventory that has to be right-sized. It takes raw orders through a data model, sales and inventory analysis, statistics and a backtested forecast, and ends in a dashboard and a list of actions with dollar values."'),
 ("Headline result",'="On the sample data it surfaces about "&TEXT(Findings!E15/1000000,"$#,##0.0")&"M a year of recurring profit opportunity ("&TEXT(Findings!E16,"0.0%")&" of gross margin) from discount discipline and lower carrying cost, "&TEXT(Findings!E8/1000000,"$#,##0.0")&"M of cash tied up in excess stock, and an FY2027 forecast whose backtest error was "&TEXT(Forecast!F10,"0.0%")&"."'),
])
r=block(ig,r,"What I built",[
 ("Data model","A fact table (Sales_Data) joined to product and dealer dimensions with INDEX/MATCH, calculated columns for net sales, cost, margin, fiscal year and delivery lateness, and named ranges so every formula reads like a sentence."),
 ("Sales analysis","Segment and region trends with YoY, target attainment, discount-band profitability, customer mix, a dealer scorecard with automatic flags, and a product Pareto (ABC)."),
 ("Operations analysis","Months of supply by product, overstock and stockout flags, excess inventory value, suggested reorder quantities, and on-time delivery by region and segment."),
 ("Statistics","Distribution of order values, correlation tests (does discounting lift volume?), trend regression per segment, z-score outlier detection."),
 ("Forecast","Seasonal index x linear trend per segment, built from first principles in formulas, validated by training on FY2024-25 and testing on FY2026."),
 ("Pivots and automation (VBA)","A macro builds 7 pivot tables, 3 pivot charts and 4 slicers that filter every pivot at once from one shared cache. Other macros refresh everything, export the dashboard to PDF, split the data into one report workbook per dealer, run a data quality check, list discount exceptions and draft a KPI email."),
 ("Decision layer","Scenario model for volume, discount and cost; an interactive dashboard; a findings page where the narrative itself is formula-driven."),
])
r=block(ig,r,"Why I built it this way",[
 ("Start from the business problem","I read John Deere's public earnings commentary first, then designed the questions: where is demand falling and growing, is discounting buying anything, is inventory in the right machines, and what should we plan for next year."),
 ("Sales and operations together","Most analyses stop at sales. Here the demand numbers drive the inventory decisions, which is how a planner or dealer-performance analyst actually works."),
 ("Transparent forecast","I chose a method I can explain cell by cell and measured its error honestly, instead of a black-box function. A forecast nobody trusts does not get used."),
 ("Ends in actions","Each finding has an owner-ready action and a dollar value with its calculation shown, because analysis only matters when someone acts on it."),
 ("Excel on purpose","Excel is where business partners live. The same design moves to SQL and Power BI: fact and dimension tables, measures, then visuals."),
])
r=block(ig,r,"Impact it can create (on the sample data)",[
 ("Protect margin",'="A discount ceiling of "&TEXT(Findings!C3,"0%")&" would recover about "&TEXT(Findings!E7/1000000,"$#,##0.0")&"M of margin a year, if the orders still close."'),
 ("Free up cash",'="Clearing stock above the overstock threshold releases "&TEXT(Findings!E8/1000000,"$#,##0.0")&"M of cash and saves about "&TEXT(Findings!E14/1000000,"$#,##0.0")&"M a year in carrying cost."'),
 ("Protect revenue",'="Reordering the under-stocked growth products protects about "&TEXT(Findings!E9/1000000,"$#,##0.0")&"M of sales; fixing West deliveries protects the experience on "&TEXT(Findings!E10/1000000,"$#,##0.0")&"M of orders."'),
 ("Better planning","A baseline forecast with a known error band and ready scenarios replaces gut-feel targets and makes stocking decisions defensible."),
 ("Time saved","One refreshable workbook replaces separate manual reports for sales, inventory and delivery. (Say this as a design benefit; do not quote hours you have not measured.)"),
])
r=block(ig,r,"How it helps a data analyst",[
 ("Full workflow in one place","Requirements, data modelling, cleaning logic, KPI design, statistics, forecasting, visualisation and storytelling: the whole job, not just a chart."),
 ("Reusable template","Paste new orders into Sales_Data and every sheet updates. The KPI definitions are documented in the formulas."),
 ("Self-service for stakeholders","Dropdown filters and scenario switches let managers answer follow-up questions without another request."),
 ("Automates the routine","The recurring jobs of a sales analytics team (refresh, distribute, check data, chase exceptions) are one click each, so analyst time goes to analysis. Read the VBA before the interview so you can walk through any macro."),
 ("Shows judgment","It states assumptions, measures forecast error, and separates correlation from causation, which is what interviewers probe."),
])
r=block(ig,r,"Questions to expect, and how to answer",[
 ("Why not FORECAST.ETS or Python?","I wanted every step visible and auditable. I would benchmark against ETS or Prophet next and keep whichever has lower holdout error."),
 ("How accurate is the forecast?",'="Backtest MAPE on FY2026 was "&TEXT(Forecast!F10,"0.0%")&" for the total. By segment it ranges from "&TEXT(MIN(Forecast!F6:F9),"0.0%")&" to "&TEXT(MAX(Forecast!F6:F9),"0.0%")&". Segments whose trend changed direction are hardest, which is why I added scenarios."'),
 ("Does discounting drive volume?",'="In this data the monthly correlation between Large Ag discount and units is "&TEXT(Statistics!B29,"0.00")&", but discounts rose as the market fell, so the cycle confounds it. I would not claim causation; I would propose a controlled test with a few dealers."'),
 ("How would you scale it?","Move the fact table to SQL or a lakehouse, build measures in Power BI, schedule refresh, add row-level security by dealer. The model design stays the same."),
 ("How do you know the numbers are right?","Totals reconcile across sheets (segment total = region total = sum of monthly actuals), lookups are checked for errors, and inputs are separated from formulas by colour."),
 ("What would you do with real data?","Validate definitions with finance (net sales, margin), check data quality, confirm thresholds with inventory planners, and review findings with dealer-facing teams before recommending anything."),
])
r=block(ig,r,"Be honest about the limits",[
 ("Synthetic data","Patterns were designed into the sample. Say so. The value you are showing is the method and the thinking."),
 ("Simple forecast","Linear trend cannot predict a cycle turning point; no external drivers (crop prices, interest rates, housing starts) are included yet."),
 ("Inventory snapshot","One point in time, no on-order or in-transit stock, no lead times."),
 ("Next steps","Add external drivers, used-equipment inventory, customer-level retention analysis, and a Power BI version."),
])

# ---------- Macro guide ----------
mg["B1"]="Macro guide: pivot tables, slicers and VBA automation"; mg["B1"].font=font(size=16,bold=True,color=NAVY)
mg["B2"]="The macros live in module modCommandCenter (Alt+F11 to read the code). They were written for Excel 2013 or later on Windows."; mg["B2"].font=font(italic=True,color="5B6770")
def gblock(r,head,rows3,hdr=None):
    for c in (2,3,4): mg.cell(r,c).fill=PatternFill("solid",fgColor=TEAL)
    mg.cell(r,2,head).font=font(size=12,bold=True,color="FFFFFF"); r+=1
    if hdr:
        header(mg,r,hdr,2); r+=1
    for a,b,c in rows3:
        for j,v in enumerate((a,b,c)):
            x=mg.cell(r,2+j,v); x.font=font(bold=(j==0)); x.alignment=Alignment(wrap_text=True,vertical="top"); x.border=BORDER
        mg.row_dimensions[r].height=max(16,14*(1+len(c)//95)); r+=1
    return r+1
r=gblock(4,"First time: turn the macros on",[
 ("1. Unblock the file","Before opening","Files downloaded from chat or email are blocked by Windows. Right-click the file > Properties > tick Unblock > OK. (Or save it in a Trusted Location.)"),
 ("2. Enable content","When Excel opens it","Click Enable Content on the yellow bar."),
 ("3. Build the reports","Setup prompt","Answer Yes to the prompt, or press Alt+F8, choose BuildPivotReports and click Run. It adds six Pivot_ sheets and a Control_Panel sheet with buttons."),
 ("4. Save","Ctrl+S","Keep the .xlsm format so the macros and pivots are saved."),
])
r=gblock(r,"What the macros do",[
 ("BuildPivotReports","Pivot tables, charts, slicers","Creates 7 pivot tables on 6 sheets (segment, region, dealer scorecard, product hierarchy, monthly trend, customer mix, on-time delivery), 3 pivot charts, and 4 slicers (Fiscal Year, Region, Segment, Customer Type). All pivots share one cache, so one slicer click filters every pivot report. Safe to re-run."),
 ("RefreshAllReports","After new data","Refreshes every pivot table and recalculates all formulas once new orders are pasted into the Sales_Data table."),
 ("ResetAllFilters","Clean slate","Clears every slicer and returns the Dashboard to FY2026 / All regions."),
 ("ExportDashboardPDF","Share with managers","Saves Dashboard and Findings as one dated PDF in the workbook's folder."),
 ("CreateDealerReports","Distribute by dealer","Creates one workbook per dealer in a Dealer_Reports folder: that dealer's orders plus a KPI summary by fiscal year."),
 ("RunDataQualityCheck","Trust the numbers","Scans every order for duplicates, missing keys, formula errors, out-of-range units or discounts, negative margin and impossible dates, and writes a Data_Quality_Log sheet."),
 ("DiscountExceptionReport","Margin control","Asks for a discount level and lists every order at or above it, sorted, with totals."),
 ("EmailKpiSummary","Weekly update","Opens an Outlook draft with headline KPIs and the top five findings. It never sends by itself."),
 ("BuildControlPanel","Buttons","Rebuilds the Control_Panel sheet with one button per macro."),
],["Macro","Use it for","What it does"])
r=gblock(r,"If Excel removes or blocks the macros",[
 ("Fallback","modCommandCenter.bas","Press Alt+F11 > File > Import File > choose modCommandCenter.bas > close the editor > Alt+F8 > BuildPivotReports. Then save as .xlsm. The result is identical."),
 ("Note","If a macro stops","Note the error message and the highlighted line (Debug), then fix or report it. Written for Excel 2013 or later on Windows; the email macro needs Outlook."),
])

for w in (ws0, ig):
    for row in w.iter_rows():
        for c in row:
            if c.column==3 and c.value and str(c.value).startswith("="): c.alignment=Alignment(wrap_text=True,vertical="top"); w.row_dimensions[c.row].height=58

for w in wb.worksheets:
    w.page_setup.orientation='landscape'; w.page_setup.fitToWidth=1; w.page_setup.fitToHeight=0; w.sheet_properties.pageSetUpPr.fitToPage=True
if PUBLIC:
    for rr in range(1, ws0.max_row+1):
        if ws0.cell(rr,2).value == "Interview_Guide": ws0.delete_rows(rr); break
    wb.remove(ig)
wb.code_name = "ThisWorkbook"
for i,w in enumerate(wb.worksheets,1): w.sheet_properties.codeName = f"Sheet{i}"
print([w.title for w in wb.worksheets])
wb.active = 0
wb.save("Sales_Operations_Command_Center.xlsx" if PUBLIC else "Sales_Operations_Command_Center.xlsx")
print("saved")

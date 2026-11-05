import sys; sys.path.insert(0,'tools')
from budget_model import ITEMS, INCOME, RESERVE_RATE
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
HDR=PatternFill("solid",fgColor="1F4E79"); HF=Font(bold=True,color="FFFFFF"); B=Font(bold=True)
thin=Side(style="thin",color="BBBBBB"); BOX=Border(left=thin,right=thin,top=thin,bottom=thin)
WRAP=Alignment(wrap_text=True,vertical="top")
def hdr(ws,row,vals,widths=None):
    for i,v in enumerate(vals,1):
        c=ws.cell(row=row,column=i,value=v); c.fill=HDR; c.font=HF; c.alignment=Alignment(wrap_text=True,vertical="center")
    if widths:
        for i,w in enumerate(widths,1): ws.column_dimensions[ws.cell(row=1,column=i).column_letter].width=w
wb=Workbook()
# 說明
ws=wb.active; ws.title="說明"
lines=["本活動｜預算與經費來源試算表（2026-11-05 建立）","",
"【先講結論】","原案（方案一）約 NT$20 萬，做不到 10 萬。","建議採方案三（戶外音樂祭＋室內比賽），總支出約 NT$11.6 萬（含 10% 預備金）。","獎金由贊助支付、野餐墊用借的、樂團自帶鼓組時，社團要籌的錢約 NT$9.3 萬。","",
"【怎麼用】","1. 黃色格子是可以改的數字（數量、單價、收入金額、狀態）。","2. 改了之後，「方案比較與缺口」分頁會自動重算。","3. 收入的「狀態」用下拉選單：已確認／申請中／預估。只有「已確認」才是真的有錢。","4. 拿到廠商報價後，把單價換掉，並把「來源類型」改成「實際報價」，備註寫廠商名稱與日期。","",
"【來源類型是什麼意思】","實際報價：廠商正式給我們的價格（目前 0 筆）。","同類活動參考：網路上別人的公開價格，只能當參考。","純假設：我們自己估的，一定要換成報價。","",
"【重要提醒】","目前所有收入都還不是「已確認」。在「已確認」收入夠之前，不要簽不能退錢的合約。","場地費目前假設免收；若學校說要收，看「場地費風險」分頁。"]
for i,l in enumerate(lines,1):
    c=ws.cell(row=i,column=1,value=l); c.alignment=Alignment(wrap_text=True)
    if l.startswith("【") or i==1: c.font=B
ws.column_dimensions["A"].width=110
# 支出明細
ws=wb.create_sheet("支出明細")
cols=["編號","類別","品項","單位","方案一\n數量","方案一\n單價","方案一\n小計","方案二\n數量","方案二\n單價","方案二\n小計","方案三\n數量","方案三\n單價","方案三\n小計","來源類型","來源說明／連結","可否刪減","備註"]
hdr(ws,1,cols,[6,9,40,6,7,9,10,7,9,10,7,9,10,12,60,18,30])
YEL=PatternFill("solid",fgColor="FFF2CC")
r=2
for n,it in enumerate(ITEMS,1):
    cat,name,unit,a,b,c,st,note,cut,rem=it
    row=[f"E{n:02d}",cat,name,unit,a[0],a[1],f"=E{r}*F{r}",b[0],b[1],f"=H{r}*I{r}",c[0],c[1],f"=K{r}*L{r}",st,note,cut,rem]
    for j,v in enumerate(row,1):
        cell=ws.cell(row=r,column=j,value=v); cell.alignment=WRAP; cell.border=BOX
        if j in (5,6,8,9,11,12,14): cell.fill=YEL
        if j in (6,7,9,10,12,13): cell.number_format="#,##0"
    r+=1
last=r-1
ws.cell(row=r,column=3,value="小計").font=B
for col in "GJM": ws[f"{col}{r}"]=f"=SUM({col}2:{col}{last})"; ws[f"{col}{r}"].number_format="#,##0"; ws[f"{col}{r}"].font=B
sub=r; r+=1
ws.cell(row=r,column=3,value="預備金（比例可改 →）").font=B
ws.cell(row=r,column=4,value=RESERVE_RATE).fill=YEL; ws.cell(row=r,column=4).number_format="0%"
for col in "GJM": ws[f"{col}{r}"]=f"=ROUND({col}{sub}*$D${r},0)"; ws[f"{col}{r}"].number_format="#,##0"
res=r; r+=1
ws.cell(row=r,column=3,value="總支出").font=B
for col in "GJM": ws[f"{col}{r}"]=f"={col}{sub}+{col}{res}"; ws[f"{col}{r}"].number_format="#,##0"; ws[f"{col}{r}"].font=Font(bold=True,color="C00000")
tot=r
dv=DataValidation(type="list",formula1='"實際報價,同類活動參考,純假設,待確認"',allow_blank=True); ws.add_data_validation(dv); dv.add(f"N2:N{last}")
ws.conditional_formatting.add(f"N2:N{last}",FormulaRule(formula=[f'$N2="純假設"'],fill=PatternFill("solid",fgColor="FCE4D6")))
ws.conditional_formatting.add(f"N2:N{last}",FormulaRule(formula=[f'$N2="實際報價"'],fill=PatternFill("solid",fgColor="C6EFCE")))
ws.freeze_panes="D2"
# 收入
wi=wb.create_sheet("收入")
hdr(wi,1,["編號","經費來源","金額","狀態","說明","來源連結","負責組"],[6,34,10,10,70,50,10])
for n,(src,amt,st,desc,link,grp) in enumerate(INCOME,1):
    row=[f"I{n:02d}",src,amt,st,desc,link,grp]
    for j,v in enumerate(row,1):
        cell=wi.cell(row=n+1,column=j,value=v); cell.alignment=WRAP; cell.border=BOX
        if j in (3,4): cell.fill=YEL
        if j==3: cell.number_format="#,##0"
il=len(INCOME)+1
wi.cell(row=il+1,column=2,value="合計（全部，含預估）").font=B; wi[f"C{il+1}"]=f"=SUM(C2:C{il})"; wi[f"C{il+1}"].number_format="#,##0"
dv2=DataValidation(type="list",formula1='"已確認,申請中,預估"',allow_blank=False); wi.add_data_validation(dv2); dv2.add(f"D2:D{il}")
wi.conditional_formatting.add(f"D2:D{il}",FormulaRule(formula=['$D2="已確認"'],fill=PatternFill("solid",fgColor="C6EFCE")))
wi.conditional_formatting.add(f"D2:D{il}",FormulaRule(formula=['$D2="申請中"'],fill=PatternFill("solid",fgColor="FFEB9C")))
wi.conditional_formatting.add(f"D2:D{il}",FormulaRule(formula=['$D2="預估"'],fill=PatternFill("solid",fgColor="FCE4D6")))
# 缺口
wg=wb.create_sheet("方案比較與缺口")
hdr(wg,1,["項目","方案一\n原案全戶外","方案二\n戶外輕量","方案三（建議）\n戶外音樂祭＋室內比賽","怎麼算"],[34,16,16,22,60])
rows=[("總支出（含預備金）",[f"=支出明細!G{tot}",f"=支出明細!J{tot}",f"=支出明細!M{tot}"],"支出明細的總支出"),
("目標總支出",[100000]*3,"MD v2：總支出約 NT$100,000"),
("超出目標多少",["=B2-B3","=C2-C3","=D2-D3"],"正數＝超過目標"),
("已確認收入",[f'=SUMIFS(收入!$C$2:$C${il},收入!$D$2:$D${il},"已確認")']*3,"只算狀態＝已確認"),
("已確認＋申請中收入",[f'=SUMIFS(收入!$C$2:$C${il},收入!$D$2:$D${il},"已確認")+SUMIFS(收入!$C$2:$C${il},收入!$D$2:$D${il},"申請中")']*3,""),
("全部收入（含預估）",[f"=SUM(收入!$C$2:$C${il})"]*3,"預估不是真的錢"),
("缺口①：只看已確認",["=B2-B5","=C2-C5","=D2-D5"],"這是現在真正的缺口"),
("缺口②：已確認＋申請中",["=B2-B6","=C2-C6","=D2-D6"],""),
("缺口③：全部預估都到位",["=B2-B7","=C2-C7","=D2-D7"],"最樂觀情況還差多少"),
("是否在 10 萬以內",['=IF(B2<=B3,"是","否")','=IF(C2<=C3,"是","否")','=IF(D2<=D3,"是","否")'],""),
("精簡執行：扣掉可刪減的獎金、野餐墊、backline 後",[f"=ROUND((支出明細!G{sub}-支出明細!G12-支出明細!G18-支出明細!G7)*(1+支出明細!D{res}),0)",f"=ROUND((支出明細!J{sub}-支出明細!J12-支出明細!J18-支出明細!J7)*(1+支出明細!D{res}),0)",f"=ROUND((支出明細!M{sub}-支出明細!M12-支出明細!M18-支出明細!M7)*(1+支出明細!D{res}),0)"],"獎金改由贊助、野餐墊用借的、樂團自帶鼓組時，社團要籌的錢"),
]
for i,(lab,vals,how) in enumerate(rows,2):
    wg.cell(row=i,column=1,value=lab).font=B
    for j,v in enumerate(vals,2):
        c=wg.cell(row=i,column=j,value=v); c.number_format="#,##0"; c.border=BOX
    wg.cell(row=i,column=5,value=how).alignment=WRAP
for rr in (8,9,10):
    for col in "BCD":
        wg.conditional_formatting.add(f"{col}{rr}",FormulaRule(formula=[f"{col}{rr}>0"],font=Font(color="C00000",bold=True)))
wg.conditional_formatting.add("B11:D11",FormulaRule(formula=['B11="否"'],fill=PatternFill("solid",fgColor="F4CCCC")))
wg.conditional_formatting.add("B11:D11",FormulaRule(formula=['B11="是"'],fill=PatternFill("solid",fgColor="C6EFCE")))
n=len(rows)+3
notes=["【各方案差在哪】",
"方案一：照原案。週三晚搭 truss 舞台＋頂棚，兩天都在書卷廣場。下雨就兩天整組延期。",
"方案二：兩天都在戶外，但不用 truss，改用舞台板＋帳篷遮雨。比較便宜，但抗風雨能力比較差。",
"方案三：週四音樂祭＋市集在戶外（當天搭當天拆）；週五比賽改到室內（例如學生活動中心，待確認）。週四下雨就移到室內，不用整組延期。"]
for k,t in enumerate(notes): wg.cell(row=n+k,column=1,value=t).font=B if k==0 else Font()
# 獎金獎品
wp=wb.create_sheet("獎金獎品")
hdr(wp,1,["獎項","名額","現金（由贊助支付）","贊助品","贊助商","狀態","備註"],[12,6,16,30,20,10,60])
pr=[("冠軍",1,5000),("亞軍",1,3000),("季軍",1,2000),("人氣獎",1,1000)]
for i,(a,b,c) in enumerate(pr,2):
    for j,v in enumerate([a,b,c,"待洽談","待洽談","預估",""],1):
        cell=wp.cell(row=i,column=j,value=v); cell.border=BOX
        if j in (3,4,5,6): cell.fill=YEL
    wp.cell(row=i,column=3).number_format="#,##0"
wp["B6"]="合計"; wp["C6"]="=SUM(C2:C5)"; wp["C6"].number_format="#,##0"; wp["B6"].font=B
dv3=DataValidation(type="list",formula1='"已確認,申請中,預估"'); wp.add_data_validation(dv3); dv3.add("F2:F5")
tips=["【規則】","1. 學校與政府補助通常不補獎金，所以獎金只用贊助款。沒有贊助就改發獎狀＋贊助品。",
"2. 稅：比賽獎金屬「競技、競賽及機會中獎之獎金」。國內居住者扣繳率 10%；每次應扣稅額不超過 NT$2,000（獎金約 NT$20,000 以下）免先扣，但主辦方仍要列單申報。",
"   因此每個獎項現金都壓在 NT$20,000 以下；領獎人要填領據（含身分證字號，只給事務行政組，不公開）。",
"   非居住者（例如外籍生）一律扣 20%，沒有免扣門檻。來源：https://smart.businessweekly.com.tw/career/blog/3005282 （詳細請以財政部各類所得扣繳率標準為準，待確認）",
"3. 獎品（實物）也算獎金所得，價值要估算並一起申報（待向學校出納確認作法）。"]
for k,t in enumerate(tips): wp.cell(row=8+k,column=1,value=t)
# 場地費風險
wv=wb.create_sheet("場地費風險")
hdr(wv,1,["方案","需要的時段數（每時段 4 小時）","每時段費用","若全額收費","說明"],[12,18,12,14,80])
data=[("方案一",7,8800,"週三晚 1＋週四 3＋週五 3 時段"),("方案二",7,8800,"同方案一"),("方案三",3,8800,"只有週四用書卷廣場；週五室內場地費另查（待確認）")]
for i,(a,b,c,d) in enumerate(data,2):
    for j,v in enumerate([a,b,c,f"=B{i}*C{i}",d],1):
        cell=wv.cell(row=i,column=j,value=v); cell.border=BOX
        if j in (2,3): cell.fill=YEL; 
        if j in (3,4): cell.number_format="#,##0"
for k,t in enumerate(["【為什麼有這一頁】",
"每時段 NT$8,800 是 114 學年度書卷廣場的收費表價格（同類參考，115 學年度價格待確認）：https://www.sports.tku.edu.tw/wp-content/uploads/2025/08/%E6%B7%A1%E6%B0%B4%E6%A0%A1%E5%9C%92_114%E5%90%84%E9%A1%9E%E5%A0%B4%E5%9C%B0%E5%8F%8A%E8%A8%AD%E5%82%99%E5%99%A8%E6%9D%90%E5%80%9F%E7%94%A8%E7%B6%AD%E8%AD%B7%E8%B2%BB%E6%94%B6%E8%B2%BB%E4%B8%80%E8%A6%BD%E8%A1%A81140811%E4%BF%AE%E8%A8%82.pdf",
"收費表寫「學生社團及班級活動：奉准舉辦校內教職員工生參與之有關校內學生活動…」免收場地維護費。",
"但同系列收費表（蘭陽校園版）寫免收的條件是「未對外收取費用，亦無本校以外其他政府機關及民間團體補助經費」。",
"→ 本活動有跨校選手、攤販販售、也想申請新北市青年局補助，三件事都可能讓「免費」不成立。一定要書面問課外組與場地管理單位。",
"→ 每時段以 4 小時計，不足 4 小時算 1 時段。"]):
    wv.cell(row=6+k,column=1,value=t)
wb.save(sys.argv[1]); print("saved")

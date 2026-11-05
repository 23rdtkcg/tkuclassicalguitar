import sys, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter as L
D=lambda s: dt.date(2026 if int(s.split('/')[0])>=11 else 2027,int(s.split('/')[0]),int(s.split('/')[1]))
G={"總":"總召、副召","B2B":"廠商對接 B2B","賽":"賽事運行","B2C":"行銷公關 B2C","行":"事務行政"}
T=[
("總","確認 115-2 學期正式行事曆，決定主日期與雨備日","11/06","11/13","","4/29–30 可能落在期中考週（第 10 週），要先確認"),
("總","帶問題清單與課外組承辦人、指導老師面談","11/09","11/20","","問題清單見「本活動_問題清單與待決事項」"),
("總","幹部會議：選定方案（一／二／三）與社費撥用金額","11/16","11/20","總-01, 總-02, 行-01","建議方案三"),
("總","確定活動正式名稱","11/16","12/04","總-03",""),
("總","每週進度會議（看本表更新狀態）","11/09","05/14","","每週固定一天"),
("總","邀請週四演出社團（5～6 團）並確認演出名單","12/01","03/12","總-03","負責組待分配（見待決事項）"),
("總","經費停損點檢查：已確認＋申請中收入是否足夠","02/22","02/26","行-03, 行-05, B2B-07","不夠就照可行性評估的降規順序刪減"),
("總","工作人員說明會＋志工訓練","04/12","04/23","B2C-05",""),
("總","活動週總彩排、T-72 天氣預判、T-24 天氣決定","04/26","04/28","","T-72＝活動前 72 小時"),
("總","活動執行（週四音樂祭＋週五比賽）","04/29","04/30","",""),
("總","結案報告與交接","05/03","05/21","行-11",""),
("行","書面詢問場地費：跨校選手、攤販販售、外部補助會不會失去免費","11/06","11/20","","最優先"),
("行","預約場地（書卷廣場、海報街、草地；方案三加學生活動中心兩天）","11/16","12/11","總-03, 行-01","主日與雨備日都要訂"),
("行","申請學校社團活動經費補助（企畫書＋經費概算表）","12/01","01/08","總-03","送件時間待課外組確認"),
("行","打電話問新北市青年局：116 年第一期何時開放、能否先送件","11/06","11/13","","115 年第二期 11/16 截止，要先問清楚"),
("行","送新北市青年局補助申請（活動前至少一個月）","01/04","03/26","行-04","依青年局公告日期調整"),
("行","確認用電來源、電量、插座位置","12/01","01/15","行-02","問總務處／場地管理單位"),
("行","學務系統送活動申請（指導老師線上審核）","03/15","04/09","行-02","規定至少活動前 5 天（不含假日），我們提早送"),
("行","公共意外責任險詢價與投保","03/01","04/16","",""),
("行","野餐墊：先借、再找贊助、最後才買；做借還登記表","03/15","04/16","B2B-07","保證金 NT$100／張"),
("行","器材借用與清點（桌椅、延長線、對講機）","04/01","04/23","行-02",""),
("行","核銷（收據寫學校抬頭與統編）","05/03","05/21","總-10",""),
("行","刪除報名個資（依同意書寫的期限）","05/24","05/31","賽-13",""),
("B2B","PALT 接洽：音控、backline、通道數","11/09","12/04","","PALT 不接就要外租，每天多約 2.8 萬"),
("B2B","舞台詢價（路口影像＋另外 2 家，用詢價單）","11/16","12/18","總-03",""),
("B2B","評審候選名單與邀請","11/16","12/31","",""),
("B2B","評審確認（簽同意書）","01/04","01/29","B2B-03",""),
("B2B","做贊助提案書與贊助方案","11/23","12/18","",""),
("B2B","請教吉他社：金韶獎怎麼拿到影視局、青年局、一銀基金會支持","11/16","12/04","",""),
("B2B","企業贊助洽談（獎金、獎品、野餐墊、飲料）","12/07","03/12","B2B-05",""),
("B2B","攤販招商（招募表單、名單確認）","02/01","03/26","行-01","不收攤位費"),
("B2B","舞台／音響簽約（含延期與取消條款）","03/01","03/12","總-07, B2B-02",""),
("B2B","市集配置圖與攤販說明","03/29","04/16","B2B-08",""),
("B2B","確認廠商進場、撤場時間表","04/12","04/23","B2B-09",""),
("賽","比賽章程初稿","11/09","12/04","","依「本活動_比賽章程草案」"),
("賽","評分表、核分表、報名表欄位定稿","11/23","12/18","賽-01",""),
("賽","章程定稿（指導老師確認）","01/04","01/22","賽-01, 總-02",""),
("賽","Google 報名表建置與測試","01/11","02/05","賽-03",""),
("賽","開放報名（約 4 週）","02/22","03/19","賽-04, B2C-02",""),
("賽","資格審查","03/22","03/26","賽-05",""),
("賽","線上海選評分（3 位評審）","03/22","04/02","賽-06, B2B-04",""),
("賽","兩人核分＋公告晉級名單","04/05","04/09","賽-07",""),
("賽","晉級通知信；收伴奏音檔、曲目、器材需求","04/06","04/16","賽-08",""),
("賽","人氣獎投票表單＋QR code","03/29","04/23","",""),
("賽","選手說明會＋公開抽出場順序（錄影）","04/19","04/23","賽-09",""),
("賽","核分演練（用假資料跑一次）","04/19","04/23","賽-02",""),
("賽","成績公告、發獎、感謝信","05/03","05/14","總-10",""),
("B2C","視覺設計與宣傳素材","12/01","01/29","總-04",""),
("B2C","比賽網站上線（靜態網站）","12/07","02/05","賽-03","原型已做好，見 GitHub"),
("B2C","社群發文排程","01/04","04/30","B2C-01",""),
("B2C","跨校宣傳（各校吉他社、熱音社、課外組）","02/15","03/19","B2C-01",""),
("B2C","志工招募","02/22","03/26","",""),
("B2C","主持人選定＋主持稿","03/15","04/23","",""),
("B2C","校內觀眾宣傳（海報、預告貼文）","03/29","04/28","",""),
("B2C","現場攝影與社群即時發文","04/29","04/30","",""),
("B2C","照片影片整理與成果貼文","05/03","05/14","B2C-08","翻唱影片不公開上網（著作權）"),
]
cnt={}; rows=[]
for g,name,s,e,pre,note in T:
    cnt[g]=cnt.get(g,0)+1; code=f"{g}-{cnt[g]:02d}"
    rows.append((code,G[g],name,"",D(s),D(e),"未開始",pre,note))
wb=Workbook(); ws=wb.active; ws.title="工作總表"
HDR=PatternFill("solid",fgColor="1F4E79"); HF=Font(bold=True,color="FFFFFF")
thin=Side(style="thin",color="D9D9D9"); BOX=Border(left=thin,right=thin,top=thin,bottom=thin)
heads=["編號","組別","工作項目","負責人","開始日","截止日","狀態","前置工作","備註"]
w=[8,14,46,10,11,11,9,16,34]
for i,h in enumerate(heads,1):
    c=ws.cell(row=2,column=i,value=h); c.fill=HDR; c.font=HF; c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    ws.column_dimensions[L(i)].width=w[i-1]
ws.cell(row=1,column=1,value="本活動｜活動總表（一列一個工作。改開始日／截止日，右邊甘特圖會自動變色）").font=Font(bold=True,size=12)
start=dt.date(2026,11,2); weeks=[]
d=start
while d<=dt.date(2027,5,31): weeks.append(d); d+=dt.timedelta(days=7)
FC=10
for k,wd in enumerate(weeks):
    col=FC+k; cl=L(col)
    c=ws.cell(row=2,column=col,value=wd); c.number_format="m/d"; c.fill=HDR; c.font=Font(color="FFFFFF",size=8,bold=True); c.alignment=Alignment(horizontal="center",textRotation=90)
    ws.column_dimensions[cl].width=3.2
    if k==0 or wd.month!=weeks[k-1].month:
        ws.cell(row=1,column=col,value=f"{wd.month}月").font=Font(bold=True)
n=len(rows); last=2+n
for i,r in enumerate(rows,3):
    for j,v in enumerate(r,1):
        c=ws.cell(row=i,column=j,value=v); c.border=BOX; c.alignment=Alignment(vertical="center",wrap_text=(j in (3,8,9)))
        if j in (5,6): c.number_format="yyyy/mm/dd"
    for k in range(len(weeks)): ws.cell(row=i,column=FC+k).border=BOX
LC=L(FC+len(weeks)-1)
dv=DataValidation(type="list",formula1='"未開始,進行中,完成,卡住"',allow_blank=False,showErrorMessage=True,errorTitle="請用下拉選單",error="只能選：未開始／進行中／完成／卡住")
ws.add_data_validation(dv); dv.add(f"G3:G500")
dvd=DataValidation(type="date",operator="greaterThan",formula1="DATE(2026,1,1)",showErrorMessage=True,error="請輸入日期，例如 2027/03/01")
ws.add_data_validation(dvd); dvd.add("E3:F500")
rng=f"{L(FC)}3:{LC}500"; f0=L(FC)
ov=f"AND($E3<>\"\",{f0}$2<=$F3,{f0}$2+6>=$E3)"
ws.conditional_formatting.add(f"A3:I500",FormulaRule(formula=['$G3="卡住"'],fill=PatternFill("solid",fgColor="F4CCCC"),font=Font(color="990000",bold=True),stopIfTrue=False))
ws.conditional_formatting.add(rng,FormulaRule(formula=[f'AND({ov},$G3="卡住")'],fill=PatternFill("solid",fgColor="E06666"),stopIfTrue=True))
ws.conditional_formatting.add(rng,FormulaRule(formula=[f'AND({ov},$G3="完成")'],fill=PatternFill("solid",fgColor="B7B7B7"),stopIfTrue=True))
colors={"總召、副召":"8E7CC3","廠商對接 B2B":"F6B26B","賽事運行":"6FA8DC","行銷公關 B2C":"93C47D","事務行政":"FFD966"}
for g,cc in colors.items():
    ws.conditional_formatting.add(rng,FormulaRule(formula=[f'AND({ov},$B3="{g}")'],fill=PatternFill("solid",fgColor=cc),stopIfTrue=True))
ws.conditional_formatting.add(rng,FormulaRule(formula=[f"AND({f0}$2<=TODAY(),{f0}$2+6>=TODAY())"],fill=PatternFill("solid",fgColor="FFF2CC")))
ws.conditional_formatting.add("G3:G500",FormulaRule(formula=['$G3="完成"'],fill=PatternFill("solid",fgColor="D9EAD3")))
ws.conditional_formatting.add("G3:G500",FormulaRule(formula=['$G3="進行中"'],fill=PatternFill("solid",fgColor="CFE2F3")))
ws.conditional_formatting.add("F3:F500",FormulaRule(formula=['AND($F3<>"",$F3<TODAY(),$G3<>"完成")'],font=Font(color="C00000",bold=True)))
ws.freeze_panes=ws["D3"]; ws.auto_filter.ref=f"A2:I{last}"
# 里程碑
wm=wb.create_sheet("里程碑")
mh=["編號","日期","里程碑（關卡）","負責組","通過條件（全部做到才算過）","沒過怎麼辦","狀態"]
for i,h in enumerate(mh,1):
    c=wm.cell(row=1,column=i,value=h); c.fill=HDR; c.font=HF
for i,wdt in enumerate([6,11,26,14,60,44,9],1): wm.column_dimensions[L(i)].width=wdt
M=[("11/13","日期定案","總召、副召","查到 115-2 行事曆；主日期不在期中考週與前一週","改用替代日期（見可行性評估）"),
("11/20","方案定案","總召、副召","幹部會議選定方案；決定社費撥用金額；課外組回覆場地費","先停在方案三規格繼續詢價"),
("12/18","場地與報價","事務行政、B2B","場地申請送出；舞台至少 2 家書面報價；PALT 回覆","PALT 不接 → 外租報價"),
("01/29","比賽準備完成","賽事運行","章程定稿；3 位評審簽同意書；網站上線預告","延後開放報名一週，不可以沒評審就開報名"),
("02/22","開放報名","賽事運行、B2C","網站＋報名表測試過；跨校宣傳開始","—"),
("02/26","經費停損點","總召、副召","已確認＋申請中收入 ≥ 建議方案社團要籌金額的 70%","照降規順序刪減；仍不夠就延到下學期或縮成一天"),
("03/12","簽約","B2B","舞台／音響合約含延期與取消條款；簽約前總召確認已確認收入","不簽不能退錢的合約"),
("03/19","報名截止","賽事運行","有效報名 ≥ 15 組（純假設門檻，可改）","延長報名一週；仍不足就決賽改 8～10 組"),
("04/09","公告晉級","賽事運行","兩人核分一致；晉級＋備取名單公告","—"),
("04/23","活動前檢查","全體","選手說明會、志工訓練、保險投保、活動申請核准、班表確定","缺一項就不開演該項目"),
("04/28","天氣決定（T-24）","總召、副召","依氣象預報決定：照常／週四改室內／延期","依應變規則執行"),
("04/30","活動完成","全體","兩天活動完成、無重大事故","—"),
("05/21","結案","事務行政","核銷完成、成果報告、器材歸還、交接檔案","—")]
for i,(d0,a,b,c,e) in enumerate(M,2):
    for j,v in enumerate([f"M{i-1:02d}",D(d0),a,b,c,e,"未開始"],1):
        cell=wm.cell(row=i,column=j,value=v); cell.alignment=Alignment(wrap_text=True,vertical="top"); cell.border=BOX
        if j==2: cell.number_format="yyyy/mm/dd"
dvm=DataValidation(type="list",formula1='"未開始,進行中,完成,卡住"'); wm.add_data_validation(dvm); dvm.add("G2:G100")
wm.conditional_formatting.add("A2:G100",FormulaRule(formula=['$G2="卡住"'],fill=PatternFill("solid",fgColor="F4CCCC")))
wm.conditional_formatting.add("A2:G100",FormulaRule(formula=['$G2="完成"'],fill=PatternFill("solid",fgColor="D9EAD3")))
# 說明
wh=wb.create_sheet("怎麼用")
for i,t in enumerate(["【怎麼用這張表】","1. 一列就是一個工作。新增工作就在最下面加一列，編號照組別往下編。",
"2. 「狀態」只能用下拉選單：未開始／進行中／完成／卡住。選「卡住」整列會變紅。",
"3. 改「開始日」「截止日」，右邊的週格子會自動上色。顏色代表組別：紫＝總召、橘＝B2B、藍＝賽事、綠＝B2C、黃＝事務行政；灰＝已完成。",
"4. 淺黃色直欄是「這一週」。截止日過了還沒完成，截止日會變紅字。",
"5. 「前置工作」寫編號，意思是：那件事做完，這件事才能開始。",
"6. 日期目前以 2027/4/29–30 為主日期排。若日期改了，只要改相關工作的日期就好。",
"7. 每週開會前，各組長把自己組的狀態更新好。","","原本的雙週版甘特圖已備份為「古吉音樂季活動總表_備份｜2026-11-05」。"],1):
    wh.cell(row=i,column=1,value=t).font=Font(bold=(i==1))
wh.column_dimensions["A"].width=100
wb.move_sheet("怎麼用",offset=-2)
wb.active=1
wb.save(sys.argv[1]); print(len(rows),"tasks",len(weeks),"weeks")

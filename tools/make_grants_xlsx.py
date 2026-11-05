import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter as L
R=[
("淡江大學學生社團活動經費補助（課外組）","學校補助","適合","經課外組核定的學生社團；以自籌為主；迎新、聯誼、聚餐、遊覽、娛樂性活動不補",
 "主辦全國性競賽每場上限 NT$40,000；區域型上限 NT$10,000（本活動算哪一類待確認）","依要點核定項目（獎金、餐費能否補待確認）",
 "待確認","2026/12 中旬（先問承辦人）","企畫書、經費概算表；活動後：成果報告、收據（學校抬頭＋統編）","B2B","事務行政","未申請",
 "https://sa.tku.edu.tw/spirit/uploads//2024/12/淡江大學學生社團活動經費補助要點20260630.pdf","條文全文因網路限制未能直接讀取，內容來自搜尋摘要，送件前要再讀一次原文"),
("新北市青年局「協助青年綜合發展補助」（多元文化類）","政府補助","適合（最優先確認）","新北市境內大專校院、學生社團、系學會等；多元文化類含公共展演、公開競技",
 "團體最高 NT$100,000（需符合：學校、公益性／關懷弱勢、3 組以上跨校或跨組織合辦等條件之一）","場地租金、舞台設備、活動攝錄、文宣印刷、交通、保險、講師鐘點等基準表項目",
 "115 年：第一期 5/25 已用罄；第二期新聞稿寫 7/1～11/16。116 年：待公告","2026/11/13 前先打電話問；依公告最早一期送件（申請書範例寫活動前一個月）",
 "申請書、計畫書（目標、活動內容、經費規劃）、資格證明","B2B","事務行政","未申請",
 "https://www.ntpc.gov.tw/ch/home.jsp?id=e8ca970cde5c00e1&dataserno=524959edfc09b70ca690467fab025ccd",
 "⚠ 拿外部政府補助可能讓校內場地失去免費資格，送件前先問課外組"),
("文化部影視及流行音樂產業局","政府補助","待確認","目前公開的補助（如流行音樂產業行銷推廣補助）以業者為主",
 "待確認","待確認","115 年行銷推廣補助 1/5～2/24（業者用，不適用）","2026/12 前請教吉他社",
 "先問吉他社金韶獎用哪個計畫","B2B","—","未申請","https://www.bamid.gov.tw/News_Content.aspx?n=3483&s=249250",
 "金韶獎曾獲影視局補助（淡江時報），但沒查到是哪個計畫"),
("第一銀行文教基金會","基金會贊助","待確認","查不到公開的學生活動贊助申請辦法；基金會多自辦音樂會",
 "待確認","待確認","無公開截止日","2026/12 寄贊助提案信",
 "贊助提案書（1～2 頁）、活動企畫摘要","B2B","—","未申請","https://wantrich.chinatimes.com/news/20251117900464-420501",
 "曾參與金韶獎；請吉他社介紹窗口"),
("企業贊助：樂器行、音響行","企業贊助","適合","—","獎品、器材折扣、現金（待洽談）","獎金、獎品、backline",
 "—","2026/12～2027/03","贊助提案書、回饋方案（logo 露出等，先問課外組能給什麼）","B2B","B2C","未申請","—",""),
("企業贊助：飲料、在地商家","企業贊助","適合","—","實物為主","工作人員飲料、印 logo 野餐墊、人氣獎獎品",
 "—","2026/12～2027/03","贊助提案書、回饋方案","B2B","B2C","未申請","—",""),
("路口影像（luko.com.tw）","廠商合作","待確認","官網只寫攝影服務；有幹部曾接洽","待確認","攝影、舞台（待確認能否做）",
 "—","2026/11～12 詢價","舞台詢價單","B2B","—","未申請","https://luko.com.tw",""),
("台積電青年築夢計畫","企業計畫","不適用（今年）","第 11 屆報名 2026/10/15 已截止","總獎金 300 萬","—","已截止",
 "下一屆約 2027 秋（給下一屆參考）","—","B2B","—","不申請","https://www.tsmc-foundation.org/udreamer/2026/","時程趕不上 2027/4 的活動"),
("教育部青年署 青年百億海外圓夢基金","政府計畫","不適用","以海外計畫為主；藝文競賽請向主管部會申請","—","—",
 "116 年第一梯次 2026/10/30～11/30","—","—","—","—","不申請","https://cava.tw/topic/news/268090",""),
("新北市文化局 青年推廣文化計畫","政府計畫","不適用","是參與文化局的推廣專案，不是補助自辦活動","—","—","—","—","—","—","—","不申請",
 "https://activity.thu.edu.tw/upload/activity/news_upload/1778030568-115年新北市青年推廣文化計畫.pdf",""),
]
wb=Workbook(); ws=wb.active; ws.title="補助與贊助時程"
H=["編號","管道","類型","適不適合","申請條件","可能金額","補助／贊助什麼","官方截止日","建議送件日","要準備的文件","主責組","協助組","狀態","來源連結","備註"]
W=[6,30,10,12,36,30,28,28,26,32,8,9,9,40,34]
HDR=PatternFill("solid",fgColor="1F4E79")
thin=Side(style="thin",color="BBBBBB"); BOX=Border(left=thin,right=thin,top=thin,bottom=thin)
for i,h in enumerate(H,1):
    c=ws.cell(row=1,column=i,value=h); c.fill=HDR; c.font=Font(bold=True,color="FFFFFF"); c.alignment=Alignment(wrap_text=True)
    ws.column_dimensions[L(i)].width=W[i-1]
for r,row in enumerate(R,2):
    for j,v in enumerate((f"F{r-1:02d}",)+row,1):
        c=ws.cell(row=r,column=j,value=v); c.alignment=Alignment(wrap_text=True,vertical="top"); c.border=BOX
dv=DataValidation(type="list",formula1='"未申請,準備中,已送件,已核准,未通過,不申請"'); ws.add_data_validation(dv); dv.add("M2:M100")
ws.conditional_formatting.add("A2:O100",FormulaRule(formula=['$M2="已核准"'],fill=PatternFill("solid",fgColor="C6EFCE")))
ws.conditional_formatting.add("A2:O100",FormulaRule(formula=['$M2="不申請"'],font=Font(color="999999")))
ws.conditional_formatting.add("D2:D100",FormulaRule(formula=['LEFT($D2,2)="適合"'],fill=PatternFill("solid",fgColor="D9EAD3")))
ws.freeze_panes="C2"
wb.save(sys.argv[1]); print("ok")

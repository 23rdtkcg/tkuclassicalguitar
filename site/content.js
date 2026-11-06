/*
 * 本活動｜網站內容設定檔
 * ------------------------------------------------------------
 * 網站上所有「會變」的內容都在這裡改。改完存檔、推到 main 就會更新。
 * 規則：公告名單只放 編號／姓名或藝名／學校／曲目。
 *       不放學號、電話、Email、在學證明。
 */
window.SITE = {
  name: "本活動",                 // 活動名稱確定後改這裡
  subtitle: "暫名：歌神杯／古吉音樂季",
  organizer: "淡江大學共鳴箱古典吉他社",
  dates: "2027/04/29（四）音樂祭　2027/04/30（五）歌唱比賽決賽",
  dateNote: "日期待學校行事曆確認，正式日期以本網站公告為準。",
  finalDate: "2027-04-30",       // 決賽日期（YYYY-MM-DD），首頁倒數用；確定後改這裡
  contactEmail: "",              // 例如 tkcg.event@gmail.com（待決）
  instagram: "",                 // 例如 https://www.instagram.com/xxx

  // ---- 報名 ----
  registrationOpen: false,       // 2027/02/22 開放報名時改成 true
  registrationPeriod: "2027/02/22（一）～ 03/19（五）23:59",
  // Google 表單 →「傳送」→「< >」嵌入 → 複製 src 網址貼在這裡
  formEmbedUrl: "",

  rulesUrl: "",                  // 完整章程（Google 文件「發布到網路」的網址）

  timeline: [
    { date: "02/22–03/19", title: "開放報名", desc: "填報名表，貼上 1～2 分鐘的 YouTube 不公開影片連結。" },
    { date: "03/22–04/02", title: "線上海選", desc: "3 位評審看影片評分。" },
    { date: "04/09", title: "公告晉級", desc: "公告 10～12 組決賽名單與備取名單，並寄 Email。" },
    { date: "04/16", title: "晉級者交資料", desc: "決賽曲目、伴奏音檔、器材需求。" },
    { date: "04/19–04/23", title: "選手說明會", desc: "線上舉行，公開抽出場順序並錄影。" },
    { date: "04/30", title: "決賽", desc: "18:30 開始，現場一輪，當晚頒獎。" }
  ],

  // ---- 公告（04/09 後填）----
  finalists: [
    // { no: "01", name: "王小明", school: "淡江大學", song: "歌名／原唱" }
  ],
  waitlist: [
    // { no: "備1", name: "", school: "" }
  ],

  // ---- 得獎（決賽後填）----
  winners: [
    // { award: "冠軍", name: "", school: "", song: "" }
  ]
};

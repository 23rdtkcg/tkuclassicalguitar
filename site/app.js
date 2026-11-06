// 讀 content.js 的內容，填進頁面。一般不需要改這個檔案。
(function () {
  var S = window.SITE || {};
  function esc(s) { return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) { return ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]; }); }
  function $(id) { return document.getElementById(id); }
  document.querySelectorAll("[data-bind]").forEach(function (el) {
    var v = S[el.getAttribute("data-bind")];
    if (v) el.textContent = v;
  });
  if (S.name) document.title = S.name + "｜大專歌唱比賽";

  // 倒數：距離決賽幾天
  var left = null;
  if (S.finalDate) {
    var p = S.finalDate.split("-").map(Number), d = new Date(p[0], p[1] - 1, p[2]), n = new Date();
    left = Math.round((d - new Date(n.getFullYear(), n.getMonth(), n.getDate())) / 864e5);
  }
  if (left !== null && left > 0) {
    $("cover-num").textContent = left; $("cover-num").hidden = false;
    $("cover-sr").textContent = "距離決賽還有 " + left + " 天：";
    $("mast-count").innerHTML = '<span class="mz-num">' + left + "</span><small>天後決賽</small>"; $("mast-count").hidden = false;
  } else {
    $("cover-label").textContent = left === 0 ? "今天決賽" : (S.name || "本活動");
  }

  var tl = $("timeline-list");
  (S.timeline || []).forEach(function (t) {
    var li = document.createElement("li");
    if (/決賽/.test(t.title)) li.className = "is-key";
    li.innerHTML = '<span class="mz-date">' + esc(t.date) + "</span><div><b>" + esc(t.title) + "</b><p>" + esc(t.desc) + "</p></div>";
    tl.appendChild(li);
  });

  if (S.rulesUrl) {
    $("rules-link").innerHTML = '<a class="mz-btn" href="' + esc(S.rulesUrl) + '" target="_blank" rel="noopener">看完整章程（另開新分頁）</a>';
  }

  var fa = $("form-area");
  if (S.registrationOpen && S.formEmbedUrl) {
    fa.innerHTML = '<iframe src="' + esc(S.formEmbedUrl) + '" title="報名表" loading="lazy">載入中…</iframe>' +
      '<p class="muted">表單沒有出現？<a href="' + esc(S.formEmbedUrl.replace("?embedded=true", "")) + '" target="_blank" rel="noopener">點這裡在新分頁開啟</a>。</p>';
  } else {
    fa.innerHTML = '<div class="mz-empty"><strong>報名尚未開放</strong><p>開放時間：' + esc(S.registrationPeriod) + "。開放後報名表會出現在這裡。</p></div>";
  }

  function table(rows, cols, empty) {
    if (!rows || !rows.length) return '<p class="mz-empty">' + empty + "</p>";
    var h = '<table class="mz-table"><thead><tr>' + cols.map(function (c) { return "<th>" + c[1] + "</th>"; }).join("") + "</tr></thead><tbody>";
    rows.forEach(function (r) { h += "<tr>" + cols.map(function (c) { return "<td>" + esc(r[c[0]]) + "</td>"; }).join("") + "</tr>"; });
    return '<div class="mz-table-wrap">' + h + "</tbody></table></div>";
  }
  $("finalists").innerHTML =
    table(S.finalists, [["no", "編號"], ["name", "姓名"], ["school", "學校"], ["song", "曲目"]], "預計 04/09 公告。") +
    (S.waitlist && S.waitlist.length ? "<h4>備取</h4>" + table(S.waitlist, [["no", "順位"], ["name", "姓名"], ["school", "學校"]], "") : "");
  $("winners").innerHTML =
    table(S.winners, [["award", "獎項"], ["name", "姓名"], ["school", "學校"], ["song", "曲目"]], "決賽當晚頒獎後公告。");

  var c = [];
  if (S.contactEmail) c.push('Email：<a href="mailto:' + esc(S.contactEmail) + '">' + esc(S.contactEmail) + "</a>");
  if (S.instagram) c.push('<a href="' + esc(S.instagram) + '" target="_blank" rel="noopener">Instagram</a>');
  $("contact").innerHTML = c.join("　") || "聯絡方式即將公布";

  // 目錄分頁：捲到哪一段，就把那一格填上訊號色
  var links = Array.prototype.slice.call(document.querySelectorAll("#nav a"));
  if ("IntersectionObserver" in window) {
    var io = new IntersectionObserver(function (es) {
      es.forEach(function (en) {
        if (!en.isIntersecting) return;
        links.forEach(function (a) { if (a.getAttribute("href") === "#" + en.target.id) a.setAttribute("aria-current", "true"); else a.removeAttribute("aria-current"); });
      });
    }, { rootMargin: "-45% 0px -50% 0px" });
    links.forEach(function (a) { var s = document.querySelector(a.getAttribute("href")); if (s) io.observe(s); });
  }
})();

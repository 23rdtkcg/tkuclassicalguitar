// 讀 content.js 的內容，填進頁面。一般不需要改這個檔案。
(function () {
  var S = window.SITE || {};
  function esc(s) { return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) { return ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]; }); }
  document.querySelectorAll("[data-bind]").forEach(function (el) {
    var v = S[el.getAttribute("data-bind")];
    if (v) el.textContent = v;
  });
  if (S.name) document.title = S.name + "｜大專歌唱比賽";

  var tl = document.getElementById("timeline-list");
  (S.timeline || []).forEach(function (t) {
    var li = document.createElement("li");
    li.innerHTML = '<span class="when">' + esc(t.date) + '</span><div><strong>' + esc(t.title) + '</strong><p>' + esc(t.desc) + "</p></div>";
    tl.appendChild(li);
  });

  if (S.rulesUrl) {
    document.getElementById("rules-link").innerHTML = '<a class="btn btn-ghost" href="' + esc(S.rulesUrl) + '" target="_blank" rel="noopener">看完整章程</a>';
  }

  var fa = document.getElementById("form-area");
  if (S.registrationOpen && S.formEmbedUrl) {
    fa.innerHTML = '<iframe src="' + esc(S.formEmbedUrl) + '" title="報名表" loading="lazy">載入中…</iframe>' +
      '<p class="muted">表單沒有出現？<a href="' + esc(S.formEmbedUrl.replace("?embedded=true", "")) + '" target="_blank" rel="noopener">點這裡在新分頁開啟</a>。</p>';
  } else {
    fa.innerHTML = '<div class="placeholder"><strong>報名尚未開放</strong><p>開放時間：' + esc(S.registrationPeriod) + '。開放後報名表會出現在這裡。</p></div>';
  }

  function table(rows, cols, empty) {
    if (!rows || !rows.length) return '<p class="placeholder">' + empty + "</p>";
    var h = "<table><thead><tr>" + cols.map(function (c) { return "<th>" + c[1] + "</th>"; }).join("") + "</tr></thead><tbody>";
    rows.forEach(function (r) { h += "<tr>" + cols.map(function (c) { return "<td>" + esc(r[c[0]]) + "</td>"; }).join("") + "</tr>"; });
    return '<div class="table-wrap">' + h + "</tbody></table></div>";
  }
  document.getElementById("finalists").innerHTML =
    table(S.finalists, [["no", "編號"], ["name", "姓名"], ["school", "學校"], ["song", "曲目"]], "預計 04/09 公告。") +
    (S.waitlist && S.waitlist.length ? "<h4>備取</h4>" + table(S.waitlist, [["no", "順位"], ["name", "姓名"], ["school", "學校"]], "") : "");
  document.getElementById("winners").innerHTML =
    table(S.winners, [["award", "獎項"], ["name", "姓名"], ["school", "學校"], ["song", "曲目"]], "決賽當晚頒獎後公告。");

  var c = [];
  if (S.contactEmail) c.push('Email：<a href="mailto:' + esc(S.contactEmail) + '">' + esc(S.contactEmail) + "</a>");
  if (S.instagram) c.push('<a href="' + esc(S.instagram) + '" target="_blank" rel="noopener">Instagram</a>');
  document.getElementById("contact").innerHTML = c.join("　") || "聯絡方式即將公布";
})();

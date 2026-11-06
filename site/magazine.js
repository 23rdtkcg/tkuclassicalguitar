/*! 本活動雜誌風 magazine.js v1.0（選用）
 * 1. 主題切換：任何 <button data-theme-set="light|system|dark"> 都會生效，並記住選擇。
 * 2. 捲動進場：加上 class="mz-reveal" 的元素，捲到畫面內時「砸進定位」。
 * 3. 提示：MZ.toast("已複製")。
 * 4. 換頁重播：MZ.replay(元素, "mz-wipe")，讓跨頁色塊重新擦入。
 * 使用者若開啟「減少動態」，2 和 4 自動停用。
 */
(function () {
  "use strict";
  var root = document.documentElement;
  var RM = window.matchMedia && matchMedia("(prefers-reduced-motion: reduce)").matches;
  root.classList.add("mz-js");

  function applyTheme(t) {
    if (t === "system") root.removeAttribute("data-theme"); else root.setAttribute("data-theme", t);
    document.querySelectorAll("[data-theme-set]").forEach(function (b) { b.setAttribute("aria-pressed", String(b.dataset.themeSet === t)); });
    try { localStorage.setItem("mz-theme", t); } catch (e) {}
  }
  var saved = "system"; try { saved = localStorage.getItem("mz-theme") || "system"; } catch (e) {}
  applyTheme(saved);
  document.addEventListener("click", function (e) { var b = e.target.closest && e.target.closest("[data-theme-set]"); if (b) applyTheme(b.dataset.themeSet); });

  function reveal() {
    var els = document.querySelectorAll(".mz-reveal:not(.is-in)");
    if (RM || !("IntersectionObserver" in window)) { els.forEach(function (el) { el.classList.add("is-in"); }); return; }
    var io = new IntersectionObserver(function (es) {
      es.forEach(function (en) { if (en.isIntersecting) { en.target.classList.add("is-in"); io.unobserve(en.target); } });
    }, { rootMargin: "0px 0px -8% 0px" });
    els.forEach(function (el) { io.observe(el); });
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", reveal); else reveal();

  var toastEl, toastT;
  function toast(msg) {
    if (!toastEl) { toastEl = document.createElement("div"); toastEl.className = "mz-toast"; toastEl.setAttribute("role", "status"); toastEl.setAttribute("aria-live", "polite"); document.body.appendChild(toastEl); }
    toastEl.textContent = msg; toastEl.classList.add("is-show");
    clearTimeout(toastT); toastT = setTimeout(function () { toastEl.classList.remove("is-show"); }, 3200);
  }
  function replay(el, cls) { if (RM || !el) return; el.classList.remove(cls); void el.offsetWidth; el.classList.add(cls); }

  window.MZ = { theme: applyTheme, reveal: reveal, toast: toast, replay: replay };
})();

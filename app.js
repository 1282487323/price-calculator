/* =========================================================
   价格计算通用版 — 计算逻辑
   忠实复现 Excel「价格计算」sheet 的 B3–B10 公式
   ========================================================= */

// 默认值（与 Excel 示例一致）
const DEFAULTS = {
  price: 15.47,   // B2 输入价格
  origin: 13,     // B4 原定价
  deduct: 0.5,    // A1 减价常数
  rate: 0.85,     // B1 折扣率
  sub: 4,         // E3 减几
};

// 数值安全读取：空 / 非法 → 0
function num(v) {
  const n = parseFloat(v);
  return Number.isFinite(n) ? n : 0;
}

// 点击「输入价格 / 原定价」后，实时从剪贴板读取最近一次复制内容并填入
// - 每次点击都实时读取（复制新内容后再点即填入最新值）
// - 取剪贴板文本中的第一个数字（适配从 Excel/网页复制带单位或前后缀的文本）
// - 若剪贴板无数字，则回退为原始去空白文本
// - 权限被拒：置 _clipDenied 标记，后续点击不再尝试（避免反复弹窗）；其余异常静默忽略
//   注：HTTPS 等安全环境下，首次授权后浏览器会记住权限，后续点击不再弹窗
let _clipDenied = false;
async function pasteFromClipboard(inp) {
  if (_clipDenied) return;
  try {
    if (!navigator.clipboard || !navigator.clipboard.readText) return;
    const text = (await navigator.clipboard.readText()).trim();
    if (!text) return;
    const m = text.match(/-?\d+(?:\.\d+)?/);
    const val = m ? m[0] : text;
    inp.value = val;
    calculate();
    copyFinalPrice();
  } catch (e) {
    // 权限被拒 / 非安全上下文 / 用户取消：标记后跳过，避免每次点击都弹窗
    _clipDenied = true;
  }
}

// 自动复制「最终价格(B6)」：仅在输入价与原定价都已填好（有效正数）时执行
// - 复制 out-b6 的纯数字文本（fmt 输出，可直接粘贴进 Excel）
// - 成功复制后弹出轻提示「已复制最终价格：xx」；失败则静默（如浏览器拦截）
async function copyFinalPrice() {
  const b2 = num(document.getElementById("in-price").value);
  const b4 = num(document.getElementById("in-origin").value);
  if (!(b2 > 0) || !(b4 > 0)) return; // 两个价格未齐全不复制
  // 差价(B5) > 0（非「低于原定价」红字）时才自动复制
  const a1 = num(document.getElementById("in-deduct").value);
  const b1 = num(document.getElementById("in-rate").value);
  const e3 = num(document.getElementById("in-sub").value);
  const { B5 } = compute(b2, b4, a1, b1, e3);
  if (!(B5 > 0)) return;
  const el = document.getElementById("out-b6");
  if (!el) return;
  const text = (el.textContent || "").trim();
  if (!text || text === "—") return;
  try {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(text);
    } else {
      legacyCopy(text);
    }
    toast("✓ 已复制最终价格：" + text);
    flashFinal();
  } catch (e) {
    // 复制被拦截（如非安全上下文 / 无手势）：给出红色提示而非无反应
    toast("⚠ 复制失败，请手动复制（Ctrl+C）", false);
  }
}

// 兼容性兜底：老浏览器无 Clipboard API 时用 execCommand 复制
function legacyCopy(text) {
  const ta = document.createElement("textarea");
  ta.value = text;
  ta.style.position = "fixed";
  ta.style.opacity = "0";
  document.body.appendChild(ta);
  ta.focus();
  ta.select();
  try { document.execCommand("copy"); } catch (e) {}
  document.body.removeChild(ta);
}

// 轻量 toast 提示（自包含，无需改 CSS）：顶部居中、彩色、更明显
function toast(msg, ok = true) {
  let t = document.getElementById("__toast");
  if (!t) {
    t = document.createElement("div");
    t.id = "__toast";
    t.style.cssText = "position:fixed;left:50%;top:18px;transform:translateX(-50%) translateY(-8px);background:#10b981;color:#fff;padding:10px 18px;border-radius:10px;font-size:15px;font-weight:600;z-index:9999;box-shadow:0 6px 18px rgba(0,0,0,.25);opacity:0;transition:opacity .2s, transform .2s;pointer-events:none;";
    document.body.appendChild(t);
  }
  t.textContent = msg;
  t.style.background = ok ? "#10b981" : "#ee5253";
  void t.offsetWidth; // 强制重绘，保证过渡可见
  t.style.opacity = "1";
  t.style.transform = "translateX(-50%) translateY(0)";
  clearTimeout(toast._t);
  toast._t = setTimeout(() => { t.style.opacity = "0"; }, 1800);
}

// 复制成功后，让「最终价格」数字卡片高亮闪一下（自包含，无需改 CSS）
function flashFinal() {
  const el = document.getElementById("out-b6");
  if (!el) return;
  const card = el.closest(".rc-feature") || el.parentElement;
  const target = card || el;
  const prevBg = target.style.background;
  const prevColor = target.style.color;
  const prevTransition = target.style.transition;
  target.style.transition = "background .15s, color .15s";
  target.style.background = "var(--accent, #4d7cfe)";
  target.style.color = "#fff";
  setTimeout(() => {
    target.style.background = prevBg;
    target.style.color = prevColor;
    target.style.transition = prevTransition;
  }, 650);
}

// ===== 本地持久化（localStorage）：刷新/关闭后再开数据不丢 =====
const STORAGE_KEY = "price-calc-state-v1";

function saveState() {
  try {
    const state = {
      inputs: {
        price: document.getElementById("in-price").value,
        origin: document.getElementById("in-origin").value,
        deduct: document.getElementById("in-deduct").value,
        rate: document.getElementById("in-rate").value,
        sub: document.getElementById("in-sub").value,
      },
      toggles: {
        b7b8: document.getElementById("toggle-b7b8").checked,
        a1b1: document.getElementById("toggle-a1b1").checked,
      },
      skc: skcList,
      history: histList,
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch (e) { /* 隐私模式 / 存储不可用：静默忽略 */ }
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return false;
    const s = JSON.parse(raw);
    if (s.inputs) {
      document.getElementById("in-price").value = s.inputs.price ?? "";
      document.getElementById("in-origin").value = s.inputs.origin ?? "";
      document.getElementById("in-deduct").value = s.inputs.deduct ?? "";
      document.getElementById("in-rate").value = s.inputs.rate ?? "";
      document.getElementById("in-sub").value = s.inputs.sub ?? "";
    }
    if (s.toggles) {
      document.getElementById("toggle-b7b8").checked = !!s.toggles.b7b8;
      document.getElementById("toggle-a1b1").checked = !!s.toggles.a1b1;
    }
    if (Array.isArray(s.skc)) {
      skcList.length = 0;
      s.skc.forEach((it) => skcList.push(it));
    }
    if (Array.isArray(s.history)) {
      histList.length = 0;
      s.history.forEach((it) => histList.push(it));
    }
    return true;
  } catch (e) { return false; }
}

// 精确小数显示：去掉多余尾零，最多保留 maxDp 位
function fmt(n, maxDp = 4) {
  if (!Number.isFinite(n)) return "—";
  const r = Math.round(n * 1e6) / 1e6;
  // 若为整数则不带小数
  if (Math.abs(r - Math.round(r)) < 1e-9) return String(Math.round(r));
  return r.toFixed(maxDp).replace(/0+$/, "").replace(/\.$/, "");
}

// B5 差价取整：模拟 Excel
//   IF(LEFT((B4-B3),1)="0", LEFT((B4-B3),5), INT(B4-B3))
//   → 差值首位为 "0"（即 0<差值<1）时取前 5 个字符保留小数，
//     否则取 INT（向下取整）。
function roundDiff(diff) {
  const s = diff.toFixed(6);          // 避免浮点噪声
  if (s[0] === "0") {                 // "0.xxxx"
    return parseFloat(s.substring(0, 5)); // 前 5 位 → 3 位小数
  }
  return Math.floor(diff);            // 非 0 开头 → INT
}

// 取字符串右起第一位并转数字（对应 RIGHT(...,1) 的 VALUE）
function rightDigit(str) {
  return parseInt(str.slice(-1), 10);
}

// 核心计算：忠实复现 Excel B3–B10。可被首页与 SKC 统计页复用。
function compute(B2, B4, A1, B1, E3) {
  // B3 计算价格 = (B2 - A1) * B1
  const B3 = (B2 - A1) * B1;

  // B5 差价(已取整)
  const B5 = roundDiff(B4 - B3);

  // B6 最终价格 = IF(B5<1, B3, IF(B5<=4, B3+B5, B3+4))
  let B6;
  if (B5 < 1) B6 = B3;
  else if (B5 <= 4) B6 = B3 + B5;
  else B6 = B3 + 4;

  // B7 是否报活动
  let B7;
  if (B5 >= 4) B7 = "加4";
  else if (B5 === 3) B7 = "加3";
  else if (B5 === 2) B7 = "加2";
  else B7 = "不报活动";

  // B8 原报活动后价格
  let B8;
  if (B7 === "不报活动") B8 = B6;
  else B8 = B6 - rightDigit(B7) - E3;

  // B9 最终报活动
  let B9;
  if (B8 >= 8) {
    B9 = B7;
  } else if (B7 === "不报活动") {
    B9 = "不报活动";
  } else {
    const n = rightDigit(B7);
    if (8 - B8 <= 1) {
      B9 = (n - 1) >= 2 ? "加" + (n - 1) : "不报活动";
    } else if (8 - B8 <= 2) {
      B9 = (n - 1) > 2 ? "加" + (n - 2) : "不报活动";
    } else {
      B9 = "不报活动";
    }
  }

  // B10 最终报活动
  let B10;
  if (B9 === "不报活动") B10 = "不报活动";
  else B10 = "减" + (rightDigit(B9) + E3);

  return { B3, B5, B6, B7, B8, B9, B10 };
}

function calculate(toHistory = false) {
  const B2 = num(document.getElementById("in-price").value);  // 输入价格
  const B4 = num(document.getElementById("in-origin").value); // 原定价
  const A1 = num(document.getElementById("in-deduct").value); // 减价常数
  const B1 = num(document.getElementById("in-rate").value);   // 折扣率
  const E3 = num(document.getElementById("in-sub").value);    // 减几

  // 输入价/原定价未填入有效正数 → 结果区全部回到占位符「—」
  // （避免用默认值硬算出一个误导性的 -0.425）
  if (!(B2 > 0) || !(B4 > 0)) {
    resetOutputs();
    // 同时清掉还未完成的防抖记录，避免空输入进入历史
    clearTimeout(recordHistory._t);
    // 即便输入价/原定价为空，也要刷新 SKC 表头（减几已变时立即联动）
    renderSkc();
    saveState();
    return;
  }

  const { B3, B5, B6, B7, B8, B9, B10 } = compute(B2, B4, A1, B1, E3);

  // ----- 渲染 -----
  document.getElementById("out-b3").textContent = fmt(B3);

  // B5 差价：为负时明确标注“低于原定价”并标红
  const outB5 = document.getElementById("out-b5");
  if (B5 < 0) {
    outB5.classList.add("neg");
    outB5.innerHTML = '<span class="neg-prefix">低于原定价</span>' + fmt(B5, 3);
  } else {
    outB5.classList.remove("neg");
    outB5.textContent = fmt(B5, 3);
  }

  document.getElementById("out-b6").textContent = fmt(B6);
  document.getElementById("out-b8").textContent = fmt(B8);

  renderTag("out-b7", B7);
  if (B5 < 0) {
    renderUnavailable("out-b9");
  } else {
    renderTag("out-b9", B9 === "不报活动" ? B9 : (B9 + " / " + B10));
  }

  // 参数变化时，SKC 分类表同步刷新
  renderSkc();

  // 仅当显式调用 calculate(true) 或点击 立即计算 时入历史
  // —— 实时输入只刷新结果不入历史，避免试算过程的中间态污染历史
  if (toHistory) {
    recordHistory({ b2: B2, b4: B4, a1: A1, b1: B1, e3: E3, b5: B5, b6: B6, b9: B9, b10: B10 });
  }
  saveState();
}

// 活动状态标签渲染：报活动时拆为 两段小贴纸（加N / 减N）；不报活动 = 黑底白字
function renderTag(id, value) {
  const el = document.getElementById(id);
  el.classList.remove("is-active", "is-none", "is-split", "is-unavailable");
  if (value === "不报活动") {
    el.textContent = value;
    el.classList.add("is-none");
    return;
  }
  // value 形如 "加2 / 减6"
  const m = /^([^\/]+)\s*\/\s*(.+)$/.exec(value);
  if (m) {
    el.innerHTML =
      '<span class="tag-piece">' + m[1].trim() + '</span>' +
      '<span class="tag-slash">/</span>' +
      '<span class="tag-piece piece-second">' + m[2].trim() + '</span>';
    el.classList.add("is-active", "is-split");
  } else {
    el.textContent = value;
    el.classList.add("is-active");
  }
}

// 差价低于原定价（B5<0）：最终报活动卡片标红显示「此商品不可用」
function renderUnavailable(id) {
  const el = document.getElementById(id);
  el.classList.remove("is-active", "is-none", "is-split", "is-unavailable");
  el.classList.add("is-unavailable");
  el.textContent = "此商品不可用";
}

// 重置：清空三个主输入项；隐藏参数保持默认值；结果区回到占位符
function resetDefaults() {
  document.getElementById("in-price").value = "";
  document.getElementById("in-origin").value = "";
  document.getElementById("in-sub").value = "";

  document.getElementById("in-deduct").value = DEFAULTS.deduct;
  document.getElementById("in-rate").value = DEFAULTS.rate;

  resetOutputs();
  saveState();
}

// ===== 顶栏实时时钟 =====
// 周映射（0=周日）
const WEEK_ZH = ["周日", "周一", "周二", "周三", "周四", "周五", "周六"];

function updateClock() {
  const d = new Date();
  const pad2 = (n) => String(n).padStart(2, "0");
  const dateEl = document.getElementById("clockDate");
  const timeEl = document.getElementById("clockTime");
  if (!dateEl || !timeEl) return;
  const dateStr = `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())} · ${WEEK_ZH[d.getDay()]}`;
  const timeStr = `${pad2(d.getHours())}:${pad2(d.getMinutes())}:${pad2(d.getSeconds())}`;
  // 仅当变化时写入 DOM（避免无谓 reflow）
  if (dateEl.textContent !== dateStr) dateEl.textContent = dateStr;
  if (timeEl.textContent !== timeStr) timeEl.textContent = timeStr;
}

// 把结果区重置为占位符
function resetOutputs() {
  document.getElementById("out-b3").textContent = "—";

  const outB5 = document.getElementById("out-b5");
  outB5.textContent = "—";
  outB5.classList.remove("neg");

  document.getElementById("out-b6").textContent = "—";
  document.getElementById("out-b8").textContent = "—";

  const outB7 = document.getElementById("out-b7");
  outB7.textContent = "—";
  outB7.classList.remove("is-active", "is-none");

  const outB9 = document.getElementById("out-b9");
  outB9.textContent = "—";
  outB9.classList.remove("is-active", "is-none", "is-unavailable");
}

// 显隐开关：A1/B1 与 B7/B8
function applyToggle() {
  const showB7B8 = document.getElementById("toggle-b7b8").checked;
  document.getElementById("card-b7").style.display = showB7B8 ? "" : "none";
  document.getElementById("card-b8").style.display = showB7B8 ? "" : "none";

  const showA1B1 = document.getElementById("toggle-a1b1").checked;
  document.getElementById("card-a1").style.display = showA1B1 ? "" : "none";
  document.getElementById("card-b1").style.display = showA1B1 ? "" : "none";
  saveState();
}

/* =========================================================
   最近计算历史：仅记录有效计算（B2>0 且 B4>0），最多 8 条
   相同参数组合重复计算时去重并置顶
   ========================================================= */
const HIST_MAX = 8;
let histList = [];        // [{ ts, b2, b4, a1, b1, e3, b5, b6, b9, b10 }]
let histNewId = 0;        // 用于给刚追加的项打 .is-new 高亮

function sameKey(a, b) {
  return a.b2 === b.b2 && a.b4 === b.b4 && a.a1 === b.a1 && a.b1 === b.b1 && a.e3 === b.e3;
}

function recordHistory(p) {
  if (!(p.b2 > 0) || !(p.b4 > 0)) return; // 输入价/原定价为空时不记录
  // 防抖：用户连续输入时只在停顿后记录一次最终值
  clearTimeout(recordHistory._t);
  recordHistory._t = setTimeout(() => {
    const entry = { ts: Date.now(), ...p };
    histList.unshift(entry);
    if (histList.length > HIST_MAX) histList = histList.slice(0, HIST_MAX);
    histNewId = entry.ts;
    renderHistory();
    saveState();
  }, 600);
}

function renderHistory() {
  const list = document.getElementById("histList");
  const count = document.getElementById("histCount");
  if (!list) return;
  count.textContent = histList.length;

  if (histList.length === 0) {
    list.innerHTML = '<div class="history-empty">暂无记录</div>';
    return;
  }

  const pad = (n) => (n < 10 ? "0" + n : "" + n);
  list.innerHTML = histList.map((h) => {
    const d = new Date(h.ts);
    const time = pad(d.getHours()) + ":" + pad(d.getMinutes());
    const isActive = h.b9 !== "不报活动";
    const tagText = h.b9 === "不报活动" ? "不报活动" : (h.b9 + " · " + h.b10);
    const isNew = h.ts === histNewId ? " is-new" : "";
    // 差价徽章：负数用红底高亮（低于原定价）
    // 用 fmt 默认（不传 maxDp）：整数如 7 直接显示 7，小数如 0.225 显示 0.225
    // 不强制 toFixed(0) 避免 0.225 被四舍五入成 0 后看似空
    const diffNeg = typeof h.b5 === "number" && h.b5 < 0;
    const diffHtml =
      '<span class="hi-diff' + (diffNeg ? " neg" : "") + '" title="' +
      (diffNeg ? "低于原定价" : "差价(已取整)") + '">' +
      (diffNeg ? "↓ " : "Δ ") + fmt(h.b5) +
      "</span>";
    return (
      '<div class="history-item' + isNew + '">' +
        '<span class="hi-time">' + time + "</span>" +
        '<span class="hi-calc">' + fmt(h.b2) + '<span class="arrow">→</span>' + fmt(h.b4) + "</span>" +
        '<span class="hi-eq">=</span>' +
        '<span class="hi-final">' + fmt(h.b6) + "</span>" +
        diffHtml +
        '<span class="hi-tag ' + (isActive ? "t-active" : "t-none") + '">' + escapeHtml(tagText) + "</span>" +
        '<span class="hi-actions">' +
          '<button class="hi-fill" type="button" data-ts="' + h.ts + '" title="把当时的参数填回输入框">回填</button>' +
          '<button class="hi-del" type="button" data-ts="' + h.ts + '" title="删除">×</button>' +
        "</span>" +
      "</div>"
    );
  }).join("");

  list.querySelectorAll(".hi-fill").forEach((b) => {
    b.addEventListener("click", () => fillFromHistory(parseInt(b.dataset.ts, 10)));
  });
  list.querySelectorAll(".hi-del").forEach((b) => {
    b.addEventListener("click", () => removeHistory(parseInt(b.dataset.ts, 10)));
  });
}

// 把历史项的 5 个参数回填到输入框（触发实时计算）
function fillFromHistory(ts) {
  const h = histList.find((x) => x.ts === ts);
  if (!h) return;
  document.getElementById("in-price").value = h.b2;
  document.getElementById("in-origin").value = h.b4;
  document.getElementById("in-deduct").value = h.a1;
  document.getElementById("in-rate").value = h.b1;
  document.getElementById("in-sub").value = h.e3;
  calculate();
}

function removeHistory(ts) {
  histList = histList.filter((h) => h.ts !== ts);
  renderHistory();
}

function clearHistory() {
  if (histList.length === 0) return;
  histList = [];
  renderHistory();
  saveState();
}

/* =========================================================
   SKC 统计页：只录入 SKC 编码，
   价格/原定价/参数全部取自首页输入，按公式自动分类
   ========================================================= */
const skcList = []; // [{ code, b2, b4, a1, b1, e3 }] 锁定添加时的全部参数，避免后续改动影响分类

// 读取当前首页全部参数（与首页共享）
function getParams() {
  return {
    B2: num(document.getElementById("in-price").value),
    B4: num(document.getElementById("in-origin").value),
    A1: num(document.getElementById("in-deduct").value),
    B1: num(document.getElementById("in-rate").value),
    E3: num(document.getElementById("in-sub").value),
  };
}

// 把 B9（最终报活动）映射到分类列索引：0=减(n+4) 1=减(n+3) 2=减(n+2) 3=不报活动
function b9ToCol(b9) {
  if (b9 === "加4") return 0;
  if (b9 === "加3") return 1;
  if (b9 === "加2") return 2;
  return 3; // 不报活动
}

// 添加一条 SKC 的核心逻辑（只录入编码；锁定添加时的全部参数）
// 返回 { ok:true } 或 { ok:false, reason, b5? }
// 供单条录入（addSkc）复用
function addOneSkc(code) {
  const { B2, B4, A1, B1, E3 } = getParams();
  if (!(B2 > 0) || !(B4 > 0)) return { ok: false, reason: "noparams" };
  // 重复值提醒：code 已存在（无论它在四列中的哪一列）→ 拒绝
  if (skcList.some((s) => s.code === code)) return { ok: false, reason: "dup" };
  // 低于原定价拒绝：当前参数下 B5<0 时不允许添加
  const r = compute(B2, B4, A1, B1, E3);
  if (r.B5 < 0) return { ok: false, reason: "neg", b5: r.B5 };
  skcList.push({ code, b2: B2, b4: B4, a1: A1, b1: B1, e3: E3 });
  return { ok: true };
}

// 单条添加（绑定「＋ 添加」按钮 / 回车）
function addSkc() {
  const code = document.getElementById("skc-code").value.trim();
  if (!code) {
    alert("请先填写 SKC 编码");
    return;
  }
  const res = addOneSkc(code);
  if (!res.ok) {
    if (res.reason === "noparams") {
      alert("请先在上方「输入参数」填好「输入价格」和「原定价」");
    } else if (res.reason === "dup") {
      alert("SKC 编码【" + code + "】已经添加过，请勿重复录入");
    } else if (res.reason === "neg") {
      alert("差价低于原定价（B5 = " + fmt(res.b5) + "），无法添加 SKC。请调整输入参数后重试");
    }
    document.getElementById("skc-code").focus();
    return;
  }
  document.getElementById("skc-code").value = "";
  document.getElementById("skc-code").focus();
  renderSkc();
  saveState();
}

// 删除一条 SKC
function removeSkc(i) {
  skcList.splice(i, 1);
  renderSkc();
  saveState();
}

// 行内修改 SKC 编码：点「改」→ 编码变输入框 → 回车/失焦提交
function startEditSkc(i) {
  const item = skcList[i];
  if (!item) return;
  const board = document.getElementById("skcBoard");
  const span = board.querySelector('.skc-code[data-i="' + i + '"]');
  if (!span) return;

  const input = document.createElement("input");
  input.type = "text";
  input.className = "skc-edit-input";
  input.value = item.code;
  let done = false;
  const commit = () => {
    if (done) return;
    done = true;
    const code = input.value.trim();
    if (!code) {
      alert("SKC 编码不能为空");
      renderSkc();
      return;
    }
    if (skcList.some((s, j) => j !== i && s.code === code)) {
      alert("SKC 编码【" + code + "】已存在，请换一个");
      renderSkc();
      return;
    }
    item.code = code;
    renderSkc();
    saveState();
  };
  const cancel = () => {
    if (done) return;
    done = true;
    renderSkc();
  };

  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") { e.preventDefault(); commit(); }
    else if (e.key === "Escape") { e.preventDefault(); cancel(); }
  });
  input.addEventListener("blur", commit);

  span.replaceWith(input);
  input.focus();
  input.select();
}

// 一键清空所有 SKC（带确认）
function clearAllSkc() {
  if (skcList.length === 0) {
    alert("当前没有可清空的 SKC");
    return;
  }
  if (!confirm("确定要清空全部 " + skcList.length + " 条 SKC 吗？此操作不可撤销")) return;
  skcList.length = 0;
  renderSkc();
  saveState();
}

// SKC 面板放大/缩小：放大=显示全部；缩小=每列只显示前 SKC_COLLAPSED_MAX 个
const SKC_COLLAPSED_MAX = 3;
let skcExpanded = false;
function toggleSkcExpand() {
  skcExpanded = !skcExpanded;
  const btn = document.getElementById("skcExpandBtn");
  if (btn) {
    btn.textContent = skcExpanded ? "收起" : "展开";
    btn.title = skcExpanded ? "缩小：每列只显示部分" : "放大：显示全部";
  }
  renderSkc();
}

// 渲染分类表
function renderSkc() {
  const board = document.getElementById("skcBoard");
  if (!board) return;
  const { E3 } = getParams(); // 当前 E3 仅用于联动表头

  // 四列表头：减(n+4) / 减(n+3) / 减(n+2) / 不报活动（按当前 E3 联动）
  const headers = ["减" + (E3 + 4), "减" + (E3 + 3), "减" + (E3 + 2), "不报活动"];
  const cols = [[], [], [], []];

  // 每条 SKC 用自己锁定时的参数算 B9，再映射到列
  // —— 已添加的 SKC 不会被上方参数变化影响
  skcList.forEach((item, i) => {
    const r = compute(item.b2, item.b4, item.a1, item.b1, item.e3);
    const col = b9ToCol(r.B9);
    cols[col].push({ i, code: item.code, b6: r.B6, b10: r.B10 });
  });

  // 缩小状态：每列限制显示数量，超出部分用「+N 更多」占位
  const limit = skcExpanded ? Infinity : SKC_COLLAPSED_MAX;

  let html = "";
  headers.forEach((h, ci) => {
    const items = cols[ci];
    html += '<div class="skc-col skc-col-' + ci + '">';
    html += '<div class="skc-col-head">' + h + ' <span class="skc-count">' + items.length + "</span></div>";
    html += '<div class="skc-col-body">';
    if (items.length === 0) {
      html += '<div class="skc-empty">—</div>';
    } else {
      const shown = items.slice(0, limit);
      shown.forEach((it) => {
        html += '<div class="skc-chip">' +
          '<span class="skc-code" data-i="' + it.i + '">' + escapeHtml(it.code) + "</span>" +
          '<span class="skc-price">' + fmt(it.b6) + "</span>" +
          '<button class="skc-edit" type="button" data-i="' + it.i + '" title="修改编码">改</button>' +
          '<button class="skc-del" type="button" data-i="' + it.i + '" title="删除">×</button>' +
          "</div>";
      });
      // 缩小且有剩余：显示「+N 更多」占位条（点击展开）
      if (items.length > limit) {
        const rest = items.length - limit;
        html += '<button class="skc-more" type="button" title="显示全部">+' + rest + ' 更多</button>';
      }
    }
    html += "</div></div>";
  });
  board.innerHTML = html;
}

// 简单转义，避免 SKC 文本破坏结构
function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])
  );
}

// 复制分类表为四列 TSV，可直接粘到 Excel（Tab 分列、换行分行）
// 直接从 skcList 全量数据生成，与「放大/缩小」显示状态无关，折叠时也复制完整内容
async function copySkcToExcel() {
  if (!skcList || skcList.length === 0) {
    flashCopyBtn("没有可复制的 SKC");
    return;
  }
  const { E3 } = getParams(); // 当前减几，仅用于联动表头
  const headers = ["减" + (E3 + 4), "减" + (E3 + 3), "减" + (E3 + 2), "不报活动"];
  const cols = [[], [], [], []];

  // 每条 SKC 用自己锁定时的参数算 B9，再映射到列（与 renderSkc 完全一致）
  skcList.forEach((item) => {
    const r = compute(item.b2, item.b4, item.a1, item.b1, item.e3);
    const col = b9ToCol(r.B9);
    cols[col].push(item.code);
  });

  const maxRows = Math.max(1, ...cols.map((a) => a.length));
  const lines = [headers.join("\t")];
  for (let r = 0; r < maxRows; r++) {
    lines.push(cols.map((arr) => (r < arr.length ? arr[r] : "")).join("\t"));
  }
  const tsv = lines.join("\n");

  let ok = false;
  try {
    await navigator.clipboard.writeText(tsv);
    ok = true;
  } catch (e) {
    // 兜底：兼容旧浏览器或非安全上下文
    const ta = document.createElement("textarea");
    ta.value = tsv;
    ta.style.position = "fixed";
    ta.style.opacity = "0";
    document.body.appendChild(ta);
    ta.select();
    try {
      document.execCommand("copy");
      ok = true;
    } catch (_) {}
    document.body.removeChild(ta);
  }
  flashCopyBtn(ok ? "已复制，去 Excel 粘贴" : "复制失败");
}

function flashCopyBtn(text) {
  const btn = document.getElementById("skcCopyBtn");
  if (!btn) return;
  const prev = btn.textContent;
  btn.textContent = text;
  btn.classList.add("is-flashed");
  setTimeout(() => {
    btn.textContent = prev;
    btn.classList.remove("is-flashed");
  }, 1600);
}

// 绑定事件
document.addEventListener("DOMContentLoaded", () => {
  // 立即计算：算 + 入历史（每次点击都新增一条，相同参数也计入）
  document.getElementById("calcBtn").addEventListener("click", () => calculate(true));
  document.getElementById("resetBtn").addEventListener("click", resetDefaults);
  document.getElementById("printBtn").addEventListener("click", () => window.print());
  // 实时输入：只算，不入历史（避免试算过程的中间态污染历史）
  // 任意输入框按回车 = 立即计算（入历史）
  document.querySelectorAll(".field input").forEach((inp) => {
    inp.addEventListener("input", () => calculate(false));
    inp.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        calculate(true);
        inp.blur(); // 回车后失焦，避免反复触发自动清空
      }
    });
  });
  // 仅「输入价格」和「原定价」在真正点击/触摸字段后获得焦点时自动清空
  // 避免窗口切换回来时浏览器自动聚焦导致误清空
  ["in-price", "in-origin"].forEach((id) => {
    const inp = document.getElementById(id);
    if (!inp) return;
    const field = inp.closest(".field");
    if (!field) return;
    let lastPointerDown = 0;
    const POINTER_FOCUS_WINDOW = 300; // ms
    field.addEventListener("pointerdown", () => { lastPointerDown = Date.now(); });
    // 兼容不支持 pointerdown 的旧浏览器
    field.addEventListener("mousedown", () => { lastPointerDown = Date.now(); });
    inp.addEventListener("focus", () => {
      if (Date.now() - lastPointerDown <= POINTER_FOCUS_WINDOW) {
        inp.value = "";
        calculate();
        pasteFromClipboard(inp);
      }
    });
  });
  // 显隐开关
  document.getElementById("toggle-b7b8").addEventListener("change", applyToggle);
  document.getElementById("toggle-a1b1").addEventListener("change", applyToggle);
  // SKC 统计：添加按钮 + 回车快捷添加
  document.getElementById("skcAddBtn").addEventListener("click", addSkc);
  document.getElementById("skc-code").addEventListener("keydown", (e) => {
    if (e.key === "Enter") { e.preventDefault(); addSkc(); }
  });
  // SKC 一键复制到 Excel
  document.getElementById("skcCopyBtn").addEventListener("click", copySkcToExcel);
  // SKC 一键清空
  document.getElementById("skcClearBtn").addEventListener("click", clearAllSkc);
  // SKC 放大/缩小切换
  document.getElementById("skcExpandBtn").addEventListener("click", toggleSkcExpand);
  // SKC 面板按钮委托监听：删除 / 修改 / +N 更多
  // 用委托可免疫 renderSkc() 频繁重绘导致的点击丢失/旧 listener 失效
  document.getElementById("skcBoard").addEventListener("click", (e) => {
    const moreBtn = e.target.closest(".skc-more");
    if (moreBtn) {
      e.preventDefault();
      if (!skcExpanded) toggleSkcExpand();
      return;
    }
    const editBtn = e.target.closest(".skc-edit");
    if (editBtn) {
      e.preventDefault();
      startEditSkc(parseInt(editBtn.dataset.i, 10));
      return;
    }
    const delBtn = e.target.closest(".skc-del");
    if (delBtn) {
      e.preventDefault();
      removeSkc(parseInt(delBtn.dataset.i, 10));
      return;
    }
  });
  // 历史：清空按钮
  document.getElementById("histClearBtn").addEventListener("click", clearHistory);
  loadState();        // 先恢复上次保存的输入/开关/SKC/历史
  applyToggle();      // 再按恢复的开关态刷新卡片显隐
  calculate(false);   // 用恢复后的参数渲染结果与 SKC 表头
  renderSkc();
  renderHistory();
  // 启动实时时钟
  updateClock();
  setInterval(updateClock, 1000);
  // PWA：注册 SW + 安装按钮逻辑
  initPwa();
  // 主题切换（彩色 / 蓝色）
  initTheme();
  // 打印前自动展开 SKC（确保全部编码都打印出来），打印后恢复原状
  let prevExpanded = null;
  window.addEventListener("beforeprint", () => {
    prevExpanded = skcExpanded;
    if (!skcExpanded) { skcExpanded = true; renderSkc(); }
  });
  window.addEventListener("afterprint", () => {
    if (prevExpanded !== null && prevExpanded !== skcExpanded) {
      skcExpanded = prevExpanded;
      renderSkc();
    }
    prevExpanded = null;
  });
});

/* =========================================================
   PWA：注册 Service Worker + 安装按钮（beforeinstallprompt）
   - 安装前：黄色「安装应用」按钮常显，呼吸灯提示
   - 安装中：调用 prompt()
   - 安装后 / 已经是 standalone：按钮自动隐藏
   ========================================================= */
function isStandalone() {
  return (
    window.matchMedia("(display-mode: standalone)").matches ||
    window.matchMedia("(display-mode: minimal-ui)").matches ||
    window.navigator.standalone === true // iOS Safari
  );
}

function initPwa() {
  // 注册 Service Worker（仅在 https / localhost 下生效；本地 file:// 不报错但会被拒）
  if ("serviceWorker" in navigator) {
    window.addEventListener("load", () => {
      navigator.serviceWorker.register("./sw.js").catch((e) => console.warn("SW 注册失败", e));
    });
  }

  const btn = document.getElementById("installBtn");
  if (!btn) return;

  // 已是 standalone（用户从桌面图标进入）→ 直接不显示
  if (isStandalone()) {
    btn.hidden = true;
    return;
  }

  // 保存 beforeinstallprompt 事件，点按钮时调用 prompt()
  let deferredPrompt = null;
  window.addEventListener("beforeinstallprompt", (e) => {
    e.preventDefault();
    deferredPrompt = e;
    btn.hidden = false;
  });

  // 用户同意安装后：appinstalled 事件触发，隐藏按钮
  window.addEventListener("appinstalled", () => {
    btn.hidden = true;
    deferredPrompt = null;
  });

  btn.addEventListener("click", async () => {
    if (!deferredPrompt) {
      // 浏览器还没派发 beforeinstallprompt（多见于隐身模式或已被用户忽略多次）
      btn.title = "请通过浏览器菜单中的「添加到主屏幕/安装应用」";
      return;
    }
    btn.disabled = true;
    btn.textContent = "安装中…";
    try {
      deferredPrompt.prompt();
      const choice = await deferredPrompt.userChoice;
      if (choice && choice.outcome === "accepted") {
        btn.hidden = true; // 即时隐藏，等 appinstalled 再确认
      } else {
        btn.textContent = "安装应用";
      }
    } catch (e) {
      btn.textContent = "安装应用";
    } finally {
      btn.disabled = false;
      deferredPrompt = null;
    }
  });
}

/* =========================================================
   主题切换：彩色版（Neobrutalism 多色） / 蓝色版（黑白+蓝强调）
   - 初始主题由 <head> 内联脚本写入 <html data-theme>，避免首屏闪烁
   - 选择持久化到 localStorage('pc_theme')，下次打开沿用
   ========================================================= */
function initTheme() {
  const KEY = "pc_theme";
  const btn = document.getElementById("themeBtn");
  const getTheme = () => document.documentElement.getAttribute("data-theme") || "blue";
  const applyBtn = () => {
    if (btn) btn.textContent = getTheme() === "color" ? "蓝色版" : "彩色版";
  };
  applyBtn();
  if (!btn) return;
  btn.addEventListener("click", () => {
    const next = getTheme() === "color" ? "blue" : "color";
    document.documentElement.setAttribute("data-theme", next);
    try { localStorage.setItem(KEY, next); } catch (e) {}
    applyBtn();
  });
}

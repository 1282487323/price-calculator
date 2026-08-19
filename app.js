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
  renderTag("out-b9", B9 === "不报活动" ? B9 : (B9 + " / " + B10));

  // 参数变化时，SKC 分类表同步刷新
  renderSkc();

  // 仅当显式调用 calculate(true) 或点击 立即计算 时入历史
  // —— 实时输入只刷新结果不入历史，避免试算过程的中间态污染历史
  if (toHistory) {
    recordHistory({ b2: B2, b4: B4, a1: A1, b1: B1, e3: E3, b5: B5, b6: B6, b9: B9, b10: B10 });
  }
}

// 活动状态标签渲染：报活动时拆为 两段小贴纸（加N / 减N）；不报活动 = 黑底白字
function renderTag(id, value) {
  const el = document.getElementById(id);
  el.classList.remove("is-active", "is-none", "is-split");
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

// 重置：清空三个主输入项；隐藏参数保持默认值；结果区回到占位符
function resetDefaults() {
  document.getElementById("in-price").value = "";
  document.getElementById("in-origin").value = "";
  document.getElementById("in-sub").value = "";

  document.getElementById("in-deduct").value = DEFAULTS.deduct;
  document.getElementById("in-rate").value = DEFAULTS.rate;

  resetOutputs();
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
  outB9.classList.remove("is-active", "is-none");
}

// 显隐开关：A1/B1 与 B7/B8
function applyToggle() {
  const showB7B8 = document.getElementById("toggle-b7b8").checked;
  document.getElementById("card-b7").style.display = showB7B8 ? "" : "none";
  document.getElementById("card-b8").style.display = showB7B8 ? "" : "none";

  const showA1B1 = document.getElementById("toggle-a1b1").checked;
  document.getElementById("card-a1").style.display = showA1B1 ? "" : "none";
  document.getElementById("card-b1").style.display = showA1B1 ? "" : "none";
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

// 添加一条 SKC（只录入编码；锁定添加时的全部参数）
function addSkc() {
  const code = document.getElementById("skc-code").value.trim();

  if (!code) {
    alert("请先填写 SKC 编码");
    return;
  }
  const { B2, B4, A1, B1, E3 } = getParams();
  if (!(B2 > 0) || !(B4 > 0)) {
    alert("请先在上方「输入参数」填好「输入价格」和「原定价」");
    return;
  }

  // 1) 重复值提醒：code 已存在（无论它在四列中的哪一列）→ 拒绝
  if (skcList.some((s) => s.code === code)) {
    alert("SKC 编码【" + code + "】已经添加过，请勿重复录入");
    document.getElementById("skc-code").focus();
    return;
  }

  // 2) 低于原定价拒绝：当前参数下 B5<0 时不允许添加
  const r = compute(B2, B4, A1, B1, E3);
  if (r.B5 < 0) {
    alert("差价低于原定价（B5 = " + fmt(r.B5) + "），无法添加 SKC。请调整输入参数后重试");
    return;
  }

  // 锁定：本次添加瞬间的全部参数快照，后续用户改动参数不影响本条
  skcList.push({ code, b2: B2, b4: B4, a1: A1, b1: B1, e3: E3 });
  document.getElementById("skc-code").value = "";
  document.getElementById("skc-code").focus();

  renderSkc();
}

// 删除一条 SKC
function removeSkc(i) {
  skcList.splice(i, 1);
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

  let html = "";
  headers.forEach((h, ci) => {
    const items = cols[ci];
    html += '<div class="skc-col skc-col-' + ci + '">';
    html += '<div class="skc-col-head">' + h + ' <span class="skc-count">' + items.length + "</span></div>";
    html += '<div class="skc-col-body">';
    if (items.length === 0) {
      html += '<div class="skc-empty">—</div>';
    } else {
      items.forEach((it) => {
        html += '<div class="skc-chip">' +
          '<span class="skc-code">' + escapeHtml(it.code) + "</span>" +
          '<span class="skc-price">' + fmt(it.b6) + "</span>" +
          '<button class="skc-del" type="button" data-i="' + it.i + '" title="删除">×</button>' +
          "</div>";
      });
    }
    html += "</div></div>";
  });
  board.innerHTML = html;

  // 绑定删除按钮
  board.querySelectorAll(".skc-del").forEach((btn) => {
    btn.addEventListener("click", () => removeSkc(parseInt(btn.dataset.i, 10)));
  });
}

// 简单转义，避免 SKC 文本破坏结构
function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])
  );
}

// 复制分类表为四列 TSV，可直接粘到 Excel（Tab 分列、换行分行）
async function copySkcToExcel() {
  const board = document.getElementById("skcBoard");
  if (!board) return;
  const cols = Array.from(board.querySelectorAll(".skc-col"));
  if (cols.length === 0) return;

  // 提取每列的列头与编码列表
  const headers = cols.map((c) =>
    c.querySelector(".skc-col-head").firstChild.textContent.trim()
  );
  const items = cols.map((c) =>
    Array.from(c.querySelectorAll(".skc-code")).map((s) => s.textContent.trim())
  );

  const maxRows = Math.max(1, ...items.map((a) => a.length));
  const lines = [];
  lines.push(headers.join("\t"));
  for (let r = 0; r < maxRows; r++) {
    lines.push(items.map((arr) => (r < arr.length ? arr[r] : "")).join("\t"));
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
  flashCopyBtn(ok ? "✓ 已复制，去 Excel 粘贴" : "✗ 复制失败");
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
  // 仅「输入价格」和「原定价」每次 focus 自动清空
  ["in-price", "in-origin"].forEach((id) => {
    const inp = document.getElementById(id);
    if (!inp) return;
    inp.addEventListener("focus", () => {
      inp.value = "";
      calculate();
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
  // 历史：清空按钮
  document.getElementById("histClearBtn").addEventListener("click", clearHistory);
  applyToggle();
  resetOutputs();
  renderSkc();
  renderHistory();
});

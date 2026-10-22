// 수식·모니터링 인자 편집기 — browser side. All parsing and file writing is done
// by the local Python server (monitoring_plan.py), so preview and exports match.
"use strict";

const STORE_KEY = "innergen-equation-editor-v1";
const KIND_LABEL = { calculated: "산출", monitoring: "모니터링", fixed: "고정" };
const DETAIL = [
  ["source", "데이터 출처"], ["method", "측정 방법/절차"], ["frequency", "모니터링 주기"],
  ["qaqc", "QA/QC 절차"], ["purpose", "데이터 목적"], ["note", "기타 의견"], ["remark", "비고(요약표)"],
];

const state = {
  title: "",
  font: "Pretendard",
  options: { form: "kcs", number_equations: true, show_values: false, label_prefix: "식 " },
  text: "",
  parameters: [],
  derived: null,
  open: new Set(),        // parameter keys with the detail row expanded
};

const $ = (id) => document.getElementById(id);
const textEl = $("text");
let timer = null;
let seq = 0;

// ---------------------------------------------------------------- utilities
function esc(s) {
  return String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
}
function toast(msg, isError = false) {
  const el = document.createElement("div");
  el.className = "toast" + (isError ? " error" : "");
  el.textContent = msg;
  document.body.appendChild(el);
  setTimeout(() => el.remove(), isError ? 6000 : 2500);
}
function save() {
  try {
    localStorage.setItem(STORE_KEY, JSON.stringify({
      title: state.title, font: state.font, options: state.options, text: state.text, parameters: state.parameters,
    }));
  } catch (_) { /* private mode etc. */ }
}
function restore() {
  try {
    const raw = localStorage.getItem(STORE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch (_) { return null; }
}
function requestBody() {
  return { text: state.text, parameters: state.parameters, title: state.title, font: state.font, options: state.options };
}
function planToText(plan) {   // mirrors monitoring_plan.plan_to_text
  const lines = [];
  const numbered = (plan.options || {}).number_equations !== false;
  for (const sec of plan.sections || []) {
    if (lines.length) lines.push("");
    lines.push((sec.level === 2 ? "## " : "# ") + (sec.name || ""));
    if (sec.description) lines.push("> " + sec.description);
    for (const eq of sec.equations || []) {
      lines.push((eq.text || eq.latex || "") + (eq.label && !numbered ? ` (${eq.label})` : ""));
    }
  }
  return lines.join("\n") + "\n";
}

// ---------------------------------------------------------------- load state
function applyPlan(plan, text) {
  state.title = plan.title || "";
  state.font = plan.font || "Pretendard";
  state.options = { ...state.options, ...(plan.options || {}) };
  state.parameters = (plan.parameters || []).map((p) => ({ ...p }));
  state.text = text ?? planToText(plan);
  state.open.clear();
  syncControls();
  derive();
}
function syncControls() {
  $("title").value = state.title;
  textEl.value = state.text;
  $("opt-number").checked = state.options.number_equations !== false;
  $("opt-values").checked = !!state.options.show_values;
  $("opt-form").value = state.options.form === "moe" ? "moe" : "kcs";
  const font = $("opt-font");
  if (![...font.options].some((o) => o.value === state.font)) font.add(new Option(state.font));
  font.value = state.font;
}

// ---------------------------------------------------------------- derive
function scheduleDerive(delay = 300) {
  clearTimeout(timer);
  timer = setTimeout(derive, delay);
}
async function derive() {
  const mine = ++seq;
  const sentText = state.text;
  $("status").textContent = "계산 중…";
  let data;
  try {
    const res = await fetch("/api/derive", {
      method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(requestBody()),
    });
    data = await res.json();
    if (!res.ok) throw new Error(data.error || res.statusText);
  } catch (err) {
    $("status").textContent = "서버 연결 실패: " + err.message;
    return;
  }
  if (mine !== seq) return;          // a newer request is on its way
  state.parameters = data.plan.parameters || [];
  absorbParamLines(sentText);
  state.derived = data.derived;
  $("status").textContent = "";
  render();
  save();
}
// Parameter rows typed as "기호 | 정의 | …" are now in the dictionary; drop them from the text.
function absorbParamLines(sentText) {
  const piped = new Set(sentText.split("\n").filter((l) => l.includes("|") && !l.trim().startsWith("//")));
  if (!piped.size) return;
  const before = textEl.value;
  const caret = textEl.selectionStart;
  const lines = before.split("\n");
  let removedBefore = 0;
  let pos = 0;
  const kept = [];
  for (const line of lines) {
    const end = pos + line.length + 1;
    if (piped.has(line)) {
      if (end <= caret) removedBefore += line.length + 1;
    } else {
      kept.push(line);
    }
    pos = end;
  }
  if (kept.length === lines.length) return;
  const next = kept.join("\n").replace(/\n{3,}/g, "\n\n");
  textEl.value = next;
  state.text = next;
  const at = Math.max(0, Math.min(next.length, caret - removedBefore));
  if (document.activeElement === textEl) textEl.setSelectionRange(at, at);
  toast(`인자 ${piped.size}개를 인자 사전에 반영했습니다`);
}

// ---------------------------------------------------------------- render
function render() {
  const d = state.derived;
  if (!d) return;
  const byKey = new Map(d.parameters.map((p) => [p.key, p]));
  renderPreview(d, byKey);
  renderIssues(d.issues);
  renderGrid(d.parameters);
}

function renderPreview(d, byKey) {
  const out = [];
  if (!d.sections.length) {
    out.push('<p class="status">왼쪽에 수식을 입력하면 여기에 편집 가능한 수식 형태로 보입니다.</p>');
  }
  const showValues = !!d.options.show_values;
  for (const sec of d.sections) {
    const h = sec.level === 2 ? "h4" : "h3";
    out.push(`<div class="sec"><${h}>${esc(sec.name)}</${h}>`);
    if (sec.description) out.push(`<p class="desc">${esc(sec.description)}</p>`);
    for (const eq of sec.equations) {
      if (eq.error) {
        out.push(`<div class="eq bad"><div><div class="src">${esc(eq.text)}</div><div class="err">${esc(eq.error)}</div></div><span class="label"></span></div>`);
      } else {
        out.push(`<div class="eq">${eq.mathml}<span class="label">${eq.label ? "(" + esc(eq.label) + ")" : ""}</span></div>`);
      }
    }
    if (sec.symbols.length) {
      out.push(`<table class="sym"><thead><tr><th>기호</th><th>정의</th>${showValues ? "<th>값</th>" : ""}<th>단위</th></tr></thead><tbody>`);
      for (const key of sec.symbols) {
        const p = byKey.get(key) || {};
        const nameCls = p.name ? "" : " missing";
        const unitCls = p.unit ? "c" : "c missing";
        out.push(`<tr><td class="c">${p.mathml || esc(key)}</td><td class="${nameCls.trim()}">${esc(p.name || "정의 입력 필요")}</td>`
          + (showValues ? `<td class="c">${esc(p.value || "")}</td>` : "")
          + `<td class="${unitCls}">${esc(p.unit || "단위?")}</td></tr>`);
      }
      out.push("</tbody></table>");
    }
    out.push("</div>");
  }
  $("preview").innerHTML = out.join("");
}

function renderIssues(issues) {
  const main = issues.filter((i) => i.level !== "info");
  const infos = issues.filter((i) => i.level === "info");
  const errors = issues.filter((i) => i.level === "error").length;
  $("issues").innerHTML = main.length
    ? main.map((i) => `<li class="${i.level}"><b>${i.level === "error" ? "오류" : "확인"}</b>${esc(i.where)} — ${esc(i.msg)}</li>`).join("")
    : '<li class="info">오류·확인 항목이 없습니다.</li>';
  $("infos").innerHTML = infos.map((i) => `<li class="info">${esc(i.where)} — ${esc(i.msg)}</li>`).join("");
  $("info-wrap").style.display = infos.length ? "" : "none";
  $("issue-count").textContent = `오류 ${errors} · 확인 ${main.length - errors}`;
  for (const id of ["btn-docx", "btn-pptx"]) {
    $(id).disabled = errors > 0;
    $(id).title = errors ? "수식 오류를 먼저 고치세요" : "";
  }
}

let gridKeys = "";
function renderGrid(rows) {
  const keys = rows.map((r) => r.key).join("\u0001") + "|" + [...state.open].join(",");
  const grid = $("grid");
  if (keys === gridKeys && grid.contains(document.activeElement)) {
    // Keep focus while typing: refresh symbols, badges and highlights only.
    for (const r of rows) {
      const tr = grid.querySelector(`tr.row[data-key="${CSS.escape(r.key)}"]`);
      if (!tr) continue;
      tr.querySelector(".c-sym").innerHTML = r.mathml || esc(r.key);
      tr.classList.toggle("auto", !!r.auto);
      tr.classList.toggle("unused", !r.used);
      const badge = tr.querySelector(".badge");
      if (badge) badge.style.display = r.kind_auto ? "" : "none";
    }
    return;
  }
  gridKeys = keys;
  const html = [];
  for (const r of rows) {
    const kindOpts = Object.entries(KIND_LABEL)
      .map(([v, l]) => `<option value="${v}"${r.kind === v ? " selected" : ""}>${l}</option>`).join("");
    html.push(`<tr class="row${r.auto ? " auto" : ""}${r.used ? "" : " unused"}" data-key="${esc(r.key)}">
      <td class="c-sym">${r.mathml || esc(r.key)}</td>
      <td class="c-symtext"><input data-f="symbol" value="${esc(r.symbol)}" title="Enter 또는 칸을 벗어나면 반영"></td>
      <td><input data-f="name" value="${esc(r.name)}" placeholder="정의 (예: y년도 사업필터 순판매량)"></td>
      <td class="c-unit"><input data-f="unit" value="${esc(r.unit)}" placeholder="단위"></td>
      <td class="c-kind"><div class="kind-wrap"><select data-f="kind">${kindOpts}</select><span class="badge" style="${r.kind_auto ? "" : "display:none"}">자동</span></div></td>
      <td class="c-value"><input data-f="value" value="${esc(r.value)}" placeholder="값"></td>
      <td class="c-more"><button type="button" class="more-btn" data-more="1">${state.open.has(r.key) ? "접기" : "세부"}</button></td>
    </tr>`);
    if (state.open.has(r.key)) {
      const fields = DETAIL.map(([f, label]) => {
        const ph = f === "purpose" ? (r.purpose_auto || "") : "";
        return `<label>${label}</label><textarea data-f="${f}" rows="1" placeholder="${esc(ph)}">${esc(r[f] || "")}</textarea>`;
      }).join("");
      html.push(`<tr class="detail" data-key="${esc(r.key)}"><td colspan="7"><div class="detail-grid">${fields}</div></td></tr>`);
    }
  }
  grid.innerHTML = html.join("") || '<tr><td colspan="7" class="status">아직 인자가 없습니다.</td></tr>';
}

// Write one field of a derived row into plan.parameters (upsert by symbol).
function editParam(key, field, value) {
  const row = state.derived.parameters.find((r) => r.key === key);
  if (!row) return;
  let param = state.parameters.find((p) => p.symbol === row.symbol);
  if (!param) {
    param = { symbol: row.symbol };
    state.parameters.push(param);
  }
  if (value === "" && field !== "symbol") delete param[field];
  else param[field] = value;
  row[field] = value;
  if (field !== "kind") row.auto = false;
  const tr = $("grid").querySelector(`tr.row[data-key="${CSS.escape(key)}"]`);
  if (tr && field !== "kind") tr.classList.remove("auto");
  if (field === "symbol") gridKeys = "";      // key changes: rebuild after derive
  scheduleDerive(field === "symbol" || field === "kind" ? 0 : 400);
}

// ---------------------------------------------------------------- events
textEl.addEventListener("input", () => { state.text = textEl.value; scheduleDerive(); });
$("title").addEventListener("input", (e) => { state.title = e.target.value; save(); });
$("opt-number").addEventListener("change", (e) => { state.options.number_equations = e.target.checked; derive(); });
$("opt-values").addEventListener("change", (e) => { state.options.show_values = e.target.checked; derive(); });
$("opt-form").addEventListener("change", (e) => { state.options.form = e.target.value; derive(); });
$("opt-font").addEventListener("change", (e) => { state.font = e.target.value; save(); });

$("grid").addEventListener("input", (e) => {
  const f = e.target.dataset.f;
  if (!f || f === "symbol" || f === "kind") return;
  const key = e.target.closest("tr").dataset.key;
  editParam(key, f, e.target.value);
});
$("grid").addEventListener("change", (e) => {
  const f = e.target.dataset.f;
  if (f !== "symbol" && f !== "kind") return;
  const key = e.target.closest("tr").dataset.key;
  if (f === "symbol" && !e.target.value.trim()) { e.target.value = state.derived.parameters.find((r) => r.key === key)?.symbol || ""; return; }
  editParam(key, f, e.target.value.trim());
});
$("grid").addEventListener("keydown", (e) => {
  if (e.key === "Enter" && e.target.dataset.f === "symbol") e.target.blur();
});
$("grid").addEventListener("click", (e) => {
  if (!e.target.dataset.more) return;
  const key = e.target.closest("tr").dataset.key;
  if (state.open.has(key)) state.open.delete(key); else state.open.add(key);
  gridKeys = "";
  renderGrid(state.derived.parameters);
});

$("btn-example").addEventListener("click", async () => {
  if (state.text.trim() && !confirm("지금 입력한 내용을 예시로 바꿀까요? (JSON 저장으로 먼저 보관할 수 있습니다)")) return;
  const res = await fetch("/api/example");
  const data = await res.json();
  applyPlan(data.plan, data.text);
});

$("file-import").addEventListener("change", async (e) => {
  const file = e.target.files[0];
  e.target.value = "";
  if (!file) return;
  if (state.text.trim() && !confirm(`'${file.name}'의 수식과 인자 표로 지금 내용을 바꿀까요?`)) return;
  $("status").textContent = "파일 읽는 중…";
  try {
    const res = await fetch("/api/import", {
      method: "POST", headers: { "X-Filename": encodeURIComponent(file.name) }, body: await file.arrayBuffer(),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || res.statusText);
    applyPlan(data.plan, data.text);
    const n = data.plan.sections.reduce((a, s) => a + s.equations.length, 0);
    toast(`수식 ${n}개, 인자 ${data.plan.parameters.length}개를 가져왔습니다`);
  } catch (err) {
    $("status").textContent = "";
    toast("가져오기 실패: " + err.message, true);
  }
});

$("btn-save").addEventListener("click", async () => {
  // Save the derived plan (sections + parameters) so monitoring_plan.py can use it directly.
  const res = await fetch("/api/derive", {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(requestBody()),
  });
  const data = await res.json();
  const plan = { ...data.plan, title: state.title, font: state.font, options: state.options };
  const blob = new Blob([JSON.stringify(plan, null, 2)], { type: "application/json" });
  download(blob, (state.title || "수식_모니터링계획") + ".json");
});

$("file-json").addEventListener("change", async (e) => {
  const file = e.target.files[0];
  e.target.value = "";
  if (!file) return;
  try {
    const plan = JSON.parse(await file.text());
    if (!plan.sections && !plan.parameters) throw new Error("수식 계획 JSON이 아닙니다");
    applyPlan(plan, plan.text);
  } catch (err) {
    toast("JSON 열기 실패: " + err.message, true);
  }
});

async function exportFile(kind) {
  const btn = $(kind === "docx" ? "btn-docx" : "btn-pptx");
  btn.disabled = true;
  try {
    const res = await fetch(`/api/export/${kind}`, {
      method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(requestBody()),
    });
    if (!res.ok) {
      const data = await res.json().catch(() => ({}));
      throw new Error(data.error || res.statusText);
    }
    const base = (state.title || "수식_모니터링계획").replace(/[\\/:*?"<>|]/g, "_");
    download(await res.blob(), `${base}.${kind}`);
    toast(kind === "docx" ? "Word 파일을 만들었습니다" : "PPT 파일을 만들었습니다");
  } catch (err) {
    toast("내보내기 실패: " + err.message, true);
  } finally {
    renderIssues(state.derived ? state.derived.issues : []);
  }
}
function download(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}
$("btn-docx").addEventListener("click", () => exportFile("docx"));
$("btn-pptx").addEventListener("click", () => exportFile("pptx"));

// ---------------------------------------------------------------- start
(async function start() {
  try {
    const res = await fetch("/api/initial");
    const data = await res.json();
    if (data.plan) { applyPlan(data.plan, data.text); return; }
  } catch (_) { /* fall through */ }
  const saved = restore();
  if (saved) {
    Object.assign(state, { ...saved, options: { ...state.options, ...(saved.options || {}) } });
    syncControls();
    derive();
    return;
  }
  const res = await fetch("/api/example");
  const data = await res.json();
  applyPlan(data.plan, data.text);
})();

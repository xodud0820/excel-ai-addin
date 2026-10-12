const MAX_CELLS = 3000;
let pending = null, backup = null, backupShapes = null;

// ---------- 이미지(도형) ----------
const selectedShapes = new Set();

async function getShapeInfo(onlyChecked) {
  return Excel.run(async ctx => {
    const shs = ctx.workbook.worksheets.getActiveWorksheet().shapes;
    shs.load("items/name,items/type,items/left,items/top,items/width,items/height");
    await ctx.sync();
    let list = shs.items.map((s, i) => ({no: i + 1, name: s.name, type: s.type, left: +s.left.toFixed(1), top: +s.top.toFixed(1), width: +s.width.toFixed(1), height: +s.height.toFixed(1)}));
    if (onlyChecked) list = list.filter(s => selectedShapes.has(s.name));
    return list;
  });
}

async function getThumb(name) {
  try {
    return await Excel.run(async ctx => {
      const r = ctx.workbook.worksheets.getActiveWorksheet().shapes.getItem(name).getAsImage(Excel.PictureFormat.png);
      await ctx.sync();
      return "data:image/png;base64," + r.value;
    });
  } catch (_) { return null; }
}

function updateCount() {
  $("selCount").textContent = `선택된 이미지: ${selectedShapes.size}개`;
}

async function listShapes() {
  try {
    const all = await getShapeInfo(false);
    const withOthers = $("incOthers").checked;
    const list = all.filter(s => withOthers || s.type === "Image");
    const names = new Set(all.map(s => s.name));
    [...selectedShapes].forEach(n => { if (!names.has(n)) selectedShapes.delete(n); });
    const box = $("shapes");
    box.innerHTML = list.length ? "" : "<i>이 시트에 이미지가 없습니다</i>";
    const tiles = [];
    list.slice(0, 40).forEach(s => {
      const d = document.createElement("div");
      d.className = "tile" + (selectedShapes.has(s.name) ? " sel" : "");
      d.dataset.name = s.name;
      d.title = `${s.name} (${s.type}, ${s.width}×${s.height}pt)`;
      d.innerHTML = `<span class="no">${s.no}</span><div class="th">...</div><small>${s.width}×${s.height}</small>`;
      d.onclick = () => {
        if (selectedShapes.has(s.name)) selectedShapes.delete(s.name); else selectedShapes.add(s.name);
        d.classList.toggle("sel"); updateCount();
      };
      box.appendChild(d); tiles.push([s, d]);
    });
    updateCount();
    for (const [s, d] of tiles) {
      const t = await getThumb(s.name);
      d.querySelector(".th").innerHTML = t ? `<img src="${t}"/>` : `<span style="font-size:11px">${s.type}</span>`;
    }
  } catch (e) { log("이미지 목록 오류: " + e.message); }
}

function selectAllShapes(on) {
  document.querySelectorAll(".tile").forEach(d => {
    if (on) selectedShapes.add(d.dataset.name); else selectedShapes.delete(d.dataset.name);
    d.classList.toggle("sel", on);
  });
  updateCount();
}

// 셀 범위를 드래그해 선택하면, 그 범위 안에 있는 이미지를 자동으로 체크
async function selectShapesInRange(silent) {
  try {
    await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      const r = ctx.workbook.getSelectedRange(); r.load("left,top,width,height,address");
      ws.shapes.load("items/name,items/left,items/top,items/width,items/height");
      await ctx.sync();
      const hit = ws.shapes.items.filter(s => {
        const cx = s.left + s.width / 2, cy = s.top + s.height / 2;
        return cx >= r.left && cx <= r.left + r.width && cy >= r.top && cy <= r.top + r.height;
      }).map(s => s.name);
      selectedShapes.clear(); hit.forEach(n => selectedShapes.add(n));
      document.querySelectorAll(".tile").forEach(d => d.classList.toggle("sel", selectedShapes.has(d.dataset.name)));
      updateCount();
      if (!silent) log(`${r.address} 범위 안의 이미지 ${hit.length}개를 선택했습니다.`);
    });
  } catch (e) {
    if (!silent) log("먼저 시트에서 셀 범위를 드래그해 선택한 뒤 눌러주세요. (이미지를 클릭한 상태면 셀 범위가 없습니다)");
  }
}

async function applyShapes(ctx, ops, saved) {
  if (!ops.length) return;
  const ws = ctx.workbook.worksheets.getActiveWorksheet();
  const shs = ws.shapes;
  shs.load("items/name,items/left,items/top,items/width,items/height,items/lockAspectRatio");
  await ctx.sync();
  const byName = Object.fromEntries(shs.items.map(s => [s.name, s]));
  const pre = [];
  for (const op of ops) {
    const s = byName[op.name];
    if (!s) throw new Error(`이미지를 찾을 수 없음: ${op.name}`);
    const need = {};
    if (op.anchorCell) { need.a = ws.getRange(op.anchorCell); need.a.load("left,top"); }
    if (op.fitToRange) { need.f = ws.getRange(op.fitToRange); need.f.load("left,top,width,height"); }
    pre.push({s, op, need});
  }
  await ctx.sync();
  for (const {s, op, need} of pre) {
    saved.push({name:s.name, left:s.left, top:s.top, width:s.width, height:s.height});
    const ow = s.width, oh = s.height;
    let lock = op.lockAspectRatio;
    s.lockAspectRatio = false;
    if (op.scalePercent) { s.width = ow * op.scalePercent/100; s.height = oh * op.scalePercent/100; }
    if (op.width != null && op.height != null) { s.width = op.width; s.height = op.height; }
    else if (op.width != null) { s.width = op.width; if (lock !== false) s.height = oh * op.width/ow; }
    else if (op.height != null) { s.height = op.height; if (lock !== false) s.width = ow * op.height/oh; }
    if (need.f) {
      const f = need.f;
      if (lock) { const k = Math.min(f.width/ow, f.height/oh); s.width = ow*k; s.height = oh*k;
        s.left = f.left + (f.width-ow*k)/2; s.top = f.top + (f.height-oh*k)/2; }
      else { s.width = f.width; s.height = f.height; s.left = f.left; s.top = f.top; }
    }
    if (need.a) { s.left = need.a.left; s.top = need.a.top; }
    if (op.left != null) s.left = op.left;
    if (op.top != null) s.top = op.top;
    if (op.placement) s.placement = op.placement;
    if (op.lockAspectRatio != null) s.lockAspectRatio = !!op.lockAspectRatio;
    if (op.z === "front") s.setZOrder("BringToFront");
    else if (op.z === "back") s.setZOrder("SendToBack");
    else if (op.z === "forward") s.setZOrder("BringForward");
    else if (op.z === "backward") s.setZOrder("SendBackward");
  }
  await ctx.sync();
}
const $ = id => document.getElementById(id);
const log = m => { $("log").textContent = m; };

Office.onReady(async () => {
  loadCfg();
  listShapes();
  $("scope").onchange = () => $("customRow").style.display = $("scope").value === "custom" ? "block" : "none";
  await Excel.run(async ctx => {
    ctx.workbook.onSelectionChanged.add(async () => refreshSel());
    await ctx.sync();
  });
  refreshSel();
});

function loadCfg() {
  ["provider","endpoint","model","key"].forEach(k => { const v = localStorage.getItem("ai_"+k); if (v) $(k).value = v; });
}
function saveCfg() {
  ["provider","endpoint","model","key"].forEach(k => localStorage.setItem("ai_"+k, $(k).value.trim()));
  log("설정 저장됨 (키는 이 브라우저에만 저장됩니다)");
}
async function refreshSel() {
  try {
    await Excel.run(async ctx => {
      const r = ctx.workbook.getSelectedRange(); r.load("address"); await ctx.sync();
      $("sel").textContent = "선택: " + r.address;
    });
    if ($("autoSel") && $("autoSel").checked) await selectShapesInRange(true);
  } catch (e) {
    $("sel").textContent = "선택: (이미지/도형이 선택됨 - 셀 범위 없음)";
  }
}

async function getTarget(ctx) {
  const ws = ctx.workbook.worksheets.getActiveWorksheet();
  const s = $("scope").value;
  if (s === "selection") return ctx.workbook.getSelectedRange();
  if (s === "used") return ws.getUsedRange();
  const a = $("custom").value.trim();
  if (!a) throw new Error("범위를 입력하세요");
  return ws.getRange(a);
}

async function readContext() {
  if ($("scope").value === "none") return emptyCtx();
  try {
    return await readContextCells();
  } catch (e) {
    // 이미지/도형이 선택된 상태에서는 셀 선택 영역이 없어 오류 -> 이미지가 체크돼 있으면 셀 없이 진행
    const imgs = await getShapeInfo(true);
    if ($("scope").value === "selection" && imgs.length) return emptyCtx();
    throw e;
  }
}

async function emptyCtx() {
  return Excel.run(async ctx => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet(); ws.load("name"); await ctx.sync();
    return { sheet: ws.name, address: "(셀 없음, 이미지 작업만)", formulas: [], header: null };
  });
}

async function readContextCells() {
  return Excel.run(async ctx => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet(); ws.load("name");
    let r = await getTarget(ctx);
    // 전체 열/행 선택 시 사용 범위와 교차
    const used = ws.getUsedRangeOrNullObject(true);
    r.load("address,rowCount,columnCount,rowIndex,columnIndex"); used.load("address"); await ctx.sync();
    let target = r;
    if (!used.isNullObject && (r.rowCount > 5000 || r.columnCount > 200)) {
      target = r.getIntersection(used);
    }
    target.load("address,formulas,rowCount,columnCount"); await ctx.sync();
    if (target.rowCount * target.columnCount > MAX_CELLS)
      throw new Error(`범위가 너무 큽니다 (${target.rowCount*target.columnCount}셀 > ${MAX_CELLS}). 범위를 줄여주세요.`);
    // 헤더(사용 범위 1행) 참고용
    let header = null;
    if (!used.isNullObject) {
      const h = used.getRow(0); h.load("values,address"); await ctx.sync(); header = {address:h.address, values:h.values[0]};
    }
    return { sheet: ws.name, address: target.address, formulas: target.formulas, header };
  });
}

const SYSTEM = `당신은 엑셀 편집 어시스턴트입니다. 사용자는 시트 범위와 자연어 지시를 줍니다.
반드시 아래 JSON만 출력하세요(코드펜스/설명 금지):
{"explanation":"변경 요약(한국어)","writes":[{"address":"시트이름!B2:B10","values":[[...],[...]]}],"shapes":[{"name":"Picture 1","width":200,"height":150,"left":10,"top":20,"anchorCell":"B2","fitToRange":"B2:D6","scalePercent":50,"lockAspectRatio":true,"placement":"TwoCell|OneCell|Absolute","z":"front|back|forward|backward"}]}
이미지 규칙(shapes는 이미지 작업 지시가 있을 때만, 없으면 생략/빈 배열):
- name은 제공된 '이미지 목록'에 있는 이름만 사용. 단위는 포인트(pt). 사용자가 cm/mm/inch로 말하면 pt로 환산(1cm=28.346pt, 1mm=2.8346pt, 1inch=72pt). 예: 19cm=538.6pt.
- 사용자가 '1번 이미지'라고 하면 목록의 no가 1인 항목의 name을 쓴다. '선택한 이미지 전부'는 목록의 모든 항목을 대상으로 한다.
- 너비와 높이를 둘 다 지정하는 지시(예: '둘 다 19cm')는 lockAspectRatio를 false로 하고 width, height를 모두 넣는다.
- 필요한 속성만 포함. anchorCell: 이미지 좌상단을 해당 셀 좌상단에 맞춤. fitToRange: 범위 크기에 맞춰 위치+크기 조정(lockAspectRatio가 true면 비율 유지하며 범위 안에 맞춤).
- scalePercent: 현재 크기 기준 배율. 여러 이미지를 '같은 크기/가로 정렬/세로로 나열/격자 배치' 하라는 지시는 width/height/left/top을 계산해 각각 지정.
- placement: TwoCell=셀과 함께 이동·크기변경, OneCell=이동만, Absolute=고정.
규칙:
- values는 address 크기와 정확히 같은 2차원 배열. 수식은 "=SUM(A1:A3)" 처럼 문자열로.
- 변경 없는 셀은 원래 값을 그대로 넣거나 writes에서 제외.
- 주어진 범위 밖(예: 옆 빈 열)에 쓰는 것은 지시에 명시된 경우에만 허용.
- 숫자는 숫자 타입으로, 빈 셀은 "".`;

async function callLLM(userMsg) {
  const provider = $("provider").value, key = $("key").value.trim();
  const endpoint = $("endpoint").value.trim(), model = $("model").value.trim();
  if (!key) throw new Error("설정에서 API Key를 입력하세요");
  let res, data, text;
  if (provider === "gemini") {
    const m = model || "gemini-2.0-flash";
    res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${m}:generateContent?key=${encodeURIComponent(key)}`, {
      method: "POST", headers: {"Content-Type":"application/json"},
      body: JSON.stringify({
        systemInstruction: {parts:[{text:SYSTEM}]},
        contents: [{role:"user", parts:[{text:userMsg}]}],
        generationConfig: {responseMimeType:"application/json", temperature:0.2}
      })});
    data = await res.json();
    if (!res.ok) throw new Error(data.error?.message || res.statusText);
    text = data.candidates?.[0]?.content?.parts?.map(p=>p.text).join("") || "";
  } else {
    if (!endpoint) throw new Error("Timely Endpoint를 입력하세요");
    res = await fetch(endpoint, {
      method: "POST", headers: {"Content-Type":"application/json","Authorization":"Bearer "+key},
      body: JSON.stringify({ model, temperature: 0.2,
        messages:[{role:"system",content:SYSTEM},{role:"user",content:userMsg}] })});
    data = await res.json();
    if (!res.ok) throw new Error(data.error?.message || JSON.stringify(data));
    text = data.choices?.[0]?.message?.content || "";
  }
  text = text.replace(/^```(?:json)?/i,"").replace(/```$/,"").trim();
  return JSON.parse(text);
}

async function run() {
  const p = $("prompt").value.trim();
  if (!p) return log("지시사항을 입력하세요");
  $("run").disabled = true; pending = null; $("apply").disabled = true;
  try {
    if ($("scope").value === "none" && selectedShapes.size === 0) throw new Error("이미지를 선택하세요 (썸네일 클릭)");
    log("범위 읽는 중...");
    const c = await readContext();
    log(`${c.address} 처리 중 (AI 호출)...`);
    const imgs = await getShapeInfo(true);
    const imgCtx = imgs.length ? `\n선택된 이미지 목록(no=시트의 이미지 번호, 이 이미지들이 작업 대상, 단위 pt):\n${JSON.stringify(imgs)}\n` : "";
    const msg = `시트: ${c.sheet}\n대상 범위: ${c.address}\n시트 1행(헤더): ${JSON.stringify(c.header?.values)}\n대상 데이터(수식 포함, 2차원):\n${JSON.stringify(c.formulas)}\n${imgCtx}\n지시: ${p}`;
    const out = await callLLM(msg);
    out.writes = Array.isArray(out.writes) ? out.writes : [];
    out.shapes = Array.isArray(out.shapes) ? out.shapes : [];
    pending = out;
    const lines = out.writes.map(w => `• ${w.address} (${w.values.length}행×${w.values[0]?.length||0}열)\n   예: ${JSON.stringify(w.values.slice(0,3))}`);
    const sl = out.shapes.map(s => `🖼 ${s.name}: ${JSON.stringify(s)}`);
    log(`${out.explanation}\n\n${lines.concat(sl).join("\n")}\n\n→ [적용]을 누르면 반영됩니다.`);
    $("apply").disabled = false;
  } catch (e) { log("오류: " + e.message); }
  $("run").disabled = false;
}

function getRng(ctx, full) {
  const [sh, addr] = full.includes("!") ? full.split("!") : [null, full];
  const ws = sh ? ctx.workbook.worksheets.getItem(sh.replace(/^'|'$/g,"")) : ctx.workbook.worksheets.getActiveWorksheet();
  return ws.getRange(addr);
}

async function apply() {
  if (!pending) return;
  try {
    const saved = [], savedShapes = [];
    await Excel.run(async ctx => {
      const rs = pending.writes.map(w => { const r = getRng(ctx, w.address); r.load("formulas"); return {r, w}; });
      await ctx.sync();
      rs.forEach(({r, w}) => {
        if (r.formulas.length !== w.values.length || r.formulas[0].length !== w.values[0].length)
          throw new Error(`크기 불일치: ${w.address}`);
        saved.push({address: w.address, old: r.formulas});
      });
      rs.forEach(({r, w}) => r.formulas = w.values);
      await ctx.sync();
      await applyShapes(ctx, pending.shapes || [], savedShapes);
    });
    backup = saved; backupShapes = savedShapes;
    log("적용 완료. 필요하면 [되돌리기]를 누르세요.");
    $("undo").disabled = false; $("apply").disabled = true; pending = null;
  } catch (e) { log("적용 오류: " + e.message); }
}

async function undo() {
  if (!backup) return;
  try {
    await Excel.run(async ctx => {
      backup.forEach(b => { getRng(ctx, b.address).formulas = b.old; });
      const shs = ctx.workbook.worksheets.getActiveWorksheet().shapes;
      (backupShapes || []).forEach(o => {
        const s = shs.getItemOrNullObject(o.name);
        s.lockAspectRatio = false; s.left = o.left; s.top = o.top; s.width = o.width; s.height = o.height;
        if (o.placement) s.placement = o.placement;
      });
      await ctx.sync();
    });
    log("원래 값/이미지 상태로 복원했습니다."); backup = null; backupShapes = null; $("undo").disabled = true;
  } catch (e) { log("복원 오류: " + e.message); }
}

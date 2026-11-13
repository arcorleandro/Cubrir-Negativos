const KEEP = ["Tipo Alm.","Nro. Item","Descripción","UM","Ubicación","Lote","Estado Lote","Stock","Fecha Alta Lote"];
const norm = s => String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g,"").trim().toLowerCase();
const $ = id => document.getElementById(id);

let rows = [], header = [], idx = {}, fileName = "stock";
let selTipo, selUbic, selProd, prodDesc = new Map();

/* ---------- Lectura del CSV ---------- */
function decode(buf){
  let t = new TextDecoder("utf-8").decode(buf);
  if (t.includes("\uFFFD")) t = new TextDecoder("windows-1252").decode(buf);
  return t.replace(/^\uFEFF/, "");
}
function detectDelimiter(text){
  const first = text.split(/\r?\n/)[0] || "";
  const c = d => first.split(d).length;
  return [";","\t",","].sort((a,b)=>c(b)-c(a))[0];
}
function parseCsv(text){
  const wb = XLSX.read(text, {type:"string", raw:true, FS:detectDelimiter(text)});
  const ws = wb.Sheets[wb.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:""});
  return data.filter(r => r.some(v => String(v).trim() !== ""));
}

async function loadFile(f){
  $("error").hidden = true;
  if (!f) return;
  try{
    const data = parseCsv(decode(await f.arrayBuffer()));
    if (data.length < 2) throw new Error("El archivo no tiene datos.");
    header = data[0].map(h => String(h).trim());
    rows = data.slice(1);

    idx = {};
    const missing = [];
    KEEP.concat(["Ubicación"]).forEach(name => {
      const i = header.findIndex(h => norm(h) === norm(name));
      if (i === -1) { if (!missing.includes(name)) missing.push(name); } else idx[name] = i;
    });
    if (missing.length) throw new Error("Faltan estas columnas en el archivo: " + missing.join(", ") + ".");

    fileName = f.name.replace(/\.csv$/i, "");
    $("fileInfo").hidden = false;
    $("fileInfo").textContent = `${f.name}: ${rows.length.toLocaleString("es")} filas cargadas.`;
    buildFilters();
  }catch(e){
    $("filters").hidden = $("out").hidden = $("fileInfo").hidden = true;
    $("error").hidden = false;
    $("error").textContent = e.message || "No se pudo leer el archivo.";
  }
}

/* ---------- Lista desplegable con selección múltiple ---------- */
function multiSelect(root, label, values, onChange, labelOf = v => v === "" ? "(vacío)" : v, ph = "Buscar…", emptyTitle = "Seleccionar…"){
  const selected = new Set();
  root.innerHTML = `<button type="button" class="ms-trigger" aria-expanded="false"><span class="t"></span><span aria-hidden="true">▾</span></button>
    <div class="ms-panel" hidden>
      <input type="search" class="field" placeholder="${ph}" aria-label="Buscar en ${label}">
      <div class="ms-bar"><button type="button" class="btn-sm" data-a="all">Todas</button><button type="button" class="btn-sm" data-a="none">Ninguna</button></div>
      <div class="ms-list"></div>
    </div>`;
  const btn = root.querySelector("button"), panel = root.querySelector(".ms-panel"),
        search = root.querySelector("input"), list = root.querySelector(".ms-list"),
        title = root.querySelector(".t");
  const refreshTitle = () => {
    title.textContent = selected.size === 0 ? emptyTitle :
      selected.size === values.length ? `Todas (${values.length})` : `${selected.size} de ${values.length}`;
    onChange();
  };
  const render = () => {
    const q = norm(search.value);
    list.innerHTML = "";
    values.filter(v => norm(labelOf(v)).includes(q)).forEach(v => {
      const l = document.createElement("label");
      const c = document.createElement("input");
      c.type = "checkbox"; c.checked = selected.has(v);
      c.onchange = () => { c.checked ? selected.add(v) : selected.delete(v); refreshTitle(); };
      l.append(c, document.createTextNode(labelOf(v)));
      list.append(l);
    });
    if (!list.children.length) list.innerHTML = '<p class="hint" style="margin:6px">Sin resultados.</p>';
  };
  btn.onclick = () => { panel.hidden = !panel.hidden; btn.setAttribute("aria-expanded", !panel.hidden); if(!panel.hidden) search.focus(); };
  search.oninput = render;
  panel.querySelectorAll("[data-a]").forEach(b => b.onclick = () => {
    const q = norm(search.value);
    values.filter(v => norm(labelOf(v)).includes(q)).forEach(v => b.dataset.a === "all" ? selected.add(v) : selected.delete(v));
    render(); refreshTitle();
  });
  document.addEventListener("click", e => { if (!root.contains(e.target)) { panel.hidden = true; btn.setAttribute("aria-expanded","false"); } });
  root.addEventListener("keydown", e => { if (e.key === "Escape") { panel.hidden = true; btn.focus(); } });
  selected.setValues = nv => {
    values = nv;
    [...selected].forEach(v => { if (!nv.includes(v)) selected.delete(v); });
    search.value = ""; render(); refreshTitle();
  };
  render(); refreshTitle();
  return selected;
}

const unique = col => [...new Set(rows.map(r => String(r[idx[col]]).trim()))].sort((a,b)=>a.localeCompare(b,"es",{numeric:true}));

function buildFilters(){
  $("filters").hidden = $("out").hidden = false;
  $("done").hidden = true;
  $("dest").onchange = update;
  selTipo = selUbic = selProd = null;
  selTipo = multiSelect($("msTipo"), "Tipo Alm.", unique("Tipo Alm."), () => selUbic ? selUbic.setValues(ubicsFor()) : update());
  selUbic = multiSelect($("msUbic"), "Ubicación", ubicsFor(), () => selProd ? selProd.setValues(prodsFor()) : update());
  selProd = multiSelect($("msProd"), "Producto", prodsFor(), update, k => prodLabel(k), "Número o descripción…", "Todos los productos");
  update();
}

// Ubicaciones que existen para los tipos de almacenamiento elegidos
const ubicsFor = () => [...new Set(rows.filter(r => selTipo.has(String(r[idx["Tipo Alm."]]).trim()))
  .map(r => String(r[idx["Ubicación"]]).trim()))].sort((a,b) => a.localeCompare(b,"es",{numeric:true}));

// Un producto = un Nro. Item (cada ítem tiene una sola descripción); se busca por número o descripción
const byTipoUbic = () => rows.filter(r => selTipo.has(String(r[idx["Tipo Alm."]]).trim()) && selUbic.has(String(r[idx["Ubicación"]]).trim()));
const prodsFor = () => {
  const m = new Map();
  byTipoUbic().forEach(r => {
    const k = cleanItem(r[idx["Nro. Item"]]);
    if (k !== "" && !m.get(k)) m.set(k, String(r[idx["Descripción"]]).trim());
  });
  prodDesc = m;
  return [...m.keys()].sort((a,b) => a.localeCompare(b,"es",{numeric:true}));
};
const prodLabel = k => prodDesc.get(k) ? `${k} – ${prodDesc.get(k)}` : k;

const isLP = u => /^(env|ela)/i.test(u);

// La lista de destino depende de los filtros; las ENV se agrupan como "Linea Productiva"
function refreshDest(){
  const d = $("dest"), prev = d.value;
  const locs = [...new Set(byTipoUbic().map(r => String(r[idx["Ubicación"]]).trim()))].filter(Boolean)
    .sort((a,b) => a.localeCompare(b,"es",{numeric:true}));
  d.innerHTML = '<option value="" disabled selected>Seleccionar…</option><option value="__none__">No mover los sobrantes</option>';
  if (locs.some(isLP)) d.add(new Option("Linea Productiva", "__LP__"));
  locs.filter(u => !isLP(u)).forEach(u => d.add(new Option(u, u)));
  if (prev && [...d.options].some(o => o.value === prev)) d.value = prev;
}

function filtered(){
  return rows.filter(r => selTipo.has(String(r[idx["Tipo Alm."]]).trim()) && selUbic.has(String(r[idx["Ubicación"]]).trim())
    && (selProd.size === 0 || selProd.has(cleanItem(r[idx["Nro. Item"]]))));
}
function update(){
  if (!selTipo || !selUbic || !selProd) return;
  refreshDest();
  const ready = selTipo.size > 0 && selUbic.size > 0 && $("dest").value !== "";
  $("gen").disabled = !ready;
  $("count").textContent = ready ? `${filtered().length.toLocaleString("es")} filas coinciden con tu selección.` : "";
}

/* ---------- Generar Excel ---------- */
function toNumber(v){
  if (typeof v === "number") return v;
  const s = String(v).trim();
  if (s === "") return "";
  const n = Number(s.includes(",") ? s.replace(/\./g,"").replace(",",".") : s);
  return Number.isFinite(n) ? n : v;
}

const cleanItem = v => String(v ?? "").replace(/[+"'`´’‘“”\s\u00a0]/g, "");
const itemOut = s => /^[1-9]\d*$/.test(s) ? Number(s) : s;
const r6 = n => Math.round(n * 1e6) / 1e6;
const IT_HEAD = ["Nro. Item","Descripción","Stock","UM","Ubicación","Lote","Ubicacion Final","Lote destino"];
const NC_HEAD = IT_HEAD.slice(0, 6);

// Cubre cada negativo con positivos del mismo ítem; el sobrante va a "dest" (si hay)
function coverNegatives(dest){
  const LP = dest === "__LP__";
  const envList = [...new Set(filtered().map(r => String(r[idx["Ubicación"]]).trim()).filter(isLP))].sort();
  const groups = new Map();
  filtered().forEach(r => {
    const st = toNumber(r[idx["Stock"]]);
    if (typeof st !== "number" || st === 0) return;
    const item = cleanItem(r[idx["Nro. Item"]]);
    if (!groups.has(item)) groups.set(item, []);
    groups.get(item).push({item, desc:r[idx["Descripción"]], um:r[idx["UM"]], ubic:String(r[idx["Ubicación"]]).trim(), lote:r[idx["Lote"]], st, fin:st});
  });
  const out = [], left = [], fin = []; let uncovered = 0, before = 0, removed = 0, created = 0;
  groups.forEach(lines => {
    // Linea Productiva: una sola ENV por ítem (la que ya tiene más positivo; si no hay, otra ENV del ítem o la primera ENV filtrada)
    let target = dest;
    if (LP) {
      const sum = {};
      lines.filter(l => l.st > 0 && isLP(l.ubic)).forEach(l => sum[l.ubic] = (sum[l.ubic] || 0) + l.st);
      const best = Object.keys(sum).sort((a,b) => sum[b] - sum[a])[0];
      target = best || (lines.find(l => isLP(l.ubic)) || {}).ubic || envList[0] || "";
    }
    // primero se usan los positivos que NO están en el destino, y los de mayor stock
    const pos = lines.filter(l => l.st > 0).sort((a,b) => ((a.ubic === target) - (b.ubic === target)) || b.st - a.st);
    const push = (p, qty, to, ld = p.lote) => out.push([itemOut(p.item), p.desc, qty, p.um, p.ubic, p.lote, to, ld]);
    lines.filter(l => l.st < 0).forEach(n => {
      let need = r6(-n.st);
      for (const p of pos) {
        if (need <= 0) break;
        if (p.st <= 0 || (p.ubic === n.ubic && String(p.lote).trim() === String(n.lote).trim())) continue;
        const q = Math.min(p.st, need);
        push(p, q, n.ubic, n.lote);
        p.st = r6(p.st - q); need = r6(need - q);
        p.fin = r6(p.fin - q); n.fin = r6(n.fin + q);
      }
      if (need > 0) { uncovered++; left.push([itemOut(n.item), n.desc, -need, n.um, n.ubic, n.lote]); }
    });
    if (target) pos.forEach(p => {
      if (p.st > 0 && p.ubic !== target) {
        push(p, p.st, target);
        p.fin = r6(p.fin - p.st);
        const d = lines.find(l => l.ubic === target && l.lote === p.lote);
        if (d) d.fin = r6(d.fin + p.st); else lines.push({...p, ubic: target, st: 0, fin: p.st, isNew: true});
      }
    });
    lines.forEach(l => {
      if (l.isNew) created++; else { before++; if (r6(l.fin) === 0) removed++; }
      if (r6(l.fin) !== 0) fin.push([itemOut(l.item), l.desc, l.fin, l.um, l.ubic, l.lote]);
    });
  });
  return {out, left, fin, uncovered, before, removed, created};
}

$("gen").onclick = () => {
  const wb = XLSX.utils.book_new();

  // 2. Stock Filtros: solo las columnas pedidas y la selección del usuario
  const cols = KEEP.map(k => idx[k]);
  const stockPos = KEEP.indexOf("Stock"), itemPos = KEEP.indexOf("Nro. Item");
  const body = filtered().map(r => cols.map((c,i) => i === stockPos ? toNumber(r[c]) : i === itemPos ? itemOut(cleanItem(r[c])) : r[c]));
  const wsF = XLSX.utils.aoa_to_sheet([KEEP, ...body]);
  wsF["!cols"] = KEEP.map((k,i) => ({wch: Math.min(40, Math.max(k.length, ...body.slice(0,500).map(r => String(r[i]).length)) + 2)}));
  XLSX.utils.book_append_sheet(wb, wsF, "Stock Filtros");

  // 3. IT Negativos: transferencias para cubrir negativos
  let dv = $("dest").value;
  dv = dv === "__none__" ? "" : dv;
  const res = coverNegatives(dv);
  const wsI = XLSX.utils.aoa_to_sheet([IT_HEAD, ...res.out]);
  wsI["!cols"] = IT_HEAD.map((k,i) => ({wch: Math.min(40, Math.max(k.length, ...res.out.slice(0,500).map(r => String(r[i]).length)) + 2)}));
  XLSX.utils.book_append_sheet(wb, wsI, "IT Negativos");

  // 4. Negativos sin cubrir: lo que quedó negativo después de las transferencias
  const wsN = XLSX.utils.aoa_to_sheet([NC_HEAD, ...res.left]);
  wsN["!cols"] = NC_HEAD.map((k,i) => ({wch: Math.min(40, Math.max(k.length, ...res.left.slice(0,500).map(r => String(r[i]).length)) + 2)}));
  XLSX.utils.book_append_sheet(wb, wsN, "Negativos sin cubrir");

  // 5. Stock final: cómo queda el stock después de las transferencias
  const wsS = XLSX.utils.aoa_to_sheet([NC_HEAD, ...res.fin]);
  wsS["!cols"] = NC_HEAD.map((k,i) => ({wch: Math.min(40, Math.max(k.length, ...res.fin.slice(0,500).map(r => String(r[i]).length)) + 2)}));
  XLSX.utils.book_append_sheet(wb, wsS, "Stock final");

  const d = new Date(), pad = n => String(n).padStart(2, "0");
  const outName = `Stock procesado ${pad(d.getDate())}-${pad(d.getMonth() + 1)}-${d.getFullYear()}.xlsx`;
  XLSX.writeFile(wb, outName);
  $("done").hidden = false;
  const nf = n => n.toLocaleString("es");
  $("done").textContent =
    `Se eliminaron ${nf(res.removed)} filas del stock por cubrir negativos (de ${nf(res.before)} a ${nf(res.fin.length)} filas).` +
    `\nIT Negativos tiene ${nf(res.out.length)} filas. ` +
    (res.uncovered ? `${nf(res.uncovered)} negativos no se pudieron cubrir del todo; están en la hoja Negativos sin cubrir.` : "Todos los negativos quedaron cubiertos.") +
    `\nGuardado como ${outName}.`;
};

/* ---------- Eventos de carga ---------- */
$("file").onchange = e => loadFile(e.target.files[0]);
const drop = $("drop");
["dragenter","dragover"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add("over"); }));
["dragleave","drop"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove("over"); }));
drop.addEventListener("drop", e => loadFile(e.dataTransfer.files[0]));

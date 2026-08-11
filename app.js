/* ============================================================
   GST Ledger — client-side app. All data lives in localStorage
   and is exportable/importable as an Excel workbook for backup.
   ============================================================ */

/* ---------- GST state codes (first 2 digits of every GSTIN) ---------- */
const STATE_CODES = [
  ["01","Jammu & Kashmir"],["02","Himachal Pradesh"],["03","Punjab"],["04","Chandigarh"],
  ["05","Uttarakhand"],["06","Haryana"],["07","Delhi"],["08","Rajasthan"],["09","Uttar Pradesh"],
  ["10","Bihar"],["11","Sikkim"],["12","Arunachal Pradesh"],["13","Nagaland"],["14","Manipur"],
  ["15","Mizoram"],["16","Tripura"],["17","Meghalaya"],["18","Assam"],["19","West Bengal"],
  ["20","Jharkhand"],["21","Odisha"],["22","Chhattisgarh"],["23","Madhya Pradesh"],["24","Gujarat"],
  ["25","Daman & Diu (legacy)"],["26","Dadra & Nagar Haveli and Daman & Diu"],["27","Maharashtra"],
  ["28","Andhra Pradesh (legacy)"],["29","Karnataka"],["30","Goa"],["31","Lakshadweep"],["32","Kerala"],
  ["33","Tamil Nadu"],["34","Puducherry"],["35","Andaman & Nicobar Islands"],["36","Telangana"],
  ["37","Andhra Pradesh"],["38","Ladakh"]
];
const STATE_NAME = Object.fromEntries(STATE_CODES);
const GST_RATES = [0, 0.25, 3, 5, 12, 18, 28];

/* ---------- Storage ---------- */
const STORAGE_KEY = "gstLedgerData_v1";

function defaultData(){
  return {
    settings: { bizName: "", bizGSTIN: "", bizState: "23" },
    vendors: [], clients: [], purchases: [], sales: []
  };
}
function loadData(){
  try{
    const raw = localStorage.getItem(STORAGE_KEY);
    if(!raw) return defaultData();
    const parsed = JSON.parse(raw);
    return Object.assign(defaultData(), parsed, {
      settings: Object.assign(defaultData().settings, parsed.settings || {})
    });
  }catch(e){
    console.error("Failed to load saved data, starting fresh.", e);
    return defaultData();
  }
}
function saveData(){
  localStorage.setItem(STORAGE_KEY, JSON.stringify(DATA));
}

let DATA = loadData();

/* ---------- Helpers ---------- */
function uid(){ return Date.now().toString(36) + Math.random().toString(36).slice(2,7); }
function round2(n){ return Math.round((n + Number.EPSILON) * 100) / 100; }
function fmt(n){ return (n||0).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
function todayISO(){ return new Date().toISOString().slice(0,10); }

function isValidGSTIN(gstin){
  return /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$/.test((gstin||"").toUpperCase());
}
function stateCodeFromGSTIN(gstin){
  if(!gstin || gstin.length < 2) return null;
  const code = gstin.slice(0,2);
  return STATE_NAME[code] ? code : null;
}
function computeGst(taxable, ratePct, inState){
  const taxAmt = round2(taxable * (ratePct/100));
  if(inState){
    const half = round2(taxAmt/2);
    return { cgst: half, sgst: round2(taxAmt-half), igst: 0 };
  }
  return { cgst: 0, sgst: 0, igst: taxAmt };
}
function toast(msg){
  const t = document.getElementById("toast");
  t.textContent = msg; t.hidden = false;
  clearTimeout(toast._h);
  toast._h = setTimeout(()=>{ t.hidden = true; }, 2600);
}
function populateStateSelect(sel){
  sel.innerHTML = STATE_CODES.map(([c,n])=>`<option value="${c}">${c} — ${n}</option>`).join("");
}

/* ============================================================
   Settings
   ============================================================ */
const settingsToggle = document.getElementById("settingsToggle");
const settingsPanel = document.getElementById("settingsPanel");
const bizStateSel = document.getElementById("bizState");
populateStateSelect(bizStateSel);

function loadSettingsIntoForm(){
  document.getElementById("bizName").value = DATA.settings.bizName;
  document.getElementById("bizGSTIN").value = DATA.settings.bizGSTIN;
  bizStateSel.value = DATA.settings.bizState || "23";
  updateSettingsSummary();
}
function updateSettingsSummary(){
  const name = DATA.settings.bizName || "your business";
  const stateName = STATE_NAME[DATA.settings.bizState] || "—";
  document.getElementById("settingsBtnLabel").textContent =
    DATA.settings.bizName ? `${DATA.settings.bizName} · ${stateName}` : "Business settings";
  document.getElementById("bizStateAuto").textContent =
    `Entries are marked in-state when the party's GSTIN starts with ${DATA.settings.bizState} (${stateName}).`;
}
settingsToggle.addEventListener("click", ()=>{
  const open = settingsPanel.hidden;
  settingsPanel.hidden = !open;
  settingsToggle.setAttribute("aria-expanded", String(open));
});
document.getElementById("saveSettings").addEventListener("click", ()=>{
  const gstin = document.getElementById("bizGSTIN").value.trim().toUpperCase();
  if(gstin && !isValidGSTIN(gstin)){
    toast("That GSTIN doesn't look valid — check the format.");
    return;
  }
  DATA.settings.bizName = document.getElementById("bizName").value.trim();
  DATA.settings.bizGSTIN = gstin;
  DATA.settings.bizState = bizStateSel.value;
  saveData();
  updateSettingsSummary();
  renderAll();
  toast("Settings saved.");
});

/* ============================================================
   Tabs
   ============================================================ */
document.querySelectorAll(".tab").forEach(btn=>{
  btn.addEventListener("click", ()=>{
    document.querySelectorAll(".tab").forEach(b=>{ b.classList.remove("active"); b.setAttribute("aria-selected","false"); });
    document.querySelectorAll(".tab-panel").forEach(p=>p.classList.remove("active"));
    btn.classList.add("active"); btn.setAttribute("aria-selected","true");
    document.getElementById("panel-"+btn.dataset.tab).classList.add("active");
    if(btn.dataset.tab === "summary") renderSummary();
  });
});

/* ============================================================
   Party master (Vendors / Clients) — shared logic
   ============================================================ */
function makePartyModule(prefix, dataKey, formId){
  const form = document.getElementById(formId);
  const nameEl = document.getElementById(`${prefix}-name`);
  const gstinEl = document.getElementById(`${prefix}-gstin`);
  const addrEl = document.getElementById(`${prefix}-address`);
  const tbody = document.getElementById(`${prefix}-tbody`);
  const empty = document.getElementById(`${prefix}-empty`);
  const cancelBtn = document.getElementById(`${prefix}-cancelEdit`);
  let editingId = null;

  function resetForm(){
    form.reset(); editingId = null;
    cancelBtn.hidden = true;
    form.querySelector('button[type="submit"]').textContent = dataKey === "vendors" ? "Save vendor" : "Save client";
  }
  cancelBtn.addEventListener("click", resetForm);

  form.addEventListener("submit", (e)=>{
    e.preventDefault();
    const gstin = gstinEl.value.trim().toUpperCase();
    if(!isValidGSTIN(gstin)){
      toast("That GSTIN doesn't look valid — check the format (15 characters).");
      return;
    }
    const state = stateCodeFromGSTIN(gstin);
    const record = {
      id: editingId || uid(),
      name: nameEl.value.trim(),
      gstin, state,
      address: addrEl.value.trim()
    };
    const list = DATA[dataKey];
    const idx = list.findIndex(v=>v.id===record.id);
    if(idx>=0) list[idx] = record; else list.push(record);
    saveData();
    resetForm();
    render();
    renderPartyDatalists();
    toast(idx>=0 ? "Updated." : "Saved.");
  });

  function render(){
    const list = DATA[dataKey];
    empty.hidden = list.length>0;
    tbody.innerHTML = list.map(v=>`
      <tr>
        <td>${escapeHtml(v.name)}</td>
        <td>${escapeHtml(v.gstin)}</td>
        <td>${escapeHtml(STATE_NAME[v.state]||"—")}</td>
        <td>${escapeHtml(v.address||"—")}</td>
        <td class="row-actions">
          <button class="icon-btn" data-act="edit" data-id="${v.id}" title="Edit">✎</button>
          <button class="icon-btn" data-act="del" data-id="${v.id}" title="Delete">✕</button>
        </td>
      </tr>`).join("");

    tbody.querySelectorAll('[data-act="edit"]').forEach(b=>b.addEventListener("click", ()=>{
      const v = DATA[dataKey].find(x=>x.id===b.dataset.id);
      nameEl.value = v.name; gstinEl.value = v.gstin; addrEl.value = v.address||"";
      editingId = v.id;
      cancelBtn.hidden = false;
      form.querySelector('button[type="submit"]').textContent = "Save changes";
      form.scrollIntoView({behavior:"smooth", block:"start"});
    }));
    tbody.querySelectorAll('[data-act="del"]').forEach(b=>b.addEventListener("click", ()=>{
      if(!confirm("Delete this record? This won't affect past entries already saved.")) return;
      DATA[dataKey] = DATA[dataKey].filter(x=>x.id!==b.dataset.id);
      saveData(); render(); renderPartyDatalists();
    }));
  }
  return { render };
}
function escapeHtml(s){
  return String(s??"").replace(/[&<>"']/g, m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[m]));
}
const vendorModule = makePartyModule("v", "vendors", "form-vendor");
const clientModule = makePartyModule("c", "clients", "form-client");

function renderPartyDatalists(){
  document.getElementById("vendorList").innerHTML =
    DATA.vendors.map(v=>`<option value="${escapeHtml(v.name)}">`).join("");
  document.getElementById("clientList").innerHTML =
    DATA.clients.map(c=>`<option value="${escapeHtml(c.name)}">`).join("");
}

/* ============================================================
   Purchase / Sale entries — shared logic
   ============================================================ */
function makeEntryModule(prefix, dataKey, partyKey, formId){
  const form = document.getElementById(formId);
  const nameEl = document.getElementById(`${prefix}-partyName`);
  const gstinEl = document.getElementById(`${prefix}-gstin`);
  const dateEl = document.getElementById(`${prefix}-date`);
  const billEl = document.getElementById(`${prefix}-billNo`);
  const taxableEl = document.getElementById(`${prefix}-taxable`);
  const rateEl = document.getElementById(`${prefix}-rate`);
  const rateCustomWrap = document.getElementById(`${prefix}-rateCustomWrap`);
  const rateCustomEl = document.getElementById(`${prefix}-rateCustom`);
  const freightEl = document.getElementById(`${prefix}-freight`);
  const preview = document.getElementById(`${prefix}-preview`);
  const dupeWarning = document.getElementById(`${prefix}-dupeWarning`);
  const tbody = document.getElementById(`${prefix}-tbody`);
  const empty = document.getElementById(`${prefix}-empty`);
  const cancelBtn = document.getElementById(`${prefix}-cancelEdit`);
  const submitBtn = document.getElementById(`${prefix}-submitBtn`);
  const searchEl = document.getElementById(`${prefix}-search`);
  const typeFilterEl = document.getElementById(`${prefix}-typeFilter`);
  const monthFilterEl = document.getElementById(`${prefix}-monthFilter`);
  let editingId = null;

  dateEl.value = todayISO();

  // Autofill GSTIN when a saved party name is chosen
  nameEl.addEventListener("input", ()=>{
    const match = DATA[partyKey].find(p=>p.name.toLowerCase() === nameEl.value.trim().toLowerCase());
    if(match) gstinEl.value = match.gstin;
    checkDupe();
  });
  rateEl.addEventListener("change", ()=>{
    rateCustomWrap.hidden = rateEl.value !== "custom";
    updatePreview();
  });
  [gstinEl, taxableEl, rateCustomEl, freightEl].forEach(el=>el.addEventListener("input", updatePreview));
  billEl.addEventListener("input", checkDupe);
  gstinEl.addEventListener("input", checkDupe);
  gstinEl.addEventListener("blur", ()=>{ gstinEl.value = gstinEl.value.trim().toUpperCase(); updatePreview(); });

  function currentRate(){
    return rateEl.value === "custom" ? (parseFloat(rateCustomEl.value)||0) : parseFloat(rateEl.value);
  }
  function isInState(){
    const partyState = stateCodeFromGSTIN(gstinEl.value.trim().toUpperCase());
    return partyState && partyState === DATA.settings.bizState;
  }
  function updatePreview(){
    const taxable = parseFloat(taxableEl.value)||0;
    const freight = parseFloat(freightEl.value)||0;
    const rate = currentRate();
    if(!taxable && !freight){ preview.classList.remove("show"); return; }
    const partyCode = stateCodeFromGSTIN(gstinEl.value.trim().toUpperCase());
    const inState = isInState();
    const {cgst, sgst, igst} = computeGst(taxable, rate, inState);
    const total = round2(taxable + cgst + sgst + igst + freight);
    let taxLine = inState
      ? `CGST ₹${fmt(cgst)} + SGST ₹${fmt(sgst)}`
      : `IGST ₹${fmt(igst)}`;
    let stateLine = partyCode
      ? (inState ? `In-state (${STATE_NAME[partyCode]})` : `Out-of-state (${STATE_NAME[partyCode]})`)
      : "Enter a valid GSTIN to detect state";
    preview.classList.add("show");
    preview.textContent = `${rate}% on ₹${fmt(taxable)} → ${taxLine} · Freight ₹${fmt(freight)} · Total ₹${fmt(total)} · ${stateLine}`;
  }
  function checkDupe(){
    const name = nameEl.value.trim().toLowerCase();
    const bill = billEl.value.trim().toLowerCase();
    if(!name || !bill){ dupeWarning.hidden = true; return; }
    const dupe = DATA[dataKey].some(e =>
      e.id !== editingId && e.party.toLowerCase()===name && e.billNo.toLowerCase()===bill);
    dupeWarning.hidden = !dupe;
  }

  function resetForm(){
    form.reset(); dateEl.value = todayISO();
    freightEl.value = "0"; rateEl.value = "18"; rateCustomWrap.hidden = true;
    editingId = null; cancelBtn.hidden = true; dupeWarning.hidden = true;
    preview.classList.remove("show");
    submitBtn.textContent = "Add entry";
  }
  cancelBtn.addEventListener("click", resetForm);

  form.addEventListener("submit", (e)=>{
    e.preventDefault();
    const gstin = gstinEl.value.trim().toUpperCase();
    if(!isValidGSTIN(gstin)){
      toast("That GSTIN doesn't look valid — check the format (15 characters).");
      return;
    }
    const state = stateCodeFromGSTIN(gstin);
    const inState = isInState();
    const taxable = round2(parseFloat(taxableEl.value)||0);
    const freight = round2(parseFloat(freightEl.value)||0);
    const rate = currentRate();
    const {cgst, sgst, igst} = computeGst(taxable, rate, inState);
    const total = round2(taxable + cgst + sgst + igst + freight);

    const record = {
      id: editingId || uid(),
      party: nameEl.value.trim(),
      gstin, state,
      date: dateEl.value,
      billNo: billEl.value.trim(),
      taxable, rate, cgst, sgst, igst, freight, total,
      inState: !!inState
    };
    const list = DATA[dataKey];
    const idx = list.findIndex(x=>x.id===record.id);
    if(idx>=0) list[idx]=record; else list.push(record);
    saveData();
    resetForm();
    render();
    toast(idx>=0 ? "Entry updated." : "Entry added.");
  });

  function matchesFilters(e){
    const q = searchEl.value.trim().toLowerCase();
    if(q && !(e.party.toLowerCase().includes(q) || e.billNo.toLowerCase().includes(q))) return false;
    const tf = typeFilterEl.value;
    if(tf==="in" && !e.inState) return false;
    if(tf==="out" && e.inState) return false;
    const mf = monthFilterEl.value;
    if(mf && e.date.slice(0,7)!==mf) return false;
    return true;
  }

  function render(){
    const list = DATA[dataKey].filter(matchesFilters).slice().sort((a,b)=> b.date.localeCompare(a.date));
    empty.hidden = list.length>0;
    tbody.innerHTML = list.map(e=>`
      <tr>
        <td>${escapeHtml(e.party)}</td>
        <td>${escapeHtml(e.gstin)}</td>
        <td>${e.date}</td>
        <td>${escapeHtml(e.billNo)}</td>
        <td class="num">${fmt(e.taxable)}</td>
        <td class="num">${fmt(e.cgst)}</td>
        <td class="num">${fmt(e.sgst)}</td>
        <td class="num">${fmt(e.igst)}</td>
        <td class="num">${fmt(e.freight)}</td>
        <td class="num">${fmt(e.total)}</td>
        <td><span class="badge ${e.inState?"badge-in":"badge-out"}">${e.inState?"In-state":"Out-of-state"}</span></td>
        <td class="row-actions">
          <button class="icon-btn" data-act="edit" data-id="${e.id}" title="Edit">✎</button>
          <button class="icon-btn" data-act="del" data-id="${e.id}" title="Delete">✕</button>
        </td>
      </tr>`).join("");

    tbody.querySelectorAll('[data-act="edit"]').forEach(b=>b.addEventListener("click", ()=>{
      const e = DATA[dataKey].find(x=>x.id===b.dataset.id);
      nameEl.value = e.party; gstinEl.value = e.gstin; dateEl.value = e.date; billEl.value = e.billNo;
      taxableEl.value = e.taxable; freightEl.value = e.freight;
      if(GST_RATES.includes(e.rate)){ rateEl.value = String(e.rate); rateCustomWrap.hidden = true; }
      else { rateEl.value = "custom"; rateCustomWrap.hidden = false; rateCustomEl.value = e.rate; }
      editingId = e.id;
      cancelBtn.hidden = false;
      submitBtn.textContent = "Save changes";
      updatePreview(); checkDupe();
      form.scrollIntoView({behavior:"smooth", block:"start"});
    }));
    tbody.querySelectorAll('[data-act="del"]').forEach(b=>b.addEventListener("click", ()=>{
      if(!confirm("Delete this entry?")) return;
      DATA[dataKey] = DATA[dataKey].filter(x=>x.id!==b.dataset.id);
      saveData(); render();
    }));
  }

  [searchEl, typeFilterEl, monthFilterEl].forEach(el=>el.addEventListener("input", render));
  resetForm();
  return { render };
}
const purchaseModule = makeEntryModule("p", "purchases", "vendors", "form-purchase");
const saleModule = makeEntryModule("s", "sales", "clients", "form-sale");

/* ============================================================
   Summary
   ============================================================ */
const sumFrom = document.getElementById("sum-from");
const sumTo = document.getElementById("sum-to");
document.getElementById("sum-clearRange").addEventListener("click", ()=>{
  sumFrom.value=""; sumTo.value=""; renderSummary();
});
[sumFrom, sumTo].forEach(el=>el.addEventListener("input", renderSummary));

function inRange(dateStr){
  const ym = dateStr.slice(0,7);
  if(sumFrom.value && ym < sumFrom.value) return false;
  if(sumTo.value && ym > sumTo.value) return false;
  return true;
}
function sumField(list, field){ return round2(list.reduce((s,e)=>s+(e[field]||0),0)); }

function renderSummary(){
  const purchases = DATA.purchases.filter(e=>inRange(e.date));
  const sales = DATA.sales.filter(e=>inRange(e.date));

  const inputGst = round2(sumField(purchases,"cgst")+sumField(purchases,"sgst")+sumField(purchases,"igst"));
  const outputGst = round2(sumField(sales,"cgst")+sumField(sales,"sgst")+sumField(sales,"igst"));
  const netPayable = round2(outputGst - inputGst);

  const cards = [
    { label: "Total purchases", value: sumField(purchases,"total") },
    { label: "Total sales", value: sumField(sales,"total") },
    { label: "Input GST (on purchases)", value: inputGst },
    { label: "Output GST (on sales)", value: outputGst },
  ];
  document.getElementById("summaryCards").innerHTML = cards.map(c=>
    `<div class="summary-card"><div class="label">${c.label}</div><div class="value">₹${fmt(c.value)}</div></div>`
  ).join("") + `
    <div class="summary-card ${netPayable>=0?"highlight":"negative"}">
      <div class="label">${netPayable>=0 ? "Net GST payable" : "Net GST credit carried forward"}</div>
      <div class="value">₹${fmt(Math.abs(netPayable))}</div>
    </div>`;

  document.getElementById("sum-purchase-table").innerHTML = `
    <tr><td>Taxable value</td><td>₹${fmt(sumField(purchases,"taxable"))}</td></tr>
    <tr><td>CGST</td><td>₹${fmt(sumField(purchases,"cgst"))}</td></tr>
    <tr><td>SGST</td><td>₹${fmt(sumField(purchases,"sgst"))}</td></tr>
    <tr><td>IGST</td><td>₹${fmt(sumField(purchases,"igst"))}</td></tr>
    <tr><td>Freight</td><td>₹${fmt(sumField(purchases,"freight"))}</td></tr>
    <tr><td>Entries</td><td>${purchases.length}</td></tr>`;
  document.getElementById("sum-sale-table").innerHTML = `
    <tr><td>Taxable value</td><td>₹${fmt(sumField(sales,"taxable"))}</td></tr>
    <tr><td>CGST</td><td>₹${fmt(sumField(sales,"cgst"))}</td></tr>
    <tr><td>SGST</td><td>₹${fmt(sumField(sales,"sgst"))}</td></tr>
    <tr><td>IGST</td><td>₹${fmt(sumField(sales,"igst"))}</td></tr>
    <tr><td>Freight</td><td>₹${fmt(sumField(sales,"freight"))}</td></tr>
    <tr><td>Entries</td><td>${sales.length}</td></tr>`;
}

/* ============================================================
   Excel export / import (SheetJS)
   ============================================================ */
const ENTRY_HEADERS = ["Party","GSTIN","State Code","Date","Bill No","Taxable","Rate %","CGST","SGST","IGST","Freight","Total","In-State"];
function entryToRow(e){
  return [e.party, e.gstin, e.state, e.date, e.billNo, e.taxable, e.rate, e.cgst, e.sgst, e.igst, e.freight, e.total, e.inState ? "Yes" : "No"];
}
function rowToEntry(row){
  const [party,gstin,state,date,billNo,taxable,rate,cgst,sgst,igst,freight,total,inState] = row;
  return {
    id: uid(), party: String(party||"").trim(), gstin: String(gstin||"").trim().toUpperCase(),
    state: String(state||stateCodeFromGSTIN(gstin)||""), date: normalizeDate(date), billNo: String(billNo||"").trim(),
    taxable: Number(taxable)||0, rate: Number(rate)||0, cgst: Number(cgst)||0, sgst: Number(sgst)||0,
    igst: Number(igst)||0, freight: Number(freight)||0, total: Number(total)||0,
    inState: String(inState).toLowerCase()==="yes"
  };
}
function normalizeDate(d){
  if(d instanceof Date) return d.toISOString().slice(0,10);
  if(typeof d === "number"){ // Excel serial date
    const dt = XLSX.SSF.parse_date_code(d);
    return `${dt.y}-${String(dt.m).padStart(2,"0")}-${String(dt.d).padStart(2,"0")}`;
  }
  return String(d||"").slice(0,10);
}
const PARTY_HEADERS = ["Name","GSTIN","State Code","Address"];
function partyToRow(p){ return [p.name, p.gstin, p.state, p.address||""]; }
function rowToParty(row){
  const [name,gstin,state,address] = row;
  return { id: uid(), name: String(name||"").trim(), gstin: String(gstin||"").trim().toUpperCase(),
    state: String(state||stateCodeFromGSTIN(gstin)||""), address: String(address||"").trim() };
}

function buildWorkbook(){
  const wb = XLSX.utils.book_new();
  const purchaseWs = XLSX.utils.aoa_to_sheet([ENTRY_HEADERS, ...DATA.purchases.map(entryToRow)]);
  const saleWs = XLSX.utils.aoa_to_sheet([ENTRY_HEADERS, ...DATA.sales.map(entryToRow)]);
  const vendorWs = XLSX.utils.aoa_to_sheet([PARTY_HEADERS, ...DATA.vendors.map(partyToRow)]);
  const clientWs = XLSX.utils.aoa_to_sheet([PARTY_HEADERS, ...DATA.clients.map(partyToRow)]);
  const settingsWs = XLSX.utils.aoa_to_sheet([
    ["Business Name","Business GSTIN","Business State Code"],
    [DATA.settings.bizName, DATA.settings.bizGSTIN, DATA.settings.bizState]
  ]);
  XLSX.utils.book_append_sheet(wb, purchaseWs, "Purchase");
  XLSX.utils.book_append_sheet(wb, saleWs, "Sale");
  XLSX.utils.book_append_sheet(wb, vendorWs, "Vendors");
  XLSX.utils.book_append_sheet(wb, clientWs, "Clients");
  XLSX.utils.book_append_sheet(wb, settingsWs, "Settings");
  return wb;
}
function exportWorkbook(){
  const wb = buildWorkbook();
  const stamp = todayISO();
  XLSX.writeFile(wb, `gst-ledger-backup-${stamp}.xlsx`);
}
document.getElementById("p-exportBtn").addEventListener("click", ()=>{
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([ENTRY_HEADERS, ...DATA.purchases.map(entryToRow)]);
  XLSX.utils.book_append_sheet(wb, ws, "Purchase");
  XLSX.writeFile(wb, `purchase-register-${todayISO()}.xlsx`);
});
document.getElementById("s-exportBtn").addEventListener("click", ()=>{
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([ENTRY_HEADERS, ...DATA.sales.map(entryToRow)]);
  XLSX.utils.book_append_sheet(wb, ws, "Sale");
  XLSX.writeFile(wb, `sale-register-${todayISO()}.xlsx`);
});
document.getElementById("sum-exportBtn").addEventListener("click", exportWorkbook);

document.getElementById("importFile").addEventListener("change", (e)=>{
  const file = e.target.files[0];
  if(!file) return;
  const reader = new FileReader();
  reader.onload = (ev)=>{
    try{
      const wb = XLSX.read(ev.target.result, { type: "array", cellDates:false });
      const sheetRows = (name)=>{
        if(!wb.Sheets[name]) return [];
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header:1, raw:true });
        return rows.slice(1).filter(r=>r && r.length && r[0]);
      };
      const newPurchases = sheetRows("Purchase").map(rowToEntry);
      const newSales = sheetRows("Sale").map(rowToEntry);
      const newVendors = sheetRows("Vendors").map(rowToParty);
      const newClients = sheetRows("Clients").map(rowToParty);
      const settingsRows = wb.Sheets["Settings"] ? XLSX.utils.sheet_to_json(wb.Sheets["Settings"], {header:1}) : null;

      if(!confirm(`Import ${newPurchases.length} purchase, ${newSales.length} sale, ${newVendors.length} vendor and ${newClients.length} client rows? This replaces your current data — export a backup first if unsure.`)) {
        e.target.value = ""; return;
      }
      DATA.purchases = newPurchases;
      DATA.sales = newSales;
      DATA.vendors = newVendors;
      DATA.clients = newClients;
      if(settingsRows && settingsRows[1]){
        DATA.settings.bizName = settingsRows[1][0]||"";
        DATA.settings.bizGSTIN = settingsRows[1][1]||"";
        DATA.settings.bizState = settingsRows[1][2]||"23";
      }
      saveData();
      renderAll();
      toast("Backup imported.");
    }catch(err){
      console.error(err);
      toast("Couldn't read that file — make sure it's a backup exported from this app.");
    }
    e.target.value = "";
  };
  reader.readAsArrayBuffer(file);
});

/* ============================================================
   Boot
   ============================================================ */
function renderAll(){
  loadSettingsIntoForm();
  vendorModule.render();
  clientModule.render();
  renderPartyDatalists();
  purchaseModule.render();
  saleModule.render();
  renderSummary();
}
renderAll();

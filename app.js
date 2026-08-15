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

/* ---------- City → state lookup (for autofill; not exhaustive — falls back to
   manual state selection for any city not in this list, or you can just type
   the state name into the city box if unsure) ---------- */
const CITY_STATE_PAIRS = [
  // Madhya Pradesh
  ["Indore","23"],["Bhopal","23"],["Gwalior","23"],["Jabalpur","23"],["Ujjain","23"],["Dewas","23"],
  ["Ratlam","23"],["Sagar","23"],["Rewa","23"],["Satna","23"],["Pithampur","23"],["Mandsaur","23"],
  ["Neemuch","23"],["Khandwa","23"],["Burhanpur","23"],["Chanderi","23"],["Maheshwar","23"],
  // Maharashtra
  ["Mumbai","27"],["Bombay","27"],["Pune","27"],["Nagpur","27"],["Nashik","27"],["Aurangabad","27"],
  ["Chhatrapati Sambhajinagar","27"],["Solapur","27"],["Bhiwandi","27"],["Ichalkaranji","27"],
  ["Malegaon","27"],["Kolhapur","27"],["Sangli","27"],["Thane","27"],["Navi Mumbai","27"],
  // Gujarat
  ["Ahmedabad","24"],["Surat","24"],["Vadodara","24"],["Baroda","24"],["Rajkot","24"],["Bhavnagar","24"],
  ["Jamnagar","24"],["Gandhinagar","24"],["Anand","24"],
  // Rajasthan
  ["Jaipur","08"],["Jodhpur","08"],["Udaipur","08"],["Kota","08"],["Bhilwara","08"],["Ajmer","08"],
  ["Bikaner","08"],["Pali","08"],["Balotra","08"],["Sanganer","08"],
  // Delhi
  ["Delhi","07"],["New Delhi","07"],
  // Uttar Pradesh
  ["Lucknow","09"],["Kanpur","09"],["Noida","09"],["Ghaziabad","09"],["Agra","09"],["Varanasi","09"],
  ["Meerut","09"],["Moradabad","09"],["Bareilly","09"],["Aligarh","09"],["Prayagraj","09"],["Allahabad","09"],
  // Punjab
  ["Ludhiana","03"],["Amritsar","03"],["Jalandhar","03"],["Patiala","03"],["Bathinda","03"],
  // Haryana
  ["Gurugram","06"],["Gurgaon","06"],["Faridabad","06"],["Panipat","06"],["Karnal","06"],["Hisar","06"],
  ["Sonipat","06"],["Panchkula","06"],
  // West Bengal
  ["Kolkata","19"],["Calcutta","19"],["Howrah","19"],["Durgapur","19"],["Siliguri","19"],["Asansol","19"],
  // Tamil Nadu
  ["Chennai","33"],["Madras","33"],["Coimbatore","33"],["Tiruppur","33"],["Tirupur","33"],["Erode","33"],
  ["Madurai","33"],["Salem","33"],["Karur","33"],
  // Karnataka
  ["Bengaluru","29"],["Bangalore","29"],["Mysuru","29"],["Mysore","29"],["Hubballi","29"],["Belagavi","29"],
  // Telangana
  ["Hyderabad","36"],["Warangal","36"],["Secunderabad","36"],
  // Andhra Pradesh
  ["Vijayawada","37"],["Visakhapatnam","37"],["Guntur","37"],["Tirupati","37"],
  // Kerala
  ["Kochi","32"],["Cochin","32"],["Thiruvananthapuram","32"],["Trivandrum","32"],["Kozhikode","32"],["Kannur","32"],
  // Bihar
  ["Patna","10"],["Gaya","10"],["Bhagalpur","10"],["Muzaffarpur","10"],
  // Jharkhand
  ["Ranchi","20"],["Jamshedpur","20"],["Dhanbad","20"],
  // Odisha
  ["Bhubaneswar","21"],["Cuttack","21"],["Rourkela","21"],
  // Chhattisgarh
  ["Raipur","22"],["Bhilai","22"],["Bilaspur","22"],["Durg","22"],
  // Assam
  ["Guwahati","18"],["Silchar","18"],
  // Uttarakhand
  ["Dehradun","05"],["Haridwar","05"],["Roorkee","05"],
  // Himachal Pradesh
  ["Shimla","02"],["Solan","02"],
  // Jammu & Kashmir
  ["Srinagar","01"],["Jammu","01"],
  // Goa
  ["Panaji","30"],["Margao","30"],["Vasco da Gama","30"],
  // Puducherry
  ["Puducherry","34"],["Pondicherry","34"],
  // Chandigarh
  ["Chandigarh","04"],
  // North-east & other UT capitals
  ["Shillong","17"],["Agartala","16"],["Imphal","14"],["Aizawl","15"],["Kohima","13"],["Itanagar","12"],["Gangtok","11"],
  ["Daman","26"],["Diu","26"],["Silvassa","26"],
  ["Port Blair","35"],["Kavaratti","31"],["Leh","38"]
];
function normalizeCity(s){ return String(s||"").trim().toLowerCase().replace(/\s+/g," "); }
const CITY_STATE = Object.fromEntries(CITY_STATE_PAIRS.map(([name,code])=>[normalizeCity(name),code]));
function stateFromCity(city){ return CITY_STATE[normalizeCity(city)] || null; }
function populateCityDatalist(){
  const dl = document.getElementById("cityList");
  const names = [...new Set(CITY_STATE_PAIRS.map(([n])=>n))].sort();
  dl.innerHTML = names.map(n=>`<option value="${escapeHtml(n)}">`).join("");
}

/* ---------- Storage ---------- */
const STORAGE_KEY = "gstLedgerData_v1";

function defaultData(){
  return {
    settings: { bizName: "", bizCity: "Indore", bizGSTIN: "", bizState: "23" },
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
function populateStateSelect(sel, includeBlank){
  const blank = includeBlank ? `<option value="">Select state…</option>` : "";
  sel.innerHTML = blank + STATE_CODES.map(([c,n])=>`<option value="${c}">${c} — ${n}</option>`).join("");
}

/* ============================================================
   Settings
   ============================================================ */
const settingsToggle = document.getElementById("settingsToggle");
const settingsPanel = document.getElementById("settingsPanel");
const bizStateSel = document.getElementById("bizState");
const bizCityEl = document.getElementById("bizCity");
const bizGSTINEl = document.getElementById("bizGSTIN");
populateStateSelect(bizStateSel);
bizCityEl.addEventListener("input", ()=>{
  const code = stateFromCity(bizCityEl.value);
  if(code) bizStateSel.value = code;
});
bizGSTINEl.addEventListener("input", ()=>{
  const code = stateCodeFromGSTIN(bizGSTINEl.value.trim().toUpperCase());
  if(code) bizStateSel.value = code;
});

function loadSettingsIntoForm(){
  document.getElementById("bizName").value = DATA.settings.bizName;
  bizCityEl.value = DATA.settings.bizCity || "";
  bizGSTINEl.value = DATA.settings.bizGSTIN;
  bizStateSel.value = DATA.settings.bizState || "23";
  updateSettingsSummary();
}
function updateSettingsSummary(){
  const name = DATA.settings.bizName || "your business";
  const stateName = STATE_NAME[DATA.settings.bizState] || "—";
  document.getElementById("settingsBtnLabel").textContent =
    DATA.settings.bizName ? `${DATA.settings.bizName} · ${stateName}` : "Business settings";
  document.getElementById("bizStateAuto").textContent =
    `Entries are marked in-state when the party's city/GSTIN resolves to ${stateName}.`;
}
settingsToggle.addEventListener("click", ()=>{
  const open = settingsPanel.hidden;
  settingsPanel.hidden = !open;
  settingsToggle.setAttribute("aria-expanded", String(open));
});
document.getElementById("saveSettings").addEventListener("click", ()=>{
  const gstin = bizGSTINEl.value.trim().toUpperCase();
  if(gstin && !isValidGSTIN(gstin)){
    toast("That GSTIN doesn't look valid — check the format.");
    return;
  }
  DATA.settings.bizName = document.getElementById("bizName").value.trim();
  DATA.settings.bizCity = bizCityEl.value.trim();
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
  const cityEl = document.getElementById(`${prefix}-city`);
  const stateEl = document.getElementById(`${prefix}-state`);
  const addrEl = document.getElementById(`${prefix}-address`);
  const tbody = document.getElementById(`${prefix}-tbody`);
  const empty = document.getElementById(`${prefix}-empty`);
  const cancelBtn = document.getElementById(`${prefix}-cancelEdit`);
  let editingId = null;

  populateStateSelect(stateEl, true);
  gstinEl.addEventListener("input", ()=>{
    const code = stateCodeFromGSTIN(gstinEl.value.trim().toUpperCase());
    if(code) stateEl.value = code; // auto-fill; user can still override below
  });
  cityEl.addEventListener("input", ()=>{
    const code = stateFromCity(cityEl.value);
    if(code) stateEl.value = code; // auto-fill; user can still override below
  });

  function resetForm(){
    form.reset(); editingId = null;
    stateEl.value = "";
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
    if(!stateEl.value){
      toast("Please select the state.");
      return;
    }
    const state = stateEl.value;
    const record = {
      id: editingId || uid(),
      name: nameEl.value.trim(),
      gstin, city: cityEl.value.trim(), state,
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
        <td>${escapeHtml(v.city||"—")}</td>
        <td>${escapeHtml(STATE_NAME[v.state]||"—")}</td>
        <td>${escapeHtml(v.address||"—")}</td>
        <td class="row-actions">
          <button class="icon-btn" data-act="edit" data-id="${v.id}" title="Edit">✎</button>
          <button class="icon-btn" data-act="del" data-id="${v.id}" title="Delete">✕</button>
        </td>
      </tr>`).join("");

    tbody.querySelectorAll('[data-act="edit"]').forEach(b=>b.addEventListener("click", ()=>{
      const v = DATA[dataKey].find(x=>x.id===b.dataset.id);
      nameEl.value = v.name; gstinEl.value = v.gstin; cityEl.value = v.city||""; stateEl.value = v.state||""; addrEl.value = v.address||"";
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
  const cityEl = document.getElementById(`${prefix}-city`);
  const stateEl = document.getElementById(`${prefix}-state`);
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
  populateStateSelect(stateEl, true);

  // Autofill GSTIN + city + state when a saved party name is chosen
  nameEl.addEventListener("input", ()=>{
    const match = DATA[partyKey].find(p=>p.name.toLowerCase() === nameEl.value.trim().toLowerCase());
    if(match){ gstinEl.value = match.gstin; if(match.city) cityEl.value = match.city; if(match.state) stateEl.value = match.state; }
    checkDupe(); updatePreview();
  });
  rateEl.addEventListener("change", ()=>{
    rateCustomWrap.hidden = rateEl.value !== "custom";
    updatePreview();
  });
  [gstinEl, cityEl, stateEl, taxableEl, rateCustomEl, freightEl].forEach(el=>el.addEventListener("input", updatePreview));
  billEl.addEventListener("input", checkDupe);
  gstinEl.addEventListener("input", ()=>{
    // Auto-fill state from GSTIN, but never overrides a state the user just picked by hand
    const code = stateCodeFromGSTIN(gstinEl.value.trim().toUpperCase());
    if(code) stateEl.value = code;
    checkDupe();
  });
  gstinEl.addEventListener("blur", ()=>{ gstinEl.value = gstinEl.value.trim().toUpperCase(); updatePreview(); });
  cityEl.addEventListener("input", ()=>{
    const code = stateFromCity(cityEl.value);
    if(code) stateEl.value = code;
  });

  function currentRate(){
    return rateEl.value === "custom" ? (parseFloat(rateCustomEl.value)||0) : parseFloat(rateEl.value);
  }
  function isInState(){
    return !!stateEl.value && stateEl.value === DATA.settings.bizState;
  }
  function updatePreview(){
    const taxable = parseFloat(taxableEl.value)||0;
    const freight = parseFloat(freightEl.value)||0;
    const rate = currentRate();
    if(!taxable && !freight){ preview.classList.remove("show"); return; }
    const inState = isInState();
    const {cgst, sgst, igst} = computeGst(taxable, rate, inState);
    const total = round2(taxable + cgst + sgst + igst + freight);
    let taxLine = inState
      ? `CGST ₹${fmt(cgst)} + SGST ₹${fmt(sgst)}`
      : `IGST ₹${fmt(igst)}`;
    let stateLine = stateEl.value
      ? (inState ? `In-state (${STATE_NAME[stateEl.value]})` : `Out-of-state (${STATE_NAME[stateEl.value]})`)
      : "Type a city (or enter a GSTIN) to auto-detect the state";
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
    freightEl.value = "0"; rateEl.value = "18"; rateCustomWrap.hidden = true; stateEl.value = "";
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
    if(!stateEl.value){
      toast("Please select the party's state.");
      return;
    }
    const state = stateEl.value;
    const inState = isInState();
    const taxable = round2(parseFloat(taxableEl.value)||0);
    const freight = round2(parseFloat(freightEl.value)||0);
    const rate = currentRate();
    const {cgst, sgst, igst} = computeGst(taxable, rate, inState);
    const total = round2(taxable + cgst + sgst + igst + freight);

    const record = {
      id: editingId || uid(),
      party: nameEl.value.trim(),
      gstin, city: cityEl.value.trim(), state,
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
        <td>${escapeHtml(e.city||"—")}</td>
        <td>${escapeHtml(STATE_NAME[e.state]||"—")}</td>
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
      nameEl.value = e.party; gstinEl.value = e.gstin; cityEl.value = e.city||""; stateEl.value = e.state||""; dateEl.value = e.date; billEl.value = e.billNo;
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
const ENTRY_HEADERS = ["Party","GSTIN","City","State Code","Date","Bill No","Taxable","Rate %","CGST","SGST","IGST","Freight","Total","In-State"];
function entryToRow(e){
  return [e.party, e.gstin, e.city||"", e.state, e.date, e.billNo, e.taxable, e.rate, e.cgst, e.sgst, e.igst, e.freight, e.total, e.inState ? "Yes" : "No"];
}
function rowToEntry(o){
  const gstin = String(o["GSTIN"]||"").trim().toUpperCase();
  const city = String(o["City"]||"").trim();
  return {
    id: uid(), party: String(o["Party"]||"").trim(), gstin,
    city, state: String(o["State Code"]||stateFromCity(city)||stateCodeFromGSTIN(gstin)||""),
    date: normalizeDate(o["Date"]), billNo: String(o["Bill No"]||"").trim(),
    taxable: Number(o["Taxable"])||0, rate: Number(o["Rate %"])||0,
    cgst: Number(o["CGST"])||0, sgst: Number(o["SGST"])||0, igst: Number(o["IGST"])||0,
    freight: Number(o["Freight"])||0, total: Number(o["Total"])||0,
    inState: String(o["In-State"]).toLowerCase()==="yes"
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
const PARTY_HEADERS = ["Name","GSTIN","City","State Code","Address"];
function partyToRow(p){ return [p.name, p.gstin, p.city||"", p.state, p.address||""]; }
function rowToParty(o){
  const gstin = String(o["GSTIN"]||"").trim().toUpperCase();
  const city = String(o["City"]||"").trim();
  return {
    id: uid(), name: String(o["Name"]||"").trim(), gstin,
    city, state: String(o["State Code"]||stateFromCity(city)||stateCodeFromGSTIN(gstin)||""),
    address: String(o["Address"]||"").trim()
  };
}

function buildWorkbook(){
  const wb = XLSX.utils.book_new();
  const purchaseWs = XLSX.utils.aoa_to_sheet([ENTRY_HEADERS, ...DATA.purchases.map(entryToRow)]);
  const saleWs = XLSX.utils.aoa_to_sheet([ENTRY_HEADERS, ...DATA.sales.map(entryToRow)]);
  const vendorWs = XLSX.utils.aoa_to_sheet([PARTY_HEADERS, ...DATA.vendors.map(partyToRow)]);
  const clientWs = XLSX.utils.aoa_to_sheet([PARTY_HEADERS, ...DATA.clients.map(partyToRow)]);
  const settingsWs = XLSX.utils.aoa_to_sheet([
    ["Business Name","Business City","Business GSTIN","Business State Code"],
    [DATA.settings.bizName, DATA.settings.bizCity||"", DATA.settings.bizGSTIN, DATA.settings.bizState]
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
      // Header-driven so backups from older versions of this app (before the
      // City column existed, or with columns in a different order) still import cleanly.
      const sheetObjs = (name)=>{
        if(!wb.Sheets[name]) return [];
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header:1, raw:true });
        const headers = rows[0]||[];
        return rows.slice(1).filter(r=>r && r.length && r[0])
          .map(r=>Object.fromEntries(headers.map((h,i)=>[h, r[i]])));
      };
      const newPurchases = sheetObjs("Purchase").map(rowToEntry);
      const newSales = sheetObjs("Sale").map(rowToEntry);
      const newVendors = sheetObjs("Vendors").map(rowToParty);
      const newClients = sheetObjs("Clients").map(rowToParty);
      const settingsRows = wb.Sheets["Settings"] ? XLSX.utils.sheet_to_json(wb.Sheets["Settings"], {header:1}) : null;

      if(!confirm(`Import ${newPurchases.length} purchase, ${newSales.length} sale, ${newVendors.length} vendor and ${newClients.length} client rows? This replaces your current data — export a backup first if unsure.`)) {
        e.target.value = ""; return;
      }
      DATA.purchases = newPurchases;
      DATA.sales = newSales;
      DATA.vendors = newVendors;
      DATA.clients = newClients;
      if(settingsRows && settingsRows[1]){
        const sHeaders = settingsRows[0]||[];
        const sVals = settingsRows[1];
        const sObj = Object.fromEntries(sHeaders.map((h,i)=>[h, sVals[i]]));
        DATA.settings.bizName = sObj["Business Name"]||"";
        DATA.settings.bizCity = sObj["Business City"]||"";
        DATA.settings.bizGSTIN = sObj["Business GSTIN"]||"";
        DATA.settings.bizState = sObj["Business State Code"]||"23";
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
populateCityDatalist();
renderAll();

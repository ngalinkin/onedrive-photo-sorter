// ===== CONFIG =====
const msalConfig = {
  auth: {
    clientId: "18eebe37-a762-4148-b425-4ee4d79674bf",        // <-- paste your Azure App Registration "Application (client) ID"
    authority: "https://login.microsoftonline.com/consumers",
    redirectUri: location.origin
  },
  cache: { cacheLocation: "localStorage" }
};
const graphScopes = ["Files.ReadWrite", "offline_access"]; // minimal

// ===== MSAL / GRAPH =====
const msalInstance = new msal.PublicClientApplication(msalConfig);
let account = null;

async function login() {
  try {
    const res = await msalInstance.loginPopup({ scopes: graphScopes });
    account = res.account;
    setStatus("Signed in");
    await initRoot();
  } catch (e) { console.error(e); setStatus("Login failed"); }
}

async function getToken() {
  if (!account) account = msalInstance.getAllAccounts()[0];
  const req = { account, scopes: graphScopes };
  try { return (await msalInstance.acquireTokenSilent(req)).accessToken; }
  catch { return (await msalInstance.acquireTokenPopup(req)).accessToken; }
}

// one retry on 429 using Retry-After
async function g(path, opts = {}, attempt = 0) {
  const token = await getToken();
  const r = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    ...opts,
    headers: { "Authorization": `Bearer ${token}`, "Content-Type": "application/json", ...(opts.headers||{}) }
  });

  if (r.status === 429 && attempt < 1) {
    const ra = parseInt(r.headers.get("Retry-After") || "2", 10) * 1000;
    await new Promise(res => setTimeout(res, isFinite(ra) ? ra : 2000));
    return g(path, opts, attempt + 1);
  }
  if (!r.ok) throw new Error(await r.text());
  try { return await r.json(); } catch { return {}; }
}
async function authedFetch(fullUrl, attempt = 0) {
  const token = await getToken();
  const r = await fetch(fullUrl, { headers: { Authorization: `Bearer ${token}` } });
  if (r.status === 429 && attempt < 1) {
    const ra = parseInt(r.headers.get("Retry-After") || "2", 10) * 1000;
    await new Promise(res => setTimeout(res, isFinite(ra) ? ra : 2000));
    return authedFetch(fullUrl, attempt + 1);
  }
  if (!r.ok) throw new Error(await r.text());
  return r.json();
}

// ===== STATE (per folder) =====
const storeKey = (folderId) => `ps:${folderId}`;
function loadState(folderId) {
  return JSON.parse(localStorage.getItem(storeKey(folderId)) ||
    `{"processed":{},"soft":{},"cursor":null,"hideProcessed":false,"filterMode":"all","pageIndex":0}`);
}
function saveState(folderId, s) { localStorage.setItem(storeKey(folderId), JSON.stringify(s)); }

// ===== DOM =====
const grid = document.getElementById("grid");
const upBtn = document.getElementById("upBtn");
const breadcrumb = document.getElementById("breadcrumb");
const subfolderSelect = document.getElementById("subfolderSelect");
const filterModeSel = document.getElementById("filterMode");
const toggleHideProcessedBtn = document.getElementById("toggleHideProcessed");
const nextPageBtn = document.getElementById("nextPage");
const prevPageBtn = document.getElementById("prevPage");
const deleteDeclinedBtn = document.getElementById("deleteDeclined");
const downloadChosenBtn = document.getElementById("downloadChosen");
const downloadNotChosenBtn = document.getElementById("downloadNotChosen");
const statusEl = document.getElementById("status");
const counterMain = document.getElementById("counterMain");
const counterPage = document.getElementById("counterPage");
const lightbox = document.getElementById("lightbox");
const lbImg = lightbox.querySelector("img");
const lbVid = lightbox.querySelector("video");
const headerEl = document.querySelector("header");
const legendEl = document.getElementById("legend");

let currentFolder = null;
let items = [];               // the 25 items currently rendered (after filtering)
let focusIdx = 0;
let nextLink = null;          // next page link for forward paging from current page
let pathStack = [];
let hideProcessed = false;
let filterMode = "all";
let totalCount = null;        // total # of files in folder (if retrievable)
let pageIndex = 0;            // 0-based
const PAGE_SIZE = 25;

// Page cache so we can go back without refetch
const pages = [];             // pages[i] = array of raw file items for page i (no filter applied)
const pageNextLinks = [];     // pageNextLinks[i] = nextLink after fetching page i

// ===== Layout: make 5x5 fit viewport =====
function layoutGrid() {
  const available = window.innerHeight - headerEl.offsetHeight - legendEl.offsetHeight - 16;
  const min = 300; // safety
  grid.style.height = Math.max(available, min) + "px";
}
window.addEventListener("resize", layoutGrid);

// ===== UI helpers =====
function setStatus(msg){ statusEl.textContent = msg; }
function setFocus(idx){
  const cells = Array.from(grid.children);
  if (!cells.length) return;
  cells.forEach(c => c.classList.remove("focus"));
  focusIdx = Math.max(0, Math.min(cells.length - 1, idx));
  if (cells[focusIdx]) cells[focusIdx].classList.add("focus");
  updateCounters();
}
function updateHideProcessedBtn(){ toggleHideProcessedBtn.textContent = `Hide processed: ${hideProcessed ? 'On' : 'Off'}`; }

function updateCounters() {
  // Absolute position: pageIndex*25 + (focusIdx+1) over totalCount
  const abs = pageIndex * PAGE_SIZE + (items.length ? (focusIdx + 1) : 0);
  const total = (totalCount ?? "—");
  counterMain.textContent = `${abs || 0} / ${total}`;
  const totalPages = totalCount ? Math.ceil(totalCount / PAGE_SIZE) : "—";
  counterPage.textContent = `(page ${pageIndex + 1} of ${totalPages})`;
  prevPageBtn.disabled = pageIndex <= 0;
  nextPageBtn.disabled = !nextLink && !pages[pageIndex + 1];
}

// ===== NAV =====
async function initRoot() {
  setStatus("Loading root…");
  const root = await g(`/me/drive/root?$select=id,name`);
  pathStack = [{ id: root.id, name: root.name || "Root" }];
  await enterCurrentFolder(true);
}

function renderBreadcrumb() {
  breadcrumb.innerHTML = pathStack.map((n,i)=>`<a href="#" data-idx="${i}">${n.name}</a>`).join(" / ");
}

async function listSubfolders(folderId) {
  const data = await g(`/me/drive/items/${folderId}/children?$select=id,name,folder&$orderby=name asc`);
  const folders = data.value.filter(x => x.folder);
  subfolderSelect.innerHTML = `<option value="">— Open subfolder —</option>` + folders.map(f=>`<option value="${f.id}">${f.name}</option>`).join("");
}

async function fetchTotalCount(folderId) {
  try {
    const r = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${folderId}/children?$top=1&$count=true&$filter=file ne null`, {
      headers: {
        Authorization: `Bearer ${await getToken()}`,
        "Consistency-Level": "eventual"
      }
    });
    if (!r.ok) throw new Error(await r.text());
    const j = await r.json();
    if (typeof j["@odata.count"] === "number") {
      totalCount = j["@odata.count"];
    }
  } catch {
    totalCount = null; // unknown; will update as we go
  }
  updateCounters();
}

async function enterCurrentFolder(firstEnter=false) {
  layoutGrid();
  const { id } = pathStack[pathStack.length - 1];
  currentFolder = id;
  renderBreadcrumb();
  await listSubfolders(id);

  // restore state
  const st = loadState(id);
  hideProcessed = !!st.hideProcessed;
  filterMode = st.filterMode || "all";
  filterModeSel.value = filterMode;
  pageIndex = st.pageIndex || 0;
  updateHideProcessedBtn();

  // reset paging state
  items = [];
  grid.innerHTML = "";
  focusIdx = 0;
  nextLink = null;
  pages.length = 0;
  pageNextLinks.length = 0;

  // try to get total count upfront
  await fetchTotalCount(id);

  // fetch first page
  await ensurePageLoaded(0);
  renderPage(0);
}

// ===== DOWNLOAD URL CACHE =====
const dlCache = new Map(); // id -> { url, ts }
const DL_TTL_MS = 2 * 60 * 1000; // short TTL
async function getDownloadUrl(id, { force = false } = {}) {
  const now = Date.now();
  const hit = dlCache.get(id);
  if (!force && hit && (now - hit.ts) < DL_TTL_MS) return hit.url;

  const meta = await g(`/me/drive/items/${id}?$select=@microsoft.graph.downloadUrl`);
  const url = meta["@microsoft.graph.downloadUrl"];
  if (!url) throw new Error("No downloadUrl");
  dlCache.set(id, { url, ts: now });
  return url;
}

// ===== FILTERING / VISIBILITY =====
function isProcessed(id, st) {
  return !!(st.processed[id] || st.soft[id]);
}
function passesFilter(it, st) {
  const mark = st.processed[it.id]; // 'F', 'X', or undefined
  const soft = !!st.soft[it.id];    // hovered processed
  switch (filterMode) {
    case "chosen":     return mark === "F";
    case "declined":   return mark === "X";
    case "unprocessed":return !mark && !soft;
    default:           return true; // all
  }
}
function shouldShowItem(it, st){
  if (!passesFilter(it, st)) return false;
  if (!hideProcessed) return true;
  return !isProcessed(it.id, st); // hides hard/soft processed
}

// ===== RENDER HELPERS =====
function applyMarkStyles(cell, mark) {
  const badge = cell.querySelector(".badge");
  cell.classList.remove("chosen","declined");
  badge.classList.remove("chosen","declined");
  if (mark === "F") { cell.classList.add("chosen");   badge.classList.add("chosen");   badge.textContent = "F"; }
  else if (mark === "X") { cell.classList.add("declined"); badge.classList.add("declined"); badge.textContent = "X"; }
  else { badge.textContent = ""; }
}

function renderItem(it, st) {
  const cell = document.createElement("div");
  cell.className = "cell";
  cell.dataset.id = it.id;

  const isVideo = !!it.video;
  const mediaEl = isVideo ? document.createElement("video") : document.createElement("img");

  if (isVideo) {
    mediaEl.muted = true;
    mediaEl.playsInline = true;
    mediaEl.controls = false;
    mediaEl.preload = "none";
    if (it._thumb) mediaEl.poster = it._thumb;
    const play = document.createElement("div");
    play.style.position = "absolute";
    play.style.inset = "0";
    play.style.display = "flex";
    play.style.alignItems = "center";
    play.style.justifyContent = "center";
    play.style.pointerEvents = "none";
    play.style.fontSize = "42px";
    play.style.color = "white";
    play.style.textShadow = "0 1px 2px rgba(0,0,0,.6)";
    play.textContent = "▶";
    cell.appendChild(play);
  } else {
    mediaEl.loading = "lazy";
    mediaEl.decoding = "async";
    mediaEl.src = it._thumb || "";
  }

  cell.appendChild(mediaEl);

  const badge = document.createElement("div");
  badge.className = "badge";
  cell.appendChild(badge);
  applyMarkStyles(cell, st.processed[it.id]);

  // hover -> mark as soft processed
  cell.addEventListener("mouseenter", () => {
    const st2 = loadState(currentFolder);
    if (!st2.soft[it.id] && !st2.processed[it.id]) {
      st2.soft[it.id] = true;
      saveState(currentFolder, st2);
    }
  });

  cell.addEventListener("click", () => {
    const idx = items.findIndex(x => x.id === it.id);
    if (idx >= 0) setFocus(idx);
  });

  grid.appendChild(cell);
}

// ===== PAGE LOADING & RENDER =====
async function fetchPageFromGraph(urlOrNextLink) {
  const data = urlOrNextLink.startsWith("https://")
    ? await authedFetch(urlOrNextLink)
    : await g(urlOrNextLink);

  const raw = data.value.filter(x => x.file);

  // thumbnails
  const thumbs = await Promise.all(raw.map(async it=>{
    try {
      const t = await g(`/me/drive/items/${it.id}/thumbnails`);
      const set = (t.value && t.value[0]) || {};
      return (set.large || set.medium || set.small || {}).url || null;
    } catch { return null; }
  }));
  raw.forEach((it,i)=>{ it._thumb = thumbs[i]; });
  return { raw, next: data["@odata.nextLink"] || null };
}

async function ensurePageLoaded(idx) {
  if (pages[idx]) return;

  // initial page
  if (idx === 0) {
    const path = `/me/drive/items/${currentFolder}/children?$top=${PAGE_SIZE}&$select=id,name,file,photo,video&$orderby=name asc`;
    setStatus("Loading items…");
    const { raw, next } = await fetchPageFromGraph(path);
    pages[0] = raw;
    pageNextLinks[0] = next;
    nextLink = next;
    if (totalCount == null) totalCount = raw.length; // estimate until we know more
    setStatus(next ? "More available" : "End of folder");
    return;
  }

  // need previous page to know the next-link to fetch from
  if (!pages[idx - 1]) await ensurePageLoaded(idx - 1);

  // if already cached next page (navigated before), we’re done
  if (pages[idx]) return;

  const link = pageNextLinks[idx - 1];
  if (!link) return; // no further pages
  setStatus("Loading items…");
  const { raw, next } = await fetchPageFromGraph(link);
  pages[idx] = raw;
  pageNextLinks[idx] = next;
  nextLink = next;

  // update totalCount if we reached end
  if (!next) {
    // we can infer total if we fetched all pages
    const known = pages.reduce((sum, p) => sum + (p ? p.length : 0), 0);
    totalCount = known;
  }
  setStatus(next ? "More available" : "End of folder");
}

function renderPage(idx) {
  const st = loadState(currentFolder);

  const src = pages[idx] || [];
  // apply filter/hideProcessed per current settings
  const filtered = src.filter(it => shouldShowItem(it, st)).slice(0, PAGE_SIZE);

  grid.innerHTML = "";
  items = [];
  for (const it of filtered) {
    renderItem(it, st);
    items.push(it);
  }

  pageIndex = idx;
  // persist where we are + settings
  saveState(currentFolder, { ...st, hideProcessed, filterMode, pageIndex, cursor: pageNextLinks[idx] || null });

  setFocus(0);
  updateCounters();
}

// next / prev navigation
async function gotoNextPage() {
  if (pages[pageIndex + 1]) {
    renderPage(pageIndex + 1);
    return;
  }
  if (!nextLink) return; // nothing more
  await ensurePageLoaded(pageIndex + 1);
  if (pages[pageIndex + 1]) renderPage(pageIndex + 1);
}
function gotoPrevPage() {
  if (pageIndex === 0) return;
  renderPage(pageIndex - 1);
}

// ===== MARK / LIGHTBOX =====
function currentItem(){ return items[focusIdx]; }

function setMark(it, newMark){
  const st = loadState(currentFolder);
  if (!newMark) delete st.processed[it.id]; else st.processed[it.id] = newMark;
  // soft stays as-is; processed is still true by virtue of hard/soft
  saveState(currentFolder, st);

  const cell = grid.children[focusIdx];
  if (cell) applyMarkStyles(cell, newMark);

  if (hideProcessed && (newMark || st.soft[it.id])){
    cell.remove();
    const idx = items.findIndex(x => x.id === it.id);
    if (idx >= 0) items.splice(idx,1);
    setFocus(Math.min(focusIdx, items.length - 1));
  }
}

function lightboxOpen(){ return lightbox.style.display === "flex"; }
function closeLightbox() { lightbox.style.display = "none"; try { lbVid.pause(); } catch {} lbVid.src = ""; lbImg.src = ""; }

let lbTicket = 0; // avoid races when toggling fast
async function toggleLightbox(){
  if (lightboxOpen()) { closeLightbox(); return; }
  const it = currentItem(); if (!it) return;

  // Fallback for formats browsers can't show inline
  const unsupportedExts = [".heic", ".nef", ".cr2", ".arw", ".orf", ".rw2", ".dng"];
  const ext = it.name ? ("." + it.name.toLowerCase().split(".").pop()) : "";
  if (unsupportedExts.includes(ext) && it._thumb) {
    lbImg.style.display = "block";
    lbVid.style.display = "none";
    lbImg.src = it._thumb;
    lightbox.style.display = "flex";
    setStatus(nextLink ? "More available" : "Ready");
    return;
  }

  const my = ++lbTicket;
  setStatus("Loading preview…");

  async function tryLoad({ forceUrl = false } = {}) {
    const url = await getDownloadUrl(it.id, { force: forceUrl });
    const isVideo = !!it.video;

    lbImg.style.display = isVideo ? "none" : "block";
    lbVid.style.display = isVideo ? "block" : "none";

    lightbox.style.display = "flex";

    return new Promise((resolve, reject) => {
      if (isVideo) {
        lbVid.preload = "metadata"; lbVid.src = url;
        const onLoaded = () => { cleanup(); resolve(); };
        const onError  = () => { cleanup(); reject(new Error("video error")); };
        function cleanup(){ lbVid.removeEventListener("loadedmetadata", onLoaded); lbVid.removeEventListener("error", onError); }
        lbVid.addEventListener("loadedmetadata", onLoaded, { once: true });
        lbVid.addEventListener("error", onError, { once: true });
      } else {
        lbImg.src = url;
        const onLoaded = () => { cleanup(); resolve(); };
        const onError  = () => { cleanup(); reject(new Error("image error")); };
        function cleanup(){ lbImg.removeEventListener("load", onLoaded); lbImg.removeEventListener("error", onError); }
        lbImg.addEventListener("load", onLoaded, { once: true });
        lbImg.addEventListener("error", onError, { once: true });
      }
    });
  }

  try { await tryLoad({ forceUrl:false }); }
  catch (e1) {
    console.warn("[LB] first load failed, retrying with fresh URL", e1);
    try { await tryLoad({ forceUrl:true }); }
    catch (e2) { console.error("[LB] second load failed", e2); closeLightbox(); alert("Could not load media. Try again."); }
  } finally {
    if (my === lbTicket) setStatus(nextLink ? "More available" : "Ready");
  }
}

// ===== BULK =====
async function deleteDeclined(){
  if (!currentFolder) return;
  const st = loadState(currentFolder);
  const ids = Object.entries(st.processed).filter(([,m])=>m==="X").map(([id])=>id);
  if (!ids.length) return alert("Nothing to delete.");
  if (!confirm(`Delete ${ids.length} files from OneDrive? This cannot be undone.`)) return;

  for (let i=0;i<ids.length;i+=20){
    const chunk = ids.slice(i,i+20);
    const token = await getToken();
    const body = { requests: chunk.map((id,idx)=>({ id:`del${i+idx}`, method:"DELETE", url:`/me/drive/items/${id}` })) };
    const r = await fetch("https://graph.microsoft.com/v1.0/$batch", {
      method:"POST", headers:{ "Authorization":`Bearer ${token}`, "Content-Type":"application/json" }, body: JSON.stringify(body)
    });
    if (!r.ok) throw new Error(await r.text());
  }
  ids.forEach(id=>{ delete st.processed[id]; });
  saveState(currentFolder, st);

  // remove from view too (if present)
  for (let i=grid.children.length-1;i>=0;i--){
    const cell = grid.children[i];
    if (ids.includes(cell.dataset.id)){
      cell.remove();
      const idx = items.findIndex(x=>x.id===cell.dataset.id);
      if (idx>=0) items.splice(idx,1);
    }
  }
  setFocus(Math.min(focusIdx, items.length - 1));
  alert("Declined files deleted.");
}

// ===== ZIP HELPERS =====
async function getName(id) {
  try { const meta = await g(`/me/drive/items/${id}`); return meta?.name || id; }
  catch { return id; }
}
async function fetchBytesFor(id) {
  try { const url1 = await getDownloadUrl(id, { force: false }); const r1 = await fetch(url1); if (r1.ok) return new Uint8Array(await r1.arrayBuffer()); } catch {}
  try { const url2 = await getDownloadUrl(id, { force: true  }); const r2 = await fetch(url2); if (r2.ok) return new Uint8Array(await r2.arrayBuffer()); } catch {}
  try { const meta = await g(`/me/drive/items/${id}`); const url3 = meta && meta["@microsoft.graph.downloadUrl"]; if (url3) { const r3 = await fetch(url3); if (r3.ok) return new Uint8Array(await r3.arrayBuffer()); } } catch {}
  try { const token = await getToken(); const r4 = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${id}/content`, { headers: { Authorization: `Bearer ${token}` } }); if (r4.ok) return new Uint8Array(await r4.arrayBuffer()); } catch {}
  throw new Error("No downloadable content");
}
async function buildZipAndSave(entries, label = "chosen") {
  setStatus("Zipping…");
  const fileMap = {}; for (const [name, bytes] of entries) fileMap[name] = bytes;
  const zipped = fflate.zipSync(fileMap, { level: 6 });
  const blob = new Blob([zipped], { type: "application/zip" });
  const suggested = `${label}_${new Date().toISOString().slice(0,19).replace(/[:T]/g, "-")}.zip`;
  if (window.showSaveFilePicker) {
    try {
      const handle = await window.showSaveFilePicker({ suggestedName: suggested, types: [{ description: "ZIP archive", accept: { "application/zip": [".zip"] } }] });
      const writable = await handle.createWritable(); await writable.write(blob); await writable.close();
      setStatus("Saved ZIP.");
      return;
    } catch (e) { console.warn("Save picker failed, falling back to download", e); }
  }
  const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = suggested; document.body.appendChild(a); a.click(); a.remove();
  setStatus("Downloaded ZIP.");
}

// ===== ZIP: chosen (all in folder) =====
async function downloadChosen() {
  if (!currentFolder) return;
  const st = loadState(currentFolder);
  const ids = Object.entries(st.processed).filter(([, m]) => m === "F").map(([id]) => id);
  if (!ids.length) return alert("Nothing chosen.");

  await zipByChunks(ids, "chosen");
}

// ===== ZIP: NOT chosen (all in folder, i.e., everything except F) =====
async function downloadNotChosen() {
  if (!currentFolder) return;
  const st = loadState(currentFolder);

  // Build set of all file IDs in folder. Use cached pages + fetch forward until exhausted.
  setStatus("Gathering items…");
  let allIds = [];
  // include already loaded pages
  for (const p of pages) if (p) allIds.push(...p.map(x => x.id));

  // walk forward to end if needed to include everything
  let cursor = pageNextLinks[pages.length - 1] ?? null;
  while (cursor) {
    const { raw, next } = await fetchPageFromGraph(cursor);
    pages.push(raw); pageNextLinks[pages.length - 1] = next; cursor = next;
    allIds.push(...raw.map(x => x.id));
  }
  // de-dupe
  allIds = Array.from(new Set(allIds));

  // NOT chosen = everything not marked F
  const notChosenIds = allIds.filter(id => st.processed[id] !== "F");
  if (!notChosenIds.length) { alert("Everything here is chosen — nothing to zip."); return; }

  await zipByChunks(notChosenIds, "not_chosen");
}

async function zipByChunks(ids, labelBase) {
  const CHUNK = 150;
  let chunkIndex = 0;
  for (let start = 0; start < ids.length; start += CHUNK) {
    const chunk = ids.slice(start, start + CHUNK);
    setStatus(`Fetching ${chunk.length} files (chunk ${chunkIndex + 1}/${Math.ceil(ids.length/CHUNK)})…`);

    const maxConcurrent = 4;
    let i = 0;
    const entries = [];
    const errors = [];
    async function worker() {
      while (i < chunk.length) {
        const my = i++; const id = chunk[my];
        try {
          const [name, bytes] = await Promise.all([getName(id), fetchBytesFor(id)]);
          entries.push([name, bytes]);
          setStatus(`Fetched ${entries.length}/${chunk.length} (chunk ${chunkIndex + 1})…`);
        } catch (e) {
          console.warn("Fetch failed for", id, e);
          errors.push({ id, error: String(e) });
        }
      }
    }
    await Promise.all(Array.from({ length: Math.min(maxConcurrent, chunk.length) }, worker));
    if (!entries.length) { alert(`Chunk ${chunkIndex + 1}: could not fetch any files.`); continue; }
    await buildZipAndSave(entries, `${labelBase}_${chunkIndex + 1}`);
    if (errors.length) alert(`Chunk ${chunkIndex + 1} finished with ${errors.length} error(s). See console.`);
    chunkIndex++;
  }
  setStatus("Ready");
}

// ===== EVENTS =====
document.getElementById("loginBtn").onclick = login;

upBtn.onclick = () => { if (pathStack.length > 1) { pathStack.pop(); enterCurrentFolder(); } };

breadcrumb.addEventListener("click", (e) => {
  const a = e.target.closest("a[data-idx]"); if (!a) return;
  const idx = parseInt(a.dataset.idx, 10);
  pathStack = pathStack.slice(0, idx + 1);
  enterCurrentFolder();
});

subfolderSelect.onchange = (e) => {
  const id = e.target.value; if (!id) return;
  const name = e.target.selectedOptions[0].textContent;
  pathStack.push({ id, name });
  enterCurrentFolder();
  e.target.value = "";
};

filterModeSel.onchange = () => {
  filterMode = filterModeSel.value;
  renderPage(pageIndex); // re-render current cached page with filter applied
};

toggleHideProcessedBtn.onclick = () => {
  hideProcessed = !hideProcessed; updateHideProcessedBtn();
  renderPage(pageIndex);
};

nextPageBtn.onclick = () => gotoNextPage();
prevPageBtn.onclick = () => gotoPrevPage();

deleteDeclinedBtn.onclick = () => deleteDeclined();
downloadChosenBtn.onclick = () => downloadChosen();
downloadNotChosenBtn.onclick = () => downloadNotChosen();

lightbox.addEventListener("click", () => { closeLightbox(); });

// Global keyboard: works without clicking first. Ignore when typing in inputs/selects.
document.addEventListener("keydown", (e) => {
  const tag = (e.target.tagName || "").toLowerCase();
  if (tag === "input" || tag === "textarea" || tag === "select" || e.isComposing) return;

  const key = e.key;
  if (["ArrowLeft","ArrowRight","ArrowUp","ArrowDown"," ","F","f","X","x","Backspace","s","S","h","H"].includes(key)) e.preventDefault();

  if (key === "ArrowRight") setFocus(focusIdx + 1);
  if (key === "ArrowLeft")  setFocus(focusIdx - 1);
  if (key === "ArrowUp")    setFocus(focusIdx - 5);
  if (key === "ArrowDown")  setFocus(focusIdx + 5);

  if (key === " ")          toggleLightbox();
  if (key === "F" || key === "f") { const it = currentItem(); if (it) setMark(it, "F"); }
  if (key === "X" || key === "x") { const it = currentItem(); if (it) setMark(it, "X"); }
  if (key === "Backspace")        { const it = currentItem(); if (it) setMark(it, null); }
  if (key === "s" || key === "S") gotoNextPage();
  if (key === "h" || key === "H") toggleHideProcessedBtn.click();
});

// Warm session
(function tryWarmAccount(){
  const acc = msalInstance.getAllAccounts()[0];
  if (acc) { account = acc; setStatus("Session found. Click Sign in to continue."); }
})();

// After DOM ready, compute initial layout
layoutGrid();

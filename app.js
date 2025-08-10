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
  return JSON.parse(localStorage.getItem(storeKey(folderId)) || `{"processed":{},"cursor":null,"hideProcessed":false}`);
}
function saveState(folderId, s) { localStorage.setItem(storeKey(folderId), JSON.stringify(s)); }

// ===== DOM =====
const grid = document.getElementById("grid");
const upBtn = document.getElementById("upBtn");
const breadcrumb = document.getElementById("breadcrumb");
const subfolderSelect = document.getElementById("subfolderSelect");
const toggleHideProcessedBtn = document.getElementById("toggleHideProcessed");
const nextPageBtn = document.getElementById("nextPage");
const deleteDeclinedBtn = document.getElementById("deleteDeclined");
const downloadChosenBtn = document.getElementById("downloadChosen");
const statusEl = document.getElementById("status");
const lightbox = document.getElementById("lightbox");
const lbImg = lightbox.querySelector("img");
const lbVid = lightbox.querySelector("video");
const headerEl = document.querySelector("header");
const legendEl = document.getElementById("legend");

let currentFolder = null;
let items = [];
let focusIdx = 0;
let nextLink = null;
let pathStack = [];
let hideProcessed = false;

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
}
function updateHideProcessedBtn(){ toggleHideProcessedBtn.textContent = `Hide processed: ${hideProcessed ? 'On' : 'Off'}`; }

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

async function enterCurrentFolder(firstEnter=false) {
  layoutGrid();
  const { id } = pathStack[pathStack.length - 1];
  currentFolder = id;
  renderBreadcrumb();
  await listSubfolders(id);

  const st = loadState(id);
  hideProcessed = !!st.hideProcessed;
  updateHideProcessedBtn();

  items = [];
  grid.innerHTML = "";
  focusIdx = 0;

  nextLink = null; // fresh load
  await loadNextPage(true);
  // persist cursor for subsequent S presses
  nextLink = st.cursor || nextLink;
}

// ===== DOWNLOAD URL CACHE =====
const dlCache = new Map(); // id -> { url, ts }
const DL_TTL_MS = 2 * 60 * 1000; // short TTL to avoid stale links
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

// ===== LOAD & RENDER =====
function shouldShowItem(it, st){ return hideProcessed ? !st.processed[it.id] : true; }

function applyMarkStyles(cell, mark) {
  const badge = cell.querySelector(".badge");
  cell.classList.remove("chosen","declined");
  badge.classList.remove("chosen","declined");
  if (mark === "F") { cell.classList.add("chosen");   badge.classList.add("chosen");   badge.textContent = "F"; }
  else if (mark === "X") { cell.classList.add("declined"); badge.classList.add("declined"); badge.textContent = "X"; }
  else { badge.textContent = ""; }
}

function renderItem(it, mark) {
  const cell = document.createElement("div");
  cell.className = "cell";
  cell.dataset.id = it.id;

  const isVideo = !!it.video;
  const mediaEl = isVideo ? document.createElement("video") : document.createElement("img");

  if (isVideo) {
    mediaEl.muted = true;
    mediaEl.playsInline = true;
    mediaEl.controls = false;
    mediaEl.preload = "none";               // don't fetch video in grid
    if (it._thumb) mediaEl.poster = it._thumb; // use thumbnail as poster
    // subtle play overlay
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
  applyMarkStyles(cell, mark);

  cell.addEventListener("click", () => {
    const idx = items.findIndex(x => x.id === it.id);
    if (idx >= 0) setFocus(idx);
  });

  grid.appendChild(cell);
}

async function loadNextPage(firstPage=false) {
  if (!currentFolder) return;
  setStatus("Loading items…");
  const st = loadState(currentFolder);

  const url = nextLink
    ? nextLink
    : `https://graph.microsoft.com/v1.0/me/drive/items/${currentFolder}/children` +
      `?$top=25&$select=id,name,file,photo,video,createdDateTime` +
      `&$orderby=name asc`;

  const data = nextLink ? await authedFetch(url) : await g(url.replace("https://graph.microsoft.com/v1.0",""));
  nextLink = data["@odata.nextLink"] || null;

  let page = data.value.filter(x => x.file);

  // thumbnails for all (photos & videos)
  const thumbs = await Promise.all(page.map(async it=>{
    try {
      const t = await g(`/me/drive/items/${it.id}/thumbnails`);
      const set = (t.value && t.value[0]) || {};
      return (set.large || set.medium || set.small || {}).url || null;
    } catch { return null; }
  }));
  page.forEach((it,i)=>{ it._thumb = thumbs[i]; });

  const visible = page.filter(it => shouldShowItem(it, st));

  // REPLACE grid with this page only (max 25)
  grid.innerHTML = "";
  items = [];
  for (const it of visible.slice(0, 25)) {
    renderItem(it, st.processed[it.id]);
    items.push(it);
  }
  setFocus(0);

  saveState(currentFolder, { ...st, cursor: nextLink });
  setStatus(nextLink ? "More available" : "End of folder");
}

// ===== MARK / LIGHTBOX =====
function currentItem(){ return items[focusIdx]; }

function setMark(it, newMark){
  const st = loadState(currentFolder);
  if (!newMark) delete st.processed[it.id]; else st.processed[it.id] = newMark;
  saveState(currentFolder, st);

  const cell = grid.children[focusIdx];
  if (cell) applyMarkStyles(cell, newMark);

  if (hideProcessed && newMark){
    cell.remove();
    const idx = items.findIndex(x => x.id === it.id);
    if (idx >= 0) items.splice(idx,1);
    setFocus(Math.min(focusIdx, items.length - 1));
  }
}

function lightboxOpen(){ return lightbox.style.display === "flex"; }

function closeLightbox() {
  lightbox.style.display = "none";
  try { lbVid.pause(); } catch {}
  lbVid.src = "";
  lbImg.src = "";
}

let lbTicket = 0; // avoid races when toggling fast

async function toggleLightbox(){
  if (lightboxOpen()) { closeLightbox(); return; }

  const it = currentItem(); 
  if (!it) return;

  // Fallback for formats browsers can't show inline
  const unsupportedExts = [".heic", ".nef", ".cr2", ".arw", ".orf", ".rw2", ".dng"];
  const ext = it.name ? ("." + it.name.toLowerCase().split(".").pop()) : "";
  if (unsupportedExts.includes(ext) && it._thumb) {
    lbImg.style.display = "block";
    lbVid.style.display = "none";
    lbImg.src = it._thumb;     // show large JPEG thumbnail instead
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
        lbVid.preload = "metadata";
        lbVid.src = url;

        const onLoaded = () => { cleanup(); resolve(); };
        const onError  = () => { cleanup(); reject(new Error("video error")); };

        function cleanup() {
          lbVid.removeEventListener("loadedmetadata", onLoaded);
          lbVid.removeEventListener("error", onError);
        }
        lbVid.addEventListener("loadedmetadata", onLoaded, { once: true });
        lbVid.addEventListener("error", onError, { once: true });
      } else {
        lbImg.src = url;

        const onLoaded = () => { cleanup(); resolve(); };
        const onError  = () => { cleanup(); reject(new Error("image error")); };

        function cleanup() {
          lbImg.removeEventListener("load", onLoaded);
          lbImg.removeEventListener("error", onError);
        }
        lbImg.addEventListener("load", onLoaded, { once: true });
        lbImg.addEventListener("error", onError, { once: true });
      }
    });
  }

  try {
    await tryLoad({ forceUrl: false });   // first attempt (cached URL ok)
  } catch (e1) {
    console.warn("[LB] first load failed, retrying with fresh URL", e1);
    try {
      await tryLoad({ forceUrl: true });  // second attempt (force refresh URL)
    } catch (e2) {
      console.error("[LB] second load failed", e2);
      closeLightbox();
      alert("Could not load media (network or URL expired). Try again.");
    }
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

async function downloadChosen() {
  if (!currentFolder) return;
  const st = loadState(currentFolder);
  const chosen = Object.entries(st.processed).filter(([, m]) => m === "F").map(([id]) => id);
  if (!chosen.length) return alert("Nothing chosen.");

  setStatus(`Fetching ${chosen.length} files…`);

  async function fetchBytes(url, attempt = 0) {
    const r = await fetch(url);
    if (r.status === 429 && attempt < 1) {
      const ra = parseInt(r.headers.get("Retry-After") || "2", 10) * 1000;
      await new Promise(res => setTimeout(res, isFinite(ra) ? ra : 2000));
      return fetchBytes(url, attempt + 1);
    }
    if (!r.ok) throw new Error(`Download failed: ${r.status}`);
    return new Uint8Array(await r.arrayBuffer());
  }

  const maxConcurrent = 4;
  let idx = 0;
  const entries = [];
  const errors = [];

  async function worker() {
    while (idx < chosen.length) {
      const i = idx++;
      try {
        const meta = await g(`/me/drive/items/${chosen[i]}?$select=name,@microsoft.graph.downloadUrl`);
        const name = meta?.name || `file_${i}`;
        const url = meta?.["@microsoft.graph.downloadUrl"];
        if (!url) throw new Error("No downloadUrl");
        const bytes = await fetchBytes(url);
        entries.push([name, bytes]);
        setStatus(`Fetched ${entries.length}/${chosen.length}…`);
      } catch (e) {
        console.error("Fetch error", e);
        errors.push({ id: chosen[i], error: e.toString() });
      }
    }
  }

  await Promise.all(Array.from({ length: Math.min(maxConcurrent, chosen.length) }, worker));

  if (!entries.length) {
    alert("Could not fetch any files.");
    setStatus("Ready");
    return;
  }

  setStatus("Zipping…");

  const fileMap = {};
  for (const [name, bytes] of entries) fileMap[name] = bytes;
  const zipped = fflate.zipSync(fileMap, { level: 6 });
  const blob = new Blob([zipped], { type: "application/zip" });

  const suggested = `chosen_${new Date().toISOString().slice(0,19).replace(/[:T]/g, "-")}.zip`;
  if (window.showSaveFilePicker) {
    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: suggested,
        types: [{ description: "ZIP archive", accept: { "application/zip": [".zip"] } }],
      });
      const writable = await handle.createWritable();
      await writable.write(blob);
      await writable.close();
      setStatus("Saved ZIP.");
    } catch (e) {
      console.warn("Save picker failed, falling back to download", e);
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = suggested;
      document.body.appendChild(a);
      a.click(); a.remove();
      setStatus("Downloaded ZIP.");
    }
  } else {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = suggested;
    document.body.appendChild(a);
    a.click(); a.remove();
    setStatus("Downloaded ZIP.");
  }

  if (errors.length) {
    console.warn("Some files failed to fetch:", errors);
    alert(`Finished with ${errors.length} fetch error(s). Check console for details.`);
  }
}

// ===== EVENTS =====
document.getElementById("loginBtn").onclick = login;

upBtn.onclick = () => {
  if (pathStack.length > 1) { pathStack.pop(); enterCurrentFolder(); }
};

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

toggleHideProcessedBtn.onclick = () => {
  hideProcessed = !hideProcessed; updateHideProcessedBtn();
  const st = loadState(currentFolder); saveState(currentFolder, { ...st, hideProcessed });
  // refresh current page according to filter (keep just these 25)
  const keep = items.slice();
  grid.innerHTML = ""; items = [];
  const st2 = loadState(currentFolder);
  for (const it of keep) if (shouldShowItem(it, st2)) { renderItem(it, st2.processed[it.id]); items.push(it); }
  setFocus(Math.min(focusIdx, items.length - 1));
};

nextPageBtn.onclick = () => loadNextPage(false);
deleteDeclinedBtn.onclick = () => deleteDeclined();
downloadChosenBtn.onclick = () => downloadChosen();
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
  if (key === "s" || key === "S") nextPageBtn.click();
  if (key === "h" || key === "H") toggleHideProcessedBtn.click();
});

// Warm session
(function tryWarmAccount(){
  const acc = msalInstance.getAllAccounts()[0];
  if (acc) { account = acc; setStatus("Session found. Click Sign in to continue."); }
})();

// After DOM ready, compute initial layout
layoutGrid();

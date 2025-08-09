// ---- CONFIG ----
const msalConfig = {
  auth: {
    clientId: "YOUR_APP_CLIENT_ID",
    authority: "https://login.microsoftonline.com/consumers", // OneDrive Personal; for Work/School use "common"
    redirectUri: location.origin
  },
  cache: { cacheLocation: "localStorage" }
};
const graphScopes = ["Files.ReadWrite", "offline_access"]; // need read+write + refresh

// ---- MSAL ----
const msalInstance = new msal.PublicClientApplication(msalConfig);
let account = null;
async function login() {
  try {
    const res = await msalInstance.loginPopup({ scopes: graphScopes });
    account = res.account;
  } catch (e) { console.error(e); setStatus("Login failed"); }
}
async function getToken() {
  if (!account) account = msalInstance.getAllAccounts()[0];
  const req = { account, scopes: graphScopes };
  try { return (await msalInstance.acquireTokenSilent(req)).accessToken; }
  catch { return (await msalInstance.acquireTokenPopup(req)).accessToken; }
}
async function g(path, opts={}) {
  const token = await getToken();
  const r = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    ...opts,
    headers: { "Authorization": `Bearer ${token}`, "Content-Type":"application/json", ...(opts.headers||{}) }
  });
  if (!r.ok) throw new Error(await r.text());
  return await r.json().catch(() => ({}));
}

// ---- STATE (per-folder) ----
const storeKey = (folderId)=>`ps:${folderId}`;
function loadState(folderId) { return JSON.parse(localStorage.getItem(storeKey(folderId)) || `{"processed":{},"cursor":null}`); }
function saveState(folderId, s) { localStorage.setItem(storeKey(folderId), JSON.stringify(s)); }

// ---- UI refs ----
const grid = document.getElementById("grid");
const folderSelect = document.getElementById("folderSelect");
const statusEl = document.getElementById("status");
const lightbox = document.getElementById("lightbox");
const lbImg = lightbox.querySelector("img");
const lbVid = lightbox.querySelector("video");

let currentFolder = null;
let mode = "unprocessed"; // or "all"
let items = [];       // current page
let focusIdx = 0;     // 0..24
let nextLink = null;  // pagination

function setStatus(msg){ statusEl.textContent = msg; }

// ---- 1) load top-level folders for selection ----
async function loadFolders() {
  setStatus("Loading folders...");
  // List root children and filter to folders
  const data = await g(`/me/drive/root/children?$select=id,name,folder`);
  const folders = data.value.filter(x => x.folder);
  folderSelect.innerHTML = folders.map(f=>`<option value="${f.id}">${f.name}</option>`).join("");
  setStatus("");
  if (folders.length) selectFolder(folders[0].id);
}

// ---- 2) selecting a folder ----
async function selectFolder(folderId){
  currentFolder = folderId;
  items = []; focusIdx = 0; grid.innerHTML = "";
  const st = loadState(folderId);
  nextLink = st.cursor; // resume if present
  await loadNextPage(true);
}

// ---- 3) page of 25 items with thumbnails ----
async function loadNextPage(first=false){
  setStatus("Loading items...");
  const url = nextLink
    ? nextLink
    : `https://graph.microsoft.com/v1.0/me/drive/items/${currentFolder}/children?$top=25&$select=id,name,file,photo,video,@microsoft.graph.downloadUrl`;
  const data = nextLink ? await authedFetch(nextLink) : await g(url.replace("https://graph.microsoft.com/v1.0",""));
  nextLink = data["@odata.nextLink"] || null;

  // if unprocessed mode, skip items you've already tagged; otherwise include all
  const st = loadState(currentFolder);
  let page = data.value.filter(x => x.file && (mode==="all" || !st.processed[x.id]));

  // fetch thumbnails for each item (best-effort)
  const thumbs = await Promise.all(page.map(async it=>{
    try {
      const t = await g(`/me/drive/items/${it.id}/thumbnails`);
      const set = (t.value && t.value[0]) || {};
      return { id: it.id, url: (set.large||set.medium||set.small||{}).url || null };
    } catch { return { id: it.id, url: null }; }
  }));

  const thumbMap = new Map(thumbs.map(t=>[t.id, t.url]));
  page.forEach(it=>{
    const cell = document.createElement("div");
    cell.className = "cell";
    cell.dataset.id = it.id;
    const isVideo = !!it.video;
    const mediaUrl = thumbMap.get(it.id) || it["@microsoft.graph.downloadUrl"];
    if (isVideo) {
      const v = document.createElement("video");
      v.src = mediaUrl; v.muted = true; v.playsInline = true;
      cell.appendChild(v);
    } else {
      const img = document.createElement("img");
      img.loading = "lazy"; img.decoding = "async"; img.src = mediaUrl;
      cell.appendChild(img);
    }
    const badge = document.createElement("div"); badge.className="badge";
    const mark = st.processed[it.id]; if (mark) badge.textContent = mark;
    cell.appendChild(badge);

    grid.appendChild(cell);
  });

  items = [...items, ...page];
  setFocus(0, first);
  setStatus(nextLink ? "More available" : "End of folder");
  // persist cursor for resume
  saveState(currentFolder, {...st, cursor: nextLink});
}

async function authedFetch(fullUrl){
  const token = await getToken();
  const r = await fetch(fullUrl, { headers: { Authorization:`Bearer ${token}` }});
  if (!r.ok) throw new Error(await r.text());
  return r.json();
}

// ---- 4) keyboard navigation + actions ----
function setFocus(delta=0, reset=false){
  const cells = Array.from(grid.children);
  cells.forEach(c=>c.classList.remove("focus"));
  if (reset) focusIdx = 0;
  else focusIdx = Math.max(0, Math.min(cells.length-1, focusIdx+delta));
  if (cells[focusIdx]) cells[focusIdx].classList.add("focus");
}
function currentItem(){ return items[focusIdx]; }

function tagCurrent(mark){ // 'F' chosen / 'X' declined
  const it = currentItem(); if (!it) return;
  const st = loadState(currentFolder);
  st.processed[it.id] = mark;
  saveState(currentFolder, st);
  // update badge
  const cell = grid.children[focusIdx];
  const badge = cell.querySelector(".badge"); badge.textContent = mark;
  // in unprocessed mode, hide and move on
  if (mode==="unprocessed") {
    cell.remove(); items.splice(focusIdx,1);
    setFocus(0);
  }
}

function openLightbox(){
  const it = currentItem(); if (!it) return;
  const isVideo = !!it.video;
  lbImg.style.display = isVideo ? "none":"block";
  lbVid.style.display = isVideo ? "block":"none";
  (isVideo ? lbVid : lbImg).src = it["@microsoft.graph.downloadUrl"];
  lightbox.style.display = "flex";
}
function closeLightbox(){ lightbox.style.display = "none"; lbVid.pause(); }

grid.addEventListener("keydown",(e)=>{
  if (e.key==="ArrowRight") { e.preventDefault(); setFocus(+1); }
  if (e.key==="ArrowLeft")  { e.preventDefault(); setFocus(-1); }
  if (e.key==="ArrowUp")    { e.preventDefault(); setFocus(-5); }
  if (e.key==="ArrowDown")  { e.preventDefault(); setFocus(+5); }
  if (e.key===" ")          { e.preventDefault(); openLightbox(); }
  if (e.key.toLowerCase()==="x"){ tagCurrent("X"); }
  if (e.key.toLowerCase()==="f"){ tagCurrent("F"); }
  if (e.key.toLowerCase()==="s"){ loadNextPage(); }
});
grid.addEventListener("click", (e)=>{ grid.focus(); });
lightbox.addEventListener("click", closeLightbox);

// ---- 5) destructive ops ----
async function deleteDeclined(){
  const st = loadState(currentFolder);
  const ids = Object.entries(st.processed).filter(([,m])=>m==="X").map(([id])=>id);
  if (!ids.length) return alert("Nothing to delete.");
  if (!confirm(`Delete ${ids.length} files from OneDrive? This cannot be undone.`)) return;

  // batch 20 deletes per Graph limits
  for (let i=0; i<ids.length; i+=20){
    const chunk = ids.slice(i,i+20);
    const token = await getToken();
    const body = {
      requests: chunk.map((id, idx)=>({
        id: `del${i+idx}`,
        method: "DELETE",
        url: `/me/drive/items/${id}`
      }))
    };
    const r = await fetch("https://graph.microsoft.com/v1.0/$batch", {
      method: "POST",
      headers: { "Authorization":`Bearer ${token}`, "Content-Type":"application/json" },
      body: JSON.stringify(body)
    });
    if (!r.ok) throw new Error(await r.text());
  }
  // clear local marks for deleted
  ids.forEach(id=>{ delete st.processed[id]; });
  saveState(currentFolder, st);
  alert("Declined files deleted.");
}

async function downloadChosen(){
  const st = loadState(currentFolder);
  const chosen = Object.entries(st.processed).filter(([,m])=>m==="F").map(([id])=>id);
  if (!chosen.length) return alert("Nothing chosen.");
  // fetch fresh metadata to get downloadUrl (short-lived)
  for (const id of chosen) {
    const it = await g(`/me/drive/items/${id}?$select=@microsoft.graph.downloadUrl,name`);
    const a = document.createElement("a");
    a.href = it["@microsoft.graph.downloadUrl"];
    a.download = it.name;
    a.style.display="none";
    document.body.appendChild(a);
    a.click(); a.remove();
  }
}

// ---- wiring ----
document.getElementById("loginBtn").onclick = async ()=>{
  await login(); await loadFolders();
};
folderSelect.onchange = (e)=> selectFolder(e.target.value);
document.getElementById("nextPage").onclick = ()=> loadNextPage();
document.getElementById("modeUnprocessed").onclick = ()=>{ mode="unprocessed"; selectFolder(currentFolder); };
document.getElementById("modeAll").onclick = ()=>{ mode="all"; selectFolder(currentFolder); };
document.getElementById("deleteDeclined").onclick = deleteDeclined;
document.getElementById("downloadChosen").onclick = downloadChosen;

// focus grid for keyboard from the start
grid.tabIndex = 0; grid.focus();

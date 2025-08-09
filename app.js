// ===== CONFIG =====
const msalConfig = {
  auth: {
    clientId: "18eebe37-a762-4148-b425-4ee4d79674bf",
    authority: "https://login.microsoftonline.com/consumers", // OneDrive Personal; for Work/School use "common"
    redirectUri: location.origin
  },
  cache: { cacheLocation: "localStorage" }
};
const graphScopes = ["Files.ReadWrite", "offline_access"]; // keep minimal

// ===== MSAL / GRAPH HELPERS =====
const msalInstance = new msal.PublicClientApplication(msalConfig);
let account = null;

async function login() {
  try {
    const res = await msalInstance.loginPopup({ scopes: graphScopes });
    account = res.account;
    setStatus("Signed in");
  } catch (e) {
    console.error(e);
    setStatus("Login failed");
  }
}

async function getToken() {
  if (!account) account = msalInstance.getAllAccounts()[0];
  const req = { account, scopes: graphScopes };
  try {
    return (await msalInstance.acquireTokenSilent(req)).accessToken;
  } catch {
    return (await msalInstance.acquireTokenPopup(req)).accessToken;
  }
}

async function g(path, opts = {}) {
  const token = await getToken();
  const r = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    ...opts,
    headers: {
      "Authorization": `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(opts.headers || {})
    }
  });
  if (!r.ok) throw new Error(await r.text());
  try { return await r.json(); } catch { return {}; }
}

async function authedFetch(fullUrl) {
  const token = await getToken();
  const r = await fetch(fullUrl, { headers: { Authorization: `Bearer ${token}` } });
  if (!r.ok) throw new Error(await r.text());
  return r.json();
}

// ===== STATE (persisted per folder) =====
const storeKey = (folderId) => `ps:${folderId}`;
function loadState(folderId) {
  return JSON.parse(localStorage.getItem(storeKey(folderId)) || `{"processed":{},"cursor":null,"hideProcessed":false}`);
}
function saveState(folderId, s) {
  localStorage.setItem(storeKey(folderId), JSON.stringify(s));
}

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

let currentFolder = null;
let items = [];         // currently rendered items (can span multiple pages loaded so far)
let focusIdx = 0;
let nextLink = null;    // pagination cursor for Graph
let pathStack = [];     // [{id,name}, ...]
let hideProcessed = false;

// ===== UI helpers =====
function setStatus(msg) { statusEl.textContent = msg; }
function setFocus(idx) {
  const cells = Array.from(grid.children);
  cells.forEach(c => c.classList.remove("focus"));
  focusIdx = Math.max(0, Math.min(cells.length - 1, idx));
  if (cells[focusIdx]) cells[focusIdx].classList.add("focus");
}

function updateHideProcessedBtn() {
  toggleHideProcessedBtn.textContent = `Hide processed: ${hideProcessed ? 'On' : 'Off'}`;
}

// ===== NAV =====
async function initRoot() {
  setStatus("Loading root…");
  const root = await g(`/me/drive/root?$select=id,name`);
  pathStack = [{ id: root.id, name: root.name || "Root" }];
  await enterCurrentFolder(true);
}

function renderBreadcrumb() {
  breadcrumb.innerHTML = pathStack
    .map((node, idx) => `<a href="#" data-idx="${idx}">${node.name}</a>`)
    .join(" / ");
}

async function listSubfolders(folderId) {
  const data = await g(`/me/drive/items/${folderId}/children?$select=id,name,folder`);
  const folders = data.value.filter(x => x.folder);
  subfolderSelect.innerHTML = `<option value="">— Open subfolder —</option>` +
    folders.map(f => `<option value="${f.id}">${f.name}</option>`).join("");
}

async function enterCurrentFolder(firstEnter = false) {
  const { id } = pathStack[pathStack.length - 1];
  currentFolder = id;
  renderBreadcrumb();
  await listSubfolders(id);

  // restore state
  const st = loadState(id);
  hideProcessed = !!st.hideProcessed;
  updateHideProcessedBtn();

  // reset current view
  items = [];
  grid.innerHTML = "";
  focusIdx = 0;

  // restore pagination cursor, but don't auto-load; we load first page always so you see something
  nextLink = null; // reset; we will load first page and then, if st.cursor exists, we'll continue S from there
  await loadNextPage(true);

  // if the saved cursor exists, keep it for future S presses (we don't jump forward automatically)
  nextLink = st.cursor || nextLink;
}

// ===== LOADING & RENDER =====
function shouldShowItem(it, st) {
  if (!hideProcessed) return true;
  const mark = st.processed[it.id];
  return !mark; // hide F/X if toggle is on
}

function renderItem(it, mark) {
  const cell = document.createElement("div");
  cell.className = "cell";
  cell.dataset.id = it.id;
  const isVideo = !!it.video;

  const mediaEl = isVideo ? document.createElement("video") : document.createElement("img");
  if (isVideo) {
    mediaEl.muted = true; mediaEl.playsInline = true; mediaEl.controls = false;
  } else {
    mediaEl.loading = "lazy"; mediaEl.decoding = "async";
  }
  mediaEl.src = it._thumb || it["@microsoft.graph.downloadUrl"]; // thumbnail or fallback
  cell.appendChild(mediaEl);

  const badge = document.createElement("div");
  badge.className = "badge";
  badge.textContent = mark || "";
  cell.appendChild(badge);

  cell.addEventListener("click", () => {
    const idx = items.findIndex(x => x.id === it.id);
    if (idx >= 0) setFocus(idx);
  });

  grid.appendChild(cell);
}

async function loadNextPage(firstPage = false) {
  if (!currentFolder) return;
  setStatus("Loading items…");

  const st = loadState(currentFolder);

  // choose URL: from cursor or fresh page
  const url = nextLink
  ? nextLink
  : `https://graph.microsoft.com/v1.0/me/drive/items/${currentFolder}/children` +
    `?$top=25&$select=id,name,file,photo,video,createdDateTime,@microsoft.graph.downloadUrl` +
    `&$orderby=createdDateTime desc`;

  const data = nextLink ? await authedFetch(url) : await g(url.replace("https://graph.microsoft.com/v1.0", ""));
  // update nextLink
  nextLink = data["@odata.nextLink"] || null;

  // pick only files
  let page = data.value.filter(x => x.file);

  // obtain thumbnails (best-effort)
  const thumbs = await Promise.all(page.map(async it => {
    try {
      const t = await g(`/me/drive/items/${it.id}/thumbnails`);
      const set = (t.value && t.value[0]) || {};
      return (set.large || set.medium || set.small || {}).url || null;
    } catch { return null; }
  }));
  page.forEach((it, i) => { it._thumb = thumbs[i]; });

  // filter according to hideProcessed
  const visible = page.filter(it => shouldShowItem(it, st));

  // append to items and render
  items = [...items, ...visible];
  for (const it of visible) {
    renderItem(it, st.processed[it.id]);
  }

  // persist cursor after successful fetch
  saveState(currentFolder, { ...st, cursor: nextLink });

  // focus first cell on initial load
  if (firstPage) setFocus(0);

  setStatus(nextLink ? "More available" : "End of folder");
}

// ===== MARKING / TOGGLING =====
function currentItem() { return items[focusIdx]; }

function setMark(it, newMark) {
  const st = loadState(currentFolder);
  if (!newMark) {
    delete st.processed[it.id];
  } else {
    st.processed[it.id] = newMark; // 'F' or 'X'
  }
  saveState(currentFolder, st);

  // update badge in-place
  const cell = grid.children[focusIdx];
  if (!cell) return;
  const badge = cell.querySelector(".badge");
  badge.textContent = newMark || "";

  // if hideProcessed is ON and we just applied a mark, hide the item visually
  if (hideProcessed) {
    if (newMark) {
      // remove from DOM + items
      cell.remove();
      const idx = items.findIndex(x => x.id === it.id);
      if (idx >= 0) items.splice(idx, 1);
      setFocus(Math.min(focusIdx, items.length - 1));
    }
  }
}

// ===== LIGHTBOX =====
async function openLightbox() {
  const it = currentItem(); if (!it) return;
  try {
    // Fetch a fresh, short-lived download URL each time you open
    const fresh = await g(`/me/drive/items/${it.id}?$select=@microsoft.graph.downloadUrl,video,photo,name`);
    const url = fresh["@microsoft.graph.downloadUrl"];

    const isVideo = !!it.video; // original metadata is fine to decide media type
    lbImg.style.display = isVideo ? "none" : "block";
    lbVid.style.display = isVideo ? "block" : "none";

    if (isVideo) {
      lbVid.src = url;
      lbVid.currentTime = 0;
      lbVid.play().catch(()=>{ /* ignore autoplay block */ });
    } else {
      lbImg.src = url;
    }
    lightbox.style.display = "flex";
  } catch (e) {
    console.error(e);
    alert("Could not load full media. Try again.");
  }
}
function closeLightbox() { lightbox.style.display = "none"; lbVid.pause(); }

// ===== BULK OPS =====
async function deleteDeclined() {
  if (!currentFolder) return;
  const st = loadState(currentFolder);
  const ids = Object.entries(st.processed).filter(([, m]) => m === "X").map(([id]) => id);
  if (!ids.length) return alert("Nothing to delete.");
  if (!confirm(`Delete ${ids.length} files from OneDrive? This cannot be undone.`)) return;

  for (let i = 0; i < ids.length; i += 20) {
    const chunk = ids.slice(i, i + 20);
    const token = await getToken();
    const body = {
      requests: chunk.map((id, idx) => ({
        id: `del${i + idx}`,
        method: "DELETE",
        url: `/me/drive/items/${id}`
      }))
    };
    const r = await fetch("https://graph.microsoft.com/v1.0/$batch", {
      method: "POST",
      headers: { "Authorization": `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify(body)
    });
    if (!r.ok) throw new Error(await r.text());
  }
  // purge local marks and remove cells corresponding to deleted items
  ids.forEach(id => { delete st.processed[id]; });
  saveState(currentFolder, st);

  // also remove deleted items from current view
  for (let i = grid.children.length - 1; i >= 0; i--) {
    const cell = grid.children[i];
    if (ids.includes(cell.dataset.id)) {
      cell.remove();
      const idx = items.findIndex(x => x.id === cell.dataset.id);
      if (idx >= 0) items.splice(idx, 1);
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
  for (const id of chosen) {
    const it = await g(`/me/drive/items/${id}?$select=@microsoft.graph.downloadUrl,name`);
    const a = document.createElement("a");
    a.href = it["@microsoft.graph.downloadUrl"];
    a.download = it.name;
    a.target = "_blank"; // let browser handle
    a.style.display = "none";
    document.body.appendChild(a);
    a.click(); a.remove();
  }
}

// ===== EVENT WIRING =====
document.getElementById("loginBtn").onclick = async () => {
  await login();
  await initRoot();
};

upBtn.onclick = () => {
  if (pathStack.length > 1) {
    pathStack.pop();
    enterCurrentFolder();
  }
};

breadcrumb.addEventListener("click", (e) => {
  const a = e.target.closest("a[data-idx]");
  if (!a) return;
  const idx = parseInt(a.dataset.idx, 10);
  pathStack = pathStack.slice(0, idx + 1);
  enterCurrentFolder();
});

subfolderSelect.onchange = (e) => {
  const id = e.target.value;
  if (!id) return;
  const opt = e.target.selectedOptions[0];
  const name = opt.textContent;
  pathStack.push({ id, name });
  enterCurrentFolder();
  e.target.value = "";
};

toggleHideProcessedBtn.onclick = () => {
  hideProcessed = !hideProcessed;
  updateHideProcessedBtn();
  // save preference
  const st = loadState(currentFolder);
  saveState(currentFolder, { ...st, hideProcessed });

  // re-render current items according to filter (simple refresh)
  const keep = items; // remember selection to refocus later
  grid.innerHTML = "";
  items = [];
  const st2 = loadState(currentFolder);
  for (const it of keep) {
    if (shouldShowItem(it, st2)) {
      renderItem(it, st2.processed[it.id]);
      items.push(it);
    }
  }
  setFocus(Math.min(focusIdx, items.length - 1));
};

nextPageBtn.onclick = () => loadNextPage(false);
deleteDeclinedBtn.onclick = () => deleteDeclined();
downloadChosenBtn.onclick = () => downloadChosen();

grid.addEventListener("keydown", (e) => {
  const key = e.key;
  if (["ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown", " ", "F", "f", "X", "x", "Backspace", "s", "S", "h", "H"].includes(key)) {
    e.preventDefault();
  }
  if (key === "ArrowRight") setFocus(focusIdx + 1);
  if (key === "ArrowLeft") setFocus(focusIdx - 1);
  if (key === "ArrowUp") setFocus(focusIdx - 5);
  if (key === "ArrowDown") setFocus(focusIdx + 5);
  if (key === " ") openLightbox();
  if (key === "F" || key === "f") {
    const it = currentItem(); if (it) setMark(it, "F");
  }
  if (key === "X" || key === "x") {
    const it = currentItem(); if (it) setMark(it, "X");
  }
  if (key === "Backspace") {
    const it = currentItem(); if (it) setMark(it, null);
  }
  if (key === "s" || key === "S") loadNextPage(false);
  if (key === "h" || key === "H") toggleHideProcessedBtn.click();
});

lightbox.addEventListener("click", closeLightbox);

// start with grid focus so arrows work immediately
grid.tabIndex = 0;
grid.addEventListener("click", () => grid.focus());

// try to reuse existing session silently (no auto-load root until user clicks Sign in)
(function tryWarmAccount() {
  const acc = msalInstance.getAllAccounts()[0];
  if (acc) { account = acc; setStatus("Session found. Click Sign in to continue."); }
})();

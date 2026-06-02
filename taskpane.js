/* ═══════════════════════════════════════════════════════════════
   CONFIG
═══════════════════════════════════════════════════════════════ */
const CONFIG = {
  // SharePoint site root
  siteUrl: "https://covalenca.sharepoint.com/sites/InnovationMarketing",

  // Exact internal name of the SharePoint List (template catalog)
  listName: "EmailTemplatesCatalog",

  // Cache TTL in ms (5 minutes)
  cacheTtl: 5 * 60 * 1000,
};

/* ═══════════════════════════════════════════════════════════════
   STATE
═══════════════════════════════════════════════════════════════ */
let allTemplates    = [];
let filteredList    = [];
let activeCategory  = "All";
let searchQuery     = "";
let cacheTimestamp  = 0;
let selectedTemplate = null;
let accessToken     = null;   // cached AAD token

/* ═══════════════════════════════════════════════════════════════
   OFFICE INIT
═══════════════════════════════════════════════════════════════ */
Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    init();
  }
});

function init() {
  bindEvents();
  loadTemplates();
}

/* ═══════════════════════════════════════════════════════════════
   AUTHENTICATION — Office SSO → AAD token for SharePoint
   
   The add-in requests an access token from Office using the
   user's existing M365 session. No login prompt appears.
   The token is then sent as a Bearer header to SharePoint REST.
═══════════════════════════════════════════════════════════════ */
async function getToken() {
  if (accessToken) return accessToken;

  return new Promise((resolve, reject) => {
    Office.context.auth.getAccessTokenAsync(
      { allowSignInPrompt: false, allowConsentPrompt: false, forMSGraphAccess: false },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          accessToken = result.value;
          resolve(accessToken);
        } else {
          reject(new Error(result.error?.message || "Auth failed"));
        }
      }
    );
  });
}

/* ═══════════════════════════════════════════════════════════════
   DATA FETCHING — SharePoint REST API with Bearer token
═══════════════════════════════════════════════════════════════ */
async function loadTemplates() {
  if (allTemplates.length > 0 && Date.now() - cacheTimestamp < CONFIG.cacheTtl) {
    renderTemplates();
    return;
  }

  showState("loading");

  try {
    let token;
    try {
      token = await getToken();
    } catch (authErr) {
      // SSO not configured yet — fall back to same-origin fetch
      // (works when taskpane is hosted on SharePoint itself)
      token = null;
    }

    const endpoint =
      `${CONFIG.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(CONFIG.listName)}')/items` +
      `?$select=Title,Category,Description,TemplateFileUrl,IsActive` +
      `&$filter=IsActive eq 1` +
      `&$orderby=Category,Title` +
      `&$top=500`;

    const headers = {
      "Accept": "application/json;odata=nometadata",
    };
    if (token) {
      headers["Authorization"] = `Bearer ${token}`;
    }

    const response = await fetch(endpoint, {
      headers,
      credentials: token ? "omit" : "include",
    });

    if (response.status === 401) {
      // Token expired or SSO not set up — show a sign-in prompt
      showState("error", "Sign-in required. Please close and reopen the panel.");
      accessToken = null;
      return;
    }

    if (!response.ok) {
      throw new Error(`SharePoint returned ${response.status}: ${response.statusText}`);
    }

    const data = await response.json();
    allTemplates   = data.value || [];
    cacheTimestamp = Date.now();

    buildCategoryPills();
    applyFilters();
    spinRefresh(false);

  } catch (err) {
    console.error("[TemplateAddin] Failed to load templates:", err);
    showState("error", err.message || "Unknown error.");
    spinRefresh(false);
  }
}

/* ═══════════════════════════════════════════════════════════════
   FETCH TEMPLATE HTML FILE
   Uses the same token so cross-origin requests to SharePoint work.
═══════════════════════════════════════════════════════════════ */
async function fetchTemplateHtml(fileUrl) {
  const headers = {};
  if (accessToken) {
    headers["Authorization"] = `Bearer ${accessToken}`;
  }

  const response = await fetch(fileUrl, {
    headers,
    credentials: accessToken ? "omit" : "include",
  });

  if (!response.ok) {
    throw new Error(`Could not fetch template file (${response.status}).`);
  }

  return await response.text();
}

/* ═══════════════════════════════════════════════════════════════
   EVENT BINDINGS
═══════════════════════════════════════════════════════════════ */
function bindEvents() {
  document.getElementById("refreshBtn").addEventListener("click", () => {
    cacheTimestamp = 0;
    accessToken    = null;
    loadTemplates();
    spinRefresh(true);
  });

  document.getElementById("searchInput").addEventListener("input", function () {
    searchQuery = this.value.trim().toLowerCase();
    applyFilters();
  });

  document.getElementById("retryBtn").addEventListener("click", () => loadTemplates());

  document.getElementById("previewClose").addEventListener("click",  closePreview);
  document.getElementById("previewCancel").addEventListener("click", closePreview);
  document.getElementById("previewInsert").addEventListener("click", insertTemplate);

  document.getElementById("previewOverlay").addEventListener("click", function (e) {
    if (e.target === this) closePreview();
  });
}

/* ═══════════════════════════════════════════════════════════════
   CATEGORY PILLS
═══════════════════════════════════════════════════════════════ */
function buildCategoryPills() {
  const bar = document.getElementById("categoryBar");
  bar.querySelectorAll(".cat-pill:not([data-cat='All'])").forEach(el => el.remove());

  const categories = [...new Set(allTemplates.map(t => t.Category).filter(Boolean))].sort();
  categories.forEach(cat => {
    const pill = document.createElement("button");
    pill.className   = "cat-pill";
    pill.dataset.cat = cat;
    pill.textContent = cat;
    pill.addEventListener("click", () => selectCategory(cat));
    bar.appendChild(pill);
  });

  bar.querySelector("[data-cat='All']").addEventListener("click", () => selectCategory("All"));
}

function selectCategory(cat) {
  activeCategory = cat;
  document.querySelectorAll(".cat-pill").forEach(p => {
    p.classList.toggle("active", p.dataset.cat === cat);
  });
  applyFilters();
}

/* ═══════════════════════════════════════════════════════════════
   FILTER + RENDER
═══════════════════════════════════════════════════════════════ */
function applyFilters() {
  filteredList = allTemplates.filter(t => {
    const matchCat    = activeCategory === "All" || t.Category === activeCategory;
    const matchSearch = !searchQuery ||
      (t.Title       || "").toLowerCase().includes(searchQuery) ||
      (t.Description || "").toLowerCase().includes(searchQuery) ||
      (t.Category    || "").toLowerCase().includes(searchQuery);
    return matchCat && matchSearch;
  });
  renderTemplates();
}

function renderTemplates() {
  const list = document.getElementById("templateList");
  list.querySelectorAll(".template-card").forEach(el => el.remove());
  hideAllStates();

  if (filteredList.length === 0) {
    showState("empty");
    return;
  }

  filteredList.forEach((template, i) => {
    list.appendChild(buildCard(template, i));
  });
}

function buildCard(template, index) {
  const card = document.createElement("div");
  card.className = "template-card";
  card.style.animationDelay = `${index * 30}ms`;
  card.innerHTML = `
    <div class="card-icon">
      <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
        <rect x="1" y="1" width="12" height="12" rx="2" stroke="currentColor" stroke-width="1.3"/>
        <path d="M3.5 4.5h7M3.5 7h5M3.5 9.5h4" stroke="currentColor" stroke-width="1.2" stroke-linecap="round"/>
      </svg>
    </div>
    <div class="card-body">
      <div class="card-name">${escHtml(template.Title || "Untitled")}</div>
      ${template.Description ? `<div class="card-desc">${escHtml(template.Description)}</div>` : ""}
      <div class="card-meta">
        ${template.Category ? `<span class="card-category">${escHtml(template.Category)}</span>` : ""}
      </div>
    </div>
    <div class="card-arrow">
      <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
        <path d="M5 3l4 4-4 4" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/>
      </svg>
    </div>`;
  card.addEventListener("click", () => openPreview(template));
  return card;
}

/* ═══════════════════════════════════════════════════════════════
   PREVIEW DRAWER
═══════════════════════════════════════════════════════════════ */
async function openPreview(template) {
  selectedTemplate = template;
  document.getElementById("previewTitle").textContent = template.Title || "Preview";
  document.getElementById("previewBody").innerHTML    = '<div class="spinner"></div>';
  document.getElementById("previewOverlay").classList.remove("hidden");

  try {
    if (!template.TemplateFileUrl) throw new Error("This template has no file URL configured.");

    const html   = await fetchTemplateHtml(template.TemplateFileUrl);
    const iframe = document.createElement("iframe");
    iframe.sandbox = "allow-same-origin";
    document.getElementById("previewBody").innerHTML = "";
    document.getElementById("previewBody").appendChild(iframe);

    const doc = iframe.contentDocument || iframe.contentWindow.document;
    doc.open(); doc.write(html); doc.close();
    iframe.style.height = (doc.body.scrollHeight + 20) + "px";

  } catch (err) {
    document.getElementById("previewBody").innerHTML =
      `<p style="color:var(--danger);font-size:12px;">${escHtml(err.message)}</p>`;
  }
}

function closePreview() {
  document.getElementById("previewOverlay").classList.add("hidden");
  selectedTemplate = null;
}

/* ═══════════════════════════════════════════════════════════════
   INSERT INTO EMAIL — Office.js
═══════════════════════════════════════════════════════════════ */
async function insertTemplate() {
  if (!selectedTemplate) return;

  const insertBtn = document.getElementById("previewInsert");
  insertBtn.textContent = "Inserting…";
  insertBtn.disabled    = true;

  try {
    const html = await fetchTemplateHtml(selectedTemplate.TemplateFileUrl);

    await new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.setAsync(
        html,
        { coercionType: Office.CoercionType.Html },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error(result.error?.message || "Insert failed."));
        }
      );
    });

    closePreview();
    showToast("✓ Template inserted");

  } catch (err) {
    console.error("[TemplateAddin] Insert failed:", err);
    showToast("Insert failed: " + (err.message || "unknown error"), true);
  } finally {
    insertBtn.textContent = "Insert template";
    insertBtn.disabled    = false;
  }
}

/* ═══════════════════════════════════════════════════════════════
   UI HELPERS
═══════════════════════════════════════════════════════════════ */
function showState(state, errorMsg) {
  hideAllStates();
  if (state === "loading") document.getElementById("stateLoading").classList.remove("hidden");
  else if (state === "empty") document.getElementById("stateEmpty").classList.remove("hidden");
  else if (state === "error") {
    document.getElementById("stateErrorMsg").textContent = errorMsg || "Could not load templates.";
    document.getElementById("stateError").classList.remove("hidden");
  }
}

function hideAllStates() {
  ["stateLoading", "stateEmpty", "stateError"].forEach(id =>
    document.getElementById(id).classList.add("hidden")
  );
}

function showToast(message, isError = false) {
  const toast = document.getElementById("toast");
  toast.textContent  = message;
  toast.style.background = isError ? "var(--danger)" : "var(--text)";
  toast.classList.remove("hidden");
  setTimeout(() => toast.classList.add("hidden"), 3000);
}

function spinRefresh(active) {
  document.getElementById("refreshBtn").classList.toggle("spinning", active);
}

function escHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

"use strict";

/* ═══════════════════════════════════════════════════════════════════════════
   taskpane.js — Xero Data Link for Excel v1.0
   Complete logic: OAuth, Excel read/write, Xero API, wizard, refresh
   ═══════════════════════════════════════════════════════════════════════════ */

// ── SECTION 1: CONSTANTS & STATE ─────────────────────────────────────────────

const APP_URL = "https://sunny-tartufo-d82fdc.netlify.app";

const XERO_AUTH_URL = "https://login.xero.com/identity/connect/authorize";
const XERO_API      = "https://api.xero.com/api.xro/2.0";
// All Xero API calls now go through the Netlify proxy (avoids Excel CORS sandbox)
const SCOPES = [
  "accounting.reports.profitandloss.read",
  "accounting.reports.balancesheet.read",
  "accounting.settings.read",
  "offline_access"
].join(" ");

const LS = {
  CID: "xdl_cid", SEC: "xdl_sec", TOK: "xdl_tok",
  TID: "xdl_tid", TNAME: "xdl_tname"
};

const MONTH_MAP = {
  jan:1,feb:2,mar:3,apr:4,may:5,jun:6,
  jul:7,aug:8,sep:9,oct:10,nov:11,dec:12
};
const MONTH_NAMES = ["","Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const MONTH_FULL  = ["","January","February","March","April","May","June",
                     "July","August","September","October","November","December"];

// Wizard state
let state = {
  step: 1,
  outputSheet: null,
  outputCell: null,
  tenants: [],
  accounts: [],           // full chart of accounts from Xero
  selectedAccountIds: [],  // AccountIDs the user checked
  trackingCategories: [],  // from Xero
  reportType: "ProfitAndLoss",
  trackingCatId: "",
  trackingOptId: "",
  startYear: 0, startMonth: 0,
  endYear: 0, endMonth: 0,
  savedConfig: null        // loaded from workbook settings
};

let _pkceVerifier = null;
let _oauthState   = null;
let _authDialog   = null;


// ── SECTION 2: OFFICE INITIALIZATION ─────────────────────────────────────────

Office.onReady(async function(info) {
  if (info.host !== Office.HostType.Excel) return;

  // Restore credentials
  const el = (id) => document.getElementById(id);
  el("inputClientId").value     = localStorage.getItem(LS.CID) || "";
  el("inputClientSecret").value = localStorage.getItem(LS.SEC) || "";

  // Set default dates (12 months ending this month)
  const now = new Date();
  const endY = now.getFullYear(), endM = now.getMonth() + 1;
  let startY = endY - 1, startM = endM + 1;
  if (startM > 12) { startM -= 12; startY++; }
  populateDateDropdowns(startY, startM, endY, endM);

  // Check for saved config in the workbook
  try {
    state.savedConfig = await loadLinkConfig();
  } catch(e) { /* no config yet */ }

  // Update connection UI
  updateConnUI();

  // If we have a saved config and are connected, show refresh panel
  if (state.savedConfig && isConnected()) {
    showRefreshPanel();
  } else {
    log("Ready. Select an output cell and connect to Xero.");
  }
});


// ── SECTION 3: CREDENTIALS ──────────────────────────────────────────────────

function saveCreds() {
  const cid = document.getElementById("inputClientId").value.trim();
  const sec = document.getElementById("inputClientSecret").value.trim();
  if (!cid || !sec) { log("Enter both Client ID and Secret.", "warn"); return; }
  localStorage.setItem(LS.CID, cid);
  localStorage.setItem(LS.SEC, sec);
  log("✓ Credentials saved.", "ok");
}


// ── SECTION 4: TOKEN MANAGEMENT ──────────────────────────────────────────────

function getToken() {
  try {
    const raw = localStorage.getItem(LS.TOK);
    if (!raw) return null;
    const tok = JSON.parse(raw);
    const exp = (tok.obtained_at || 0) + (tok.expires_in || 1800) - 60;
    return (Date.now() / 1000 < exp) ? tok : null;
  } catch { return null; }
}

function saveToken(tok) {
  tok.obtained_at = Date.now() / 1000;
  localStorage.setItem(LS.TOK, JSON.stringify(tok));
}

function isConnected() {
  return !!getToken() && !!localStorage.getItem(LS.TID);
}

function disconnect() {
  [LS.TOK, LS.TID, LS.TNAME].forEach(k => localStorage.removeItem(k));
  state.tenants = [];
  updateConnUI();
  log("Disconnected.");
}


// ── SECTION 5: CONNECTION UI ─────────────────────────────────────────────────

function updateConnUI() {
  const connected = isConnected();
  const org       = localStorage.getItem(LS.TNAME) || "";
  const badge     = document.getElementById("connBadge");
  const btnConn   = document.getElementById("btnConnect");
  const btnDisc   = document.getElementById("btnDisconnect");
  const btnNext1  = document.getElementById("btnNext1");
  const orgWrap   = document.getElementById("orgSelectorWrap");

  if (connected) {
    badge.textContent = "● Connected: " + org;
    badge.className   = "badge badge-success";
    btnConn.style.display = "none";
    btnDisc.style.display = "inline-flex";
    btnNext1.disabled = !state.outputCell; // need location AND connection
  } else {
    badge.textContent = "● Not connected";
    badge.className   = "badge badge-neutral";
    btnConn.style.display = "block";
    btnDisc.style.display = "none";
    btnNext1.disabled = true;
  }

  // Multi-org dropdown
  if (connected && state.tenants.length > 1) {
    orgWrap.style.display = "block";
  } else {
    orgWrap.style.display = "none";
  }
}


// ── SECTION 6: OAUTH FLOW ────────────────────────────────────────────────────

function genVerifier() {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~";
  const arr = new Uint8Array(64);
  crypto.getRandomValues(arr);
  return Array.from(arr, n => chars[n % chars.length]).join("");
}

async function genChallenge(v) {
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(v));
  return btoa(String.fromCharCode(...new Uint8Array(buf)))
    .replace(/\+/g, "-").replace(/\//g, "_").replace(/=/g, "");
}

async function connectXero() {
  const cid = localStorage.getItem(LS.CID);
  const sec = localStorage.getItem(LS.SEC);
  if (!cid || !sec) { saveCreds(); if (!localStorage.getItem(LS.CID)) return; }

  _pkceVerifier = genVerifier();
  _oauthState   = crypto.randomUUID();
  const challenge = await genChallenge(_pkceVerifier);

  const params = new URLSearchParams({
    response_type: "code", client_id: cid,
    redirect_uri: APP_URL + "/auth-dialog.html",
    scope: SCOPES, state: _oauthState,
    code_challenge: challenge, code_challenge_method: "S256"
  });

  log("Opening Xero login…");
  const btn = document.getElementById("btnConnect");
  btn.disabled = true; btn.textContent = "⏳ Waiting…";

  Office.context.ui.displayDialogAsync(
    XERO_AUTH_URL + "?" + params,
    { height: 60, width: 40, displayInIframe: false },
    function(result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        log("✗ Could not open login: " + result.error.message, "err");
        resetConnBtn(); return;
      }
      _authDialog = result.value;
      _authDialog.addEventHandler(Office.EventType.DialogMessageReceived, onAuthMsg);
      _authDialog.addEventHandler(Office.EventType.DialogEventReceived, function(e) {
        if (e.error === 12006) { log("Login cancelled.", "warn"); resetConnBtn(); }
      });
    }
  );
}

function resetConnBtn() {
  const btn = document.getElementById("btnConnect");
  btn.disabled = false; btn.textContent = "🔗 Connect to Xero";
}

async function onAuthMsg(args) {
  // Parse the message BEFORE closing the dialog
  let msg;
  try {
    msg = JSON.parse(args.message);
  } catch(e) {
    log("✗ Bad auth response: " + e.message, "err");
    resetConnBtn(); return;
  }

  // Close dialog after parsing message
  if (_authDialog) { try { _authDialog.close(); } catch(e) {} }
  resetConnBtn();

  if (msg.type === "error") { log("✗ " + msg.message, "err"); return; }
  if (msg.type !== "code")  { log("✗ Unexpected type: " + msg.type, "err"); return; }
  if (msg.state !== _oauthState) { log("✗ Security check failed.", "err"); return; }

  log("Auth code received. Connecting to Xero…");

  // Retry helper — dialog closing can briefly interrupt network
  async function tryFetch(url, options, retries) {
    for (var i = 0; i <= retries; i++) {
      try {
        return await fetch(url, options);
      } catch(e) {
        if (i === retries) throw new Error("Network error after " + (retries+1) + " attempts: " + e.message);
        log("Network blip — retrying (" + (i+1) + "/" + retries + ")…", "warn");
        await new Promise(function(r) { setTimeout(r, 1200); });
      }
    }
  }

  try {
    // Exchange auth code for access token via Netlify function
    var tokenUrl = APP_URL + "/.netlify/functions/token";
    log("Calling: " + tokenUrl);

    var resp = await tryFetch(tokenUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        code:         msg.code,
        verifier:     _pkceVerifier,
        redirect_uri: APP_URL + "/auth-dialog.html"
      })
    }, 3);

    log("Response: HTTP " + resp.status);

    var respText = await resp.text();
    var tok;
    try { tok = JSON.parse(respText); }
    catch(e) { throw new Error("Non-JSON response: " + respText.substring(0, 200)); }

    if (!resp.ok) {
      throw new Error("Token exchange failed (" + resp.status + "): " +
        (tok.error || JSON.stringify(tok)));
    }
    if (!tok.access_token) {
      throw new Error("No access_token in response: " + JSON.stringify(tok));
    }

    saveToken(tok);
    log("✓ Token received.", "ok");

    // Get list of Xero organisations via Netlify proxy
    var tenResp = await tryFetch(APP_URL + "/.netlify/functions/xero", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ endpoint: "connections", token: tok.access_token })
    }, 2);

    state.tenants = await tenResp.json();
    if (!state.tenants || !state.tenants.length) {
      throw new Error("No Xero organisations found for this account.");
    }

    localStorage.setItem(LS.TID,   state.tenants[0].tenantId);
    localStorage.setItem(LS.TNAME, state.tenants[0].tenantName);

    if (state.tenants.length > 1) {
      var sel = document.getElementById("selOrg");
      sel.innerHTML = state.tenants.map(function(t) {
        return '<option value="' + t.tenantId + '">' + t.tenantName + '</option>';
      }).join("");
    }

    log("✓ Connected: " + state.tenants[0].tenantName, "ok");
    updateConnUI();

  } catch(e) {
    log("✗ " + e.message, "err");
    log("Check: Netlify env vars set? Xero redirect URI correct? Try again.", "warn");
  }
}


function selectOrg() {
  const sel = document.getElementById("selOrg");
  const t = state.tenants.find(t => t.tenantId === sel.value);
  if (t) {
    localStorage.setItem(LS.TID, t.tenantId);
    localStorage.setItem(LS.TNAME, t.tenantName);
    log("Switched to: " + t.tenantName);
    updateConnUI();
  }
}


// ── SECTION 7: XERO API CALLS ───────────────────────────────────────────────

async function xeroGet(path, params) {
  // Routes through Netlify proxy to avoid CORS restrictions in Excel sandbox
  const tok    = getToken();
  const tenant = localStorage.getItem(LS.TID);
  if (!tok || !tenant) throw new Error("Not connected to Xero.");

  const resp = await fetch(APP_URL + "/.netlify/functions/xero", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      endpoint: path,
      token:    tok.access_token,
      tenantId: tenant,
      params:   params || null
    })
  });
  if (!resp.ok) throw new Error("Proxy error " + resp.status + " on " + path);
  return resp.json();
}

async function fetchAccounts() {
  log("Loading chart of accounts…");
  const data = await xeroGet("Accounts");
  state.accounts = (data.Accounts || [])
    .filter(a => a.Status === "ACTIVE")
    .sort((a, b) => {
      if (a.Type !== b.Type) return a.Type < b.Type ? -1 : 1;
      return (a.Code || "").localeCompare(b.Code || "");
    });
  log("✓ " + state.accounts.length + " active accounts loaded.", "ok");
}

async function fetchTrackingCategories() {
  const data = await xeroGet("TrackingCategories");
  state.trackingCategories = data.TrackingCategories || [];
  if (state.trackingCategories.length > 0) {
    document.getElementById("cardTracking").style.display = "flex";
    const sel = document.getElementById("selTrackingCat");
    sel.innerHTML = '<option value="">No filter — show all data</option>' +
      state.trackingCategories.map(tc =>
        '<option value="' + tc.TrackingCategoryID + '">' + tc.Name + '</option>'
      ).join("");
    log("✓ " + state.trackingCategories.length + " tracking categories loaded.", "ok");
  }
}

function onTrackingCatChange() {
  const catId = document.getElementById("selTrackingCat").value;
  const wrap  = document.getElementById("trackingOptionWrap");
  if (!catId) { wrap.style.display = "none"; return; }

  const cat = state.trackingCategories.find(c => c.TrackingCategoryID === catId);
  if (!cat || !cat.Options || !cat.Options.length) { wrap.style.display = "none"; return; }

  wrap.style.display = "block";
  const sel = document.getElementById("selTrackingOpt");
  sel.innerHTML = cat.Options.map(o =>
    '<option value="' + o.TrackingOptionID + '">' + o.Name + '</option>'
  ).join("");
}

async function fetchReport(reportType, year, month) {
  const fromDate = year + "-" + pad2(month) + "-01";
  const lastDay  = new Date(year, month, 0).getDate();
  const toDate   = year + "-" + pad2(month) + "-" + pad2(lastDay);

  const params = {
    fromDate: fromDate, toDate: toDate,
    standardLayout: "true", paymentsOnly: "false"
  };

  // Add tracking filter if set
  if (state.trackingCatId && state.trackingOptId) {
    params.trackingCategoryID = state.trackingCatId;
    params.trackingOptionID   = state.trackingOptId;
  }

  const data = await xeroGet("Reports/" + reportType, params);
  const results = {};

  function walk(rows) {
    if (!rows) return;
    for (const row of rows) {
      const rt = row.RowType || "";
      if (rt === "Section" || rt === "SummaryRow") {
        walk(row.Rows);
      } else if (rt === "Row") {
        const cells = row.Cells || [];
        if (cells.length >= 2) {
          const name = (cells[0].Value || "").trim();
          // Extract AccountID from attributes if available
          let accountId = null;
          const attrs = cells[0].Attributes || [];
          for (const a of attrs) {
            if (a.Id === "account") accountId = a.Value;
          }
          const val = parseFloat((cells[1].Value || "0").replace(/,/g, "")) || 0;
          if (accountId) {
            results[accountId] = { name: name, amount: val };
          }
          if (name) {
            results["name:" + name.toLowerCase()] = { name: name, amount: val };
          }
        }
      }
    }
  }
  walk((data.Reports || [{}])[0].Rows || []);
  return results;
}


// ── SECTION 8: WIZARD STEP MANAGEMENT ────────────────────────────────────────

async function goStep(n) {
  // Validation before advancing
  if (n === 2 && state.step === 1) {
    if (!state.outputCell) { log("Please capture a cell selection first.", "warn"); return; }
    if (!isConnected()) { log("Please connect to Xero first.", "warn"); return; }
    // Load accounts and tracking categories when entering step 2
    try {
      document.getElementById("accountsLoading").style.display = "block";
      document.getElementById("accountsList").style.display = "none";
      await fetchAccounts();
      await fetchTrackingCategories();
      renderAccountsList();
    } catch(e) {
      log("✗ Failed to load accounts: " + e.message, "err");
      return;
    }
  }

  if (n === 3 && state.step === 2) {
    // Collect selected accounts
    collectSelectedAccounts();
    if (state.selectedAccountIds.length === 0) {
      log("Please select at least one account.", "warn"); return;
    }
    // Collect tracking selection
    state.trackingCatId = document.getElementById("selTrackingCat").value;
    state.trackingOptId = document.getElementById("selTrackingOpt")?.value || "";
  }

  if (n === 4 && state.step === 3) {
    collectDates();
    if (!state.startYear || !state.endYear) {
      log("Please select a date range.", "warn"); return;
    }
    renderFetchSummary();
  }

  state.step = n;

  // Update step indicators
  for (let i = 1; i <= 4; i++) {
    const item = document.getElementById("si" + i);
    const dot  = document.getElementById("sd" + i);
    item.className = "step-item";
    if (i < n)       { item.classList.add("done");   dot.textContent = "✓"; }
    else if (i === n) { item.classList.add("active"); dot.textContent = String(i); }
    else              { dot.textContent = String(i); }
  }

  // Show/hide panels
  for (let i = 1; i <= 4; i++) {
    document.getElementById("panel" + i).style.display = (i === n) ? "flex" : "none";
  }
  document.getElementById("panelRefresh").style.display = "none";

  // Scroll to top
  document.getElementById("mainContent").scrollTop = 0;
}


// ── SECTION 9: ACCOUNT LIST RENDERING ────────────────────────────────────────

function onReportTypeChange() {
  state.reportType = document.querySelector('input[name="reportType"]:checked').value;
  if (state.accounts.length > 0) renderAccountsList();
}

function renderAccountsList() {
  const container = document.getElementById("accountsList");
  const reportType = state.reportType;

  // Filter accounts by type based on report selection
  let relevantTypes;
  if (reportType === "ProfitAndLoss") {
    relevantTypes = ["REVENUE", "DIRECTCOSTS", "EXPENSE", "OVERHEADS", "OTHERINCOME", "OTHEREXPENSES"];
  } else {
    relevantTypes = ["BANK", "CURRENT", "CURRLIAB", "EQUITY", "FIXED", "LIABILITY",
                     "NONCURRENT", "OTHERASSET", "PREPAYMENT", "TERMLIAB"];
  }

  const filtered = state.accounts.filter(a => relevantTypes.includes(a.Type));

  // Group by class (REVENUE, EXPENSE, ASSET, LIABILITY, EQUITY)
  const groups = {};
  for (const a of filtered) {
    const cls = a.Class || a.Type;
    if (!groups[cls]) groups[cls] = [];
    groups[cls].push(a);
  }

  // Pretty names for groups
  const groupLabels = {
    REVENUE: "Revenue", EXPENSE: "Expenses",
    ASSET: "Assets", LIABILITY: "Liabilities", EQUITY: "Equity"
  };

  let html = "";
  for (const [cls, accts] of Object.entries(groups)) {
    const label = groupLabels[cls] || cls;
    const groupId = "grp_" + cls;
    html += '<div class="acct-group-header">' +
      '<input type="checkbox" checked onchange="toggleGroup(\'' + groupId + '\', this.checked)"/> ' +
      label + ' (' + accts.length + ')</div>';
    for (const a of accts) {
      html += '<div class="acct-row" data-group="' + groupId + '">' +
        '<input type="checkbox" checked value="' + a.AccountID + '" class="acct-cb"/>' +
        '<span class="acct-code">' + (a.Code || "") + '</span>' +
        '<span class="acct-name">' + a.Name + '</span></div>';
    }
  }

  container.innerHTML = html;
  container.style.display = "block";
  document.getElementById("accountsLoading").style.display = "none";
}

function toggleGroup(groupId, checked) {
  const rows = document.querySelectorAll('.acct-row[data-group="' + groupId + '"] .acct-cb');
  rows.forEach(cb => cb.checked = checked);
}

function collectSelectedAccounts() {
  const checkboxes = document.querySelectorAll(".acct-cb:checked");
  state.selectedAccountIds = Array.from(checkboxes).map(cb => cb.value);
}


// ── SECTION 10: DATE MANAGEMENT ──────────────────────────────────────────────

function populateDateDropdowns(defStartY, defStartM, defEndY, defEndM) {
  const startSel = document.getElementById("selStartMonth");
  const endSel   = document.getElementById("selEndMonth");

  // Generate months from 5 years ago to 2 years ahead
  const now = new Date();
  const fromY = now.getFullYear() - 5;
  const toY   = now.getFullYear() + 2;
  let options = "";

  for (let y = toY; y >= fromY; y--) {
    for (let m = 12; m >= 1; m--) {
      const label = MONTH_NAMES[m] + " " + y;
      const val   = y + "-" + m;
      const selStart = (y === defStartY && m === defStartM) ? " selected" : "";
      const selEnd   = (y === defEndY && m === defEndM) ? " selected" : "";
      options += '<option value="' + val + '"' + selStart + '>' + label + '</option>\n';
    }
  }

  // End dropdown gets its own options with different defaults
  let endOptions = "";
  for (let y = toY; y >= fromY; y--) {
    for (let m = 12; m >= 1; m--) {
      const label = MONTH_NAMES[m] + " " + y;
      const val   = y + "-" + m;
      const sel   = (y === defEndY && m === defEndM) ? " selected" : "";
      endOptions += '<option value="' + val + '"' + sel + '>' + label + '</option>\n';
    }
  }

  startSel.innerHTML = options;
  endSel.innerHTML   = endOptions;
  updateDatePreview();
}

function collectDates() {
  const startVal = document.getElementById("selStartMonth").value.split("-");
  const endVal   = document.getElementById("selEndMonth").value.split("-");
  state.startYear  = parseInt(startVal[0]);
  state.startMonth = parseInt(startVal[1]);
  state.endYear    = parseInt(endVal[0]);
  state.endMonth   = parseInt(endVal[1]);
}

function updateDatePreview() {
  collectDates();
  const months = countMonths(state.startYear, state.startMonth, state.endYear, state.endMonth);
  const el = document.getElementById("datePreview");
  if (months <= 0) {
    el.textContent = "⚠ End date must be after start date";
    el.style.color = "var(--danger-fg)";
  } else if (months > 60) {
    el.textContent = months + " columns — consider a shorter range (max ~36 recommended)";
    el.style.color = "var(--warning-fg)";
  } else {
    el.textContent = months + " monthly column" + (months > 1 ? "s" : "");
    el.style.color = "var(--brand)";
  }
}

function countMonths(sy, sm, ey, em) {
  return (ey - sy) * 12 + (em - sm) + 1;
}

function getMonthsList(sy, sm, ey, em) {
  const list = [];
  let y = sy, m = sm;
  while (y < ey || (y === ey && m <= em)) {
    list.push({ year: y, month: m });
    m++;
    if (m > 12) { m = 1; y++; }
  }
  return list;
}


// ── SECTION 11: EXCEL — CAPTURE SELECTION ────────────────────────────────────

async function captureSelection() {
  try {
    await Excel.run(async ctx => {
      const range = ctx.workbook.getSelectedRange();
      range.load("address");
      const sheet = range.worksheet;
      sheet.load("name");
      await ctx.sync();

      // Parse address — may be "Sheet1!A1" or just "A1"
      let addr = range.address;
      let sheetName = sheet.name;

      // Strip sheet name prefix if present in address
      if (addr.includes("!")) {
        addr = addr.split("!")[1];
      }
      // Take just the first cell if a range is selected
      if (addr.includes(":")) {
        addr = addr.split(":")[0];
      }

      state.outputSheet = sheetName;
      state.outputCell  = addr;

      const display = document.getElementById("locationDisplay");
      display.textContent = "📍 " + addr + " on " + sheetName;
      display.classList.remove("empty");

      // Enable next button if connected
      document.getElementById("btnNext1").disabled = !isConnected();

      log("Output location: " + addr + " on " + sheetName);
    });
  } catch(e) {
    log("✗ Could not read selection: " + e.message, "err");
  }
}


// ── SECTION 12: EXCEL — WRITE DATA ───────────────────────────────────────────

async function writeToExcel(allData, months) {
  await Excel.run(async ctx => {
    const sheet = ctx.workbook.worksheets.getItem(state.outputSheet);
    const startAddr = state.outputCell;

    // Parse start cell to get column and row
    const startCol = colLetterToNum(startAddr.replace(/[0-9]/g, ""));
    const startRow = parseInt(startAddr.replace(/[^0-9]/g, ""));

    const numCols = months.length;
    const selectedAccounts = state.accounts.filter(
      a => state.selectedAccountIds.includes(a.AccountID)
    );

    // Group selected accounts by class
    const groupLabels = {
      REVENUE: "Revenue", EXPENSE: "Expenses",
      ASSET: "Assets", LIABILITY: "Liabilities", EQUITY: "Equity"
    };

    const groups = {};
    const groupOrder = state.reportType === "ProfitAndLoss"
      ? ["REVENUE", "EXPENSE"]
      : ["ASSET", "LIABILITY", "EQUITY"];

    for (const a of selectedAccounts) {
      const cls = a.Class || a.Type;
      if (!groups[cls]) groups[cls] = [];
      groups[cls].push(a);
    }

    let currentRow = startRow;
    let cellsWritten = 0;
    const isBs = state.reportType === "BalanceSheet";

    // ── Write header row (month labels) ─────────────────────────────────
    for (let i = 0; i < numCols; i++) {
      const m = months[i];
      const cell = sheet.getCell(currentRow - 1, startCol - 1 + 1 + i);
      if (isBs) {
        const lastDay = new Date(m.year, m.month, 0).getDate();
        cell.values = [["As at " + lastDay + " " + MONTH_NAMES[m.month] + " " + m.year]];
      } else {
        cell.values = [[MONTH_NAMES[m.month] + " " + m.year]];
      }
      cell.format.font.bold = true;
      cell.format.font.size = 10;
      cell.format.horizontalAlignment = "Center";
      cell.format.fill.color = "#1B3A5C";
      cell.format.font.color = "#FFFFFF";
    }

    // Account label header
    const acctHeader = sheet.getCell(currentRow - 1, startCol - 1);
    acctHeader.values = [["Account"]];
    acctHeader.format.font.bold = true;
    acctHeader.format.font.size = 10;
    acctHeader.format.fill.color = "#1B3A5C";
    acctHeader.format.font.color = "#FFFFFF";

    currentRow++;

    // Track rows for grand total calculations
    const groupTotalRows = {};

    // ── Write each group ───────────────────────────────────────────────
    for (const cls of groupOrder) {
      const accts = groups[cls];
      if (!accts || accts.length === 0) continue;

      // Group header row (bold, grey background)
      const headerCell = sheet.getCell(currentRow - 1, startCol - 1);
      headerCell.values = [[groupLabels[cls] || cls]];
      headerCell.format.font.bold = true;
      headerCell.format.font.size = 10;
      headerCell.format.fill.color = "#E8E8E8";
      // Fill the rest of the header row too
      for (let i = 0; i < numCols; i++) {
        sheet.getCell(currentRow - 1, startCol - 1 + 1 + i).format.fill.color = "#E8E8E8";
      }
      currentRow++;

      const firstDataRow = currentRow;

      // Individual account rows
      for (const acct of accts) {
        // Account name in first column
        const nameCell = sheet.getCell(currentRow - 1, startCol - 1);
        nameCell.values = [["  " + acct.Name]];
        nameCell.format.font.size = 10;
        nameCell.format.font.color = "#242424";

        // Data cells for each month
        for (let i = 0; i < numCols; i++) {
          const m = months[i];
          const periodKey = m.year + "-" + m.month;
          const periodData = allData[periodKey] || {};

          // Try matching by AccountID first, then by name
          let amount = null;
          if (periodData[acct.AccountID]) {
            amount = periodData[acct.AccountID].amount;
          } else {
            const nameKey = "name:" + acct.Name.toLowerCase();
            if (periodData[nameKey]) {
              amount = periodData[nameKey].amount;
            }
          }

          const dataCell = sheet.getCell(currentRow - 1, startCol - 1 + 1 + i);
          if (amount !== null && amount !== undefined) {
            dataCell.values = [[amount]];
            dataCell.numberFormat = [['#,##0;(#,##0);"-"']];
            cellsWritten++;
          } else {
            dataCell.values = [[null]];
            dataCell.format.fill.color = "#FFF4CE"; // amber tint
          }
          dataCell.format.horizontalAlignment = "Right";
          dataCell.format.font.size = 10;
        }
        currentRow++;
      }

      const lastDataRow = currentRow - 1;

      // Subtotal row (bold, with SUM formulas)
      const totalLabel = sheet.getCell(currentRow - 1, startCol - 1);
      totalLabel.values = [["Total " + (groupLabels[cls] || cls)]];
      totalLabel.format.font.bold = true;
      totalLabel.format.font.size = 10;
      totalLabel.format.fill.color = "#F0F0F0";

      for (let i = 0; i < numCols; i++) {
        const colLetter = numToColLetter(startCol + 1 + i);
        const totalCell = sheet.getCell(currentRow - 1, startCol - 1 + 1 + i);
        totalCell.formulas = [["=SUM(" + colLetter + firstDataRow + ":" + colLetter + lastDataRow + ")"]];
        totalCell.numberFormat = [['#,##0;(#,##0);"-"']];
        totalCell.format.font.bold = true;
        totalCell.format.font.size = 10;
        totalCell.format.fill.color = "#F0F0F0";
        totalCell.format.horizontalAlignment = "Right";
        totalCell.format.borders.getItem("EdgeBottom").style = "Thin";
      }
      totalLabel.format.borders.getItem("EdgeBottom").style = "Thin";

      groupTotalRows[cls] = currentRow;
      currentRow++;

      // Spacer row
      currentRow++;
    }

    // ── Grand total rows (P&L: Gross Profit, Net Profit | BS: Net Assets) ──
    if (state.reportType === "ProfitAndLoss") {
      const revRow = groupTotalRows["REVENUE"];
      const expRow = groupTotalRows["EXPENSE"];

      if (revRow && expRow) {
        // Net Profit = Total Revenue - Total Expenses
        const npLabel = sheet.getCell(currentRow - 1, startCol - 1);
        npLabel.values = [["Net Profit"]];
        npLabel.format.font.bold = true;
        npLabel.format.font.size = 11;
        npLabel.format.fill.color = "#D9E2EC";

        for (let i = 0; i < numCols; i++) {
          const colLetter = numToColLetter(startCol + 1 + i);
          const npCell = sheet.getCell(currentRow - 1, startCol - 1 + 1 + i);
          npCell.formulas = [["=" + colLetter + revRow + "-" + colLetter + expRow]];
          npCell.numberFormat = [['#,##0;(#,##0);"-"']];
          npCell.format.font.bold = true;
          npCell.format.font.size = 11;
          npCell.format.fill.color = "#D9E2EC";
          npCell.format.horizontalAlignment = "Right";
          npCell.format.borders.getItem("EdgeBottom").style = "Medium";
          npCell.format.borders.getItem("EdgeTop").style = "Thin";
        }
        npLabel.format.borders.getItem("EdgeBottom").style = "Medium";
        npLabel.format.borders.getItem("EdgeTop").style = "Thin";
      }
    }

    // Update org name and timestamp
    const orgName = localStorage.getItem(LS.TNAME) || "Xero";

    await ctx.sync();
    return cellsWritten;
  });
}


// ── SECTION 13: FETCH ORCHESTRATION ──────────────────────────────────────────

function renderFetchSummary() {
  const months = countMonths(state.startYear, state.startMonth, state.endYear, state.endMonth);
  const selCount = state.selectedAccountIds.length;
  const org = localStorage.getItem(LS.TNAME) || "";
  const isBs = state.reportType === "BalanceSheet";
  const trackingName = getTrackingLabel();

  let html = "<strong>Organisation:</strong> " + org + "<br>" +
    "<strong>Report:</strong> " + (isBs ? "Balance Sheet" : "Profit & Loss") + "<br>" +
    "<strong>Accounts:</strong> " + selCount + " selected<br>" +
    "<strong>Period:</strong> " + MONTH_NAMES[state.startMonth] + " " + state.startYear +
    " → " + MONTH_NAMES[state.endMonth] + " " + state.endYear +
    " (" + months + " columns)<br>" +
    "<strong>Output:</strong> " + state.outputCell + " on " + state.outputSheet;

  if (trackingName) {
    html += "<br><strong>Tracking:</strong> " + trackingName;
  }

  document.getElementById("fetchSummary").innerHTML = html;
}

function getTrackingLabel() {
  if (!state.trackingCatId) return "";
  const cat = state.trackingCategories.find(c => c.TrackingCategoryID === state.trackingCatId);
  if (!cat) return "";
  const opt = (cat.Options || []).find(o => o.TrackingOptionID === state.trackingOptId);
  return cat.Name + ": " + (opt ? opt.Name : "All");
}

async function doFetch() {
  const btn = document.getElementById("btnFetch");
  btn.disabled = true; btn.textContent = "⏳ Fetching…";

  document.getElementById("cardProgress").style.display = "flex";
  document.getElementById("cardResults").style.display = "none";

  const months = getMonthsList(state.startYear, state.startMonth, state.endYear, state.endMonth);
  const allData = {};
  let fetchErrors = 0;

  try {
    for (let i = 0; i < months.length; i++) {
      const m = months[i];
      const label = MONTH_FULL[m.month] + " " + m.year;
      const pct = Math.round(((i + 1) / months.length) * 85);

      setProgress(pct, "Fetching " + label + "…", (i + 1) + " of " + months.length);
      log("Fetching " + label + "…");

      try {
        allData[m.year + "-" + m.month] = await fetchReport(state.reportType, m.year, m.month);
        const count = Object.keys(allData[m.year + "-" + m.month]).length;
        log("  ✓ " + (count / 2) + " accounts", "ok"); // divide by 2 since we store both ID and name keys
      } catch(e) {
        log("  ✗ " + label + ": " + e.message, "err");
        allData[m.year + "-" + m.month] = {};
        fetchErrors++;
      }

      // Small delay to respect rate limits for large ranges
      if (months.length > 30 && i > 0 && i % 30 === 0) {
        log("Pausing briefly for rate limits…", "warn");
        await sleep(2000);
      }
    }

    setProgress(90, "Writing to Excel…", "");
    log("Writing to spreadsheet…");

    const cellsWritten = await writeToExcel(allData, months);

    // Detect unmatched accounts — check which selected accounts appear in Xero data
    const selectedAccts = state.accounts.filter(
      a => state.selectedAccountIds.includes(a.AccountID)
    );
    const unmatched = [];
    for (const acct of selectedAccts) {
      let found = false;
      for (const key of Object.keys(allData)) {
        const pd = allData[key];
        if (pd[acct.AccountID] || pd["name:" + acct.Name.toLowerCase()]) {
          found = true; break;
        }
      }
      if (!found) unmatched.push(acct.Name);
    }

    // Save the link config for refresh
    await saveLinkConfig();

    setProgress(100, "Complete!", "");

    // Show results
    document.getElementById("statCells").textContent    = cellsWritten || "✓";
    document.getElementById("statMonths").textContent   = months.length;
    document.getElementById("statAccounts").textContent = state.selectedAccountIds.length;
    document.getElementById("cardResults").style.display = "flex";

    // Show unmatched accounts if any
    const unmatchedEl = document.getElementById("unmatchedList");
    if (unmatched.length > 0) {
      unmatchedEl.innerHTML = '<div class="card-title" style="color:var(--warning-fg)">⚠ Unmatched Accounts</div>' +
        '<p class="hint">These accounts were not found in Xero\'s data. Cells are highlighted amber. Check spelling matches your Xero chart of accounts.</p>' +
        '<div style="font-size:11px;color:var(--fg2);line-height:1.8;margin-top:4px">' +
        unmatched.map(n => "• " + n).join("<br>") + '</div>';
      unmatchedEl.style.display = "block";
    } else {
      unmatchedEl.style.display = "none";
    }

    log("✓ Done — " + (cellsWritten || 0) + " cells updated.", "ok");
    if (unmatched.length > 0) {
      log("⚠ " + unmatched.length + " account(s) not found in Xero data.", "warn");
    }
    if (fetchErrors > 0) {
      log("⚠ " + fetchErrors + " month(s) had errors.", "warn");
    }

  } catch(e) {
    log("✗ " + e.message, "err");
    setProgress(0, "", "");
  } finally {
    btn.disabled = false;
    btn.textContent = "🔄 Fetch Data from Xero";
  }
}


// ── SECTION 14: LINK CONFIG STORAGE & REFRESH ────────────────────────────────

async function saveLinkConfig() {
  const config = {
    version: "1.0",
    reportType: state.reportType,
    tenantId:   localStorage.getItem(LS.TID),
    tenantName: localStorage.getItem(LS.TNAME),
    selectedAccountIds: state.selectedAccountIds,
    trackingCatId: state.trackingCatId,
    trackingOptId: state.trackingOptId,
    startYear:  state.startYear,
    startMonth: state.startMonth,
    endYear:    state.endYear,
    endMonth:   state.endMonth,
    outputSheet: state.outputSheet,
    outputCell:  state.outputCell,
    trackingLabel: getTrackingLabel(),
    lastRefreshed: new Date().toISOString()
  };

  await Excel.run(async ctx => {
    ctx.workbook.settings.add("xeroDataLink", JSON.stringify(config));
    await ctx.sync();
  });

  state.savedConfig = config;
  log("✓ Link configuration saved in workbook.", "ok");
}

async function loadLinkConfig() {
  return await Excel.run(async ctx => {
    const setting = ctx.workbook.settings.getItemOrNullObject("xeroDataLink");
    setting.load("value");
    await ctx.sync();
    if (setting.isNullObject) return null;
    return JSON.parse(setting.value);
  });
}

function showRefreshPanel() {
  const cfg = state.savedConfig;
  if (!cfg) return;

  const isBs = cfg.reportType === "BalanceSheet";
  const label = isBs ? "Balance Sheet" : "Profit & Loss";
  const refreshDate = cfg.lastRefreshed
    ? new Date(cfg.lastRefreshed).toLocaleDateString("en-NZ", { day: "numeric", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" })
    : "Never";

  let html = "<strong>Report:</strong> " + label + "<br>" +
    "<strong>Organisation:</strong> " + (cfg.tenantName || "") + "<br>" +
    "<strong>Accounts:</strong> " + (cfg.selectedAccountIds || []).length + " selected<br>" +
    "<strong>Period:</strong> " + MONTH_NAMES[cfg.startMonth] + " " + cfg.startYear +
    " → " + MONTH_NAMES[cfg.endMonth] + " " + cfg.endYear + "<br>" +
    "<strong>Output:</strong> " + cfg.outputCell + " on " + cfg.outputSheet + "<br>";

  if (cfg.trackingLabel) {
    html += "<strong>Tracking:</strong> " + cfg.trackingLabel + "<br>";
  }
  html += "<strong>Last refreshed:</strong> " + refreshDate;

  document.getElementById("refreshSummary").innerHTML = html;

  // Check if a new month is available
  const now = new Date();
  const currentY = now.getFullYear(), currentM = now.getMonth() + 1;
  const endY = cfg.endYear, endM = cfg.endMonth;

  if (currentY > endY || (currentY === endY && currentM > endM)) {
    const newMonthLabel = MONTH_NAMES[currentM] + " " + currentY;
    const banner = document.getElementById("refreshNewMonth");
    banner.innerHTML = "💡 New month available: <strong>" + newMonthLabel +
      "</strong>. Click below to extend and refresh.";
    banner.style.display = "block";
    document.getElementById("btnRefreshNow").textContent = "🔄 Extend to " + newMonthLabel + " & Refresh";
  }

  // Hide wizard panels, show refresh panel
  for (let i = 1; i <= 4; i++) {
    document.getElementById("panel" + i).style.display = "none";
  }
  document.getElementById("panelRefresh").style.display = "flex";
  document.getElementById("stepTrack").style.display = "none";

  log("Saved data link loaded. Click Refresh to update with latest Xero data.");
}

function showWizard() {
  document.getElementById("panelRefresh").style.display = "none";
  document.getElementById("stepTrack").style.display = "flex";

  // Restore state from saved config
  if (state.savedConfig) {
    const cfg = state.savedConfig;
    state.reportType       = cfg.reportType;
    state.selectedAccountIds = cfg.selectedAccountIds || [];
    state.trackingCatId    = cfg.trackingCatId || "";
    state.trackingOptId    = cfg.trackingOptId || "";
    state.startYear        = cfg.startYear;
    state.startMonth       = cfg.startMonth;
    state.endYear          = cfg.endYear;
    state.endMonth         = cfg.endMonth;
    state.outputSheet      = cfg.outputSheet;
    state.outputCell       = cfg.outputCell;
  }

  goStep(1);
}

async function doRefresh() {
  const cfg = state.savedConfig;
  if (!cfg) { log("No saved configuration found.", "warn"); return; }
  if (!isConnected()) { log("Please connect to Xero first.", "warn"); return; }

  // Restore state from saved config
  state.reportType       = cfg.reportType;
  state.selectedAccountIds = cfg.selectedAccountIds;
  state.trackingCatId    = cfg.trackingCatId || "";
  state.trackingOptId    = cfg.trackingOptId || "";
  state.outputSheet      = cfg.outputSheet;
  state.outputCell       = cfg.outputCell;

  // Check if we should extend the date range
  const now = new Date();
  const currentY = now.getFullYear(), currentM = now.getMonth() + 1;

  if (currentY > cfg.endYear || (currentY === cfg.endYear && currentM > cfg.endMonth)) {
    state.endYear  = currentY;
    state.endMonth = currentM;
    log("Extended date range to " + MONTH_NAMES[currentM] + " " + currentY);
  } else {
    state.endYear  = cfg.endYear;
    state.endMonth = cfg.endMonth;
  }
  state.startYear  = cfg.startYear;
  state.startMonth = cfg.startMonth;

  // Show progress on refresh panel
  const btn = document.getElementById("btnRefreshNow");
  btn.disabled = true; btn.textContent = "⏳ Refreshing…";

  const months = getMonthsList(state.startYear, state.startMonth, state.endYear, state.endMonth);
  const allData = {};

  try {
    for (let i = 0; i < months.length; i++) {
      const m = months[i];
      const label = MONTH_FULL[m.month] + " " + m.year;
      log("Fetching " + label + "…");

      try {
        allData[m.year + "-" + m.month] = await fetchReport(state.reportType, m.year, m.month);
        log("  ✓ done", "ok");
      } catch(e) {
        log("  ✗ " + e.message, "err");
        allData[m.year + "-" + m.month] = {};
      }

      if (months.length > 30 && i > 0 && i % 30 === 0) {
        await sleep(2000);
      }
    }

    log("Writing to spreadsheet…");
    await writeToExcel(allData, months);
    await saveLinkConfig();

    log("✓ Refresh complete.", "ok");

    // Reload the refresh panel with updated info
    state.savedConfig = await loadLinkConfig();
    showRefreshPanel();

  } catch(e) {
    log("✗ " + e.message, "err");
  } finally {
    btn.disabled = false;
    btn.textContent = "🔄 Refresh Now";
  }
}


// ── SECTION 15: UTILITY FUNCTIONS ────────────────────────────────────────────

function pad2(n) { return String(n).padStart(2, "0"); }

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

function colLetterToNum(letters) {
  let n = 0;
  for (let i = 0; i < letters.length; i++) {
    n = n * 26 + (letters.charCodeAt(i) - 64);
  }
  return n;
}

function numToColLetter(num) {
  let s = "";
  while (num > 0) {
    let rem = (num - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    num = Math.floor((num - 1) / 26);
  }
  return s;
}

function setProgress(pct, label, sub) {
  document.getElementById("progressBar").style.width = Math.min(pct, 100) + "%";
  if (label) document.getElementById("progressLabel").textContent = label;
  if (sub !== undefined) document.getElementById("progressSub").textContent = sub || "";
}

function log(msg, type) {
  const area = document.getElementById("logArea");
  const line = document.createElement("div");
  const t = new Date().toLocaleTimeString("en-NZ", { hour12: false });
  line.className = "log-line" + (type ? " log-" + type : "");
  line.textContent = t + "  " + msg;
  area.appendChild(line);
  area.scrollTop = area.scrollHeight;
}

function clearLog() {
  document.getElementById("logArea").innerHTML = '<div class="log-line">Log cleared.</div>';
}

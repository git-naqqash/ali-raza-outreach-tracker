/* D:\Lead Tracker\app.js */

// Authentication Credentials
const LOGIN_USER = "contact.naqqash@gmail.com";
const LOGIN_PASS = "@@@03314200250";

// Valid outreach stages
const VALID_STAGES = [
  "Found (Lead collected only)",
  "First Message Sent",
  "First Follow-up Sent",
  "Second Follow-up Sent",
  "Engaged",
  "Samples Sent",
  "Warm Lead",
  "Follow-up Due",
  "Replied",
  "Not Now",
  "Archived"
];

// Helper to format a Date object as a local YYYY-MM-DD string
function formatLocalDate(d) {
  if (!d || !(d instanceof Date) || isNaN(d.getTime())) return "";
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

// Helper to format dates relative to today
function getOffsetDateString(daysOffset) {
  const d = new Date();
  d.setDate(d.getDate() + daysOffset);
  return formatLocalDate(d);
}

// Parse dates from excel (handles serial number and text/strings)
function parseExcelDate(val) {
  if (val === undefined || val === null || val === "") return null;
  
  if (val instanceof Date) {
    return formatLocalDate(val);
  }
  
  const numVal = parseFloat(val);
  if ((typeof val === 'number' && !isNaN(val)) || (typeof val === 'string' && !isNaN(val) && !isNaN(numVal) && numVal > 0 && numVal < 100000)) {
    const serial = numVal;
    const utc_days = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);
    const localDate = new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate());
    return formatLocalDate(localDate);
  }

  const str = String(val).trim();
  if (!str) return null;

  // Try checking DD/MM/YYYY pattern (e.g., 25/06/2026 or 25-06-2026)
  const dmyRegex = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/;
  const dmyMatch = str.match(dmyRegex);
  if (dmyMatch) {
    const day = parseInt(dmyMatch[1], 10);
    const month = parseInt(dmyMatch[2], 10);
    const year = parseInt(dmyMatch[3], 10);
    
    if (day >= 1 && day <= 31 && month >= 1 && month <= 12) {
      const d = new Date(year, month - 1, day);
      if (!isNaN(d.getTime())) {
        return formatLocalDate(d);
      }
    }
  }

  // Try checking YYYY-MM-DD pattern (e.g., 2026-06-25 or 2026/06/25)
  const ymdRegex = /^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/;
  const ymdMatch = str.match(ymdRegex);
  if (ymdMatch) {
    const year = parseInt(ymdMatch[1], 10);
    const month = parseInt(ymdMatch[2], 10);
    const day = parseInt(ymdMatch[3], 10);
    
    if (day >= 1 && day <= 31 && month >= 1 && month <= 12) {
      const d = new Date(year, month - 1, day);
      if (!isNaN(d.getTime())) {
        return formatLocalDate(d);
      }
    }
  }

  const parsedMs = Date.parse(str);
  if (!isNaN(parsedMs)) {
    return formatLocalDate(new Date(parsedMs));
  }

  return null;
}

// Add working days (Mon-Fri) to a date
function addWorkingDays(startDateStr, days) {
  let date = new Date(startDateStr);
  if (isNaN(date.getTime())) {
    date = new Date();
  }
  let added = 0;
  while (added < days) {
    date.setDate(date.getDate() + 1);
    const day = date.getDay();
    if (day !== 0 && day !== 6) { // Not Sunday (0) and not Saturday (6)
      added++;
    }
  }
  return formatLocalDate(date);
}

// Add calendar days to a date
function addDays(startDateStr, days) {
  let date = new Date(startDateStr);
  if (isNaN(date.getTime())) {
    date = new Date();
  }
  date.setDate(date.getDate() + days);
  return formatLocalDate(date);
}

// Check if a SheetJS cell style is colored red (background or fill)
function checkRowRedStyle(cell) {
  if (!cell || !cell.s) return false;
  const fill = cell.s.fill;
  if (fill && fill.fgColor) {
    const rgb = fill.fgColor.rgb;
    if (rgb && (rgb.startsWith("FF") || rgb.startsWith("ff") || rgb.toLowerCase() === "red")) {
      return true;
    }
  }
  return false;
}

// Normalize URL to prevent local relative path routing (Vercel 404s)
function normalizeUrl(url) {
  if (!url) return "";
  const trimmed = url.trim();
  if (!trimmed) return "";
  if (trimmed.startsWith("http://") || trimmed.startsWith("https://")) {
    return trimmed;
  }
  return "https://" + trimmed;
}

// Format WhatsApp click-to-chat URL
function getWhatsAppLink(number) {
  if (!number) return "";
  const cleaned = number.replace(/[\s\+\(\)\-\[\]]/g, ""); // Keep only digits
  return `https://wa.me/${cleaned}`;
}

// Default Seed Data for the Outreach CRM
const DEFAULT_LEADS = [
  {
    dateAdded: getOffsetDateString(-10),
    name: "Roberto C. (Scale Coaching)",
    market: "Italy",
    channel: "Instagram",
    mainLink: "https://instagram.com/robertocoach_scale",
    niche: "Business Coach",
    source: "Instagram Search",
    priority: "A",
    stage: "Warm Lead",
    lastActionDate: getOffsetDateString(-1),
    nextAction: "Send DM",
    nextActionDate: getOffsetDateString(0), // Today
    replyStatus: "Interested",
    notes: "Very interested in interactive workbook for his high-ticket masterclass. Already has a rough outline. Wants a quick pricing structure. Main contact Roberto.",
    contactPerson: "Roberto",
    email: "roberto@scalecoaching.it",
    whatsappNumber: "+39 347 123 4567",
    extraLink: "https://scalecoaching.it",
    messageSent: "Ciao Roberto! Ti andrebbe di dare un'occhiata a un workbook Canva di prova che ho fatto per altri coach?",
    followUpCount: 1
  },
  {
    dateAdded: getOffsetDateString(-12),
    name: "KDP Publishing Partners",
    market: "US",
    channel: "Email",
    mainLink: "https://kdp-publishing-partners-test.com",
    niche: "KDP Publisher",
    source: "Google",
    priority: "B",
    stage: "Replied",
    lastActionDate: getOffsetDateString(-4),
    nextAction: "Send samples",
    nextActionDate: getOffsetDateString(0), // Today
    replyStatus: "Samples requested",
    notes: "Self-publishing team focused on niche journals and workbook series. Requested samples of previous books in cooking and mindfulness niches. Italian market testing.",
    contactPerson: "Jane Smith",
    email: "submissions@kdppartners.com",
    whatsappNumber: "",
    extraLink: "",
    messageSent: "Sent initial pitch email outlining workbook drafting and formatting services.",
    followUpCount: 0
  },
  {
    dateAdded: getOffsetDateString(-5),
    name: "Dr. Andrea Gruber (Mindset Academy)",
    market: "Austria",
    channel: "LinkedIn",
    mainLink: "https://linkedin.com/in/andrea-gruber-mindset",
    niche: "Course Creator",
    source: "LinkedIn Search",
    priority: "A",
    stage: "First Message Sent",
    lastActionDate: getOffsetDateString(-5),
    nextAction: "Send follow-up",
    nextActionDate: getOffsetDateString(-1), // Overdue (Yesterday)
    replyStatus: "No reply",
    notes: "Strong fit. Sells \u20AC2k mindset courses. No active lead magnet on website. Sent first pitch on LinkedIn about workbook companion design.",
    contactPerson: "Andrea Gruber",
    email: "",
    whatsappNumber: "",
    extraLink: "https://mindset-academy.at",
    messageSent: "Hi Andrea, loved your post on cognitive reframing. Sells \u20AC2k mindset courses...",
    followUpCount: 0
  },
  {
    dateAdded: getOffsetDateString(-15),
    name: "Verlagsgruppe K. (Munich)",
    market: "Germany",
    channel: "Email",
    mainLink: "https://verlag-muenchen-test.de",
    niche: "Publishing Business",
    source: "Google",
    priority: "B",
    stage: "Engaged",
    lastActionDate: getOffsetDateString(-3),
    nextAction: "Send CV",
    nextActionDate: getOffsetDateString(2), // In 2 days
    replyStatus: "Replied",
    notes: "Discussed white-label novel formatting and workbook companion design for their self-help segment. Requested CV and rate sheet by Friday.",
    contactPerson: "Marcus Weber",
    email: "weber@verlagmuenchen.de",
    whatsappNumber: "+49 89 123456",
    extraLink: "",
    messageSent: "Sent email discussing our Canva design capabilities and formatting structures.",
    followUpCount: 0
  },
  {
    dateAdded: getOffsetDateString(-20),
    name: "Peak Performance Agency",
    market: "UK",
    channel: "WhatsApp",
    mainLink: "https://peak-performance-agency.co.uk",
    niche: "Content Agency",
    source: "Website",
    priority: "C",
    stage: "Not Now",
    lastActionDate: getOffsetDateString(-8),
    nextAction: "Wait",
    nextActionDate: getOffsetDateString(10), // Future
    replyStatus: "Not now",
    notes: "Expressed interest but budgets are frozen until next quarter. Keep warm and message back in 2 months. WhatsApp contact chosen because phone is active on their business bio.",
    contactPerson: "David H.",
    email: "hello@peakperformance.co.uk",
    whatsappNumber: "+44 7911 123456",
    extraLink: "",
    messageSent: "First pitch on WhatsApp. They replied saying 'Not now, check back in August'.",
    followUpCount: 1
  },
  {
    dateAdded: getOffsetDateString(-30),
    name: "Bestseller Labs",
    market: "US",
    channel: "LinkedIn",
    mainLink: "https://linkedin.com/company/bestsellerlabs",
    niche: "KDP Business",
    source: "LinkedIn Search",
    priority: "D",
    stage: "Archived",
    lastActionDate: getOffsetDateString(-20),
    nextAction: "Archive",
    nextActionDate: getOffsetDateString(-10),
    replyStatus: "Not interested",
    notes: "Spoke to lead, they do all writing in-house and do not use white-label writers. Archiving.",
    contactPerson: "Elena Vance",
    email: "elena@bestsellerlabs.com",
    whatsappNumber: "",
    extraLink: "",
    messageSent: "LinkedIn connection pitch followed by DM.",
    followUpCount: 1
  }
];

// Initialize State
let leads = [];
let activeTab = "All"; // All, Instagram, LinkedIn, Email, WhatsApp
let activeQuickFilter = "All"; // All, Today, FollowUp, A, Warm, Archived
let activeTodayFilter = "All"; // All, Email, WhatsApp, Instagram, LinkedIn, A, Warm, Overdue
let currentSortColumn = null;
let currentSortDirection = "asc";

// ============================================================
// NEON DATABASE SYNC ENGINE
// All DB calls go through /api/ serverless routes — no CORS,
// no CDN dependency, works in all browsers.
// ============================================================
let cloudReady = false;

// ── Sync status badge ────────────────────────────────────────
function setSyncStatus(status) {
  const badge = document.getElementById("syncStatusBadge");
  if (!badge) return;
  if (status === "connected") {
    badge.className = "sync-badge sync-connected";
    badge.textContent = "☁ Cloud Sync: Connected";
  } else if (status === "syncing") {
    badge.className = "sync-badge sync-syncing";
    badge.textContent = "↻ Cloud Sync: Syncing…";
  } else if (status === "offline") {
    badge.className = "sync-badge sync-offline";
    badge.textContent = "⚠ Cloud Sync: Offline";
  } else {
    badge.className = "sync-badge sync-checking";
    badge.textContent = "⬡ Cloud Sync: Checking…";
  }
}

// ── UUID generator ───────────────────────────────────────────
function generateId() {
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    return (c === "x" ? r : (r & 0x3) | 0x8).toString(16);
  });
}

function ensureLeadId(lead) {
  if (!lead.id) lead.id = generateId();
  return lead;
}

// ── Helpers: show/hide upload banner + error panel ───────────
function showUploadBanner(count) {
  const banner = document.getElementById("uploadLocalBanner");
  const btn    = document.getElementById("uploadLocalToCloudBtn");
  if (banner) banner.style.display = "flex";
  if (btn)    btn.textContent = `☁ Upload ${count} Local Leads to Cloud`;
}
function hideUploadBanner() {
  const banner = document.getElementById("uploadLocalBanner");
  if (banner) banner.style.display = "none";
}
function showCloudError(msg) {
  const errDiv = document.getElementById("cloudErrorDisplay");
  if (errDiv) { errDiv.style.display = "block"; errDiv.innerHTML = msg; }
}
function hideCloudError() {
  const errDiv = document.getElementById("cloudErrorDisplay");
  if (errDiv) { errDiv.style.display = "none"; errDiv.innerHTML = ""; }
}

// ── Core API fetch helper ─────────────────────────────────────
async function apiFetch(path, options = {}) {
  const defaults = { headers: { "Content-Type": "application/json" } };
  const merged = {
    ...defaults,
    ...options,
    headers: { ...defaults.headers, ...(options.headers || {}) }
  };
  if (merged.body && typeof merged.body !== "string") {
    merged.body = JSON.stringify(merged.body);
  }
  const res = await fetch(path, merged);
  const json = await res.json();
  if (!res.ok) throw new Error(json.error || `HTTP ${res.status}`);
  return json;
}

// ── Sync all current leads to cloud ──────────────────────────
async function syncLeadsToCloud() {
  if (!cloudReady) return;
  try {
    setSyncStatus("syncing");
    leads.forEach(ensureLeadId);
    await apiFetch("/api/leads", { method: "POST", body: { leads } });
    setSyncStatus("connected");
  } catch (err) {
    console.error("[Cloud] syncLeadsToCloud error:", err.message);
    setSyncStatus("offline");
  }
}

// ── Upload local localStorage data to cloud ───────────────────
async function uploadLocalDataToCloud() {
  const banner = document.getElementById("uploadLocalBanner");
  const btn    = document.getElementById("uploadLocalToCloudBtn");
  if (btn) { btn.disabled = true; btn.textContent = "↻ Uploading…"; }
  setSyncStatus("syncing");
  hideCloudError();

  try {
    const localLeads   = JSON.parse(localStorage.getItem("ali_raza_leads")   || "[]");

    localLeads.forEach(ensureLeadId);
    localStorage.setItem("ali_raza_leads",   JSON.stringify(localLeads));

    await apiFetch("/api/leads", { method: "POST", body: { leads: localLeads } });

    const data = await apiFetch("/api/leads");
    if (data.leads && data.leads.length > 0) {
      leads = data.leads;
      localStorage.setItem("ali_raza_leads", JSON.stringify(leads));
    }

    hideUploadBanner();
    cloudReady = true;
    setSyncStatus("connected");
    updateDashboard();
    renderLeads();
    renderTodayActions();
    showToast(`✅ Uploaded ${localLeads.length} leads to cloud!`, "success");

  } catch (err) {
    console.error("[Cloud] uploadLocalDataToCloud failed:", err);
    showToast("Upload failed: " + err.message, "error");
    setSyncStatus("offline");
    showCloudError("Upload error: " + err.message);
    if (btn) { btn.disabled = false; btn.textContent = "☁ Retry Upload"; }
  }
}
window.uploadLocalDataToCloud = uploadLocalDataToCloud;

// ── Test Cloud Connection button ──────────────────────────────
async function testCloudConnection() {
  setSyncStatus("syncing");
  hideCloudError();
  showToast("Testing cloud connection…", "info");

  try {
    const data = await apiFetch("/api/test");
    console.log("[Cloud Test] SUCCESS:", data);
    cloudReady = true;
    setSyncStatus("connected");
    showToast("✅ " + (data.message || "Cloud connected!"), "success");
    hideCloudError();
  } catch (err) {
    console.error("[Cloud Test] FAILED:", err);
    setSyncStatus("offline");
    cloudReady = false;

    let msg = err.message || String(err);
    if (msg.toLowerCase().includes("failed to fetch") || msg.toLowerCase().includes("networkerror")) {
      msg =
        "<b>Network error: app cannot reach /api/test</b><br><br>" +
        "1. Is this deployed on Vercel? API routes only work after deployment.<br>" +
        "2. Check Vercel → Project → <b>Environment Variables</b><br>" +
        "&nbsp;&nbsp;→ Add: <code>DATABASE_URL</code> = your Neon connection string<br>" +
        "3. Redeploy after adding the env var.";
    } else if (msg.includes("DATABASE_URL")) {
      msg =
        "<b>DATABASE_URL not set</b><br>" +
        "Go to Vercel → Project Settings → Environment Variables<br>" +
        "Add <code>DATABASE_URL</code> with your Neon connection string.";
    }
    showToast("Cloud test failed: " + err.message, "error");
    showCloudError(msg);
  }
}
window.testCloudConnection = testCloudConnection;

// ── On-login async init: fetch from cloud, fall back to localStorage ─
async function initAppData() {
  setSyncStatus("syncing");
  hideCloudError();

  try {
    await apiFetch("/api/test");
    cloudReady = true;

    const leadsData = await apiFetch("/api/leads");
    const localLeads = JSON.parse(localStorage.getItem("ali_raza_leads") || "[]");

    if (leadsData.leads && leadsData.leads.length > 0) {
      leads = leadsData.leads;
      localStorage.setItem("ali_raza_leads", JSON.stringify(leads));
      hideUploadBanner();
    } else if (localLeads.length > 0) {
      showUploadBanner(localLeads.length);
    }

    setSyncStatus("connected");

  } catch (err) {
    console.warn("[Cloud] initAppData failed, using local backup:", err.message);
    cloudReady = false;
    setSyncStatus("offline");

    let errMsg = err.message || String(err);
    if (errMsg.includes("DATABASE_URL")) {
      errMsg = "<b>DATABASE_URL not configured.</b><br>Go to Vercel → Project Settings → Environment Variables → Add DATABASE_URL.";
    } else if (errMsg.toLowerCase().includes("failed to fetch")) {
      errMsg = "<b>API not reachable.</b> API routes (/api/*) only work on Vercel or with <code>vercel dev</code> locally.";
    } else {
      errMsg = "Cloud error: " + errMsg;
    }
    showCloudError(errMsg);
    return;
  }

  updateDashboard();
  renderLeads();
  renderTodayActions();
}

// ============================================================
// END NEON DATABASE SYNC ENGINE
// ============================================================

// On Page Load
document.addEventListener("DOMContentLoaded", () => {
  loadData();            // Immediately load localStorage (fast, no flicker)
  setupEventListeners();
  initStorageNotice();   // Hide storage notice if previously dismissed
  // initTableResizableColumns(); // Column resizing has been disabled in favor of stable CSS widths
  initColumnVisibility(); // Initialize column visibility from preferences
  updateDashboard();
  renderLeads();
  renderTodayActions();
  checkAuth();
  initAppData();         // Async: sync with Neon cloud in background
});

// ── Storage Notice: dismiss with animation, remember in localStorage ──
function dismissStorageNotice() {
  const banner = document.getElementById("storageNoticeBanner");
  if (!banner) return;
  banner.classList.add("dismissed");
  setTimeout(() => { banner.style.display = "none"; }, 320);
  localStorage.setItem("ali_raza_notice_dismissed", "true");
}
function initStorageNotice() {
  const banner = document.getElementById("storageNoticeBanner");
  if (!banner) return;
  if (localStorage.getItem("ali_raza_notice_dismissed") === "true") {
    banner.style.display = "none";
  }
}
window.dismissStorageNotice = dismissStorageNotice;

// ── Column Visibility dropdown and toggling logic ──────────────────────
function toggleColVisibilityDropdown(event) {
  event.stopPropagation();
  const dropdown = document.getElementById("colVisibilityMenu");
  if (!dropdown) return;
  const wasOpen = dropdown.classList.contains("show");
  
  // Close all other dropdowns
  document.querySelectorAll(".dropdown-content.show").forEach(d => {
    if (d.id !== "colVisibilityMenu") d.classList.remove("show");
  });
  
  if (wasOpen) {
    dropdown.classList.remove("show");
  } else {
    dropdown.classList.add("show");
  }
}

function applyColumnVisibility() {
  let settings = {};
  try {
    settings = JSON.parse(localStorage.getItem("ali_raza_col_visibility")) || {};
  } catch (e) {
    settings = {};
  }
  
  const colCheckboxes = document.querySelectorAll(".col-toggle-checkbox");
  colCheckboxes.forEach((cb) => {
    const colId = cb.getAttribute("data-column");
    const isVisible = settings[colId] !== false;
    
    // Toggle all elements (headers, cells, mobile card fields) with data-col attribute
    const elements = document.querySelectorAll(`[data-col="${colId}"]`);
    elements.forEach(el => {
      el.style.display = isVisible ? "" : "none";
    });
  });

  // Clean up separators in mobile card headers (.market-lbl) based on neighboring field visibility
  document.querySelectorAll(".market-lbl").forEach(lbl => {
    const visibleSpans = Array.from(lbl.children).filter(span => {
      return !span.classList.contains("col-sep") && span.style.display !== "none";
    });
    
    // Hide all separators initially
    lbl.querySelectorAll(".col-sep").forEach(sep => {
      sep.style.display = "none";
    });
    
    // Show separator after each visible span, except the last one
    for (let i = 0; i < visibleSpans.length - 1; i++) {
      const currentSpan = visibleSpans[i];
      const nextSep = currentSpan.nextElementSibling;
      if (nextSep && nextSep.classList.contains("col-sep")) {
        nextSep.style.display = "inline";
      }
    }
  });

  // Resizer and dynamic table width adjustments have been disabled to let CSS handle layouts.
}

function getMinColumnWidth(colId) {
  if (colId === "col-name") return 260;
  if (colId === "col-contactPerson") return 140;
  if (colId === "col-contact") return 140;
  if (colId === "col-market") return 100;
  if (colId === "col-niche") return 120;
  if (colId === "col-source") return 90;
  if (colId === "col-priority") return 85;
  if (colId === "col-stage") return 110;
  if (colId === "col-nextAction") return 150;
  if (colId === "col-nextActionDate") return 130;
  if (colId === "col-replyStatus") return 125;
  if (colId === "col-notes") return 250;
  if (colId === "col-actions") return 120;
  return 50; // Checkbox column fallback
}

function getDefaultColumnWidth(colId) {
  if (colId === "col-name") return 260;
  if (colId === "col-contactPerson") return 140;
  if (colId === "col-contact") return 140;
  if (colId === "col-market") return 100;
  if (colId === "col-niche") return 120;
  if (colId === "col-source") return 90;
  if (colId === "col-priority") return 85;
  if (colId === "col-stage") return 110;
  if (colId === "col-nextAction") return 150;
  if (colId === "col-nextActionDate") return 130;
  if (colId === "col-replyStatus") return 125;
  if (colId === "col-notes") return 250;
  if (colId === "col-actions") return 120;
  return 100;
}

function adjustTableWidthToColumns(table) {
  if (!table) return;
  const ths = table.querySelectorAll("thead th");
  let totalWidth = 0;
  ths.forEach(th => {
    if (th.style.display !== "none") {
      const widthVal = parseFloat(th.style.width) || th.offsetWidth || 0;
      totalWidth += widthVal;
    }
  });
  
  const scrollContainer = table.parentElement;
  if (scrollContainer && scrollContainer.classList.contains("leads-table-scroll")) {
    const containerWidth = scrollContainer.clientWidth;
    if (containerWidth && totalWidth < containerWidth) {
      table.style.width = "100%";
      return;
    }
  }
  table.style.width = totalWidth + "px";
}

function initTableResizableColumns() {
  const table = document.querySelector("table.leads-table");
  if (!table) return;

  const ths = table.querySelectorAll("thead th");
  const colWidthsKey = "ali_raza_col_widths";
  let savedWidths = {};
  
  try {
    savedWidths = JSON.parse(localStorage.getItem(colWidthsKey)) || {};
  } catch (e) {
    savedWidths = {};
  }

  ths.forEach((th, idx) => {
    const colId = th.getAttribute("data-col") || `col-idx-${idx}`;
    
    // Apply saved width or initialize (lock first and actions columns)
    if (idx === 0) {
      th.style.width = "50px";
    } else if (colId === "col-actions") {
      th.style.width = "120px";
    } else if (savedWidths[colId]) {
      th.style.width = savedWidths[colId] + "px";
    } else {
      th.style.width = getDefaultColumnWidth(colId) + "px";
    }

    // Inject resizer handle (skip first checkbox column and last actions column)
    if (idx !== 0 && colId !== "col-actions") {
      if (!th.querySelector(".resizer")) {
        const resizer = document.createElement("div");
        resizer.className = "resizer";
        th.appendChild(resizer);

        // Drag event handling
        resizer.addEventListener("mousedown", function(e) {
          e.preventDefault();
          e.stopPropagation();
          
          const startX = e.pageX;
          const startWidth = th.offsetWidth;
          
          resizer.classList.add("resizing");
          document.body.style.cursor = "col-resize";
          document.body.style.userSelect = "none";

          function onMouseMove(moveEvent) {
            const dx = moveEvent.pageX - startX;
            const minWidth = getMinColumnWidth(colId);
            const newWidth = Math.max(minWidth, startWidth + dx);
            th.style.width = newWidth + "px";
            adjustTableWidthToColumns(table);
          }

          function onMouseUp() {
            resizer.classList.remove("resizing");
            document.body.style.cursor = "";
            document.body.style.userSelect = "";
            
            document.removeEventListener("mousemove", onMouseMove);
            document.removeEventListener("mouseup", onMouseUp);
            
            // Save column widths to localStorage
            try {
              const currentWidths = JSON.parse(localStorage.getItem(colWidthsKey)) || {};
              currentWidths[colId] = th.offsetWidth;
              localStorage.setItem(colWidthsKey, JSON.stringify(currentWidths));
            } catch (err) {
              console.error("Failed to save column widths:", err);
            }
          }

          document.addEventListener("mousemove", onMouseMove);
          document.addEventListener("mouseup", onMouseUp);
        });
      }
    }
  });

  adjustTableWidthToColumns(table);
}

function toggleColumnVisibility(colId, isVisible) {
  let settings = {};
  try {
    settings = JSON.parse(localStorage.getItem("ali_raza_col_visibility")) || {};
  } catch (e) {
    settings = {};
  }
  settings[colId] = isVisible;
  localStorage.setItem("ali_raza_col_visibility", JSON.stringify(settings));
  applyColumnVisibility();
}

function initColumnVisibility() {
  let settings = {};
  try {
    settings = JSON.parse(localStorage.getItem("ali_raza_col_visibility")) || {};
  } catch (e) {
    settings = {};
  }
  
  const colCheckboxes = document.querySelectorAll(".col-toggle-checkbox");
  colCheckboxes.forEach(cb => {
    const colId = cb.getAttribute("data-column");
    const isVisible = settings[colId] !== false;
    cb.checked = isVisible;
  });
  
  applyColumnVisibility();
}

// Authentication Status Toggle
function checkAuth() {
  const isLoggedIn = localStorage.getItem("ali_raza_logged_in") === "true";
  const loginContainer = document.getElementById("loginContainer");
  const appContainer = document.querySelector(".app-container");
  
  if (isLoggedIn) {
    if (loginContainer) loginContainer.classList.add("hidden");
    if (appContainer) appContainer.classList.remove("hidden");
  } else {
    if (loginContainer) loginContainer.classList.remove("hidden");
    if (appContainer) appContainer.classList.add("hidden");
  }
}

// Load Leads from LocalStorage or Seed Defaults
function loadData() {
  // Load Leads
  try {
    const stored = localStorage.getItem("ali_raza_leads");
    if (stored) {
      leads = JSON.parse(stored);
      // Ensure backwards compatibility for Source
      leads.forEach(lead => {
        if (!lead.source) lead.source = "Other";
      });
    } else {
      leads = [...DEFAULT_LEADS];
      saveData();
    }
  } catch (err) {
    console.error("Error parsing leads from localStorage, resetting to defaults", err);
    leads = [...DEFAULT_LEADS];
    saveData();
  }
}


// Save Leads to LocalStorage + fire-and-forget cloud sync
function saveData() {
  localStorage.setItem("ali_raza_leads", JSON.stringify(leads));
  syncLeadsToCloud(); // async, fire-and-forget
}

// Calculate and Render Dashboard Metrics
function updateDashboard() {
  const todayStr = getOffsetDateString(0);

  // Raw counts
  const total = leads.length;
  
  // Filter active leads (not archived)
  const activeLeads = leads.filter(l => l.stage !== 'Archived');

  const instagram = activeLeads.filter(l => l.channel === 'Instagram').length;
  const linkedin = activeLeads.filter(l => l.channel === 'LinkedIn').length;
  const email = activeLeads.filter(l => l.channel === 'Email').length;
  const whatsapp = activeLeads.filter(l => l.channel === 'WhatsApp').length;
  
  // Actions due today or overdue (and not archived)
  const todayActions = activeLeads.filter(l => l.nextActionDate <= todayStr).length;

  // Follow-ups due: Stage is Follow-up Due OR Next Action is Send follow-up
  const followUps = activeLeads.filter(l => l.stage === 'Follow-up Due' || l.nextAction === 'Send follow-up').length;

  const followupSentCount = activeLeads.filter(l => l.stage === 'First Follow-up Sent' || l.stage === 'Second Follow-up Sent').length;

  const aLeads = activeLeads.filter(l => l.priority === 'A').length;
  const warmLeads = activeLeads.filter(l => l.stage === 'Warm Lead').length;

  // Apply to DOM
  document.querySelector("#metricTotal .metric-value").textContent = total;
  document.querySelector("#metricInstagram .metric-value").textContent = instagram;
  document.querySelector("#metricLinkedin .metric-value").textContent = linkedin;
  document.querySelector("#metricEmail .metric-value").textContent = email;
  document.querySelector("#metricWhatsapp .metric-value").textContent = whatsapp;
  document.querySelector("#metricToday .metric-value").textContent = todayActions;
  document.querySelector("#metricFollowUps .metric-value").textContent = followUps;
  if (document.querySelector("#metricFollowupSent .metric-value")) {
    document.querySelector("#metricFollowupSent .metric-value").textContent = followupSentCount;
  }
  document.querySelector("#metricALeads .metric-value").textContent = aLeads;
  document.querySelector("#metricWarmLeads .metric-value").textContent = warmLeads;
}

// Unified Filter Logic
function getFilteredLeads() {
  const todayStr = getOffsetDateString(0);
  const searchQuery = document.getElementById("searchInput").value.trim().toLowerCase();
  
  // Advanced filter values
  const filterChannel = document.getElementById("filterChannel").value;
  const filterMarket = document.getElementById("filterMarket").value;
  const filterSource = document.getElementById("filterSource").value;
  const filterContactPerson = document.getElementById("filterContactPerson").value.trim().toLowerCase();
  const filterPriority = document.getElementById("filterPriority").value;
  const filterStage = document.getElementById("filterStage").value;
  const filterReply = document.getElementById("filterReply").value;
  const filterActionDate = document.getElementById("filterActionDate").value;
  const filterNextAction = document.getElementById("filterNextAction").value;

  const filtered = leads.filter((lead, index) => {
    // Keep track of index for operations
    lead.originalIndex = index;

    // 1. Channel Tab filtering (sync with tab nav)
    if (activeTab !== "All" && activeTab !== "LeadFinder" && activeTab !== "TodayMode" && lead.channel !== activeTab) return false;

    // 2. Search Query matching
    if (searchQuery) {
      const matchName = lead.name.toLowerCase().includes(searchQuery);
      const matchNiche = lead.niche.toLowerCase().includes(searchQuery);
      const matchNotes = lead.notes.toLowerCase().includes(searchQuery);
      const matchSource = lead.source && lead.source.toLowerCase().includes(searchQuery);
      const matchPerson = lead.contactPerson && lead.contactPerson.toLowerCase().includes(searchQuery);
      const matchEmail = lead.email && lead.email.toLowerCase().includes(searchQuery);
      const matchPhone = lead.whatsappNumber && lead.whatsappNumber.toLowerCase().includes(searchQuery);

      if (!matchName && !matchNiche && !matchNotes && !matchPerson && !matchEmail && !matchPhone && !matchSource) {
        return false;
      }
    }

    // 3. Quick Filters
    if (activeQuickFilter === "All") {
      // By default, hide archived leads in general list unless explicitly selecting "Archived" quick filter
      if (lead.stage === "Archived") return false;
    } else if (activeQuickFilter === "Today") {
      if (lead.nextActionDate > todayStr || lead.stage === "Archived") return false;
    } else if (activeQuickFilter === "FollowUp") {
      const isFollowUp = lead.stage === "Follow-up Due" || lead.nextAction === "Send follow-up";
      if (!isFollowUp || lead.stage === "Archived") return false;
    } else if (activeQuickFilter === "A") {
      if (lead.priority !== "A" || lead.stage === "Archived") return false;
    } else if (activeQuickFilter === "Warm") {
      if (lead.stage !== "Warm Lead" || lead.stage === "Archived") return false;
    } else if (activeQuickFilter === "Archived") {
      if (lead.stage !== "Archived") return false;
    }

    // 4. Advanced Filters (Dropdowns)
    if (filterChannel !== "All" && lead.channel !== filterChannel) return false;
    if (filterMarket !== "All" && lead.market !== filterMarket) return false;
    if (filterSource !== "All" && lead.source !== filterSource) return false;
    if (filterPriority !== "All" && lead.priority !== filterPriority) return false;
    if (filterStage !== "All" && lead.stage !== filterStage) return false;
    if (filterReply !== "All" && lead.replyStatus !== filterReply) return false;
    if (filterNextAction !== "All" && lead.nextAction !== filterNextAction) return false;
    
    // Action Date filter
    if (filterActionDate !== "All") {
      if (!lead.nextActionDate) return false;
      if (filterActionDate === "Overdue") {
        if (lead.nextActionDate > todayStr) return false;
      } else if (filterActionDate === "Today") {
        if (lead.nextActionDate !== todayStr) return false;
      } else if (filterActionDate === "Tomorrow") {
        const tomorrowStr = getOffsetDateString(1);
        if (lead.nextActionDate !== tomorrowStr) return false;
      } else if (filterActionDate === "ThisWeek") {
        const weekLaterStr = getOffsetDateString(7);
        if (lead.nextActionDate < todayStr || lead.nextActionDate > weekLaterStr) return false;
      } else if (filterActionDate === "NoDate") {
        return false; // has a date, so filter out
      }
    }

    // 5. Contact Person Filter
    if (filterContactPerson) {
      if (!lead.contactPerson) return false;
      if (!lead.contactPerson.toLowerCase().includes(filterContactPerson)) return false;
    }

    return true;
  });

  if (currentSortColumn === "contactPerson") {
    filtered.sort((a, b) => {
      const valA = (a.contactPerson || "").trim().toLowerCase();
      const valB = (b.contactPerson || "").trim().toLowerCase();
      if (!valA && !valB) return 0;
      if (!valA) return 1;
      if (!valB) return -1;
      if (valA < valB) return currentSortDirection === "asc" ? -1 : 1;
      if (valA > valB) return currentSortDirection === "asc" ? 1 : -1;
      return 0;
    });
  }

  return filtered;
}

// Generate Badges and styling elements
function getChannelBadge(channel) {
  const iconMap = {
    Instagram: `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><rect x="2" y="2" width="20" height="20" rx="5" ry="5"></rect><path d="M16 11.37A4 4 0 1 1 12.63 8 4 4 0 0 1 16 11.37z"></path><line x1="17.5" y1="6.5" x2="17.51" y2="6.5"></line></svg>`,
    LinkedIn: `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M16 8a6 6 0 0 1 6 6v7h-4v-7a2 2 0 0 0-2-2 2 2 0 0 0-2 2v7h-4v-7a6 6 0 0 1 6-6z"></path><rect x="2" y="9" width="4" height="12"></rect><circle cx="4" cy="4" r="2"></circle></svg>`,
    Email: `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"></path><polyline points="22,6 12,13 2,6"></polyline></svg>`,
    WhatsApp: `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 11.5a8.38 8.38 0 0 1-.9 3.8 8.5 8.5 0 0 1-7.6 4.7 8.38 8.38 0 0 1-3.8-.9L3 21l1.9-5.7a8.38 8.38 0 0 1-.9-3.8 8.5 8.5 0 0 1 4.7-7.6 8.38 8.38 0 0 1 3.8-.9h.5a8.48 8.48 0 0 1 8 8v.5z"></path></svg>`
  };
  return `<span class="badge badge-channel ${channel.toLowerCase()}">${iconMap[channel] || ''} ${channel}</span>`;
}

function getPriorityBadge(priority) {
  return `<span class="priority-marker p-${priority.toLowerCase()}" title="Priority ${priority}">${priority}</span>`;
}

function getStageBadge(stage) {
  const cssClass = stage.toLowerCase().replace(/ /g, '-');
  return `<span class="badge stage-badge s-${cssClass}">${stage}</span>`;
}

function updateStageHint(stage) {
  const hintEl = document.getElementById("leadStageHint");
  if (!hintEl) return;
  
  const hints = {
    "Found (Lead collected only)": "💡 <strong>Found</strong>: Lead collected only. No message sent yet.",
    "First Message Sent": "💡 <strong>First Message Sent</strong>: First outreach message/email/DM/WhatsApp sent. Waiting for first reply.",
    "First Follow-up Sent": "💡 <strong>First Follow-up Sent</strong>: First follow-up or clarification message sent after no reply.",
    "Second Follow-up Sent": "💡 <strong>Second Follow-up Sent</strong>: Second and final follow-up sent. After this, wait or archive.",
    "Engaged": "💡 <strong>Engaged</strong>: Client replied or showed basic interest, but no samples/test sent yet.",
    "Samples Sent": "💡 <strong>Samples Sent</strong>: CV, portfolio, samples, or example work sent. Waiting for review.",
    "Warm Lead": "💡 <strong>Warm Lead</strong>: Client asks for pricing, timeline, process, brief, outline, or gives a test project.",
    "Follow-up Due": "💡 <strong>Follow-up Due</strong>: Follow-up is needed now.",
    "Replied": "💡 <strong>Replied</strong>: Client replied and needs review/action.",
    "Not Now": "💡 <strong>Not Now</strong>: Client said not interested now, maybe later, or no current need.",
    "Archived": "💡 <strong>Archived</strong>: Dead, duplicate, irrelevant, wrong fit, bounced, or closed lead."
  };
  
  hintEl.innerHTML = hints[stage] || "";
}

function getReplyBadge(status) {
  const cssClass = status.toLowerCase().replace(/ /g, '-');
  return `<span class="reply-status-text ${cssClass}">${status}</span>`;
}

function renderNotesCellContent(notes, originalIndex) {
  if (!notes) return '<span class="notes-empty">-</span>';
  const cleanNotes = notes.trim();
  if (!cleanNotes) return '<span class="notes-empty">-</span>';
  
  return `<p class="notes-full-text">${escapeHtml(cleanNotes)}</p>`;
}

// Render Leads Grid (Table and Mobile Cards)
function renderLeads() {
  clearBulkSelection();
  const isLeadFinder = activeTab === "LeadFinder";
  const isTodayMode = activeTab === "TodayMode";
  const isStandardView = !isLeadFinder && !isTodayMode;

  // Toggle the tab action button container and text dynamically
  const tabActions = document.getElementById("tabActionsContainer");
  const btnText = document.getElementById("importLeadsBtnText");
  if (tabActions && btnText) {
    if (activeTab === "All" || isLeadFinder || isTodayMode) {
      tabActions.style.display = "none";
    } else {
      tabActions.style.display = "inline-flex";
      btnText.textContent = `Import ${activeTab} Leads`;
    }
  }

  const filtered = getFilteredLeads();
  const exportFilteredBtn = document.getElementById("exportFilteredLeadsBtn");
  if (exportFilteredBtn) {
    exportFilteredBtn.disabled = (filtered.length === 0);
  }
  const tableBody = document.getElementById("leadsTableBody");
  const cardsContainer = document.getElementById("leadsCardsContainer");
  const countBadge = document.getElementById("leadsCount");

  // Update headers / titles
  if (countBadge) {
    countBadge.textContent = (isLeadFinder || isTodayMode) ? "" : `${filtered.length} leads`;
  }
  if (document.getElementById("activeTabTitle")) {
    document.getElementById("activeTabTitle").textContent = isLeadFinder ? "Lead Finder" : (isTodayMode ? "Today Mode" : (activeTab === "All" ? "Dashboard" : `${activeTab} Leads`));
  }

  // Toggle Today's Action Center section visibility (only on Today Mode tab)
  const todaySection = document.getElementById("todayModeSection");
  if (todaySection) {
    todaySection.style.display = isTodayMode ? "block" : "none";
  }

  // Toggle Lead Finder section
  const finderSection = document.getElementById("leadFinderSection");
  if (finderSection) {
    finderSection.style.display = isLeadFinder ? "block" : "none";
  }

  // Toggle other dashboard/main view elements
  const storageNotice = document.getElementById("storageNoticeContainer");
  if (storageNotice) {
    storageNotice.style.display = isStandardView ? "block" : "none";
  }

  const dashGrid = document.querySelector(".dashboard-grid");
  if (dashGrid) dashGrid.style.display = isStandardView ? "grid" : "none";

  const searchBox = document.querySelector(".search-box");
  if (searchBox) searchBox.style.display = isStandardView ? "flex" : "none";

  const filterToggle = document.querySelector(".filters-toggle-row");
  if (filterToggle) filterToggle.style.display = isStandardView ? "flex" : "none";

  const advPanel = document.getElementById("advancedFiltersPanel");
  if (advPanel) advPanel.style.display = isStandardView ? "" : "none";

  const mainContent = document.querySelector(".main-content-section");
  if (mainContent) mainContent.style.display = isStandardView ? "block" : "none";

  // Update table header text dynamically
  const contactHeader = document.getElementById("dynamicContactHeader");
  if (contactHeader) {
    const textSpan = contactHeader.querySelector(".th-text") || contactHeader;
    if (activeTab === "Email") {
      textSpan.textContent = "Email Address";
    } else if (activeTab === "WhatsApp") {
      textSpan.textContent = "WhatsApp Number";
    } else if (activeTab === "Instagram" || activeTab === "LinkedIn") {
      textSpan.textContent = "Profile Link";
    } else {
      textSpan.textContent = "Channel";
    }
  }

  tableBody.innerHTML = "";
  cardsContainer.innerHTML = "";

  if (filtered.length === 0) {
    const emptyHtml = `
      <tr>
        <td colspan="15" style="text-align: center;">
          <div class="empty-state">
            <span class="empty-state-icon">\uD83D\uDD0D</span>
            <h3>No leads found</h3>
            <p>Try clearing your search query, adjusting your active filters, or adding a new lead.</p>
          </div>
        </td>
      </tr>
    `;
    tableBody.innerHTML = emptyHtml;

    cardsContainer.innerHTML = `
      <div class="empty-state">
        <span class="empty-state-icon">\uD83D\uDD0D</span>
        <h3>No leads found</h3>
        <p>Try clearing filters or adding a new lead.</p>
      </div>
    `;
    return;
  }

  // Loop & Render
  filtered.forEach(lead => {
    // Determine what to render in the second column based on activeTab
    let contactCellHtml = "";
    if (activeTab === "Email") {
      const email = lead.email ? String(lead.email).trim() : "";
      contactCellHtml = `<td data-col="col-contact">${email ? `<a href="mailto:${email}" style="font-weight: 600; color: var(--color-royal-blue);">${email}</a>` : '-'}</td>`;
    } else if (activeTab === "WhatsApp") {
      const phone = lead.whatsappNumber ? String(lead.whatsappNumber).trim() : "";
      const waLink = getWhatsAppLink(phone);
      contactCellHtml = `<td data-col="col-contact">
        <div style="font-weight: 600;">${phone || "-"}</div>
        ${phone ? `<a href="${waLink}" target="_blank" rel="noopener noreferrer" style="font-size: 11px; color: var(--color-teal-green); font-weight: 600; display: inline-flex; align-items: center; gap: 4px; margin-top: 2px;">
          Open WhatsApp <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="margin-left: 2px; display: inline-block; vertical-align: middle;"><path d="M15 3h6v6"/><path d="M10 14 21 3"/><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/></svg>
        </a>` : ''}
      </td>`;
    } else if (activeTab === "Instagram" || activeTab === "LinkedIn") {
      const link = lead.mainLink ? String(lead.mainLink).trim() : "";
      const normalized = normalizeUrl(link);
      const displayText = link ? (link.length > 30 ? link.substring(0, 30) + "..." : link) : "";
      contactCellHtml = `<td data-col="col-contact">${link ? `<a href="${normalized}" target="_blank" rel="noopener noreferrer" style="font-weight: 600; color: var(--color-royal-blue);">${displayText} <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="margin-left: 2px; display: inline-block; vertical-align: middle;"><path d="M15 3h6v6"/><path d="M10 14 21 3"/><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/></svg></a>` : '-'}</td>`;
    } else {
      contactCellHtml = `<td data-col="col-contact">${getChannelBadge(lead.channel)}</td>`;
    }

    // 1. Table Row (Desktop)
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td style="text-align: center;"><input type="checkbox" class="lead-checkbox" data-index="${lead.originalIndex}" onchange="updateSelectedLeadsCount()"></td>
      <td data-col="col-name">
        <div style="font-weight: 700; color: var(--color-deep-navy);">${lead.name || "Unnamed Lead / Company"}</div>
        ${lead.mainLink ? `<a href="${normalizeUrl(lead.mainLink)}" target="_blank" rel="noopener noreferrer" style="font-size: 11px; display: inline-flex; align-items: center; gap: 4px; margin-top: 2px;">
          Open Link <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="margin-left: 2px; display: inline-block; vertical-align: middle;"><path d="M15 3h6v6"/><path d="M10 14 21 3"/><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/></svg>
        </a>` : ''}
      </td>
      <td data-col="col-contactPerson">${lead.contactPerson ? escapeHtml(lead.contactPerson) : '—'}</td>
      ${contactCellHtml}
      <td data-col="col-market"><strong>${lead.market}</strong></td>
      <td data-col="col-niche">${lead.niche}</td>
      <td data-col="col-source"><span style="font-size: 12.5px; font-weight: 600; color: var(--color-priority-c);">${lead.source || "Other"}</span></td>
      <td data-col="col-priority">${getPriorityBadge(lead.priority)}</td>
      <td data-col="col-stage">${getStageBadge(lead.stage)}</td>
      <td data-col="col-nextAction" style="font-weight: 600;">${lead.nextAction}</td>
      <td data-col="col-nextActionDate" style="white-space: nowrap;">
        <span style="font-weight: 600; color: ${lead.nextActionDate && lead.nextActionDate <= getOffsetDateString(0) ? '#ef4444' : 'inherit'}">
          ${lead.nextActionDate || '-'}
        </span>
      </td>
      <td data-col="col-replyStatus">${getReplyBadge(lead.replyStatus)}</td>
      <td data-col="col-notes" class="col-notes-cell" title="${escapeHtml(lead.notes || '')}">
        ${renderNotesCellContent(lead.notes, lead.originalIndex)}
      </td>
      <td data-col="col-actions">
        <div class="action-buttons" style="display: flex; gap: 4px; align-items: center; justify-content: center;">
          ${getQuickActionsDropdownHtml(lead)}
          <button class="action-btn edit-btn" onclick="openEditModal(${lead.originalIndex})" title="Edit Lead">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 20h9M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg>
          </button>
          <button class="action-btn delete-btn" onclick="deleteLead(${lead.originalIndex})" title="Delete Lead">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path><line x1="10" y1="11" x2="10" y2="17"></line><line x1="14" y1="11" x2="14" y2="17"></line></svg>
          </button>
        </div>
      </td>
    `;
    tableBody.appendChild(tr);

    // 2. Card Stack (Mobile)
    const card = document.createElement("div");
    card.className = "lead-card-mobile";
    card.innerHTML = `
      <div class="lead-card-header" style="display: flex; align-items: center; justify-content: space-between; gap: 8px;">
        <input type="checkbox" class="lead-checkbox" data-index="${lead.originalIndex}" onchange="updateSelectedLeadsCount()" style="width: 16px; height: 16px; cursor: pointer; margin-right: 4px;">
        <div class="lead-card-title" style="flex: 1;">
          <h3 data-col="col-name">${lead.name || "Unnamed Lead / Company"}</h3>
          <span class="market-lbl">
            <span data-col="col-market">${lead.market || "-"}</span>
            <span class="col-sep" style="margin: 0 4px;">•</span>
            <span data-col="col-niche">${lead.niche || "-"}</span>
            <span class="col-sep" style="margin: 0 4px;">•</span>
            <span data-col="col-source">${lead.source || "Other"}</span>
          </span>
        </div>
        <div class="lead-card-badges">
          <span data-col="col-priority">${getPriorityBadge(lead.priority)}</span>
          <span data-col="col-contact">${getChannelBadge(lead.channel)}</span>
        </div>
      </div>
      
      <div class="lead-card-details-grid">
        <div class="lead-card-detail-item" data-col="col-contactPerson">
          <span class="lbl">Contact Person</span>
          <span class="val">${lead.contactPerson || '—'}</span>
        </div>
        <div class="lead-card-detail-item" data-col="col-stage">
          <span class="lbl">Stage</span>
          <span class="val">${getStageBadge(lead.stage)}</span>
        </div>
        <div class="lead-card-detail-item" data-col="col-replyStatus">
          <span class="lbl">Reply Status</span>
          <span class="val">${getReplyBadge(lead.replyStatus)}</span>
        </div>
        <div class="lead-card-detail-item" data-col="col-nextAction">
          <span class="lbl">Next Action</span>
          <span class="val" style="font-weight: 700;">${lead.nextAction}</span>
        </div>
        <div class="lead-card-detail-item" data-col="col-nextActionDate">
          <span class="lbl">Action Date</span>
          <span class="val" style="font-weight: 700; color: ${lead.nextActionDate && lead.nextActionDate <= getOffsetDateString(0) ? '#ef4444' : 'inherit'}">${lead.nextActionDate || '-'}</span>
        </div>
      </div>

      <div class="lead-card-notes" data-col="col-notes">
        <strong>Notes:</strong> ${lead.notes}
      </div>

      ${lead.contactPerson || lead.email || lead.whatsappNumber || lead.dateAdded || lead.lastActionDate ? `
      <div class="lead-card-notes" style="font-size: 11px; background-color: var(--color-off-white); padding: 6px; border-radius: var(--radius-sm);">
        ${lead.contactPerson ? `<span><strong>Contact:</strong> ${lead.contactPerson}<br></span>` : ''}
        ${lead.email ? `<span data-col="col-contact"><strong>Email:</strong> ${lead.email}<br></span>` : ''}
        ${lead.whatsappNumber ? `<span data-col="col-contact"><strong>WhatsApp:</strong> ${lead.whatsappNumber}<br></span>` : ''}
        ${lead.dateAdded ? `<span><strong>Added:</strong> ${lead.dateAdded} | </span>` : ''}
        ${lead.lastActionDate ? `<span><strong>Last Action:</strong> ${lead.lastActionDate}</span>` : ''}
      </div>` : ''}

      <div class="lead-card-actions" data-col="col-actions" style="display: flex; align-items: center; gap: 6px;">
        ${getQuickActionsDropdownHtml(lead)}
        ${lead.mainLink ? `<a href="${normalizeUrl(lead.mainLink)}" target="_blank" rel="noopener noreferrer" class="btn btn-secondary" style="padding: 6px 12px; font-size: 12px; margin-right: auto;">Open Link <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="margin-left: 2px; display: inline-block; vertical-align: middle;"><path d="M15 3h6v6"/><path d="M10 14 21 3"/><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/></svg></a>` : ''}
        <button class="btn btn-secondary btn-icon-only" onclick="openEditModal(${lead.originalIndex})" title="Edit Lead">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 20h9M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg>
        </button>
        <button class="btn btn-secondary btn-icon-only btn-danger-outline" onclick="deleteLead(${lead.originalIndex})" title="Delete Lead">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path><line x1="10" y1="11" x2="10" y2="17"></line><line x1="14" y1="11" x2="14" y2="17"></line></svg>
        </button>
      </div>
    `;
    cardsContainer.appendChild(card);
  });

  if (isTodayMode) {
    renderTodayActions();
  }
  
  if (typeof applyColumnVisibility === "function") {
    applyColumnVisibility();
  }
}

// Render "Today's Actions" section in the new Today Mode columns
function renderTodayActions() {
  const todayStr = getOffsetDateString(0);
  
  const todayModeCount = document.getElementById("todayModeCount");
  const firstMessagesCount = document.getElementById("firstMessagesCount");
  const followupsCount = document.getElementById("followupsCount");
  const warmCount = document.getElementById("warmCount");
  const overdueCount = document.getElementById("overdueCount");
  
  const firstList = document.getElementById("firstMessagesList");
  const followupsList = document.getElementById("followupsList");
  const warmList = document.getElementById("warmList");
  const overdueList = document.getElementById("overdueList");

  if (!firstList || !followupsList || !warmList || !overdueList) return;

  // Filter criteria: Next Action Date is today or earlier AND Stage is not Archived
  let todayLeads = leads
    .map((lead, index) => ({ ...lead, originalIndex: index }))
    .filter(lead => {
      if (!lead.nextActionDate) return false;
      if (lead.stage === "Archived") return false;
      const leadDate = lead.nextActionDate.substring(0, 10);
      return leadDate <= todayStr;
    });

  // Apply activeTodayFilter
  if (activeTodayFilter === "Email") {
    todayLeads = todayLeads.filter(l => l.channel === "Email");
  } else if (activeTodayFilter === "WhatsApp") {
    todayLeads = todayLeads.filter(l => l.channel === "WhatsApp");
  } else if (activeTodayFilter === "Instagram") {
    todayLeads = todayLeads.filter(l => l.channel === "Instagram");
  } else if (activeTodayFilter === "LinkedIn") {
    todayLeads = todayLeads.filter(l => l.channel === "LinkedIn");
  } else if (activeTodayFilter === "A") {
    todayLeads = todayLeads.filter(l => l.priority === "A");
  } else if (activeTodayFilter === "Warm") {
    todayLeads = todayLeads.filter(l => l.stage === "Warm Lead");
  } else if (activeTodayFilter === "Overdue") {
    todayLeads = todayLeads.filter(l => l.nextActionDate.substring(0, 10) < todayStr);
  }

  if (todayModeCount) todayModeCount.textContent = `${todayLeads.length} active`;

  // Helper to determine first message action type
  const isFirstAction = (action) => {
    const act = (action || "").toLowerCase();
    return act.includes("first") || 
           act.includes("dm") || 
           act.includes("connection") || 
           act.includes("email") || 
           act.includes("whatsapp") || 
           act.includes("like") || 
           (act.includes("follow") && !act.includes("follow-up") && !act.includes("followup"));
  };

  // Partition leads
  // 1. Overdue Actions: dates older than today
  const overdueActions = todayLeads.filter(l => l.nextActionDate.substring(0, 10) < todayStr);
  
  // Leads matching today's date
  const currentLeads = todayLeads.filter(l => l.nextActionDate.substring(0, 10) === todayStr);

  // 2. Send First Messages: today's date and is first contact action
  const firstMessages = currentLeads.filter(l => l.stage !== "Warm Lead" && isFirstAction(l.nextAction));
  
  // 3. Warm Leads: today's date and Stage is Warm Lead
  const warmLeads = currentLeads.filter(l => l.stage === "Warm Lead");
  
  // 4. Follow-ups Due: today's date and not first contact action
  const followups = currentLeads.filter(l => l.stage !== "Warm Lead" && !isFirstAction(l.nextAction));

  if (firstMessagesCount) firstMessagesCount.textContent = firstMessages.length;
  if (followupsCount) followupsCount.textContent = followups.length;
  if (warmCount) warmCount.textContent = warmLeads.length;
  if (overdueCount) overdueCount.textContent = overdueActions.length;

  // Helper to render workflow cards
  const renderColumnList = (container, listLeads, columnType) => {
    container.innerHTML = "";
    if (listLeads.length === 0) {
      let emptyMsg = "No actions pending";
      if (columnType === "first") emptyMsg = "No new messages to send";
      else if (columnType === "followup") emptyMsg = "No follow-ups due";
      else if (columnType === "warm") emptyMsg = "No warm leads pending";
      else if (columnType === "overdue") emptyMsg = "No overdue actions pending";
      
      container.innerHTML = `
        <div style="text-align: center; padding: 16px; color: var(--color-priority-c); font-size: 12px; background: rgba(11, 31, 58, 0.02); border-radius: var(--radius-sm); border: 1px dashed rgba(11, 31, 58, 0.06);">
          ${emptyMsg}
        </div>
      `;
      return;
    }

    listLeads.forEach(lead => {
      const card = document.createElement("div");
      card.className = "workflow-card";
      
      // Determine contact detail
      let contactHtml = "";
      if (lead.channel === "Email") {
        contactHtml = `<div><strong>Email:</strong> ${lead.email ? `<a href="mailto:${lead.email}">${lead.email}</a>` : '-'}</div>`;
      } else if (lead.channel === "WhatsApp") {
        contactHtml = `<div><strong>WhatsApp:</strong> ${lead.whatsappNumber || '-'}</div>`;
      } else {
        const linkText = lead.mainLink ? (lead.mainLink.length > 25 ? lead.mainLink.substring(0, 25) + "..." : lead.mainLink) : "-";
        contactHtml = `<div><strong>Profile:</strong> ${lead.mainLink ? `<a href="${normalizeUrl(lead.mainLink)}" target="_blank">${linkText}</a>` : '-'}</div>`;
      }

      // Generate buttons based on channel
      let buttonsHtml = "";
      if (lead.channel === "Email") {
        buttonsHtml = `
          <button class="btn btn-primary" onclick="markSentEmail(${lead.originalIndex})" title="Mark email sent & reschedule follow-up"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg> Mark Sent</button>
          <button class="btn btn-secondary" onclick="openLeadLink(${lead.originalIndex})" title="Open lead's website/link"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 3h6v6"/><path d="M10 14 21 3"/><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/></svg> Open Link</button>
          <button class="btn btn-secondary" onclick="setFollowupCalendar(${lead.originalIndex})" title="Reschedule next action date"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="18" height="18" x="3" y="4" rx="2" ry="2"/><line x1="16" x2="16" y1="2" y2="6"/><line x1="8" x2="8" y1="2" y2="6"/><line x1="3" x2="21" y1="10" y2="10"/></svg> Reschedule</button>
          <button class="btn btn-danger-outline" onclick="archiveLead(${lead.originalIndex})" title="Archive lead"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="20" height="5" x="2" y="3" rx="1"/><path d="M4 8v11a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8"/><line x1="10" x2="14" y1="12" y2="12"/></svg> Archive</button>
        `;
      } else if (lead.channel === "WhatsApp") {
        buttonsHtml = `
          <button class="btn btn-secondary" onclick="openWhatsAppChat(${lead.originalIndex})" title="Open click-to-chat WhatsApp link"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M7.9 20A9 9 0 1 0 4 16.1L2 22Z"/></svg> Open WA</button>
          <button class="btn btn-primary" onclick="markSentWhatsApp(${lead.originalIndex})" title="Mark WhatsApp sent & reschedule follow-up"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg> Mark Sent</button>
          <button class="btn btn-secondary" onclick="setFollowupCalendar(${lead.originalIndex})" title="Reschedule next action date"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="18" height="18" x="3" y="4" rx="2" ry="2"/><line x1="16" x2="16" y1="2" y2="6"/><line x1="8" x2="8" y1="2" y2="6"/><line x1="3" x2="21" y1="10" y2="10"/></svg> Reschedule</button>
          <button class="btn btn-danger-outline" onclick="archiveLead(${lead.originalIndex})" title="Archive lead"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="20" height="5" x="2" y="3" rx="1"/><path d="M4 8v11a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8"/><line x1="10" x2="14" y1="12" y2="12"/></svg> Archive</button>
        `;
      } else if (lead.channel === "Instagram") {
        buttonsHtml = `
          <button class="btn btn-secondary" onclick="openLeadLink(${lead.originalIndex})" title="Open Instagram profile in new tab"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 3h6v6"/><path d="M10 14 21 3"/><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/></svg> Open IG</button>
          <button class="btn btn-secondary" onclick="markInstagramCommented(${lead.originalIndex})" title="Log a comment action on lead's post"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg> Commented</button>
          <button class="btn btn-secondary" onclick="markInstagramFollowed(${lead.originalIndex})" title="Log a followed action on lead's profile"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><line x1="19" x2="19" y1="8" y2="14"/><line x1="22" x2="16" y1="11" y2="11"/></svg> Followed</button>
          <button class="btn btn-primary" onclick="markSentDM(${lead.originalIndex})" title="Mark Instagram DM sent & reschedule follow-up"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg> DM Sent</button>
          <button class="btn btn-danger-outline" onclick="archiveLead(${lead.originalIndex})" title="Archive lead"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="20" height="5" x="2" y="3" rx="1"/><path d="M4 8v11a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8"/><line x1="10" x2="14" y1="12" y2="12"/></svg> Archive</button>
        `;
      } else if (lead.channel === "LinkedIn") {
        buttonsHtml = `
          <button class="btn btn-secondary" onclick="openLeadLink(${lead.originalIndex})" title="Open LinkedIn profile in new tab"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 3h6v6"/><path d="M10 14 21 3"/><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/></svg> Open LI</button>
          <button class="btn btn-secondary" onclick="markLinkedInConnectionSent(${lead.originalIndex})" title="Mark connection request sent & wait 3 days"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><line x1="19" x2="19" y1="8" y2="14"/><line x1="22" x2="16" y1="11" y2="11"/></svg> Conn Sent</button>
          <button class="btn btn-primary" onclick="markSentLinkedInMessage(${lead.originalIndex})" title="Mark message sent & reschedule follow-up"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg> Msg Sent</button>
          <button class="btn btn-danger-outline" onclick="archiveLead(${lead.originalIndex})" title="Archive lead"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="20" height="5" x="2" y="3" rx="1"/><path d="M4 8v11a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8"/><line x1="10" x2="14" y1="12" y2="12"/></svg> Archive</button>
        `;
      }

      card.innerHTML = `
        <div class="workflow-card-header">
          <div class="workflow-card-name" title="${escapeHtml(lead.name)}">${escapeHtml(lead.name)}</div>
          <div class="workflow-card-badges">
            ${getPriorityBadge(lead.priority)}
            ${getChannelBadge(lead.channel)}
          </div>
        </div>
        <div class="workflow-card-details">
          <div><strong>Niche:</strong> ${escapeHtml(lead.niche || '-')} (${escapeHtml(lead.market || '-')})</div>
          <div><strong>Next Action:</strong> <span style="color: var(--color-royal-blue); font-weight: 700;">${escapeHtml(lead.nextAction || '-')}</span></div>
          ${contactHtml}
        </div>
        ${lead.notes ? `<div class="workflow-card-notes" title="Notes"><strong>Notes:</strong> ${escapeHtml(lead.notes)}</div>` : ''}
        <div class="workflow-card-buttons">
          ${buttonsHtml}
        </div>
      `;
      container.appendChild(card);
    });
  };

  renderColumnList(firstList, firstMessages, "first");
  renderColumnList(followupsList, followups, "followup");
  renderColumnList(warmList, warmLeads, "warm");
  renderColumnList(overdueList, overdueActions, "overdue");
}

// Modal open: Add Mode
function openAddModal() {
  const form = document.getElementById("leadForm");
  form.reset();

  const todayStr = getOffsetDateString(0);
  document.getElementById("leadDateAdded").value = todayStr;
  document.getElementById("leadLastActionDate").value = todayStr;
  document.getElementById("leadSource").value = "Google";
  document.getElementById("leadNextActionDate").value = getOffsetDateString(1); // Default to tomorrow

  document.getElementById("leadIndex").value = "";
  document.getElementById("modalTitle").textContent = "Add Outreach Lead";
  updateStageHint("Found (Lead collected only)"); // Initialize default stage hint
  document.getElementById("leadModal").classList.add("active");
}

// Modal open: Edit Mode
function openEditModal(index) {
  const lead = leads[index];
  
  document.getElementById("leadIndex").value = index;
  document.getElementById("leadDateAdded").value = lead.dateAdded || "";
  document.getElementById("leadName").value = lead.name || "";
  document.getElementById("leadMarket").value = lead.market || "";
  document.getElementById("leadChannel").value = lead.channel || "Instagram";
  document.getElementById("leadMainLink").value = lead.mainLink || "";
  document.getElementById("leadNiche").value = lead.niche || "";
  document.getElementById("leadSource").value = lead.source || "Other";
  document.getElementById("leadPriority").value = lead.priority || "A";
  document.getElementById("leadStage").value = lead.stage || "Found (Lead collected only)";
  document.getElementById("leadLastActionDate").value = lead.lastActionDate || "";
  document.getElementById("leadNextAction").value = lead.nextAction || "Send DM";
  document.getElementById("leadNextActionDate").value = lead.nextActionDate || "";
  document.getElementById("leadReplyStatus").value = lead.replyStatus || "No reply";
  document.getElementById("leadNotes").value = lead.notes || "";

  // Optional Fields
  document.getElementById("leadContactPerson").value = lead.contactPerson || "";
  document.getElementById("leadEmail").value = lead.email || "";
  document.getElementById("leadWhatsappNumber").value = lead.whatsappNumber || "";
  document.getElementById("leadExtraLink").value = lead.extraLink || "";
  document.getElementById("leadFollowUpCount").value = lead.followUpCount || 0;
  document.getElementById("leadMessageSent").value = lead.messageSent || "";

  // Make sure details drawer is closed or reset
  const details = document.querySelector("details");
  if (details) details.removeAttribute("open");

  document.getElementById("modalTitle").textContent = "Edit Outreach Lead";
  updateStageHint(lead.stage || "Found (Lead collected only)"); // Load selected stage hint
  document.getElementById("leadModal").classList.add("active");
}

// Close Modal
function closeModal() {
  document.getElementById("leadModal").classList.remove("active");
}

// Archive Lead (quick toggle)
function archiveLead(index) {
  const leadName = leads[index].name;
  leads[index].stage = "Archived";
  leads[index].nextAction = "Archive";
  
  saveData();
  updateDashboard();
  renderLeads();
  renderTodayActions();
  showToast(`Archived lead "${leadName}"`, "success");
}

// Delete Lead (permanently)
function deleteLead(index) {
  const leadName = leads[index].name;
  const confirmed = confirm("Are you sure you want to permanently delete this lead? This action cannot be undone.");
  if (confirmed) {
    leads.splice(index, 1);
    saveData();
    updateDashboard();
    renderLeads();
    renderTodayActions();
    showToast(`Permanently deleted lead "${leadName}"`, "success");
  }
}

// Toast System
function showToast(message, type = "success") {
  const container = document.getElementById("toastContainer");
  const toast = document.createElement("div");
  toast.className = `toast ${type}`;
  toast.innerHTML = `
    <span>${type === 'success' ? 'Γ£ô' : 'Γä╣'}</span>
    <span>${message}</span>
  `;
  container.appendChild(toast);

  // Auto remove after 3s
  setTimeout(() => {
    toast.style.animation = 'toast-in 0.3s cubic-bezier(0.4, 0, 0.2, 1) reverse forwards';
    setTimeout(() => {
      toast.remove();
    }, 300);
  }, 3000);
}

// Setup Event Listeners
function setupEventListeners() {
  const addBtn = document.getElementById("addLeadBtn");
  const closeBtn = document.getElementById("closeModalBtn");
  const cancelBtn = document.getElementById("cancelModalBtn");
  const form = document.getElementById("leadForm");
  const searchInput = document.getElementById("searchInput");
  const exportBtn = document.getElementById("exportCsvBtn");
  
  // Login form submission
  const loginForm = document.getElementById("loginForm");
  const loginError = document.getElementById("loginError");
  if (loginForm) {
    loginForm.addEventListener("submit", (e) => {
      e.preventDefault();
      console.log("Login attempted");
      const usernameInput = document.getElementById("loginUsername").value.trim();
      const passwordInput = document.getElementById("loginPassword").value.trim();
      
      if (usernameInput === LOGIN_USER && passwordInput === LOGIN_PASS) {
        console.log("Login successful");
        localStorage.setItem("ali_raza_logged_in", "true");
        if (loginError) loginError.classList.remove("active");
        checkAuth();
        showToast("Welcome back, Ali Raza!", "success");
      } else {
        console.log("Login failed");
        if (loginError) {
          loginError.textContent = "Invalid username or password.";
          loginError.classList.add("active");
        }
      }
    });
  }

  // Logout button click
  const logoutBtn = document.getElementById("logoutBtn");
  if (logoutBtn) {
    logoutBtn.addEventListener("click", () => {
      const confirmed = confirm("Are you sure you want to log out?");
      if (confirmed) {
        localStorage.removeItem("ali_raza_logged_in");
        if (loginForm) loginForm.reset();
        if (loginError) loginError.classList.remove("active");
        checkAuth();
        showToast("Logged out successfully.", "info");
      }
    });
  }

  // Modal toggles
  addBtn.addEventListener("click", openAddModal);
  closeBtn.addEventListener("click", closeModal);
  cancelBtn.addEventListener("click", closeModal);

  // Stage tooltips & dynamic auto-transitions
  const leadStageSelect = document.getElementById("leadStage");
  const leadReplyStatusSelect = document.getElementById("leadReplyStatus");
  const leadNextActionSelect = document.getElementById("leadNextAction");

  if (leadStageSelect) {
    leadStageSelect.addEventListener("change", () => {
      const stageVal = leadStageSelect.value;
      updateStageHint(stageVal);
      
      if (stageVal === "First Message Sent") {
        const replyStatusEl = document.getElementById("leadReplyStatus");
        if (replyStatusEl) replyStatusEl.value = "No reply";
        const nextActionEl = document.getElementById("leadNextAction");
        if (nextActionEl) nextActionEl.value = "Wait";
        showToast("Stage changed to 'First Message Sent'. Reply status set to 'No reply', Next Action set to 'Wait'.", "info");
      } else if (stageVal === "First Follow-up Sent" || stageVal === "Second Follow-up Sent") {
        const replyStatusEl = document.getElementById("leadReplyStatus");
        if (replyStatusEl) replyStatusEl.value = "No reply";
        const nextActionEl = document.getElementById("leadNextAction");
        if (nextActionEl) nextActionEl.value = "Wait";
        const followUpCountEl = document.getElementById("leadFollowUpCount");
        if (followUpCountEl) {
          const currentCount = parseInt(followUpCountEl.value) || 0;
          followUpCountEl.value = currentCount + 1;
        }
        showToast(`Stage changed to '${stageVal}'. Reply status set to 'No reply', Next Action set to 'Wait', and Follow-up Count incremented.`, "info");
      }
    });
  }

  if (leadReplyStatusSelect && leadStageSelect) {
    leadReplyStatusSelect.addEventListener("change", () => {
      const status = leadReplyStatusSelect.value;
      const currentStage = leadStageSelect.value;
      
      // Rule 1: If client only asks for CV/portfolio/samples, stage = Engaged
      if (status === "CV requested" || status === "Samples requested") {
        if (currentStage === "Found (Lead collected only)" || currentStage === "First Message Sent" || currentStage === "First Follow-up Sent" || currentStage === "Second Follow-up Sent") {
          leadStageSelect.value = "Engaged";
          updateStageHint("Engaged");
        }
      }
      
      // Rule 5: If client asks price, deadline, process, or sends brief/outline, stage = Warm Lead
      if (status === "Price asked") {
        if (currentStage !== "Archived") {
          leadStageSelect.value = "Warm Lead";
          updateStageHint("Warm Lead");
        }
      }

      // If client replies, set stage to Engaged or Replied
      if (status === "Replied" || status === "Interested") {
        if (currentStage === "Found (Lead collected only)" || currentStage === "First Message Sent" || currentStage === "First Follow-up Sent" || currentStage === "Second Follow-up Sent") {
          leadStageSelect.value = (status === "Interested") ? "Engaged" : "Replied";
          updateStageHint(leadStageSelect.value);
        }
      }
    });
  }

  if (leadNextActionSelect && leadStageSelect) {
    leadNextActionSelect.addEventListener("change", () => {
      const nextAction = leadNextActionSelect.value;
      const currentStage = leadStageSelect.value;
      
      // Rule 3: If client gives a test project, stage = Warm Lead
      if (nextAction === "Do test project" || nextAction === "Prepare proposal" || nextAction === "Send proposal") {
        if (currentStage !== "Archived") {
          leadStageSelect.value = "Warm Lead";
          updateStageHint("Warm Lead");
        }
      } else if (nextAction === "Submit test project") {
        // Rule 4: After test project is submitted, keep stage = Warm Lead and set Next Action = Wait for feedback
        leadNextActionSelect.value = "Wait for feedback";
        if (currentStage !== "Archived") {
          leadStageSelect.value = "Warm Lead";
          updateStageHint("Warm Lead");
        }
        showToast("Test project submitted! Action set to 'Wait for feedback'.", "info");
      } else if (nextAction === "Wait for feedback") {
        if (currentStage !== "Archived") {
          leadStageSelect.value = "Warm Lead";
          updateStageHint("Warm Lead");
        }
      }
    });
  }
  
  // Close on overlay click
  document.getElementById("leadModal").addEventListener("click", (e) => {
    if (e.target.id === "leadModal") closeModal();
  });

  // Save / Form Submit
  form.addEventListener("submit", (e) => {
    e.preventDefault();
    
    const indexVal = document.getElementById("leadIndex").value;
    
    const leadData = {
      dateAdded: document.getElementById("leadDateAdded").value,
      name: document.getElementById("leadName").value.trim(),
      market: document.getElementById("leadMarket").value,
      channel: document.getElementById("leadChannel").value,
      mainLink: normalizeUrl(document.getElementById("leadMainLink").value),
      niche: document.getElementById("leadNiche").value.trim(),
      source: document.getElementById("leadSource").value,
      priority: document.getElementById("leadPriority").value,
      stage: document.getElementById("leadStage").value,
      lastActionDate: document.getElementById("leadLastActionDate").value,
      nextAction: document.getElementById("leadNextAction").value,
      nextActionDate: document.getElementById("leadNextActionDate").value,
      replyStatus: document.getElementById("leadReplyStatus").value,
      notes: document.getElementById("leadNotes").value.trim(),
      
      // Optional fields
      contactPerson: document.getElementById("leadContactPerson").value.trim(),
      email: document.getElementById("leadEmail").value.trim(),
      whatsappNumber: document.getElementById("leadWhatsappNumber").value.trim(),
      extraLink: normalizeUrl(document.getElementById("leadExtraLink").value),
      followUpCount: parseInt(document.getElementById("leadFollowUpCount").value) || 0,
      messageSent: document.getElementById("leadMessageSent").value.trim()
    };

    if (indexVal === "") {
      // Add mode
      leads.push(leadData);
      showToast(`Added lead "${leadData.name}"`, "success");
    } else {
      // Edit mode
      const idx = parseInt(indexVal);
      leads[idx] = leadData;
      showToast(`Updated lead "${leadData.name}"`, "success");
    }

    saveData();
    closeModal();
    updateDashboard();
    renderLeads();
    renderTodayActions();
  });

  // Search input typing
  searchInput.addEventListener("input", () => {
    renderLeads();
  });

  // Contact Person input typing with 100ms debounce
  const filterContactPersonInput = document.getElementById("filterContactPerson");
  const clearFilterContactPersonBtn = document.getElementById("clearFilterContactPerson");
  if (filterContactPersonInput && clearFilterContactPersonBtn) {
    let cpTimeout;
    filterContactPersonInput.addEventListener("input", () => {
      if (filterContactPersonInput.value.trim() !== "") {
        clearFilterContactPersonBtn.style.display = "block";
      } else {
        clearFilterContactPersonBtn.style.display = "none";
      }
      clearTimeout(cpTimeout);
      cpTimeout = setTimeout(() => {
        renderLeads();
      }, 100);
    });

    clearFilterContactPersonBtn.addEventListener("click", () => {
      filterContactPersonInput.value = "";
      clearFilterContactPersonBtn.style.display = "none";
      renderLeads();
    });
  }

  // Tabs navigation
  const tabBtns = document.querySelectorAll(".tab-btn");
  tabBtns.forEach(btn => {
    btn.addEventListener("click", () => {
      tabBtns.forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
      
      activeTab = btn.getAttribute("data-channel");
      
      // Sync with advanced filter select channel
      const filterChanSelect = document.getElementById("filterChannel");
      if (activeTab === "All") {
        filterChanSelect.value = "All";
      } else {
        filterChanSelect.value = activeTab;
      }

      renderLeads();
    });
  });

  // Sync advanced channel dropdown to tabs if user edits inside advanced panel
  document.getElementById("filterChannel").addEventListener("change", (e) => {
    const val = e.target.value;
    activeTab = val;
    tabBtns.forEach(btn => {
      btn.classList.remove("active");
      if (btn.getAttribute("data-channel") === val) {
        btn.classList.add("active");
      }
    });
    renderLeads();
  });

  // Connect remaining advanced filter drop downs
  ["filterMarket", "filterSource", "filterPriority", "filterStage", "filterReply", "filterActionDate", "filterNextAction"].forEach(id => {
    document.getElementById(id).addEventListener("change", () => {
      renderLeads();
    });
  });

  // Advanced Filters toggle collapse
  const toggleFiltersBtn = document.getElementById("toggleFiltersBtn");
  const advancedPanel = document.getElementById("advancedFiltersPanel");
  toggleFiltersBtn.addEventListener("click", () => {
    const isCollapsed = advancedPanel.classList.toggle("collapsed");
    toggleFiltersBtn.querySelector("span").textContent = isCollapsed ? "Show Filters" : "Hide Filters";
  });

  // Quick Filters toggles
  const quickFilters = [
    { id: "qFilterAll", value: "All" },
    { id: "qFilterToday", value: "Today" },
    { id: "qFilterFollowUp", value: "FollowUp" },
    { id: "qFilterA", value: "A" },
    { id: "qFilterWarm", value: "Warm" },
    { id: "qFilterArchived", value: "Archived" }
  ];

  quickFilters.forEach(qf => {
    const btn = document.getElementById(qf.id);
    btn.addEventListener("click", () => {
      quickFilters.forEach(x => document.getElementById(x.id).classList.remove("active"));
      btn.classList.add("active");
      activeQuickFilter = qf.value;
      
      // If user wants archived, sync the advanced filter dropdown stage to Archived as well to make it intuitive
      const filterStageSelect = document.getElementById("filterStage");
      if (activeQuickFilter === "Archived") {
        filterStageSelect.value = "Archived";
      } else {
        // Reset advanced stage filter if switching back
        if (filterStageSelect.value === "Archived") {
          filterStageSelect.value = "All";
        }
      }

      renderLeads();
    });
  });
  // CSV Export trigger

  const exportFilteredLeadsBtn = document.getElementById("exportFilteredLeadsBtn");
  if (exportFilteredLeadsBtn) {
    exportFilteredLeadsBtn.addEventListener("click", exportFilteredLeads);
  }

  const exportAllLeadsBtn = document.getElementById("exportAllLeadsBtn");
  if (exportAllLeadsBtn) {
    exportAllLeadsBtn.addEventListener("click", exportAllLeads);
  }

  // --- Backup JSON System Event Listeners (wired directly to the visible header buttons) ---
  const exportBackupJsonBtn = document.getElementById("exportBackupJsonBtn");
  if (exportBackupJsonBtn) {
    exportBackupJsonBtn.addEventListener("click", exportBackupJSON);
  }

  const importBackupJsonBtn = document.getElementById("importBackupJsonBtn");
  if (importBackupJsonBtn) {
    importBackupJsonBtn.addEventListener("click", () => {
      const fileInput = document.getElementById("backupFileInput");
      if (fileInput) fileInput.click();
    });
  }


  const backupFileInput = document.getElementById("backupFileInput");
  if (backupFileInput) {
    backupFileInput.addEventListener("change", handleBackupFileSelect);
  }

  const cancelBackupImportBtn = document.getElementById("cancelBackupImportBtn");
  if (cancelBackupImportBtn) {
    cancelBackupImportBtn.addEventListener("click", closeBackupImportModal);
  }

  const closeBackupImportModalBtn = document.getElementById("closeBackupImportModalBtn");
  if (closeBackupImportModalBtn) {
    closeBackupImportModalBtn.addEventListener("click", closeBackupImportModal);
  }

  const confirmBackupImportBtn = document.getElementById("confirmBackupImportBtn");
  if (confirmBackupImportBtn) {
    confirmBackupImportBtn.addEventListener("click", confirmBackupImport);
  }

  const backupImportPreviewModal = document.getElementById("backupImportPreviewModal");
  if (backupImportPreviewModal) {
    backupImportPreviewModal.addEventListener("click", (e) => {
      if (e.target.id === "backupImportPreviewModal") {
        closeBackupImportModal();
      }
    });
  }

  const backModeMerge = document.getElementById("backModeMerge");
  if (backModeMerge) {
    backModeMerge.addEventListener("change", updateBackupPreviewCounts);
  }

  const backModeReplace = document.getElementById("backModeReplace");
  if (backModeReplace) {
    backModeReplace.addEventListener("change", updateBackupPreviewCounts);
  }

  const backImportDuplicatesAnyway = document.getElementById("backImportDuplicatesAnyway");
  if (backImportDuplicatesAnyway) {
    backImportDuplicatesAnyway.addEventListener("change", updateBackupPreviewCounts);
  }

  // Import Leads Button
  const importLeadsBtn = document.getElementById("importLeadsBtn");
  if (importLeadsBtn) {
    importLeadsBtn.addEventListener("click", () => {
      const fileInput = document.getElementById("importFileInput");
      if (fileInput) fileInput.click();
    });
  }

  // Import manual file change listener
  const importFileInput = document.getElementById("importFileInput");
  if (importFileInput) {
    importFileInput.addEventListener("change", (e) => {
      const file = e.target.files[0];
      if (!file) return;
      
      const reader = new FileReader();
      reader.onload = (event) => {
        const arrayBuffer = event.target.result;
        processImportData(arrayBuffer, file.name);
      };
      reader.readAsArrayBuffer(file);
    });
  }
  
  // Close Import Modal
  const closeImportModalBtn = document.getElementById("closeImportModalBtn");
  if (closeImportModalBtn) {
    closeImportModalBtn.addEventListener("click", closeImportPreview);
  }
  
  const cancelImportBtn = document.getElementById("cancelImportBtn");
  if (cancelImportBtn) {
    cancelImportBtn.addEventListener("click", closeImportPreview);
  }
  
  const confirmImportBtn = document.getElementById("confirmImportBtn");
  if (confirmImportBtn) {
    confirmImportBtn.addEventListener("click", confirmImport);
  }

  // Overlay click to close
  const importPreviewModal = document.getElementById("importPreviewModal");
  if (importPreviewModal) {
    importPreviewModal.addEventListener("click", (e) => {
      if (e.target.id === "importPreviewModal") closeImportPreview();
    });
  }



  // --- Lead Finder Event Listeners ---
  const finderMarket = document.getElementById("finderMarket");
  const finderTargetType = document.getElementById("finderTargetType");
  const finderKeyword = document.getElementById("finderKeyword");
  
  if (finderMarket) {
    finderMarket.addEventListener("change", () => {
      const captureMarket = document.getElementById("captureMarket");
      if (captureMarket) captureMarket.value = finderMarket.value;
      updateChipsVisibility();
      generateSearchLinks();
    });
  }
  if (finderTargetType) finderTargetType.addEventListener("change", generateSearchLinks);
  if (finderKeyword) finderKeyword.addEventListener("input", generateSearchLinks);

  document.querySelectorAll(".keyword-chip").forEach(chip => {
    chip.addEventListener("click", () => {
      document.querySelectorAll(".keyword-chip").forEach(c => c.classList.remove("active"));
      chip.classList.add("active");
      
      if (finderKeyword) finderKeyword.value = chip.getAttribute("data-keyword");
      
      const captureMarket = document.getElementById("captureMarket");
      if (captureMarket && finderMarket) captureMarket.value = finderMarket.value;
      
      generateSearchLinks();
    });
  });

  const quickCaptureForm = document.getElementById("quickCaptureForm");
  if (quickCaptureForm) {
    quickCaptureForm.addEventListener("submit", (e) => {
      e.preventDefault();
      
      const emailVal = document.getElementById("captureEmail").value.trim().toLowerCase();
      const phoneVal = document.getElementById("captureWhatsApp").value.trim().replace(/[\s\+\(\)\-\[\]]/g, "");
      const mainLinkVal = normalizeUrl(document.getElementById("captureMainLink").value).toLowerCase();
      const igLinkVal = normalizeUrl(document.getElementById("captureIgLink").value).toLowerCase();
      const liLinkVal = normalizeUrl(document.getElementById("captureLiLink").value).toLowerCase();
      const webLinkVal = normalizeUrl(document.getElementById("captureWebsite").value).toLowerCase();
      
      let isDuplicate = false;
      for (let l of leads) {
        if (emailVal && l.email && l.email.trim().toLowerCase() === emailVal) {
          isDuplicate = true;
          break;
        }
        if (phoneVal) {
          const existingPhone = (l.whatsappNumber || "").trim().replace(/[\s\+\(\)\-\[\]]/g, "");
          if (existingPhone && existingPhone === phoneVal) {
            isDuplicate = true;
            break;
          }
        }
        if (mainLinkVal && l.mainLink && l.mainLink.trim().toLowerCase() === mainLinkVal) {
          isDuplicate = true;
          break;
        }
        const linksToCheck = [mainLinkVal, igLinkVal, liLinkVal, webLinkVal].filter(Boolean);
        const existingLinks = [l.mainLink, l.extraLink].map(x => (x || "").trim().toLowerCase()).filter(Boolean);
        for (let link of linksToCheck) {
          if (existingLinks.includes(link)) {
            isDuplicate = true;
            break;
          }
        }
        if (isDuplicate) break;
      }
      
      if (isDuplicate) {
        const proceed = confirm("A lead with matching contact details or link already exists in the tracker. Do you still want to save this new lead?");
        if (!proceed) return;
      }
      
      const channel = document.getElementById("captureChannel").value;
      let nextAction = "Send DM";
      if (channel === "LinkedIn") nextAction = "Send connection request";
      else if (channel === "Email") nextAction = "Send pitch email";
      else if (channel === "WhatsApp") nextAction = "Send WhatsApp pitch";
      
      const keyword = finderKeyword ? finderKeyword.value.trim() : "";
      const notesPrefix = keyword ? `[Keyword: ${keyword}] ` : "";
      const notesVal = notesPrefix + document.getElementById("captureNotes").value.trim();
      
      const newLead = {
        dateAdded: getOffsetDateString(0),
        name: document.getElementById("captureName").value.trim(),
        market: document.getElementById("captureMarket").value,
        channel: channel,
        mainLink: normalizeUrl(document.getElementById("captureMainLink").value),
        niche: document.getElementById("captureNiche").value.trim(),
        source: "Lead Finder",
        priority: document.getElementById("capturePriority").value,
        stage: "Found (Lead collected only)",
        lastActionDate: "",
        nextAction: nextAction,
        nextActionDate: getOffsetDateString(0),
        replyStatus: "No reply",
        notes: notesVal,
        
        email: document.getElementById("captureEmail").value.trim(),
        whatsappNumber: document.getElementById("captureWhatsApp").value.trim(),
        extraLink: normalizeUrl(document.getElementById("captureWebsite").value || document.getElementById("captureIgLink").value || document.getElementById("captureLiLink").value),
        contactPerson: "",
        followUpCount: 0,
        messageSent: ""
      };
      
      leads.unshift(newLead);
      saveData();
      quickCaptureForm.reset();
      
      document.querySelectorAll(".keyword-chip").forEach(c => c.classList.remove("active"));
      if (finderKeyword) finderKeyword.value = "";
      if (document.getElementById("searchButtonsContainer")) {
        document.getElementById("searchButtonsContainer").innerHTML = `
          <div style="font-size: 12px; color: var(--color-priority-c); text-align: center; padding: 12px; background: var(--color-off-white); border-radius: var(--radius-sm);">
            Type a keyword or click a chip above to generate safe search buttons.
          </div>
        `;
      }
      
      updateDashboard();
      renderLeads();
      renderTodayActions();
      showToast(`Lead "${newLead.name}" captured successfully!`, "success");
    });
  }

  updateChipsVisibility();

  // --- Column Visibility Dropdown Event Listeners ---
  const colVisibilityBtn = document.getElementById("colVisibilityBtn");
  if (colVisibilityBtn) {
    colVisibilityBtn.addEventListener("click", toggleColVisibilityDropdown);
  }

  const colVisibilityMenu = document.getElementById("colVisibilityMenu");
  if (colVisibilityMenu) {
    colVisibilityMenu.addEventListener("click", (e) => {
      e.stopPropagation();
    });
  }

  const colCheckboxes = document.querySelectorAll(".col-toggle-checkbox");
  colCheckboxes.forEach(cb => {
    cb.addEventListener("change", (e) => {
      const colId = e.target.getAttribute("data-column");
      const isVisible = e.target.checked;
      toggleColumnVisibility(colId, isVisible);
    });
  });

  const thContactPerson = document.getElementById("thContactPerson");
  if (thContactPerson) {
    thContactPerson.addEventListener("click", () => {
      if (currentSortColumn === "contactPerson") {
        currentSortDirection = currentSortDirection === "asc" ? "desc" : "asc";
      } else {
        currentSortColumn = "contactPerson";
        currentSortDirection = "asc";
      }
      
      const indicator = document.getElementById("sortIndicatorContactPerson");
      if (indicator) {
        indicator.textContent = currentSortDirection === "asc" ? " ▲" : " ▼";
      }
      
      renderLeads();
    });
  }
}

// Escaping values for CSV
function escapeCsvValue(val) {
  if (val === undefined || val === null) return "";
  let str = String(val);
  // Replace double quotes with duplicate double quotes
  str = str.replace(/"/g, '""');
  // If the cell contains commas, new lines or double quotes, wrap it in double quotes
  if (str.includes(",") || str.includes("\n") || str.includes('"')) {
    str = `"${str}"`;
  }
  return str;
}

// CSV Export Logic
function exportToCSV() {
  if (leads.length === 0) {
    showToast("No leads to export.", "error");
    return;
  }

  // Header column titles
  const headers = [
    "Date Added", "Lead Name/Company", "Contact Person", "Market", 
    "Channel", "Main Link", "Extra Link", "Niche", "Source", "Priority", 
    "Stage", "Last Action Date", "Next Action", "Next Action Date", 
    "Reply Status", "Email", "WhatsApp Number", "Follow-up Count", 
    "Exact Message Sent", "Notes"
  ];

  // Convert objects to rows
  const csvRows = [headers.join(",")];
  
  leads.forEach(lead => {
    const row = [
      escapeCsvValue(lead.dateAdded),
      escapeCsvValue(lead.name),
      escapeCsvValue(lead.contactPerson),
      escapeCsvValue(lead.market),
      escapeCsvValue(lead.channel),
      escapeCsvValue(lead.mainLink),
      escapeCsvValue(lead.extraLink),
      escapeCsvValue(lead.niche),
      escapeCsvValue(lead.source || "Other"),
      escapeCsvValue(lead.priority),
      escapeCsvValue(lead.stage),
      escapeCsvValue(lead.lastActionDate),
      escapeCsvValue(lead.nextAction),
      escapeCsvValue(lead.nextActionDate),
      escapeCsvValue(lead.replyStatus),
      escapeCsvValue(lead.email),
      escapeCsvValue(lead.whatsappNumber),
      escapeCsvValue(lead.followUpCount),
      escapeCsvValue(lead.messageSent),
      escapeCsvValue(lead.notes)
    ];
    csvRows.push(row.join(","));
  });

  // Join and create Blob
  const csvContent = "\uFEFF" + csvRows.join("\n"); // add BOM for correct Excel encoding
  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  
  // Trigger download
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  const dateStr = getOffsetDateString(0);
  
  link.setAttribute("href", url);
  link.setAttribute("download", `Ali_Raza_Outreach_Leads_${dateStr}.csv`);
  link.style.visibility = "hidden";
  
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  
  showToast("CSV Exported successfully!", "success");
}

async function fetchLatestLeads() {
  try {
    const data = await apiFetch("/api/leads");
    if (data && data.leads) {
      leads = data.leads;
      localStorage.setItem("ali_raza_leads", JSON.stringify(leads));
      return true;
    }
  } catch (err) {
    console.error("Failed to fetch fresh leads from Neon database:", err);
  }
  return false;
}

function downloadCSVFile(leadsList, filename) {
  const headers = [
    "Date Added", "Lead Name/Company", "Contact Person", "Market", 
    "Channel", "Main Link", "Extra Link", "Niche", "Source", "Priority", 
    "Stage", "Last Action Date", "Next Action", "Next Action Date", 
    "Reply Status", "Email", "WhatsApp Number", "Follow-up Count", 
    "Exact Message Sent", "Notes"
  ];
  
  const csvRows = [headers.join(",")];
  
  leadsList.forEach(lead => {
    const row = [
      escapeCsvValue(lead.dateAdded || ""),
      escapeCsvValue(lead.name || "Unnamed Lead / Company"),
      escapeCsvValue(lead.contactPerson || ""),
      escapeCsvValue(lead.market || ""),
      escapeCsvValue(lead.channel || ""),
      escapeCsvValue(lead.mainLink || ""),
      escapeCsvValue(lead.extraLink || ""),
      escapeCsvValue(lead.niche || ""),
      escapeCsvValue(lead.source || "Other"),
      escapeCsvValue(lead.priority || "B"),
      escapeCsvValue(lead.stage || "Found (Lead collected only)"),
      escapeCsvValue(lead.lastActionDate || ""),
      escapeCsvValue(lead.nextAction || ""),
      escapeCsvValue(lead.nextActionDate || ""),
      escapeCsvValue(lead.replyStatus || "No reply"),
      escapeCsvValue(lead.email || ""),
      escapeCsvValue(lead.whatsappNumber || ""),
      escapeCsvValue(lead.followUpCount !== undefined && lead.followUpCount !== null ? lead.followUpCount : 0),
      escapeCsvValue(lead.messageSent || ""),
      escapeCsvValue(lead.notes || "")
    ];
    csvRows.push(row.join(","));
  });

  const csvContent = "\uFEFF" + csvRows.join("\n"); // add BOM for correct Excel encoding
  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  
  link.setAttribute("href", url);
  link.setAttribute("download", filename);
  link.style.visibility = "hidden";
  
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

async function exportFilteredLeads() {
  const exportBtn = document.getElementById("exportFilteredLeadsBtn");
  const originalBtnText = exportBtn ? exportBtn.innerHTML : "Export Filtered Leads";
  
  if (exportBtn) {
    exportBtn.disabled = true;
    exportBtn.textContent = "Syncing...";
  }

  // Use the online database as the source of truth
  const success = await fetchLatestLeads();
  if (!success) {
    showToast("Could not sync with online database before export. Using local data.", "warning");
  }

  const filtered = getFilteredLeads();
  if (filtered.length === 0) {
    showToast("No leads match the current filters. Adjust filters before exporting.", "error");
    if (exportBtn) {
      exportBtn.disabled = true;
      exportBtn.innerHTML = originalBtnText;
    }
    return;
  }

  // Before download, show a short confirmation message
  showToast(`Exporting ${filtered.length} filtered leads to CSV.`, "info");
  
  const dateStr = getOffsetDateString(0);
  const filename = `Ali_Raza_Filtered_Leads_${dateStr}.csv`;
  
  downloadCSVFile(filtered, filename);

  if (exportBtn) {
    exportBtn.disabled = (filtered.length === 0);
    exportBtn.innerHTML = originalBtnText;
  }
}

async function exportAllLeads() {
  const exportBtn = document.getElementById("exportAllLeadsBtn");
  const originalBtnText = exportBtn ? exportBtn.innerHTML : "Export All Leads";

  if (exportBtn) {
    exportBtn.disabled = true;
    exportBtn.textContent = "Syncing...";
  }

  // Use the online database as the source of truth
  const success = await fetchLatestLeads();
  if (!success) {
    showToast("Could not sync with online database before export. Using local data.", "warning");
  }

  if (leads.length === 0) {
    showToast("No leads to export.", "error");
    if (exportBtn) {
      exportBtn.disabled = false;
      exportBtn.innerHTML = originalBtnText;
    }
    return;
  }

  // Show confirmation message
  showToast(`Exporting ${leads.length} leads to CSV.`, "info");

  const dateStr = getOffsetDateString(0);
  const filename = `Ali_Raza_Filtered_Leads_${dateStr}.csv`;

  downloadCSVFile(leads, filename);

  if (exportBtn) {
    exportBtn.disabled = false;
    exportBtn.innerHTML = originalBtnText;
  }
}

// --- IMPORT LEADS IMPLEMENTATION ---

const SYNONYMS = {
  name: ["leadnamecompany", "leadcompany", "leadname", "company", "businessname", "name", "profilename", "account"],
  mainLink: ["website", "link", "url", "mainlink", "profile", "profilelink", "websiteprofile", "instagram", "linkedin", "instagramlink", "linkedinlink"],
  contactPerson: ["contactperson", "founder", "owner", "personname", "contact"],
  email: ["email", "emailaddress", "contactemail", "mail", "e-mail"],
  whatsappNumber: ["whatsapp", "whatsappnumber", "phone", "phonenumber", "mobile", "contactnumber", "numero", "telefono"],
  market: ["market"],
  niche: ["niche", "category", "industry", "leadtype", "type"],
  source: ["source", "leadsource"],
  priority: ["priority"],
  stage: ["stage", "status"],
  nextAction: ["nextaction"],
  nextActionDate: ["nextactiondate", "followupdate", "nextfollowupdate"],
  replyStatus: ["replystatus", "result"],
  notes: ["notes", "remarks", "comments", "context"],
  messageSent: ["messagesent", "personalizedfirstmessage", "firstmessage", "dm", "emailmessage", "whatsappmessage", "exactmessagesent"],
  dateAdded: ["dateadded", "date", "created", "createddate", "timecreated", "added"],
  lastActionDate: ["lastactiondate", "lastaction", "dateoflastaction", "lastcontact", "lastcontactdate"],
  channel: ["channel", "platform", "outreachchannel", "sourcechannel"]
};

let pendingValidLeads = [];
let pendingDuplicateLeads = [];
let pendingInvalidLeads = [];
let pendingSkippedCount = 0;
let importMarketStats = {};

function normalizeHeader(str) {
  if (str === undefined || str === null) return "";
  return str.toString().toLowerCase().replace(/[\s\-_/\\(),.?!;:'"&`+*@#^%[\]{}|]/g, "");
}

function normalizeChannel(str) {
  if (!str) return "";
  const s = String(str).trim().toLowerCase();
  if (s === "email") return "Email";
  if (s === "whatsapp" || s === "whats app") return "WhatsApp";
  if (s === "linkedin" || s === "linked in") return "LinkedIn";
  if (s === "instagram" || s === "ig") return "Instagram";
  return "";
}

const VALID_MARKETS = [
  "US", "UK", "Canada", "Australia", "Ireland", "New Zealand", "Singapore",
  "Netherlands", "Germany", "Austria", "Switzerland", "Belgium", "Denmark",
  "Sweden", "Norway", "Finland", "Italy", "Portugal", "UAE", "South Africa",
  "Malaysia", "Philippines", "Hong Kong", "Saudi Arabia", "Other"
];

const lowerToExactMarket = {};
VALID_MARKETS.forEach(m => {
  lowerToExactMarket[m.toLowerCase()] = m;
});

function normalizeMarket(val) {
  if (val === undefined || val === null) return "";
  const clean = val.toString().trim();
  if (clean === "") return "";
  const lower = clean.toLowerCase();
  
  if (lower === "usa" || lower === "united states" || lower === "u.s.a.") {
    return "US";
  }
  if (lower === "united kingdom" || lower === "u.k.") {
    return "UK";
  }
  if (lower === "united arab emirates") {
    return "UAE";
  }
  if (lower === "ksa") {
    return "Saudi Arabia";
  }
  
  if (lowerToExactMarket[lower]) {
    return lowerToExactMarket[lower];
  }
  
  return "Other";
}

function normalizeUrlForComparison(url) {
  if (!url) return "";
  let clean = url.toString().trim().toLowerCase();
  clean = clean.replace(/^(https?:\/\/)?(www\.)?/, "");
  clean = clean.replace(/\/$/, "");
  return clean;
}

function normalizeNameForComparison(name) {
  if (!name) return "";
  let clean = name.toString().toLowerCase().trim();
  clean = clean.replace(/[\.,\/#!$%\^&\*;:{}=\-_`~()]/g, "");
  clean = clean.replace(/\s+/g, " ");
  return clean.trim();
}

function normalizeEmailForComparison(email) {
  if (!email) return "";
  return email.toString().trim().toLowerCase();
}

function checkDuplicateForImport(lead) {
  const emailNorm = normalizeEmailForComparison(lead.email);
  let matchedLead = null;
  let reason = "";

  // Priority 1: Email Address matches existing lead email
  if (emailNorm) {
    matchedLead = leads.find(l => normalizeEmailForComparison(l.email) === emailNorm);
    if (matchedLead) {
      reason = `Email Address matches existing lead email address: "${lead.email}"`;
      return { isDup: true, matchedLead, reason };
    }
  }

  // Priority 2: Main Profile Link matches existing lead main link
  const linkNorm = normalizeUrlForComparison(lead.mainLink);
  if (linkNorm) {
    matchedLead = leads.find(l => normalizeUrlForComparison(l.mainLink) === linkNorm);
    if (matchedLead) {
      reason = `Main Profile Link matches existing lead profile link: "${lead.mainLink}"`;
      return { isDup: true, matchedLead, reason };
    }
  }

  // Priority 3: When email and main profile link are both empty, name + market match
  if (!emailNorm && !linkNorm) {
    const nameNorm = normalizeNameForComparison(lead.name);
    const marketNorm = lead.market ? lead.market.trim().toLowerCase() : "";
    if (nameNorm) {
      matchedLead = leads.find(l => {
        return normalizeNameForComparison(l.name) === nameNorm &&
               (l.market ? l.market.trim().toLowerCase() : "") === marketNorm;
      });
      if (matchedLead) {
        reason = `Lead Name/Company and Market match existing lead: "${lead.name}" in Market "${lead.market}"`;
        return { isDup: true, matchedLead, reason };
      }
    }
  }

  return { isDup: false };
}

function mapHeaders(headers) {
  const mapping = {};
  const normalizedHeaders = headers.map(h => normalizeHeader(h));

  // Priority mapping for 'name'
  const namePriority = ["leadnamecompany", "leadcompany", "leadname", "company", "businessname", "name", "profilename", "account"];
  let nameIdx = -1;
  for (const p of namePriority) {
    const idx = normalizedHeaders.indexOf(p);
    if (idx !== -1) {
      nameIdx = idx;
      break;
    }
  }
  if (nameIdx !== -1) {
    mapping["name"] = nameIdx;
  }

  // Map other keys using synonyms
  for (const [key, synonyms] of Object.entries(SYNONYMS)) {
    if (key === "name") continue; // handled above
    for (let idx = 0; idx < normalizedHeaders.length; idx++) {
      const normH = normalizedHeaders[idx];
      if (normH && synonyms.includes(normH)) {
        mapping[key] = idx;
        break; // map to the first match
      }
    }
  }
  return mapping;
}

function processImportData(arrayBuffer, filename) {
  try {
    const data = new Uint8Array(arrayBuffer);
    const workbook = XLSX.read(data, { type: 'array', cellStyles: true, cellDates: true });
    
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    
    // header: 1 formats it as an array of arrays
    const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
    
    if (rawRows.length === 0) {
      showToast("The imported file is empty.", "error");
      return;
    }
    
    const headers = rawRows[0];
    const rows = rawRows.slice(1);
    
    // Map headers to column indices
    const colMapping = mapHeaders(headers);
    
    pendingValidLeads = [];
    pendingDuplicateLeads = [];
    pendingInvalidLeads = [];
    pendingSkippedCount = 0;
    
    importMarketStats = {
      withVal: 0,
      blank: 0,
      first5: []
    };
    VALID_MARKETS.forEach(m => {
      importMarketStats[m] = 0;
    });
    
    let totalRows = 0;
    
    rows.forEach((row, rowIdx) => {
      // Skip completely empty rows
      if (row.length === 0 || row.every(val => val === undefined || val === null || String(val).trim() === "")) {
        return;
      }
      
      totalRows++;
      
      const getVal = (field) => {
        const colIdx = colMapping[field];
        if (colIdx === undefined) return "";
        const val = row[colIdx];
        return val !== undefined && val !== null ? val : "";
      };
      
      const rowName = getVal("name") ? String(getVal("name")).trim() : "";
      const rowEmail = getVal("email") ? String(getVal("email")).trim() : "";
      const rowMainLink = getVal("mainLink") ? String(getVal("mainLink")).trim() : "";
      
      // Validation: at least one of Name, Email, or Profile Link. Skip only if all 3 are empty.
      if (!rowName && !rowEmail && !rowMainLink) {
        pendingInvalidLeads.push({
          rowNum: rowIdx + 2,
          reason: "Missing Name, Email, and Link"
        });
        return;
      }
      
      // Map basic values
      const rowContactPerson = getVal("contactPerson") ? String(getVal("contactPerson")).trim() : "";
      const rowWhatsapp = getVal("whatsappNumber") ? String(getVal("whatsappNumber")).trim() : "";
      const rawMarket = getVal("market") ? String(getVal("market")).trim() : "";
      const rowMarket = normalizeMarket(rawMarket);
      
      // Calculate market preview statistics
      if (rowMarket) {
        importMarketStats.withVal++;
        if (importMarketStats[rowMarket] !== undefined) {
          importMarketStats[rowMarket]++;
        } else {
          importMarketStats.Other++;
        }
      } else {
        importMarketStats.blank++;
      }
      if (importMarketStats.first5.length < 5) {
        importMarketStats.first5.push(rowMarket || "(blank)");
      }
      
      const rowNiche = getVal("niche") ? String(getVal("niche")).trim() : "";
      const rowSource = getVal("source") ? String(getVal("source")).trim() : "";
      const rowPriority = getVal("priority") ? String(getVal("priority")).trim() : "";
      const rowStageRaw = getVal("stage") ? String(getVal("stage")).trim() : "";
      const rowNextAction = getVal("nextAction") ? String(getVal("nextAction")).trim() : "";
      const rowNextActionDate = getVal("nextActionDate") ? String(getVal("nextActionDate")).trim() : "";
      const rowLastActionDate = getVal("lastActionDate") ? String(getVal("lastActionDate")).trim() : "";
      const rowReplyStatus = getVal("replyStatus") ? String(getVal("replyStatus")).trim() : "";
      const rowNotes = getVal("notes") ? String(getVal("notes")).trim() : "";
      const rowMessageSent = getVal("messageSent") ? String(getVal("messageSent")).trim() : "";
      const rowDateAdded = getVal("dateAdded") ? String(getVal("dateAdded")).trim() : "";
      const rowChannelRaw = getVal("channel") ? String(getVal("channel")).trim() : "";
      
      // Determine channel
      let channel = normalizeChannel(rowChannelRaw);
      if (!channel) {
        channel = (activeTab && activeTab !== "All" && activeTab !== "LeadFinder" && activeTab !== "TodayMode") ? activeTab : "Email";
      }
 
      // Determine stage
      let stage = "Found (Lead collected only)";
      if (rowStageRaw) {
        const matchedStage = VALID_STAGES.find(s => s.toLowerCase() === rowStageRaw.toLowerCase());
        if (matchedStage) {
          stage = matchedStage;
        }
      }
 
      // Base lead mapping
      const lead = {
        dateAdded: parseExcelDate(rowDateAdded) || getOffsetDateString(0),
        name: rowName || "Unnamed Lead / Company",
        market: rowMarket,
        channel: channel,
        mainLink: rowMainLink ? normalizeUrl(rowMainLink) : "",
        niche: rowNiche,
        source: rowSource || "Import",
        priority: rowPriority || "B",
        stage: stage,
        lastActionDate: parseExcelDate(rowLastActionDate) || getOffsetDateString(0),
        nextAction: rowNextAction,
        nextActionDate: parseExcelDate(rowNextActionDate),
        replyStatus: rowReplyStatus || "No reply",
        notes: rowNotes,
        contactPerson: rowContactPerson,
        email: rowEmail,
        whatsappNumber: rowWhatsapp,
        extraLink: "",
        messageSent: rowMessageSent,
        followUpCount: 0
      };
 
      // Set default nextAction if empty based on channel
      if (!lead.nextAction) {
        if (lead.channel === "LinkedIn") lead.nextAction = "Send connection request";
        else if (lead.channel === "Email") lead.nextAction = "Send pitch email";
        else if (lead.channel === "WhatsApp") lead.nextAction = "Send WhatsApp pitch";
        else if (lead.channel === "Instagram") lead.nextAction = "Send DM";
        else lead.nextAction = "Send first message";
      }
 
      // Set default nextActionDate if empty
      if (!lead.nextActionDate) {
        lead.nextActionDate = getOffsetDateString(0);
      }
 
      // Process unmapped columns into notes
      const extraNotesList = [];
      headers.forEach((header, idx) => {
        const isMapped = Object.values(colMapping).includes(idx);
        if (!isMapped) {
          const val = row[idx];
          if (val !== undefined && val !== null && String(val).trim() !== "") {
            const colLetter = String.fromCharCode(65 + idx);
            const headerName = header && String(header).trim() ? String(header).trim() : `Col ${colLetter}`;
            extraNotesList.push(`${headerName}: ${val}`);
          }
        }
      });
      if (extraNotesList.length > 0) {
        lead.notes = (lead.notes ? lead.notes + "\n" : "") + extraNotesList.join(" | ");
      }
 
      // Duplicate checks
      const dupCheck = checkDuplicateForImport(lead);
      if (dupCheck.isDup) {
        pendingDuplicateLeads.push({
          lead: lead,
          matchedLead: dupCheck.matchedLead,
          reason: dupCheck.reason,
          resolution: "skip", // Default choice must be 'skip'
          rowIdx: rowIdx + 2
        });
      } else {
        pendingValidLeads.push(lead);
      }
    });
 
    pendingSkippedCount = pendingInvalidLeads.length;
    showImportPreview(totalRows);
 
  } catch (err) {
    console.error("Error parsing spreadsheet file:", err);
    showToast(`Error reading spreadsheet: ${err.message}`, "error");
  }
}

function escapeHtml(str) {
  if (!str) return "";
  return str.toString()
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function showImportPreview(totalRows) {
  // Update Total detected count
  document.getElementById("prevTotalRows").textContent = totalRows;

  // Update Market Details
  const prevMarketWithVal = document.getElementById("prevMarketWithVal");
  if (prevMarketWithVal) prevMarketWithVal.textContent = importMarketStats.withVal;
  
  const prevMarketBlank = document.getElementById("prevMarketBlank");
  if (prevMarketBlank) prevMarketBlank.textContent = importMarketStats.blank;
  
  const prevFirst5Markets = document.getElementById("prevFirst5Markets");
  if (prevFirst5Markets) {
    prevFirst5Markets.textContent = importMarketStats.first5.length > 0 ? importMarketStats.first5.join(", ") : "None";
  }
  
  VALID_MARKETS.forEach(m => {
    const id = "dist" + m.replace(/\s+/g, "");
    const el = document.getElementById(id);
    if (el) el.textContent = importMarketStats[m] || 0;
  });
  
  // Render Section C (invalid rows)
  const prevSkippedDiv = document.getElementById("prevSkippedRowsList");
  if (prevSkippedDiv) {
    if (pendingInvalidLeads.length === 0) {
      prevSkippedDiv.innerHTML = "No invalid rows found.";
    } else {
      prevSkippedDiv.innerHTML = pendingInvalidLeads.map(item => {
        return `Row ${item.rowNum}: Skipped — ${escapeHtml(item.reason)}`;
      }).join("<br>");
    }
  }

  // Render duplicates count badge
  const countBadge = document.getElementById("duplicatesCountBadge");
  if (countBadge) {
    countBadge.textContent = `(${pendingDuplicateLeads.length} found)`;
  }

  // Render Duplicates Resolution table in Section B
  const dupContainer = document.getElementById("duplicatesListContainer");
  if (dupContainer) {
    if (pendingDuplicateLeads.length === 0) {
      dupContainer.innerHTML = `<div style="padding: 20px; text-align: center; color: var(--color-priority-c); font-size: 13px;">No duplicates detected in this file.</div>`;
    } else {
      let tableHtml = `
        <table style="width: 100%; border-collapse: collapse; font-size: 13px; text-align: left;">
          <thead>
            <tr style="border-bottom: 2px solid rgba(11, 31, 58, 0.1); background-color: rgba(11, 31, 58, 0.02);">
              <th style="padding: 8px; width: 50px;">Row</th>
              <th style="padding: 8px;">Imported Lead Info</th>
              <th style="padding: 8px;">Matched Existing Lead</th>
              <th style="padding: 8px; width: 200px;">Conflict Action</th>
            </tr>
          </thead>
          <tbody>
      `;
      pendingDuplicateLeads.forEach((dup, idx) => {
        tableHtml += `
          <tr style="border-bottom: 1px solid rgba(11, 31, 58, 0.05); vertical-align: top;">
            <td style="padding: 8px; font-weight: bold;">${dup.rowIdx}</td>
            <td style="padding: 8px;">
              <div style="font-weight: 600; color: var(--color-deep-navy);">${escapeHtml(dup.lead.name)}</div>
              <div style="font-size: 11px; color: var(--color-priority-c);">${escapeHtml(dup.lead.email || dup.lead.mainLink || 'No email/link')}</div>
            </td>
            <td style="padding: 8px;">
              <div style="font-weight: 600; color: var(--color-deep-navy);">${escapeHtml(dup.matchedLead.name)}</div>
              <div style="font-size: 11px; color: var(--color-priority-c);">${escapeHtml(dup.matchedLead.email || 'No email')}</div>
              <div style="font-size: 11px; color: #b45309; font-style: italic; margin-top: 4px;">Reason: ${escapeHtml(dup.reason)}</div>
            </td>
            <td style="padding: 8px;">
              <select class="dup-resolution-select" data-index="${idx}" style="padding: 4px 8px; border-radius: var(--radius-sm); border: var(--border-light); font-size: 12px; background: #fff; width: 100%;">
                <option value="skip" ${dup.resolution === 'skip' ? 'selected' : ''}>Skip Duplicate</option>
                <option value="update" ${dup.resolution === 'update' ? 'selected' : ''}>Update Existing Lead</option>
                <option value="create" ${dup.resolution === 'create' ? 'selected' : ''}>Create New Lead Anyway</option>
              </select>
            </td>
          </tr>
        `;
      });
      tableHtml += `</tbody></table>`;
      dupContainer.innerHTML = tableHtml;

      // Add change event listeners to individual selects
      const selects = dupContainer.querySelectorAll(".dup-resolution-select");
      selects.forEach(select => {
        select.addEventListener("change", (e) => {
          const idx = parseInt(e.target.getAttribute("data-index"));
          const val = e.target.value;
          pendingDuplicateLeads[idx].resolution = val;
          updateImportSummary(totalRows);
        });
      });
    }
  }

  // Clone and bind bulk resolution buttons
  const oldSkip = document.getElementById("bulkSkipBtn");
  if (oldSkip) {
    const newSkip = oldSkip.cloneNode(true);
    oldSkip.parentNode.replaceChild(newSkip, oldSkip);
    newSkip.addEventListener("click", () => {
      pendingDuplicateLeads.forEach(dup => dup.resolution = "skip");
      document.querySelectorAll(".dup-resolution-select").forEach(sel => sel.value = "skip");
      updateImportSummary(totalRows);
    });
  }
  
  const oldUpdate = document.getElementById("bulkUpdateBtn");
  if (oldUpdate) {
    const newUpdate = oldUpdate.cloneNode(true);
    oldUpdate.parentNode.replaceChild(newUpdate, oldUpdate);
    newUpdate.addEventListener("click", () => {
      pendingDuplicateLeads.forEach(dup => dup.resolution = "update");
      document.querySelectorAll(".dup-resolution-select").forEach(sel => sel.value = "update");
      updateImportSummary(totalRows);
    });
  }
  
  const oldCreate = document.getElementById("bulkCreateBtn");
  if (oldCreate) {
    const newCreate = oldCreate.cloneNode(true);
    oldCreate.parentNode.replaceChild(newCreate, oldCreate);
    newCreate.addEventListener("click", () => {
      pendingDuplicateLeads.forEach(dup => dup.resolution = "create");
      document.querySelectorAll(".dup-resolution-select").forEach(sel => sel.value = "create");
      updateImportSummary(totalRows);
    });
  }

  // Initialize summary counts and dynamic fields
  updateImportSummary(totalRows);
  
  const modal = document.getElementById("importPreviewModal");
  if (modal) modal.classList.add("active");
}

function updateImportSummary(totalRows) {
  let newLeadsCount = pendingValidLeads.length;
  let updateLeadsCount = 0;
  let skippedLeadsCount = 0;

  pendingDuplicateLeads.forEach(dup => {
    if (dup.resolution === "create") {
      newLeadsCount++;
    } else if (dup.resolution === "update") {
      updateLeadsCount++;
    } else {
      skippedLeadsCount++;
    }
  });

  const invalidRowsCount = pendingSkippedCount;

  // Update Summary Stats Table
  document.getElementById("summaryNewCount").textContent = newLeadsCount;
  document.getElementById("summaryUpdateCount").textContent = updateLeadsCount;
  document.getElementById("summarySkipCount").textContent = skippedLeadsCount;
  document.getElementById("summaryInvalidCount").textContent = invalidRowsCount;

  // Render First 5 Mapped New Names
  const newNames = [];
  pendingValidLeads.forEach(l => newNames.push(l.name));
  pendingDuplicateLeads.forEach(dup => {
    if (dup.resolution === "create") {
      newNames.push(dup.lead.name);
    }
  });
  const first5New = newNames.slice(0, 5);
  const prevNamesUl = document.getElementById("prevFirst5Names");
  if (prevNamesUl) {
    prevNamesUl.innerHTML = first5New.map(name => `<li>${escapeHtml(name)}</li>`).join("");
    if (first5New.length === 0) {
      prevNamesUl.innerHTML = "<li>No new leads will be created.</li>";
    }
  }

  // Render First 5 Leads to Update
  const updateNames = [];
  pendingDuplicateLeads.forEach(dup => {
    if (dup.resolution === "update") {
      updateNames.push(`${dup.matchedLead.name} (Matched by Row ${dup.rowIdx})`);
    }
  });
  const first5Updates = updateNames.slice(0, 5);
  const prevUpdatesUl = document.getElementById("prevFirst5Updates");
  const updatesSection = document.getElementById("leadsToUpdateSection");
  if (prevUpdatesUl && updatesSection) {
    if (first5Updates.length > 0) {
      updatesSection.style.display = "block";
      prevUpdatesUl.innerHTML = first5Updates.map(name => `<li>${escapeHtml(name)}</li>`).join("");
    } else {
      updatesSection.style.display = "none";
      prevUpdatesUl.innerHTML = "";
    }
  }

  // Dynamic Confirm Button text and state
  const confirmBtn = document.getElementById("confirmImportBtn");
  if (confirmBtn) {
    const totalActions = newLeadsCount + updateLeadsCount;
    if (totalActions === 0) {
      confirmBtn.disabled = true;
      confirmBtn.textContent = "No Valid Leads to Import";
    } else {
      confirmBtn.disabled = false;
      if (updateLeadsCount > 0) {
        confirmBtn.textContent = `Import ${newLeadsCount} New Leads and Update ${updateLeadsCount} Leads`;
      } else {
        confirmBtn.textContent = `Import ${newLeadsCount} New Leads`;
      }
    }
  }
}

function prepareDuplicateAsNew(imported) {
  const newLead = { ...imported };
  delete newLead.id; // clear any existing ID reference just in case
  
  // Add note automatically
  const todayStr = getOffsetDateString(0);
  const dupNote = `Possible duplicate imported manually on ${todayStr}.`;
  newLead.notes = (newLead.notes ? newLead.notes + "\n\n" : "") + dupNote;
  
  return newLead;
}

function updateExistingLeadObj(existing, imported) {
  // Update fields if present in imported cell
  if (imported.name) existing.name = imported.name;
  if (imported.contactPerson) existing.contactPerson = imported.contactPerson;
  if (imported.market) existing.market = imported.market;
  if (imported.channel) existing.channel = imported.channel;
  if (imported.mainLink) existing.mainLink = imported.mainLink;
  if (imported.extraLink) existing.extraLink = imported.extraLink;
  if (imported.niche) existing.niche = imported.niche;
  if (imported.source) existing.source = imported.source;
  if (imported.priority) existing.priority = imported.priority;
  
  // For Outreach Stage: Keep the existing stage unless the CSV contains a valid stage value
  if (imported.stage) {
    const isValid = VALID_STAGES.includes(imported.stage);
    if (isValid) {
      existing.stage = imported.stage;
    }
  }

  if (imported.lastActionDate) existing.lastActionDate = imported.lastActionDate;
  if (imported.nextAction) existing.nextAction = imported.nextAction;
  if (imported.nextActionDate) existing.nextActionDate = imported.nextActionDate;
  if (imported.replyStatus) existing.replyStatus = imported.replyStatus;
  if (imported.email) existing.email = imported.email;
  if (imported.whatsappNumber) existing.whatsappNumber = imported.whatsappNumber;
  
  // Follow-up count
  if (imported.followUpCount !== undefined && imported.followUpCount !== null && imported.followUpCount !== "") {
    const countVal = parseInt(imported.followUpCount);
    if (!isNaN(countVal)) {
      existing.followUpCount = countVal;
    }
  }

  if (imported.messageSent) existing.messageSent = imported.messageSent;

  // Notes handling: append under Imported update header
  const importedNote = imported.notes ? imported.notes.trim() : "";
  if (importedNote) {
    const existingNotesStr = existing.notes ? existing.notes.toString() : "";
    if (!existingNotesStr.includes(importedNote)) {
      const todayStr = getOffsetDateString(0);
      const appendStr = `Imported update — ${todayStr}:\n${importedNote}`;
      existing.notes = (existingNotesStr ? existingNotesStr + "\n\n" : "") + appendStr;
    }
  }
}

async function confirmImport() {
  const confirmBtn = document.getElementById("confirmImportBtn");
  const originalBtnText = confirmBtn ? confirmBtn.textContent : "Confirm Import";
  if (confirmBtn) {
    confirmBtn.disabled = true;
    confirmBtn.textContent = "Saving...";
  }

  // Transaction: save a deep copy of original leads list
  const originalLeads = JSON.parse(JSON.stringify(leads));

  try {
    let newLeadsCount = 0;
    let updatedLeadsCount = 0;
    let skippedLeadsCount = 0;

    // 1. Process Section A (new leads)
    const newLeadsToCreate = [];
    pendingValidLeads.forEach(l => {
      const copy = { ...l };
      ensureLeadId(copy);
      newLeadsToCreate.push(copy);
      newLeadsCount++;
    });

    // 2. Process Section B (duplicates based on conflict action)
    pendingDuplicateLeads.forEach(dup => {
      if (dup.resolution === "create") {
        const copy = prepareDuplicateAsNew(dup.lead);
        ensureLeadId(copy);
        newLeadsToCreate.push(copy);
        newLeadsCount++;
      } else if (dup.resolution === "update") {
        // Find existing lead by ID
        const existingId = dup.matchedLead.id;
        const existing = leads.find(l => l.id === existingId);
        if (existing) {
          updateExistingLeadObj(existing, dup.lead);
          updatedLeadsCount++;
        }
      } else {
        // resolution is 'skip'
        skippedLeadsCount++;
      }
    });

    // Add new leads to local array
    leads.push(...newLeadsToCreate);

    // 3. Attempt direct sync with online database
    setSyncStatus("syncing");
    await apiFetch("/api/leads", { method: "POST", body: { leads } });
    setSyncStatus("connected");

    // Sync succeeded: save to LocalStorage
    localStorage.setItem("ali_raza_leads", JSON.stringify(leads));

    // Success behavior:
    // Reset all advanced filters
    document.getElementById("searchInput").value = "";
    document.getElementById("filterChannel").value = "All";
    document.getElementById("filterMarket").value = "All";
    document.getElementById("filterSource").value = "All";
    document.getElementById("filterContactPerson").value = "";
    if (document.getElementById("clearFilterContactPerson")) {
      document.getElementById("clearFilterContactPerson").style.display = "none";
    }
    document.getElementById("filterPriority").value = "All";
    document.getElementById("filterStage").value = "All";
    document.getElementById("filterReply").value = "All";
    document.getElementById("filterActionDate").value = "All";
    document.getElementById("filterNextAction").value = "All";

    // Reset quick filters active class to All
    const quickFilters = [
      { id: "qFilterAll", value: "All" },
      { id: "qFilterToday", value: "Today" },
      { id: "qFilterFollowUp", value: "FollowUp" },
      { id: "qFilterA", value: "A" },
      { id: "qFilterWarm", value: "Warm" },
      { id: "qFilterArchived", value: "Archived" }
    ];
    quickFilters.forEach(x => {
      const btn = document.getElementById(x.id);
      if (btn) {
        if (x.value === "All") btn.classList.add("active");
        else btn.classList.remove("active");
      }
    });
    activeQuickFilter = "All";

    // Switch to the Email Leads tab
    activeTab = "Email";
    const tabBtns = document.querySelectorAll(".tab-btn");
    tabBtns.forEach(btn => {
      btn.classList.remove("active");
      if (btn.getAttribute("data-channel") === "Email") {
        btn.classList.add("active");
      }
    });
    const filterChanSelect = document.getElementById("filterChannel");
    if (filterChanSelect) filterChanSelect.value = "Email";

    // Refresh views
    updateDashboard();
    renderLeads();
    renderTodayActions();

    // Success toast alert:
    showToast(`${newLeadsCount} new leads imported, ${updatedLeadsCount} existing leads updated, and ${skippedLeadsCount + pendingSkippedCount} duplicate rows skipped successfully.`, "success");

  } catch (err) {
    console.error("Error saving imported leads to cloud:", err);
    // Rollback local memory leads list on error
    leads.length = 0;
    leads.push(...originalLeads);
    setSyncStatus("offline");
    showToast(`Failed to sync imported leads to database: ${err.message}`, "error");
  } finally {
    if (confirmBtn) {
      confirmBtn.disabled = false;
      confirmBtn.textContent = originalBtnText;
    }
    closeImportPreview();
  }
}

function closeImportPreview() {
  const modal = document.getElementById("importPreviewModal");
  if (modal) modal.classList.remove("active");
  pendingValidLeads = [];
  pendingSkippedCount = 0;
  const fileInput = document.getElementById("importFileInput");
  if (fileInput) fileInput.value = "";
}

// --- BACKUP JSON IMPLEMENTATION ---
let pendingBackupData = null;

// Helper to merge lead objects without overwriting populated fields with empty values
function mergeLeadObjects(existingLead, newLead) {
  const merged = { ...existingLead };
  Object.keys(newLead).forEach(key => {
    const newVal = newLead[key];
    if (newVal !== undefined && newVal !== null && String(newVal).trim() !== "") {
      merged[key] = newVal;
    }
  });
  return merged;
}

// Check if a backup lead matches an existing lead by email, phone, main link, or name
function checkBackupLeadDuplicate(newLead) {
  const emailVal = newLead.email ? newLead.email.trim().toLowerCase() : "";
  const phoneVal = newLead.whatsappNumber ? newLead.whatsappNumber.trim().replace(/[\s\+\(\)\-\[\]]/g, "") : "";
  const mainLinkVal = newLead.mainLink ? newLead.mainLink.trim().toLowerCase() : "";
  const nameVal = newLead.name ? newLead.name.trim().toLowerCase() : "";
  
  for (let i = 0; i < leads.length; i++) {
    const l = leads[i];
    
    // Check Email
    if (emailVal && l.email && l.email.trim().toLowerCase() === emailVal) {
      return { type: "duplicate", index: i, reason: "Email match" };
    }
    
    // Check WhatsApp
    if (phoneVal) {
      const existingPhone = (l.whatsappNumber || "").trim().replace(/[\s\+\(\)\-\[\]]/g, "");
      if (existingPhone && existingPhone === phoneVal) {
        return { type: "duplicate", index: i, reason: "WhatsApp match" };
      }
    }
    
    // Check Main Link
    if (mainLinkVal && l.mainLink && l.mainLink.trim().toLowerCase() === mainLinkVal) {
      return { type: "duplicate", index: i, reason: "Main link match" };
    }
    
    // Check Name
    if (nameVal && l.name && l.name.trim().toLowerCase() === nameVal) {
      return { type: "duplicate", index: i, reason: "Name match" };
    }
  }
  return null;
}

// Export Backup JSON
function exportBackupJSON() {
  const data = {
    appName: "Ali Raza Outreach Tracker",
    appVersion: "1.0.0",
    backupCreatedDate: new Date().toISOString(),
    leads: leads,
    settings: {
      theme: localStorage.getItem("theme") || "dark",
      ali_raza_logged_in: localStorage.getItem("ali_raza_logged_in")
    }
  };
  const jsonStr = JSON.stringify(data, null, 2);
  const blob = new Blob([jsonStr], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  
  const dateStr = getOffsetDateString(0);
  const filename = `ali-raza-outreach-backup-${dateStr}.json`;
  
  const link = document.createElement("a");
  link.setAttribute("href", url);
  link.setAttribute("download", filename);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
  
  showToast("Backup JSON exported successfully!", "success");
}

// Handle backup file select
function handleBackupFileSelect(event) {
  const file = event.target.files[0];
  if (!file) return;
  
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = JSON.parse(e.target.result);
      if (!data || !Array.isArray(data.leads)) {
        showToast("Invalid backup file format. Leads array not found.", "error");
        return;
      }
      
      pendingBackupData = {
        leads: Array.isArray(data.leads) ? data.leads : []
      };
      
      showBackupImportPreview();
    } catch (err) {
      showToast("Error parsing backup file: " + err.message, "error");
    }
  };
  reader.readAsText(file);
}

// Show backup import preview summary in the modal
function showBackupImportPreview() {
  if (!pendingBackupData) return;
  
  const leadsInBackup = pendingBackupData.leads;
  
  let duplicateCount = 0;
  
  leadsInBackup.forEach(lead => {
    const dupCheck = checkBackupLeadDuplicate(lead);
    if (dupCheck) {
      duplicateCount++;
    }
  });
  
  document.getElementById("backLeadsFound").textContent = leadsInBackup.length;
  document.getElementById("backExistingLeads").textContent = leads.length;
  document.getElementById("backDuplicatesFound").textContent = duplicateCount;
  
  // Reset select radio to default 'merge' and duplicate checkbox to unchecked
  const radioMerge = document.getElementById("backModeMerge");
  if (radioMerge) radioMerge.checked = true;
  
  const chkDupAnyway = document.getElementById("backImportDuplicatesAnyway");
  if (chkDupAnyway) chkDupAnyway.checked = false;
  
  updateBackupPreviewCounts();
  
  const modal = document.getElementById("backupImportPreviewModal");
  if (modal) modal.classList.add("active");
}

// Update counts dynamically when toggle mode (Merge vs Replace) or toggle import duplicates checkbox
function updateBackupPreviewCounts() {
  if (!pendingBackupData) return;
  
  const radioMerge = document.getElementById("backModeMerge");
  const mode = (radioMerge && radioMerge.checked) ? "merge" : "replace";
  const importDuplicates = document.getElementById("backImportDuplicatesAnyway")?.checked || false;
  const leadsInBackup = pendingBackupData.leads;
  
  let finalCount = 0;
  if (mode === "replace") {
    finalCount = leadsInBackup.length;
  } else {
    // Merge mode
    if (importDuplicates) {
      finalCount = leads.length + leadsInBackup.length;
    } else {
      let newCount = 0;
      leadsInBackup.forEach(lead => {
        const dupCheck = checkBackupLeadDuplicate(lead);
        if (!dupCheck) {
          newCount++;
        }
      });
      finalCount = leads.length + newCount;
    }
  }
  
  document.getElementById("backFinalCount").textContent = finalCount;
}

// Confirm and execute the JSON import
function confirmBackupImport() {
  if (!pendingBackupData) return;
  
  const radioMerge = document.getElementById("backModeMerge");
  const mode = (radioMerge && radioMerge.checked) ? "merge" : "replace";
  
  if (mode === "replace") {
    const warningText = "This will replace all current leads in this browser. Export a backup first. Continue?";
    if (!confirm(warningText)) {
      return;
    }
    
    // Replace mode
    leads = pendingBackupData.leads;
    
    showToast("Application data replaced successfully!", "success");
  } else {
    // Merge mode
    let addedCount = 0;
    let skippedCount = 0;
    
    const importDuplicates = document.getElementById("backImportDuplicatesAnyway")?.checked || false;
    
    pendingBackupData.leads.forEach(newLead => {
      const dupCheck = checkBackupLeadDuplicate(newLead);
      if (dupCheck) {
        if (importDuplicates) {
          leads.push(newLead);
          addedCount++;
        } else {
          skippedCount++;
        }
      } else {
        leads.push(newLead);
        addedCount++;
      }
    });
    
    let toastMsg = `Merged backup successfully: added ${addedCount} leads`;
    if (skippedCount > 0) {
      toastMsg += `, skipped ${skippedCount} duplicates`;
    }
    toastMsg += `.`;
    showToast(toastMsg, "success");
  }
  
  saveData();
  closeBackupImportModal();
  
  // Refresh views
  updateDashboard();
  renderLeads();
  renderTodayActions();
}

function closeBackupImportModal() {
  const modal = document.getElementById("backupImportPreviewModal");
  if (modal) modal.classList.remove("active");
  document.getElementById("backupFileInput").value = "";
  pendingBackupData = null;
}

window.toggleBackupDropdown = function(event) {
  event.stopPropagation();
  const dropdown = document.getElementById("backupDropdown");
  if (!dropdown) return;
  const wasOpen = dropdown.classList.contains("show");
  
  document.querySelectorAll(".dropdown-content.show").forEach(d => {
    if (d.id !== "backupDropdown") d.classList.remove("show");
  });
  
  if (wasOpen) {
    dropdown.classList.remove("show");
  } else {
    dropdown.classList.add("show");
  }
};

// Global functions accessible from HTML (inline onclick)
window.openEditModal = openEditModal;
window.archiveLead = archiveLead;
window.deleteLead = deleteLead;
// Simple HTML escaper helper to prevent XSS
function escapeHtml(text) {
  if (typeof text !== 'string') return text;
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

// Helpers to identify action categories
function isFirstAction(action) {
  if (!action) return false;
  const act = action.toLowerCase();
  return act.includes("email") || act.includes("whatsapp") || act.includes("dm") || act.includes("connection");
}

function isFollowupAction(action) {
  if (!action) return false;
  const act = action.toLowerCase();
  return act.includes("follow-up") || act.includes("follow up");
}

function isWarmStatus(status) {
  if (!status) return false;
  const s = status.toLowerCase();
  const warmStatuses = [
    "replied", "interested", "samples requested", "cv requested", 
    "price asked", "test project", "warm lead", "samples sent", "cv sent"
  ];
  return warmStatuses.some(ws => s.includes(ws));
}

// Lead Finder Safe Link Generator
function generateSearchLinks() {
  const market = document.getElementById("finderMarket").value;
  const targetType = document.getElementById("finderTargetType").value;
  const keyword = document.getElementById("finderKeyword").value.trim();
  const container = document.getElementById("searchButtonsContainer");
  
  if (!keyword) {
    container.innerHTML = `
      <div style="font-size: 12px; color: var(--color-priority-c); text-align: center; padding: 12px; background: var(--color-off-white); border-radius: var(--radius-sm);">
        Type a keyword or click a chip above to generate safe search buttons.
      </div>
    `;
    return;
  }
  
  let googleDomain = "google.com";
  if (market === "Italy") googleDomain = "google.it";
  else if (market === "UK") googleDomain = "google.co.uk";
  else if (market === "Germany") googleDomain = "google.de";
  else if (market === "Austria") googleDomain = "google.at";
  
  const targetPart = targetType !== "Other" ? `"${targetType}"` : "";
  const marketPart = market !== "Other" ? `"${market}"` : "";
  
  const igQuery = `site:instagram.com ${targetPart} "${keyword}"`.replace(/\s+/g, ' ').trim();
  const liQuery = `site:linkedin.com/in/ ${targetPart} "${keyword}"`.replace(/\s+/g, ' ').trim();
  const webQuery = `${targetPart} "${keyword}" ${marketPart}`.replace(/\s+/g, ' ').trim();
  
  const igUrl = `https://www.${googleDomain}/search?q=${encodeURIComponent(igQuery)}`;
  const liUrl = `https://www.${googleDomain}/search?q=${encodeURIComponent(liQuery)}`;
  const webUrl = `https://www.${googleDomain}/search?q=${encodeURIComponent(webQuery)}`;
  
  container.innerHTML = `
    <div style="display: grid; grid-template-columns: 1fr; gap: 10px;">
      <a href="${igUrl}" target="_blank" rel="noopener noreferrer" class="btn btn-secondary" style="display: flex; align-items: center; justify-content: space-between; font-size: 13px; text-decoration: none;">
        <span style="display: flex; align-items: center; gap: 8px;">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="lucide lucide-camera"><path d="M14.5 4h-5L7 7H4a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2h-3l-2.5-3z"/><circle cx="12" cy="13" r="3"/></svg>
          Google Instagram Search
        </span>
        <span style="font-size: 11px; color: var(--color-priority-c); font-weight: normal; margin-left: 8px;">site:instagram.com ...</span>
      </a>
      <a href="${liUrl}" target="_blank" rel="noopener noreferrer" class="btn btn-secondary" style="display: flex; align-items: center; justify-content: space-between; font-size: 13px; text-decoration: none;">
        <span style="display: flex; align-items: center; gap: 8px;">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="lucide lucide-briefcase"><path d="M16 20V4a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"/><rect width="20" height="14" x="2" y="6" rx="2"/></svg>
          Google LinkedIn Search
        </span>
        <span style="font-size: 11px; color: var(--color-priority-c); font-weight: normal; margin-left: 8px;">site:linkedin.com/in/ ...</span>
      </a>
      <a href="${webUrl}" target="_blank" rel="noopener noreferrer" class="btn btn-secondary" style="display: flex; align-items: center; justify-content: space-between; font-size: 13px; text-decoration: none;">
        <span style="display: flex; align-items: center; gap: 8px;">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="lucide lucide-globe"><circle cx="12" cy="12" r="10"/><path d="M12 2a14.5 14.5 0 0 0 0 20 14.5 14.5 0 0 0 0-20"/><path d="M2 12h20"/></svg>
          Google Web Search
        </span>
        <span style="font-size: 11px; color: var(--color-priority-c); font-weight: normal; margin-left: 8px;">"${keyword}" ...</span>
      </a>
    </div>
  `;
}

// Update keyword chips visibility based on market
function updateChipsVisibility() {
  const finderMarket = document.getElementById("finderMarket");
  const market = finderMarket ? finderMarket.value : "Italy";
  const italyGroup = document.getElementById("italyChips")?.closest("div");
  const usGroup = document.getElementById("usChips")?.closest("div");
  
  if (italyGroup && usGroup) {
    if (market === "Italy") {
      italyGroup.style.display = "block";
      usGroup.style.display = "none";
    } else if (market === "US" || market === "UK") {
      italyGroup.style.display = "none";
      usGroup.style.display = "block";
    } else {
      italyGroup.style.display = "block";
      usGroup.style.display = "block";
    }
  }
}

// Clipboard helper
function copyTextToClipboard(text) {
  if (navigator.clipboard && navigator.clipboard.writeText) {
    return navigator.clipboard.writeText(text);
  } else {
    const textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.style.position = "fixed";
    document.body.appendChild(textarea);
    textarea.focus();
    textarea.select();
    try {
      document.execCommand('copy');
    } catch (err) {
      console.error('Fallback copy failed', err);
    }
    document.body.removeChild(textarea);
    return Promise.resolve();
  }
}
// Quick Actions Dropdown Builder
function getQuickActionsDropdownHtml(lead) {
  const index = lead.originalIndex;
  
  let copySection = "";
  if (lead.email) {
    copySection += `<button class="dropdown-item" onclick="copyLeadEmail(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/></svg> Copy Email</button>`;
  }
  if (lead.whatsappNumber) {
    copySection += `<button class="dropdown-item" onclick="copyLeadWhatsApp(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/></svg> Copy WhatsApp</button>`;
  }
  if (lead.mainLink || lead.extraLink) {
    copySection += `<button class="dropdown-item" onclick="copyLeadWebsite(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/></svg> Copy Website</button>`;
  }
  
  copySection += `
    <button class="dropdown-item" onclick="copyLeadFirstMessage(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/></svg> Copy First Message</button>
  `;

  let markSection = `
    <button class="dropdown-item" onclick="markFirstMessageSent(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 6 9 17l-5-5"/></svg> Mark First Message Sent</button>
    <button class="dropdown-item" onclick="markFirstFollowupSent(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8"/><path d="M16 3h5v5"/><path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16"/><path d="M8 21H3v-5"/></svg> Mark First Follow-up Sent</button>
    <button class="dropdown-item" onclick="markSecondFollowupSent(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8"/><path d="M16 3h5v5"/><path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16"/><path d="M8 21H3v-5"/></svg> Mark Second Follow-up Sent</button>
  `;

  let channelSection = "";
  if (lead.channel === "Email") {
    channelSection = `
      <button class="dropdown-item" onclick="sendSamples(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/><polyline points="3.27 6.96 12 12.01 20.73 6.96"/><line x1="12" x2="12" y1="22.08" y2="12"/></svg> Send Samples</button>
      <button class="dropdown-item" onclick="sendCV(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="M10 9H8"/><path d="M16 13H8"/><path d="M16 17H8"/></svg> Send CV</button>
      <button class="dropdown-item" onclick="openLeadLink(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M15 3h6v6"/><path d="M10 14 21 3"/><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/></svg> Open Website</button>
    `;
  } else if (lead.channel === "WhatsApp") {
    channelSection = `
      <button class="dropdown-item" onclick="openWhatsAppChat(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M7.9 20A9 9 0 1 0 4 16.1L2 22Z"/></svg> Open WhatsApp</button>
      <button class="dropdown-item" onclick="sendSamples(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/><polyline points="3.27 6.96 12 12.01 20.73 6.96"/><line x1="12" x2="12" y1="22.08" y2="12"/></svg> Send Samples</button>
    `;
  } else if (lead.channel === "Instagram") {
    channelSection = `
      <button class="dropdown-item" onclick="openLeadLink(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M15 3h6v6"/><path d="M10 14 21 3"/><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/></svg> Open Profile</button>
      <button class="dropdown-item" onclick="markInstagramCommented(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg> Mark Commented</button>
      <button class="dropdown-item" onclick="markInstagramFollowed(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><polyline points="16 11 18 13 22 9"/></svg> Mark Followed</button>
    `;
  } else if (lead.channel === "LinkedIn") {
    channelSection = `
      <button class="dropdown-item" onclick="openLeadLink(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M15 3h6v6"/><path d="M10 14 21 3"/><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/></svg> Open Profile</button>
      <button class="dropdown-item" onclick="markLinkedInConnectionSent(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><line x1="19" x2="19" y1="8" y2="14"/><line x1="22" x2="16" y1="11" y2="11"/></svg> Mark Connection Sent</button>
    `;
  }

  const itemsHtml = `
    ${copySection}
    <div style="border-top: 1px dashed rgba(11,31,58,0.08); margin: 4px 0;"></div>
    ${markSection}
    <div style="border-top: 1px dashed rgba(11,31,58,0.08); margin: 4px 0;"></div>
    ${channelSection}
    <div style="border-top: 1px dashed rgba(11,31,58,0.08); margin: 4px 0;"></div>
    <button class="dropdown-item" onclick="archiveLead(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect width="20" height="5" x="2" y="3" rx="1"/><path d="M4 8v11a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8"/><line x1="10" x2="14" y1="12" y2="12"/></svg> Archive</button>
  `;

  return `
    <div class="quick-actions-dropdown">
      <button type="button" class="action-btn quick-btn" onclick="toggleDropdown(${index}, event)" title="Quick outreach actions" style="display: flex; align-items: center; justify-content: center; width: 26px; height: 26px; border-radius: var(--radius-sm); border: var(--border-light); background-color: var(--color-off-white); cursor: pointer; font-size: 13px; color: var(--color-royal-blue);">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="fill: currentColor;"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg>
      </button>
      <div id="dropdown-${index}" class="dropdown-content">
        ${itemsHtml}
      </div>
    </div>
  `;
}

// Toggle Quick Actions dropdown view
window.toggleDropdown = function(index, event) {
  event.stopPropagation();
  const dropdown = document.getElementById(`dropdown-${index}`);
  if (!dropdown) return;
  const wasOpen = dropdown.classList.contains("show");
  
  document.querySelectorAll(".dropdown-content.show").forEach(d => {
    d.classList.remove("show");
  });
  
  if (!wasOpen) {
    dropdown.classList.add("show");
  }
};

// Set Active Quick Filter inside Today Mode
window.setTodayFilter = function(filterVal) {
  activeTodayFilter = filterVal;
  
  const buttons = document.querySelectorAll(".today-mode-filters .quick-filter-btn");
  buttons.forEach(btn => {
    if (btn.getAttribute("data-today-filter") === filterVal) {
      btn.classList.add("active");
    } else {
      btn.classList.remove("active");
    }
  });
  
  renderTodayActions();
};

// Clipboard copy and stage marking helpers
window.copyLeadEmail = function(originalIndex) {
  const lead = leads[originalIndex];
  if (!lead || !lead.email) return;
  copyTextToClipboard(lead.email).then(() => {
    showToast("Copied", "success");
  }).catch(() => {
    showToast("Failed to copy email", "error");
  });
};

window.copyLeadWhatsApp = function(originalIndex) {
  const lead = leads[originalIndex];
  if (!lead || !lead.whatsappNumber) return;
  copyTextToClipboard(lead.whatsappNumber).then(() => {
    showToast("Copied", "success");
  }).catch(() => {
    showToast("Failed to copy WhatsApp number", "error");
  });
};

window.copyLeadWebsite = function(originalIndex) {
  const lead = leads[originalIndex];
  if (!lead) return;
  const url = lead.mainLink || lead.extraLink || "";
  if (!url) return;
  copyTextToClipboard(url).then(() => {
    showToast("Copied", "success");
  }).catch(() => {
    showToast("Failed to copy website", "error");
  });
};

window.copyLeadFirstMessage = function(originalIndex) {
  const lead = leads[originalIndex];
  if (!lead) return;
  const msg = lead.messageSent || lead.notes || "";
  if (msg) {
    copyTextToClipboard(msg).then(() => {
      showToast("Copied", "success");
    }).catch(() => {
      showToast("Failed to copy message", "error");
    });
  } else {
    showToast("No message sent yet.", "warning");
  }
};

window.markFirstMessageSent = function(originalIndex) {
  const lead = leads[originalIndex];
  if (!lead) return;
  const today = getOffsetDateString(0);
  lead.stage = "First Message Sent";
  lead.lastActionDate = today;
  lead.nextAction = "Wait";
  lead.replyStatus = "No reply";
  
  if (lead.channel === "Email") {
    lead.nextActionDate = addWorkingDays(today, 5);
  } else if (lead.channel === "LinkedIn") {
    lead.nextActionDate = addDays(today, 3);
  } else {
    lead.nextActionDate = addDays(today, 7);
  }
  
  saveData();
  updateDashboard();
  renderLeads();
  renderTodayActions();
  showToast("First message sent marked!", "success");
};

window.markFirstFollowupSent = function(originalIndex) {
  const lead = leads[originalIndex];
  if (!lead) return;
  const today = getOffsetDateString(0);
  lead.stage = "First Follow-up Sent";
  lead.followUpCount = (lead.followUpCount || 0) + 1;
  lead.lastActionDate = today;
  lead.nextAction = "Wait";
  lead.replyStatus = "No reply";
  lead.nextActionDate = addDays(today, 7);
  
  saveData();
  updateDashboard();
  renderLeads();
  renderTodayActions();
  showToast("First follow-up sent marked!", "success");
};

window.markSecondFollowupSent = function(originalIndex) {
  const lead = leads[originalIndex];
  if (!lead) return;
  const today = getOffsetDateString(0);
  lead.stage = "Second Follow-up Sent";
  lead.followUpCount = (lead.followUpCount || 0) + 1;
  lead.lastActionDate = today;
  lead.nextAction = "Wait";
  lead.replyStatus = "No reply";
  lead.nextActionDate = addDays(today, 7);
  
  saveData();
  updateDashboard();
  renderLeads();
  renderTodayActions();
  showToast("Second follow-up sent marked!", "success");
};

// Open profile or main website URL
function openLeadLink(originalIndex) {
  const lead = leads[originalIndex];
  if (!lead) return;
  const url = lead.mainLink || lead.extraLink;
  if (url) {
    window.open(normalizeUrl(url), "_blank", "noopener,noreferrer");
  } else {
    showToast("No website or profile link available.", "error");
  }
}

// Open click-to-chat WhatsApp link
function openWhatsAppChat(originalIndex) {
  const lead = leads[originalIndex];
  if (!lead) return;
  if (lead.whatsappNumber) {
    window.open(getWhatsAppLink(lead.whatsappNumber), "_blank", "noopener,noreferrer");
  } else {
    showToast("No WhatsApp number available.", "error");
  }
}

// Prompt calendar next action date
function setFollowupCalendar(originalIndex) {
  const lead = leads[originalIndex];
  if (!lead) return;
  
  const currentVal = lead.nextActionDate || getOffsetDateString(0);
  const input = prompt(`Reschedule next action date for "${lead.name}":\nEnter a date (YYYY-MM-DD) or number of days (e.g. "+3" or "3"):`, currentVal);
  
  if (input === null) return;
  
  let newDate = "";
  const cleaned = input.trim();
  if (/^\+?\d+$/.test(cleaned)) {
    const offset = parseInt(cleaned);
    newDate = addDays(getOffsetDateString(0), offset);
  } else {
    const d = new Date(cleaned);
    if (!isNaN(d.getTime())) {
      newDate = formatLocalDate(d);
    } else {
      alert("Invalid date format. Please use YYYY-MM-DD or a number of days.");
      return;
    }
  }
  
  lead.nextActionDate = newDate;
  saveData();
  updateDashboard();
  renderLeads();
  renderTodayActions();
  showToast(`Rescheduled follow-up to ${newDate}`, "success");
}

// Auto follow-up sent functions with rules
function markSentEmail(originalIndex) {
  const lead = leads[originalIndex];
  if (!lead) return;
  
  const today = getOffsetDateString(0);
  lead.lastActionDate = today;
  lead.replyStatus = "No reply";
  
  const isFirstEmail = (lead.stage === "Found (Lead collected only)" || lead.stage === "Engaged");
  if (isFirstEmail) {
    lead.stage = "First Message Sent";
    lead.followUpCount = 0;
    lead.nextAction = "Send follow-up";
    lead.nextActionDate = addWorkingDays(today, 5);
  } else {
    lead.followUpCount = (lead.followUpCount || 0) + 1;
    if (lead.followUpCount === 1) {
      lead.stage = "First Follow-up Sent";
    } else {
      lead.stage = "Second Follow-up Sent";
    }
    lead.nextAction = "Wait";
    lead.nextActionDate = addWorkingDays(today, 7); // Default wait period after follow-up sent
    
    const limit = 2; // email limit
    if (lead.followUpCount >= limit) {
      lead.nextAction = "Archive";
      alert(`Follow-up limit of ${limit} reached for ${lead.name}. Consider archiving this lead.`);
    }
  }
  
  lead.messageSent = `Sent email on ${today}`;
  
  saveData();
  updateDashboard();
  renderLeads();
  renderTodayActions();
  showToast(isFirstEmail ? "First email marked sent!" : `Follow-up email #${lead.followUpCount} marked sent!`, "success");
}

function markSentWhatsApp(originalIndex) {
  const lead = leads[originalIndex];
  if (!lead) return;
  
  const today = getOffsetDateString(0);
  lead.lastActionDate = today;
  lead.stage = "First Message Sent";
  lead.replyStatus = "No reply";
  lead.nextAction = "Send follow-up";
  lead.nextActionDate = addDays(today, 7);
  lead.messageSent = `Sent WhatsApp on ${today}`;
  
  saveData();
  updateDashboard();
  renderLeads();
  renderTodayActions();
  showToast("WhatsApp marked sent! Follow-up scheduled (+7 days).", "success");
}

function markInstagramCommented(originalIndex) {
  const lead = leads[originalIndex];
  if (!lead) return;
  
  const today = getOffsetDateString(0);
  lead.lastActionDate = today;
  
  const commentNote = `[${today}] Commented on Instagram post.`;
  if (lead.notes) {
    lead.notes = lead.notes.trim() + "\n" + commentNote;
  } else {
    lead.notes = commentNote;
  }
  
  saveData();
  updateDashboard();
  renderLeads();
  renderTodayActions();
  showToast("Instagram commented logged!", "success");
}

function markInstagramFollowed(originalIndex) {
  const lead = leads[originalIndex];
  if (!lead) return;
  
  const today = getOffsetDateString(0);
  lead.lastActionDate = today;
  
  const followNote = `[${today}] Followed Instagram profile.`;
  if (lead.notes) {
    lead.notes = lead.notes.trim() + "\n" + followNote;
  } else {
    lead.notes = followNote;
  }
  
  saveData();
  updateDashboard();
  renderLeads();
  renderTodayActions();
  showToast("Instagram follow logged!", "success");
}

function markSentDM(originalIndex) {
  const lead = leads[originalIndex];
  if (!lead) return;
  
  const today = getOffsetDateString(0);
  lead.lastActionDate = today;
  lead.stage = "First Message Sent";
  lead.replyStatus = "No reply";
  lead.nextAction = "Send follow-up";
  lead.nextActionDate = addDays(today, 7);
  lead.messageSent = `Sent Instagram DM on ${today}`;
  
  saveData();
  updateDashboard();
  renderLeads();
  renderTodayActions();
  showToast("Instagram DM marked sent! Follow-up scheduled (+7 days).", "success");
}

function markLinkedInConnectionSent(originalIndex) {
  const lead = leads[originalIndex];
  if (!lead) return;
  
  const today = getOffsetDateString(0);
  lead.lastActionDate = today;
  lead.stage = "First Message Sent";
  lead.replyStatus = "No reply";
  lead.nextAction = "Wait";
  lead.nextActionDate = addDays(today, 3);
  lead.messageSent = `Sent LinkedIn connection request on ${today}`;
  
  saveData();
  updateDashboard();
  renderLeads();
  renderTodayActions();
  showToast("LinkedIn connection sent! Next action Wait (+3 days).", "success");
}

function markSentLinkedInMessage(originalIndex) {
  const lead = leads[originalIndex];
  if (!lead) return;
  
  const today = getOffsetDateString(0);
  lead.lastActionDate = today;
  lead.stage = "First Message Sent";
  lead.replyStatus = "No reply";
  lead.nextAction = "Wait";
  lead.nextActionDate = addDays(today, 3);
  lead.messageSent = `Sent LinkedIn message on ${today}`;
  
  saveData();
  updateDashboard();
  renderLeads();
  renderTodayActions();
  showToast("LinkedIn message marked sent! Next action Wait (+3 days).", "success");
}

function markFollowupSent(originalIndex) {
  const lead = leads[originalIndex];
  if (!lead) return;
  
  const today = getOffsetDateString(0);
  lead.lastActionDate = today;
  lead.followUpCount = (lead.followUpCount || 0) + 1;
  lead.replyStatus = "No reply";
  if (lead.followUpCount === 1) {
    lead.stage = "First Follow-up Sent";
  } else {
    lead.stage = "Second Follow-up Sent";
  }
  lead.nextAction = "Wait";
  lead.nextActionDate = addDays(today, 7); // Default wait period in follow-up sent state
  
  let limit = 2;
  if (lead.channel === "WhatsApp") limit = 1;
  else if (lead.channel === "Instagram") limit = 1;
  else if (lead.channel === "LinkedIn") limit = 2;
  
  if (lead.followUpCount >= limit) {
    lead.nextAction = "Archive";
    alert(`Follow-up limit of ${limit} reached for ${lead.name}. Consider archiving this lead.`);
  }
  
  saveData();
  updateDashboard();
  renderLeads();
  renderTodayActions();
  showToast(`Follow-up #${lead.followUpCount} marked sent!`, "success");
}

function sendSamples(originalIndex) {
  const lead = leads[originalIndex];
  if (!lead) return;
  
  const today = getOffsetDateString(0);
  lead.lastActionDate = today;
  lead.stage = "Samples Sent";
  lead.replyStatus = "Samples requested";
  lead.nextAction = "Send follow-up";
  if (lead.channel === "Email") {
    lead.nextActionDate = addWorkingDays(today, 5);
  } else {
    lead.nextActionDate = addDays(today, 7);
  }
  
  saveData();
  updateDashboard();
  renderLeads();
  renderTodayActions();
  showToast("Samples marked sent!", "success");
}

function sendCV(originalIndex) {
  const lead = leads[originalIndex];
  if (!lead) return;
  
  const today = getOffsetDateString(0);
  lead.lastActionDate = today;
  lead.stage = "Samples Sent"; // Rule 2: After sending CV/portfolio/samples, stage = Samples Sent
  lead.replyStatus = "CV requested";
  lead.nextAction = "Send follow-up";
  lead.nextActionDate = addWorkingDays(today, 5);
  
  saveData();
  updateDashboard();
  renderLeads();
  renderTodayActions();
  showToast("CV marked sent!", "success");
}

// Checkbox selections helpers
function getSelectedLeadIndexes() {
  const checkboxes = document.querySelectorAll(".lead-checkbox:checked");
  const indexes = new Set();
  checkboxes.forEach(cb => {
    indexes.add(parseInt(cb.dataset.index));
  });
  return Array.from(indexes);
}

window.updateSelectedLeadsCount = function() {
  const indexes = getSelectedLeadIndexes();
  const count = indexes.length;
  const toolbar = document.getElementById("bulkActionsToolbar");
  const selectedCountBadge = document.getElementById("bulkSelectedCount");
  
  // Synchronize checkbox checked states between table view and mobile card view
  const selectedSet = new Set(indexes);
  document.querySelectorAll(".lead-checkbox").forEach(cb => {
    const idx = parseInt(cb.dataset.index);
    if (!isNaN(idx)) {
      cb.checked = selectedSet.has(idx);
    }
  });

  document.querySelectorAll(".leads-table tbody tr").forEach(tr => {
    const cb = tr.querySelector(".lead-checkbox");
    if (cb && cb.checked) {
      tr.classList.add("selected-row");
    } else {
      tr.classList.remove("selected-row");
    }
  });

  // Update Select All master checkbox state based on visible checkboxes currently rendered
  const selectAllCheckbox = document.getElementById("selectAllLeads");
  if (selectAllCheckbox) {
    const tableBody = document.getElementById("leadsTableBody");
    if (tableBody) {
      const visibleCheckboxes = tableBody.querySelectorAll(".lead-checkbox");
      if (visibleCheckboxes.length > 0) {
        const allChecked = Array.from(visibleCheckboxes).every(cb => cb.checked);
        selectAllCheckbox.checked = allChecked;
      } else {
        selectAllCheckbox.checked = false;
      }
    } else {
      selectAllCheckbox.checked = false;
    }
  }
  
  if (count > 0) {
    if (toolbar) toolbar.classList.remove("hidden");
    if (selectedCountBadge) selectedCountBadge.textContent = `${count} selected`;
  } else {
    if (toolbar) toolbar.classList.add("hidden");
  }
};

window.toggleSelectAllLeads = function(masterCheckbox) {
  const tableBody = document.getElementById("leadsTableBody");
  const cardsContainer = document.getElementById("leadsCardsContainer");
  
  if (tableBody) {
    const checkboxes = tableBody.querySelectorAll(".lead-checkbox");
    checkboxes.forEach(cb => {
      cb.checked = masterCheckbox.checked;
    });
  }
  if (cardsContainer) {
    const checkboxes = cardsContainer.querySelectorAll(".lead-checkbox");
    checkboxes.forEach(cb => {
      cb.checked = masterCheckbox.checked;
    });
  }
  updateSelectedLeadsCount();
};

window.clearBulkSelection = function() {
  const checkboxes = document.querySelectorAll(".lead-checkbox");
  checkboxes.forEach(cb => {
    cb.checked = false;
  });
  const selectAllCheckbox = document.getElementById("selectAllLeads");
  if (selectAllCheckbox) selectAllCheckbox.checked = false;
  
  const fieldSelect = document.getElementById("bulkDateFieldSelect");
  const dateInput = document.getElementById("bulkDateInput");
  if (fieldSelect) fieldSelect.value = "";
  if (dateInput) dateInput.value = "";
  
  const nextActionSelect = document.getElementById("bulkNextActionSelect");
  if (nextActionSelect) nextActionSelect.value = "";
  
  updateSelectedLeadsCount();
};

// Bulk action commands implementation
async function bulkMoveToChannel(targetChannel) {
  if (!targetChannel) return;
  const indexes = getSelectedLeadIndexes();
  if (indexes.length === 0) return;

  const selectedLeads = indexes.map(idx => leads[idx]);
  const ids = selectedLeads.map(l => l.id).filter(Boolean);

  if (ids.length === 0) {
    showToast("Error: Selected leads are missing unique IDs.", "error");
    return;
  }

  try {
    if (cloudReady) {
      setSyncStatus("syncing");
      await apiFetch("/api/leads", {
        method: "PATCH",
        body: { ids, channel: targetChannel }
      });
      setSyncStatus("connected");
    }

    // Update local state
    indexes.forEach(idx => {
      if (leads[idx]) {
        leads[idx].channel = targetChannel;
      }
    });

    // Save locally
    localStorage.setItem("ali_raza_leads", JSON.stringify(leads));

    // Refresh UI & dashboard
    updateDashboard();
    renderLeads();
    renderTodayActions();

    showToast(`Successfully moved ${ids.length} leads to ${targetChannel}!`, "success");
  } catch (err) {
    console.error("[Bulk Move] failed:", err);
    setSyncStatus("offline");
    showToast(`Failed to move leads: ${err.message || err}`, "error");
  } finally {
    const channelSelect = document.getElementById("bulkChannelSelect");
    if (channelSelect) channelSelect.value = "";
    clearBulkSelection();
  }
}

window.applyBulkChange = function(field, value) {
  if (!value) return;
  
  if (field === 'channel') {
    bulkMoveToChannel(value);
    return;
  }
  
  const indexes = getSelectedLeadIndexes();
  if (indexes.length === 0) return;
  
  indexes.forEach(idx => {
    if (leads[idx]) leads[idx][field] = value;
  });
  
  saveData();
  updateDashboard();
  renderLeads();
  renderTodayActions();
  
  if (document.getElementById("bulkStageSelect")) document.getElementById("bulkStageSelect").value = "";
  if (document.getElementById("bulkPrioritySelect")) document.getElementById("bulkPrioritySelect").value = "";
  if (document.getElementById("bulkReplySelect")) document.getElementById("bulkReplySelect").value = "";
  if (document.getElementById("bulkChannelSelect")) document.getElementById("bulkChannelSelect").value = "";
  if (document.getElementById("bulkDateFieldSelect")) document.getElementById("bulkDateFieldSelect").value = "";
  if (document.getElementById("bulkDateInput")) document.getElementById("bulkDateInput").value = "";
  if (document.getElementById("bulkNextActionSelect")) document.getElementById("bulkNextActionSelect").value = "";
  
  showToast(`Updated ${field} for ${indexes.length} leads!`, "success");
  clearBulkSelection();
};

window.applyBulkDateChange = async function() {
  const fieldSelect = document.getElementById("bulkDateFieldSelect");
  const dateInput = document.getElementById("bulkDateInput");
  
  if (!fieldSelect || !dateInput) return;
  
  const field = fieldSelect.value;
  const dateValue = dateInput.value;
  
  if (!field) {
    showToast("Please select a date field to update.", "error");
    return;
  }
  
  if (!dateValue) {
    showToast("Please select a valid date.", "error");
    return;
  }
  
  const indexes = getSelectedLeadIndexes();
  if (indexes.length === 0) {
    showToast("No leads selected.", "error");
    return;
  }
  
  const selectedLeads = indexes.map(idx => leads[idx]);
  const ids = selectedLeads.map(l => l.id).filter(Boolean);
  
  try {
    if (cloudReady && ids.length > 0) {
      setSyncStatus("syncing");
      await apiFetch("/api/leads", {
        method: "PATCH",
        body: { 
          ids, 
          updates: { [field]: dateValue }
        }
      });
      setSyncStatus("connected");
    }
    
    // Update local memory
    indexes.forEach(idx => {
      if (leads[idx]) {
        leads[idx][field] = dateValue;
      }
    });
    
    // Save to localStorage
    localStorage.setItem("ali_raza_leads", JSON.stringify(leads));
    
    // Update UI
    updateDashboard();
    renderLeads();
    renderTodayActions();
    
    showToast(`Updated date for ${indexes.length} leads!`, "success");
    
    // Reset selectors & selection
    fieldSelect.value = "";
    dateInput.value = "";
    clearBulkSelection();
  } catch (err) {
    console.error("[Bulk Date Change] failed:", err);
    setSyncStatus("offline");
    showToast(`Failed to update dates: ${err.message || err}`, "error");
  }
};

window.applyBulkNextAction = async function(value) {
  if (!value) return;
  const indexes = getSelectedLeadIndexes();
  if (indexes.length === 0) {
    showToast("No leads selected.", "error");
    return;
  }
  
  const newDate = prompt("Enter next action date (YYYY-MM-DD) or number of days (e.g. +3):");
  if (newDate === null) {
    const nextActionSelect = document.getElementById("bulkNextActionSelect");
    if (nextActionSelect) nextActionSelect.value = "";
    return;
  }
  
  let parsedDate = "";
  const cleaned = newDate.trim();
  if (/^\+?\d+$/.test(cleaned)) {
    const offset = parseInt(cleaned);
    parsedDate = addDays(getOffsetDateString(0), offset);
  } else {
    const d = new Date(cleaned);
    if (!isNaN(d.getTime())) {
      parsedDate = formatLocalDate(d);
    } else {
      alert("Invalid date format.");
      const nextActionSelect = document.getElementById("bulkNextActionSelect");
      if (nextActionSelect) nextActionSelect.value = "";
      return;
    }
  }
  
  const selectedLeads = indexes.map(idx => leads[idx]);
  const ids = selectedLeads.map(l => l.id).filter(Boolean);
  
  try {
    if (cloudReady && ids.length > 0) {
      setSyncStatus("syncing");
      await apiFetch("/api/leads", {
        method: "PATCH",
        body: { 
          ids, 
          updates: { 
            nextAction: value,
            nextActionDate: parsedDate
          }
        }
      });
      setSyncStatus("connected");
    }
    
    // Update local memory
    indexes.forEach(idx => {
      if (leads[idx]) {
        leads[idx].nextAction = value;
        leads[idx].nextActionDate = parsedDate;
      }
    });
    
    // Save locally
    localStorage.setItem("ali_raza_leads", JSON.stringify(leads));
    
    // Update UI
    updateDashboard();
    renderLeads();
    renderTodayActions();
    
    showToast(`Updated next action for ${indexes.length} leads!`, "success");
  } catch (err) {
    console.error("[Bulk Next Action] failed:", err);
    setSyncStatus("offline");
    showToast(`Failed to update next actions: ${err.message || err}`, "error");
  } finally {
    const nextActionSelect = document.getElementById("bulkNextActionSelect");
    if (nextActionSelect) nextActionSelect.value = "";
    clearBulkSelection();
  }
};

window.bulkMarkFirstMessageSent = function() {
  const indexes = getSelectedLeadIndexes();
  if (indexes.length === 0) return;
  
  const today = getOffsetDateString(0);
  
  indexes.forEach(idx => {
    const lead = leads[idx];
    if (!lead) return;
    
    lead.lastActionDate = today;
    lead.stage = "First Message Sent";
    lead.replyStatus = "No reply";
    lead.followUpCount = 0;
    
    if (lead.channel === "Email") {
      lead.nextAction = "Send follow-up";
      lead.nextActionDate = addWorkingDays(today, 5);
    } else if (lead.channel === "LinkedIn") {
      lead.nextAction = "Wait";
      lead.nextActionDate = addDays(today, 3);
    } else {
      lead.nextAction = "Send follow-up";
      lead.nextActionDate = addDays(today, 7);
    }
  });
  
  saveData();
  updateDashboard();
  renderLeads();
  renderTodayActions();
  showToast(`Marked first message sent for ${indexes.length} leads!`, "success");
  clearBulkSelection();
};

window.bulkMarkFirstFollowupSent = function() {
  const indexes = getSelectedLeadIndexes();
  if (indexes.length === 0) return;
  
  const today = getOffsetDateString(0);
  
  indexes.forEach(idx => {
    const lead = leads[idx];
    if (!lead) return;
    
    lead.lastActionDate = today;
    lead.followUpCount = (lead.followUpCount || 0) + 1;
    lead.replyStatus = "No reply";
    lead.stage = "First Follow-up Sent";
    lead.nextAction = "Wait";
    lead.nextActionDate = addDays(today, 7);
  });
  
  saveData();
  updateDashboard();
  renderLeads();
  renderTodayActions();
  
  showToast(`Marked first follow-up sent for ${indexes.length} leads!`, "success");
  clearBulkSelection();
};

window.bulkMarkSecondFollowupSent = function() {
  const indexes = getSelectedLeadIndexes();
  if (indexes.length === 0) return;
  
  const today = getOffsetDateString(0);
  let limitWarningTriggered = false;
  
  indexes.forEach(idx => {
    const lead = leads[idx];
    if (!lead) return;
    
    lead.lastActionDate = today;
    lead.followUpCount = (lead.followUpCount || 0) + 1;
    lead.replyStatus = "No reply";
    lead.stage = "Second Follow-up Sent";
    lead.nextAction = "Wait";
    lead.nextActionDate = addDays(today, 7); // Default wait period in follow-up sent state
    
    let limit = 2;
    if (lead.channel === "WhatsApp") limit = 1;
    else if (lead.channel === "Instagram") limit = 1;
    else if (lead.channel === "LinkedIn") limit = 2;
    
    if (lead.followUpCount >= limit) {
      lead.nextAction = "Archive";
      limitWarningTriggered = true;
    }
  });
  
  saveData();
  updateDashboard();
  renderLeads();
  renderTodayActions();
  
  if (limitWarningTriggered) {
    alert("One or more leads reached their follow-up limit. Consider archiving them.");
  }
  
  showToast(`Marked second follow-up sent for ${indexes.length} leads!`, "success");
  clearBulkSelection();
};

window.bulkArchiveSelected = function() {
  const indexes = getSelectedLeadIndexes();
  if (indexes.length === 0) return;
  
  const confirmArchive = confirm("Are you sure you want to archive selected leads?");
  if (!confirmArchive) return;
  
  indexes.forEach(idx => {
    if (leads[idx]) leads[idx].stage = "Archived";
  });
  
  saveData();
  updateDashboard();
  renderLeads();
  renderTodayActions();
  showToast(`Archived ${indexes.length} leads!`, "success");
  clearBulkSelection();
};

window.bulkDeleteSelected = function() {
  const indexes = getSelectedLeadIndexes();
  if (indexes.length === 0) return;
  
  const confirmDelete = confirm("Are you sure you want to permanently delete selected leads? This cannot be undone.");
  if (!confirmDelete) return;
  
  indexes.sort((a, b) => b - a);
  
  indexes.forEach(idx => {
    leads.splice(idx, 1);
  });
  
  saveData();
  updateDashboard();
  renderLeads();
  renderTodayActions();
  showToast(`Deleted ${indexes.length} leads!`, "success");
  clearBulkSelection();
};

window.bulkExportCSV = function() {
  const indexes = getSelectedLeadIndexes();
  if (indexes.length === 0) return;
  
  const selectedLeads = indexes.map(idx => leads[idx]).filter(Boolean);
  if (selectedLeads.length === 0) return;
  
  const headers = [
    "Date Added", "Lead Name/Company", "Contact Person", "Market", 
    "Channel", "Main Link", "Extra Link", "Niche", "Source", "Priority", 
    "Stage", "Last Action Date", "Next Action", "Next Action Date", 
    "Reply Status", "Notes", "Message Sent", "Follow-up Count"
  ];
  
  const rows = selectedLeads.map(lead => [
    lead.dateAdded || "",
    lead.name || "",
    lead.contactPerson || "",
    lead.market || "",
    lead.channel || "",
    lead.mainLink || "",
    lead.extraLink || "",
    lead.niche || "",
    lead.source || "",
    lead.priority || "",
    lead.stage || "",
    lead.lastActionDate || "",
    lead.nextAction || "",
    lead.nextActionDate || "",
    lead.replyStatus || "",
    lead.notes || "",
    lead.messageSent || "",
    lead.followUpCount !== undefined ? lead.followUpCount : 0
  ]);
  
  let csvContent = "data:text/csv;charset=utf-8," 
    + [headers.join(","), ...rows.map(r => r.map(val => escapeCsvValue(val)).join(","))].join("\n");
    
  const encodedUri = encodeURI(csvContent);
  const link = document.createElement("a");
  link.setAttribute("href", encodedUri);
  link.setAttribute("download", `outreach_leads_selected_${getOffsetDateString(0)}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  
  showToast(`Exported ${selectedLeads.length} leads to CSV!`, "success");
  clearBulkSelection();
};

// Dismiss dropdowns when clicking outside
document.addEventListener("click", (e) => {
  if (!e.target.closest(".quick-actions-dropdown") && !e.target.closest(".column-visibility-dropdown")) {
    document.querySelectorAll(".dropdown-content.show").forEach(d => {
      d.classList.remove("show");
    });
  }
});

// Expose functions to window
window.addDays = addDays;
window.openLeadLink = openLeadLink;
window.openWhatsAppChat = openWhatsAppChat;
window.setFollowupCalendar = setFollowupCalendar;
window.markSentEmail = markSentEmail;
window.markSentWhatsApp = markSentWhatsApp;
window.markInstagramCommented = markInstagramCommented;
window.markInstagramFollowed = markInstagramFollowed;
window.markSentDM = markSentDM;
window.markLinkedInConnectionSent = markLinkedInConnectionSent;
window.markSentLinkedInMessage = markSentLinkedInMessage;
window.markFollowupSent = markFollowupSent;
window.sendSamples = sendSamples;
window.sendCV = sendCV;
window.toggleDropdown = toggleDropdown;
window.setTodayFilter = setTodayFilter;
window.toggleSelectAllLeads = toggleSelectAllLeads;
window.updateSelectedLeadsCount = updateSelectedLeadsCount;
window.clearBulkSelection = clearBulkSelection;
window.applyBulkChange = applyBulkChange;
window.applyBulkNextAction = applyBulkNextAction;
window.bulkMarkFirstMessageSent = bulkMarkFirstMessageSent;
window.bulkMarkFirstFollowupSent = bulkMarkFirstFollowupSent;
window.bulkMarkSecondFollowupSent = bulkMarkSecondFollowupSent;
window.bulkArchiveSelected = bulkArchiveSelected;
window.bulkDeleteSelected = bulkDeleteSelected;
window.bulkExportCSV = bulkExportCSV;

// Backup operations exports
window.exportBackupJSON = exportBackupJSON;
window.handleBackupFileSelect = handleBackupFileSelect;
window.showBackupImportPreview = showBackupImportPreview;
window.updateBackupPreviewCounts = updateBackupPreviewCounts;
window.confirmBackupImport = confirmBackupImport;
window.closeBackupImportModal = closeBackupImportModal;


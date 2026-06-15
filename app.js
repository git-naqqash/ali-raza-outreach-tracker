/* D:\Lead Tracker\app.js */

// Authentication Credentials
const LOGIN_USER = "contact.naqqash@gmail.com";
const LOGIN_PASS = "@@@03314200250";

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
  
  if (typeof val === 'number' || (!isNaN(val) && !isNaN(parseFloat(val)))) {
    const serial = parseFloat(val);
    const utc_days = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);
    const localDate = new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate());
    return formatLocalDate(localDate);
  }

  const str = String(val).trim();
  if (!str) return null;

  const parsedMs = Date.parse(str);
  if (!isNaN(parsedMs)) {
    return formatLocalDate(new Date(parsedMs));
  }

  // Handle formats like "25 may 2026" or "13-May-2026"
  const wordMonthRegex = /^(\d{1,2})\s+([a-zA-Z]+)\s+(\d{4})/i;
  const match = str.match(wordMonthRegex);
  if (match) {
    const day = parseInt(match[1]);
    const monthStr = match[2].toLowerCase().substring(0, 3);
    const year = parseInt(match[3]);
    const months = {
      jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11
    };
    if (months[monthStr] !== undefined) {
      const d = new Date(year, months[monthStr], day);
      if (!isNaN(d.getTime())) {
        return formatLocalDate(d);
      }
    }
  }

  // Handle format with text like "5/13/2026, got replied..."
  const dateRegex = /(\d{1,2})[\/\-]\d{1,2}[\/\-]\d{4}/;
  const dateMatch = str.match(dateRegex);
  if (dateMatch) {
    const d = new Date(dateMatch[0]);
    if (!isNaN(d.getTime())) {
      return formatLocalDate(d);
    }
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

// Default seed data for Message Scripts
const DEFAULT_SCRIPTS = [
  {
    id: "default-ig-dm",
    title: "Instagram DM (Found Pitch)",
    channel: "Instagram",
    type: "First Message",
    body: "Ciao [Name]! Love your content on [Topic], especially your post about [Reference]. I help Italian coaches package their expertise into professional PDF workbooks and lead magnets to scale their client acquisition. Just wanted to see if you have any books/lead magnets planned for this quarter?",
    isDefault: true
  },
  {
    id: "default-li-conn",
    title: "LinkedIn Connection Request",
    channel: "LinkedIn",
    type: "Connection Request",
    body: "Hi [Name], came across your profile and noticed your focus on [Niche]. I write white-label ebooks and interactive workbooks for business agencies and publishing teams in the [Market] region. Would love to connect and keep in touch!",
    isDefault: true
  },
  {
    id: "default-li-after",
    title: "LinkedIn Message (After Accepting)",
    channel: "LinkedIn",
    type: "After Accepting",
    body: "Thanks for connecting, [Name]! Many agencies and course creators struggle to write high-quality books/lead magnets that engage readers and drive backend sales. I outline, structure, and write white-label books so you can publish faster under your name. Have you ever considered launching an ebook or Canva workbook for your audience?",
    isDefault: true
  },
  {
    id: "default-email-en",
    title: "English First Email (Agencies/Coaches)",
    channel: "Email",
    type: "First Message",
    body: "Subject: Ebook writing & Canva workbook support for [Company]\n\nHi [Name],\n\nI came across [Company] and love the work you are doing for clients in [Niche]. \n\nI work as a freelance white-label book writer. I help publishing teams, content agencies, and coaches design and draft high-conversion ebooks, workbook PDFs, and lead magnets. This lets you offer publishing services to your clients or grow your own backend program without doing any writing yourself.\n\nI\u2019d love to send over a brief PDF sample workbook that shows my writing style and Canva layouts. \n\nWould it be okay to share the link with you?\n\nBest regards,\nAli Raza\nEbook & Lead Magnet Writer\nhttps://alirazawriter.com (Placeholder link)",
    isDefault: true
  },
  {
    id: "default-email-it",
    title: "Italian First Email (Primo Contatto)",
    channel: "Email",
    type: "First Message",
    body: "Oggetto: Supporto stesura Ebook e Workbook Canva per [Company]\n\nGentile [Name],\n\nHo visto il vostro lavoro su [Company] e ho pensato di scrivervi.\n\nLavoro come scrittore freelance white-label, specializzato nella creazione di ebook, workbook interattivi Canva e lead magnet per coach, agenzie e case editrici. Scrivo i testi in inglese o italiano e realizzo il layout grafico, consentendovi di pubblicare a vostro nome o offrire questo servizio ai clienti senza occuparvene internamente.\n\nPosso inviarvi una breve anteprima dei miei lavori e dei workbook Canva realizzati per altri progetti simili?\n\nUn cordiale saluto,\nAli Raza\nScrittore di Ebook & Lead Magnet\nhttps://alirazawriter.com",
    isDefault: true
  },
  {
    id: "default-email-follow",
    title: "Email Follow-up (Polite Nudge)",
    channel: "Email",
    type: "Follow-up",
    body: "Subject: Re: Ebook writing & Canva workbook support for [Company]\n\nHi [Name],\n\nI wanted to quickly bump this in case it got buried in your inbox. I know you're busy growing [Company].\n\nJust to recap: I write and design high-quality white-label ebooks and workbooks so you can deliver premium publishing content to your clients or scale your backend funnel.\n\nIf you have 2 minutes, I'd love to drop a quick link to a sample workbook. Would it be worth checking out?\n\nBest,\nAli Raza",
    isDefault: true
  },
  {
    id: "default-wa-first",
    title: "WhatsApp First Message (Public Contacts Only)",
    channel: "WhatsApp",
    type: "First Message",
    body: "Hi [Name]! Ali Raza here. Found your business number on your page. I specialize in writing white-label ebooks and interactive workbooks for coaches & course creators. I noticed you publish digital content and wanted to ask if you currently need support outsourcing workbook copywriting or Canva formatting? Appreciate your time!",
    isDefault: true
  },
  {
    id: "default-wa-follow",
    title: "WhatsApp Follow-up",
    channel: "WhatsApp",
    type: "Follow-up",
    body: "Hi [Name], just checking if you had a moment to see my previous message about ebook/workbook support? No pressure at all\u2014happy to share a quick PDF sample if you ever want to expand your digital offerings. Have a great day!",
    isDefault: true
  },
  {
    id: "default-italian-sample-reply",
    title: "Italian Sample Reply",
    channel: "Email",
    type: "Sample Reply",
    body: "Buongiorno [Nome],\n\ncerto, posso inviarvi 2 sample per valutare struttura, stile e qualit\u00e0 del lavoro.\n\nPer comodit\u00e0, qui trova anche il mio portfolio:\nlinktr.ee/aliraza.ebooks\n\nResto disponibile se desiderate un esempio pi\u00f9 vicino al vostro tipo di progetto.\n\nCordiali saluti,\nAli Raza",
    isDefault: true
  },
  {
    id: "default-english-sample-reply",
    title: "English Sample Reply",
    channel: "Email",
    type: "Sample Reply",
    body: "Hi [Name],\n\nSure, I can send 2 relevant samples so you can review the structure, writing style, and quality.\n\nYou can also view my portfolio here:\nlinktr.ee/aliraza.ebooks\n\nIf needed, I can share a sample closer to your type of project.\n\nBest regards,\nAli Raza",
    isDefault: true
  },
  {
    id: "default-italian-cv-reply",
    title: "Italian CV Sent Reply",
    channel: "Email",
    type: "CV Reply",
    body: "Buongiorno [Nome],\n\ngrazie mille per la risposta.\n\nLe invio il mio CV e i dettagli della mia attivit\u00e0.\n\nLavoro come supporto white-label per contenuti in inglese: ebook, workbook, Canva book e lead magnet per team editoriali, self-publishing, coach, agenzie e content business.\n\nPer comodit\u00e0, qui trova anche il mio portfolio:\nlinktr.ee/aliraza.ebooks\n\nResto disponibile anche per una piccola prova gratuita, cos\u00ec potete valutare qualit\u00e0, struttura e tempi di consegna.\n\nCordiali saluti,\nAli Raza",
    isDefault: true
  },
  {
    id: "default-price-reply",
    title: "Price Reply",
    channel: "General",
    type: "Price Reply",
    body: "Hi [Name],\n\nThe price depends on length, structure, niche, deadline, and whether formatting or Canva layout is included.\n\nIf you send me the topic, word count or page count, outline, and deadline, I can give you a clear quote.\n\nBest regards,\nAli Raza",
    isDefault: true
  },
  {
    id: "default-not-now-reply",
    title: "Not Now Reply",
    channel: "General",
    type: "Not Now Reply",
    body: "No problem at all, thank you for letting me know.\n\nI\u2019ll stay available if you ever need extra white-label support for ebook, workbook, Canva book, or lead magnet projects in the future.\n\nBest regards,\nAli Raza",
    isDefault: true
  }
];

// Scripts State variables
let scripts = [];
let editingScriptId = null;

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

// ── Sync all current scripts to cloud ────────────────────────
async function syncScriptsToCloud() {
  if (!cloudReady) return;
  try {
    await apiFetch("/api/scripts", { method: "POST", body: { scripts } });
  } catch (err) {
    console.error("[Cloud] syncScriptsToCloud error:", err.message);
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
    const localScripts = JSON.parse(localStorage.getItem("ali_raza_scripts") || "[]");

    localLeads.forEach(ensureLeadId);
    localScripts.forEach(s => { if (!s.id) s.id = generateId(); });
    localStorage.setItem("ali_raza_leads",   JSON.stringify(localLeads));
    localStorage.setItem("ali_raza_scripts", JSON.stringify(localScripts));

    await apiFetch("/api/leads", { method: "POST", body: { leads: localLeads } });

    if (localScripts.length > 0) {
      await apiFetch("/api/scripts", { method: "POST", body: { scripts: localScripts } }).catch(e =>
        console.warn("[Cloud] scripts upload warning:", e.message)
      );
    }

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

    const scriptsData = await apiFetch("/api/scripts");
    if (scriptsData.scripts && scriptsData.scripts.length > 0) {
      scripts = scriptsData.scripts;
      localStorage.setItem("ali_raza_scripts", JSON.stringify(scripts));
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
  renderScripts();
}

// ============================================================
// END NEON DATABASE SYNC ENGINE
// ============================================================

// On Page Load
document.addEventListener("DOMContentLoaded", () => {
  loadData();            // Immediately load localStorage (fast, no flicker)
  setupEventListeners();
  initStorageNotice();   // Hide storage notice if previously dismissed
  initTableResizableColumns(); // Initialize resizable columns first
  initColumnVisibility(); // Initialize column visibility from preferences
  updateDashboard();
  renderLeads();
  renderTodayActions();
  renderScripts();
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

  // Highlight the last visible th to adjust resizer positioning and avoid scroll container clipping
  const table = document.querySelector("table.leads-table");
  if (table) {
    const ths = table.querySelectorAll("thead th");
    ths.forEach(th => th.classList.remove("last-visible-th"));
    const visibleThs = Array.from(ths).filter(th => {
      return th.style.display !== "none";
    });
    if (visibleThs.length > 0) {
      visibleThs[visibleThs.length - 1].classList.add("last-visible-th");
    }
  }

  adjustTableWidthToColumns(table);
}

function getMinColumnWidth(colId) {
  if (colId === "col-name") return 180;
  if (colId === "col-contact") return 120;
  if (colId === "col-market") return 90;
  if (colId === "col-niche") return 90;
  if (colId === "col-source") return 80;
  if (colId === "col-priority") return 80;
  if (colId === "col-stage") return 100;
  if (colId === "col-nextAction") return 130;
  if (colId === "col-nextActionDate") return 110;
  if (colId === "col-replyStatus") return 100;
  if (colId === "col-notes") return 160;
  if (colId === "col-actions") return 120;
  return 40; // Checkbox column fallback
}

function getDefaultColumnWidth(colId) {
  if (colId === "col-name") return 180;
  if (colId === "col-contact") return 140;
  if (colId === "col-market") return 100;
  if (colId === "col-niche") return 100;
  if (colId === "col-source") return 100;
  if (colId === "col-priority") return 90;
  if (colId === "col-stage") return 100;
  if (colId === "col-nextAction") return 140;
  if (colId === "col-nextActionDate") return 120;
  if (colId === "col-replyStatus") return 110;
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

  // Load Scripts
  try {
    const storedScripts = localStorage.getItem("ali_raza_scripts");
    if (storedScripts) {
      scripts = JSON.parse(storedScripts);
    } else {
      scripts = JSON.parse(JSON.stringify(DEFAULT_SCRIPTS));
      saveScripts();
    }
  } catch (err) {
    console.error("Error parsing scripts from localStorage, resetting to defaults", err);
    scripts = JSON.parse(JSON.stringify(DEFAULT_SCRIPTS));
    saveScripts();
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
  const filterPriority = document.getElementById("filterPriority").value;
  const filterStage = document.getElementById("filterStage").value;
  const filterReply = document.getElementById("filterReply").value;
  const filterActionDate = document.getElementById("filterActionDate").value;
  const filterNextAction = document.getElementById("filterNextAction").value;

  return leads.filter((lead, index) => {
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

    return true;
  });
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

function getReplyBadge(status) {
  const cssClass = status.toLowerCase().replace(/ /g, '-');
  return `<span class="reply-status-text ${cssClass}">${status}</span>`;
}

function renderNotesCellContent(notes, originalIndex) {
  if (!notes) return '<span class="notes-empty">-</span>';
  const cleanNotes = notes.trim();
  if (!cleanNotes) return '<span class="notes-empty">-</span>';
  
  if (cleanNotes.length <= 80) {
    return `<span class="notes-text">${escapeHtml(cleanNotes)}</span>`;
  }
  
  const shortNotes = cleanNotes.substring(0, 75) + "...";
  return `
    <span class="notes-text-short" id="notes-short-${originalIndex}">${escapeHtml(shortNotes)}</span>
    <span class="notes-text-full" id="notes-full-${originalIndex}" style="white-space: pre-wrap; display: none;">${escapeHtml(cleanNotes)}</span>
    <button type="button" class="notes-toggle-btn" onclick="toggleNotes(${originalIndex}, event)">View more</button>
  `;
}

window.toggleNotes = function(index, event) {
  if (event) {
    event.preventDefault();
    event.stopPropagation();
  }
  const shortSpan = document.getElementById(`notes-short-${index}`);
  const fullSpan = document.getElementById(`notes-full-${index}`);
  const btn = event.target;
  
  if (shortSpan && fullSpan && btn) {
    const isExpanded = shortSpan.style.display === "none";
    if (isExpanded) {
      shortSpan.style.display = "inline";
      fullSpan.style.display = "none";
      btn.textContent = "View more";
    } else {
      shortSpan.style.display = "none";
      fullSpan.style.display = "inline";
      btn.textContent = "View less";
    }
  }
};

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
      tabActions.style.display = "block";
      btnText.textContent = `Import ${activeTab} Leads`;
    }
  }

  const filtered = getFilteredLeads();
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
        <td colspan="9" style="text-align: center;">
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
          <button class="btn btn-secondary" onclick="copyPersonalizedScript(${lead.originalIndex})" title="Copy customized outreach email message"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/></svg> Copy Script</button>
          <button class="btn btn-primary" onclick="markSentEmail(${lead.originalIndex})" title="Mark email sent & reschedule follow-up"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg> Mark Sent</button>
          <button class="btn btn-secondary" onclick="openLeadLink(${lead.originalIndex})" title="Open lead's website/link"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 3h6v6"/><path d="M10 14 21 3"/><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/></svg> Open Link</button>
          <button class="btn btn-secondary" onclick="setFollowupCalendar(${lead.originalIndex})" title="Reschedule next action date"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="18" height="18" x="3" y="4" rx="2" ry="2"/><line x1="16" x2="16" y1="2" y2="6"/><line x1="8" x2="8" y1="2" y2="6"/><line x1="3" x2="21" y1="10" y2="10"/></svg> Reschedule</button>
          <button class="btn btn-danger-outline" onclick="archiveLead(${lead.originalIndex})" title="Archive lead"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="20" height="5" x="2" y="3" rx="1"/><path d="M4 8v11a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8"/><line x1="10" x2="14" y1="12" y2="12"/></svg> Archive</button>
        `;
      } else if (lead.channel === "WhatsApp") {
        buttonsHtml = `
          <button class="btn btn-secondary" onclick="copyPersonalizedScript(${lead.originalIndex})" title="Copy customized WhatsApp pitch message"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/></svg> Copy Script</button>
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
          <button class="btn btn-secondary" onclick="copyPersonalizedScript(${lead.originalIndex})" title="Copy customized Instagram DM script"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/></svg> Copy DM</button>
          <button class="btn btn-primary" onclick="markSentDM(${lead.originalIndex})" title="Mark Instagram DM sent & reschedule follow-up"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg> DM Sent</button>
          <button class="btn btn-danger-outline" onclick="archiveLead(${lead.originalIndex})" title="Archive lead"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="20" height="5" x="2" y="3" rx="1"/><path d="M4 8v11a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8"/><line x1="10" x2="14" y1="12" y2="12"/></svg> Archive</button>
        `;
      } else if (lead.channel === "LinkedIn") {
        buttonsHtml = `
          <button class="btn btn-secondary" onclick="openLeadLink(${lead.originalIndex})" title="Open LinkedIn profile in new tab"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 3h6v6"/><path d="M10 14 21 3"/><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/></svg> Open LI</button>
          <button class="btn btn-secondary" onclick="markLinkedInConnectionSent(${lead.originalIndex})" title="Mark connection request sent & wait 3 days"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><line x1="19" x2="19" y1="8" y2="14"/><line x1="22" x2="16" y1="11" y2="11"/></svg> Conn Sent</button>
          <button class="btn btn-secondary" onclick="copyPersonalizedScript(${lead.originalIndex})" title="Copy customized LinkedIn script"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/></svg> Copy Script</button>
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
  document.getElementById("leadModal").classList.add("active");
}

// Modal open: Edit Mode
function openEditModal(index) {
  const lead = leads[index];
  
  document.getElementById("leadIndex").value = index;
  document.getElementById("leadDateAdded").value = lead.dateAdded || "";
  document.getElementById("leadName").value = lead.name || "";
  document.getElementById("leadMarket").value = lead.market || "Italy";
  document.getElementById("leadChannel").value = lead.channel || "Instagram";
  document.getElementById("leadMainLink").value = lead.mainLink || "";
  document.getElementById("leadNiche").value = lead.niche || "";
  document.getElementById("leadSource").value = lead.source || "Other";
  document.getElementById("leadPriority").value = lead.priority || "A";
  document.getElementById("leadStage").value = lead.stage || "Found";
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

  // --- Scripts Section Toolbar & Accordion Event Delegation ---

  // Search & Filter listeners
  const scriptSearchInput = document.getElementById("scriptSearchInput");
  const filterScriptChannel = document.getElementById("filterScriptChannel");
  const filterScriptType = document.getElementById("filterScriptType");

  if (scriptSearchInput) {
    scriptSearchInput.addEventListener("input", renderScripts);
  }
  if (filterScriptChannel) {
    filterScriptChannel.addEventListener("change", renderScripts);
  }
  if (filterScriptType) {
    filterScriptType.addEventListener("change", renderScripts);
  }

  // Delegated Accordion Events (Toggle, Copy, Edit, Cancel, Save, Delete)
  const scriptsAccordion = document.getElementById("scriptsAccordion");
  if (scriptsAccordion) {
    scriptsAccordion.addEventListener("click", (e) => {
      // 1. Accordion Trigger (Toggle Open/Close)
      const trigger = e.target.closest(".accordion-trigger");
      if (trigger) {
        const parent = trigger.closest(".accordion-item");
        const scriptId = parent.dataset.id;
        
        // If this item is being edited, prevent toggle collapse
        if (editingScriptId === scriptId) return;
        
        const isActive = parent.classList.contains("active");
        
        // Close all other accordion items
        document.querySelectorAll("#scriptsAccordion .accordion-item").forEach(item => {
          if (item !== parent) {
            item.classList.remove("active");
            const content = item.querySelector(".accordion-content");
            if (content) content.style.maxHeight = null;
          }
        });
        
        parent.classList.toggle("active");
        const content = parent.querySelector(".accordion-content");
        if (parent.classList.contains("active")) {
          content.style.maxHeight = content.scrollHeight + "px";
        } else {
          content.style.maxHeight = null;
        }
        return;
      }
      
      // 2. Copy Button
      const copyBtn = e.target.closest(".copy-btn");
      if (copyBtn) {
        const scriptId = copyBtn.dataset.id;
        const script = scripts.find(s => s.id === scriptId);
        if (script) {
          navigator.clipboard.writeText(script.body).then(() => {
            copyBtn.textContent = "Copied!";
            copyBtn.classList.add("btn-success");
            copyBtn.classList.remove("btn-secondary");
            
            setTimeout(() => {
              copyBtn.textContent = "Copy Script";
              copyBtn.classList.remove("btn-success");
              copyBtn.classList.add("btn-secondary");
            }, 2000);
            
            showToast("Script body copied!", "success");
          }).catch(() => {
            showToast("Failed to copy script.", "error");
          });
        }
        return;
      }
      
      // 3. Edit Button
      const editBtn = e.target.closest(".edit-btn");
      if (editBtn) {
        const scriptId = editBtn.dataset.id;
        editingScriptId = scriptId;
        renderScripts();
        
        // Ensure the accordion item is active and has full height visible
        const parent = editBtn.closest(".accordion-item");
        parent.classList.add("active");
        const content = parent.querySelector(".accordion-content");
        if (content) content.style.maxHeight = "none";
        return;
      }
      
      // 4. Cancel Edit Button
      const cancelEditBtn = e.target.closest(".cancel-edit-btn");
      if (cancelEditBtn) {
        editingScriptId = null;
        renderScripts();
        return;
      }
      
      // 5. Save Edit Button
      const saveEditBtn = e.target.closest(".save-edit-btn");
      if (saveEditBtn) {
        const scriptId = saveEditBtn.dataset.id;
        const parent = saveEditBtn.closest(".accordion-item");
        
        const titleInput = parent.querySelector(".edit-script-title");
        const channelSelect = parent.querySelector(".edit-script-channel");
        const typeSelect = parent.querySelector(".edit-script-type");
        const bodyTextarea = parent.querySelector(".edit-script-body");
        
        const updatedTitle = titleInput.value.trim();
        const updatedChannel = channelSelect.value;
        const updatedType = typeSelect.value;
        const updatedBody = bodyTextarea.value;
        
        if (!updatedTitle || !updatedBody) {
          showToast("Title and Body are required.", "error");
          return;
        }
        
        const scriptIndex = scripts.findIndex(s => s.id === scriptId);
        if (scriptIndex !== -1) {
          scripts[scriptIndex].title = updatedTitle;
          scripts[scriptIndex].channel = updatedChannel;
          scripts[scriptIndex].type = updatedType;
          scripts[scriptIndex].body = updatedBody;
          
          saveScripts();
          editingScriptId = null;
          renderScripts();
          showToast("Script updated successfully!", "success");
        }
        return;
      }
      
      // 6. Delete Button
      const deleteBtn = e.target.closest(".delete-btn");
      if (deleteBtn) {
        const scriptId = deleteBtn.dataset.id;
        const script = scripts.find(s => s.id === scriptId);
        if (script) {
          const confirmed = confirm("Are you sure you want to delete this script?");
          if (confirmed) {
            scripts = scripts.filter(s => s.id !== scriptId);
            saveScripts();
            if (editingScriptId === scriptId) {
              editingScriptId = null;
            }
            renderScripts();
            showToast(`Deleted script "${script.title}"`, "success");
          }
        }
        return;
      }
    });
  }

  // "+ Add New Script" Modal Controls
  const addNewScriptBtn = document.getElementById("addNewScriptBtn");
  const scriptModal = document.getElementById("scriptModal");
  const closeScriptModalBtn = document.getElementById("closeScriptModalBtn");
  const cancelScriptModalBtn = document.getElementById("cancelScriptModalBtn");
  const scriptForm = document.getElementById("scriptForm");

  if (addNewScriptBtn && scriptModal) {
    addNewScriptBtn.addEventListener("click", () => {
      if (scriptForm) scriptForm.reset();
      scriptModal.classList.add("active");
    });
  }
  
  const closeScriptModal = () => {
    if (scriptModal) scriptModal.classList.remove("active");
  };

  if (closeScriptModalBtn) {
    closeScriptModalBtn.addEventListener("click", closeScriptModal);
  }
  if (cancelScriptModalBtn) {
    cancelScriptModalBtn.addEventListener("click", closeScriptModal);
  }
  
  // Close script modal on overlay click
  if (scriptModal) {
    scriptModal.addEventListener("click", (e) => {
      if (e.target.id === "scriptModal") closeScriptModal();
    });
  }

  // Script Modal Form Submission
  if (scriptForm) {
    scriptForm.addEventListener("submit", (e) => {
      e.preventDefault();
      
      const newTitle = document.getElementById("newScriptTitle").value.trim();
      const newChannel = document.getElementById("newScriptChannel").value;
      const newType = document.getElementById("newScriptType").value;
      const newBody = document.getElementById("newScriptBody").value;
      
      if (!newTitle || !newBody) {
        showToast("Title and Body are required.", "error");
        return;
      }
      
      const newScript = {
        id: "custom-" + Date.now(),
        title: newTitle,
        channel: newChannel,
        type: newType,
        body: newBody,
        isDefault: false
      };
      
      scripts.unshift(newScript); // Add custom to beginning
      saveScripts();
      closeScriptModal();
      renderScripts();
      showToast(`Added script "${newTitle}"`, "success");
    });
  }

  // "Reset Default Scripts" Action
  const resetScriptsBtn = document.getElementById("resetScriptsBtn");
  if (resetScriptsBtn) {
    resetScriptsBtn.addEventListener("click", () => {
      const confirmed = confirm("This will restore the original default scripts. Your custom scripts may be kept separately.");
      if (confirmed) {
        // Keep custom scripts
        const customScripts = scripts.filter(s => !s.isDefault);
        
        // Pristine default copy
        const pristineDefaults = JSON.parse(JSON.stringify(DEFAULT_SCRIPTS));
        
        // defaults first, then custom ones
        scripts = [...pristineDefaults, ...customScripts];
        
        saveScripts();
        editingScriptId = null;
        renderScripts();
        showToast("Default scripts restored!", "success");
      }
    });
  }

  // CSV Export trigger
  exportBtn.addEventListener("click", exportToCSV);

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

  // Checkbox listeners in import preview modal
  const chkImportDupAnyway = document.getElementById("chkImportDupAnyway");
  if (chkImportDupAnyway) {
    chkImportDupAnyway.addEventListener("change", updateImportPreviewCounts);
  }
  const chkUpdateExistingDup = document.getElementById("chkUpdateExistingDup");
  if (chkUpdateExistingDup) {
    chkUpdateExistingDup.addEventListener("change", updateImportPreviewCounts);
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
        stage: "Found",
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

// --- IMPORT LEADS IMPLEMENTATION ---

const SYNONYMS = {
  name: ["lead", "company", "company name", "business name", "name", "profile name", "account"],
  mainLink: ["website", "link", "url", "main link", "profile", "profile link", "website/profile", "instagram", "linkedin", "instagram link", "linkedin link"],
  contactPerson: ["contact person", "founder", "owner", "person name"],
  email: ["email", "email address", "contact email", "mail", "e-mail"],
  whatsappNumber: ["whatsapp", "whatsapp number", "phone", "phone number", "mobile", "contact number", "numero", "telefono"],
  market: ["market", "country", "location"],
  niche: ["niche", "category", "industry", "lead type", "type"],
  source: ["source", "lead source"],
  priority: ["priority"],
  stage: ["stage", "status"],
  nextAction: ["next action"],
  nextActionDate: ["next action date", "follow-up date", "next follow-up date"],
  replyStatus: ["reply status", "result"],
  notes: ["notes", "remarks", "comments", "context"],
  messageSent: ["message sent", "personalized first message", "first message", "dm", "email message", "whatsapp message"]
};

let pendingValidLeads = [];
let pendingDuplicateLeads = [];

function mapHeaders(headers) {
  const mapping = {};
  headers.forEach((header, idx) => {
    if (header === undefined || header === null) return;
    const cleanHeader = header.toString().trim().toLowerCase();
    for (const [key, synonyms] of Object.entries(SYNONYMS)) {
      if (synonyms.includes(cleanHeader)) {
        if (mapping[key] === undefined) {
          mapping[key] = idx;
        }
      }
    }
  });
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
    
    let totalRows = 0;
    let missingInfoCount = 0;
    
    const validLeads = [];
    const duplicateLeads = []; // Array of { lead, matchedIndex }
    
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
      
      const rowName = getVal("name");
      const rowMainLink = getVal("mainLink");
      const rowContactPerson = getVal("contactPerson");
      const rowEmail = getVal("email");
      const rowWhatsapp = getVal("whatsappNumber");
      const rowMarket = getVal("market");
      const rowNiche = getVal("niche");
      const rowSource = getVal("source");
      const rowPriority = getVal("priority");
      const rowStage = getVal("stage");
      const rowNextAction = getVal("nextAction");
      const rowNextActionDate = getVal("nextActionDate");
      const rowReplyStatus = getVal("replyStatus");
      const rowNotes = getVal("notes");
      const rowMessageSent = getVal("messageSent");
      
      // Skip row if it has absolutely no contact details
      if (!rowName && !rowMainLink && !rowContactPerson && !rowEmail && !rowWhatsapp) {
        return;
      }
      
      // Check if main link or lead name is missing
      const isMissingInfo = !rowMainLink || !rowName;
      if (isMissingInfo) {
        missingInfoCount++;
      }
      
      // Parse dates from file
      const parsedDate = parseExcelDate(rowNextActionDate);
      
      // Check red row color
      let isRedRow = false;
      const colIndices = Object.values(colMapping);
      for (const colIdx of colIndices) {
        const cellAddress = XLSX.utils.encode_cell({ r: rowIdx + 1, c: colIdx });
        const cell = worksheet[cellAddress];
        if (cell && cell.s) {
          if (checkRowRedStyle(cell)) {
            isRedRow = true;
            break;
          }
        }
      }
      
      // Base lead mapping
      const lead = {
        dateAdded: getOffsetDateString(0),
        name: rowName ? String(rowName).trim() : "Unnamed Lead / Company",
        market: rowMarket ? String(rowMarket).trim() : "Italy",
        channel: activeTab, // Assigned to current tab
        mainLink: rowMainLink ? normalizeUrl(String(rowMainLink)) : "",
        niche: rowNiche ? String(rowNiche).trim() : "",
        source: rowSource ? String(rowSource).trim() : "Import",
        priority: rowPriority ? String(rowPriority).trim() : "B",
        stage: rowStage ? String(rowStage).trim() : "Found",
        lastActionDate: getOffsetDateString(0),
        nextAction: rowNextAction ? String(rowNextAction).trim() : "",
        nextActionDate: parsedDate || "",
        replyStatus: rowReplyStatus ? String(rowReplyStatus).trim() : "No reply",
        notes: rowNotes ? String(rowNotes).trim() : "",
        contactPerson: rowContactPerson ? String(rowContactPerson).trim() : "",
        email: rowEmail ? String(rowEmail).trim() : "",
        whatsappNumber: rowWhatsapp ? String(rowWhatsapp).trim() : "",
        extraLink: "",
        messageSent: rowMessageSent ? String(rowMessageSent).trim() : "",
        followUpCount: 0
      };
      
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
      
      // Channel-specific default overrides
      if (lead.channel === "Email") {
        const hasDate = (colMapping.date !== undefined && String(row[colMapping.date]).trim() !== "");
        const hasMsg = (colMapping.messageSent !== undefined && String(row[colMapping.messageSent]).trim() !== "");
        
        if (hasDate || hasMsg) {
          lead.stage = "First Message Sent";
        } else {
          lead.stage = rowStage ? String(rowStage).trim() : "Found";
        }
        lead.replyStatus = "No reply";
        lead.nextAction = "Send follow-up";
        
        // Date mapping logic
        if (colMapping.date !== undefined && row[colMapping.date]) {
          const emailDate = parseExcelDate(row[colMapping.date]);
          if (emailDate) {
            lead.lastActionDate = emailDate;
            lead.dateAdded = emailDate;
            lead.firstMessageDate = emailDate;
            if (!lead.nextActionDate) {
              lead.nextActionDate = addWorkingDays(emailDate, 5);
            }
          }
        }
      } else if (lead.channel === "WhatsApp") {
        lead.stage = rowStage ? String(rowStage).trim() : "Ready to WhatsApp";
        lead.replyStatus = "No reply";
        lead.nextAction = "Send WhatsApp";
        if (rowMessageSent) {
          lead.messageSent = String(rowMessageSent).trim();
        }
      } else if (lead.channel === "Instagram") {
        lead.stage = rowStage ? String(rowStage).trim() : "Found";
        lead.nextAction = "Like/comment";
      } else if (lead.channel === "LinkedIn") {
        lead.stage = rowStage ? String(rowStage).trim() : "Found";
        lead.nextAction = "Send connection request";
      }
      
      // Duplicate checks
      const emailLower = rowEmail ? String(rowEmail).trim().toLowerCase() : "";
      const phoneLower = rowWhatsapp ? String(rowWhatsapp).trim().toLowerCase() : "";
      const linkLower = rowMainLink ? String(rowMainLink).trim().toLowerCase() : "";
      const nameLower = rowName ? String(rowName).trim().toLowerCase() : "";
      
      let matchedIndex = -1;
      
      // Search in existing leads
      for (let i = 0; i < leads.length; i++) {
        const l = leads[i];
        const dbEmail = l.email ? String(l.email).trim().toLowerCase() : "";
        const dbPhone = l.whatsappNumber ? String(l.whatsappNumber).trim().toLowerCase() : "";
        const dbLink = l.mainLink ? String(l.mainLink).trim().toLowerCase() : "";
        const dbName = l.name ? String(l.name).trim().toLowerCase() : "";
        
        if (emailLower && dbEmail === emailLower) {
          matchedIndex = i;
          break;
        } else if (phoneLower && dbPhone === phoneLower) {
          matchedIndex = i;
          break;
        } else if (linkLower && dbLink === linkLower) {
          matchedIndex = i;
          break;
        } else if (!emailLower && !phoneLower && !linkLower) {
          if (nameLower && dbName === nameLower) {
            matchedIndex = i;
            break;
          }
        }
      }
      
      // Search in already accumulated valid leads from this file
      let isDupInSession = false;
      if (matchedIndex === -1) {
        for (let i = 0; i < validLeads.length; i++) {
          const l = validLeads[i];
          const sessEmail = l.email ? String(l.email).trim().toLowerCase() : "";
          const sessPhone = l.whatsappNumber ? String(l.whatsappNumber).trim().toLowerCase() : "";
          const sessLink = l.mainLink ? String(l.mainLink).trim().toLowerCase() : "";
          const sessName = l.name ? String(l.name).trim().toLowerCase() : "";
          
          if (emailLower && sessEmail === emailLower) {
            isDupInSession = true;
            break;
          } else if (phoneLower && sessPhone === phoneLower) {
            isDupInSession = true;
            break;
          } else if (linkLower && sessLink === linkLower) {
            isDupInSession = true;
            break;
          } else if (!emailLower && !phoneLower && !linkLower) {
            if (nameLower && sessName === nameLower) {
              isDupInSession = true;
              break;
            }
          }
        }
      }
      
      if (matchedIndex !== -1 || isDupInSession) {
        duplicateLeads.push({ lead, matchedIndex });
      } else {
        validLeads.push(lead);
      }
    });
    
    showImportPreview(totalRows, validLeads, duplicateLeads, missingInfoCount);
    
  } catch (err) {
    console.error("Error parsing spreadsheet file:", err);
    showToast(`Error reading spreadsheet: ${err.message}`, "error");
  }
}

function showImportPreview(totalRows, validLeads, duplicateLeads, missingInfo) {
  pendingValidLeads = validLeads;
  pendingDuplicateLeads = duplicateLeads;
  
  document.getElementById("prevTotalRows").textContent = totalRows;
  document.getElementById("prevValidLeads").textContent = validLeads.length;
  document.getElementById("prevDuplicatesFound").textContent = duplicateLeads.length;
  document.getElementById("prevMissingInfo").textContent = missingInfo;
  
  // Reset checkboxes to default unchecked
  document.getElementById("chkImportDupAnyway").checked = false;
  document.getElementById("chkUpdateExistingDup").checked = false;
  
  // Calculate and update counts based on defaults
  updateImportPreviewCounts();
  
  const modal = document.getElementById("importPreviewModal");
  if (modal) modal.classList.add("active");
}

function updateImportPreviewCounts() {
  const importDupAnyway = document.getElementById("chkImportDupAnyway").checked;
  const updateExistingDup = document.getElementById("chkUpdateExistingDup").checked;
  
  const totalDup = pendingDuplicateLeads.length;
  
  let dupSkipped = 0;
  let dupToImport = 0;
  let finalImportCount = pendingValidLeads.length;
  
  if (updateExistingDup) {
    // prioritize "Update existing duplicates", do not create duplicate records
    dupSkipped = 0;
    dupToImport = 0;
  } else if (importDupAnyway) {
    dupSkipped = 0;
    dupToImport = totalDup;
    finalImportCount = pendingValidLeads.length + totalDup;
  } else {
    dupSkipped = totalDup;
    dupToImport = 0;
  }
  
  document.getElementById("prevDuplicatesSkipped").textContent = dupSkipped;
  document.getElementById("prevDuplicatesToImport").textContent = dupToImport;
  document.getElementById("prevFinalCount").textContent = finalImportCount;
}

function confirmImport() {
  const importDupAnyway = document.getElementById("chkImportDupAnyway").checked;
  const updateExistingDup = document.getElementById("chkUpdateExistingDup").checked;
  
  const leadsToPush = [...pendingValidLeads];
  let updateCount = 0;
  
  if (updateExistingDup) {
    // Update existing duplicates in place
    pendingDuplicateLeads.forEach(dupItem => {
      const idx = dupItem.matchedIndex;
      if (idx !== -1 && idx < leads.length) {
        const existing = leads[idx];
        const imported = dupItem.lead;
        
        // Update fields of the existing lead with the imported row's values
        if (imported.name) existing.name = imported.name;
        if (imported.market) existing.market = imported.market;
        if (imported.mainLink) existing.mainLink = imported.mainLink;
        if (imported.niche) existing.niche = imported.niche;
        if (imported.source) existing.source = imported.source;
        if (imported.priority) existing.priority = imported.priority;
        if (imported.stage) existing.stage = imported.stage;
        if (imported.lastActionDate) existing.lastActionDate = imported.lastActionDate;
        if (imported.nextAction) existing.nextAction = imported.nextAction;
        if (imported.nextActionDate) existing.nextActionDate = imported.nextActionDate;
        if (imported.replyStatus) existing.replyStatus = imported.replyStatus;
        if (imported.contactPerson) existing.contactPerson = imported.contactPerson;
        if (imported.email) existing.email = imported.email;
        if (imported.whatsappNumber) existing.whatsappNumber = imported.whatsappNumber;
        if (imported.messageSent) existing.messageSent = imported.messageSent;
        
        // Append notes: preserve existing, append imported notes
        const importedNotes = imported.notes ? String(imported.notes).trim() : "";
        if (importedNotes) {
          existing.notes = (existing.notes ? existing.notes + "\n" : "") + "Updated from import: " + importedNotes;
        }
        
        updateCount++;
      } else {
        // duplicate in file session itself but doesn't exist in DB -> import as new
        leadsToPush.push(dupItem.lead);
      }
    });
  } else if (importDupAnyway) {
    // Import duplicates as separate records
    pendingDuplicateLeads.forEach(dupItem => {
      const dupLead = { ...dupItem.lead };
      // Add note
      dupLead.notes = (dupLead.notes ? dupLead.notes + "\n" : "") + "Imported as duplicate by user choice.";
      leadsToPush.push(dupLead);
    });
  }
  
  if (leadsToPush.length > 0 || updateCount > 0) {
    if (leadsToPush.length > 0) {
      leads.push(...leadsToPush);
    }
    saveData();
    updateDashboard();
    renderLeads();
    renderTodayActions();
    
    let msg = `Successfully imported ${leadsToPush.length} new leads`;
    if (updateCount > 0) {
      msg += ` and updated ${updateCount} existing leads`;
    }
    showToast(msg + "!", "success");
  } else {
    showToast("No leads imported or updated.", "info");
  }
  
  closeImportPreview();
}

function closeImportPreview() {
  const modal = document.getElementById("importPreviewModal");
  if (modal) modal.classList.remove("active");
  pendingValidLeads = [];
  pendingDuplicateLeads = [];
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
    scripts: scripts,
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
      if (!data || (!Array.isArray(data.leads) && !Array.isArray(data.scripts))) {
        showToast("Invalid backup file format. Leads or scripts array not found.", "error");
        return;
      }
      
      pendingBackupData = {
        leads: Array.isArray(data.leads) ? data.leads : [],
        scripts: Array.isArray(data.scripts) ? data.scripts : []
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
  const scriptsInBackup = pendingBackupData.scripts;
  
  let duplicateCount = 0;
  
  leadsInBackup.forEach(lead => {
    const dupCheck = checkBackupLeadDuplicate(lead);
    if (dupCheck) {
      duplicateCount++;
    }
  });
  
  document.getElementById("backLeadsFound").textContent = leadsInBackup.length;
  document.getElementById("backScriptsFound").textContent = scriptsInBackup.length;
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
    const warningText = "This will replace all current leads/scripts in this browser. Export a backup first. Continue?";
    if (!confirm(warningText)) {
      return;
    }
    
    // Replace mode
    leads = pendingBackupData.leads;
    if (pendingBackupData.scripts.length > 0) {
      scripts = pendingBackupData.scripts;
    }
    
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
    
    // Merge scripts
    let scriptsAdded = 0;
    pendingBackupData.scripts.forEach(newScript => {
      const existingIdx = scripts.findIndex(s => s.id === newScript.id || s.title.toLowerCase() === newScript.title.toLowerCase());
      if (existingIdx !== -1) {
        scripts[existingIdx] = newScript;
      } else {
        scripts.push(newScript);
        scriptsAdded++;
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
  saveScripts();
  closeBackupImportModal();
  
  // Refresh views
  updateDashboard();
  renderLeads();
  renderTodayActions();
  renderScripts();
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

// Save Scripts to LocalStorage + fire-and-forget cloud sync
function saveScripts() {
  localStorage.setItem("ali_raza_scripts", JSON.stringify(scripts));
  syncScriptsToCloud(); // async, fire-and-forget
}

// Render outreach message scripts dynamically
function renderScripts() {
  const container = document.getElementById("scriptsAccordion");
  if (!container) return;

  const searchQuery = document.getElementById("scriptSearchInput")?.value.toLowerCase().trim() || "";
  const channelFilter = document.getElementById("filterScriptChannel")?.value || "All";
  const typeFilter = document.getElementById("filterScriptType")?.value || "All";

  // Filter scripts
  const filtered = scripts.filter(script => {
    // Search matching title or body
    const matchesSearch = !searchQuery || 
      script.title.toLowerCase().includes(searchQuery) || 
      script.body.toLowerCase().includes(searchQuery);
    
    // Channel filter
    const matchesChannel = channelFilter === "All" || script.channel === channelFilter;
    
    // Type filter
    const matchesType = typeFilter === "All" || script.type === typeFilter;

    return matchesSearch && matchesChannel && matchesType;
  });

  // Clear container
  container.innerHTML = "";

  if (filtered.length === 0) {
    container.innerHTML = `
      <div style="text-align: center; padding: 24px; color: var(--color-priority-c); font-size: 14px; background-color: var(--color-off-white); border-radius: var(--radius-md); border: var(--border-light);">
        No outreach scripts found matching your filters.
      </div>
    `;
    return;
  }

  // Render each script
  filtered.forEach(script => {
    const isEditing = editingScriptId === script.id;
    
    const accordionItem = document.createElement("div");
    accordionItem.className = `accordion-item${isEditing ? " active" : ""}`;
    accordionItem.dataset.id = script.id;

    // Header Trigger
    const triggerBtn = document.createElement("button");
    triggerBtn.className = "accordion-trigger";
    triggerBtn.style.display = "flex";
    triggerBtn.style.alignItems = "center";
    triggerBtn.style.justifyContent = "space-between";
    triggerBtn.style.width = "100%";
    triggerBtn.style.gap = "8px";
    
    if (isEditing) {
      triggerBtn.style.cursor = "default";
    }

    triggerBtn.innerHTML = `
      <span style="display: flex; align-items: center; gap: 8px; flex-wrap: wrap; pointer-events: none;">
        <span class="script-title-display" style="font-weight: 600;">${escapeHtml(script.title)}</span>
        <span class="script-badge channel">${escapeHtml(script.channel)}</span>
        <span class="script-badge type">${escapeHtml(script.type)}</span>
      </span>
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width: 14px; height: 14px; flex-shrink: 0; pointer-events: none;"><polyline points="6 9 12 15 18 9"/></svg>
    `;

    const contentDiv = document.createElement("div");
    contentDiv.className = "accordion-content";
    if (isEditing) {
      contentDiv.style.maxHeight = "none";
    }

    if (isEditing) {
      // Edit form layout
      contentDiv.innerHTML = `
        <div class="script-body editing-form" style="display: flex; flex-direction: column; gap: 12px; padding: 16px 18px;">
          <div class="form-group" style="margin: 0; display: flex; flex-direction: column; gap: 4px;">
            <label style="font-size: 11px; font-weight: 700; text-transform: uppercase; color: var(--color-priority-c); display: block;">Script Title</label>
            <input type="text" class="edit-script-title" value="${escapeHtml(script.title)}" style="width: 100%; padding: 8px 12px; font-size: 13px; border-radius: var(--radius-sm); border: var(--border-light); outline: none;">
          </div>
          
          <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px;">
            <div class="form-group" style="margin: 0; display: flex; flex-direction: column; gap: 4px;">
              <label style="font-size: 11px; font-weight: 700; text-transform: uppercase; color: var(--color-priority-c); display: block;">Channel</label>
              <select class="edit-script-channel" style="width: 100%; padding: 8px 12px; font-size: 13px; border-radius: var(--radius-sm); border: var(--border-light); background-color: var(--color-white); outline: none;">
                <option value="Instagram" ${script.channel === 'Instagram' ? 'selected' : ''}>Instagram</option>
                <option value="LinkedIn" ${script.channel === 'LinkedIn' ? 'selected' : ''}>LinkedIn</option>
                <option value="Email" ${script.channel === 'Email' ? 'selected' : ''}>Email</option>
                <option value="WhatsApp" ${script.channel === 'WhatsApp' ? 'selected' : ''}>WhatsApp</option>
                <option value="General" ${script.channel === 'General' ? 'selected' : ''}>General</option>
              </select>
            </div>
            <div class="form-group" style="margin: 0; display: flex; flex-direction: column; gap: 4px;">
              <label style="font-size: 11px; font-weight: 700; text-transform: uppercase; color: var(--color-priority-c); display: block;">Script Type</label>
              <select class="edit-script-type" style="width: 100%; padding: 8px 12px; font-size: 13px; border-radius: var(--radius-sm); border: var(--border-light); background-color: var(--color-white); outline: none;">
                <option value="First Message" ${script.type === 'First Message' ? 'selected' : ''}>First Message</option>
                <option value="Follow-up" ${script.type === 'Follow-up' ? 'selected' : ''}>Follow-up</option>
                <option value="Connection Request" ${script.type === 'Connection Request' ? 'selected' : ''}>Connection Request</option>
                <option value="After Accepting" ${script.type === 'After Accepting' ? 'selected' : ''}>After Accepting</option>
                <option value="Sample Reply" ${script.type === 'Sample Reply' ? 'selected' : ''}>Sample Reply</option>
                <option value="CV Reply" ${script.type === 'CV Reply' ? 'selected' : ''}>CV Reply</option>
                <option value="Price Reply" ${script.type === 'Price Reply' ? 'selected' : ''}>Price Reply</option>
                <option value="Not Now Reply" ${script.type === 'Not Now Reply' ? 'selected' : ''}>Not Now Reply</option>
                <option value="Custom" ${script.type === 'Custom' ? 'selected' : ''}>Custom</option>
              </select>
            </div>
          </div>
          
          <div class="form-group" style="margin: 0; display: flex; flex-direction: column; gap: 4px;">
            <label style="font-size: 11px; font-weight: 700; text-transform: uppercase; color: var(--color-priority-c); display: block;">Message Body</label>
            <textarea class="edit-script-body" rows="6" style="width: 100%; padding: 8px 12px; font-size: 13px; border-radius: var(--radius-sm); border: var(--border-light); font-family: var(--font-body); line-height: 1.5; resize: vertical; min-height: 100px; outline: none;">${escapeHtml(script.body)}</textarea>
          </div>
          
          <div style="display: flex; justify-content: flex-end; gap: 8px; margin-top: 4px; flex-wrap: wrap;">
            <button class="btn btn-secondary btn-sm cancel-edit-btn" data-id="${script.id}">
              Cancel
            </button>
            <button class="btn btn-primary btn-sm save-edit-btn" data-id="${script.id}">
              Save Changes
            </button>
          </div>
        </div>
      `;
    } else {
      // Read mode layout
      contentDiv.innerHTML = `
        <div class="script-body">
          <div class="script-text-container" style="position: relative;">
            <pre style="font-family: var(--font-body); white-space: pre-wrap; word-break: break-word; line-height: 1.6; margin: 0;">${escapeHtml(script.body)}</pre>
          </div>
          <div class="script-actions" style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px;">
            <div>
              <span style="font-size: 11px; color: var(--color-priority-c); font-weight: 600;">
                ${script.isDefault ? "Default Script" : "Custom Script"}
              </span>
            </div>
            <div style="display: flex; gap: 8px; flex-wrap: wrap;">
              <button class="btn btn-secondary btn-sm copy-btn" data-id="${script.id}">
                Copy Script
              </button>
              <button class="btn btn-secondary btn-sm edit-btn" data-id="${script.id}">
                Edit
              </button>
              <button class="btn btn-danger-outline btn-sm delete-btn" data-id="${script.id}">
                Delete
              </button>
            </div>
          </div>
        </div>
      `;
    }

    accordionItem.appendChild(triggerBtn);
    accordionItem.appendChild(contentDiv);
    container.appendChild(accordionItem);
  });
}

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

// Helper to get script with replaced placeholders
function getPersonalizedScript(lead, script) {
  if (!script) return "";
  
  const nameVal = lead.contactPerson ? lead.contactPerson : (lead.name || "there");
  const companyVal = lead.name || "";
  const nicheVal = lead.niche || "";
  const marketVal = lead.market || "";
  
  let body = script.body;
  body = body.replace(/\[Name\]/gi, nameVal);
  body = body.replace(/\[Nome\]/gi, nameVal);
  body = body.replace(/\[Company\]/gi, companyVal);
  body = body.replace(/\[Nome Azienda\]/gi, companyVal);
  body = body.replace(/\[Azienda\]/gi, companyVal);
  body = body.replace(/\[Niche\]/gi, nicheVal);
  body = body.replace(/\[Market\]/gi, marketVal);
  
  return body;
}

// Smart Message Copy
function copyPersonalizedScript(originalIndex, scriptType) {
  const lead = leads[originalIndex];
  if (!lead) return;

  let scriptTitle = "";
  let typeKey = "";
  let channelKey = "";

  if (scriptType === "Email") {
    scriptTitle = lead.market === "Italy" ? "Italian First Email (Primo Contatto)" : "English First Email (Agencies/Coaches)";
    typeKey = "First Message";
    channelKey = "Email";
  } else if (scriptType === "WhatsApp") {
    scriptTitle = "WhatsApp First Message (Public Contacts Only)";
    typeKey = "First Message";
    channelKey = "WhatsApp";
  } else if (scriptType === "Instagram" || scriptType === "DM") {
    scriptTitle = "Instagram DM (Found Pitch)";
    typeKey = "First Message";
    channelKey = "Instagram";
  } else if (scriptType === "LinkedIn") {
    if (lead.stage === "Found" || lead.stage === "Engaged") {
      scriptTitle = "LinkedIn Connection Request";
      typeKey = "Connection Request";
    } else {
      scriptTitle = "LinkedIn Message (After Accepting)";
      typeKey = "After Accepting";
    }
    channelKey = "LinkedIn";
  } else if (scriptType === "Follow-up") {
    if (lead.channel === "Email") scriptTitle = "Email Follow-up (Polite Nudge)";
    else if (lead.channel === "WhatsApp") scriptTitle = "WhatsApp Follow-up";
    else {
      typeKey = "Follow-up";
      channelKey = lead.channel;
    }
  } else if (scriptType === "Sample Reply") {
    scriptTitle = lead.market === "Italy" ? "Italian Sample Reply" : "English Sample Reply";
    typeKey = "Sample Reply";
  } else if (scriptType === "CV Reply") {
    scriptTitle = "Italian CV Sent Reply";
    typeKey = "CV Reply";
  } else if (scriptType === "Price Reply") {
    scriptTitle = "Price Reply";
    typeKey = "Price Reply";
  }

  let script = null;
  if (scriptTitle) {
    script = scripts.find(s => s.title.toLowerCase() === scriptTitle.toLowerCase());
  }
  if (!script && typeKey) {
    script = scripts.find(s => s.type === typeKey && (channelKey ? s.channel === channelKey : true));
  }
  if (!script && typeKey) {
    script = scripts.find(s => s.type === typeKey);
  }
  if (!script) {
    script = scripts[0];
  }

  if (!script) {
    showToast("No template script found. Please add a script first.", "error");
    return;
  }

  const text = getPersonalizedScript(lead, script);
  copyTextToClipboard(text).then(() => {
    showToast("Message copied", "success");
  }).catch(() => {
    showToast("Failed to copy script.", "error");
  });
}

// Quick Actions Dropdown Builder
function getQuickActionsDropdownHtml(lead) {
  const index = lead.originalIndex;
  let itemsHtml = "";

  if (lead.channel === "Email") {
    itemsHtml = `
      <button class="dropdown-item" onclick="copyPersonalizedScript(${index}, 'Email')"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/></svg> Copy Email Script</button>
      <button class="dropdown-item" onclick="copyPersonalizedScript(${index}, 'Follow-up')"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/></svg> Copy Follow-up</button>
      <button class="dropdown-item" onclick="copyPersonalizedScript(${index}, 'Sample Reply')"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/></svg> Copy Sample Reply</button>
      <button class="dropdown-item" onclick="copyPersonalizedScript(${index}, 'CV Reply')"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/></svg> Copy CV Reply</button>
      <button class="dropdown-item" onclick="copyPersonalizedScript(${index}, 'Price Reply')"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/></svg> Copy Price Reply</button>
      <div style="border-top: 1px dashed rgba(11,31,58,0.08); margin: 4px 0;"></div>
      <button class="dropdown-item" onclick="markSentEmail(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg> Mark Email Sent</button>
      <button class="dropdown-item" onclick="markFollowupSent(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8"/><path d="M16 3h5v5"/><path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16"/><path d="M8 21H3v-5"/></svg> Mark Follow-up Sent</button>
      <button class="dropdown-item" onclick="sendSamples(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/><polyline points="3.27 6.96 12 12.01 20.73 6.96"/><line x1="12" x2="12" y1="22.08" y2="12"/></svg> Send Samples</button>
      <button class="dropdown-item" onclick="sendCV(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="M10 9H8"/><path d="M16 13H8"/><path d="M16 17H8"/></svg> Send CV</button>
      <div style="border-top: 1px dashed rgba(11,31,58,0.08); margin: 4px 0;"></div>
      <button class="dropdown-item" onclick="openLeadLink(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 3h6v6"/><path d="M10 14 21 3"/><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/></svg> Open Website</button>
      <button class="dropdown-item" onclick="archiveLead(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="20" height="5" x="2" y="3" rx="1"/><path d="M4 8v11a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8"/><line x1="10" x2="14" y1="12" y2="12"/></svg> Archive</button>
    `;
  } else if (lead.channel === "WhatsApp") {
    itemsHtml = `
      <button class="dropdown-item" onclick="copyPersonalizedScript(${index}, 'WhatsApp')"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/></svg> Copy WhatsApp Script</button>
      <button class="dropdown-item" onclick="copyPersonalizedScript(${index}, 'Follow-up')"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/></svg> Copy Follow-up</button>
      <button class="dropdown-item" onclick="copyPersonalizedScript(${index}, 'Price Reply')"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/></svg> Copy Price Reply</button>
      <div style="border-top: 1px dashed rgba(11,31,58,0.08); margin: 4px 0;"></div>
      <button class="dropdown-item" onclick="openWhatsAppChat(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M7.9 20A9 9 0 1 0 4 16.1L2 22Z"/></svg> Open WhatsApp</button>
      <button class="dropdown-item" onclick="markSentWhatsApp(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg> Mark WhatsApp Sent</button>
      <button class="dropdown-item" onclick="markFollowupSent(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8"/><path d="M16 3h5v5"/><path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16"/><path d="M8 21H3v-5"/></svg> Mark Follow-up Sent</button>
      <button class="dropdown-item" onclick="sendSamples(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/><polyline points="3.27 6.96 12 12.01 20.73 6.96"/><line x1="12" x2="12" y1="22.08" y2="12"/></svg> Send Samples</button>
      <div style="border-top: 1px dashed rgba(11,31,58,0.08); margin: 4px 0;"></div>
      <button class="dropdown-item" onclick="archiveLead(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="20" height="5" x="2" y="3" rx="1"/><path d="M4 8v11a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8"/><line x1="10" x2="14" y1="12" y2="12"/></svg> Archive</button>
    `;
  } else if (lead.channel === "Instagram") {
    itemsHtml = `
      <button class="dropdown-item" onclick="copyPersonalizedScript(${index}, 'Instagram')"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/></svg> Copy DM Script</button>
      <button class="dropdown-item" onclick="copyPersonalizedScript(${index}, 'Follow-up')"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/></svg> Copy Follow-up</button>
      <div style="border-top: 1px dashed rgba(11,31,58,0.08); margin: 4px 0;"></div>
      <button class="dropdown-item" onclick="openLeadLink(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 3h6v6"/><path d="M10 14 21 3"/><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/></svg> Open Profile</button>
      <button class="dropdown-item" onclick="markInstagramCommented(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg> Mark Commented</button>
      <button class="dropdown-item" onclick="markInstagramFollowed(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><polyline points="16 11 18 13 22 9"/></svg> Mark Followed</button>
      <button class="dropdown-item" onclick="markSentDM(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg> Mark DM Sent</button>
      <button class="dropdown-item" onclick="markFollowupSent(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8"/><path d="M16 3h5v5"/><path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16"/><path d="M8 21H3v-5"/></svg> Mark Follow-up Sent</button>
      <div style="border-top: 1px dashed rgba(11,31,58,0.08); margin: 4px 0;"></div>
      <button class="dropdown-item" onclick="archiveLead(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="20" height="5" x="2" y="3" rx="1"/><path d="M4 8v11a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8"/><line x1="10" x2="14" y1="12" y2="12"/></svg> Archive</button>
    `;
  } else if (lead.channel === "LinkedIn") {
    itemsHtml = `
      <button class="dropdown-item" onclick="copyPersonalizedScript(${index}, 'LinkedIn')"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/></svg> Copy LinkedIn Script</button>
      <button class="dropdown-item" onclick="copyPersonalizedScript(${index}, 'Follow-up')"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/></svg> Copy Follow-up</button>
      <button class="dropdown-item" onclick="copyPersonalizedScript(${index}, 'Price Reply')"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/></svg> Copy Price Reply</button>
      <div style="border-top: 1px dashed rgba(11,31,58,0.08); margin: 4px 0;"></div>
      <button class="dropdown-item" onclick="openLeadLink(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 3h6v6"/><path d="M10 14 21 3"/><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/></svg> Open Profile</button>
      <button class="dropdown-item" onclick="markLinkedInConnectionSent(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><line x1="19" x2="19" y1="8" y2="14"/><line x1="22" x2="16" y1="11" y2="11"/></svg> Mark Connection Sent</button>
      <button class="dropdown-item" onclick="markSentLinkedInMessage(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg> Mark Message Sent</button>
      <button class="dropdown-item" onclick="markFollowupSent(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8"/><path d="M16 3h5v5"/><path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16"/><path d="M8 21H3v-5"/></svg> Mark Follow-up Sent</button>
      <div style="border-top: 1px dashed rgba(11,31,58,0.08); margin: 4px 0;"></div>
      <button class="dropdown-item" onclick="archiveLead(${index})"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="20" height="5" x="2" y="3" rx="1"/><path d="M4 8v11a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8"/><line x1="10" x2="14" y1="12" y2="12"/></svg> Archive</button>
    `;
  }

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
  
  const isFirstEmail = (lead.stage === "Found" || lead.stage === "Engaged");
  if (isFirstEmail) {
    lead.stage = "First Message Sent";
    lead.followUpCount = 0;
    lead.nextAction = "Send follow-up";
    lead.nextActionDate = addWorkingDays(today, 5);
  } else {
    lead.stage = "Follow-up Due";
    lead.followUpCount = (lead.followUpCount || 0) + 1;
    
    const limit = 2; // email limit
    if (lead.followUpCount >= limit) {
      lead.nextAction = "Archive";
      lead.stage = "Follow-up Sent";
      alert(`Follow-up limit of ${limit} reached for ${lead.name}. Consider archiving this lead.`);
    } else {
      lead.nextAction = "Send follow-up";
      lead.nextActionDate = addWorkingDays(today, 5);
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
  
  let limit = 2;
  if (lead.channel === "WhatsApp") limit = 1;
  else if (lead.channel === "Instagram") limit = 1;
  else if (lead.channel === "LinkedIn") limit = 2;
  
  if (lead.followUpCount >= limit) {
    lead.nextAction = "Archive";
    lead.stage = "Follow-up Sent";
    alert(`Follow-up limit of ${limit} reached for ${lead.name}. Consider archiving this lead.`);
  } else {
    lead.nextAction = "Send follow-up";
    if (lead.channel === "Email") {
      lead.nextActionDate = addWorkingDays(today, 5);
    } else if (lead.channel === "LinkedIn") {
      lead.nextActionDate = addDays(today, 5);
    } else {
      lead.nextActionDate = addDays(today, 7);
    }
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

window.bulkMarkFollowupSent = function() {
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
    
    let limit = 2;
    if (lead.channel === "WhatsApp") limit = 1;
    else if (lead.channel === "Instagram") limit = 1;
    else if (lead.channel === "LinkedIn") limit = 2;
    
    if (lead.followUpCount >= limit) {
      lead.nextAction = "Archive";
      lead.stage = "Follow-up Sent";
      limitWarningTriggered = true;
    } else {
      lead.nextAction = "Send follow-up";
      if (lead.channel === "Email") {
        lead.nextActionDate = addWorkingDays(today, 5);
      } else if (lead.channel === "LinkedIn") {
        lead.nextActionDate = addDays(today, 5);
      } else {
        lead.nextActionDate = addDays(today, 7);
      }
    }
  });
  
  saveData();
  updateDashboard();
  renderLeads();
  renderTodayActions();
  
  if (limitWarningTriggered) {
    alert("One or more leads reached their follow-up limit. Consider archiving them.");
  }
  
  showToast(`Marked follow-up sent for ${indexes.length} leads!`, "success");
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
window.copyPersonalizedScript = copyPersonalizedScript;
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
window.bulkMarkFollowupSent = bulkMarkFollowupSent;
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


/* D:\Lead Tracker\app.js */

// Helper to format dates relative to today
function getOffsetDateString(daysOffset) {
  const d = new Date();
  d.setDate(d.getDate() + daysOffset);
  return d.toISOString().split('T')[0];
}

// Parse dates from excel (handles serial number and text/strings)
function parseExcelDate(val) {
  if (val === undefined || val === null || val === "") return null;
  
  if (val instanceof Date) {
    return val.toISOString().split('T')[0];
  }
  
  if (typeof val === 'number' || (!isNaN(val) && !isNaN(parseFloat(val)))) {
    const serial = parseFloat(val);
    const utc_days = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);
    const localDate = new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate());
    return localDate.toISOString().split('T')[0];
  }

  const str = String(val).trim();
  if (!str) return null;

  const parsedMs = Date.parse(str);
  if (!isNaN(parsedMs)) {
    return new Date(parsedMs).toISOString().split('T')[0];
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
        return d.toISOString().split('T')[0];
      }
    }
  }

  // Handle format with text like "5/13/2026, got replied..."
  const dateRegex = /(\d{1,2})[\/\-]\d{1,2}[\/\-]\d{4}/;
  const dateMatch = str.match(dateRegex);
  if (dateMatch) {
    const d = new Date(dateMatch[0]);
    if (!isNaN(d.getTime())) {
      return d.toISOString().split('T')[0];
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
  return date.toISOString().split('T')[0];
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
    notes: "Strong fit. Sells €2k mindset courses. No active lead magnet on website. Sent first pitch on LinkedIn about workbook companion design.",
    contactPerson: "Andrea Gruber",
    email: "",
    whatsappNumber: "",
    extraLink: "https://mindset-academy.at",
    messageSent: "Hi Andrea, loved your post on cognitive reframing. Sells €2k mindset courses...",
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

// On Page Load
document.addEventListener("DOMContentLoaded", () => {
  loadData();
  setupEventListeners();
  updateDashboard();
  renderLeads();
  renderTodayActions();
});

// Load Leads from LocalStorage or Seed Defaults
function loadData() {
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
}

// Save Leads to LocalStorage
function saveData() {
  localStorage.setItem("ali_raza_leads", JSON.stringify(leads));
}

// Calculate and Render Dashboard Metrics
function updateDashboard() {
  const todayStr = new Date().toISOString().split('T')[0];

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
  const todayStr = new Date().toISOString().split('T')[0];
  const searchQuery = document.getElementById("searchInput").value.trim().toLowerCase();
  
  // Advanced filter values
  const filterChannel = document.getElementById("filterChannel").value;
  const filterMarket = document.getElementById("filterMarket").value;
  const filterSource = document.getElementById("filterSource").value;
  const filterPriority = document.getElementById("filterPriority").value;
  const filterStage = document.getElementById("filterStage").value;
  const filterReply = document.getElementById("filterReply").value;
  const filterActionDate = document.getElementById("filterActionDate").value;

  return leads.filter((lead, index) => {
    // Keep track of index for operations
    lead.originalIndex = index;

    // 1. Channel Tab filtering (sync with tab nav)
    if (activeTab !== "All" && lead.channel !== activeTab) return false;

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

// Render Leads Grid (Table and Mobile Cards)
function renderLeads() {
  // Toggle the tab action button container and text dynamically
  const tabActions = document.getElementById("tabActionsContainer");
  const btnText = document.getElementById("importLeadsBtnText");
  if (tabActions && btnText) {
    if (activeTab === "All") {
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
  countBadge.textContent = `${filtered.length} leads`;
  document.getElementById("activeTabTitle").textContent = activeTab === "All" ? "Dashboard" : `${activeTab} Leads`;

  // Toggle Today's Action Center section visibility (only on Dashboard tab)
  const todaySection = document.querySelector(".today-actions-section");
  if (todaySection) {
    todaySection.style.display = activeTab === "All" ? "block" : "none";
  }

  // Update table header text dynamically
  const contactHeader = document.getElementById("dynamicContactHeader");
  if (contactHeader) {
    if (activeTab === "Email") {
      contactHeader.textContent = "Email Address";
    } else if (activeTab === "WhatsApp") {
      contactHeader.textContent = "WhatsApp Number";
    } else if (activeTab === "Instagram" || activeTab === "LinkedIn") {
      contactHeader.textContent = "Profile Link";
    } else {
      contactHeader.textContent = "Channel";
    }
  }

  tableBody.innerHTML = "";
  cardsContainer.innerHTML = "";

  if (filtered.length === 0) {
    const emptyHtml = `
      <tr>
        <td colspan="9" style="text-align: center;">
          <div class="empty-state">
            <span class="empty-state-icon">🔍</span>
            <h3>No leads found</h3>
            <p>Try clearing your search query, adjusting your active filters, or adding a new lead.</p>
          </div>
        </td>
      </tr>
    `;
    tableBody.innerHTML = emptyHtml;

    cardsContainer.innerHTML = `
      <div class="empty-state">
        <span class="empty-state-icon">🔍</span>
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
      contactCellHtml = `<td>${email ? `<a href="mailto:${email}" style="font-weight: 600; color: var(--color-royal-blue);">${email}</a>` : '-'}</td>`;
    } else if (activeTab === "WhatsApp") {
      const phone = lead.whatsappNumber ? String(lead.whatsappNumber).trim() : "";
      const waLink = getWhatsAppLink(phone);
      contactCellHtml = `<td>
        <div style="font-weight: 600;">${phone || "-"}</div>
        ${phone ? `<a href="${waLink}" target="_blank" rel="noopener noreferrer" style="font-size: 11px; color: var(--color-teal-green); font-weight: 600; display: inline-flex; align-items: center; gap: 4px; margin-top: 2px;">
          Open WhatsApp ↗
        </a>` : ''}
      </td>`;
    } else if (activeTab === "Instagram" || activeTab === "LinkedIn") {
      const link = lead.mainLink ? String(lead.mainLink).trim() : "";
      const normalized = normalizeUrl(link);
      const displayText = link ? (link.length > 30 ? link.substring(0, 30) + "..." : link) : "";
      contactCellHtml = `<td>${link ? `<a href="${normalized}" target="_blank" rel="noopener noreferrer" style="font-weight: 600; color: var(--color-royal-blue);">${displayText} ↗</a>` : '-'}</td>`;
    } else {
      contactCellHtml = `<td>${getChannelBadge(lead.channel)}</td>`;
    }

    // 1. Table Row (Desktop)
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>
        <div style="font-weight: 700; color: var(--color-deep-navy);">${lead.name || "Unnamed Lead / Company"}</div>
        ${lead.mainLink ? `<a href="${normalizeUrl(lead.mainLink)}" target="_blank" rel="noopener noreferrer" style="font-size: 11px; display: inline-flex; align-items: center; gap: 4px; margin-top: 2px;">
          Open Link ↗
        </a>` : ''}
      </td>
      ${contactCellHtml}
      <td><strong>${lead.market}</strong></td>
      <td>${lead.niche}</td>
      <td><span style="font-size: 12.5px; font-weight: 600; color: var(--color-priority-c);">${lead.source || "Other"}</span></td>
      <td>${getPriorityBadge(lead.priority)}</td>
      <td>${getStageBadge(lead.stage)}</td>
      <td style="font-weight: 600;">${lead.nextAction}</td>
      <td style="white-space: nowrap;">
        <span style="font-weight: 600; color: ${lead.nextActionDate && lead.nextActionDate <= getOffsetDateString(0) ? '#ef4444' : 'inherit'}">
          ${lead.nextActionDate || '-'}
        </span>
      </td>
      <td>${getReplyBadge(lead.replyStatus)}</td>
      <td>
        <div class="action-buttons">
          <button class="action-btn edit-btn" onclick="openEditModal(${lead.originalIndex})" title="Edit Lead">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 20h9M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg>
          </button>
          <button class="action-btn archive-btn" onclick="archiveLead(${lead.originalIndex})" title="Archive Lead">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2" ry="2"/><line x1="9" y1="9" x2="15" y2="15"/><line x1="15" y1="9" x2="9" y2="15"/></svg>
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
      <div class="lead-card-header">
        <div class="lead-card-title">
          <h3>${lead.name || "Unnamed Lead / Company"}</h3>
          <span class="market-lbl">${lead.market || "-"} • ${lead.niche || "-"} • ${lead.source || "Other"}</span>
        </div>
        <div class="lead-card-badges">
          ${getPriorityBadge(lead.priority)}
          ${getChannelBadge(lead.channel)}
        </div>
      </div>
      
      <div class="lead-card-details-grid">
        <div class="lead-card-detail-item">
          <span class="lbl">Stage</span>
          <span class="val">${getStageBadge(lead.stage)}</span>
        </div>
        <div class="lead-card-detail-item">
          <span class="lbl">Reply Status</span>
          <span class="val">${getReplyBadge(lead.replyStatus)}</span>
        </div>
        <div class="lead-card-detail-item">
          <span class="lbl">Next Action</span>
          <span class="val" style="font-weight: 700;">${lead.nextAction}</span>
        </div>
        <div class="lead-card-detail-item">
          <span class="lbl">Action Date</span>
          <span class="val" style="font-weight: 700; color: ${lead.nextActionDate && lead.nextActionDate <= getOffsetDateString(0) ? '#ef4444' : 'inherit'}">${lead.nextActionDate || '-'}</span>
        </div>
      </div>

      <div class="lead-card-notes">
        <strong>Notes:</strong> ${lead.notes}
      </div>

      ${lead.contactPerson || lead.email || lead.whatsappNumber || lead.dateAdded || lead.lastActionDate ? `
      <div class="lead-card-notes" style="font-size: 11px; background-color: var(--color-off-white); padding: 6px; border-radius: var(--radius-sm);">
        ${lead.contactPerson ? `<strong>Contact:</strong> ${lead.contactPerson}<br>` : ''}
        ${lead.email ? `<strong>Email:</strong> ${lead.email}<br>` : ''}
        ${lead.whatsappNumber ? `<strong>WhatsApp:</strong> ${lead.whatsappNumber}<br>` : ''}
        ${lead.dateAdded ? `<strong>Added:</strong> ${lead.dateAdded} | ` : ''}
        ${lead.lastActionDate ? `<strong>Last Action:</strong> ${lead.lastActionDate}` : ''}
      </div>` : ''}

      <div class="lead-card-actions">
        ${lead.mainLink ? `<a href="${normalizeUrl(lead.mainLink)}" target="_blank" rel="noopener noreferrer" class="btn btn-secondary" style="padding: 6px 12px; font-size: 12px; margin-right: auto;">Open Link ↗</a>` : ''}
        <button class="btn btn-secondary btn-icon-only" onclick="openEditModal(${lead.originalIndex})" title="Edit Lead">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 20h9M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg>
        </button>
        <button class="btn btn-secondary btn-icon-only" onclick="archiveLead(${lead.originalIndex})" title="Archive Lead">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2" ry="2"/><line x1="9" y1="9" x2="15" y2="15"/><line x1="15" y1="9" x2="9" y2="15"/></svg>
        </button>
        <button class="btn btn-secondary btn-icon-only btn-danger-outline" onclick="deleteLead(${lead.originalIndex})" title="Delete Lead">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path><line x1="10" y1="11" x2="10" y2="17"></line><line x1="14" y1="11" x2="14" y2="17"></line></svg>
        </button>
      </div>
    `;
    cardsContainer.appendChild(card);
  });
}

// Render "Today's Actions" section
function renderTodayActions() {
  const todayStr = new Date().toISOString().split('T')[0];
  const tableBody = document.getElementById("todayTableBody");
  const cardsContainer = document.getElementById("todayCardsContainer");
  const countBadge = document.getElementById("todayActionsCount");

  // Filter criteria: Next Action Date is today or earlier AND Stage is not Archived
  const todayLeads = leads
    .map((lead, index) => ({ ...lead, originalIndex: index }))
    .filter(lead => lead.nextActionDate <= todayStr && lead.stage !== "Archived");

  countBadge.textContent = `${todayLeads.length} active`;

  tableBody.innerHTML = "";
  cardsContainer.innerHTML = "";

  if (todayLeads.length === 0) {
    const emptyHtml = `
      <tr>
        <td colspan="9" style="text-align: center;">
          <div class="empty-state">
            <span class="empty-state-icon">🎉</span>
            <h3>Inbox Zero! All actions completed.</h3>
            <p>You have no pending or overdue outreach actions for today.</p>
          </div>
        </td>
      </tr>
    `;
    tableBody.innerHTML = emptyHtml;
    document.getElementById("todayDesktopContainer").style.display = "block";

    cardsContainer.innerHTML = `
      <div class="empty-state">
        <span class="empty-state-icon">🎉</span>
        <h3>Inbox Zero! All tasks done.</h3>
      </div>
    `;
    return;
  }

  // Render Table & Cards
  todayLeads.forEach(lead => {
    // 1. Table
    const tr = document.createElement("tr");
    tr.style.backgroundColor = "rgba(214, 168, 79, 0.03)"; // slight gold highlight
    tr.innerHTML = `
      <td>
        <div style="font-weight: 700; color: var(--color-deep-navy);">${lead.name || "Unnamed Lead / Company"}</div>
        ${lead.mainLink ? `<a href="${normalizeUrl(lead.mainLink)}" target="_blank" rel="noopener noreferrer" style="font-size: 11px; display: inline-flex; align-items: center; gap: 4px; margin-top: 2px;">
          Open Link ↗
        </a>` : ''}
      </td>
      <td>${getChannelBadge(lead.channel)}</td>
      <td><strong>${lead.market}</strong></td>
      <td>${lead.niche}</td>
      <td><span style="font-size: 12.5px; font-weight: 600; color: var(--color-priority-c);">${lead.source || "Other"}</span></td>
      <td>${getPriorityBadge(lead.priority)}</td>
      <td>${getStageBadge(lead.stage)}</td>
      <td style="font-weight: 700; color: var(--color-royal-blue);">${lead.nextAction}</td>
      <td style="white-space: nowrap;">
        <span style="font-weight: 800; color: #ef4444;">
          ${lead.nextActionDate || '-'} (Overdue/Due)
        </span>
      </td>
      <td>${getReplyBadge(lead.replyStatus)}</td>
      <td>
        <div class="action-buttons">
          <button class="action-btn edit-btn" onclick="openEditModal(${lead.originalIndex})" title="Edit Lead">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 20h9M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg>
          </button>
          <button class="action-btn archive-btn" onclick="archiveLead(${lead.originalIndex})" title="Archive Lead">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2" ry="2"/><line x1="9" y1="9" x2="15" y2="15"/><line x1="15" y1="9" x2="9" y2="15"/></svg>
          </button>
          <button class="action-btn delete-btn" onclick="deleteLead(${lead.originalIndex})" title="Delete Lead">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path><line x1="10" y1="11" x2="10" y2="17"></line><line x1="14" y1="11" x2="14" y2="17"></line></svg>
          </button>
        </div>
      </td>
    `;
    tableBody.appendChild(tr);

    // 2. Mobile Cards
    const card = document.createElement("div");
    card.className = "lead-card-mobile";
    card.style.borderLeft = "4px solid var(--color-soft-gold)";
    card.innerHTML = `
      <div class="lead-card-header">
        <div class="lead-card-title">
          <h3>${lead.name || "Unnamed Lead / Company"}</h3>
          <span class="market-lbl">${lead.market || "-"} • ${lead.niche || "-"} • ${lead.source || "Other"}</span>
        </div>
        <div class="lead-card-badges">
          ${getPriorityBadge(lead.priority)}
          ${getChannelBadge(lead.channel)}
        </div>
      </div>
      
      <div class="lead-card-details-grid">
        <div class="lead-card-detail-item">
          <span class="lbl">Stage</span>
          <span class="val">${getStageBadge(lead.stage)}</span>
        </div>
        <div class="lead-card-detail-item">
          <span class="lbl">Next Action</span>
          <span class="val" style="font-weight: 700; color: var(--color-royal-blue);">${lead.nextAction}</span>
        </div>
        <div class="lead-card-detail-item" style="grid-column: span 2;">
          <span class="lbl">Action Date</span>
          <span class="val" style="font-weight: 800; color: #ef4444;">${lead.nextActionDate || '-'} (Due Now)</span>
        </div>
      </div>

      <div class="lead-card-notes">
        <strong>Notes:</strong> ${lead.notes}
      </div>

      <div class="lead-card-actions">
        ${lead.mainLink ? `<a href="${normalizeUrl(lead.mainLink)}" target="_blank" rel="noopener noreferrer" class="btn btn-secondary" style="padding: 6px 12px; font-size: 12px; margin-right: auto;">Open Link ↗</a>` : ''}
        <button class="btn btn-secondary btn-icon-only" onclick="openEditModal(${lead.originalIndex})" title="Edit Lead">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 20h9M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg>
        </button>
        <button class="btn btn-secondary btn-icon-only" onclick="archiveLead(${lead.originalIndex})" title="Archive Lead">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2" ry="2"/><line x1="9" y1="9" x2="15" y2="15"/><line x1="15" y1="9" x2="9" y2="15"/></svg>
        </button>
        <button class="btn btn-secondary btn-icon-only btn-danger-outline" onclick="deleteLead(${lead.originalIndex})" title="Delete Lead">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path><line x1="10" y1="11" x2="10" y2="17"></line><line x1="14" y1="11" x2="14" y2="17"></line></svg>
        </button>
      </div>
    `;
    cardsContainer.appendChild(card);
  });
}

// Modal open: Add Mode
function openAddModal() {
  const form = document.getElementById("leadForm");
  form.reset();

  const todayStr = new Date().toISOString().split('T')[0];
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
    <span>${type === 'success' ? '✓' : 'ℹ'}</span>
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
      mainLink: document.getElementById("leadMainLink").value.trim(),
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
      extraLink: document.getElementById("leadExtraLink").value.trim(),
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
  ["filterMarket", "filterSource", "filterPriority", "filterStage", "filterReply", "filterActionDate"].forEach(id => {
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

  // Script copying action buttons
  const copyBtns = document.querySelectorAll(".copy-script-btn");
  copyBtns.forEach(btn => {
    btn.addEventListener("click", () => {
      const targetId = btn.getAttribute("data-target");
      const tooltipId = btn.getAttribute("data-tooltip");
      
      const scriptText = document.getElementById(targetId).textContent;
      
      navigator.clipboard.writeText(scriptText).then(() => {
        // Show tooltip
        const tooltip = document.getElementById(tooltipId);
        tooltip.classList.add("active");
        
        btn.textContent = "Copied!";
        btn.classList.add("btn-success");
        
        setTimeout(() => {
          tooltip.classList.remove("active");
          btn.textContent = "Copy Script";
          btn.classList.remove("btn-success");
        }, 2000);
        
        showToast("Script copied to clipboard!", "success");
      }).catch(err => {
        showToast("Failed to copy script.", "error");
      });
    });
  });

  // Script Accordion toggle
  const accordionTriggers = document.querySelectorAll(".accordion-trigger");
  accordionTriggers.forEach(trigger => {
    trigger.addEventListener("click", () => {
      const parent = trigger.parentElement;
      const isActive = parent.classList.contains("active");
      
      // Close all
      document.querySelectorAll(".accordion-item").forEach(item => {
        item.classList.remove("active");
        item.querySelector(".accordion-content").style.maxHeight = null;
      });

      if (!isActive) {
        parent.classList.add("active");
        const content = parent.querySelector(".accordion-content");
        content.style.maxHeight = content.scrollHeight + "px";
      }
    });
  });

  // CSV Export trigger
  exportBtn.addEventListener("click", exportToCSV);

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
  const dateStr = new Date().toISOString().split("T")[0];
  
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
        mainLink: rowMainLink ? String(rowMainLink).trim() : "",
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

// Global functions accessible from HTML (inline onclick)
window.openEditModal = openEditModal;
window.archiveLead = archiveLead;
window.deleteLead = deleteLead;

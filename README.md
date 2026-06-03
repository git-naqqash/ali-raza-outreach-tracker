# Ali Raza Outreach Tracker

A simple, fast, and mobile-friendly outreach tracker and mini CRM designed specifically for Ali Raza's freelance white-label writing business.

## Business Context
This application tracks leads and manages client acquisition for white-label English ebooks, interactive workbook PDFs, Canva books, and lead magnets. It is optimized for targeting:
- Coaches and Course Creators
- Agencies (Content, Branding, Marketing)
- Self-Publishing and KDP Publishing Teams
- Publishing Houses and Content Agencies

**Primary Target Markets**: Italy, United States, United Kingdom, Austria, and Germany.

---

## Features

1. **Branded Dashboard**: Live counters for total active leads, channel breakdowns (Instagram, LinkedIn, Email, WhatsApp), and daily alerts (Today's Actions, Pending Follow-ups, A-Leads, and Warm Leads).
2. **Channel Filtering Tab System**: Quickly switch between channels to filter leads by outreach platform.
3. **Control Center & Search**: Full-text search matching name, company, contact person, niche, or notes, combined with advanced multi-select dropdown filters (Market, Priority, Stage, Reply Status, and Next Action Date).
4. **Single Unified Leads database**: Built on a single database structure saved safely to the browser's `localStorage` so that data is persistent and never lost upon page refresh.
5. **Interactive Tables & Mobile Layouts**:
   - **Desktop**: Compact tabular list for comprehensive workspace management.
   - **Mobile**: Touch-friendly card stack structure that displays notes, details, and shortcuts.
6. **Today's Actions Panel**: An active checklist highlighting leads whose action date is today or overdue (and are not archived).
7. **Message Scripts Repository**: A built-in accordion featuring high-converting DMs, connection notes, first emails, and follow-ups in English and Italian, equipped with quick-copy clipboard buttons.
8. **Safe Outreach Guidelines**: Static checklist of outreach safety rules, reminding you of optimal daily sending limits, multi-channel messaging rules, and WhatsApp guidelines.
9. **Export to CSV**: Easily backup your entire CRM data table into an Excel-friendly CSV sheet with one click.

---

## Tech Stack & Styling
- **Core**: HTML5, Vanilla JavaScript.
- **Styling**: Vanilla CSS3 using custom properties (CSS variables) for modern typography and the **Ali Raza Book Content** branding palette:
  - Deep Navy: `#0B1F3A`
  - Royal Blue: `#1E5BFF`
  - Teal Green: `#0FAE96`
  - Off-white: `#F7F9FC`
  - Soft Gold: `#D6A84F`

---

## How to Run the App
Since this app is a client-side Single Page Application (SPA), it does not require any backend databases or compilation steps.

1. Locate the project folder `D:\Lead Tracker`.
2. Double-click `index.html` to open the CRM directly in any modern web browser (Chrome, Edge, Safari, Firefox).
3. Alternatively, you can run a local development server in the directory if you wish to host it on a local IP:
   ```bash
   npx serve .
   # or
   python -m http.server 8000
   ```

-- Ali Raza Outreach Tracker — Neon PostgreSQL Schema
-- Run this once in the Neon SQL Editor to create required tables.

-- ── LEADS ──────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS leads (
  id               TEXT PRIMARY KEY,
  lead_name        TEXT DEFAULT '',
  contact_person   TEXT DEFAULT '',
  date_added       TEXT DEFAULT '',
  market           TEXT DEFAULT '',
  channel          TEXT DEFAULT '',
  main_link        TEXT DEFAULT '',
  niche            TEXT DEFAULT '',
  source           TEXT DEFAULT '',
  priority         TEXT DEFAULT '',
  stage            TEXT DEFAULT '',
  last_action_date TEXT DEFAULT '',
  next_action      TEXT DEFAULT '',
  next_action_date TEXT DEFAULT '',
  reply_status     TEXT DEFAULT '',
  notes            TEXT DEFAULT '',
  email            TEXT DEFAULT '',
  whatsapp_number  TEXT DEFAULT '',
  extra_link       TEXT DEFAULT '',
  followup_count   INTEGER DEFAULT 0,
  message_sent     TEXT DEFAULT '',
  created_at       TIMESTAMPTZ DEFAULT NOW()
);

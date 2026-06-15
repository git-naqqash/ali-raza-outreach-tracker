// api/leads.js — Vercel Serverless Function
// GET    /api/leads          → fetch all leads
// POST   /api/leads          → upsert all leads (body: { leads: [...] })
// DELETE /api/leads?id=xxx   → delete one lead by id
// DELETE /api/leads           → delete multiple (body: { ids: [...] })

const { neon } = require("@neondatabase/serverless");

// ── CORS headers (same origin on Vercel, but safe to include) ─────
function setCors(res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, DELETE, PUT, PATCH, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
}

// ── Field map: JS camelCase → DB snake_case row object ────────────
function leadToRow(lead) {
  return {
    id:               lead.id               || "",
    lead_name:        lead.name             || "",
    contact_person:   lead.contactPerson    || "",
    date_added:       lead.dateAdded        || "",
    market:           lead.market           || "",
    channel:          lead.channel          || "",
    main_link:        lead.mainLink         || "",
    niche:            lead.niche            || "",
    source:           lead.source           || "",
    priority:         lead.priority         || "",
    stage:            lead.stage            || "",
    last_action_date: lead.lastActionDate   || "",
    next_action:      lead.nextAction       || "",
    next_action_date: lead.nextActionDate   || "",
    reply_status:     lead.replyStatus      || "",
    notes:            lead.notes            || "",
    email:            lead.email            || "",
    whatsapp_number:  lead.whatsappNumber   || "",
    extra_link:       lead.extraLink        || "",
    followup_count:   lead.followUpCount    || 0,
    message_sent:     lead.messageSent      || ""
  };
}

// ── Field map: DB row → JS camelCase object ───────────────────────
function rowToLead(row) {
  return {
    id:             row.id,
    name:           row.lead_name         || "",
    contactPerson:  row.contact_person    || "",
    dateAdded:      row.date_added        || "",
    market:         row.market            || "",
    channel:        row.channel           || "",
    mainLink:       row.main_link         || "",
    niche:          row.niche             || "",
    source:         row.source            || "Other",
    priority:       row.priority          || "",
    stage:          row.stage             || "",
    lastActionDate: row.last_action_date  || "",
    nextAction:     row.next_action       || "",
    nextActionDate: row.next_action_date  || "",
    replyStatus:    row.reply_status      || "",
    notes:          row.notes             || "",
    email:          row.email             || "",
    whatsappNumber: row.whatsapp_number   || "",
    extraLink:      row.extra_link        || "",
    followUpCount:  row.followup_count    || 0,
    messageSent:    row.message_sent      || ""
  };
}

// ── Upsert one lead row ───────────────────────────────────────────
async function upsertLead(sql, lead) {
  const r = leadToRow(lead);
  await sql`
    INSERT INTO leads (
      id, lead_name, contact_person, date_added, market, channel,
      main_link, niche, source, priority, stage,
      last_action_date, next_action, next_action_date,
      reply_status, notes, email, whatsapp_number, extra_link,
      followup_count, message_sent
    ) VALUES (
      ${r.id}, ${r.lead_name}, ${r.contact_person}, ${r.date_added},
      ${r.market}, ${r.channel}, ${r.main_link}, ${r.niche}, ${r.source},
      ${r.priority}, ${r.stage}, ${r.last_action_date}, ${r.next_action},
      ${r.next_action_date}, ${r.reply_status}, ${r.notes}, ${r.email},
      ${r.whatsapp_number}, ${r.extra_link}, ${r.followup_count}, ${r.message_sent}
    )
    ON CONFLICT (id) DO UPDATE SET
      lead_name        = EXCLUDED.lead_name,
      contact_person   = EXCLUDED.contact_person,
      date_added       = EXCLUDED.date_added,
      market           = EXCLUDED.market,
      channel          = EXCLUDED.channel,
      main_link        = EXCLUDED.main_link,
      niche            = EXCLUDED.niche,
      source           = EXCLUDED.source,
      priority         = EXCLUDED.priority,
      stage            = EXCLUDED.stage,
      last_action_date = EXCLUDED.last_action_date,
      next_action      = EXCLUDED.next_action,
      next_action_date = EXCLUDED.next_action_date,
      reply_status     = EXCLUDED.reply_status,
      notes            = EXCLUDED.notes,
      email            = EXCLUDED.email,
      whatsapp_number  = EXCLUDED.whatsapp_number,
      extra_link       = EXCLUDED.extra_link,
      followup_count   = EXCLUDED.followup_count,
      message_sent     = EXCLUDED.message_sent
  `;
}

// ── Main handler ──────────────────────────────────────────────────
module.exports = async function handler(req, res) {
  setCors(res);
  if (req.method === "OPTIONS") return res.status(200).end();

  if (!process.env.DATABASE_URL) {
    return res.status(500).json({ error: "DATABASE_URL not configured in Vercel environment variables." });
  }

  const sql = neon(process.env.DATABASE_URL);

  // ── GET /api/leads ─────────────────────────────────────────────
  if (req.method === "GET") {
    try {
      const rows = await sql`SELECT * FROM leads ORDER BY created_at ASC`;
      return res.status(200).json({ leads: rows.map(rowToLead) });
    } catch (err) {
      console.error("[api/leads GET]", err);
      return res.status(500).json({ error: err.message });
    }
  }

  // ── POST /api/leads ────────────────────────────────────────────
  // Body: { leads: [...] }  → full sync (upsert all + delete orphans)
  // Body: { lead: {...} }   → single upsert
  if (req.method === "POST") {
    try {
      const body = req.body || {};

      // Single lead upsert
      if (body.lead) {
        await upsertLead(sql, body.lead);
        return res.status(200).json({ ok: true, upserted: 1 });
      }

      // Full sync: upsert all + delete any IDs not in the payload
      if (body.leads && Array.isArray(body.leads)) {
        const incomingLeads = body.leads;

        // Fetch existing IDs
        const existing = await sql`SELECT id FROM leads`;
        const existingIds = new Set(existing.map(r => r.id));
        const incomingIds = new Set(incomingLeads.map(l => l.id).filter(Boolean));

        // Delete orphans (leads deleted on client)
        const toDelete = [...existingIds].filter(id => !incomingIds.has(id));
        if (toDelete.length > 0) {
          for (const id of toDelete) {
            await sql`DELETE FROM leads WHERE id = ${id}`;
          }
        }

        // Upsert all incoming leads
        for (const lead of incomingLeads) {
          if (lead.id) await upsertLead(sql, lead);
        }

        return res.status(200).json({
          ok: true,
          upserted: incomingLeads.length,
          deleted: toDelete.length
        });
      }

      return res.status(400).json({ error: "Body must contain 'lead' or 'leads'" });
    } catch (err) {
      console.error("[api/leads POST]", err);
      return res.status(500).json({ error: err.message });
    }
  }

  // ── PATCH /api/leads ───────────────────────────────────────────
  if (req.method === "PATCH" || req.method === "PUT") {
    try {
      const body = req.body || {};
      const { ids, channel, updates } = body;

      const mergedUpdates = { ...updates };
      if (channel) {
        mergedUpdates.channel = channel;
      }

      if (!ids || !Array.isArray(ids) || ids.length === 0 || Object.keys(mergedUpdates).length === 0) {
        return res.status(400).json({ error: "Missing ids or updates payload" });
      }

      const DB_FIELD_MAP = {
        name:             "lead_name",
        contactPerson:    "contact_person",
        dateAdded:        "date_added",
        market:           "market",
        channel:          "channel",
        mainLink:         "main_link",
        niche:            "niche",
        source:           "source",
        priority:         "priority",
        stage:            "stage",
        lastActionDate:   "last_action_date",
        nextAction:       "next_action",
        nextActionDate:   "next_action_date",
        replyStatus:      "reply_status",
        notes:            "notes",
        email:            "email",
        whatsappNumber:   "whatsapp_number",
        extraLink:        "extra_link",
        followUpCount:    "followup_count",
        messageSent:      "message_sent"
      };

      const dbUpdates = {};
      for (const [key, val] of Object.entries(mergedUpdates)) {
        const dbCol = DB_FIELD_MAP[key] || key;
        dbUpdates[dbCol] = val;
      }

      for (const id of ids) {
        for (const [col, val] of Object.entries(dbUpdates)) {
          if (col === "channel") {
            await sql`UPDATE leads SET channel = ${val} WHERE id = ${id}`;
          } else if (col === "next_action_date") {
            await sql`UPDATE leads SET next_action_date = ${val} WHERE id = ${id}`;
          } else if (col === "last_action_date") {
            await sql`UPDATE leads SET last_action_date = ${val} WHERE id = ${id}`;
          } else if (col === "date_added") {
            await sql`UPDATE leads SET date_added = ${val} WHERE id = ${id}`;
          } else if (col === "stage") {
            await sql`UPDATE leads SET stage = ${val} WHERE id = ${id}`;
          } else if (col === "priority") {
            await sql`UPDATE leads SET priority = ${val} WHERE id = ${id}`;
          } else if (col === "reply_status") {
            await sql`UPDATE leads SET reply_status = ${val} WHERE id = ${id}`;
          }
        }
      }

      return res.status(200).json({ ok: true, updated: ids.length });
    } catch (err) {
      console.error("[api/leads PATCH/PUT]", err);
      return res.status(500).json({ error: err.message });
    }
  }

  // ── DELETE /api/leads?id=xxx ───────────────────────────────────
  if (req.method === "DELETE") {
    try {
      const id = req.query && req.query.id;
      const ids = req.body && req.body.ids;

      if (id) {
        await sql`DELETE FROM leads WHERE id = ${id}`;
        return res.status(200).json({ ok: true, deleted: id });
      }

      if (ids && Array.isArray(ids) && ids.length > 0) {
        for (const delId of ids) {
          await sql`DELETE FROM leads WHERE id = ${delId}`;
        }
        return res.status(200).json({ ok: true, deleted: ids.length });
      }

      return res.status(400).json({ error: "Provide ?id=xxx or body { ids: [...] }" });
    } catch (err) {
      console.error("[api/leads DELETE]", err);
      return res.status(500).json({ error: err.message });
    }
  }

  return res.status(405).json({ error: "Method not allowed" });
};

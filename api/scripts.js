// api/scripts.js — Vercel Serverless Function
// GET    /api/scripts        → fetch all scripts
// POST   /api/scripts        → upsert scripts (body: { scripts: [...] } or { script: {...} })
// DELETE /api/scripts?id=xxx → delete one script

const { neon } = require("@neondatabase/serverless");

function setCors(res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, DELETE, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
}

async function upsertScript(sql, script) {
  const id         = script.id         || "";
  const title      = script.title      || "";
  const channel    = script.channel    || "";
  const type       = script.type       || "";
  const body       = script.body       || "";
  const is_default = script.isDefault  || false;

  await sql`
    INSERT INTO scripts (id, title, channel, type, body, is_default)
    VALUES (${id}, ${title}, ${channel}, ${type}, ${body}, ${is_default})
    ON CONFLICT (id) DO UPDATE SET
      title      = EXCLUDED.title,
      channel    = EXCLUDED.channel,
      type       = EXCLUDED.type,
      body       = EXCLUDED.body,
      is_default = EXCLUDED.is_default
  `;
}

function rowToScript(row) {
  return {
    id:        row.id,
    title:     row.title      || "",
    channel:   row.channel    || "",
    type:      row.type       || "",
    body:      row.body       || "",
    isDefault: row.is_default || false
  };
}

module.exports = async function handler(req, res) {
  setCors(res);
  if (req.method === "OPTIONS") return res.status(200).end();

  if (!process.env.DATABASE_URL) {
    return res.status(500).json({ error: "DATABASE_URL not configured." });
  }

  const sql = neon(process.env.DATABASE_URL);

  // ── GET /api/scripts ──────────────────────────────────────────
  if (req.method === "GET") {
    try {
      const rows = await sql`SELECT * FROM scripts ORDER BY created_at ASC`;
      return res.status(200).json({ scripts: rows.map(rowToScript) });
    } catch (err) {
      console.error("[api/scripts GET]", err);
      return res.status(500).json({ error: err.message });
    }
  }

  // ── POST /api/scripts ─────────────────────────────────────────
  if (req.method === "POST") {
    try {
      const body = req.body || {};

      if (body.script) {
        await upsertScript(sql, body.script);
        return res.status(200).json({ ok: true, upserted: 1 });
      }

      if (body.scripts && Array.isArray(body.scripts)) {
        const existing = await sql`SELECT id FROM scripts`;
        const existingIds = new Set(existing.map(r => r.id));
        const incomingIds = new Set(body.scripts.map(s => s.id).filter(Boolean));

        const toDelete = [...existingIds].filter(id => !incomingIds.has(id));
        for (const id of toDelete) {
          await sql`DELETE FROM scripts WHERE id = ${id}`;
        }

        for (const script of body.scripts) {
          if (script.id) await upsertScript(sql, script);
        }

        return res.status(200).json({ ok: true, upserted: body.scripts.length, deleted: toDelete.length });
      }

      return res.status(400).json({ error: "Body must contain 'script' or 'scripts'" });
    } catch (err) {
      console.error("[api/scripts POST]", err);
      return res.status(500).json({ error: err.message });
    }
  }

  // ── DELETE /api/scripts?id=xxx ────────────────────────────────
  if (req.method === "DELETE") {
    try {
      const id = req.query && req.query.id;
      if (!id) return res.status(400).json({ error: "Provide ?id=xxx" });
      await sql`DELETE FROM scripts WHERE id = ${id}`;
      return res.status(200).json({ ok: true, deleted: id });
    } catch (err) {
      console.error("[api/scripts DELETE]", err);
      return res.status(500).json({ error: err.message });
    }
  }

  return res.status(405).json({ error: "Method not allowed" });
};

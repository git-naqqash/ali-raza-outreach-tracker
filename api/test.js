// api/test.js — Vercel Serverless Function
// GET /api/test  →  verifies Neon DB is reachable
const { neon } = require("@neondatabase/serverless");

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "GET") return res.status(405).json({ error: "Method not allowed" });

  if (!process.env.DATABASE_URL) {
    return res.status(500).json({
      ok: false,
      error: "DATABASE_URL environment variable is not set in Vercel project settings."
    });
  }

  try {
    const sql = neon(process.env.DATABASE_URL);
    const result = await sql`SELECT NOW() AS time, current_database() AS db`;
    return res.status(200).json({
      ok: true,
      time: result[0].time,
      db: result[0].db,
      message: "Neon database connected successfully!"
    });
  } catch (err) {
    console.error("[api/test] DB error:", err);
    return res.status(500).json({
      ok: false,
      error: err.message || String(err)
    });
  }
};

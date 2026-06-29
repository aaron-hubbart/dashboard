import express from "express";
import session from "express-session";
import fetch from "node-fetch";
import cors from "cors";
import fs from "fs/promises";
import path from "path";
import crypto from "crypto";

const app = express();
const PORT = 3001;
const DATA_FILE = process.env.DATA_FILE || "/data/dashboard.json";

const {
  ANTHROPIC_API_KEY,
  AZURE_CLIENT_ID,
  AZURE_CLIENT_SECRET,
  AZURE_TENANT_ID,
  AZURE_REDIRECT_URI = "https://dashboard.es-sandbox.com/auth/callback",
  SESSION_SECRET = "wbLMI9Ry6DW9GPFsuEd0zH0viGsC2mE2",
  SKIP_AUTH = "false",
  APP_VERSION = "dev",
  ASANA_TOKEN = "",
  ZOOM_TOKEN = "",
  BLINKO_TOKEN = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJyb2xlIjoic3VwZXJhZG1pbiIsIm5hbWUiOiJhaHViYmFydCIsInN1YiI6IjEiLCJleHAiOjQ5MjI5MDM3ODIsImlhdCI6MTc2OTMwMzc4Mn0.BeCaFOP7Gb4FlaNbKuXYaRozy4EYgM7R20EvSQ3ByQE",
  BLINKO_URL = "http://35.225.239.191:1111",
  ZOOM_WEBHOOK_SECRET = "",

} = process.env;

// ── Data persistence ───────────────────────────────────────
async function readData() {
  try {
    const raw = await fs.readFile(DATA_FILE, "utf8");
    return JSON.parse(raw);
  } catch {
    return {};
  }
}

async function writeData(data) {
  await fs.mkdir(path.dirname(DATA_FILE), { recursive: true });
  await fs.writeFile(DATA_FILE, JSON.stringify(data), "utf8");
}

// ── Middleware ──────────────────────────────────────────────
app.set("trust proxy", 1);
app.use(express.json({ limit: "10mb" }));
app.use(cors({ origin: "https://dashboard.es-sandbox.com", credentials: true }));
app.use(session({
  secret: SESSION_SECRET,
  resave: false,
  saveUninitialized: false,
  cookie: {
    secure: true,
    httpOnly: true,
    sameSite: "none",
    maxAge: 8 * 60 * 60 * 1000
  }
}));

app.use((req, res, next) => {
  console.log(`${req.method} ${req.path}`);
  next();
});

// ── Auth guard ─────────────────────────────────────────────
function requireAuth(req, res, next) {
  if (SKIP_AUTH === "true" || req.session.accessToken) return next();
  return res.status(401).json({ error: "Not authenticated" });
}

// ── Health / version ───────────────────────────────────────
app.get("/api/health", (req, res) => res.json({ ok: true, version: APP_VERSION, time: new Date().toISOString() }));
app.get("/api/version", (req, res) => res.json({ version: APP_VERSION }));

// ── Persistent data ────────────────────────────────────────
app.get("/api/data", requireAuth, async (req, res) => {
  const data = await readData();
  res.json({ success: true, data });
});

app.post("/api/data", requireAuth, async (req, res) => {
  try {
    await writeData(req.body);
    res.json({ success: true });
  } catch (err) {
    console.error("Data write error:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ── Admin status ───────────────────────────────────────────
app.get("/api/admin/status", requireAuth, (req, res) => {
  res.json({
    version: APP_VERSION,
    skipAuth: SKIP_AUTH === "true",
    m365SignedIn: !!req.session.accessToken,
    asanaTokenSet: !!ASANA_TOKEN,
    zoomTokenSet: !!ZOOM_TOKEN,
    blinkoTokenSet: !!BLINKO_TOKEN,
	zoomWebhookSet: !!ZOOM_WEBHOOK_SECRET,
  });
});

app.get("/api/admin/test/asana", requireAuth, async (req, res) => {
  if (!ASANA_TOKEN) return res.json({ ok: false, error: "ASANA_TOKEN not set" });
  const r = await fetch("https://app.asana.com/api/1.0/users/me", {
    headers: { Authorization: `Bearer ${ASANA_TOKEN}`, Accept: "application/json" }
  });
  const data = await r.json();
  if (!r.ok) return res.json({ ok: false, error: data?.errors?.[0]?.message || "Auth failed" });
  res.json({ ok: true, user: data.data?.name, email: data.data?.email });
});

app.get("/api/admin/test/m365", requireAuth, async (req, res) => {
  if (!req.session.accessToken) return res.json({ ok: false, error: "Not signed in to Microsoft" });
  const r = await fetch("https://graph.microsoft.com/v1.0/me?$select=displayName,mail", {
    headers: { Authorization: `Bearer ${req.session.accessToken}` }
  });
  const data = await r.json();
  if (!r.ok) return res.json({ ok: false, error: data?.error?.message || "Auth failed" });
  res.json({ ok: true, user: data.displayName, email: data.mail });
});

// ── Directory: employee search (used by customer Account Team fields) ─────
// Searches Azure AD for people matching a name or email prefix.
// Requires User.ReadBasic.All permission on the Azure app registration.
app.get("/api/directory/search", requireAuth, async (req, res) => {
  if (!req.session.accessToken)
    return res.status(401).json({ error: "Microsoft sign-in required" });
  await refreshIfNeeded(req);
  const q = (req.query.q || "").trim();
  if (q.length < 2) return res.json({ success: true, data: [] });
  // Use /users with OData filter — works for displayName or mail prefix
  const filter = `startswith(displayName,'${q.replace(/'/g, "''")}') or startswith(mail,'${q.replace(/'/g, "''")}')`;
  const r = await fetch(
    `https://graph.microsoft.com/v1.0/users?$filter=${encodeURIComponent(filter)}&$select=id,displayName,mail,jobTitle&$top=10`,
    { headers: { Authorization: `Bearer ${req.session.accessToken}` } }
  );
  const data = await r.json();
  if (!r.ok) {
    console.error("Directory search error:", data?.error?.message);
    return res.status(r.status).json({ success: false, error: data?.error?.message });
  }
  res.json({
    success: true,
    data: (data.value || []).map(u => ({
      id: u.id, name: u.displayName, email: u.mail, title: u.jobTitle || "",
    })),
  });
});

// ── Auth ───────────────────────────────────────────────────
app.get("/auth/status", (req, res) => {
  res.json({ authenticated: SKIP_AUTH === "true" || !!req.session.accessToken });
});

app.get("/auth/login", (req, res) => {
  if (SKIP_AUTH === "true") return res.redirect("/");
  const params = new URLSearchParams({
    client_id: AZURE_CLIENT_ID,
    response_type: "code",
    redirect_uri: AZURE_REDIRECT_URI,
    scope: "openid profile email Calendars.ReadWrite User.Read OnlineMeetings.Read offline_access",
    response_mode: "query",
  });
  res.redirect(`https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/authorize?${params}`);
});

app.get("/auth/callback", async (req, res) => {
  const { code, error } = req.query;
  if (error) return res.status(400).send(`Auth error: ${error}`);
  const tokenRes = await fetch(
    `https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: AZURE_CLIENT_ID,
        client_secret: AZURE_CLIENT_SECRET,
        code,
        redirect_uri: AZURE_REDIRECT_URI,
        grant_type: "authorization_code",
      }),
    }
  );
  const tokens = await tokenRes.json();
  if (tokens.error) return res.status(400).send(`Token error: ${tokens.error_description}`);
  req.session.accessToken = tokens.access_token;
  req.session.refreshToken = tokens.refresh_token;
  req.session.expiresAt = Date.now() + tokens.expires_in * 1000;
  console.log("Auth callback: session saved, sessionID:", req.sessionID);
  req.session.save(err => {
    if (err) console.error("Session save error:", err);
    res.redirect("/");
  });
});

app.get("/auth/logout", (req, res) => {
  req.session.destroy();
  res.redirect(`https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/logout?post_logout_redirect_uri=https://dashboard.es-sandbox.com`);
});

// ── Token refresh ──────────────────────────────────────────
async function refreshIfNeeded(req) {
  if (!req.session.expiresAt || Date.now() < req.session.expiresAt - 60000) return;
  const tokenRes = await fetch(
    `https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: AZURE_CLIENT_ID,
        client_secret: AZURE_CLIENT_SECRET,
        refresh_token: req.session.refreshToken,
        grant_type: "refresh_token",
        scope: "Calendars.ReadWrite User.Read OnlineMeetings.Read offline_access",
      }),
    }
  );
  const tokens = await tokenRes.json();
  if (!tokens.error) {
    req.session.accessToken = tokens.access_token;
    req.session.refreshToken = tokens.refresh_token ?? req.session.refreshToken;
    req.session.expiresAt = Date.now() + tokens.expires_in * 1000;
  }
}

// ── Calendar helpers ───────────────────────────────────────

// Read user's stored timezone; falls back to UTC
// Read the user's stored timezone preference; falls back to UTC
async function getUserTimezone() {
  try { return (await readData()).timezone || "UTC"; }
  catch { return "UTC"; }
}

// Get today's date string (YYYY-MM-DD) in a given IANA timezone.
function todayInTimezone(tz) {
  try {
    const parts = new Intl.DateTimeFormat("en-CA", {
      timeZone: tz, year: "numeric", month: "2-digit", day: "2-digit"
    }).formatToParts(new Date());
    return parts.map(p => p.value).join(""); // yields "YYYY-MM-DD" via en-CA
  } catch {
    return new Date().toISOString().split("T")[0];
  }
}

// Return the UTC offset in milliseconds for a given timezone at a given timestamp.
// Positive = timezone is ahead of UTC, negative = behind.
function getUTCOffsetMs(tz, timestamp) {
  const fmt = new Intl.DateTimeFormat("en-CA", {
    timeZone: tz,
    year: "numeric", month: "2-digit", day: "2-digit",
    hour: "2-digit", minute: "2-digit", second: "2-digit",
    hour12: false,
  });
  const p = Object.fromEntries(fmt.formatToParts(new Date(timestamp)).map(x => [x.type, x.value]));
  const localAsUTC = Date.UTC(
    parseInt(p.year), parseInt(p.month) - 1, parseInt(p.day),
    parseInt(p.hour), parseInt(p.minute), parseInt(p.second)
  );
  return localAsUTC - timestamp; // e.g. CDT = -18000000 (-5h)
}

// Build exact UTC start/end for a calendar day (YYYY-MM-DD) in the user's timezone.
// Example: "2025-05-27" in America/Chicago (CDT=UTC-5)
//   → start: 2025-05-27T05:00:00Z  (local midnight)
//   → end:   2025-05-28T04:59:59Z  (local 23:59:59)
function dayBounds(dateStr, tz) {
  const [y, mo, d] = dateStr.split("-").map(Number);
  const approxMs = Date.UTC(y, mo - 1, d, 12, 0, 0); // use noon to avoid DST edge cases
  const offsetMs = getUTCOffsetMs(tz, approxMs);
  // local midnight = UTC midnight - offset
  const startMs = Date.UTC(y, mo - 1, d, 0, 0, 0) - offsetMs;
  const endMs   = startMs + 24 * 3600 * 1000 - 1000; // +23:59:59
  return { start: new Date(startMs).toISOString(), end: new Date(endMs).toISOString() };
}

// Map a Graph calendar event to the shape the frontend expects.
// IMPORTANT: do NOT append "Z" to dateTime — Graph returns times already in
// the requested timezone (via Prefer header). Adding "Z" wrongly marks local
// time as UTC, which shifts all events by your UTC offset.
function mapCalendarEvent(e) {
  return {
    id:              e.id,
    subject:         e.subject,
    start:           e.start?.dateTime,
    end:             e.end?.dateTime,
    isAllDay:        e.isAllDay || false,
    categories:      e.categories || [],
    attendees:       (e.attendees || []).map(a => a.emailAddress?.address).filter(Boolean),
    isOnlineMeeting: e.isOnlineMeeting || false,
    body:            e.body?.content
                       ? e.body.content.replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim().slice(0, 1000)
                       : "",
  };
}

// GET /api/calendar/today — optional ?date=YYYY-MM-DD for day navigation
app.get("/api/calendar/today", requireAuth, async (req, res) => {
  if (!req.session.accessToken)
    return res.status(401).json({ error: "Calendar requires Microsoft sign-in" });
  await refreshIfNeeded(req);
  const tz      = await getUserTimezone();
  // Client always passes ?date — fallback computes today in the user's timezone
  const dateStr = req.query.date || todayInTimezone(tz);
  const { start, end } = dayBounds(dateStr, tz);
  const graphRes = await fetch(
    `https://graph.microsoft.com/v1.0/me/calendarView` +
    `?startDateTime=${encodeURIComponent(start)}&endDateTime=${encodeURIComponent(end)}` +
    `&$select=id,subject,start,end,attendees,isOnlineMeeting,isAllDay,categories,body` +
    `&$orderby=start/dateTime&$top=50`,
    { headers: { Authorization: `Bearer ${req.session.accessToken}`, Prefer: `outlook.timezone="${tz}"` } }
  );
  if (!graphRes.ok) {
    const err = await graphRes.json();
    console.error("Graph calendar error:", err);
    return res.status(graphRes.status).json({ success: false, error: err });
  }
  const data = await graphRes.json();
  res.json({ success: true, data: (data.value || []).map(mapCalendarEvent) });
});

// GET /api/calendar/day — ?date=YYYY-MM-DD or legacy ?start=&end= (time suggestions)
app.get("/api/calendar/day", requireAuth, async (req, res) => {
  if (!req.session.accessToken)
    return res.status(401).json({ error: "Calendar requires Microsoft sign-in" });
  await refreshIfNeeded(req);
  const tz = await getUserTimezone();
  let start, end;
  if (req.query.date) {
    ({ start, end } = dayBounds(req.query.date, tz));
  } else {
    start = req.query.start;
    end   = req.query.end;
    if (!start || !end) return res.status(400).json({ error: "date or start+end required" });
  }
  const graphRes = await fetch(
    `https://graph.microsoft.com/v1.0/me/calendarView` +
    `?startDateTime=${encodeURIComponent(start)}&endDateTime=${encodeURIComponent(end)}` +
    `&$select=id,subject,start,end,attendees,isOnlineMeeting,isAllDay,categories,body` +
    `&$orderby=start/dateTime&$top=50`,
    { headers: { Authorization: `Bearer ${req.session.accessToken}`, Prefer: `outlook.timezone="${tz}"` } }
  );
  if (!graphRes.ok) {
    const err = await graphRes.json();
    return res.status(graphRes.status).json({ success: false, error: err });
  }
  const data = await graphRes.json();
  res.json({ success: true, data: (data.value || []).map(mapCalendarEvent) });
});

// GET /api/calendar/categories — Outlook master categories
app.get("/api/calendar/categories", requireAuth, async (req, res) => {
  if (!req.session.accessToken)
    return res.status(401).json({ error: "Calendar requires Microsoft sign-in" });
  await refreshIfNeeded(req);
  const r = await fetch("https://graph.microsoft.com/v1.0/me/outlook/masterCategories",
    { headers: { Authorization: `Bearer ${req.session.accessToken}` } }
  );
  const data = await r.json();
  if (!r.ok) return res.status(r.status).json({ success: false, error: data });
  res.json({ success: true, data: (data.value || []).map(c => ({ displayName: c.displayName, color: c.color })) });
});


function asanaFetch(path, opts = {}) {
  return fetch(`https://app.asana.com/api/1.0${path}`, {
    ...opts,
    headers: {
      Authorization: `Bearer ${ASANA_TOKEN}`,
      Accept: "application/json",
      "Content-Type": "application/json",
      ...(opts.headers || {}),
    },
  });
}

app.get("/api/asana/portfolios", requireAuth, async (req, res) => {
  if (!ASANA_TOKEN) return res.status(400).json({ error: "ASANA_TOKEN not configured" });
  const wsRes = await asanaFetch("/workspaces");
  const wsData = await wsRes.json();
  const workspaceGid = wsData.data?.[0]?.gid;
  if (!workspaceGid) return res.status(400).json({ error: "No workspace found" });
  const r = await asanaFetch(`/portfolios?workspace=${workspaceGid}&owner=me&opt_fields=gid,name&limit=100`);
  const data = await r.json();
  if (!r.ok) return res.status(r.status).json({ error: data?.errors?.[0]?.message });
  res.json({ success: true, data: data.data || [] });
});

async function getProjectsFromPortfolio(gid, depth = 0) {
  if (depth > 3) return [];
  const r = await asanaFetch(`/portfolios/${gid}/items?opt_fields=gid,name,resource_type&limit=100`);
  const data = await r.json();
  if (!r.ok) return [];
  const projects = [];
  for (const item of (data.data || [])) {
    if (item.resource_type === "project") {
      projects.push(item);
    } else if (item.resource_type === "portfolio") {
      const nested = await getProjectsFromPortfolio(item.gid, depth + 1);
      projects.push(...nested);
    }
  }
  return projects;
}

app.get("/api/asana/portfolios/:gid/projects", requireAuth, async (req, res) => {
  if (!ASANA_TOKEN) return res.status(400).json({ error: "ASANA_TOKEN not configured" });
  const projects = await getProjectsFromPortfolio(req.params.gid);
  res.json({ success: true, data: projects });
});

app.get("/api/asana/tasks", requireAuth, async (req, res) => {
  if (!ASANA_TOKEN) return res.status(400).json({ error: "ASANA_TOKEN not configured" });
  const due = req.query.due_before || new Date(Date.now() + 7 * 86400000).toISOString().split("T")[0];
  const projectGids = req.query.project_gids ? req.query.project_gids.split(",").filter(Boolean) : [];
  if (projectGids.length > 0) {
    const allTasks = [];
    for (const gid of projectGids) {
      const r = await asanaFetch(`/tasks?project=${gid}&completed_since=now&opt_fields=gid,name,completed,due_on,assignee.name,projects.gid,projects.name,memberships.section.gid,memberships.section.name,memberships.project.gid&limit=100`);
      const data = await r.json();
      if (r.ok && data.data) allTasks.push(...data.data.filter(t => !t.due_on || t.due_on <= due));
    }
    const seen = new Set();
    return res.json({ success: true, data: allTasks.filter(t => { if (seen.has(t.gid)) return false; seen.add(t.gid); return true; }) });
  }
  const wsRes = await asanaFetch("/workspaces");
  const wsData = await wsRes.json();
  const workspaceGid = wsData.data?.[0]?.gid;
  if (!workspaceGid) return res.status(400).json({ error: "No workspace found" });
  const r = await asanaFetch(`/tasks?assignee=me&workspace=${workspaceGid}&completed_since=now&due_on.before=${due}&opt_fields=gid,name,due_on,projects.name,projects.gid&limit=100`);
  const data = await r.json();
  if (!r.ok) return res.status(r.status).json({ error: data?.errors?.[0]?.message });
  res.json({ success: true, data: data.data || [] });
});

app.get("/api/asana/projects", requireAuth, async (req, res) => {
  if (!ASANA_TOKEN) return res.status(400).json({ error: "ASANA_TOKEN not configured" });
  const wsRes = await asanaFetch("/workspaces");
  const wsData = await wsRes.json();
  const workspaceGid = wsData.data?.[0]?.gid;
  if (!workspaceGid) return res.status(400).json({ error: "No workspace found" });
  const r = await asanaFetch(`/projects?workspace=${workspaceGid}&archived=false&opt_fields=gid,name&limit=100`);
  const data = await r.json();
  if (!r.ok) return res.status(r.status).json({ error: data?.errors?.[0]?.message });
  res.json({ success: true, data: data.data || [] });
});

app.get("/api/asana/projects/:gid/tasks", requireAuth, async (req, res) => {
  if (!ASANA_TOKEN) return res.status(400).json({ error: "ASANA_TOKEN not configured" });
  const r = await asanaFetch(`/projects/${req.params.gid}/tasks?opt_fields=gid,name,completed,due_on,assignee.name,memberships.section.gid,memberships.section.name&limit=100`);
  const data = await r.json();
  if (!r.ok) return res.status(r.status).json({ error: data?.errors?.[0]?.message });
  res.json({ success: true, data: data.data || [] });
});

app.get("/api/asana/projects/:gid/sections", requireAuth, async (req, res) => {
  const r = await asanaFetch(`/projects/${req.params.gid}/sections?opt_fields=gid,name&limit=100`);
  const data = await r.json();
  if (!r.ok) return res.status(r.status).json({ error: data?.errors?.[0]?.message });
  res.json({ success: true, data: data.data });
});

app.patch("/api/asana/tasks/:gid/complete", requireAuth, async (req, res) => {
  if (!ASANA_TOKEN) return res.status(400).json({ error: "ASANA_TOKEN not configured" });
  const r = await asanaFetch(`/tasks/${req.params.gid}`, { method: "PUT", body: JSON.stringify({ data: { completed: true } }) });
  const data = await r.json();
  if (!r.ok) return res.status(r.status).json({ error: data?.errors?.[0]?.message });
  res.json({ success: true });
});

app.patch("/api/asana/tasks/:gid/due", requireAuth, async (req, res) => {
  if (!ASANA_TOKEN) return res.status(400).json({ error: "ASANA_TOKEN not configured" });
  const { due_on } = req.body;
  const r = await asanaFetch(`/tasks/${req.params.gid}`, { method: "PUT", body: JSON.stringify({ data: { due_on: due_on || null } }) });
  const data = await r.json();
  if (!r.ok) return res.status(r.status).json({ error: data?.errors?.[0]?.message });
  res.json({ success: true });
});

// ── Blinko direct REST ─────────────────────────────────────
//const BLINKO_URL = process.env.BLINKO_URL || "http://35.225.239.191:1111";

app.post("/api/blinko/notes", requireAuth, async (req, res) => {
  if (!BLINKO_TOKEN) return res.status(400).json({ error: "BLINKO_TOKEN not configured" });
  const { content } = req.body;
  const r = await fetch(`${BLINKO_URL}/api/v1/note/upsert`, {
    method: "POST",
    headers: { "Content-Type": "application/json", "Authorization": `Bearer ${BLINKO_TOKEN}` },
    body: JSON.stringify({ content, type: 0 })
  });
  const data = await r.json();
  if (!r.ok) return res.status(r.status).json({ error: data?.message || "Blinko error" });
  res.json({ success: true, data: data });
});

// POST /api/asana/tasks — create a single task
app.post("/api/asana/tasks", requireAuth, async (req, res) => {
  if (!ASANA_TOKEN) return res.status(400).json({ error: "ASANA_TOKEN not configured" });
  const { name, due_on, notes: taskNotes } = req.body;
  const r = await asanaFetch("/tasks", {
    method: "POST",
    body: JSON.stringify({ data: { name, due_on: due_on || null, notes: taskNotes || "", assignee: "me" } })
  });
  const data = await r.json();
  if (!r.ok) return res.status(r.status).json({ error: data?.errors?.[0]?.message });
  res.json({ success: true, data: data.data });
});

// ── Zoom direct REST ───────────────────────────────────────
const ZOOM_BASE = "https://api.zoom.us/v2";

async function zoomFetch(path, opts = {}) {
  return fetch(`${ZOOM_BASE}${path}`, {
    ...opts,
    headers: { "Authorization": `Bearer ${ZOOM_TOKEN}`, "Content-Type": "application/json", ...(opts.headers || {}) }
  });
}

// GET /api/zoom/meetings?date=YYYY-MM-DD&q=meeting+title
app.get("/api/zoom/meetings", requireAuth, async (req, res) => {
  if (!ZOOM_TOKEN) return res.status(400).json({ error: "ZOOM_TOKEN not configured" });
  const { date, q } = req.query;
  const from = date || new Date().toISOString().split("T")[0];
  const to = from;
  const r = await zoomFetch(`/users/me/meetings?type=previous_meetings&from=${from}&to=${to}&page_size=30`);
  const data = await r.json();
  if (!r.ok) return res.status(r.status).json({ error: data?.message || "Zoom error" });
  // Filter by title match if q provided
  const meetings = (data.meetings || []).filter(m =>
    !q || m.topic?.toLowerCase().includes(q.toLowerCase())
  ).map(m => ({ id: m.id, uuid: m.uuid, topic: m.topic, start_time: m.start_time, duration: m.duration }));
  res.json({ success: true, data: meetings });
});

// GET /api/zoom/meetings/:id/summary
app.get("/api/zoom/meetings/:id/summary", requireAuth, async (req, res) => {
  if (!ZOOM_TOKEN) return res.status(400).json({ error: "ZOOM_TOKEN not configured" });
  const r = await zoomFetch(`/meetings/${req.params.id}/meeting_summary`);
  const data = await r.json();
  if (!r.ok) return res.status(r.status).json({ error: data?.message || "No summary available" });
  const text = [
    data.summary_overview,
    data.summary_details?.map(d => `${d.label}: ${d.summary}`).join("\n"),
    data.next_steps?.map(s => `- ${s.summary}`).join("\n")
  ].filter(Boolean).join("\n\n");
  res.json({ success: true, data: text || null });
});




// ── Zoom webhook ───────────────────────────────────────────
// Receives meeting.summary_completed events from Zoom.
// No auth token needed — validated by ZOOM_WEBHOOK_SECRET.
app.post("/api/zoom/webhook", express.raw({ type: "application/json" }), async (req, res) => {
  const body = req.body;
  const bodyStr = body.toString();

  // Zoom URL validation challenge (fires when you first register the endpoint)
  let parsed;
  try { parsed = JSON.parse(bodyStr); } catch { return res.status(400).send("Bad JSON"); }

  if (parsed.event === "endpoint.url_validation") {
    if (!ZOOM_WEBHOOK_SECRET) return res.status(500).json({ error: "ZOOM_WEBHOOK_SECRET not configured" });
    const hash = crypto.createHmac("sha256", ZOOM_WEBHOOK_SECRET)
      .update(`v0:${req.headers["x-zm-request-timestamp"]}:${bodyStr}`)
      .digest("hex");
    return res.json({ plainToken: parsed.payload.plainToken, encryptedToken: `v0=${hash}` });
  }

  // Validate signature on all other events
  if (ZOOM_WEBHOOK_SECRET) {
    const ts = req.headers["x-zm-request-timestamp"];
    const sig = req.headers["x-zm-signature"];
    const expected = "v0=" + crypto.createHmac("sha256", ZOOM_WEBHOOK_SECRET)
      .update(`v0:${ts}:${bodyStr}`)
      .digest("hex");
    if (sig !== expected) {
      console.warn("Zoom webhook: invalid signature");
      return res.status(401).json({ error: "Invalid signature" });
    }
  }

  if (parsed.event !== "meeting.summary_completed") {
    return res.status(200).json({ ok: true, ignored: true });
  }

  const payload = parsed.payload?.object;
  if (!payload) return res.status(400).json({ error: "No payload" });

  const topic       = payload.topic || "Untitled meeting";
  const startTime   = payload.start_time || new Date().toISOString();
  const hostEmail   = payload.host_email || "";
  const summary     = payload.summary_overview || "";
  const details     = (payload.summary_details || []).map(d => `**${d.label}:** ${d.summary}`).join("\n\n");
  const nextSteps   = (payload.next_steps || []).map(s => `- [ ] ${s.summary}`).join("\n");

  const content = [
    `# ${topic}`,
    `**Date:** ${new Date(startTime).toLocaleDateString("en-US", { weekday: "long", year: "numeric", month: "long", day: "numeric" })}`,
    hostEmail ? `**Host:** ${hostEmail}` : "",
    "",
    summary ? `## Summary\n${summary}` : "",
    details ? `## Details\n${details}` : "",
    nextSteps ? `## Next Steps\n${nextSteps}` : "",
  ].filter(Boolean).join("\n");

  console.log(`Zoom webhook: summary_completed for "${topic}" — saving to Blinko`);

  if (!BLINKO_TOKEN) {
    console.error("Zoom webhook: BLINKO_TOKEN not set, cannot save note");
    return res.status(500).json({ error: "BLINKO_TOKEN not configured" });
  }

  try {
    const bRes = await fetch(`${BLINKO_URL}/api/v1/note/upsert`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "Authorization": `Bearer ${BLINKO_TOKEN}` },
      body: JSON.stringify({ content, type: 1 })
    });
    const bData = await bRes.json();
    if (!bRes.ok) {
      console.error("Zoom webhook: Blinko error:", bData);
      return res.status(500).json({ error: "Failed to save to Blinko" });
    }
    console.log(`Zoom webhook: note saved to Blinko for "${topic}"`);
    res.json({ ok: true });
  } catch (e) {
    console.error("Zoom webhook: exception:", e.message);
    res.status(500).json({ error: e.message });
  }
});


// ── Anthropic proxy ────────────────────────────────────────
app.post("/api/claude", requireAuth, async (req, res) => {
  await refreshIfNeeded(req);
  const body = { ...req.body };
  const m365Mode = req.headers["x-m365-mode"] || "direct";
  console.log(`/api/claude model=${body.model} mcp=${JSON.stringify((body.mcp_servers||[]).map(s=>s.name))} m365=${m365Mode}`);
  if (Array.isArray(body.mcp_servers)) {
    body.mcp_servers = body.mcp_servers.map(s => {
      if (s.url?.includes("microsoft365")) {
        if (m365Mode === "direct" && req.session.accessToken) return { ...s, authorization_token: req.session.accessToken };
        return s;
      }
      if (s.url?.includes("asana") && ASANA_TOKEN) return { ...s, authorization_token: ASANA_TOKEN };
      if (s.url?.includes("zoom") && ZOOM_TOKEN) return { ...s, authorization_token: ZOOM_TOKEN };
      if ((s.url?.includes("blinko") || s.name === "blinko") && BLINKO_TOKEN) return { ...s, authorization_token:BLINKO_TOKEN };
      return s;
    });
  }
  const upstream = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": ANTHROPIC_API_KEY,
      "anthropic-version": "2023-06-01",
      "anthropic-beta": "mcp-client-2025-04-04",
    },
    body: JSON.stringify(body),
  });
  if (!upstream.ok) {
    const errText = await upstream.text();
    console.error(`Anthropic ${upstream.status}:`, errText);
    return res.status(upstream.status).send(errText);
  }
  console.log(`Anthropic ${upstream.status} ok`);
  res.status(upstream.status);
  upstream.headers.forEach((v, k) => {
    if (!["content-encoding", "transfer-encoding", "connection"].includes(k)) res.setHeader(k, v);
  });
  upstream.body.pipe(res);
});

// ── Error handler ──────────────────────────────────────────
app.use((err, req, res, next) => {
  console.error("Express error:", err.message);
  res.status(err.status || 500).json({ error: err.message });
});

app.listen(PORT, () => {
  console.log(`Proxy v${APP_VERSION} listening on :${PORT}`);
  console.log(`ASANA_TOKEN set: ${!!ASANA_TOKEN} (length: ${ASANA_TOKEN.length})`);
  console.log(`SKIP_AUTH: ${SKIP_AUTH}`);
});
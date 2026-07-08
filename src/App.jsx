import { useState, useEffect, useRef } from "react";

const STORAGE_KEY = "dash_v5";
const ASANA_MCP_URL = "https://mcp.asana.com/v2/mcp";
const M365_MCP = "https://microsoft365.mcp.claude.com/mcp";
const ZOOM_MCP = "https://mcp.zoom.us/mcp/zoom/streamable";
const ZOOM_RECORDINGS_MCP = "https://mcp-us.zoom.us/mcp/zoom/streamable";
const BLINKO_MCP = "http://35.225.239.191:1111/mcp";

function getTodayISO() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
}
function getWeekEndISO() {
  const d = new Date(); d.setDate(d.getDate() + 7);
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
}
function dateToISO(d) {
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
}
function fmtHours(h) { if (!h) return "0h"; return Number.isInteger(h) ? `${h}h` : `${h.toFixed(1)}h`; }

// ── Customer panel sub-components ─────────────────────────────────────────
// Defined at module level so React sees stable references across renders.
// If defined inside the component they'd be recreated every render, causing
// input fields to lose focus after every keystroke.
function PanelSection({ label, accent, border, children }) {
  return (
    <div style={{ marginBottom: 16 }}>
      <div style={{ fontSize: 10, fontWeight: 700, color: accent, textTransform: "uppercase", letterSpacing: "0.08em", paddingBottom: 5, marginBottom: 8, borderBottom: `1px solid ${accent}33` }}>
        {label}
      </div>
      {children}
    </div>
  );
}

function FieldLabel({ icon, color, children }) {
  return (
    <label style={{ fontSize: 11, fontWeight: 600, color: color || "#888", textTransform: "uppercase", letterSpacing: "0.05em", display: "flex", alignItems: "center", gap: 5, marginBottom: 4 }}>
      {icon && <i className={`ti ${icon}`} style={{ fontSize: 12 }} />}
      {children}
    </label>
  );
}

function loadLocal() {
  try { return JSON.parse(localStorage.getItem(STORAGE_KEY)) || {}; } catch { return {}; }
}
function saveLocal(d) { try { localStorage.setItem(STORAGE_KEY, JSON.stringify(d)); } catch {} }

let saveTimer = null;
function scheduleSave(data) {
  saveLocal(data);
  clearTimeout(saveTimer);
  saveTimer = setTimeout(async () => {
    try {
      await fetch("/api/data", {
        method: "POST", credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(data)
      });
    } catch (e) { console.error("Server save failed:", e); }
  }, 1500);
}

async function callClaude(messages, system, mcpServers = [], m365Mode = "direct") {
  const res = await fetch("/api/claude", {
    method: "POST", credentials: "include",
    headers: { "Content-Type": "application/json", "x-m365-mode": m365Mode },
    body: JSON.stringify({ model: "claude-sonnet-4-6", max_tokens: 4000, system, messages, mcp_servers: mcpServers })
  });
  return res.json();
}

async function agentCall(instruction, mcpServers, m365Mode = "direct") {
  const sys = `You are a productivity assistant. Execute the instruction using available MCP tools.
Respond ONLY with valid JSON (no markdown fences, no preamble):
{"success": true, "data": <result>} or {"success": false, "error": "reason"}`;
  const data = await callClaude([{ role: "user", content: instruction }], sys, mcpServers, m365Mode);
  const text = (data.content || []).filter(b => b.type === "text").map(b => b.text).join("");
  try { return JSON.parse(text.replace(/```json|```/g, "").trim()); }
  catch { return { success: false, error: "Parse error: " + text.slice(0, 300) }; }
}

// For calls with no MCP servers — uses a stricter system prompt that prevents
// Claude from attempting to fetch URLs or use tools that aren't available.
async function generateCall(instruction) {
  const sys = `You are a productivity assistant that generates structured data.
You have NO tools available. Do NOT attempt to fetch URLs, call APIs, or use any tools.
If a URL is provided in the input, treat it as a reference only — do not try to access it.
Respond ONLY with valid JSON (no markdown fences, no preamble, no explanation):
{"success": true, "data": <result>} or {"success": false, "error": "reason"}`;
  const data = await callClaude([{ role: "user", content: instruction }], sys, [], "direct");
  const text = (data.content || []).filter(b => b.type === "text").map(b => b.text).join("");
  try { return JSON.parse(text.replace(/```json|```/g, "").trim()); }
  catch { return { success: false, error: "Parse error: " + text.slice(0, 300) }; }
}

async function asanaREST(path) {
  const r = await fetch(path, { credentials: "include" });
  return r.json();
}

const DEFAULT_BUCKETS = [
  { id: "boa", name: "Bank of America", type: "customer", domains: ["bankofamerica.com", "bofa.com"] },
  { id: "internal-sales", name: "Sales / Presales", type: "internal", domains: [] },
  { id: "internal-eng", name: "Engineering", type: "internal", domains: [] },
  { id: "internal-gen", name: "General Internal", type: "internal", domains: [] },
];

const STATUS_C = {
  "On Track": { bg: "#14532d22", text: "#15803d", border: "#16a34a" },
  "At Risk":  { bg: "#78350f22", text: "#d97706", border: "#f59e0b" },
  "Blocked":  { bg: "#7f1d1d22", text: "#ef4444", border: "#f87171" },
};

const TABS = ["meetings", "goals", "projects", "customers", "time", "admin"];
const TAB_LABELS = { meetings: "Meetings", goals: "Goals", projects: "Projects", customers: "Customers", time: "Time", admin: "⚙ Admin" };

export default function App() {
  const saved = loadLocal();

  // ── UI ─────────────────────────────────────────────────────
  const [dark, setDark] = useState(saved.dark ?? true);
  const [tab, setTab] = useState("meetings");
  const [authed, setAuthed] = useState(null);
  const [version, setVersion] = useState("...");
  const [loading, setLoading] = useState({});
  const [dataLoaded, setDataLoaded] = useState(false);
  const setLoad = (k, v) => setLoading(l => ({ ...l, [k]: v }));

  // ── Config ─────────────────────────────────────────────────
  const [asanaMode, setAsanaMode] = useState(saved.asanaMode || "rest");
  const [m365Mode, setM365Mode] = useState(saved.m365Mode || "direct");
  const [zoomMode, setZoomMode] = useState(saved.zoomMode || "claude"); // "claude" | "token"
  const [selectedPortfolios, setSelectedPortfolios] = useState(saved.selectedPortfolios || []);
  const [portfolios, setPortfolios] = useState(saved.portfolios || []);

  // ── Goals ──────────────────────────────────────────────────
  const [goals, setGoals] = useState(saved.goals || []);
  const [asanaTasks, setAsanaTasks] = useState([]);
  const [showPicker, setShowPicker] = useState(false);
  const [expandedGroups, setExpandedGroups] = useState({});
  const [newGoalText, setNewGoalText] = useState("");

  // ── Projects ───────────────────────────────────────────────
  const [projects, setProjects] = useState(saved.projects || []);
  const [projTasks, setProjTasks] = useState({});   // gid → [task]  (flat, for legacy view)
  const [projSections, setProjSections] = useState({}); // gid → [{ gid, name, tasks[] }]
  const [expanded, setExpanded] = useState({});
  const [goalsProjects, setGoalsProjects] = useState(saved.goalsProjects || {});
  // goalsSections: { [sectionGid]: bool } — true = include in goals picker
  const [goalsSections, setGoalsSections] = useState(saved.goalsSections || {});

  // ── Meetings ───────────────────────────────────────────────
  const [meetings, setMeetings] = useState(saved.meetings || []);
  const [meetingMeta, setMeetingMeta] = useState(saved.meetingMeta || {});
  // Store the fetched date explicitly — never derive from offset so filter always matches fetch
  const [meetingDay, setMeetingDay] = useState(getTodayISO());
  const meetingDateRef = useRef(null);
  const suggestionDateRef = useRef(null);
  const [outlookCategories, setOutlookCategories] = useState(saved.outlookCategories || []);
  const [excludedCategories, setExcludedCategories] = useState(saved.excludedCategories || []);
  const [loadingCategories, setLoadingCategories] = useState(false);

  // ── Time ───────────────────────────────────────────────────
  const [buckets, setBuckets] = useState(saved.buckets || DEFAULT_BUCKETS);
  const [timeEntries, setTimeEntries] = useState(saved.timeEntries || []);
  const [newBucket, setNewBucket] = useState({ name: "", type: "customer", domains: "" });
  const [newEntry, setNewEntry] = useState({ bucketId: "", hours: "", note: "", date: getTodayISO() });
  const [weekOffset, setWeekOffset] = useState(0);
  // suggestionDay: ISO string for the meeting suggestions block (independent of log form)
  const [suggestionDay, setSuggestionDay] = useState(getTodayISO());
  const [editingEntry, setEditingEntry] = useState(null);
  const [dismissedSuggestions, setDismissedSuggestions] = useState(saved.dismissedSuggestions || {});
  const [suggestedMeetings, setSuggestedMeetings] = useState([]);
  const [taskDueDates, setTaskDueDates] = useState({});
  const [loadingSuggestions, setLoadingSuggestions] = useState(false);

  // ── Process panel ──────────────────────────────────────────
  // { meetingId, meeting, steps: { zoom, notes, blinko, asana } }
  // each step: { status: "idle"|"running"|"done"|"error"|"skipped", message: "", result: null }
  const [processPanel, setProcessPanel] = useState(null);

  // ── Zoom Recordings ────────────────────────────────────────
  const [meetingsSubTab, setMeetingsSubTab] = useState("calendar"); // "calendar" | "recordings"
  const [recordings, setRecordings] = useState([]);
  const [recordingsRange, setRecordingsRange] = useState(() => {
    const to = new Date();
    const from = new Date();
    from.setDate(from.getDate() - 30);
    const fmt = d => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
    return { from: fmt(from), to: fmt(to) };
  });
  const [selectedRecording, setSelectedRecording] = useState(null);
  const [recordingAssets, setRecordingAssets] = useState(null); // { summary, transcript }
  const [recordingsAssetsTab, setRecordingsAssetsTab] = useState("summary");

  // ── Admin ──────────────────────────────────────────────────
  const [adminStatus, setAdminStatus] = useState(null);
  const [testResults, setTestResults] = useState({});
  const [timezone, setTimezone] = useState(saved.timezone || "America/Chicago");
  const [directorySearchEnabled, setDirectorySearchEnabled] = useState(saved.directorySearchEnabled ?? false);
  // Track collapsed state for admin sections — all start collapsed to keep admin clean
  const [adminCollapsed, setAdminCollapsed] = useState({ categories: true, portfolios: true, buckets: true });

  // ── Customers ──────────────────────────────────────────────
  // Each customer: { id, name, notes, sfAccountId, sfSynced,
  //   financials: { arr, renewalDate, renewalAmount, ltv, ela },
  //   team: { ae, se, csm, tam },
  //   links: { sf_account, sf_opps[], slack_primary, slack_supporting[],
  //            asana, gdrive, infra360, briefing, briefingDate, jira_all, jira_open } }
  const [customers, setCustomers] = useState(saved.customers || []);
  const [customerSearch, setCustomerSearch] = useState("");
  const [editingCustomer, setEditingCustomer] = useState(null);
  const [customerDraft, setCustomerDraft] = useState(null);
  const [sfSyncing, setSfSyncing] = useState(false);
  // { fieldKey: [{ id, name, email, title }] } — dropdown suggestions per team field
  const [teamSuggestions, setTeamSuggestions] = useState({});
  const [customerPanelWidth, setCustomerPanelWidth] = useState(480);
  const [missingDataModal, setMissingDataModal] = useState(null); // null | { customerName, loading, prompt, result, error }
  const [customerColumns, setCustomerColumns] = useState(saved.customerColumns || [{ label: "All Customers", industry: "*" }]);

  // ── Derived ────────────────────────────────────────────────
  const today = getTodayISO();

  function getDayISO(offset = 0) {
    const d = new Date();
    d.setDate(d.getDate() + offset);
    return dateToISO(d);
  }

  // Navigate by skipping weekends
  function nextWeekday(offset, direction) {
    let d = new Date();
    d.setDate(d.getDate() + offset + direction);
    while (d.getDay() === 0 || d.getDay() === 6) d.setDate(d.getDate() + direction);
    const diff = Math.round((d - new Date(new Date().toDateString())) / 86400000);
    return diff;
  }

  // Convert an ISO date string to offset from today
  function isoToOffset(iso) {
    const target = new Date(iso + "T12:00:00");
    const todayDate = new Date(new Date().toDateString());
    return Math.round((target - todayDate) / 86400000);
  }

  function getWeekDays(offset = 0) {
    const now = new Date();
    const day = now.getDay();
    const mon = new Date(now);
    mon.setDate(now.getDate() - (day === 0 ? 6 : day - 1) + offset * 7);
    return Array.from({ length: 5 }, (_, i) => {
      const d = new Date(mon); d.setDate(mon.getDate() + i);
      return dateToISO(d);
    });
  }

  const weekDays = getWeekDays(weekOffset);
  const weekLabel = weekOffset === 0 ? "This week" : weekOffset === -1 ? "Last week"
    : `Week of ${new Date(weekDays[0] + "T12:00:00").toLocaleDateString("en-US", { month: "short", day: "numeric" })}`;
  const weekMonday = weekDays[0];

  // Log-entry date comes from the form's own date field (newEntry.date)
  const logEntries = timeEntries.filter(e => e.date === (newEntry.date || today));
  const logBucketTotals = {};
  buckets.forEach(b => { logBucketTotals[b.id] = 0; });
  logEntries.forEach(e => { logBucketTotals[e.bucketId] = (logBucketTotals[e.bucketId] || 0) + (e.hours || 0); });
  const logTotalHours = Object.values(logBucketTotals).reduce((s, v) => s + v, 0);

  const isSuggestionToday = suggestionDay === today;
  const suggestionDayLabel = new Date(suggestionDay + "T12:00:00").toLocaleDateString("en-US", { weekday: "short", month: "short", day: "numeric" });

  // Meetings for suggestion day — filter dismissed + excluded categories
  const dayMeetings = (suggestedMeetings.length > 0 ? suggestedMeetings
    : meetings.filter(m => m.start?.split("T")[0] === suggestionDay))
    .filter(m => !dismissedSuggestions[m.id])
    .filter(m => !(m.categories || []).some(c => excludedCategories.includes(c)));
  const timeSuggestions = dayMeetings
    .filter(m => m.durationMins > 0)
    .map(m => ({ meetingId: m.id, bucketId: m.suggestedBucket || "internal-gen", hours: Math.round((m.durationMins / 60) * 4) / 4, note: m.subject }));

  const meetingDayLabel = new Date(meetingDay + "T12:00:00").toLocaleDateString("en-US", { weekday: "short", month: "long", day: "numeric" });
  const isMeetingToday = meetingDay === today;

  // Outlook category colour presets → CSS colours
  const OUTLOOK_COLOR_MAP = {
    preset0: "#e74856", preset1: "#ff8c00", preset2: "#ffb900", preset3: "#fff100",
    preset4: "#16c60c", preset5: "#00b7c3", preset6: "#0078d7", preset7: "#8764b8",
    preset8: "#b146c2", preset9: "#e81123", preset10: "#a4262c", preset11: "#ca5010",
    preset12: "#8e562e", preset13: "#847545", preset14: "#107c10", preset15: "#004b1c",
    preset16: "#004e8c", preset17: "#001d3d", preset18: "#004e8c", preset19: "#32145a",
    preset20: "#470d21", none: "#888",
  };
  const selectedProjectGids = projects.map(p => p.gid);
  const doneGoals = goals.filter(g => g.done).length;
  const goalPct = goals.length > 0 ? Math.round((doneGoals / goals.length) * 100) : 0;

  // ── Effects ────────────────────────────────────────────────
  useEffect(() => {
    fetch("/auth/status", { credentials: "include" })
      .then(r => r.json())
      .then(async d => {
        setAuthed(d.authenticated);
        if (d.authenticated) {
          try {
            const r = await fetch("/api/data", { credentials: "include" });
            const j = await r.json();
            if (j.success && j.data && Object.keys(j.data).length > 0) {
              const s = j.data;
              if (s.dark !== undefined) setDark(s.dark);
              if (s.goals) setGoals(s.goals);
              if (s.projects) setProjects(s.projects);
              if (s.meetings) setMeetings(s.meetings);
              if (s.meetingMeta) setMeetingMeta(s.meetingMeta);
              if (s.buckets) setBuckets(s.buckets);
              if (s.timeEntries) setTimeEntries(s.timeEntries);
              if (s.asanaMode) setAsanaMode(s.asanaMode);
              if (s.m365Mode) setM365Mode(s.m365Mode);
              if (s.zoomMode) setZoomMode(s.zoomMode);
              if (s.selectedPortfolios) setSelectedPortfolios(s.selectedPortfolios);
              if (s.portfolios) setPortfolios(s.portfolios);
              if (s.portfolios) setPortfolios(s.portfolios);
              if (s.dismissedSuggestions) setDismissedSuggestions(s.dismissedSuggestions);
              if (s.goalsProjects) setGoalsProjects(s.goalsProjects);
              if (s.goalsSections) setGoalsSections(s.goalsSections);
              if (s.outlookCategories) setOutlookCategories(s.outlookCategories);
              if (s.excludedCategories) setExcludedCategories(s.excludedCategories);
              if (s.timezone) setTimezone(s.timezone);
              if (s.directorySearchEnabled !== undefined) setDirectorySearchEnabled(s.directorySearchEnabled);
              if (s.customers) setCustomers(s.customers);
              if (s.customerColumns) setCustomerColumns(s.customerColumns);
              saveLocal(s);
            }
          } catch (e) { console.error("Server load failed, using localStorage:", e); }
          setDataLoaded(true);
          fetchMeetings(today); // always fetch fresh on load
        }
      })
      .catch(() => setAuthed(false));
    fetch("/api/version")
      .then(r => r.json()).then(d => setVersion(d.version || "unknown")).catch(() => setVersion("unknown"));
  }, []);

  useEffect(() => {
    if (!dataLoaded) return;
    const data = { dark, goals, projects, meetings, meetingMeta, buckets, timeEntries, asanaMode, m365Mode, zoomMode, selectedPortfolios, portfolios, dismissedSuggestions, goalsProjects, goalsSections, outlookCategories, excludedCategories, timezone, directorySearchEnabled, customers, customerColumns };
    scheduleSave(data);
  }, [dark, goals, projects, meetings, meetingMeta, buckets, timeEntries, asanaMode, m365Mode, selectedPortfolios, portfolios, dismissedSuggestions, excludedCategories, outlookCategories, timezone, customers, customerColumns]);

  // keep newEntry.date initialized to today on load

  // ── Theme ──────────────────────────────────────────────────
  const bg = dark ? "#0f0f0f" : "#f8f8f6";
  const surface = dark ? "#1a1a1a" : "#ffffff";
  const surface2 = dark ? "#242424" : "#f1efe8";
  const border = dark ? "#2e2e2e" : "#e5e3da";
  const tp = dark ? "#f0f0ee" : "#1a1a1a";
  const ts = dark ? "#888" : "#666";
  const tt = dark ? "#555" : "#aaa";
  const accent = "#1d9e75";

  const card = { background: surface, border: `0.5px solid ${border}`, borderRadius: 12, padding: "1rem 1.25rem", marginBottom: "1rem" };
  const pill = (a) => ({ padding: "6px 14px", borderRadius: 99, fontSize: 13, cursor: "pointer", fontWeight: a ? 500 : 400, background: a ? accent : "transparent", color: a ? "#fff" : ts, border: `0.5px solid ${a ? accent : border}`, transition: "all 0.15s" });
  const btn = (c) => ({ fontSize: 13, padding: "5px 12px", background: c || "transparent", color: c ? "#fff" : ts, border: `0.5px solid ${c || border}`, borderRadius: 6, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 5 });
  const inp = { background: surface2, border: `0.5px solid ${border}`, color: tp, borderRadius: 6, padding: "7px 10px", fontSize: 13, outline: "none", width: "100%", boxSizing: "border-box" };

  const Spinner = () => <i className="ti ti-loader-2" style={{ animation: "spin 1s linear infinite" }} />;
  const Badge = ({ label, color, bg: b2 }) => <span style={{ fontSize: 11, padding: "2px 8px", borderRadius: 99, background: b2, color, fontWeight: 500, whiteSpace: "nowrap" }}>{label}</span>;
  const StatusDot = ({ ok }) => <span style={{ width: 8, height: 8, borderRadius: "50%", background: ok === true ? accent : ok === false ? "#ef4444" : tt, display: "inline-block", marginRight: 6, flexShrink: 0 }} />;
  const ModeToggle = ({ value, onChange, options }) => (
    <div style={{ display: "inline-flex", border: `0.5px solid ${border}`, borderRadius: 6, overflow: "hidden", fontSize: 12 }}>
      {options.map((o, i) => (
        <button key={o.value} onClick={() => onChange(o.value)}
          style={{ padding: "5px 12px", cursor: "pointer", border: "none", borderRight: i < options.length - 1 ? `0.5px solid ${border}` : "none", background: value === o.value ? accent : surface2, color: value === o.value ? "#fff" : ts, fontWeight: value === o.value ? 500 : 400 }}>
          {o.label}
        </button>
      ))}
    </div>
  );

  // Inline due date picker with shortcuts — used in Goals picker and Projects task list
  const DueDatePicker = ({ value, onChange }) => (
    <div style={{ display: "flex", alignItems: "center", gap: 5, flexWrap: "wrap", marginTop: 4 }}>
      <input type="date" value={value || ""} onChange={e => onChange(e.target.value)}
        style={{ ...inp, width: "auto", padding: "2px 6px", fontSize: 11, color: value && value < today ? "#ef4444" : ts }} />
      {[{ label: "Tomorrow", days: 1 }, { label: "+1 wk", days: 7 }, { label: "+1 mo", days: 30 }].map(s => (
        <button key={s.label} onClick={() => onChange(dateOffset(s.days))}
          style={{ fontSize: 11, padding: "2px 7px", background: "transparent", color: ts, border: `0.5px solid ${border}`, borderRadius: 4, cursor: "pointer", whiteSpace: "nowrap" }}>
          {s.label}
        </button>
      ))}
    </div>
  );

  // ── Auth gate ──────────────────────────────────────────────
  if (authed === null) return <div style={{ background: bg, minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center" }}><style>{`body,html{background:${bg};margin:0;padding:0}`}</style><p style={{ color: ts }}>Checking authentication…</p></div>;
  if (authed === false) return (
    <div style={{ background: bg, minHeight: "100vh", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 16 }}>
      <style>{`body,html{background:${bg};margin:0;padding:0}`}</style>
      <h1 style={{ color: tp, fontSize: 22, fontWeight: 500, margin: 0 }}>TAM Dashboard</h1>
      <p style={{ color: ts, margin: 0 }}>Sign in with your Microsoft account to continue.</p>
      <a href="/auth/login" style={{ background: accent, color: "#fff", padding: "10px 24px", borderRadius: 8, textDecoration: "none", fontSize: 15, fontWeight: 500 }}>Sign in with Microsoft</a>
    </div>
  );

  // ── Meetings ───────────────────────────────────────────────
  function guessBucket(attendees) {
    for (const email of (attendees || [])) {
      const domain = email.split("@")[1]?.toLowerCase();
      if (!domain) continue;
      for (const b of buckets) { if (b.domains.some(d => d.toLowerCase() === domain)) return b.id; }
    }
    return "internal-gen";
  }

  function normaliseMeeting(m) {
    let start = m.start || "";
    let end   = m.end   || "";
    const isAllDay = m.isAllDay ||
      (/T00:00:00(\.0+)?Z?$/.test(start) && /T00:00:00(\.0+)?Z?$/.test(end));
    if (isAllDay) {
      start = start.split("T")[0] + "T12:00:00";
      end   = end.split("T")[0]   + "T12:00:00";
    }
    return { ...m, start, end, isAllDay: !!isAllDay,
      durationMins: Math.round((new Date(end) - new Date(start)) / 60000),
      suggestedBucket: guessBucket(m.attendees) };
  }

  async function fetchMeetings(targetDate) {
    const date = targetDate || meetingDay;
    // Set the displayed day to exactly what we're fetching — prevents filter mismatch
    setMeetingDay(date);
    setLoad("meetings", true);
    if (m365Mode === "direct") {
      try {
        const res = await fetch(`/api/calendar/today?date=${date}`, { credentials: "include" });
        const json = await res.json();
        if (json.success && Array.isArray(json.data))
          setMeetings(json.data.map(normaliseMeeting));
        else
          setMeetings([]);
      } catch (e) { console.error(e); setMeetings([]); }
    } else {
      const res = await agentCall(`Get all calendar events for ${date}. Return array with id, subject, start (ISO), end (ISO), isAllDay (bool), categories (string[]), attendees (array of email strings), isOnlineMeeting.`, [{ type: "url", url: M365_MCP, name: "m365" }], m365Mode);
      if (res.success && Array.isArray(res.data))
        setMeetings(res.data.map(normaliseMeeting));
      else
        setMeetings([]);
    }
    setLoad("meetings", false);
  }

  function navigateMeetingDay(direction) {
    const d = new Date(meetingDay + "T12:00:00");
    d.setDate(d.getDate() + direction);
    while (d.getDay() === 0 || d.getDay() === 6) d.setDate(d.getDate() + direction);
    fetchMeetings(dateToISO(d));
  }

  async function fetchCategories() {
    setLoadingCategories(true);
    try {
      const res = await fetch("/api/calendar/categories", { credentials: "include" });
      const json = await res.json();
      if (json.success && Array.isArray(json.data)) setOutlookCategories(json.data);
    } catch (e) { console.error(e); }
    setLoadingCategories(false);
  }

  function toggleExcludedCategory(name) {
    setExcludedCategories(prev =>
      prev.includes(name) ? prev.filter(c => c !== name) : [...prev, name]
    );
  }

  // ── Process panel helpers ──────────────────────────────────
  const STEP_DEFS = [
    { key: "zoom",   label: "Fetch Zoom summary",    icon: "ti-brand-zoom" },
    { key: "notes",  label: "Generate meeting notes", icon: "ti-notes" },
    { key: "blinko", label: "Save to Blinko",         icon: "ti-database" },
    { key: "asana",  label: "Create Asana tasks",     icon: "ti-checkbox" },
  ];

  function initPanel(meeting) {
    const steps = {};
    STEP_DEFS.forEach(s => { steps[s.key] = { status: "idle", message: "Not started", result: null }; });
    setProcessPanel({ meetingId: meeting.id, meeting, steps });
  }

  function setPanelStep(meetingId, key, patch) {
    setProcessPanel(p => p?.meetingId === meetingId
      ? { ...p, steps: { ...p.steps, [key]: { ...p.steps[key], ...patch } } }
      : p
    );
  }

  // ── Zoom Recordings via MCP ────────────────────────────────
  async function fetchRecordings() {
    setLoad("recordings", true);
    setRecordings([]);
    setSelectedRecording(null);
    setRecordingAssets(null);
    try {
      const result = await agentCall(
        `List my Zoom cloud recordings from ${recordingsRange.from} to ${recordingsRange.to}. ` +
        `Use recordings_list with userId="me". ` +
        `Return a JSON array of objects with fields: meetingId (string), topic (string), startTime (ISO string), duration (number, minutes). ` +
        `If there are no recordings return an empty array [].`,
        [{ type: "url", url: ZOOM_RECORDINGS_MCP, name: "zoom-recordings-mcp" }]
      );
      setRecordings(Array.isArray(result.data) ? result.data : []);
    } catch (e) {
      console.error("Recordings fetch error:", e);
      setRecordings([]);
    } finally {
      setLoad("recordings", false);
    }
  }

  async function fetchRecordingAssets(rec) {
    setSelectedRecording(rec);
    setRecordingAssets(null);
    setRecordingsAssetsTab("summary");
    setLoad("recordingAssets", true);
    try {
      const [summaryResult, transcriptResult] = await Promise.all([
        agentCall(
          `Get the AI meeting summary for Zoom meeting ID ${rec.meetingId}. ` +
          `Use get_meeting_assets. Return the summary text as the data field (a plain string).`,
          [{ type: "url", url: ZOOM_RECORDINGS_MCP, name: "zoom-recordings-mcp" }]
        ),
        agentCall(
          `Get the transcript for Zoom meeting ID ${rec.meetingId}. ` +
          `Use get_recording_resource. Return the transcript text as the data field (a plain string), ` +
          `with speaker names on their own lines where available.`,
          [{ type: "url", url: ZOOM_RECORDINGS_MCP, name: "zoom-recordings-mcp" }]
        ),
      ]);
      setRecordingAssets({
        summary: (summaryResult.success && summaryResult.data) ? String(summaryResult.data) : "No summary available for this recording.",
        transcript: (transcriptResult.success && transcriptResult.data) ? String(transcriptResult.data) : "No transcript available for this recording.",
      });
    } catch (e) {
      console.error("Recording assets error:", e);
      setRecordingAssets({ summary: "Failed to load summary.", transcript: "Failed to load transcript." });
    } finally {
      setLoad("recordingAssets", false);
    }
  }

  async function runStepZoom(meeting) {
    // Zoom summaries are now auto-saved to Blinko via webhook when the AI summary is ready.
    // This step lets you optionally link an existing Blinko note or Zoom URL to the meeting,
    // or skip if the webhook will handle it.
    setPanelStep(meeting.id, "zoom", {
      status: "needs_input",
      message: adminStatus?.zoomWebhookSet
        ? "Zoom summaries auto-save via webhook. Paste a Blinko/Zoom URL to link now, or skip."
        : "Paste a Zoom notes URL if available, or skip.",
      candidates: [],
      urlInput: ""
    });
    return "pending";
  }

  // Called when user resolves a Zoom disambiguation
  async function resolveZoomStep(meeting, resolution) {
    const meta = meetingMeta[meeting.id] || {};

    if (resolution.type === "skip") {
      setPanelStep(meeting.id, "zoom", { status: "skipped", message: "Skipped — no Zoom notes.", result: null, candidates: null });
      const notes = await runStepNotes(meeting, null);
      const blinkoUrl = await runStepBlinko(meeting, notes);
      const asanaTaskIds = await runStepAsana(meeting, notes);
      setMeetingMeta(m => ({ ...m, [meeting.id]: { ...meta, processed: true, processedAt: new Date().toISOString(), hasZoomSummary: false, blinkoUrl, notes, asanaTaskIds, followUpsLogged: asanaTaskIds.length > 0 } }));
      return;
    }

    if (resolution.type === "url") {
      const zoomContext = `Zoom notes URL: ${resolution.url} (user-provided; treat as the source for this meeting's notes and decisions)`;
      setPanelStep(meeting.id, "zoom", { status: "done", message: "URL recorded.", result: zoomContext, candidates: null });
      const notes = await runStepNotes(meeting, zoomContext);
      const blinkoUrl = await runStepBlinko(meeting, notes);
      const asanaTaskIds = await runStepAsana(meeting, notes);
      setMeetingMeta(m => ({ ...m, [meeting.id]: { ...meta, processed: true, processedAt: new Date().toISOString(), hasZoomSummary: true, blinkoUrl, notes, asanaTaskIds, followUpsLogged: asanaTaskIds.length > 0 } }));
      return;
    }

    if (resolution.type === "pick") {
      const summary = await fetchZoomSummary(meeting, resolution.candidate);
      const notes = await runStepNotes(meeting, summary);
      const blinkoUrl = await runStepBlinko(meeting, notes);
      const asanaTaskIds = await runStepAsana(meeting, notes);
      setMeetingMeta(m => ({ ...m, [meeting.id]: { ...meta, processed: true, processedAt: new Date().toISOString(), hasZoomSummary: !!summary, blinkoUrl, notes, asanaTaskIds, followUpsLogged: asanaTaskIds.length > 0 } }));
      return;
    }
  }

  async function runStepNotes(meeting, zoomContext) {
    setPanelStep(meeting.id, "notes", { status: "running", message: "Generating meeting notes…" });

    const zoomSection = zoomContext
      ? `Zoom context (reference only — do not attempt to fetch): ${zoomContext}`
      : "Zoom context: Not available";

    const bodySection = meeting.body
      ? `Meeting description/agenda:\n${meeting.body}`
      : "";

    const notesRes = await generateCall(
      `Generate structured meeting notes for the following meeting.

Meeting title: ${meeting.subject}
Date: ${meeting.start}
Attendees: ${(meeting.attendees || []).join(", ") || "Unknown"}
${bodySection}
${zoomSection}

Use the meeting description/agenda and Zoom context above as your primary source.
If a Zoom notes URL was provided, reference it in the summary but do not attempt to fetch it.
If little context is available, base the summary on what can be inferred from the title and attendees.

Return a JSON object in this exact shape:
{
  "summary": "3-5 sentence summary of the meeting",
  "decisions": ["decision 1", "decision 2"],
  "actionItems": [{"title": "task name", "assignee": "person or empty string", "dueDate": "YYYY-MM-DD or empty string"}]
}`
    );

    if (notesRes.success && notesRes.data) {
      let notes = null;
      try { notes = typeof notesRes.data === "string" ? JSON.parse(notesRes.data.replace(/```json|```/g, "").trim()) : notesRes.data; } catch { notes = null; }
      if (notes) {
        const aiCount = (notes.actionItems || []).length;
        const decCount = (notes.decisions || []).length;
        setPanelStep(meeting.id, "notes", { status: "done", message: `Notes generated · ${decCount} decision${decCount !== 1 ? "s" : ""} · ${aiCount} action item${aiCount !== 1 ? "s" : ""}`, result: notes });
        return notes;
      }
    }
    setPanelStep(meeting.id, "notes", { status: "error", message: notesRes.error || "Failed to generate notes." });
    return null;
  }

  async function runStepBlinko(meeting, notes) {
    if (!notes) { setPanelStep(meeting.id, "blinko", { status: "skipped", message: "No notes to save — skipped." }); return null; }
    setPanelStep(meeting.id, "blinko", { status: "running", message: "Saving note to Blinko…" });
    const content = `# ${meeting.subject}\n**Date:** ${meeting.start}\n**Attendees:** ${(meeting.attendees || []).join(", ")}\n\n## Summary\n${notes.summary}\n\n## Decisions\n${(notes.decisions || []).map(d => `- ${d}`).join("\n")}\n\n## Action Items\n${(notes.actionItems || []).map(a => `- [ ] ${a.title}${a.assignee ? ` (${a.assignee})` : ""}${a.dueDate ? ` — due ${a.dueDate}` : ""}`).join("\n")}`;
    try {
      const r = await fetch("/api/blinko/notes", {
        method: "POST", credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ content })
      });
      const data = await r.json();
      if (r.ok && data.success) {
        setPanelStep(meeting.id, "blinko", { status: "done", message: "Note saved to Blinko.", result: data.data });
        return data.data;
      }
      setPanelStep(meeting.id, "blinko", { status: "error", message: data.error || `Blinko returned ${r.status}` });
    } catch (e) {
      setPanelStep(meeting.id, "blinko", { status: "error", message: e.message || "Blinko request failed." });
    }
    return null;
  }

  async function runStepAsana(meeting, notes) {
    if (!notes?.actionItems?.length) { setPanelStep(meeting.id, "asana", { status: "skipped", message: "No action items — skipped." }); return []; }
    setPanelStep(meeting.id, "asana", { status: "running", message: `Creating ${notes.actionItems.length} task${notes.actionItems.length !== 1 ? "s" : ""} in Asana…` });
    const createdGids = [];
    const errors = [];
    for (const item of notes.actionItems) {
      try {
        const r = await fetch("/api/asana/tasks", {
          method: "POST", credentials: "include",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ name: item.title, due_on: item.dueDate || null, notes: item.assignee ? `Assignee: ${item.assignee}` : "" })
        });
        const data = await r.json();
        if (r.ok && data.success) createdGids.push(data.data.gid);
        else errors.push(item.title);
      } catch { errors.push(item.title); }
    }
    if (createdGids.length > 0) {
      const msg = errors.length > 0
        ? `${createdGids.length} task${createdGids.length !== 1 ? "s" : ""} created · ${errors.length} failed`
        : `${createdGids.length} task${createdGids.length !== 1 ? "s" : ""} created in Asana.`;
      setPanelStep(meeting.id, "asana", { status: "done", message: msg, result: createdGids });
      return createdGids;
    }
    setPanelStep(meeting.id, "asana", { status: "error", message: `Failed to create tasks: ${errors.join(", ")}` });
    return [];
  }

  async function processMeeting(meeting) {
    initPanel(meeting);
    const meta = meetingMeta[meeting.id] || {};
    const zoomResult = await runStepZoom(meeting);
    // If Zoom needs disambiguation, the pipeline pauses — resolveZoomStep continues it
    if (zoomResult === "pending") return;
    const notes        = await runStepNotes(meeting, zoomResult);
    const blinkoUrl    = await runStepBlinko(meeting, notes);
    const asanaTaskIds = await runStepAsana(meeting, notes);
    setMeetingMeta(m => ({ ...m, [meeting.id]: { ...meta, processed: true, processedAt: new Date().toISOString(), hasZoomSummary: !!zoomResult, blinkoUrl, notes, asanaTaskIds, followUpsLogged: asanaTaskIds.length > 0 } }));
  }

  async function rerunStep(meeting, key) {
    const meta = meetingMeta[meeting.id] || {};
    const currentSteps = processPanel?.steps || {};
    if (key === "zoom") { await runStepZoom(meeting); return; }
    if (key === "notes") { const z = currentSteps.zoom?.result || null; await runStepNotes(meeting, z); return; }
    if (key === "blinko") { const n = currentSteps.notes?.result || meta.notes || null; await runStepBlinko(meeting, n); return; }
    if (key === "asana") { const n = currentSteps.notes?.result || meta.notes || null; await runStepAsana(meeting, n); return; }
  }

  // ── Goals ──────────────────────────────────────────────────
  async function fetchAsanaTasks() {
    setLoad("tasks", true);
    // Build the project GID filter from goalsProjects toggles.
    // Fall back to ONLY the projects currently loaded (from portfolio filter), never the whole workspace.
    const goalsEnabled = projects.filter(p => goalsProjects[p.gid] === true).map(p => p.gid);
    const filterGids = goalsEnabled.length > 0 ? goalsEnabled : projects.map(p => p.gid);
    if (filterGids.length === 0) { setAsanaTasks([]); setLoad("tasks", false); return; }
    if (asanaMode === "rest") {
      const pp = `&project_gids=${filterGids.join(",")}`;
      const res = await asanaREST(`/api/asana/tasks?due_before=${getWeekEndISO()}${pp}`);
      if (res.success) setAsanaTasks(res.data);
    } else {
      const res = await agentCall(`Get my incomplete Asana tasks due on or before ${getWeekEndISO()} in projects: ${filterGids.join(",")}. Include gid, name, due_on, projects (name, gid), memberships (section gid and name).`, [{ type: "url", url: ASANA_MCP_URL, name: "asana" }]);
      if (res.success && Array.isArray(res.data)) setAsanaTasks(res.data);
    }
    setLoad("tasks", false);
  }

  function dateOffset(days) { const d = new Date(); d.setDate(d.getDate() + days); return dateToISO(d); }

  async function updateGoalDue(goalId, asanaGid, newDate) {
    setGoals(g => g.map(x => x.id === goalId ? { ...x, dueOn: newDate } : x));
    if (asanaGid && asanaMode === "rest") {
      await fetch(`/api/asana/tasks/${asanaGid}/due`, { method: "PATCH", credentials: "include", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ due_on: newDate }) });
    }
  }

  async function updateTaskDue(taskGid, newDate) {
    setTaskDueDates(d => ({ ...d, [taskGid]: newDate }));
    if (asanaMode === "rest") {
      await fetch(`/api/asana/tasks/${taskGid}/due`, {
        method: "PATCH", credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ due_on: newDate })
      });
    }
  }

  async function updateProjectTaskDue(projectGid, taskGid, newDate) {
    setProjTasks(pt => ({
      ...pt,
      [projectGid]: (pt[projectGid] || []).map(t => t.gid === taskGid ? { ...t, due_on: newDate } : t)
    }));
    if (asanaMode === "rest") {
      await fetch(`/api/asana/tasks/${taskGid}/due`, {
        method: "PATCH", credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ due_on: newDate })
      });
    }
  }

  async function toggleGoal(goal) {
    setGoals(g => g.map(x => x.id === goal.id ? { ...x, done: !x.done } : x));
    if (!goal.done && goal.asanaGid) {
      if (asanaMode === "rest") await fetch(`/api/asana/tasks/${goal.asanaGid}/complete`, { method: "PATCH", credentials: "include" });
      else await agentCall(`Mark Asana task ${goal.asanaGid} as completed.`, [{ type: "url", url: ASANA_MCP_URL, name: "asana" }]);
    }
  }

  // ── Projects ───────────────────────────────────────────────
  async function fetchProjects() {
    setLoad("projects", true);
    if (asanaMode === "rest") {
      if (selectedPortfolios.length > 0) {
        const all = [];
        for (const pgid of selectedPortfolios) { const r = await asanaREST(`/api/asana/portfolios/${pgid}/projects`); if (r.success) all.push(...r.data); }
        const seen = new Set();
        setProjects(all.filter(p => { if (seen.has(p.gid)) return false; seen.add(p.gid); return true; }).map(ap => ({ ...ap, status: projects.find(p => p.gid === ap.gid)?.status || "On Track" })));
      } else {
        const r = await asanaREST("/api/asana/projects");
        if (r.success) setProjects(r.data.map(ap => ({ ...ap, status: projects.find(p => p.gid === ap.gid)?.status || "On Track" })));
      }
    } else {
      const r = await agentCall("Get all projects in my Asana workspace. Return array with gid, name.", [{ type: "url", url: ASANA_MCP_URL, name: "asana" }]);
      if (r.success && Array.isArray(r.data)) setProjects(r.data.map(ap => ({ ...ap, status: projects.find(p => p.gid === ap.gid)?.status || "On Track" })));
    }
    setLoad("projects", false);
  }

  async function expandProject(gid) {
    setExpanded(e => ({ ...e, [gid]: !e[gid] }));
    if (!projSections[gid] && !expanded[gid]) {
      setLoad("pt_" + gid, true);
      if (asanaMode === "rest") {
        // Fetch sections then tasks per section
        const sRes = await asanaREST(`/api/asana/projects/${gid}/sections`);
        const taskRes = await asanaREST(`/api/asana/projects/${gid}/tasks`);
        if (taskRes.success) {
          const allTasks = taskRes.data.filter(t => !t.completed);
          setProjTasks(pt => ({ ...pt, [gid]: allTasks }));
          if (sRes.success && sRes.data?.length > 0) {
            // Group tasks by section using memberships
            const sections = sRes.data.map(s => ({
              gid: s.gid, name: s.name,
              tasks: allTasks.filter(t => (t.memberships || []).some(m => m.section?.gid === s.gid))
            }));
            // Tasks with no matching section go into an "Other" bucket
            const assignedGids = new Set(sections.flatMap(s => s.tasks.map(t => t.gid)));
            const unassigned = allTasks.filter(t => !assignedGids.has(t.gid));
            if (unassigned.length > 0) sections.push({ gid: gid + "_other", name: "Other", tasks: unassigned });
            setProjSections(ps => ({ ...ps, [gid]: sections }));
          } else {
            // No sections — treat all tasks as one unlabeled group
            setProjSections(ps => ({ ...ps, [gid]: [{ gid: gid + "_all", name: null, tasks: allTasks }] }));
          }
        }
      } else {
        const r = await agentCall(`Get all incomplete tasks in Asana project ${gid} with their section memberships. Return array with gid, name, completed, assignee (name), due_on, memberships (array of {section: {gid, name}}).`, [{ type: "url", url: ASANA_MCP_URL, name: "asana" }]);
        if (r.success) {
          const allTasks = (r.data || []).filter(t => !t.completed);
          setProjTasks(pt => ({ ...pt, [gid]: allTasks }));
          // Group by section from memberships
          const sectionMap = {};
          allTasks.forEach(t => {
            const sec = t.memberships?.[0]?.section;
            const key = sec?.gid || gid + "_other";
            const name = sec?.name || "Other";
            if (!sectionMap[key]) sectionMap[key] = { gid: key, name, tasks: [] };
            sectionMap[key].tasks.push(t);
          });
          setProjSections(ps => ({ ...ps, [gid]: Object.values(sectionMap) }));
        }
      }
      setLoad("pt_" + gid, false);
    }
  }

  function toggleGoalsProject(pgid) {
    const newVal = !goalsProjects[pgid];
    setGoalsProjects(g => ({ ...g, [pgid]: newVal }));
    // When enabling a project, auto-enable all its known sections
    if (newVal && projSections[pgid]) {
      const updates = {};
      projSections[pgid].forEach(s => { updates[s.gid] = true; });
      setGoalsSections(gs => ({ ...gs, ...updates }));
    }
    // When disabling a project, also disable all its sections
    if (!newVal && projSections[pgid]) {
      const updates = {};
      projSections[pgid].forEach(s => { updates[s.gid] = false; });
      setGoalsSections(gs => ({ ...gs, ...updates }));
    }
  }

  async function completeProjectTask(projectGid, taskGid) {
    setProjTasks(pt => ({ ...pt, [projectGid]: (pt[projectGid] || []).filter(t => t.gid !== taskGid) }));
    if (asanaMode === "rest") await fetch(`/api/asana/tasks/${taskGid}/complete`, { method: "PATCH", credentials: "include" });
    else await agentCall(`Mark Asana task ${taskGid} as completed.`, [{ type: "url", url: ASANA_MCP_URL, name: "asana" }]);
  }

  // ── Time ───────────────────────────────────────────────────
  async function fetchSuggestionsForDay(dateISO) {
    setSuggestionDay(dateISO);
    setLoadingSuggestions(true);
    setSuggestedMeetings([]);
    if (m365Mode === "direct") {
      try {
        const r = await fetch(`/api/calendar/today?date=${dateISO}`, { credentials: "include" });
        const json = await r.json();
        if (json.success && Array.isArray(json.data))
          setSuggestedMeetings(json.data.map(normaliseMeeting));
      } catch (e) { console.error("Suggestions fetch error:", e); }
    } else {
      const res = await agentCall(`Get all calendar events for ${dateISO}. Return array with id, subject, start (ISO), end (ISO), isAllDay (bool), categories (string[]), attendees (array of email strings), isOnlineMeeting.`, [{ type: "url", url: M365_MCP, name: "m365" }], m365Mode);
      if (res.success && Array.isArray(res.data))
        setSuggestedMeetings(res.data.map(normaliseMeeting));
    }
    setLoadingSuggestions(false);
  }

  function navigateSuggestionDay(direction) {
    const d = new Date(suggestionDay + "T12:00:00");
    d.setDate(d.getDate() + direction);
    while (d.getDay() === 0 || d.getDay() === 6) d.setDate(d.getDate() + direction);
    fetchSuggestionsForDay(dateToISO(d));
  }

  function dismissSuggestion(meetingId) {
    setDismissedSuggestions(d => ({ ...d, [meetingId]: true }));
  }

  function saveEditEntry() {
    if (!editingEntry) return;
    setTimeEntries(es => es.map(e => e.id === editingEntry.id ? { ...e, bucketId: editingEntry.bucketId, hours: parseFloat(editingEntry.hours), note: editingEntry.note } : e));
    setEditingEntry(null);
  }

  const EntryEditForm = ({ e, size = "normal" }) => {
    const p = size === "small" ? "3px 6px" : "4px 6px";
    const fs = size === "small" ? 12 : 13;
    return (
      <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
        <select value={editingEntry.bucketId} onChange={ev => setEditingEntry(x => ({ ...x, bucketId: ev.target.value }))} style={{ ...inp, flex: "1 1 120px", padding: p, fontSize: fs }}>
          {["customer", "internal"].map(type => (
            <optgroup key={type} label={type.charAt(0).toUpperCase() + type.slice(1)}>
              {buckets.filter(bk => bk.type === type).map(bk => <option key={bk.id} value={bk.id}>{bk.name}</option>)}
            </optgroup>
          ))}
        </select>
        <input type="number" min="0.25" step="0.25" value={editingEntry.hours} onChange={ev => setEditingEntry(x => ({ ...x, hours: ev.target.value }))} style={{ ...inp, flex: "0 0 60px", padding: p, fontSize: fs }} />
        <input value={editingEntry.note} onChange={ev => setEditingEntry(x => ({ ...x, note: ev.target.value }))} style={{ ...inp, flex: "1 1 120px", padding: p, fontSize: fs }} placeholder="Note" />
        <button style={{ ...btn(accent), fontSize: fs }} onClick={saveEditEntry}>Save</button>
        <button style={{ ...btn(), fontSize: fs }} onClick={() => setEditingEntry(null)}>Cancel</button>
      </div>
    );
  };

  // ── Admin ──────────────────────────────────────────────────
  async function loadAdminStatus() {
    setLoad("admin", true);
    const r = await fetch("/api/admin/status", { credentials: "include" });
    setAdminStatus(await r.json());
    setLoad("admin", false);
  }

  async function testConnection(svc) {
    setTestResults(t => ({ ...t, [svc]: "testing" }));
    const r = await fetch(`/api/admin/test/${svc}`, { credentials: "include" });
    const d = await r.json();
    setTestResults(t => ({ ...t, [svc]: d }));
  }

  async function loadPortfolios() {
    setLoad("portfolios", true);
    const r = await asanaREST("/api/asana/portfolios");
    if (r.success) setPortfolios(r.data);
    setLoad("portfolios", false);
  }


  // ── Customer helpers ─────────────────────────────────────────────
    // Multi-link field — plain function (not a component) to prevent focus loss
    function multiLinkField({ fieldKey, label, icon, placeholder }) {
      const vals = customerDraft.links?.[fieldKey] || [""];
      return (
        <div style={{ marginBottom: 10 }}>
          <FieldLabel icon={icon} color={ts}>{label}</FieldLabel>
          {vals.map((url, i) => (
            <div key={i} style={{ display: "flex", gap: 4, marginBottom: 4 }}>
              <input style={{ ...inp, flex: 1, fontSize: 12 }} placeholder={placeholder} value={url}
                onChange={e => setCustomerDraft(d => {
                  const arr = [...(d.links?.[fieldKey] || [""])];
                  arr[i] = e.target.value;
                  return { ...d, links: { ...d.links, [fieldKey]: arr } };
                })} />
              {vals.length > 1 && (
                <button onClick={() => setCustomerDraft(d => {
                  const arr = (d.links?.[fieldKey] || []).filter((_, j) => j !== i);
                  return { ...d, links: { ...d.links, [fieldKey]: arr.length ? arr : [""] } };
                })} style={{ ...btn(), padding: "4px 6px", color: "#ef4444", borderColor: "#ef444466" }}>
                  <i className="ti ti-x" style={{ fontSize: 11 }} />
                </button>
              )}
            </div>
          ))}
          <button style={{ ...btn(), fontSize: 11, marginTop: 2 }}
            onClick={() => setCustomerDraft(d => ({ ...d, links: { ...d.links, [fieldKey]: [...(d.links?.[fieldKey] || [""]), ""] } }))}>
            <i className="ti ti-plus" style={{ fontSize: 11 }} /> Add another
          </button>
        </div>
      );
    }

    // Team field — plain function (not a component) to prevent focus loss
    function teamField({ roleKey, label }) {
      const val = customerDraft.team?.[roleKey] || "";
      const sugs = teamSuggestions[roleKey] || [];
      const searchError = teamSuggestions[roleKey + "_err"] || null;
      let searchTimer = null;

      function onSearch(q) {
        setCustomerDraft(d => ({ ...d, team: { ...d.team, [roleKey]: q } }));
        clearTimeout(searchTimer);
        if (!directorySearchEnabled || q.length < 2) {
          setTeamSuggestions(s => ({ ...s, [roleKey]: [], [roleKey + "_err"]: null }));
          return;
        }
        searchTimer = setTimeout(async () => {
          try {
            const r = await fetch(`/api/directory/search?q=${encodeURIComponent(q)}`, { credentials: "include" });
            const j = await r.json();
            if (r.status === 403) setTeamSuggestions(s => ({ ...s, [roleKey]: [], [roleKey + "_err"]: "permission" }));
            else if (r.status === 404) setTeamSuggestions(s => ({ ...s, [roleKey]: [], [roleKey + "_err"]: "notfound" }));
            else if (j.success) setTeamSuggestions(s => ({ ...s, [roleKey]: j.data, [roleKey + "_err"]: null }));
          } catch { setTeamSuggestions(s => ({ ...s, [roleKey]: [], [roleKey + "_err"]: null })); }
        }, 300);
      }

      function pick(person) {
        setCustomerDraft(d => ({ ...d, team: { ...d.team, [roleKey]: person.name } }));
        setTeamSuggestions(s => ({ ...s, [roleKey]: [], [roleKey + "_err"]: null }));
      }

      return (
        <div style={{ marginBottom: 10, position: "relative" }}>
          <FieldLabel color={ts}>{label}</FieldLabel>
          <input style={{ ...inp, width: "100%", fontSize: 12 }}
            placeholder={searchError === "permission" ? `Type name (search unavailable)` : `Search ${label} name…`}
            value={val}
            onChange={e => onSearch(e.target.value)}
            onBlur={() => setTimeout(() => setTeamSuggestions(s => ({ ...s, [roleKey]: [] })), 150)}
          />
          {sugs.length > 0 && (
            <div style={{ position: "absolute", top: "100%", left: 0, right: 0, zIndex: 50, background: surface, border: `0.5px solid ${border}`, borderRadius: 6, boxShadow: "0 4px 16px rgba(0,0,0,0.3)", marginTop: 2 }}>
              {sugs.map(p => (
                <div key={p.id} onMouseDown={() => pick(p)}
                  style={{ padding: "7px 10px", cursor: "pointer", borderBottom: `0.5px solid ${border}` }}>
                  <p style={{ fontSize: 12, fontWeight: 500, color: tp, margin: 0 }}>{p.name}</p>
                  <p style={{ fontSize: 11, color: ts, margin: 0 }}>{p.title ? `${p.title} · ` : ""}{p.email}</p>
                </div>
              ))}
            </div>
          )}
        </div>
      );
    }

    // ── Missing fields audit ─────────────────────────────
    // Returns { total, byGroup: { CRM, Team, Financials, Slack, Files, Issues } }
    function auditMissing(c) {
      const lx = c.links ? migrateLinks(c.links) : blankDraft().links;
      const fin = c.financials || {};
      const team = c.team || {};
      const opps = (lx.sf_opps || []).filter(Boolean);
      const groups = {
        CRM:        [["Salesforce Account", !lx.sf_account], ["Opportunity", !opps.length]],
        General:    [["Vertical / Industry", !c.vertical]],
        Team:       [["AE", !team.ae], ["SE", !team.se], ["CSM", !team.csm], ["TAM", !team.tam]],
        Financials: [["ARR", !fin.arr], ["LTV", !fin.ltv], ["Renewal Date", !fin.renewalDate], ["Renewal Amount", !fin.renewalAmount]],
        Slack:      [["Primary Channel", !lx.slack_primary?.url]],
        Files:      [["Google Drive", !lx.gdrive], ["Infra 360", !lx.infra360], ["Account Briefing", !lx.briefing]],
        Issues:     [["Jira (all)", !lx.jira_all], ["Jira (open)", !lx.jira_open]],
        Tools:      [["Asana", !lx.asana]],
      };
      const byGroup = {};
      let total = 0;
      for (const [group, fields] of Object.entries(groups)) {
        const missing = fields.filter(([, isMissing]) => isMissing).map(([name]) => name);
        if (missing.length) { byGroup[group] = missing; total += missing.length; }
      }
      return { total, byGroup };
    }

    async function findMissingData(customer) {
      const audit = auditMissing(customer);
      const lx = migrateLinks(customer.links || {});
      const fin = customer.financials || {};
      const team = customer.team || {};

      const missingList = Object.entries(audit.byGroup)
        .map(([g, fields]) => `${g}: ${fields.join(", ")}`)
        .join("\n");

      const knownInfo = [
        customer.notes && `Notes: ${customer.notes}`,
        lx.sf_account && `Salesforce Account URL: ${lx.sf_account}`,
        team.ae && `AE: ${team.ae}`,
        team.csm && `CSM: ${team.csm}`,
        fin.arr && `ARR: $${Number(fin.arr).toLocaleString()}`,
        fin.renewalDate && `Next Renewal: ${fin.renewalDate}`,
      ].filter(Boolean).join("\n");

      const prompt = `You are a Customer Success assistant. I need help filling in missing data for a customer account.

Customer: ${customer.name}
${knownInfo ? `\nKnown information:\n${knownInfo}` : ""}

Missing fields that need to be found or confirmed:
${missingList}

Please provide:
1. Suggested sources or strategies for finding each missing field
2. Any information you can infer or suggest based on the customer name and known context
3. Questions I should ask the customer or internal team to fill these gaps

Be specific and practical. If you recognize the company, include relevant context.`;

      setMissingDataModal({ customerName: customer.name, loading: true, prompt, result: null, error: null });

      const res = await generateCall(prompt);
      if (res.success && res.data) {
        const text = typeof res.data === "string" ? res.data : JSON.stringify(res.data);
        setMissingDataModal(m => ({ ...m, loading: false, result: text }));
      } else {
        setMissingDataModal(m => ({ ...m, loading: false, error: res.error || "Failed to generate response." }));
      }
    }

    // ── blank draft ──────────────────────────────────────
    function blankDraft(name = "", sfAccountId = "", sfSynced = false) {
      return {
        name, notes: "", vertical: "", sfAccountId, sfSynced,
        financials: { arr: "", renewalDate: "", renewalAmount: "", ltv: "", ela: false },
        team: { ae: "", se: "", csm: "", tam: "" },
        links: {
          sf_account: "", sf_opps: [""],
          slack_primary: { url: "", name: "" },
          slack_supporting: [{ url: "", name: "" }],
          asana: "", gdrive: "",
          infra360: "", briefing: "", briefingDate: "",
          jira_all: "", jira_open: "",
        },
      };
    }

    // Migrate legacy string slack values to object shape
    function migrateLinks(links) {
      if (!links) return blankDraft().links;
      let sl_p = links.slack_primary;
      if (typeof sl_p === "string") sl_p = { url: sl_p, name: "" };
      else sl_p = sl_p || { url: "", name: "" };
      let sl_s = links.slack_supporting;
      if (!Array.isArray(sl_s)) sl_s = [{ url: "", name: "" }];
      else sl_s = sl_s.map(v => typeof v === "string" ? { url: v, name: "" } : (v || { url: "", name: "" }));
      if (sl_s.length === 0) sl_s = [{ url: "", name: "" }];
      return { ...blankDraft().links, ...links, slack_primary: sl_p, slack_supporting: sl_s };
    }

    const filtered = customers
      .filter(c => !customerSearch || c.name.toLowerCase().includes(customerSearch.toLowerCase()))
      .sort((a, b) => a.name.localeCompare(b.name));

    const isNew = editingCustomer === "new";
    const isEditing = editingCustomer !== null;

    function saveCustomer() {
      if (!customerDraft?.name?.trim()) return;
      if (isNew) {
        setCustomers(cs => [...cs, { ...customerDraft, id: Date.now().toString() }]);
      } else {
        setCustomers(cs => cs.map(c => c.id === editingCustomer ? { ...customerDraft, id: c.id } : c));
      }
      setEditingCustomer(null);
      setCustomerDraft(null);
    }

    function deleteCustomer(id) {
      setCustomers(cs => cs.filter(c => c.id !== id));
      if (editingCustomer === id) { setEditingCustomer(null); setCustomerDraft(null); }
    }

    async function syncFromSalesforce() {
      setSfSyncing(true);
      try {
        const res = await fetch("/api/sf/accounts", { credentials: "include" });
        const json = await res.json();
        if (json.success && Array.isArray(json.data)) {
          setCustomers(existing => {
            const updated = [...existing];
            for (const sfAcc of json.data) {
              const idx = updated.findIndex(c => c.sfAccountId === sfAcc.id);
              if (idx >= 0) {
                updated[idx] = { ...updated[idx], name: sfAcc.name, sfSynced: true,
                  links: { ...updated[idx].links, sf_account: sfAcc.url } };
              } else {
                updated.push({ ...blankDraft(sfAcc.name, sfAcc.id, true),
                  id: Date.now().toString() + Math.random(),
                  links: { ...blankDraft().links, sf_account: sfAcc.url } });
              }                  }
            return updated;
          });
        } else {
          alert(json.error || "Salesforce sync failed — see Admin for setup details.");
        }
      } catch (e) {
        alert("Salesforce sync error: " + e.message);
      }
      setSfSyncing(false);
    }

    function LinkChip({ href, icon, label, sub }) {
      if (!href) return null;
      return (
        <a href={href} target="_blank" rel="noreferrer"
          style={{ display: "inline-flex", alignItems: "center", gap: 5, fontSize: 12, color: accent, textDecoration: "none", padding: "3px 8px", border: `0.5px solid ${accent}44`, borderRadius: 5, background: accent+"0d", flexShrink: 0 }}>
          <i className={`ti ${icon}`} style={{ fontSize: 12 }} />
          <span>{label}</span>
          {sub && <span style={{ fontSize: 10, color: ts, marginLeft: 3 }}>{sub}</span>}
        </a>
      );
    }


  // ── Render ─────────────────────────────────────────────────
  return (
    <div style={{ background: bg, minHeight: "100vh", padding: "1.5rem 1rem", fontFamily: "system-ui, sans-serif", color: tp, transition: "background 0.2s" }}>
      <style>{`body,html{background:${bg};margin:0;padding:0} @keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}} input::placeholder{color:${tt}} select option{background:${surface};color:${tp}} a{color:inherit}`}</style>
      <div style={{ maxWidth: (tab === "meetings" && processPanel) || (tab === "customers" && editingCustomer) ? 1200 : 700, margin: "0 auto", transition: "max-width 0.25s" }}>

        {/* Header */}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: "1.5rem" }}>
          <div>
            <p style={{ fontSize: 13, color: ts, margin: "0 0 2px" }}>{new Date().toLocaleDateString("en-US", { weekday: "long", month: "long", day: "numeric" })}</p>
            <h1 style={{ fontSize: 22, fontWeight: 500, margin: "0 0 3px", color: tp }}>{new Date().getHours() < 12 ? "Good morning" : new Date().getHours() < 17 ? "Good afternoon" : "Good evening"} 👋</h1>
            <p style={{ fontSize: 13, color: ts, margin: 0 }}>{doneGoals}/{goals.length} goals · {fmtHours(timeEntries.filter(e => e.date===today).reduce((s,e) => s+(e.hours||0),0))} tracked · {meetings.filter(m => !(m.categories||[]).some(c => excludedCategories.includes(c))).length} meetings</p>
          </div>
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            <a href="/auth/logout" style={{ fontSize: 12, color: tt, textDecoration: "none", padding: "6px 10px", border: `0.5px solid ${border}`, borderRadius: 6 }}>Sign out</a>
            <button onClick={() => setDark(d => !d)} style={{ background: surface2, border: `0.5px solid ${border}`, borderRadius: 8, padding: "8px 10px", cursor: "pointer", color: ts, fontSize: 18 }}>
              <i className={`ti ti-${dark ? "sun" : "moon"}`} />
            </button>
          </div>
        </div>

        {/* Metrics */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 10, marginBottom: "1rem" }}>
          {[{ label: "Meetings today", value: meetings.filter(m => !(m.categories||[]).some(c => excludedCategories.includes(c))).length }, { label: "Processed", value: meetings.filter(m => meetingMeta[m.id]?.processed).length }, { label: "Goals done", value: `${doneGoals}/${goals.length}` }, { label: "Hours tracked", value: fmtHours(timeEntries.filter(e => e.date===today).reduce((s,e) => s+(e.hours||0),0)) }].map(c => (
            <div key={c.label} style={{ background: surface2, borderRadius: 8, padding: "0.75rem 1rem" }}>
              <p style={{ fontSize: 11, color: ts, margin: "0 0 4px" }}>{c.label}</p>
              <p style={{ fontSize: 20, fontWeight: 500, margin: 0, color: tp }}>{c.value}</p>
            </div>
          ))}
        </div>

        {/* Tabs */}
        <div style={{ display: "flex", gap: 8, marginBottom: "1rem", flexWrap: "wrap" }}>
          {TABS.map(t => <button key={t} style={pill(tab === t)} onClick={() => { setTab(t); if (t === "admin") loadAdminStatus(); }}>{TAB_LABELS[t]}</button>)}
        </div>

        {/* ── MEETINGS ── */}
        {tab === "meetings" && (
          <div>
            {/* Sub-tab switcher */}
            <div style={{ display: "flex", gap: 6, marginBottom: 12 }}>
              {[["calendar", "ti-calendar", "Calendar"], ["recordings", "ti-video", "Recordings"]].map(([key, icon, label]) => (
                <button key={key} onClick={() => setMeetingsSubTab(key)}
                  style={{ ...btn(), ...(meetingsSubTab === key ? { background: accent + "22", color: accent, borderColor: accent } : {}) }}>
                  <i className={`ti ${icon}`} style={{ fontSize: 13 }} /> {label}
                </button>
              ))}
            </div>

            {/* ── Calendar sub-tab ── */}
            {meetingsSubTab === "calendar" && (
          <div style={{ display: "flex", gap: 12, alignItems: "flex-start" }}>
            {/* Meeting list */}
            <div style={{ ...card, flex: 1, minWidth: 0, marginBottom: 0 }}>
              {/* Day navigation header */}
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                  <button onClick={() => navigateMeetingDay(-1)}
                    style={{ ...btn(), padding: "5px 8px" }}><i className="ti ti-chevron-left" /></button>
                  <div style={{ position: "relative" }}>
                    <button
                      onClick={() => { try { meetingDateRef.current?.showPicker(); } catch { meetingDateRef.current?.click(); } }}
                      style={{ ...btn(), minWidth: 140, justifyContent: "center", fontWeight: isMeetingToday ? 500 : 400, color: isMeetingToday ? accent : tp }}>
                      <i className="ti ti-calendar" style={{ fontSize: 12 }} />
                      {isMeetingToday ? "Today" : meetingDayLabel}
                    </button>
                    <input
                      ref={meetingDateRef}
                      type="date"
                      value={meetingDay}
                      onChange={e => { if (e.target.value) fetchMeetings(e.target.value); }}
                      style={{ position: "absolute", top: 0, left: 0, width: "100%", height: "100%", opacity: 0, cursor: "pointer", pointerEvents: "none" }}
                    />
                  </div>
                  <button onClick={() => navigateMeetingDay(1)}
                    style={{ ...btn(), padding: "5px 8px" }}><i className="ti ti-chevron-right" /></button>
                  {!isMeetingToday && (
                    <button onClick={() => fetchMeetings(today)}
                      style={{ ...btn(), fontSize: 12, padding: "4px 10px", color: accent, borderColor: accent }}>Today</button>
                  )}
                </div>
                <button style={btn()} onClick={() => fetchMeetings(meetingDay)}>
                  {loading.meetings ? <Spinner /> : <i className="ti ti-refresh" />} Sync
                </button>
              </div>

              {/* Category exclusion indicator */}
              {excludedCategories.length > 0 && (
                <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 10, flexWrap: "wrap" }}>
                  <span style={{ fontSize: 11, color: tt }}>Hiding:</span>
                  {excludedCategories.map(c => {
                    const cat = outlookCategories.find(oc => oc.displayName === c);
                    const color = OUTLOOK_COLOR_MAP[cat?.color] || OUTLOOK_COLOR_MAP.none;
                    return <Badge key={c} label={c} color={color} bg={color + "33"} />;
                  })}
                </div>
              )}

              {loading.meetings && <p style={{ fontSize: 13, color: ts }}>Loading…</p>}
              {!loading.meetings && (() => {
                const visibleMeetings = meetings
                  .filter(m => {
                    // Apply category exclusions — meeting is hidden if ALL its categories are excluded,
                    // or if it has at least one excluded category. A meeting with no categories is never hidden.
                    const cats = m.categories || [];
                    if (cats.length === 0) return true;
                    return !cats.some(c => excludedCategories.includes(c));
                  });
                if (visibleMeetings.length === 0)
                  return <p style={{ fontSize: 13, color: ts }}>No meetings on {isMeetingToday ? "today's" : "this day's"} calendar{excludedCategories.length > 0 ? " (some hidden by category filter)" : ""}.</p>;
                return visibleMeetings.map(m => {
                  const meta = meetingMeta[m.id] || {};
                  const isActive = processPanel?.meetingId === m.id;
                  const isRunning = isActive && Object.values(processPanel.steps).some(s => s.status === "running");
                  const st = m.isAllDay ? "All day" : new Date(m.start).toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" });
                  const et = m.isAllDay ? "" : new Date(m.end).toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" });
                  const timeLabel = m.isAllDay ? "All day" : `${st} – ${et} · ${m.durationMins}m`;
                  // Show category colour dot if meeting has categories
                  const catColor = (() => {
                    const c = (m.categories || [])[0];
                    if (!c) return null;
                    const cat = outlookCategories.find(oc => oc.displayName === c);
                    return OUTLOOK_COLOR_MAP[cat?.color] || OUTLOOK_COLOR_MAP.none;
                  })();
                  return (
                    <div key={m.id} style={{ borderBottom: `0.5px solid ${border}`, margin: "0 -1.25rem", padding: "10px 1.25rem", background: isActive ? accent + "08" : "transparent", transition: "background 0.2s" }}>
                      <div style={{ display: "flex", alignItems: "flex-start", gap: 10 }}>
                        {catColor && <div style={{ width: 3, borderRadius: 99, background: catColor, alignSelf: "stretch", flexShrink: 0, marginTop: 2 }} />}
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <p style={{ fontSize: 14, fontWeight: 500, margin: "0 0 4px", color: tp, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{m.subject}</p>
                          <p style={{ fontSize: 12, color: ts, margin: "0 0 6px" }}>{timeLabel}</p>
                          <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                            {meta.processed ? <Badge label="Processed ✓" color={accent} bg={accent+"22"} /> : <Badge label="Unprocessed" color={tt} bg={surface2} />}
                            {meta.hasZoomSummary && <Badge label="Zoom ✓" color={accent} bg={accent+"22"} />}
                            {meta.blinkoUrl && <Badge label="Blinko ✓" color="#7c3aed" bg="#7c3aed22" />}
                            {meta.followUpsLogged && <Badge label={`${(meta.asanaTaskIds||[]).length} tasks → Asana`} color="#2563eb" bg="#2563eb22" />}
                            {(m.categories || []).map(c => {
                              const cat = outlookCategories.find(oc => oc.displayName === c);
                              const col = OUTLOOK_COLOR_MAP[cat?.color] || OUTLOOK_COLOR_MAP.none;
                              return <Badge key={c} label={c} color={col} bg={col + "33"} />;
                            })}
                          </div>
                        </div>
                        <div style={{ display: "flex", flexDirection: "column", gap: 6, flexShrink: 0 }}>
                          <button style={btn(isActive ? accent : undefined)} onClick={() => { if (!isActive) processMeeting(m); else setProcessPanel(null); }} disabled={isRunning}>
                            {isRunning ? <Spinner /> : <i className={`ti ${isActive ? "ti-x" : "ti-robot"}`} />}
                            {isRunning ? " Running…" : isActive ? " Close" : " Process"}
                          </button>
                          {meta.processed && !isActive && (
                            <button style={btn()} onClick={() => initPanel(m)}>
                              <i className="ti ti-eye" /> Details
                            </button>
                          )}
                        </div>
                      </div>
                    </div>
                  );
                });
              })()}
            </div>

            {/* Process panel — right pane, sticky */}
            {processPanel && (() => {
              const m = processPanel.meeting;
              const steps = processPanel.steps;
              const anyRunning = Object.values(steps).some(s => s.status === "running");
              const anyNeedsInput = Object.values(steps).some(s => s.status === "needs_input");
              return (
                <div style={{ width: 272, flexShrink: 0, background: surface, border: `0.5px solid ${border}`, borderRadius: 12, overflow: "hidden", position: "sticky", top: 20 }}>
                  {/* Header */}
                  <div style={{ padding: "0.85rem 1rem", borderBottom: `0.5px solid ${border}`, background: surface2 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 2 }}>
                      <span style={{ fontSize: 13, fontWeight: 600, color: tp }}>Processing steps</span>
                      <button onClick={() => setProcessPanel(null)} style={{ background: "none", border: "none", cursor: "pointer", color: tt, fontSize: 16, padding: 0, lineHeight: 1 }}><i className="ti ti-x" /></button>
                    </div>
                    <p style={{ fontSize: 11, color: ts, margin: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{m.subject}</p>
                  </div>

                  {/* Steps */}
                  <div style={{ padding: "0.75rem 1rem" }}>
                    {STEP_DEFS.map(def => {
                      const s = steps[def.key];
                      const stepColors  = { idle: tt, running: "#f59e0b", done: accent, error: "#ef4444", skipped: ts, needs_input: "#f59e0b" };
                      const stepIcons   = { idle: "ti-circle", running: "ti-loader-2", done: "ti-circle-check", error: "ti-circle-x", skipped: "ti-circle-dashed", needs_input: "ti-alert-circle" };
                      const color = stepColors[s.status];
                      const icon  = stepIcons[s.status];
                      const spinning = s.status === "running";
                      const needsInput = s.status === "needs_input";
                      return (
                        <div key={def.key} style={{ marginBottom: 14 }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 3 }}>
                            <i className={`ti ${icon}`} style={{ color, fontSize: 14, flexShrink: 0, animation: spinning ? "spin 1s linear infinite" : "none" }} />
                            <span style={{ fontSize: 12, fontWeight: 600, color: tp, flex: 1 }}>{def.label}</span>
                            <button onClick={() => rerunStep(m, def.key)} disabled={spinning || needsInput} title="Re-run this step"
                              style={{ background: "none", border: `0.5px solid ${border}`, borderRadius: 4, cursor: (spinning || needsInput) ? "default" : "pointer", color: (spinning || needsInput) ? tt : ts, fontSize: 10, padding: "1px 5px", opacity: (spinning || needsInput) ? 0.4 : 1 }}>
                              <i className="ti ti-refresh" />
                            </button>
                          </div>
                          <p style={{ fontSize: 11, color, margin: "0 0 0 21px", lineHeight: 1.4 }}>{s.message}</p>

                          {/* Zoom disambiguation UI */}
                          {def.key === "zoom" && needsInput && (() => {
                            const candidates = s.candidates || [];
                            return (
                              <div style={{ marginLeft: 21, marginTop: 6, background: surface2, borderRadius: 6, padding: "8px 10px", fontSize: 11 }}>
                                {candidates.length > 0 && (
                                  <div style={{ marginBottom: 8 }}>
                                    {candidates.map((c, ci) => (
                                      <button key={ci} onClick={() => resolveZoomStep(m, { type: "pick", candidate: c })}
                                        style={{ display: "block", width: "100%", textAlign: "left", background: surface, border: `0.5px solid ${border}`, borderRadius: 5, padding: "5px 8px", marginBottom: 4, cursor: "pointer", color: tp }}>
                                        <span style={{ fontSize: 12, fontWeight: 500, display: "block" }}>{c.topic || c.title}</span>
                                        {(c.start_time || c.date) && <span style={{ fontSize: 10, color: ts }}>{(c.start_time || c.date)?.split("T")[0]}{c.duration ? ` · ${Math.round(c.duration)} min` : ""}</span>}
                                      </button>
                                    ))}
                                  </div>
                                )}
                                <div style={{ display: "flex", gap: 4, marginBottom: 6 }}>
                                  <input
                                    placeholder="Paste Zoom notes URL…"
                                    value={s.urlInput || ""}
                                    onChange={e => setPanelStep(m.id, "zoom", { urlInput: e.target.value })}
                                    style={{ ...inp, flex: 1, fontSize: 11, padding: "4px 7px" }}
                                  />
                                  <button
                                    onClick={() => { if (s.urlInput?.trim()) resolveZoomStep(m, { type: "url", url: s.urlInput.trim() }); }}
                                    disabled={!s.urlInput?.trim()}
                                    style={{ ...btn(accent), fontSize: 11, padding: "4px 8px", opacity: s.urlInput?.trim() ? 1 : 0.4 }}>
                                    Use
                                  </button>
                                </div>
                                <button onClick={() => resolveZoomStep(m, { type: "skip" })}
                                  style={{ ...btn(), fontSize: 10, padding: "2px 8px", width: "100%", justifyContent: "center", color: tt }}>
                                  Skip — no notes exist for this meeting
                                </button>
                              </div>
                            );
                          })()}

                          {/* Result previews */}
                          {s.status === "done" && s.result && def.key === "notes" && (
                            <div style={{ marginLeft: 21, marginTop: 5, background: surface2, borderRadius: 6, padding: "6px 8px", fontSize: 11, color: ts, lineHeight: 1.5 }}>
                              <p style={{ margin: "0 0 3px", fontWeight: 600, color: tp }}>Summary</p>
                              <p style={{ margin: "0 0 5px" }}>{s.result.summary}</p>
                              {(s.result.decisions||[]).length > 0 && <>
                                <p style={{ margin: "0 0 2px", fontWeight: 600, color: tp }}>Decisions</p>
                                {s.result.decisions.map((d,i) => <p key={i} style={{ margin: "0 0 1px" }}>• {d}</p>)}
                              </>}
                              {(s.result.actionItems||[]).length > 0 && <>
                                <p style={{ margin: "5px 0 2px", fontWeight: 600, color: tp }}>Action items</p>
                                {s.result.actionItems.map((a,i) => <p key={i} style={{ margin: "0 0 1px" }}>☐ {a.title}{a.assignee?` · ${a.assignee}`:""}{a.dueDate?` · due ${a.dueDate}`:""}</p>)}
                              </>}
                            </div>
                          )}
                          {s.status === "done" && s.result && def.key === "zoom" && (
                            <div style={{ marginLeft: 21, marginTop: 4, background: surface2, borderRadius: 6, padding: "5px 8px", fontSize: 11, color: ts, lineHeight: 1.4, maxHeight: 80, overflowY: "auto" }}>
                              {(typeof s.result === "string" ? s.result : JSON.stringify(s.result)).slice(0, 280)}…
                            </div>
                          )}
                          {s.status === "done" && def.key === "asana" && Array.isArray(s.result) && s.result.length > 0 && (
                            <p style={{ margin: "3px 0 0 21px", fontSize: 11, color: accent }}>{s.result.length} task{s.result.length !== 1 ? "s" : ""} created</p>
                          )}
                        </div>
                      );
                    })}
                  </div>

                  {/* Footer */}
                  <div style={{ padding: "0.75rem 1rem", borderTop: `0.5px solid ${border}` }}>
                    <button style={{ ...btn(accent), width: "100%", justifyContent: "center" }}
                      onClick={() => processMeeting(m)} disabled={anyRunning || anyNeedsInput}>
                      <i className="ti ti-player-play" /> {anyRunning ? "Running…" : anyNeedsInput ? "Waiting for input…" : "Run all steps"}
                    </button>
                  </div>
                </div>
              );
            })()}
          </div>
            )} {/* end calendar sub-tab */}

            {/* ── Recordings sub-tab ── */}
            {meetingsSubTab === "recordings" && (() => {
              const hasToken = adminStatus?.zoomTokenSet;
              return (
                <div style={{ display: "flex", gap: 12, alignItems: "flex-start" }}>
                  {/* Left: list */}
                  <div style={{ ...card, width: 280, flexShrink: 0, marginBottom: 0 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
                      <span style={{ fontSize: 14, fontWeight: 600, color: tp }}>Cloud Recordings</span>
                      {!hasToken && (
                        <span style={{ fontSize: 11, color: "#ef4444" }}>ZOOM_TOKEN not set</span>
                      )}
                    </div>
                    <div style={{ display: "flex", gap: 6, marginBottom: 10 }}>
                      <div style={{ flex: 1 }}>
                        <div style={{ fontSize: 10, color: ts, marginBottom: 3 }}>FROM</div>
                        <input type="date" value={recordingsRange.from}
                          onChange={e => setRecordingsRange(r => ({ ...r, from: e.target.value }))}
                          style={{ width: "100%", background: surface2, border: `1px solid ${border}`, color: tp, borderRadius: 5, padding: "4px 6px", fontSize: 12 }} />
                      </div>
                      <div style={{ flex: 1 }}>
                        <div style={{ fontSize: 10, color: ts, marginBottom: 3 }}>TO</div>
                        <input type="date" value={recordingsRange.to}
                          onChange={e => setRecordingsRange(r => ({ ...r, to: e.target.value }))}
                          style={{ width: "100%", background: surface2, border: `1px solid ${border}`, color: tp, borderRadius: 5, padding: "4px 6px", fontSize: 12 }} />
                      </div>
                    </div>
                    <button style={{ ...btn(accent), width: "100%", justifyContent: "center", marginBottom: 12 }}
                      onClick={fetchRecordings} disabled={loading.recordings || !hasToken}>
                      {loading.recordings ? <><Spinner /> Fetching…</> : <><i className="ti ti-refresh" /> Fetch Recordings</>}
                    </button>
                    {!hasToken && (
                      <p style={{ fontSize: 11, color: ts, margin: 0 }}>
                        Add <code style={{ background: surface2, padding: "1px 4px", borderRadius: 3 }}>ZOOM_TOKEN</code> to the k8s secret to enable recordings.
                      </p>
                    )}
                    {recordings.length === 0 && !loading.recordings && hasToken && (
                      <p style={{ fontSize: 12, color: ts, textAlign: "center", paddingTop: 20 }}>No recordings loaded. Select a date range and fetch.</p>
                    )}
                    {recordings.map(rec => (
                      <button key={rec.meetingId} onClick={() => fetchRecordingAssets(rec)}
                        style={{ ...btn(), width: "100%", justifyContent: "flex-start", textAlign: "left", gap: 8, marginBottom: 4, padding: "8px 10px",
                          ...(selectedRecording?.meetingId === rec.meetingId ? { background: accent + "22", borderColor: accent, color: accent } : {}) }}>
                        <i className="ti ti-video" style={{ flexShrink: 0, fontSize: 13 }} />
                        <span style={{ minWidth: 0 }}>
                          <div style={{ fontSize: 12, fontWeight: 500, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{rec.topic || "Untitled"}</div>
                          <div style={{ fontSize: 10, color: selectedRecording?.meetingId === rec.meetingId ? accent + "aa" : ts }}>
                            {rec.startTime ? new Date(rec.startTime).toLocaleDateString() : ""}
                            {rec.duration ? ` · ${rec.duration}m` : ""}
                          </div>
                        </span>
                      </button>
                    ))}
                  </div>

                  {/* Right: assets */}
                  {selectedRecording && (
                    <div style={{ ...card, flex: 1, minWidth: 0, marginBottom: 0 }}>
                      <div style={{ marginBottom: 12 }}>
                        <div style={{ fontSize: 15, fontWeight: 600, color: tp, marginBottom: 2 }}>{selectedRecording.topic || "Recording"}</div>
                        <div style={{ fontSize: 11, color: ts }}>
                          {selectedRecording.startTime ? new Date(selectedRecording.startTime).toLocaleString() : ""}
                          {selectedRecording.duration ? ` · ${selectedRecording.duration} min` : ""}
                        </div>
                      </div>
                      <div style={{ display: "flex", gap: 4, borderBottom: `1px solid ${border}`, marginBottom: 12 }}>
                        {["summary", "transcript"].map(t => (
                          <button key={t} onClick={() => setRecordingsAssetsTab(t)}
                            style={{ ...btn(), borderBottom: recordingsAssetsTab === t ? `2px solid ${accent}` : "2px solid transparent",
                              borderRadius: "4px 4px 0 0", color: recordingsAssetsTab === t ? accent : ts, fontSize: 12, padding: "5px 12px" }}>
                            {t === "summary" ? "Summary" : "Transcript"}
                          </button>
                        ))}
                      </div>
                      {loading.recordingAssets ? (
                        <div style={{ display: "flex", justifyContent: "center", padding: 40 }}><Spinner /></div>
                      ) : recordingAssets ? (
                        <pre style={{ whiteSpace: "pre-wrap", fontSize: 12, lineHeight: 1.65, color: tp, margin: 0, fontFamily: "inherit", maxHeight: 480, overflowY: "auto" }}>
                          {recordingsAssetsTab === "summary" ? recordingAssets.summary : recordingAssets.transcript}
                        </pre>
                      ) : (
                        <p style={{ fontSize: 12, color: ts }}>Select a recording to view its summary and transcript.</p>
                      )}
                    </div>
                  )}

                  {!selectedRecording && recordings.length > 0 && (
                    <div style={{ ...card, flex: 1, minWidth: 0, marginBottom: 0, display: "flex", alignItems: "center", justifyContent: "center" }}>
                      <p style={{ fontSize: 12, color: ts }}>Select a recording to view its summary and transcript.</p>
                    </div>
                  )}
                </div>
              );
            })()}
          </div>
        )}

        {/* ── GOALS ── */}
        {tab === "goals" && (
          <div style={card}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
              <div>
                <h2 style={{ fontSize: 16, fontWeight: 500, margin: "0 0 2px", color: tp }}>Today's goals</h2>
                {(() => {
                  const goalsEnabled = projects.filter(p => goalsProjects[p.gid]).length;
                  return goalsEnabled > 0
                    ? <p style={{ fontSize: 11, color: ts, margin: 0 }}>{goalsEnabled} project{goalsEnabled !== 1 ? "s" : ""} in Goals filter · due this week</p>
                    : projects.length > 0
                    ? <p style={{ fontSize: 11, color: ts, margin: 0 }}>All {projects.length} projects · due this week</p>
                    : null;
                })()}
              </div>
              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <span style={{ fontSize: 12, color: goals.length >= 5 ? accent : ts }}>{goals.length}/5 goals</span>
                {goals.length < 5 && (
                  <button style={btn()} onClick={() => { fetchAsanaTasks(); setShowPicker(true); }}>
                    {loading.tasks ? <Spinner /> : <i className="ti ti-cloud-download" />} Load from Asana
                  </button>
                )}
              </div>
            </div>

            {goals.length > 0 && <div style={{ height: 4, background: surface2, borderRadius: 99, marginBottom: 12, overflow: "hidden" }}><div style={{ height: "100%", width: `${goalPct}%`, background: accent, borderRadius: 99, transition: "width 0.3s" }} /></div>}

            {showPicker && goals.length < 5 && (
              <div style={{ background: surface2, border: `0.5px solid ${border}`, borderRadius: 8, padding: "0.75rem", marginBottom: 12 }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
                  <p style={{ fontSize: 13, fontWeight: 500, margin: 0, color: tp }}>Tasks due this week · {5 - goals.length} slot{5 - goals.length !== 1 ? "s" : ""} remaining</p>
                  <button onClick={() => setShowPicker(false)} style={{ background: "none", border: "none", cursor: "pointer", color: ts, fontSize: 16 }}><i className="ti ti-x" /></button>
                </div>
                {loading.tasks && <p style={{ fontSize: 13, color: ts }}>Loading…</p>}
                {!loading.tasks && asanaTasks.length === 0 && <p style={{ fontSize: 13, color: ts }}>No tasks found.</p>}
                {!loading.tasks && (() => {
                  // Group tasks by their matched loaded project
                  const loadedProjGids = new Set(projects.map(p => p.gid));
                  const projGroups = {};
                  asanaTasks.forEach(t => {
                    const matchedProj = (t.projects || []).find(p => loadedProjGids.has(p.gid));
                    const projName = matchedProj?.name || t.projects?.[0]?.name || "No project";
                    const projGid  = matchedProj?.gid  || t.projects?.[0]?.gid  || "_none";
                    if (!projGroups[projGid]) projGroups[projGid] = { name: projName, gid: projGid, tasks: [] };
                    projGroups[projGid].tasks.push(t);
                  });

                  return Object.values(projGroups).map(proj => {
                    const isOpen = !!expandedGroups[proj.gid];

                    // Apply goalsSections filter using task memberships directly
                    const filteredTasks = proj.tasks.filter(t => {
                      const secGid = (t.memberships || []).find(m => m.section?.gid)?.section?.gid;
                      if (!secGid) return true;
                      return goalsSections[secGid] !== false;
                    });

                    // Group filtered tasks by section, using memberships on the task itself.
                    // Preserve Asana section order by collecting sections in the order they appear.
                    const sectionOrder = [];
                    const sectionMap = {};
                    filteredTasks.forEach(t => {
                      // Find the membership for this project specifically
                      const membership = (t.memberships || []).find(m =>
                        m.project?.gid === proj.gid || !m.project // fallback for tasks with single membership
                      );
                      const secGid  = membership?.section?.gid  || null;
                      const secName = membership?.section?.name || null;
                      const key = secGid || "__none__";
                      if (!sectionMap[key]) {
                        sectionMap[key] = { gid: secGid, name: secName, tasks: [] };
                        sectionOrder.push(key);
                      }
                      sectionMap[key].tasks.push(t);
                    });
                    const sectionGroups = sectionOrder.map(k => sectionMap[k]).filter(s => s.tasks.length > 0);

                    const totalShown = filteredTasks.length;
                    const addedCount = filteredTasks.filter(t => goals.find(g => g.asanaGid === t.gid)).length;
                    return (
                      <div key={proj.gid} style={{ marginBottom: 2 }}>
                        {/* Project header */}
                        <button onClick={() => setExpandedGroups(g => ({ ...g, [proj.gid]: !g[proj.gid] }))}
                          style={{ display: "flex", alignItems: "center", gap: 6, width: "100%", background: "none", border: "none", cursor: "pointer", padding: "7px 0", borderBottom: `0.5px solid ${border}` }}>
                          <i className="ti ti-chevron-right" style={{ fontSize: 12, color: tt, transition: "transform 0.15s", transform: isOpen ? "rotate(90deg)" : "none", flexShrink: 0 }} />
                          <span style={{ fontSize: 13, color: tp, fontWeight: 500, flex: 1, textAlign: "left" }}>{proj.name}</span>
                          <span style={{ fontSize: 11, color: tt }}>{totalShown} task{totalShown !== 1 ? "s" : ""}{addedCount > 0 ? ` · ${addedCount} added` : ""}</span>
                        </button>

                        {/* Sections → tasks */}
                        {isOpen && sectionGroups.map(({ gid: secGid, name: secName, tasks }) => (
                          <div key={secGid || "none"}>
                            {secName && (
                              <p style={{ fontSize: 11, fontWeight: 600, color: ts, textTransform: "uppercase", letterSpacing: "0.04em", margin: "8px 0 4px 18px" }}>{secName}</p>
                            )}
                            {tasks.map(t => {
                              const added = goals.find(g => g.asanaGid === t.gid);
                              const due = taskDueDates[t.gid] ?? t.due_on;
                              return (
                                <div key={t.gid} style={{ padding: "6px 0 6px 18px", borderBottom: `0.5px solid ${border}` }}>
                                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                                    <div style={{ flex: 1 }}>
                                      <p style={{ fontSize: 13, margin: 0, color: added ? tt : tp, textDecoration: added ? "line-through" : "none" }}>{t.name}</p>
                                    </div>
                                    <button disabled={!!added} onClick={() => {
                                      if (!added) {
                                        setGoals(g => [...g, { id: Date.now(), text: t.name, done: false, asanaGid: t.gid, project: proj.name, dueOn: taskDueDates[t.gid] ?? t.due_on }]);
                                        if (goals.length + 1 >= 5) setShowPicker(false);
                                      }
                                    }} style={{ fontSize: 12, padding: "4px 10px", background: added ? "transparent" : accent, color: added ? tt : "#fff", border: `0.5px solid ${added ? border : accent}`, borderRadius: 6, cursor: added ? "default" : "pointer", flexShrink: 0 }}>
                                      {added ? "Added" : "Add"}
                                    </button>
                                  </div>
                                  {!added && <DueDatePicker value={due} onChange={d => updateTaskDue(t.gid, d)} />}
                                </div>
                              );
                            })}
                          </div>
                        ))}
                      </div>
                    );
                  });
                })()}
              </div>
            )}

            {goals.length >= 5 && <div style={{ background: accent+"15", border: `0.5px solid ${accent}44`, borderRadius: 8, padding: "8px 12px", marginBottom: 12, fontSize: 13, color: accent }}>✓ 5 goals set for today — great work planning your day!</div>}

            {goals.map(g => (
              <div key={g.id} style={{ padding: "10px 0", borderBottom: `0.5px solid ${border}` }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <input type="checkbox" checked={g.done} onChange={() => toggleGoal(g)} style={{ width: 16, height: 16, accentColor: accent, cursor: "pointer", flexShrink: 0 }} />
                  <div style={{ flex: 1, minWidth: 0 }}>
                    {g.asanaGid ? (
                      <a href={`https://app.asana.com/0/0/${g.asanaGid}/f`} target="_blank" rel="noreferrer"
                        style={{ fontSize: 14, color: g.done ? tt : tp, textDecoration: g.done ? "line-through" : "none", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", display: "block" }}>
                        {g.text}
                      </a>
                    ) : (
                      <p style={{ fontSize: 14, margin: 0, color: g.done ? tt : tp, textDecoration: g.done ? "line-through" : "none", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{g.text}</p>
                    )}
                    {g.project && <p style={{ fontSize: 11, color: tt, margin: "1px 0 0" }}>{g.project}</p>}
                  </div>
                  <button onClick={() => setGoals(gs => gs.filter(x => x.id !== g.id))} style={{ background: "none", border: "none", cursor: "pointer", color: tt, fontSize: 15, padding: 4 }}><i className="ti ti-x" /></button>
                </div>
                {!g.done && (
                  <div style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 6, marginLeft: 26, flexWrap: "wrap" }}>
                    <i className="ti ti-calendar" style={{ fontSize: 12, color: tt }} />
                    <input type="date" value={g.dueOn || ""} onChange={e => updateGoalDue(g.id, g.asanaGid, e.target.value)}
                      style={{ ...inp, width: "auto", padding: "3px 6px", fontSize: 12, color: g.dueOn && g.dueOn < today ? "#ef4444" : ts }} />
                    {[{ label: "Tomorrow", days: 1 }, { label: "Next week", days: 7 }, { label: "Next month", days: 30 }].map(s => (
                      <button key={s.label} onClick={() => updateGoalDue(g.id, g.asanaGid, dateOffset(s.days))}
                        style={{ fontSize: 11, padding: "3px 8px", background: "transparent", color: ts, border: `0.5px solid ${border}`, borderRadius: 4, cursor: "pointer", whiteSpace: "nowrap" }}>
                        {s.label}
                      </button>
                    ))}
                  </div>
                )}
              </div>
            ))}

            <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
              <input style={{ ...inp, flex: 1 }} placeholder="Add a manual goal…" value={newGoalText} onChange={e => setNewGoalText(e.target.value)}
                onKeyDown={e => { if (e.key === "Enter" && newGoalText.trim()) { setGoals(g => [...g, { id: Date.now(), text: newGoalText.trim(), done: false }]); setNewGoalText(""); }}} />
              <button style={btn(accent)} onClick={() => { if (newGoalText.trim()) { setGoals(g => [...g, { id: Date.now(), text: newGoalText.trim(), done: false }]); setNewGoalText(""); }}}>Add</button>
            </div>
          </div>
        )}

        {/* ── PROJECTS ── */}
        {tab === "projects" && (
          <div style={card}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
              <div>
                <h2 style={{ fontSize: 16, fontWeight: 500, margin: "0 0 2px", color: tp }}>Projects</h2>
                {selectedPortfolios.length > 0 && <p style={{ fontSize: 11, color: ts, margin: 0 }}>{selectedPortfolios.length} portfolio{selectedPortfolios.length !== 1 ? "s" : ""} selected</p>}
              </div>
              <button style={btn()} onClick={fetchProjects}>{loading.projects ? <Spinner /> : <i className="ti ti-refresh" />} Sync Asana</button>
            </div>
            {projects.length === 0 && <p style={{ fontSize: 13, color: ts }}>No projects — sync from Asana.</p>}
            {projects.map(p => {
              const isExp = expanded[p.gid];
              const sc = STATUS_C[p.status || "On Track"];
              const sections = projSections[p.gid] || [];
              const goalsOn = !!goalsProjects[p.gid];
              // Count how many sections are enabled for goals
              const enabledSections = sections.filter(s => goalsSections[s.gid] !== false);
              const allTaskCount = sections.reduce((n, s) => n + s.tasks.length, 0);
              return (
                <div key={p.gid} style={{ borderBottom: `0.5px solid ${border}`, paddingBottom: 10, marginBottom: 10 }}>
                  {/* Project row */}
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <button onClick={() => expandProject(p.gid)} style={{ background: "none", border: "none", cursor: "pointer", color: ts, fontSize: 14, padding: 2, flexShrink: 0, transition: "transform 0.15s", transform: isExp ? "rotate(90deg)" : "none" }}><i className="ti ti-chevron-right" /></button>
                    <span style={{ flex: 1, fontSize: 14, fontWeight: 500, color: tp }}>{p.name}</span>
                    {isExp && allTaskCount > 0 && <span style={{ fontSize: 12, color: ts }}>{allTaskCount} open</span>}
                    {/* Goals project toggle */}
                    <button
                      onClick={() => toggleGoalsProject(p.gid)}
                      title={goalsOn ? "Remove from Goals picker" : "Include in Goals picker"}
                      style={{ fontSize: 11, padding: "2px 8px", borderRadius: 99, cursor: "pointer", border: `0.5px solid ${goalsOn ? accent : border}`, background: goalsOn ? accent+"22" : "transparent", color: goalsOn ? accent : tt, whiteSpace: "nowrap" }}>
                      {goalsOn ? `✓ Goals${sections.length > 0 ? ` (${enabledSections.length}/${sections.length})` : ""}` : "Goals"}
                    </button>
                    <select value={p.status || "On Track"} onChange={e => setProjects(ps => ps.map(x => x.gid === p.gid ? { ...x, status: e.target.value } : x))}
                      style={{ fontSize: 12, padding: "2px 6px", background: sc.bg, color: sc.text, border: `0.5px solid ${sc.border}`, borderRadius: 6, cursor: "pointer" }}>
                      {Object.keys(STATUS_C).map(s => <option key={s}>{s}</option>)}
                    </select>
                  </div>

                  {/* Expanded: sections → tasks */}
                  {isExp && (
                    <div style={{ marginLeft: 22, marginTop: 8 }}>
                      {loading["pt_" + p.gid] && <p style={{ fontSize: 12, color: ts }}>Loading…</p>}
                      {!loading["pt_" + p.gid] && sections.length === 0 && <p style={{ fontSize: 12, color: tt }}>No incomplete tasks.</p>}
                      {sections.map(sec => {
                        const secGoalsOn = goalsSections[sec.gid] !== false; // default true if project is on
                        const showSecGoals = goalsOn; // only show section toggle when project is Goals-enabled
                        return (
                          <div key={sec.gid} style={{ marginBottom: 10 }}>
                            {/* Section header — only shown when project has >1 section or section has a name */}
                            {sec.name && (
                              <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "4px 0 4px", borderBottom: `0.5px solid ${border}` }}>
                                <span style={{ fontSize: 11, fontWeight: 600, color: ts, flex: 1, textTransform: "uppercase", letterSpacing: "0.04em" }}>{sec.name}</span>
                                <span style={{ fontSize: 11, color: tt }}>{sec.tasks.length} task{sec.tasks.length !== 1 ? "s" : ""}</span>
                                {showSecGoals && (
                                  <button
                                    onClick={() => setGoalsSections(gs => ({ ...gs, [sec.gid]: !secGoalsOn }))}
                                    title={secGoalsOn ? "Exclude section from Goals" : "Include section in Goals"}
                                    style={{ fontSize: 10, padding: "1px 7px", borderRadius: 99, cursor: "pointer", border: `0.5px solid ${secGoalsOn ? accent : border}`, background: secGoalsOn ? accent+"22" : "transparent", color: secGoalsOn ? accent : tt, whiteSpace: "nowrap" }}>
                                    {secGoalsOn ? "✓" : "+"}
                                  </button>
                                )}
                              </div>
                            )}
                            {/* Tasks */}
                            {sec.tasks.map(t => (
                              <div key={t.gid} style={{ padding: "6px 0", borderBottom: `0.5px solid ${border}` }}>
                                <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                                  <input type="checkbox" checked={false} onChange={() => completeProjectTask(p.gid, t.gid)} style={{ width: 15, height: 15, accentColor: accent, cursor: "pointer", flexShrink: 0 }} />
                                  <span style={{ flex: 1, fontSize: 13, color: tp }}>{t.name}</span>
                                  {t.assignee?.name && <span style={{ fontSize: 11, color: tt }}>{t.assignee.name}</span>}
                                </div>
                                <div style={{ marginLeft: 23 }}>
                                  <DueDatePicker value={t.due_on || ""} onChange={d => updateProjectTaskDue(p.gid, t.gid, d)} />
                                </div>
                              </div>
                            ))}
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        )}

        {/* ── CUSTOMERS ── */}
        {tab === "customers" && (
          <>
          <div style={{ display: "flex", gap: 12, alignItems: "flex-start" }}>
            {/* ── Customer columns ── */}
            <div style={{ flex: 1, minWidth: 0, display: "flex", flexDirection: "column", gap: 12 }}>
              {/* Header — always full width of this column */}
              <div style={{ ...card, marginBottom: 0 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <h2 style={{ fontSize: 16, fontWeight: 500, margin: 0, color: tp }}>
                    Customers <span style={{ fontSize: 13, fontWeight: 400, color: ts }}>({customers.length})</span>
                  </h2>
                  <div style={{ display: "flex", gap: 8 }}>
                    <button style={btn()} onClick={syncFromSalesforce} disabled={sfSyncing}
                      title="Sync accounts from Salesforce where you are the TAM">
                      {sfSyncing ? <Spinner /> : <i className="ti ti-refresh" />} Sync SF
                    </button>
                    <button style={btn(accent)} onClick={() => { setCustomerDraft(blankDraft()); setEditingCustomer("new"); }}>
                      <i className="ti ti-plus" /> Add manually
                    </button>
                  </div>
                </div>
                <input style={{ ...inp, width: "100%", marginTop: 12, fontSize: 13 }}
                  placeholder="Search customers…" value={customerSearch}
                  onChange={e => setCustomerSearch(e.target.value)} />
              </div>

              {/* Column grid — each column is the same width as the header (COL_W px).
                  Uses the standard full-bleed breakout: negative margins to escape the
                  700px container, then re-centers the fixed-width grid on the viewport. */}
              {(() => {
                const namedIndustries = customerColumns.filter(col => col.industry !== "*").map(col => col.industry.toLowerCase());

                function colMatches(col, c) {
                  if (col.industry === "*") return !namedIndustries.some(ind => (c.vertical || "").toLowerCase() === ind);
                  return (c.vertical || "").toLowerCase() === col.industry.toLowerCase();
                }

                const COL_W = 700;
                const GAP   = 16;
                const n = customerColumns.length;
                const totalW = n * COL_W + (n - 1) * GAP;

                return (
                  <div style={{
                    // Full-bleed breakout from the 700px container
                    position: "relative",
                    left: "50%",
                    right: "50%",
                    marginLeft: "-50vw",
                    marginRight: "-50vw",
                    width: "100vw",
                    // Inner grid centered
                    display: "flex",
                    justifyContent: "center",
                  }}>
                    <div style={{
                      display: "grid",
                      gridTemplateColumns: `repeat(${n}, ${COL_W}px)`,
                      gap: GAP,
                      alignItems: "start",
                      width: totalW,
                    }}>
                    {customerColumns.map((col, ci) => {
                      const colCustomers = filtered.filter(c => colMatches(col, c));
                      return (
                        <div key={ci} style={{ ...card, marginBottom: 0 }}>
                          {/* Column header */}
                          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12, paddingBottom: 10, borderBottom: `0.5px solid ${border}` }}>
                            <span style={{ fontSize: 13, fontWeight: 700, color: tp, flex: 1 }}>{col.label}</span>
                            <span style={{ fontSize: 11, color: ts }}>{colCustomers.length}</span>
                          </div>

                          {colCustomers.length === 0 && (
                            <p style={{ fontSize: 12, color: tt, margin: 0 }}>
                              {customers.length === 0 ? "No customers yet." : customerSearch ? "No results." : "No customers in this vertical."}
                            </p>
                          )}

                          {colCustomers.map(c => {
                            const isActive = editingCustomer === c.id;
                            const team = c.team || {};
                            const audit = auditMissing(c);
                            const tooltipLines = audit.total === 0
                              ? "All fields complete ✓"
                              : [`${audit.total} field${audit.total !== 1 ? "s" : ""} missing:`,
                                 ...Object.entries(audit.byGroup).map(([g, fs]) => `  ${g}: ${fs.join(", ")}`)
                                ].join("\n");

                            return (
                              <div key={c.id} style={{ borderBottom: `0.5px solid ${border}`, margin: "0 -1.25rem", padding: "10px 1.25rem", background: isActive ? accent+"08" : "transparent", transition: "background 0.15s" }}>
                                <div style={{ display: "flex", alignItems: "flex-start", gap: 8 }}>
                                  <div style={{ flex: 1, minWidth: 0 }}>
                                    <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 3, flexWrap: "wrap" }}>
                                      <span title={tooltipLines}
                                        style={{ fontSize: 14, fontWeight: 600, color: tp, cursor: "default",
                                          borderBottom: audit.total > 0 ? `1px dashed ${tt}` : "none" }}>
                                        {c.name}
                                      </span>
                                      {audit.total > 0 && (
                                        <span title={tooltipLines}
                                          style={{ fontSize: 10, color: audit.total > 5 ? "#ef4444" : "#f59e0b",
                                            background: audit.total > 5 ? "#ef444422" : "#f59e0b22",
                                            border: `0.5px solid ${audit.total > 5 ? "#ef444466" : "#f59e0b66"}`,
                                            borderRadius: 99, padding: "1px 6px", cursor: "default", flexShrink: 0 }}>
                                          {audit.total}
                                        </span>
                                      )}
                                      {c.sfSynced && <span title="Synced from Salesforce" style={{ fontSize: 11, color: ts }}><i className="ti ti-cloud-check" /></span>}
                                    </div>
                                    {/* Vertical badge */}
                                    {c.vertical && (
                                      <span style={{ fontSize: 10, color: ts, background: surface2, border: `0.5px solid ${border}`, borderRadius: 4, padding: "1px 6px", display: "inline-block", marginBottom: 4 }}>
                                        {c.vertical}
                                      </span>
                                    )}
                                    {/* Team row */}
                                    {(team.ae || team.se || team.csm || team.tam) && (
                                      <p style={{ fontSize: 11, color: ts, margin: "0 0 3px" }}>
                                        {[team.ae && `AE: ${team.ae}`, team.se && `SE: ${team.se}`, team.csm && `CSM: ${team.csm}`, team.tam && `TAM: ${team.tam}`].filter(Boolean).join(" · ")}
                                      </p>
                                    )}
                                    {/* Financials row */}
                                    {(() => {
                                      const fin = c.financials || {};
                                      const fmt$ = v => v ? `$${Number(v).toLocaleString()}` : null;
                                      const parts = [
                                        fin.arr        && `ARR ${fmt$(fin.arr)}`,
                                        fin.ltv        && `LTV ${fmt$(fin.ltv)}`,
                                        fin.renewalDate && `Renews ${new Date(fin.renewalDate + "T12:00:00").toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" })}`,
                                        fin.renewalAmount && `(${fmt$(fin.renewalAmount)})`,
                                        fin.ela        && "ELA",
                                      ].filter(Boolean);
                                      if (!parts.length) return null;
                                      const daysToRenewal = fin.renewalDate
                                        ? Math.ceil((new Date(fin.renewalDate + "T12:00:00") - new Date()) / 86400000)
                                        : null;
                                      const renewalColor = daysToRenewal !== null && daysToRenewal <= 90
                                        ? (daysToRenewal <= 30 ? "#ef4444" : "#f59e0b")
                                        : ts;
                                      return <p style={{ fontSize: 11, margin: "0 0 3px", color: renewalColor }}>{parts.join(" · ")}</p>;
                                    })()}
                                    {c.notes && <p style={{ fontSize: 12, color: ts, margin: "0 0 6px", lineHeight: 1.5 }}>{c.notes}</p>}
                                    {/* Link grid */}
                                    {(() => {
                                      const lx = migrateLinks(c.links);
                                      const slPrimary = lx.slack_primary;
                                      const slSupporting = (lx.slack_supporting || []).filter(s => s.url || s.name);
                                      const opps = (lx.sf_opps || []).filter(Boolean);
                                      const briefingDate = lx.briefingDate
                                        ? new Date(lx.briefingDate+"T12:00:00").toLocaleDateString("en-US",{month:"short",day:"numeric",year:"numeric"})
                                        : null;
                                      function Chip({ href, icon, label, sub, missing }) {
                                        const style = {
                                          display: "inline-flex", alignItems: "center", gap: 4, fontSize: 10,
                                          padding: "2px 6px", borderRadius: 4, flexShrink: 0,
                                          textDecoration: "none", whiteSpace: "nowrap",
                                          ...(missing
                                            ? { color: ts, border: `0.5px solid ${ts}55`, background: "transparent", opacity: 0.65, cursor: "default" }
                                            : { color: accent, border: `0.5px solid ${accent}44`, background: accent+"0d" })
                                        };
                                        const inner = (<><i className={`ti ${icon}`} style={{ fontSize: 10 }} /><span>{label}</span>{sub && <span style={{ fontSize: 9, opacity: 0.7, marginLeft: 1 }}>{sub}</span>}</>);
                                        if (missing) return <span style={style}>{inner}</span>;
                                        return <a href={href} target="_blank" rel="noreferrer" style={style}>{inner}</a>;
                                      }
                                      function LinkRow({ label, children }) {
                                        return (
                                          <div style={{ display: "flex", alignItems: "center", gap: 4, marginBottom: 3 }}>
                                            <span style={{ fontSize: 9, color: tp, textTransform: "uppercase", letterSpacing: "0.06em", fontWeight: 700, width: 44, flexShrink: 0, textAlign: "right" }}>{label}</span>
                                            <div style={{ display: "flex", gap: 3, flexWrap: "wrap" }}>{children}</div>
                                          </div>
                                        );
                                      }
                                      return (
                                        <div style={{ marginTop: 4 }}>
                                          <LinkRow label="CRM">
                                            <Chip href={lx.sf_account} icon="ti-building-store" label="SF" missing={!lx.sf_account} />
                                            {opps.length > 0 ? opps.map((u,i) => <Chip key={i} href={u} icon="ti-currency-dollar" label={opps.length > 1 ? `Opp ${i+1}` : "Opp"} />) : <Chip icon="ti-currency-dollar" label="Opp" missing />}
                                          </LinkRow>
                                          <LinkRow label="Collab">
                                            <Chip href={slPrimary.url} icon="ti-brand-slack" label={slPrimary.name || "Slack"} missing={!slPrimary.url} />
                                            {slSupporting.map((s,i) => <Chip key={i} href={s.url} icon="ti-brand-slack" label={s.name || `+${i+1}`} missing={!s.url} />)}
                                            <Chip href={lx.asana} icon="ti-checkbox" label="Asana" missing={!lx.asana} />
                                          </LinkRow>
                                          <LinkRow label="Files">
                                            <Chip href={lx.gdrive} icon="ti-brand-google-drive" label="Drive" missing={!lx.gdrive} />
                                            <Chip href={lx.infra360} icon="ti-server" label="I360" missing={!lx.infra360} />
                                            <Chip href={lx.briefing} icon="ti-file-description" label="Brief" sub={briefingDate} missing={!lx.briefing} />
                                          </LinkRow>
                                          <LinkRow label="Issues">
                                            <Chip href={lx.jira_all} icon="ti-brand-jira" label="All" missing={!lx.jira_all} />
                                            <Chip href={lx.jira_open} icon="ti-brand-jira" label="Open" missing={!lx.jira_open} />
                                          </LinkRow>
                                        </div>
                                      );
                                    })()}
                                  </div>
                                  <div style={{ display: "flex", flexDirection: "column", gap: 4, flexShrink: 0 }}>
                                    <button style={btn(isActive ? accent : undefined)} onClick={() => {
                                      if (isActive) { setEditingCustomer(null); setCustomerDraft(null); }
                                      else { setEditingCustomer(c.id); setCustomerDraft({
                                        ...blankDraft(), ...c,
                                        financials: { arr: "", renewalDate: "", renewalAmount: "", ltv: "", ela: false, ...c.financials },
                                        team: { ae: "", se: "", csm: "", tam: "", ...c.team },
                                        links: migrateLinks(c.links),
                                      }); }
                                    }}>
                                      <i className={`ti ${isActive ? "ti-x" : "ti-pencil"}`} />
                                    </button>
                                    <button style={{ ...btn(), color: "#ef4444", borderColor: "#ef444466" }}
                                      onClick={() => { if (confirm(`Delete ${c.name}?`)) deleteCustomer(c.id); }}>
                                      <i className="ti ti-trash" />
                                    </button>
                                  </div>
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      );
                    })}
                  </div>
                  </div>
                );
              })()}
            </div>


            {/* ── Edit / Add panel ── */}
            {isEditing && customerDraft && (
              <div style={{ width: customerPanelWidth, flexShrink: 0, display: "flex", alignItems: "stretch", position: "sticky", top: 20, alignSelf: "flex-start" }}>
                {/* Drag-to-resize grip on the left edge */}
                <div
                  onMouseDown={e => {
                    e.preventDefault();
                    const startX = e.clientX;
                    const startW = customerPanelWidth;
                    const onMove = mv => setCustomerPanelWidth(Math.max(380, Math.min(820, startW - (mv.clientX - startX))));
                    const onUp = () => { window.removeEventListener("mousemove", onMove); window.removeEventListener("mouseup", onUp); };
                    window.addEventListener("mousemove", onMove);
                    window.addEventListener("mouseup", onUp);
                  }}
                  style={{ width: 8, flexShrink: 0, cursor: "col-resize", display: "flex", alignItems: "center", justifyContent: "center", borderRadius: "6px 0 0 6px", background: "transparent", transition: "background 0.15s" }}
                  onMouseEnter={e => { e.currentTarget.style.background = accent + "55"; }}
                  onMouseLeave={e => { e.currentTarget.style.background = "transparent"; }}
                  title="Drag to resize panel"
                >
                  <div style={{ width: 2, height: 32, borderRadius: 99, background: accent + "88" }} />
                </div>

                {/* Panel body */}
                <div style={{ flex: 1, minWidth: 0, background: surface, border: `0.5px solid ${border}`, borderRadius: "0 12px 12px 0", overflow: "hidden", display: "flex", flexDirection: "column", boxSizing: "border-box" }}>
                  {/* Header */}
                  <div style={{ padding: "0.85rem 1rem", borderBottom: `0.5px solid ${border}`, background: surface2, display: "flex", justifyContent: "space-between", alignItems: "center", boxSizing: "border-box", gap: 8 }}>
                    <span style={{ fontSize: 13, fontWeight: 600, color: tp, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", flex: 1 }}>{isNew ? "Add customer" : `Edit — ${customerDraft.name || "customer"}`}</span>
                    <div style={{ display: "flex", gap: 6, flexShrink: 0, alignItems: "center" }}>
                      {!isNew && (
                        <button
                          onClick={() => findMissingData(customerDraft)}
                          title="Generate a prompt to find missing data for this customer"
                          style={{ ...btn(), fontSize: 11, padding: "3px 8px", color: "#f59e0b", borderColor: "#f59e0b66" }}>
                          <i className="ti ti-search" /> Find missing
                        </button>
                      )}
                      <button onClick={() => { setEditingCustomer(null); setCustomerDraft(null); setTeamSuggestions({}); }}
                        style={{ background: "none", border: "none", cursor: "pointer", color: tt, fontSize: 16, padding: 0, flexShrink: 0 }}>
                        <i className="ti ti-x" />
                      </button>
                    </div>
                  </div>

                  <div style={{ padding: "0.85rem 1rem", maxHeight: "82vh", overflowY: "auto", overflowX: "hidden", flex: 1, boxSizing: "border-box", width: "100%" }}>

                  {/* ── General ── */}
                  <PanelSection label="General" accent={accent} border={border}>
                    <div style={{ marginBottom: 10 }}>
                      <FieldLabel color={ts}>Customer name</FieldLabel>
                      <input style={{ ...inp, width: "100%" }} placeholder="Acme Corp"
                        value={customerDraft.name}
                        onChange={e => setCustomerDraft(d => ({ ...d, name: e.target.value }))} />
                    </div>
                    <div style={{ marginBottom: 10 }}>
                      <FieldLabel icon="ti-building" color={ts}>Vertical / Industry</FieldLabel>
                      <input style={{ ...inp, width: "100%", fontSize: 12 }}
                        placeholder="e.g. Financial Services, Healthcare, Manufacturing…"
                        list="vertical-suggestions"
                        value={customerDraft.vertical || ""}
                        onChange={e => setCustomerDraft(d => ({ ...d, vertical: e.target.value }))} />
                      <datalist id="vertical-suggestions">
                        {["Financial Services", "Banking", "Insurance", "Healthcare", "Life Sciences",
                          "Manufacturing", "Retail", "Technology", "Telecommunications",
                          "Government", "Energy & Utilities", "Transportation & Logistics",
                          "Media & Entertainment", "Professional Services", "Education"].map(v => (
                          <option key={v} value={v} />
                        ))}
                      </datalist>
                    </div>
                    <div>
                      <FieldLabel color={ts}>Notes</FieldLabel>
                      <textarea style={{ ...inp, width: "100%", resize: "vertical", minHeight: 60, fontSize: 12 }}
                        placeholder="Key context, health notes…"
                        value={customerDraft.notes || ""}
                        onChange={e => setCustomerDraft(d => ({ ...d, notes: e.target.value }))} />
                    </div>
                  </PanelSection>

                  {/* ── Account Information ── */}
                  <PanelSection label="Account Information" accent={accent} border={border}>
                    {/* Row 1: ARR + LTV */}
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 12 }}>
                      <div>
                        <FieldLabel icon="ti-chart-bar" color={ts}>ARR</FieldLabel>
                        <div style={{ position: "relative" }}>
                          <span style={{ position: "absolute", left: 8, top: "50%", transform: "translateY(-50%)", fontSize: 12, color: tt, pointerEvents: "none" }}>$</span>
                          <input style={{ ...inp, width: "100%", fontSize: 12, paddingLeft: 18 }}
                            placeholder="0" type="number" min="0"
                            value={customerDraft.financials?.arr || ""}
                            onChange={e => setCustomerDraft(d => ({ ...d, financials: { ...d.financials, arr: e.target.value } }))} />
                        </div>
                      </div>
                      <div>
                        <FieldLabel icon="ti-trending-up" color={ts}>LTV</FieldLabel>
                        <div style={{ position: "relative" }}>
                          <span style={{ position: "absolute", left: 8, top: "50%", transform: "translateY(-50%)", fontSize: 12, color: tt, pointerEvents: "none" }}>$</span>
                          <input style={{ ...inp, width: "100%", fontSize: 12, paddingLeft: 18 }}
                            placeholder="0" type="number" min="0"
                            value={customerDraft.financials?.ltv || ""}
                            onChange={e => setCustomerDraft(d => ({ ...d, financials: { ...d.financials, ltv: e.target.value } }))} />
                        </div>
                      </div>
                    </div>

                    {/* Row 2: Renewal Date + Renewal Amount */}
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 12 }}>
                      <div>
                        <FieldLabel icon="ti-calendar-due" color={ts}>Next Renewal Date</FieldLabel>
                        <input type="date" style={{ ...inp, width: "100%", fontSize: 12 }}
                          value={customerDraft.financials?.renewalDate || ""}
                          onChange={e => setCustomerDraft(d => ({ ...d, financials: { ...d.financials, renewalDate: e.target.value } }))} />
                      </div>
                      <div>
                        <FieldLabel icon="ti-currency-dollar" color={ts}>Renewal Amount</FieldLabel>
                        <div style={{ position: "relative" }}>
                          <span style={{ position: "absolute", left: 8, top: "50%", transform: "translateY(-50%)", fontSize: 12, color: tt, pointerEvents: "none" }}>$</span>
                          <input style={{ ...inp, width: "100%", fontSize: 12, paddingLeft: 18 }}
                            placeholder="0" type="number" min="0"
                            value={customerDraft.financials?.renewalAmount || ""}
                            onChange={e => setCustomerDraft(d => ({ ...d, financials: { ...d.financials, renewalAmount: e.target.value } }))} />
                        </div>
                      </div>
                    </div>

                    {/* ELA toggle */}
                    <div onClick={() => setCustomerDraft(d => ({ ...d, financials: { ...d.financials, ela: !d.financials?.ela } }))}
                      style={{ display: "flex", alignItems: "center", gap: 10, cursor: "pointer", padding: "6px 0", userSelect: "none" }}>
                      <div style={{ width: 36, height: 20, borderRadius: 99, flexShrink: 0, transition: "background 0.2s",
                        background: customerDraft.financials?.ela ? accent : border, position: "relative" }}>
                        <div style={{ position: "absolute", top: 2, width: 16, height: 16, borderRadius: "50%", background: "#fff",
                          transition: "left 0.2s", left: customerDraft.financials?.ela ? 18 : 2 }} />
                      </div>
                      <div>
                        <p style={{ fontSize: 13, fontWeight: 500, color: tp, margin: 0 }}>Enterprise License Agreement (ELA)</p>
                        <p style={{ fontSize: 11, color: ts, margin: 0 }}>{customerDraft.financials?.ela ? "ELA in place" : "Not an ELA customer"}</p>
                      </div>
                    </div>
                  </PanelSection>
                  {/* ── Account Team ── */}
                  <PanelSection label="Account Team" accent={accent} border={border}>
                    {(() => {
                      if (!directorySearchEnabled) return <p style={{ fontSize: 11, color: tt, margin: "0 0 8px" }}>Directory search is disabled. Enable it in Admin → M365 Configuration if you have <code>User.ReadBasic.All</code> on your Azure app. Type names manually below.</p>;
                      const hasPermErr = Object.entries(teamSuggestions).some(([k,v]) => k.endsWith("_err") && v === "permission");
                      const hasNotFound = Object.entries(teamSuggestions).some(([k,v]) => k.endsWith("_err") && v === "notfound");
                      if (hasNotFound) return <div style={{ background: "#f59e0b22", border: "0.5px solid #f59e0b66", borderRadius: 6, padding: "6px 10px", marginBottom: 8 }}><p style={{ fontSize: 11, color: "#f59e0b", margin: 0 }}>Directory search endpoint not found — redeploy with the latest server.js.</p></div>;
                      if (hasPermErr) return <div style={{ background: "#f59e0b22", border: "0.5px solid #f59e0b66", borderRadius: 6, padding: "6px 10px", marginBottom: 8 }}><p style={{ fontSize: 11, color: "#f59e0b", margin: 0 }}>Employee search returned 403. Add <strong>User.ReadBasic.All</strong> to your Azure app under API Permissions, then re-enable in Admin.</p></div>;
                      return <p style={{ fontSize: 11, color: tt, margin: "0 0 8px" }}>Type a name to search active Camunda employees.</p>;
                    })()}
                    {teamField({ roleKey: "ae", label: "Account Executive (AE)" })}
                    {teamField({ roleKey: "se", label: "Solutions Engineer (SE)" })}
                    {teamField({ roleKey: "csm", label: "Customer Success Manager (CSM)" })}
                    {teamField({ roleKey: "tam", label: "Technical Account Manager (TAM)" })}
                  </PanelSection>

                  {/* ── Salesforce ── */}
                  <PanelSection label="Salesforce" accent={accent} border={border}>
                    <div style={{ marginBottom: 10 }}>
                      <FieldLabel icon="ti-building-store" color={ts}>Account</FieldLabel>
                      <input style={{ ...inp, width: "100%", fontSize: 12 }}
                        placeholder="https://camunda.lightning.force.com/lightning/r/Account/..."
                        value={customerDraft.links?.sf_account || ""}
                        onChange={e => setCustomerDraft(d => ({ ...d, links: { ...d.links, sf_account: e.target.value } }))} />
                    </div>
                    {multiLinkField({ fieldKey: "sf_opps", label: "Opportunity", icon: "ti-currency-dollar", placeholder: "https://camunda.lightning.force.com/lightning/r/Opportunity/..." })}
                  </PanelSection>

                  {/* ── Slack ── */}
                  <PanelSection label="Slack" accent={accent} border={border}>
                    <div style={{ marginBottom: 10 }}>
                      <FieldLabel icon="ti-brand-slack" color={ts}>Primary channel</FieldLabel>
                      <input style={{ ...inp, width: "100%", fontSize: 12, marginBottom: 4 }}
                        placeholder="Channel name (e.g. #acme-corp)"
                        value={customerDraft.links?.slack_primary?.name || ""}
                        onChange={e => setCustomerDraft(d => ({ ...d, links: { ...d.links, slack_primary: { ...(d.links?.slack_primary||{}), name: e.target.value } } }))} />
                      <input style={{ ...inp, width: "100%", fontSize: 12 }}
                        placeholder="https://camunda.slack.com/channels/..."
                        value={customerDraft.links?.slack_primary?.url || ""}
                        onChange={e => setCustomerDraft(d => ({ ...d, links: { ...d.links, slack_primary: { ...(d.links?.slack_primary||{}), url: e.target.value } } }))} />
                    </div>
                    <div>
                      <FieldLabel icon="ti-brand-slack" color={ts}>Supporting channels</FieldLabel>
                      {(customerDraft.links?.slack_supporting || [{ url: "", name: "" }]).map((ch, i) => (
                        <div key={i} style={{ marginBottom: 8, paddingBottom: 8, borderBottom: i < (customerDraft.links?.slack_supporting||[]).length - 1 ? `0.5px solid ${border}` : "none" }}>
                          <div style={{ display: "flex", gap: 4, marginBottom: 4 }}>
                            <input style={{ ...inp, flex: 1, fontSize: 12 }}
                              placeholder="Channel name"
                              value={ch.name || ""}
                              onChange={e => setCustomerDraft(d => {
                                const arr = [...(d.links?.slack_supporting || [])];
                                arr[i] = { ...arr[i], name: e.target.value };
                                return { ...d, links: { ...d.links, slack_supporting: arr } };
                              })} />
                            {(customerDraft.links?.slack_supporting||[]).length > 1 && (
                              <button onClick={() => setCustomerDraft(d => {
                                const arr = (d.links?.slack_supporting||[]).filter((_,j) => j !== i);
                                return { ...d, links: { ...d.links, slack_supporting: arr.length ? arr : [{ url: "", name: "" }] } };
                              })} style={{ ...btn(), padding: "4px 6px", color: "#ef4444", borderColor: "#ef444466" }}>
                                <i className="ti ti-x" style={{ fontSize: 11 }} />
                              </button>
                            )}
                          </div>
                          <input style={{ ...inp, width: "100%", fontSize: 12 }}
                            placeholder="https://camunda.slack.com/channels/..."
                            value={ch.url || ""}
                            onChange={e => setCustomerDraft(d => {
                              const arr = [...(d.links?.slack_supporting || [])];
                              arr[i] = { ...arr[i], url: e.target.value };
                              return { ...d, links: { ...d.links, slack_supporting: arr } };
                            })} />
                        </div>
                      ))}
                      <button style={{ ...btn(), fontSize: 11, marginTop: 2 }}
                        onClick={() => setCustomerDraft(d => ({ ...d, links: { ...d.links, slack_supporting: [...(d.links?.slack_supporting||[]), { url: "", name: "" }] } }))}>
                        <i className="ti ti-plus" style={{ fontSize: 11 }} /> Add channel
                      </button>
                    </div>
                  </PanelSection>

                  {/* ── Project tools ── */}
                  <PanelSection label="Project Tools" accent={accent} border={border}>
                    <div style={{ marginBottom: 10 }}>
                      <FieldLabel icon="ti-checkbox" color={ts}>Asana board</FieldLabel>
                      <input style={{ ...inp, width: "100%", fontSize: 12 }}
                        placeholder="https://app.asana.com/0/..."
                        value={customerDraft.links?.asana || ""}
                        onChange={e => setCustomerDraft(d => ({ ...d, links: { ...d.links, asana: e.target.value } }))} />
                    </div>
                    <div style={{ marginBottom: 10 }}>
                      <FieldLabel icon="ti-brand-jira" color={ts}>Jira — all tickets</FieldLabel>
                      <input style={{ ...inp, width: "100%", fontSize: 12 }}
                        placeholder="https://camunda.atlassian.net/issues/?jql=..."
                        value={customerDraft.links?.jira_all || ""}
                        onChange={e => setCustomerDraft(d => ({ ...d, links: { ...d.links, jira_all: e.target.value } }))} />
                    </div>
                    <div>
                      <FieldLabel icon="ti-brand-jira" color={ts}>Jira — open tickets</FieldLabel>
                      <input style={{ ...inp, width: "100%", fontSize: 12 }}
                        placeholder="https://camunda.atlassian.net/issues/?jql=...+AND+statusCategory+!%3D+Done"
                        value={customerDraft.links?.jira_open || ""}
                        onChange={e => setCustomerDraft(d => ({ ...d, links: { ...d.links, jira_open: e.target.value } }))} />
                    </div>
                  </PanelSection>

                  {/* ── Files & documents ── */}
                  <PanelSection label="Files & Documents" accent={accent} border={border}>
                    <div style={{ marginBottom: 10 }}>
                      <FieldLabel icon="ti-brand-google-drive" color={ts}>Google Drive</FieldLabel>
                      <input style={{ ...inp, width: "100%", fontSize: 12 }}
                        placeholder="https://drive.google.com/drive/folders/..."
                        value={customerDraft.links?.gdrive || ""}
                        onChange={e => setCustomerDraft(d => ({ ...d, links: { ...d.links, gdrive: e.target.value } }))} />
                    </div>
                    <div style={{ marginBottom: 10 }}>
                      <FieldLabel icon="ti-server" color={ts}>Infrastructure 360 Report</FieldLabel>
                      <input style={{ ...inp, width: "100%", fontSize: 12 }} placeholder="https://..."
                        value={customerDraft.links?.infra360 || ""}
                        onChange={e => setCustomerDraft(d => ({ ...d, links: { ...d.links, infra360: e.target.value } }))} />
                    </div>
                    <div>
                      <FieldLabel icon="ti-file-description" color={ts}>Account Briefing</FieldLabel>
                      <input style={{ ...inp, width: "100%", fontSize: 12, marginBottom: 5 }} placeholder="https://..."
                        value={customerDraft.links?.briefing || ""}
                        onChange={e => setCustomerDraft(d => ({ ...d, links: { ...d.links, briefing: e.target.value } }))} />
                      <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                        <span style={{ fontSize: 11, color: tt, flexShrink: 0 }}>Generated:</span>
                        <input type="date" style={{ ...inp, fontSize: 12, padding: "3px 6px", flex: 1 }}
                          value={customerDraft.links?.briefingDate || ""}
                          onChange={e => setCustomerDraft(d => ({ ...d, links: { ...d.links, briefingDate: e.target.value } }))} />
                        <button style={{ ...btn(), fontSize: 11, padding: "3px 8px" }}
                          onClick={() => setCustomerDraft(d => ({ ...d, links: { ...d.links, briefingDate: getTodayISO() } }))}>
                          Today
                        </button>
                      </div>
                    </div>
                  </PanelSection>

                </div>

                {/* Footer */}
                <div style={{ padding: "0.75rem 1rem", borderTop: `0.5px solid ${border}`, display: "flex", gap: 8, boxSizing: "border-box" }}>
                  <button style={{ ...btn(accent), flex: 1, justifyContent: "center" }} onClick={saveCustomer}>
                    <i className="ti ti-device-floppy" /> {isNew ? "Add customer" : "Save changes"}
                  </button>
                </div>
                </div>
              </div>
            )}
          </div>

          {/* ── Missing data modal ── */}
          {missingDataModal && (
            <div style={{ position: "fixed", inset: 0, zIndex: 200, display: "flex", alignItems: "center", justifyContent: "center", background: "rgba(0,0,0,0.6)", padding: 24 }}
              onClick={e => { if (e.target === e.currentTarget) setMissingDataModal(null); }}>
              <div style={{ background: surface, border: `0.5px solid ${border}`, borderRadius: 12, width: "100%", maxWidth: 680, maxHeight: "85vh", display: "flex", flexDirection: "column", overflow: "hidden", boxShadow: "0 8px 40px rgba(0,0,0,0.5)" }}>
                {/* Header */}
                <div style={{ padding: "0.9rem 1.2rem", borderBottom: `0.5px solid ${border}`, background: surface2, display: "flex", alignItems: "center", gap: 10 }}>
                  <i className="ti ti-search" style={{ color: "#f59e0b", fontSize: 16, flexShrink: 0 }} />
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <p style={{ fontSize: 14, fontWeight: 600, color: tp, margin: 0 }}>Missing data — {missingDataModal.customerName}</p>
                    <p style={{ fontSize: 11, color: ts, margin: 0 }}>
                      {missingDataModal.loading ? "Analyzing…" : missingDataModal.result ? "Claude's suggestions — copy the prompt below to run in claude.ai" : "Prompt ready to copy"}
                    </p>
                  </div>
                  <button onClick={() => setMissingDataModal(null)}
                    style={{ background: "none", border: "none", cursor: "pointer", color: tt, fontSize: 18, padding: 0, flexShrink: 0 }}>
                    <i className="ti ti-x" />
                  </button>
                </div>

                {/* Body */}
                <div style={{ overflowY: "auto", flex: 1, display: "flex", flexDirection: "column" }}>
                  {/* Claude's response — shown when available */}
                  {(missingDataModal.loading || missingDataModal.result || missingDataModal.error) && (
                    <div style={{ padding: "1rem 1.2rem", borderBottom: `0.5px solid ${border}` }}>
                      <p style={{ fontSize: 11, fontWeight: 600, color: accent, textTransform: "uppercase", letterSpacing: "0.06em", margin: "0 0 8px" }}>Claude's suggestions</p>
                      {missingDataModal.loading && (
                        <div style={{ display: "flex", alignItems: "center", gap: 10, color: ts }}>
                          <Spinner /> <span style={{ fontSize: 13 }}>Generating suggestions…</span>
                        </div>
                      )}
                      {missingDataModal.error && (
                        <p style={{ fontSize: 13, color: "#ef4444", margin: 0 }}>{missingDataModal.error}</p>
                      )}
                      {missingDataModal.result && (
                        <div style={{ fontSize: 13, color: tp, lineHeight: 1.7, whiteSpace: "pre-wrap" }}>
                          {missingDataModal.result}
                        </div>
                      )}
                    </div>
                  )}

                  {/* The actual prompt — always shown so user can see what gets copied */}
                  <div style={{ padding: "1rem 1.2rem" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                      <p style={{ fontSize: 11, fontWeight: 600, color: accent, textTransform: "uppercase", letterSpacing: "0.06em", margin: 0 }}>Prompt — copy to claude.ai</p>
                      <button style={{ ...btn(), fontSize: 11, padding: "2px 8px" }}
                        onClick={() => { navigator.clipboard?.writeText(missingDataModal.prompt || ""); }}>
                        <i className="ti ti-copy" /> Copy
                      </button>
                    </div>
                    <pre style={{ fontSize: 12, color: ts, lineHeight: 1.6, whiteSpace: "pre-wrap", wordBreak: "break-word", background: surface2, border: `0.5px solid ${border}`, borderRadius: 6, padding: "10px 12px", margin: 0 }}>
                      {missingDataModal.prompt}
                    </pre>
                  </div>
                </div>

                {/* Footer */}
                <div style={{ padding: "0.75rem 1.2rem", borderTop: `0.5px solid ${border}`, display: "flex", justifyContent: "flex-end" }}>
                  <button style={btn(accent)} onClick={() => setMissingDataModal(null)}>
                    Close
                  </button>
                </div>
              </div>
            </div>
          )}
          </>
        )}




        {/* ── TIME ── */}
        {tab === "time" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>

            {/* ── Block 1: Log time manually ── */}
            <div style={card}>
              <h2 style={{ fontSize: 16, fontWeight: 500, margin: "0 0 14px", color: tp }}>Log time</h2>

              {/* Log form */}
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 14 }}>
                <select value={newEntry.bucketId} onChange={e => setNewEntry(n => ({ ...n, bucketId: e.target.value }))} style={{ ...inp, flex: "1 1 150px" }}>
                  <option value="">Select bucket…</option>
                  {["customer","internal"].map(type => (
                    <optgroup key={type} label={type.charAt(0).toUpperCase()+type.slice(1)}>
                      {buckets.filter(b => b.type === type).map(b => <option key={b.id} value={b.id}>{b.name}</option>)}
                    </optgroup>
                  ))}
                </select>
                <input style={{ ...inp, flex: "0 0 72px" }} type="number" min="0.25" step="0.25" placeholder="Hours" value={newEntry.hours} onChange={e => setNewEntry(n => ({ ...n, hours: e.target.value }))} />
                <input style={{ ...inp, flex: "1 1 160px" }} placeholder="Note (optional)" value={newEntry.note} onChange={e => setNewEntry(n => ({ ...n, note: e.target.value }))} />
                <input style={{ ...inp, flex: "0 0 140px" }} type="date" value={newEntry.date} onChange={e => setNewEntry(n => ({ ...n, date: e.target.value }))} />
                <button style={btn(accent)} onClick={() => {
                  if (!newEntry.bucketId || !newEntry.hours) return;
                  setTimeEntries(e => [...e, { id: Date.now(), date: newEntry.date || today, bucketId: newEntry.bucketId, hours: parseFloat(newEntry.hours), note: newEntry.note, source: "manual" }]);
                  setNewEntry(n => ({ ...n, hours: "", note: "" }));
                }}>Log</button>
              </div>

              {/* Entries for selected date */}
              {logEntries.length > 0 && (
                <>
                  <div style={{ borderTop: `0.5px solid ${border}`, paddingTop: 10, marginBottom: 6 }}>
                    {["customer","internal"].map(type => {
                      const tb = buckets.filter(b => b.type === type && (logBucketTotals[b.id]||0) > 0);
                      if (!tb.length) return null;
                      return (
                        <div key={type} style={{ marginBottom: 8 }}>
                          <p style={{ fontSize: 11, color: tt, textTransform: "uppercase", letterSpacing: "0.05em", margin: "0 0 4px", fontWeight: 500 }}>{type}</p>
                          {tb.map(b => (
                            <div key={b.id} style={{ display: "flex", justifyContent: "space-between", padding: "4px 0", borderBottom: `0.5px solid ${border}` }}>
                              <span style={{ fontSize: 13, color: tp }}>{b.name}</span>
                              <span style={{ fontSize: 13, fontWeight: 500, color: accent }}>{fmtHours(logBucketTotals[b.id])}</span>
                            </div>
                          ))}
                        </div>
                      );
                    })}
                    <div style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", borderTop: `0.5px solid ${border}`, marginTop: 4 }}>
                      <span style={{ fontSize: 13, fontWeight: 600, color: tp }}>Total</span>
                      <span style={{ fontSize: 14, fontWeight: 600, color: tp }}>{fmtHours(logTotalHours)}</span>
                    </div>
                  </div>
                  <div>
                    <p style={{ fontSize: 12, color: tt, margin: "0 0 4px" }}>Entries</p>
                    {logEntries.map(e => {
                      const b = buckets.find(x => x.id === e.bucketId);
                      return (
                        <div key={e.id} style={{ padding: "5px 0", borderBottom: `0.5px solid ${border}` }}>
                          {editingEntry?.id === e.id ? <EntryEditForm e={e} /> : (
                            <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13 }}>
                              <span style={{ flex: 1, color: tp }}>{e.note || b?.name || e.bucketId}</span>
                              <span style={{ color: ts, fontSize: 12 }}>{b?.name}</span>
                              <span style={{ fontWeight: 500, color: tp, minWidth: 32, textAlign: "right" }}>{fmtHours(e.hours)}</span>
                              <button onClick={() => setEditingEntry({ id: e.id, bucketId: e.bucketId, hours: e.hours, note: e.note || "" })} style={{ background: "none", border: "none", cursor: "pointer", color: ts, fontSize: 14, padding: 2 }}><i className="ti ti-pencil" /></button>
                              <button onClick={() => setTimeEntries(en => en.filter(x => x.id !== e.id))} style={{ background: "none", border: "none", cursor: "pointer", color: tt, fontSize: 14, padding: 2 }}><i className="ti ti-trash" /></button>
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                </>
              )}
              {logEntries.length === 0 && <p style={{ fontSize: 13, color: tt }}>No time logged for {newEntry.date === today ? "today" : newEntry.date} yet.</p>}
            </div>

            {/* ── Block 2: Meeting suggestions ── */}
            <div style={card}>
              {/* Suggestions header with its own date navigator */}
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
                <div>
                  <h2 style={{ fontSize: 16, fontWeight: 500, margin: "0 0 2px", color: tp }}>Meeting suggestions</h2>
                  <p style={{ fontSize: 11, color: ts, margin: 0 }}>Log time from meetings on any day</p>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                  {!isSuggestionToday && (
                    <button onClick={() => fetchSuggestionsForDay(today)} style={{ ...btn(), fontSize: 12, padding: "4px 10px", color: accent, borderColor: accent }}>Today</button>
                  )}
                  <button onClick={() => navigateSuggestionDay(-1)} style={{ ...btn(), padding: "5px 8px" }}><i className="ti ti-chevron-left" /></button>
                  <div style={{ position: "relative" }}>
                    <button onClick={() => { try { suggestionDateRef.current?.showPicker(); } catch { suggestionDateRef.current?.click(); } }}
                      style={{ ...btn(), minWidth: 110, justifyContent: "center", fontWeight: isSuggestionToday ? 500 : 400, color: isSuggestionToday ? accent : tp }}>
                      <i className="ti ti-calendar" style={{ fontSize: 12 }} />
                      {isSuggestionToday ? "Today" : suggestionDayLabel}
                    </button>
                    <input
                      ref={suggestionDateRef}
                      type="date"
                      value={suggestionDay}
                      onChange={e => { if (e.target.value) fetchSuggestionsForDay(e.target.value); }}
                      style={{ position: "absolute", top: 0, left: 0, width: "100%", height: "100%", opacity: 0, cursor: "pointer", pointerEvents: "none" }}
                    />
                  </div>
                  <button onClick={() => navigateSuggestionDay(1)} style={{ ...btn(), padding: "5px 8px" }}><i className="ti ti-chevron-right" /></button>
                  <button style={{ ...btn(), padding: "5px 8px" }} onClick={() => fetchSuggestionsForDay(suggestionDay)}>
                    {loadingSuggestions ? <Spinner /> : <i className="ti ti-refresh" />}
                  </button>
                </div>
              </div>

              {loadingSuggestions && <p style={{ fontSize: 13, color: ts }}>Loading meetings…</p>}
              {!loadingSuggestions && timeSuggestions.length === 0 && (
                <p style={{ fontSize: 13, color: tt }}>No meetings found for {isSuggestionToday ? "today" : suggestionDayLabel}. Click refresh or navigate to a different day.</p>
              )}
              {timeSuggestions.map(s => {
                const b = buckets.find(x => x.id === s.bucketId);
                return (
                  <div key={s.meetingId} style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 0", borderBottom: `0.5px solid ${border}` }}>
                    <div style={{ flex: 1 }}>
                      <p style={{ fontSize: 13, margin: "0 0 2px", color: tp }}>{s.note}</p>
                      <p style={{ fontSize: 11, color: ts, margin: 0 }}>{fmtHours(s.hours)} → {b?.name || s.bucketId}</p>
                    </div>
                    <button style={btn(accent)} onClick={() => {
                      setTimeEntries(e => [...e, { id: Date.now(), date: suggestionDay, bucketId: s.bucketId, hours: s.hours, note: s.note, source: "suggested" }]);
                      dismissSuggestion(s.meetingId);
                    }}>Log</button>
                    <button style={btn()} onClick={() => dismissSuggestion(s.meetingId)}><i className="ti ti-x" /></button>
                  </div>
                );
              })}
            </div>

            {/* ── Block 3: Weekly summary ── */}
            <div style={card}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
                <h2 style={{ fontSize: 16, fontWeight: 500, margin: 0, color: tp }}>
                  Weekly summary
                  <span style={{ fontSize: 13, fontWeight: 400, color: ts, marginLeft: 10 }}>week of Mon {new Date(weekMonday+"T12:00:00").toLocaleDateString("en-US",{month:"short",day:"numeric"})}</span>
                </h2>
                <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                  <button onClick={() => setWeekOffset(w => w - 1)} style={{ ...btn(), padding: "5px 8px" }}><i className="ti ti-chevron-left" /></button>
                  <span style={{ fontSize: 13, color: tp, minWidth: 80, textAlign: "center" }}>{weekLabel}</span>
                  <button onClick={() => setWeekOffset(w => Math.min(0, w + 1))} disabled={weekOffset === 0} style={{ ...btn(), padding: "5px 8px", opacity: weekOffset === 0 ? 0.3 : 1 }}><i className="ti ti-chevron-right" /></button>
                </div>
              </div>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                  <thead>
                    <tr>
                      <th style={{ textAlign: "left", padding: "0 12px 10px 0", color: ts, fontWeight: 600, fontSize: 12, whiteSpace: "nowrap", minWidth: 130 }}>Category</th>
                      {weekDays.map(d => (
                        <th key={d} style={{ textAlign: "right", padding: "0 0 10px 8px", color: d===today ? accent : ts, fontWeight: d===today ? 700 : 500, fontSize: 12, whiteSpace: "nowrap" }}>
                          {new Date(d+"T12:00:00").toLocaleDateString("en-US",{weekday:"short"})}<br/>
                          <span style={{ fontSize: 11, fontWeight: 400 }}>{new Date(d+"T12:00:00").toLocaleDateString("en-US",{month:"numeric",day:"numeric"})}</span>
                        </th>
                      ))}
                      <th style={{ textAlign: "right", padding: "0 0 10px 8px", color: tp, fontWeight: 700, fontSize: 12 }}>Total</th>
                    </tr>
                  </thead>
                  <tbody>
                    {buckets.map(b => {
                      const dh = weekDays.map(d => timeEntries.filter(e => e.date===d && e.bucketId===b.id).reduce((s,e) => s+(e.hours||0),0));
                      const rt = dh.reduce((s,h) => s+h, 0);
                      if (rt === 0) return null;
                      return (
                        <tr key={b.id} style={{ borderTop: `0.5px solid ${border}` }}>
                          <td style={{ padding: "7px 12px 7px 0", color: tp, whiteSpace: "nowrap" }}>
                            <span style={{ fontSize: 11, padding: "1px 5px", borderRadius: 99, background: b.type==="customer" ? "#2563eb22" : surface2, color: b.type==="customer" ? "#2563eb" : ts, marginRight: 5 }}>{b.type==="customer"?"C":"I"}</span>
                            {b.name}
                          </td>
                          {dh.map((h,i) => <td key={i} style={{ textAlign: "right", padding: "7px 0 7px 8px", color: h>0 ? (weekDays[i]===today ? accent : tp) : tt, fontSize: h>0 ? 13 : 12 }}>{h>0 ? fmtHours(h) : "—"}</td>)}
                          <td style={{ textAlign: "right", padding: "7px 0 7px 8px", fontWeight: 600, color: accent }}>{fmtHours(rt)}</td>
                        </tr>
                      );
                    })}
                    <tr style={{ borderTop: `1.5px solid ${border}` }}>
                      <td style={{ padding: "8px 12px 4px 0", fontWeight: 700, color: tp, fontSize: 13 }}>Total</td>
                      {weekDays.map(d => {
                        const dt = timeEntries.filter(e => e.date===d).reduce((s,e) => s+(e.hours||0), 0);
                        return <td key={d} style={{ textAlign: "right", padding: "8px 0 4px 8px", fontWeight: 600, color: d===today ? accent : tp }}>{dt>0 ? fmtHours(dt) : "—"}</td>;
                      })}
                      <td style={{ textAlign: "right", padding: "8px 0 4px 8px", fontWeight: 700, color: tp }}>
                        {fmtHours(weekDays.reduce((s,d) => s+timeEntries.filter(e => e.date===d).reduce((ss,e) => ss+(e.hours||0),0), 0))}
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>

              {/* Daily entry detail below the table */}
              {weekDays.some(d => timeEntries.some(e => e.date===d)) && (
                <div style={{ marginTop: 18 }}>
                  <p style={{ fontSize: 12, color: tt, margin: "0 0 10px", textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 500 }}>Daily entries</p>
                  {weekDays.map(d => {
                    const entries = timeEntries.filter(e => e.date===d);
                    if (!entries.length) return null;
                    const dt = entries.reduce((s,e) => s+(e.hours||0), 0);
                    const isT = d === today;
                    return (
                      <div key={d} style={{ marginBottom: 12 }}>
                        <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
                          <span style={{ fontSize: 13, fontWeight: 500, color: isT ? accent : tp }}>{new Date(d+"T12:00:00").toLocaleDateString("en-US",{weekday:"long",month:"short",day:"numeric"})}</span>
                          <span style={{ fontSize: 13, fontWeight: 500, color: tp }}>{fmtHours(dt)}</span>
                        </div>
                        {entries.map(e => {
                          const b = buckets.find(x => x.id===e.bucketId);
                          return (
                            <div key={e.id} style={{ padding: "5px 0 5px 10px", borderLeft: `2px solid ${isT ? accent : border}` }}>
                              {editingEntry?.id === e.id ? <EntryEditForm e={e} size="small" /> : (
                                <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12 }}>
                                  <span style={{ flex: 1, color: tp }}>{e.note||b?.name||e.bucketId}</span>
                                  <span style={{ color: ts }}>{b?.name}</span>
                                  <span style={{ fontWeight: 500, color: tp }}>{fmtHours(e.hours)}</span>
                                  <button onClick={() => setEditingEntry({ id: e.id, bucketId: e.bucketId, hours: e.hours, note: e.note||"" })} style={{ background: "none", border: "none", cursor: "pointer", color: ts, fontSize: 13, padding: 2 }}><i className="ti ti-pencil" /></button>
                                  <button onClick={() => setTimeEntries(en => en.filter(x => x.id!==e.id))} style={{ background: "none", border: "none", cursor: "pointer", color: tt, fontSize: 13, padding: 2 }}><i className="ti ti-trash" /></button>
                                </div>
                              )}
                            </div>
                          );
                        })}
                      </div>
                    );
                  })}
                </div>
              )}
            </div>

          </div>
        )}

        {/* ── ADMIN ── */}
        {tab === "admin" && (
          <>
            <div style={card}>
              <h2 style={{ fontSize: 16, fontWeight: 500, margin: "0 0 16px", color: tp }}>M365 Configuration</h2>
              <div style={{ marginBottom: 16 }}>
                <label style={{ fontSize: 13, fontWeight: 500, color: tp, display: "block", marginBottom: 6 }}>Timezone</label>
                <p style={{ fontSize: 11, color: ts, margin: "0 0 8px" }}>Used for calendar event display and Graph API queries. Select your local timezone.</p>
                <select value={timezone} onChange={e => setTimezone(e.target.value)}
                  style={{ ...inp, width: "100%", fontSize: 13 }}>
                  {[
                    ["America/New_York",     "Eastern Time (ET)"],
                    ["America/Chicago",      "Central Time (CT)"],
                    ["America/Denver",       "Mountain Time (MT)"],
                    ["America/Phoenix",      "Arizona (no DST)"],
                    ["America/Los_Angeles",  "Pacific Time (PT)"],
                    ["America/Anchorage",    "Alaska Time (AKT)"],
                    ["Pacific/Honolulu",     "Hawaii Time (HST)"],
                    ["Europe/London",        "London (GMT/BST)"],
                    ["Europe/Paris",         "Central European (CET)"],
                    ["Europe/Berlin",        "Berlin (CET)"],
                    ["Europe/Amsterdam",     "Amsterdam (CET)"],
                    ["Europe/Stockholm",     "Stockholm (CET)"],
                    ["Europe/Zurich",        "Zurich (CET)"],
                    ["Asia/Dubai",           "Dubai (GST)"],
                    ["Asia/Kolkata",         "India (IST)"],
                    ["Asia/Singapore",       "Singapore (SGT)"],
                    ["Asia/Tokyo",           "Japan (JST)"],
                    ["Australia/Sydney",     "Sydney (AEDT)"],
                    ["UTC",                  "UTC"],
                  ].map(([val, label]) => (
                    <option key={val} value={val}>{label} — {val}</option>
                  ))}
                </select>
                {adminStatus?.timezone && adminStatus.timezone !== timezone && (
                  <p style={{ fontSize: 11, color: "#f59e0b", margin: "6px 0 0" }}>
                    Server is using <strong>{adminStatus.timezone}</strong> — save and redeploy to apply the new timezone server-side.
                  </p>
                )}
              </div>

              {/* Directory search toggle */}
              <div style={{ borderTop: `0.5px solid ${border}`, paddingTop: 14 }}>
                <div onClick={() => setDirectorySearchEnabled(v => !v)}
                  style={{ display: "flex", alignItems: "center", gap: 12, cursor: "pointer", userSelect: "none" }}>
                  <div style={{ width: 36, height: 20, borderRadius: 99, flexShrink: 0, transition: "background 0.2s",
                    background: directorySearchEnabled ? accent : border, position: "relative" }}>
                    <div style={{ position: "absolute", top: 2, width: 16, height: 16, borderRadius: "50%", background: "#fff",
                      transition: "left 0.2s", left: directorySearchEnabled ? 18 : 2 }} />
                  </div>
                  <div>
                    <p style={{ fontSize: 13, fontWeight: 500, color: tp, margin: 0 }}>Employee directory search</p>
                    <p style={{ fontSize: 11, color: ts, margin: 0 }}>
                      {directorySearchEnabled
                        ? "Enabled — Account Team fields search Azure AD as you type."
                        : "Disabled — Account Team fields accept free text only."}
                    </p>
                  </div>
                </div>
                {directorySearchEnabled && (
                  <p style={{ fontSize: 11, color: tt, margin: "8px 0 0 48px" }}>
                    Requires <code style={{ background: surface2, padding: "1px 4px", borderRadius: 3 }}>User.ReadBasic.All</code> on your Azure app registration. If you see 403 errors, disable this and add the permission first.
                  </p>
                )}
              </div>
            </div>

            <div style={card}>
              <h2 style={{ fontSize: 16, fontWeight: 500, margin: "0 0 16px", color: tp }}>Connections</h2>
              <div style={{ marginBottom: 20 }}>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <StatusDot ok={adminStatus?.asanaTokenSet} />
                    <span style={{ fontSize: 14, fontWeight: 500, color: tp }}>Asana</span>
                    {testResults.asana && testResults.asana !== "testing" && <Badge label={testResults.asana.ok ? `✓ ${testResults.asana.user}` : `✗ ${testResults.asana.error}`} color={testResults.asana.ok?accent:"#ef4444"} bg={testResults.asana.ok?accent+"22":"#ef444422"} />}
                  </div>
                  <button style={btn()} onClick={() => testConnection("asana")}>{testResults.asana==="testing"?<Spinner />:"Test"}</button>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <span style={{ fontSize: 12, color: ts }}>Auth mode:</span>
                  <ModeToggle value={asanaMode} onChange={setAsanaMode} options={[{value:"rest",label:"Direct REST (PAT)"},{value:"mcp",label:"MCP OAuth"}]} />
                </div>
                <p style={{ fontSize: 11, color: tt, margin: "6px 0 0" }}>{asanaMode==="rest"?"Uses ASANA_TOKEN env var to call Asana REST API directly.":"Uses Asana MCP OAuth server."}</p>
              </div>
              <div style={{ marginBottom: 20, paddingTop: 16, borderTop: `0.5px solid ${border}` }}>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <StatusDot ok={adminStatus?.m365SignedIn} />
                    <span style={{ fontSize: 14, fontWeight: 500, color: tp }}>Microsoft 365</span>
                    {testResults.m365 && testResults.m365 !== "testing" && <Badge label={testResults.m365.ok ? `✓ ${testResults.m365.user}` : `✗ ${testResults.m365.error}`} color={testResults.m365.ok?accent:"#ef4444"} bg={testResults.m365.ok?accent+"22":"#ef444422"} />}
                  </div>
                  <div style={{ display: "flex", gap: 8 }}>
                    {!adminStatus?.m365SignedIn && <a href="/auth/login" style={{ ...btn(accent), textDecoration: "none" }}>Sign in</a>}
                    <button style={btn()} onClick={() => testConnection("m365")}>{testResults.m365==="testing"?<Spinner />:"Test"}</button>
                  </div>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <span style={{ fontSize: 12, color: ts }}>Calendar mode:</span>
                  <ModeToggle value={m365Mode} onChange={setM365Mode} options={[{value:"direct",label:"Direct Graph API"},{value:"claude",label:"Claude MCP"}]} />
                </div>
                <p style={{ fontSize: 11, color: tt, margin: "6px 0 0" }}>{m365Mode==="direct"?"Calls Microsoft Graph API directly using your Azure OAuth token.":"Uses Claude's built-in M365 MCP connector. Only works via Claude.ai."}</p>
              </div>

              {/* M365 Calendar configuration — collapsible */}
              <div style={{ paddingTop: 16, borderTop: `0.5px solid ${border}` }}>
                <button onClick={() => setAdminCollapsed(s => ({ ...s, categories: !s.categories }))}
                  style={{ display: "flex", alignItems: "center", gap: 8, width: "100%", background: "none", border: "none", cursor: "pointer", padding: 0, marginBottom: adminCollapsed.categories ? 0 : 12 }}>
                  <i className="ti ti-chevron-right" style={{ fontSize: 13, color: ts, transition: "transform 0.2s", transform: adminCollapsed.categories ? "none" : "rotate(90deg)", flexShrink: 0 }} />
                  <div style={{ flex: 1, textAlign: "left" }}>
                    <h3 style={{ fontSize: 14, fontWeight: 600, margin: 0, color: tp }}>Calendar category filters</h3>
                    {adminCollapsed.categories && excludedCategories.length > 0 && (
                      <span style={{ fontSize: 11, color: accent }}>{excludedCategories.length} categor{excludedCategories.length !== 1 ? "ies" : "y"} hidden</span>
                    )}
                    {adminCollapsed.categories && excludedCategories.length === 0 && (
                      <span style={{ fontSize: 11, color: tt }}>None hidden</span>
                    )}
                  </div>
                  <button style={{ ...btn(), fontSize: 11, padding: "2px 8px" }} onClick={e => { e.stopPropagation(); fetchCategories(); }}>
                    {loadingCategories ? <Spinner /> : <i className="ti ti-refresh" />}
                  </button>
                </button>
                {!adminCollapsed.categories && (
                  <>
                    <p style={{ fontSize: 11, color: ts, margin: "0 0 8px" }}>Checked categories are hidden from the Meetings tab.</p>
                    {outlookCategories.length === 0 && !loadingCategories && (
                      <p style={{ fontSize: 12, color: tt, margin: "4px 0 0" }}>No categories loaded — click the refresh button above.</p>
                    )}
                    {loadingCategories && <p style={{ fontSize: 12, color: ts }}>Loading…</p>}
                    {outlookCategories.length > 0 && (
                      <div>
                        {outlookCategories.map(cat => {
                          const color = OUTLOOK_COLOR_MAP[cat.color] || OUTLOOK_COLOR_MAP.none;
                          const excluded = excludedCategories.includes(cat.displayName);
                          return (
                            <div key={cat.displayName} onClick={() => toggleExcludedCategory(cat.displayName)}
                              style={{ display: "flex", alignItems: "center", gap: 10, padding: "7px 0", borderBottom: `0.5px solid ${border}`, cursor: "pointer" }}>
                              <div style={{ width: 12, height: 12, borderRadius: "50%", background: color, flexShrink: 0 }} />
                              <span style={{ flex: 1, fontSize: 13, color: excluded ? tt : tp, textDecoration: excluded ? "line-through" : "none" }}>{cat.displayName}</span>
                              <div style={{ width: 16, height: 16, borderRadius: 4, border: `1.5px solid ${excluded ? "#ef4444" : border}`, background: excluded ? "#ef444422" : "transparent", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                                {excluded && <i className="ti ti-x" style={{ fontSize: 10, color: "#ef4444" }} />}
                              </div>
                            </div>
                          );
                        })}
                        {excludedCategories.length > 0 && (
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 8 }}>
                            <span style={{ fontSize: 12, color: tt }}>{excludedCategories.length} categor{excludedCategories.length !== 1 ? "ies" : "y"} hidden</span>
                            <button style={{ ...btn(), fontSize: 11, color: "#ef4444", borderColor: "#ef444466" }}
                              onClick={() => setExcludedCategories([])}>Clear all</button>
                          </div>
                        )}
                      </div>
                    )}
                  </>
                )}
              </div>
              <div style={{ paddingTop: 16, borderTop: `0.5px solid ${border}` }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
                  <StatusDot ok={adminStatus?.zoomWebhookSet} />
                  <span style={{ fontSize: 14, fontWeight: 500, color: tp }}>Zoom</span>
                  {adminStatus?.zoomWebhookSet
                    ? <span style={{ fontSize: 12, color: accent }}>Webhook active — summaries auto-save to Blinko</span>
                    : <span style={{ fontSize: 12, color: tt }}>Webhook not configured</span>
                  }
                </div>
                <p style={{ fontSize: 11, color: tt, margin: "0 0 4px" }}>
                  When a Zoom meeting AI summary is ready, Zoom fires a webhook to this dashboard and the summary is saved automatically as a Blinko note.
                </p>
                <p style={{ fontSize: 11, color: ts, margin: "0 0 10px", fontFamily: "monospace", background: surface2, padding: "4px 8px", borderRadius: 4 }}>
                  Webhook URL: https://dashboard.es-sandbox.com/api/zoom/webhook
                </p>
                <p style={{ fontSize: 11, color: tt, margin: 0 }}>
                  To set up: create a Webhook-only app at marketplace.zoom.us, subscribe to <code style={{ background: surface2, padding: "1px 4px", borderRadius: 3 }}>meeting.summary_completed</code>, set the URL above, and add the Secret Token to the k8s secret as <code style={{ background: surface2, padding: "1px 4px", borderRadius: 3 }}>ZOOM_WEBHOOK_SECRET</code>.
                </p>
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 12 }}>
                  <StatusDot ok={adminStatus?.blinkoTokenSet} />
                  <span style={{ fontSize: 14, fontWeight: 500, color: tp }}>Blinko</span>
                  <span style={{ fontSize: 12, color: tt }}>{adminStatus?.blinkoTokenSet ? "Token configured" : "BLINKO_TOKEN not set — add to k8s secret"}</span>
                </div>
              </div>
            </div>

            <div style={card}>
              <button onClick={() => setAdminCollapsed(s => ({ ...s, portfolios: !s.portfolios }))}
                style={{ display: "flex", alignItems: "center", gap: 8, width: "100%", background: "none", border: "none", cursor: "pointer", padding: 0, marginBottom: adminCollapsed.portfolios ? 0 : 12 }}>
                <i className="ti ti-chevron-right" style={{ fontSize: 13, color: ts, transition: "transform 0.2s", transform: adminCollapsed.portfolios ? "none" : "rotate(90deg)", flexShrink: 0 }} />
                <div style={{ flex: 1, textAlign: "left" }}>
                  <h2 style={{ fontSize: 16, fontWeight: 500, margin: 0, color: tp }}>Portfolio filter</h2>
                  {adminCollapsed.portfolios && (
                    <span style={{ fontSize: 11, color: ts }}>
                      {selectedPortfolios.length > 0 ? `${selectedPortfolios.length} portfolio${selectedPortfolios.length !== 1 ? "s" : ""} selected` : "All projects shown"}
                    </span>
                  )}
                </div>
                <button style={{ ...btn(), fontSize: 11, padding: "2px 8px" }} onClick={e => { e.stopPropagation(); loadPortfolios(); }}>
                  {loading.portfolios ? <Spinner /> : <i className="ti ti-refresh" />}
                </button>
              </button>
              {!adminCollapsed.portfolios && (
                <>
                  <p style={{ fontSize: 12, color: ts, margin: "0 0 10px" }}>Select portfolios to filter projects and goals. Leave all unchecked to show everything.</p>
                  {portfolios.length === 0 && <p style={{ fontSize: 13, color: ts, marginTop: 4 }}>No portfolios loaded — click the refresh button above.</p>}
                  {portfolios.map(p => {
                    const sel = selectedPortfolios.includes(p.gid);
                    return (
                      <div key={p.gid} style={{ display: "flex", alignItems: "center", gap: 10, padding: "7px 0", borderBottom: `0.5px solid ${border}` }}>
                        <input type="checkbox" checked={sel} onChange={() => setSelectedPortfolios(sp => sel ? sp.filter(g => g!==p.gid) : [...sp,p.gid])} style={{ width: 15, height: 15, accentColor: accent, cursor: "pointer", flexShrink: 0 }} />
                        <span style={{ fontSize: 13, color: tp, flex: 1 }}>{p.name}</span>
                        {sel && <Badge label="Active" color={accent} bg={accent+"22"} />}
                      </div>
                    );
                  })}
                  {selectedPortfolios.length > 0 && <p style={{ fontSize: 12, color: ts, margin: "10px 0 0" }}>{selectedPortfolios.length} portfolio{selectedPortfolios.length!==1?"s":""} selected · re-sync Projects tab to apply</p>}
                </>
              )}
            </div>

            <div style={card}>
              <button onClick={() => setAdminCollapsed(s => ({ ...s, buckets: !s.buckets }))}
                style={{ display: "flex", alignItems: "center", gap: 8, width: "100%", background: "none", border: "none", cursor: "pointer", padding: 0, marginBottom: adminCollapsed.buckets ? 0 : 14 }}>
                <i className="ti ti-chevron-right" style={{ fontSize: 13, color: ts, transition: "transform 0.2s", transform: adminCollapsed.buckets ? "none" : "rotate(90deg)", flexShrink: 0 }} />
                <div style={{ flex: 1, textAlign: "left" }}>
                  <h2 style={{ fontSize: 16, fontWeight: 500, margin: 0, color: tp }}>Time buckets</h2>
                  {adminCollapsed.buckets && (
                    <span style={{ fontSize: 11, color: ts }}>{buckets.length} bucket{buckets.length !== 1 ? "s" : ""} configured</span>
                  )}
                </div>
              </button>
              {!adminCollapsed.buckets && (
                <>
                  {["customer", "internal"].map(type => {
                    const group = [...buckets.filter(b => b.type === type)].sort((a, b) => a.name.localeCompare(b.name));
                    if (!group.length) return null;
                    return (
                      <div key={type} style={{ marginBottom: 8 }}>
                        <p style={{ fontSize: 11, fontWeight: 600, color: type === "customer" ? "#2563eb" : ts, textTransform: "uppercase", letterSpacing: "0.05em", margin: "8px 0 4px" }}>{type}</p>
                        {group.map(b => (
                          <div key={b.id} style={{ padding: "8px 0", borderBottom: `0.5px solid ${border}` }}>
                            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
                              <span style={{ flex: 1, fontSize: 13, fontWeight: 500, color: tp }}>{b.name}</span>
                              <button onClick={() => setBuckets(bk => bk.filter(x => x.id!==b.id))} style={{ background: "none", border: "none", cursor: "pointer", color: tt, fontSize: 14 }}><i className="ti ti-trash" /></button>
                            </div>
                            <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                              <span style={{ fontSize: 11, color: tt, flexShrink: 0 }}>Domains:</span>
                              <input style={{ ...inp, fontSize: 12, padding: "4px 8px" }} placeholder="e.g. acme.com" defaultValue={b.domains.join(", ")}
                                onBlur={e => setBuckets(bk => bk.map(x => x.id===b.id ? { ...x, domains: e.target.value.split(",").map(d => d.trim()).filter(Boolean) } : x))} />
                            </div>
                          </div>
                        ))}
                      </div>
                    );
                  })}
                  <div style={{ marginTop: 12, display: "flex", gap: 8, flexWrap: "wrap" }}>
                    <input style={{ ...inp, flex: "2 1 140px" }} placeholder="Bucket name" value={newBucket.name} onChange={e => setNewBucket(n => ({ ...n, name: e.target.value }))} />
                    <select style={{ ...inp, flex: "1 1 100px" }} value={newBucket.type} onChange={e => setNewBucket(n => ({ ...n, type: e.target.value }))}><option value="customer">Customer</option><option value="internal">Internal</option></select>
                    <input style={{ ...inp, flex: "2 1 160px" }} placeholder="Domains (comma-sep)" value={newBucket.domains} onChange={e => setNewBucket(n => ({ ...n, domains: e.target.value }))} />
                    <button style={btn(accent)} onClick={() => { if (!newBucket.name.trim()) return; setBuckets(b => [...b, { id: Date.now().toString(), name: newBucket.name.trim(), type: newBucket.type, domains: newBucket.domains.split(",").map(d => d.trim()).filter(Boolean) }]); setNewBucket({ name: "", type: "customer", domains: "" }); }}>Add</button>
                  </div>
                </>
              )}
            </div>

            {/* ── Customer column configuration ── */}
            <div style={card}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
                <div>
                  <h2 style={{ fontSize: 16, fontWeight: 500, margin: "0 0 2px", color: tp }}>Customer columns</h2>
                  <p style={{ fontSize: 12, color: ts, margin: 0 }}>Split the Customers tab into columns by vertical. Up to 3 columns. "All Others" catches any unmatched verticals.</p>
                </div>
              </div>

              {customerColumns.map((col, ci) => {
                // Collect all verticals from existing customers for the dropdown
                const allVerticals = [...new Set(customers.map(c => c.vertical).filter(Boolean))].sort();
                return (
                  <div key={ci} style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 8, padding: "8px 10px", background: surface2, borderRadius: 8, border: `0.5px solid ${border}` }}>
                    <span style={{ fontSize: 12, color: tt, flexShrink: 0, width: 20, textAlign: "center" }}>{ci + 1}</span>
                    <input style={{ ...inp, flex: "1 1 120px", fontSize: 12 }}
                      placeholder="Column label"
                      value={col.label}
                      onChange={e => setCustomerColumns(cols => cols.map((c, i) => i === ci ? { ...c, label: e.target.value } : c))} />
                    <select style={{ ...inp, flex: "1 1 160px", fontSize: 12 }}
                      value={col.industry}
                      onChange={e => setCustomerColumns(cols => cols.map((c, i) => i === ci ? { ...c, industry: e.target.value } : c))}>
                      <option value="*">All Others</option>
                      {allVerticals.map(v => <option key={v} value={v}>{v}</option>)}
                      {/* Also allow typing a custom vertical not yet in the list */}
                      {col.industry !== "*" && !allVerticals.includes(col.industry) && (
                        <option value={col.industry}>{col.industry} (custom)</option>
                      )}
                    </select>
                    <button onClick={() => setCustomerColumns(cols => cols.filter((_, i) => i !== ci))}
                      disabled={customerColumns.length <= 1}
                      style={{ ...btn(), padding: "4px 8px", color: "#ef4444", borderColor: "#ef444466", opacity: customerColumns.length <= 1 ? 0.4 : 1 }}>
                      <i className="ti ti-trash" style={{ fontSize: 13 }} />
                    </button>
                  </div>
                );
              })}

              {customerColumns.length < 3 && (
                <button style={{ ...btn(), fontSize: 12, marginTop: 4 }}
                  onClick={() => setCustomerColumns(cols => [...cols, { label: "New Column", industry: "*" }])}>
                  <i className="ti ti-plus" /> Add column
                </button>
              )}
              <p style={{ fontSize: 11, color: tt, margin: "10px 0 0" }}>
                Set a vertical in each column's dropdown to show only customers with that vertical. Set to "All Others" to show customers that don't match any named column. Verticals in the dropdown come from your customers' Vertical / Industry fields.
              </p>
            </div>

            <div style={card}>
              <h2 style={{ fontSize: 16, fontWeight: 500, margin: "0 0 12px", color: tp }}>App info</h2>
              {[
                { label: "Timezone", value: timezone },
                { label: "Version", value: version },
                { label: "Auth bypass (SKIP_AUTH)", value: adminStatus?.skipAuth ? "Enabled" : "Disabled" },
                { label: "Asana mode", value: asanaMode === "rest" ? "Direct REST API" : "MCP OAuth" },
                { label: "M365 mode", value: m365Mode === "direct" ? "Direct Graph API" : "Claude MCP" },
                { label: "Zoom", value: adminStatus?.zoomWebhookSet ? "Webhook active" : "Webhook not configured" },
                { label: "Portfolio filter", value: selectedPortfolios.length > 0 ? `${selectedPortfolios.length} selected` : "All projects" },
                { label: "Projects loaded", value: projects.length },
              ].map(r => (
                <div key={r.label} style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", borderBottom: `0.5px solid ${border}` }}>
                  <span style={{ fontSize: 13, color: ts }}>{r.label}</span>
                  <span style={{ fontSize: 13, fontWeight: 500, color: tp }}>{r.value}</span>
                </div>
              ))}
            </div>
          </>
        )}

        <div style={{ textAlign: "center", padding: "1rem 0 0.5rem", borderTop: `0.5px solid ${border}`, marginTop: "0.5rem" }}>
          <span style={{ fontSize: 11, color: tt }}>TAM Dashboard · v{version}</span>
        </div>

      </div>
    </div>
  );
}
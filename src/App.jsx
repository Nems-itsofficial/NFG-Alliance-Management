import { useState, useEffect, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer } from "recharts";
import {
  LayoutDashboard, Users, TrendingUp, CalendarDays, Plus, X, Pencil,
  Trash2, Snowflake, Search, Settings, Save,
  ArrowUp, ArrowDown, Minus, Check, UserX, RotateCcw, Ban, Download, Upload, LogOut, Zap, CheckSquare, Swords, ClipboardList
} from "lucide-react";
import { supabase } from "./supabaseClient.js";
import Login from "./Login.jsx";

const RANKS = ["R1", "R2", "R3", "R4", "R5"];
const FURNACE_LEVELS = ["", ...Array.from({ length: 30 }, (_, i) => `F${i + 1}`), ...Array.from({ length: 10 }, (_, i) => `FC${i + 1}`)];
const TROOP_TIERS = ["", ...Array.from({ length: 12 }, (_, i) => `T${i + 1}`)];
const FC_TROOP_LEVELS = ["", ...Array.from({ length: 10 }, (_, i) => `FC${i + 1}`)];
const CLASSES = [{ key: "infantry", label: "Infantry" }, { key: "marksman", label: "Marksman" }, { key: "lancer", label: "Lancer" }];
const EMPTY_CLASS = { troopTier: "", t12Skills: 0, fcTroopLevel: "" };

const EVENT_TYPES = ["Foundry Battle", "Canyon Clash", "Tyrant Battle", "SvS Battle", "Custom"];
// Matches public/templates/NFG_Canyon_Template.xlsx exactly: team key, display label,
// seat count, and the row where that team's seats start in the template (column C = Players).
const CANYON_TEAMS = [
  { key: "anchor", label: "Anchor Team", seats: 6, startRow: 4 },
  { key: "harass1", label: "Harassment 1 (Right Side)", seats: 5, startRow: 10 },
  { key: "harass2", label: "Harassment 2 (Left Side)", seats: 5, startRow: 15 },
  { key: "harass3", label: "Harassment 3 (Middle / Support)", seats: 5, startRow: 20 },
  { key: "wall", label: "Wall Team", seats: 14, startRow: 25 },
];
const sessionPlaceholder = (type) => (type === "Foundry Battle" || type === "Canyon Clash") ? 'e.g. "Legion 1" or "Legion 2"' : 'e.g. "Team 1" or "Team 2"';

const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
const todayStr = () => new Date().toISOString().slice(0, 10);
const fmtDate = (d) => {
  if (!d) return "—";
  const dt = new Date(d + "T00:00:00");
  if (isNaN(dt)) return d;
  return dt.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" });
};
const daysAgo = (d) => {
  if (!d) return Infinity;
  const dt = new Date(d + "T00:00:00");
  if (isNaN(dt)) return Infinity;
  return Math.floor((Date.now() - dt.getTime()) / 86400000);
};
const fmtNum = (n) => (n === null || n === undefined || n === "" || isNaN(n)) ? "—" : Number(n).toLocaleString();
const fmtCompact = (n) => {
  if (!n || isNaN(n)) return "0";
  if (n >= 1e9) return (n / 1e9).toFixed(2) + "B";
  if (n >= 1e6) return (n / 1e6).toFixed(1) + "M";
  if (n >= 1e3) return (n / 1e3).toFixed(0) + "K";
  return String(Math.round(n));
};
const formatTenure = (joinDate, endDate) => {
  if (!joinDate) return "Unknown";
  const start = new Date(joinDate + "T00:00:00");
  const end = endDate ? new Date(endDate + "T00:00:00") : new Date();
  if (isNaN(start)) return "Unknown";
  const days = Math.floor((end - start) / 86400000);
  if (days < 0) return "Unknown";
  if (days < 30) return `${days}d`;
  const months = Math.floor(days / 30);
  if (months < 12) return `${months}mo`;
  const years = Math.floor(months / 12), remM = months % 12;
  return remM > 0 ? `${years}y ${remM}mo` : `${years}y`;
};
const TenureLabel = ({ joinDate, endDate }) => {
  const val = formatTenure(joinDate, endDate);
  return val === "Unknown" ? <span style={{ color: "var(--steel-dim)", fontStyle: "italic" }}>Unknown</span> : <span>{val}</span>;
};
const SKILL_LABELS = ["No skill", "1st skill", "2nd skill", "3rd skill"];

// ---------------------------------------------------------------- data layer (Supabase)
const rowToMember = (r) => ({ id: r.id, name: r.name, gameId: r.game_id || "", rank: r.rank, status: r.status, customRole: r.custom_role || "", joinDate: r.join_date || "", leftDate: r.left_date || "", notes: r.notes || "" });
const memberToRow = (m) => ({ name: m.name, game_id: m.gameId || "", rank: m.rank, status: m.status, custom_role: m.customRole || "", join_date: m.joinDate || null, left_date: m.leftDate || null, notes: m.notes || "" });
const rowToGrowth = (r) => ({ memberId: r.member_id, power: r.power ?? "", previousPower: r.previous_power ?? "", furnaceLevel: r.furnace_level || "", classes: r.classes || {}, updatedDate: r.updated_date || "" });
const growthToRow = (g) => ({ member_id: g.memberId, power: g.power === "" ? null : g.power, previous_power: g.previousPower === "" ? null : g.previousPower, furnace_level: g.furnaceLevel || "", classes: g.classes || {}, updated_date: g.updatedDate || null });
const rowToEvent = (r) => ({ id: r.id, date: r.date, type: r.type, name: r.name, session: r.session || "", mode: r.mode || "score" });
const eventToRow = (e) => ({ date: e.date, type: e.type, name: e.name, session: e.session || "", mode: e.mode || "score" });
const rowToPart = (r) => ({ id: r.id, eventId: r.event_id, memberId: r.member_id, signedUp: !!r.signed_up, attended: !!r.attended, partial: !!r.partial, score: r.score ?? "", note: r.note || "", strategy: r.strategy || "" });
const partToRow = (p) => ({ event_id: p.eventId, member_id: p.memberId, signed_up: !!p.signedUp, attended: !!p.attended, partial: !!p.partial, score: p.score === "" || p.score === undefined ? null : p.score, note: p.note || "", strategy: p.strategy || "" });
const rowToCanyon = (r) => ({ id: r.id, name: r.name, date: r.date || "", seats: r.seats || {} });
const canyonToRow = (c) => ({ name: c.name, date: c.date || null, seats: c.seats || {} });

async function fetchAllData() {
  const [settingsRes, membersRes, growthRes, eventsRes, partRes, canyonRes] = await Promise.all([
    supabase.from("settings").select("*").eq("id", 1).single(),
    supabase.from("members").select("*").order("created_at", { ascending: true }),
    supabase.from("growth").select("*"),
    supabase.from("events").select("*").order("date", { ascending: false }),
    supabase.from("participation").select("*"),
    supabase.from("canyon_assignments").select("*").order("created_at", { ascending: false }),
  ]);
  const config = settingsRes.data
    ? { allianceName: settingsRes.data.alliance_name || "", leaderName: settingsRes.data.leader_name || "", inactivityDays: settingsRes.data.inactivity_days ?? 10, leaverRetentionDays: settingsRes.data.leaver_retention_days ?? 90, rankLabels: settingsRes.data.rank_labels || {} }
    : { allianceName: "", leaderName: "", inactivityDays: 10, leaverRetentionDays: 90, rankLabels: {} };
  return {
    config,
    members: (membersRes.data || []).map(rowToMember),
    growth: (growthRes.data || []).map(rowToGrowth),
    events: (eventsRes.data || []).map(rowToEvent),
    participation: (partRes.data || []).map(rowToPart),
    canyonAssignments: (canyonRes.data || []).map(rowToCanyon),
  };
}


const STYLE = `
.wsc { --bg:#0B1420; --bg-elev:#0F1C2C; --panel:#132234; --panel-2:#18293C;
  --border:#24384C; --border-soft:#1B2C3E; --frost:#6FCBEA; --frost-dim:#3E6E85;
  --white:#E9F3F7; --steel:#8397AA; --steel-dim:#5C7086; --amber:#E8A33D;
  --danger:#E2604F; --success:#5FBF8C;
  --font-display: "Segoe UI Semibold", "SF Pro Display", -apple-system, "Helvetica Neue", Arial, sans-serif;
  --font-body: -apple-system, BlinkMacSystemFont, "Segoe UI", "Inter", Roboto, "Helvetica Neue", Arial, sans-serif;
  --font-mono: "Cascadia Code", "SFMono-Regular", Consolas, "Liberation Mono", monospace;
  background:var(--bg); color:var(--white); font-family:var(--font-body);
  min-height:100vh; overflow:hidden; border:none; border-radius:0;
  display:flex; }
.wsc * { box-sizing:border-box; }
.wsc-side { width:200px; flex-shrink:0; background:var(--bg-elev); border-right:1px solid var(--border-soft);
  padding:18px 12px; display:flex; flex-direction:column; gap:4px; }
.wsc-brand { display:flex; align-items:center; gap:8px; padding:4px 8px 18px 8px; }
.wsc-brand-mark { width:26px; height:26px; border-radius:7px; background:linear-gradient(180deg,var(--frost),var(--frost-dim));
  display:flex; align-items:center; justify-content:center; flex-shrink:0; }
.wsc-brand-text { font-family:var(--font-display); font-weight:700; font-size:14px; letter-spacing:0.06em;
  text-transform:uppercase; line-height:1.2; }
.wsc-brand-sub { font-size:11px; color:var(--steel-dim); letter-spacing:0.04em; }
.wsc-brand-mobile { display:none; align-items:center; gap:6px; font-family:var(--font-display); font-weight:700;
  font-size:12px; color:var(--frost); letter-spacing:0.04em; text-transform:uppercase; margin-bottom:2px; }
.wsc-nav-btn { display:flex; align-items:center; gap:10px; padding:9px 10px; border-radius:8px; border:none;
  background:transparent; color:var(--steel); font-size:14px; font-weight:500; cursor:pointer; text-align:left;
  font-family:var(--font-body); transition:background .12s,color .12s; }
.wsc-nav-btn:hover { background:var(--panel-2); color:var(--white); }
.wsc-nav-btn.active { background:var(--panel-2); color:var(--frost); box-shadow:inset 2px 0 0 var(--frost); }
.wsc-nav-count { margin-left:auto; font-family:var(--font-mono); font-size:11.5px; color:var(--steel-dim); }
.wsc-main { flex:1; min-width:0; display:flex; flex-direction:column; }
.wsc-topbar { padding:14px 22px; border-bottom:1px solid var(--border-soft); display:flex; align-items:center;
  justify-content:space-between; gap:12px; flex-shrink:0; }
.wsc-title { font-family:var(--font-display); font-size:17.5px; font-weight:700; letter-spacing:0.02em; }
.wsc-title-sub { font-size:12px; color:var(--steel-dim); margin-top:2px; }
.wsc-body { flex:1; overflow-y:auto; padding:20px 22px 28px; }
.wsc-card { background:var(--panel); border:1px solid var(--border-soft); border-radius:11px; padding:16px 18px; }
.wsc-grid { display:grid; gap:12px; }
.wsc-stat-label { font-size:12px; color:var(--steel-dim); text-transform:uppercase; letter-spacing:0.06em; margin-bottom:6px; }
.wsc-stat-val { font-family:var(--font-mono); font-size:26px; font-weight:600; color:var(--white); }
.wsc-btn { display:inline-flex; align-items:center; gap:6px; padding:8px 13px; border-radius:8px;
  border:1px solid var(--border); background:var(--panel-2); color:var(--white); font-size:13.5px; font-weight:600;
  cursor:pointer; font-family:var(--font-body); transition:border-color .12s,background .12s; white-space:nowrap; }
.wsc-btn:hover { border-color:var(--frost-dim); }
.wsc-btn-primary { background:var(--frost); color:#08202C; border-color:var(--frost); }
.wsc-btn-primary:hover { background:#8AD6EE; }
.wsc-btn-danger { background:transparent; color:var(--danger); border-color:#4A2A26; }
.wsc-btn-danger:hover { background:#251616; border-color:var(--danger); }
.wsc-btn-sm { padding:5px 9px; font-size:12.5px; }
.wsc-btn-icon { padding:7px; }
.wsc-input, .wsc-select, .wsc-textarea { width:100%; background:var(--bg-elev); border:1px solid var(--border);
  color:var(--white); border-radius:7px; padding:8px 10px; font-size:14px; font-family:var(--font-body); }
.wsc-input:focus, .wsc-select:focus, .wsc-textarea:focus { outline:none; border-color:var(--frost-dim); }
.wsc-label { font-size:12px; color:var(--steel); margin-bottom:5px; display:block; font-weight:600;
  text-transform:uppercase; letter-spacing:0.04em; }
.wsc-field { margin-bottom:12px; }
.wsc-table { width:100%; border-collapse:collapse; font-size:14px; }
.wsc-table th { text-align:left; font-size:11.5px; text-transform:uppercase; letter-spacing:0.05em;
  color:var(--steel-dim); font-weight:600; padding:8px 10px; border-bottom:1px solid var(--border-soft); }
.wsc-table td { padding:9px 10px; border-bottom:1px solid var(--border-soft); vertical-align:middle; }
.wsc-table tr:last-child td { border-bottom:none; }
.wsc-table tr:hover td { background:var(--panel-2); }
.wsc-badge { display:inline-flex; align-items:center; gap:5px; padding:3px 9px; border-radius:20px;
  font-size:12px; font-weight:700; font-family:var(--font-mono); }
.wsc-modal-wrap { position:fixed; inset:0; background:rgba(4,10,16,0.72); display:flex; align-items:center;
  justify-content:center; z-index:50; padding:20px; }
.wsc-modal { background:var(--panel); border:1px solid var(--border); border-radius:13px; padding:22px;
  width:100%; max-width:460px; max-height:85vh; overflow-y:auto; }
.wsc-modal.wide { max-width:640px; }
.wsc-modal h3 { font-family:var(--font-display); font-size:16px; margin:0 0 16px; letter-spacing:0.02em; }
.wsc-empty { text-align:center; padding:40px 20px; color:var(--steel-dim); }
.wsc-empty-title { color:var(--steel); font-weight:600; margin-bottom:4px; font-size:15px; }
.wsc-search { display:flex; align-items:center; gap:8px; background:var(--bg-elev); border:1px solid var(--border);
  border-radius:8px; padding:0 10px; }
.wsc-search input { border:none; background:transparent; padding:8px 0; font-size:14px; color:var(--white); flex:1; }
.wsc-search input:focus { outline:none; }
.wsc-pill { font-size:11.5px; padding:2px 8px; border-radius:20px; font-weight:700; letter-spacing:0.02em; white-space:nowrap; display:inline-block; }
.wsc-role-badge { display:inline-flex; align-items:center; padding:3px 9px; border-radius:20px; font-size:11.5px; font-weight:700;
  line-height:1; font-family:var(--font-mono); vertical-align:middle;
  letter-spacing:0.03em; text-transform:uppercase; border:1px solid; background:rgba(255,255,255,0.02); }
.wsc-checkbox { width:15px; height:15px; accent-color:var(--frost); cursor:pointer; }
.wsc-bulk-bar { display:flex; align-items:center; gap:10px; background:#6FCBEA14; border:1px solid var(--frost-dim);
  border-radius:8px; padding:8px 12px; margin-bottom:12px; font-size:13.5px; color:var(--frost); }
.wsc-class-block { border:1px solid var(--border-soft); border-radius:9px; padding:10px 12px; margin-bottom:10px; }
.wsc-class-title { font-size:13px; font-weight:700; color:var(--frost); margin-bottom:8px; }
.wsc-scroll::-webkit-scrollbar { width:8px; height:8px; }
.wsc-scroll::-webkit-scrollbar-thumb { background:var(--border); border-radius:8px; }
.wsc-grid-hero { display:grid; gap:12px; grid-template-columns: 220px 1fr; align-items:stretch; }
.wsc-grid-pair { display:grid; gap:12px; grid-template-columns: 1fr 1fr; margin-top:14px; }
.wsc-bottomnav { display:none; }
@media (max-width: 760px) {
  .wsc-side { display:none; }
  .wsc-body { padding:14px 14px 80px; }
  .wsc-topbar { padding:12px 14px; }
  .wsc-title-row { display:flex; flex-direction:column; }
  .wsc-brand-mobile { display:flex; }
  .wsc-grid-hero { grid-template-columns: 1fr; }
  .wsc-grid-pair { grid-template-columns: 1fr; margin-top:0; }
  .wsc-modal { max-width:94vw !important; padding:16px; }
  .wsc-modal.wide { max-width:94vw !important; }
  .wsc-bottomnav { display:flex; position:fixed; bottom:0; left:0; right:0; background:var(--bg-elev);
    border-top:1px solid var(--border-soft); padding:6px 4px calc(6px + env(safe-area-inset-bottom, 0px));
    z-index:40; justify-content:space-around; }
  .wsc-bottomnav-btn { display:flex; flex-direction:column; align-items:center; gap:2px; background:none;
    border:none; color:var(--steel); font-size:11px; font-family:var(--font-body); padding:4px 8px; cursor:pointer;
    position:relative; flex:1; }
  .wsc-bottomnav-btn.active { color:var(--frost); }
  .wsc-bottomnav-count { position:absolute; top:-2px; right:14px; background:var(--danger); color:#fff;
    font-size:9px; font-weight:700; border-radius:20px; padding:1px 4px; font-family:var(--font-mono); }
}
.wsc-power-hero { display:flex; align-items:center; gap:18px; padding:20px 26px; border-radius:14px;
  background: linear-gradient(135deg, rgba(111,203,234,0.14), rgba(111,203,234,0.02) 60%);
  border:1px solid var(--frost-dim); margin-bottom:14px; }
.wsc-power-icon { width:50px; height:50px; border-radius:13px; flex-shrink:0;
  background:linear-gradient(160deg,var(--frost),var(--frost-dim));
  display:flex; align-items:center; justify-content:center; color:#08202C;
  box-shadow:0 0 22px rgba(111,203,234,0.35); }
.wsc-power-label { font-size:12px; color:var(--steel-dim); text-transform:uppercase; letter-spacing:0.09em; font-weight:700; }
.wsc-power-value { font-family:var(--font-mono); font-size:38px; font-weight:700; color:var(--frost);
  text-shadow:0 0 20px rgba(111,203,234,0.5); line-height:1.15; letter-spacing:0.01em; }
.wsc-power-sub { font-size:12.5px; color:var(--steel-dim); margin-top:2px; }
`;

function RankBadge({ rank }) {
  const colors = { R5: "#F2C94C", R4: "#6FCBEA", R3: "#8397AA", R2: "#5C7086", R1: "#3E4E5E" };
  const c = colors[rank] || "#5C7086";
  return <span className="wsc-badge" style={{ background: `${c}22`, color: c }}>{rank}</span>;
}
const NEON_COLORS = ["#39FF88", "#00E5FF", "#FF3EC8", "#FFD23F", "#B14EFF", "#FF5F5F", "#4EFFE0"];
const roleColor = (label) => {
  if (!label) return "#5C7086";
  let hash = 0;
  for (let i = 0; i < label.length; i++) hash = (hash * 31 + label.charCodeAt(i)) >>> 0;
  return NEON_COLORS[hash % NEON_COLORS.length];
};
function RoleBadge({ label }) {
  if (!label) return <span style={{ color: "var(--steel-dim)" }}>—</span>;
  const c = roleColor(label);
  return (
    <span className="wsc-role-badge" style={{ color: c, borderColor: c, boxShadow: `0 0 6px ${c}66, 0 0 1px ${c}` }}>
      {label}
    </span>
  );
}
function Modal({ title, onClose, children, wide }) {
  return (
    <div className="wsc-modal-wrap" onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div className={`wsc-modal ${wide ? "wide" : ""}`}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <h3>{title}</h3>
          <button className="wsc-btn wsc-btn-icon" onClick={onClose} aria-label="Close"><X size={14} /></button>
        </div>
        {children}
      </div>
    </div>
  );
}
function EmptyState({ title, body }) {
  return <div className="wsc-empty"><div className="wsc-empty-title">{title}</div><div style={{ fontSize: 13.5 }}>{body}</div></div>;
}
function ReadinessGauge({ score }) {
  const color = score >= 70 ? "#5FBF8C" : score >= 40 ? "#E8A33D" : "#E2604F";
  const r = 60, cx = 74, cy = 74, startAngle = 135, endAngle = 405;
  const toRad = (deg) => (deg * Math.PI) / 180;
  const arcPoint = (deg) => [cx + r * Math.cos(toRad(deg)), cy + r * Math.sin(toRad(deg))];
  const [sx, sy] = arcPoint(startAngle), [ex, ey] = arcPoint(endAngle);
  const valueAngle = startAngle + (Math.min(100, Math.max(0, score)) / 100) * (endAngle - startAngle);
  const [vx, vy] = arcPoint(valueAngle);
  return (
    <svg viewBox="0 0 148 148" width="148" height="148">
      <path d={`M ${sx} ${sy} A ${r} ${r} 0 1 1 ${ex} ${ey}`} fill="none" stroke="#1B2C3E" strokeWidth="10" strokeLinecap="round" />
      <path d={`M ${sx} ${sy} A ${r} ${r} 0 ${score / 100 > 0.5 ? 1 : 0} 1 ${vx} ${vy}`} fill="none" stroke={color} strokeWidth="10" strokeLinecap="round" />
      <text x="74" y="70" textAnchor="middle" fontFamily="SFMono-Regular, Consolas, monospace" fontSize="26" fontWeight="700" fill="#E9F3F7">{score}</text>
      <text x="74" y="90" textAnchor="middle" fontFamily="Arial, sans-serif" fontSize="10" letterSpacing="0.05em" fill="#8397AA">READINESS</text>
    </svg>
  );
}
function BottomNav({ tab, setTab, leaverCount }) {
  const items = [
    { id: "dashboard", label: "Home", icon: LayoutDashboard },
    { id: "roster", label: "Roster", icon: Users },
    { id: "growth", label: "Growth", icon: TrendingUp },
    { id: "events", label: "Events", icon: CalendarDays },
    { id: "assignments", label: "Plans", icon: ClipboardList },
    { id: "leavers", label: "Leavers", icon: UserX, count: leaverCount },
  ];
  return (
    <div className="wsc-bottomnav">
      {items.map((it) => (
        <button key={it.id} className={`wsc-bottomnav-btn ${tab === it.id ? "active" : ""}`} onClick={() => setTab(it.id)}>
          <it.icon size={18} />
          <span>{it.label}</span>
          {!!it.count && <span className="wsc-bottomnav-count">{it.count}</span>}
        </button>
      ))}
    </div>
  );
}
function Sidebar({ tab, setTab, allianceName, leaderName, leaverCount }) {
  const items = [
    { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
    { id: "roster", label: "Roster", icon: Users },
    { id: "leavers", label: "Leavers", icon: UserX, count: leaverCount },
    { id: "growth", label: "Growth", icon: TrendingUp },
    { id: "events", label: "Events", icon: CalendarDays },
    { id: "assignments", label: "Assignments", icon: ClipboardList },
  ];
  return (
    <div className="wsc-side">
      <div className="wsc-brand">
        <div className="wsc-brand-mark"><Snowflake size={14} color="#08202C" strokeWidth={2.5} /></div>
        <div>
          <div className="wsc-brand-text">{allianceName || "Alliance"}</div>
          <div className="wsc-brand-sub">{leaderName ? `Led by ${leaderName}` : "Command"}</div>
        </div>
      </div>
      {items.map((it) => (
        <button key={it.id} className={`wsc-nav-btn ${tab === it.id ? "active" : ""}`} onClick={() => setTab(it.id)}>
          <it.icon size={15} />{it.label}{!!it.count && <span className="wsc-nav-count">{it.count}</span>}
        </button>
      ))}
    </div>
  );
}
function ConfigModal({ config, onClose, onSave, onExportExcel, onExportSummary, onExportBackup, onImportBackup }) {
  const [name, setName] = useState(config.allianceName || "");
  const [leader, setLeader] = useState(config.leaderName || "");
  const [retention, setRetention] = useState(config.leaverRetentionDays ?? 90);
  const [rankLabels, setRankLabels] = useState({ ...(config.rankLabels || {}) });
  const [importMsg, setImportMsg] = useState("");
  const fileRef = useRef(null);
  const setRankLabel = (rank, val) => setRankLabels((r) => ({ ...r, [rank]: val }));

  const handleImportClick = () => fileRef.current?.click();
  const handleFileChange = async (e) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    if (!window.confirm("Importing replaces all current roster, growth, event, and settings data in this app. Continue?")) return;
    try { await onImportBackup(file); setImportMsg("Import complete."); }
    catch (err) { setImportMsg("Couldn't read that file — make sure it's a backup exported from this app."); }
  };

  return (
    <Modal title="Alliance settings" onClose={onClose}>
      <div className="wsc-field"><label className="wsc-label">Alliance name</label>
        <input className="wsc-input" value={name} onChange={(e) => setName(e.target.value)} placeholder="e.g. Frostbound" /></div>
      <div className="wsc-field"><label className="wsc-label">Alliance leader name</label>
        <input className="wsc-input" value={leader} onChange={(e) => setLeader(e.target.value)} placeholder="e.g. Chief Frost" /></div>
      <div className="wsc-field"><label className="wsc-label">Leaver data retention (days)</label>
        <input className="wsc-input" type="number" min="1" value={retention} onChange={(e) => setRetention(Number(e.target.value))} />
        <div style={{ fontSize: 12.5, color: "var(--steel-dim)", marginTop: 5 }}>Members marked "left" stay on the Leavers tab this long, then are wiped on next load.</div></div>
      <div className="wsc-field">
        <label className="wsc-label">Rank labels (optional)</label>
        <div style={{ fontSize: 12, color: "var(--steel-dim)", marginBottom: 8 }}>Give R1–R5 a custom display name (e.g. R5 → "Alliance Leader"). Leave blank to just show "R5".</div>
        {RANKS_DESC.map((r) => (
          <div key={r} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
            <RankBadge rank={r} />
            <input className="wsc-input" style={{ flex: 1 }} value={rankLabels[r] || ""} onChange={(e) => setRankLabel(r, e.target.value)} placeholder={r} />
          </div>
        ))}
      </div>
      <div style={{ display: "flex", gap: 8, marginTop: 18, justifyContent: "flex-end" }}>
        <button className="wsc-btn" onClick={onClose}>Cancel</button>
        <button className="wsc-btn wsc-btn-primary" onClick={() => onSave({ allianceName: name, leaderName: leader, inactivityDays: config.inactivityDays ?? 10, leaverRetentionDays: retention, rankLabels })}><Save size={13} /> Save</button>
      </div>
      <div style={{ borderTop: "1px solid var(--border-soft)", marginTop: 20, paddingTop: 16 }}>
        <label className="wsc-label">Data</label>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 8 }}>
          <button className="wsc-btn wsc-btn-sm" onClick={onExportExcel}><Download size={12} /> Export Excel — full data (.xlsx)</button>
          <button className="wsc-btn wsc-btn-sm" onClick={onExportSummary}><Download size={12} /> Export Excel — summary (.xlsx)</button>
          <button className="wsc-btn wsc-btn-sm" onClick={onExportBackup}><Download size={12} /> Export backup (.json)</button>
          <button className="wsc-btn wsc-btn-sm" onClick={handleImportClick}><Upload size={12} /> Import backup</button>
          <input ref={fileRef} type="file" accept="application/json" style={{ display: "none" }} onChange={handleFileChange} />
        </div>
        <div style={{ fontSize: 12.5, color: "var(--steel-dim)" }}>
          Full data is every row, raw. Summary is a one-page digest — snapshot stats, who's reliable, who's not, troop spread, recent turnout — meant for sharing with officers who don't need the raw tables. The .json backup is a full, exact copy of everything here — use it to restore into this app later, or as a record if you ever move off Claude.
          {importMsg && <div style={{ color: "var(--frost)", marginTop: 6 }}>{importMsg}</div>}
        </div>
      </div>
    </Modal>
  );
}
function MemberModal({ member, onClose, onSave, onDelete }) {
  const [form, setForm] = useState(member || { name: "", gameId: "", rank: "R1", status: "active", customRole: "", joinDate: "", leftDate: "", notes: "" });
  const isEdit = !!member;
  const set = (k, v) => setForm((f) => ({ ...f, [k]: v }));
  const canSave = form.name.trim().length > 0;
  const isLeft = form.status === "left";
  const toggleLeft = (checked) => {
    if (checked && !form.leftDate) set("leftDate", todayStr());
    setForm((f) => ({ ...f, status: checked ? "left" : "active" }));
  };
  return (
    <Modal title={isEdit ? "Edit member" : "Add member"} onClose={onClose}>
      <div className="wsc-field"><label className="wsc-label">In-game name</label>
        <input className="wsc-input" value={form.name} onChange={(e) => set("name", e.target.value)} placeholder="Chief name" /></div>
      <div className="wsc-field"><label className="wsc-label">Game ID (optional)</label>
        <input className="wsc-input" value={form.gameId} onChange={(e) => set("gameId", e.target.value)} placeholder="123456789" /></div>
      <div style={{ display: "flex", gap: 10 }}>
        <div className="wsc-field" style={{ flex: 1 }}><label className="wsc-label">Rank</label>
          <select className="wsc-select" value={form.rank} onChange={(e) => set("rank", e.target.value)}>{RANKS.map((r) => <option key={r} value={r}>{r}</option>)}</select></div>
        <div className="wsc-field" style={{ flex: 1 }}><label className="wsc-label">Custom role (optional)</label>
          <input className="wsc-input" value={form.customRole || ""} onChange={(e) => set("customRole", e.target.value)} placeholder="e.g. Rally Lead" /></div>
      </div>
      <div style={{ display: "flex", gap: 10 }}>
        <div className="wsc-field" style={{ flex: 1 }}><label className="wsc-label">Join date</label>
          <input className="wsc-input" type="date" value={form.joinDate} onChange={(e) => set("joinDate", e.target.value)} />
          <div style={{ fontSize: 12, color: "var(--steel-dim)", marginTop: 4 }}>Blank by default — set it if you know it, otherwise tenure just shows "Unknown" instead of guessing.</div></div>
        {isLeft && (
          <div className="wsc-field" style={{ flex: 1 }}><label className="wsc-label">Left date</label>
            <input className="wsc-input" type="date" value={form.leftDate} onChange={(e) => set("leftDate", e.target.value)} /></div>
        )}
      </div>
      <div className="wsc-field" style={{ display: "flex", alignItems: "center", gap: 8 }}>
        <input type="checkbox" className="wsc-checkbox" id="left-toggle" checked={isLeft} onChange={(e) => toggleLeft(e.target.checked)} />
        <label htmlFor="left-toggle" style={{ fontSize: 13.5, color: "var(--steel)", cursor: "pointer" }}>This member has left the alliance</label>
      </div>
      <div className="wsc-field"><label className="wsc-label">Notes</label>
        <textarea className="wsc-textarea" rows={2} value={form.notes} onChange={(e) => set("notes", e.target.value)} placeholder="Timezone, anything worth remembering" /></div>
      <div style={{ display: "flex", gap: 8, marginTop: 18, justifyContent: "space-between" }}>
        {isEdit ? <button className="wsc-btn wsc-btn-danger" onClick={() => onDelete(form.id)}><Trash2 size={13} /> Delete permanently</button> : <span />}
        <div style={{ display: "flex", gap: 8 }}>
          <button className="wsc-btn" onClick={onClose}>Cancel</button>
          <button className="wsc-btn wsc-btn-primary" disabled={!canSave} style={{ opacity: canSave ? 1 : 0.5 }}
            onClick={() => canSave && onSave({ ...form, id: form.id || uid() })}><Save size={13} /> Save</button>
        </div>
      </div>
    </Modal>
  );
}
function BulkLeaveModal({ count, onClose, onConfirm }) {
  const [date, setDate] = useState(todayStr());
  return (
    <Modal title={`Mark ${count} member${count > 1 ? "s" : ""} as left`} onClose={onClose}>
      <div className="wsc-field"><label className="wsc-label">Left date (applied to all selected)</label>
        <input className="wsc-input" type="date" value={date} onChange={(e) => setDate(e.target.value)} /></div>
      <div style={{ fontSize: 12.5, color: "var(--steel-dim)", marginBottom: 10 }}>They'll move to the Leavers tab and drop off active tracking.</div>
      <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
        <button className="wsc-btn" onClick={onClose}>Cancel</button>
        <button className="wsc-btn wsc-btn-primary" onClick={() => onConfirm(date)}><Save size={13} /> Confirm</button>
      </div>
    </Modal>
  );
}
const CLASS_COLORS = { infantry: "#6FCBEA", marksman: "#E8A33D", lancer: "#B14EFF" };
function EndgameProgressChart({ members, growth }) {
  const byMember = useMemo(() => { const map = {}; growth.forEach((g) => { map[g.memberId] = g; }); return map; }, [growth]);
  const buckets = ["Below FC", ...FC_TROOP_LEVELS.filter(Boolean)];
  const chartData = useMemo(() => {
    const data = buckets.map((b) => ({ label: b, infantry: 0, marksman: 0, lancer: 0 }));
    members.forEach((m) => {
      const g = byMember[m.id];
      CLASSES.forEach((c) => {
        const cls = g?.classes?.[c.key];
        const bucketLabel = cls?.fcTroopLevel || "Below FC";
        const row = data.find((d) => d.label === bucketLabel);
        if (row) row[c.key] += 1;
      });
    });
    return data;
  }, [members, byMember]);
  return (
    <div className="wsc-card">
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10, flexWrap: "wrap", gap: 8 }}>
        <div className="wsc-stat-label" style={{ margin: 0 }}>Endgame troop progress</div>
        <div style={{ display: "flex", gap: 12 }}>
          {CLASSES.map((c) => (
            <div key={c.key} style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12, color: "var(--steel)" }}>
              <span style={{ width: 8, height: 8, borderRadius: 2, background: CLASS_COLORS[c.key], display: "inline-block" }} />{c.label}
            </div>
          ))}
        </div>
      </div>
      {members.length === 0 ? <EmptyState title="No members yet" body="Add members to see endgame progress here." /> : (
        <ResponsiveContainer width="100%" height={190}>
          <BarChart data={chartData} margin={{ top: 4, right: 8, left: -18, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#1B2C3E" />
            <XAxis dataKey="label" tick={{ fill: "#8397AA", fontSize: 10 }} axisLine={{ stroke: "#24384C" }} tickLine={false} interval={0} angle={-35} textAnchor="end" height={50} />
            <YAxis tick={{ fill: "#8397AA", fontSize: 11 }} axisLine={{ stroke: "#24384C" }} tickLine={false} allowDecimals={false} />
            <Tooltip contentStyle={{ background: "#132234", border: "1px solid #24384C", borderRadius: 8, fontSize: 12 }} labelStyle={{ color: "#E9F3F7" }} />
            <Bar dataKey="infantry" name="Infantry" fill={CLASS_COLORS.infantry} radius={[3, 3, 0, 0]} />
            <Bar dataKey="marksman" name="Marksman" fill={CLASS_COLORS.marksman} radius={[3, 3, 0, 0]} />
            <Bar dataKey="lancer" name="Lancer" fill={CLASS_COLORS.lancer} radius={[3, 3, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      )}
    </div>
  );
}
function T12SkillCard({ members, growth }) {
  const byMember = useMemo(() => { const map = {}; growth.forEach((g) => { map[g.memberId] = g; }); return map; }, [growth]);
  const rows = useMemo(() => CLASSES.map((c) => {
    const counts = [0, 0, 0, 0];
    members.forEach((m) => {
      const cls = byMember[m.id]?.classes?.[c.key];
      if (cls?.troopTier === "T12") counts[cls.t12Skills || 0] += 1;
    });
    return { label: c.label, key: c.key, counts, total: counts.reduce((a, b) => a + b, 0) };
  }), [members, byMember]);
  const anyT12 = rows.some((r) => r.total > 0);
  return (
    <div className="wsc-card">
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
        <Zap size={15} color="var(--frost)" /><div className="wsc-stat-label" style={{ margin: 0 }}>T12 skill progress</div>
      </div>
      {!anyT12 ? <EmptyState title="No T12 troops logged yet" body="Once members hit T12, their unlocked skills (0–3) show here." /> : (
        <table className="wsc-table">
          <thead><tr><th>Class</th><th>No skill</th><th>1st skill</th><th>2nd skill</th><th>3rd skill</th></tr></thead>
          <tbody>{rows.map((r) => (
            <tr key={r.key}>
              <td style={{ fontWeight: 600, color: CLASS_COLORS[r.key] }}>{r.label}</td>
              {r.counts.map((n, i) => <td key={i} style={{ fontFamily: "var(--font-mono)", color: i === 3 && n > 0 ? "var(--success)" : "var(--steel)" }}>{n}</td>)}
            </tr>
          ))}</tbody>
        </table>
      )}
    </div>
  );
}

function Dashboard({ members, growth, events, participation, config }) {
  const activeMembers = members.filter((m) => m.status !== "left");
  const leaverCount = members.length - activeMembers.length;
  const noShows = useMemo(() => {
    const counts = {};
    participation.forEach((p) => { if (p.signedUp && !p.attended) counts[p.memberId] = (counts[p.memberId] || 0) + 1; });
    return Object.entries(counts).map(([memberId, count]) => ({ member: activeMembers.find((m) => m.id === memberId), count }))
      .filter((r) => r.member && r.count >= 2).sort((a, b) => b.count - a.count).slice(0, 8);
  }, [participation, activeMembers]);
  const activeCount = activeMembers.length;
  const latestEvent = events.length ? [...events].sort((a, b) => b.date.localeCompare(a.date))[0] : null;
  const latestEventAttendance = latestEvent && activeCount > 0 ? participation.filter((p) => p.eventId === latestEvent.id && p.attended).length / activeCount : null;
  const readiness = latestEventAttendance !== null ? Math.round(latestEventAttendance * 100) : 0;
  const recentEvents = useMemo(() => [...events].sort((a, b) => b.date.localeCompare(a.date)).slice(0, 6).map((ev) => {
    const rows = participation.filter((p) => p.eventId === ev.id);
    const signed = rows.filter((p) => p.signedUp).length, attended = rows.filter((p) => p.attended).length;
    const rate = activeCount > 0 ? Math.round((attended / activeCount) * 100) : 0;
    return { ev, signed, attended, rate };
  }), [events, participation, activeCount]);
  const avgTurnout = useMemo(() => {
    if (events.length === 0 || activeCount === 0) return 0;
    const total = events.reduce((sum, e) => sum + participation.filter((p) => p.eventId === e.id && p.attended).length / activeCount, 0);
    return Math.round((total / events.length) * 100);
  }, [events, participation, activeCount]);
  const totalPower = useMemo(() => {
    const byMember = {}; growth.forEach((g) => { byMember[g.memberId] = g; });
    return activeMembers.reduce((sum, m) => {
      const p = byMember[m.id]?.power;
      return sum + (p !== "" && p !== undefined && p !== null ? Number(p) : 0);
    }, 0);
  }, [growth, activeMembers]);

  return (
    <div>
      <div className="wsc-power-hero">
        <div className="wsc-power-icon"><Swords size={24} /></div>
        <div>
          <div className="wsc-power-label">Total Alliance Power</div>
          <div className="wsc-power-value">{fmtCompact(totalPower)}</div>
          <div className="wsc-power-sub">Across {activeMembers.length} member{activeMembers.length !== 1 ? "s" : ""}</div>
        </div>
      </div>
      <div className="wsc-grid" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(150px,1fr))", marginBottom: 14 }}>
        <div className="wsc-card"><div className="wsc-stat-label">Total members</div><div className="wsc-stat-val">{activeMembers.length}</div></div>
        <div className="wsc-card"><div className="wsc-stat-label">Events logged</div><div className="wsc-stat-val">{events.length}</div></div>
        <div className="wsc-card"><div className="wsc-stat-label">Avg turnout</div><div className="wsc-stat-val">{events.length > 0 ? `${avgTurnout}%` : "—"}</div></div>
        <div className="wsc-card"><div className="wsc-stat-label">Leavers on file</div><div className="wsc-stat-val">{leaverCount}</div></div>
      </div>
      <div className="wsc-grid-hero">
        <div className="wsc-card" style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 6 }}>
          <ReadinessGauge score={isNaN(readiness) ? 0 : readiness} />
          <div style={{ fontSize: 12, color: "var(--steel-dim)", textAlign: "center" }}>{latestEvent ? "Turnout at the most recent event." : "Log an event to see turnout here."}</div>
        </div>
        <EndgameProgressChart members={activeMembers} growth={growth} />
      </div>
      <div className="wsc-grid-pair">
        <div className="wsc-card">
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}><Ban size={15} color="var(--danger)" /><div className="wsc-stat-label" style={{ margin: 0 }}>Frequent no-shows</div></div>
          {noShows.length === 0 ? <EmptyState title="No repeat no-shows" body="Members who sign up but don't attend twice or more will show here." /> : (
            <table className="wsc-table"><thead><tr><th>Member</th><th>No-shows</th></tr></thead>
              <tbody>{noShows.map(({ member, count }) => <tr key={member.id}><td>{member.name}</td><td style={{ color: "var(--danger)", fontFamily: "var(--font-mono)" }}>{count}</td></tr>)}</tbody></table>
          )}
        </div>
        <T12SkillCard members={activeMembers} growth={growth} />
      </div>
      <div className="wsc-card" style={{ marginTop: 14 }}>
        <div className="wsc-stat-label" style={{ marginBottom: 10 }}>Recent event turnout</div>
        {recentEvents.length === 0 ? <EmptyState title="No events logged" body="Create one on the Events tab to start tracking who shows up." /> : (
          <table className="wsc-table"><thead><tr><th>Date</th><th>Event</th><th>Signed up</th><th>Attended</th><th>Turnout</th></tr></thead>
            <tbody>{recentEvents.map(({ ev, signed, attended, rate }) => (
              <tr key={ev.id}><td style={{ color: "var(--steel)" }}>{fmtDate(ev.date)}</td><td style={{ fontWeight: 600 }}>{ev.name}{ev.session ? ` · ${ev.session}` : ""}</td>
                <td style={{ fontFamily: "var(--font-mono)" }}>{signed}</td><td style={{ fontFamily: "var(--font-mono)", color: "var(--success)" }}>{attended}</td>
                <td style={{ fontFamily: "var(--font-mono)" }}>{rate}%</td></tr>
            ))}</tbody></table>
        )}
      </div>
    </div>
  );
}
const RANKS_DESC = [...RANKS].reverse();
const ROSTER_COLS = [
  { key: "role", label: "Role", width: 130, align: "center" },
  { key: "power", label: "Power", width: 130, align: "right" },
  { key: "joined", label: "Joined", width: 120, align: "center" },
  { key: "tenure", label: "Tenure", width: 100, align: "center" },
  { key: "last", label: "Last activity", width: 130, align: "center" },
];
function RankGroup({ rank, rankLabel, list, selectMode, selected, onToggleOne, onToggleAllInGroup, onEdit, lastActivityByMember, powerByMember }) {
  const allSelected = list.length > 0 && list.every((m) => selected.has(m.id));
  return (
    <div className="wsc-card" style={{ padding: 0, overflow: "hidden", marginBottom: 12 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "12px 16px", borderBottom: "1px solid var(--border-soft)" }}>
        <RankBadge rank={rank} />
        {rankLabel && <span style={{ fontSize: 13, color: "var(--frost)", fontWeight: 600 }}>{rankLabel}</span>}
        <span style={{ fontSize: 13, color: "var(--steel-dim)" }}>·</span>
        <span style={{ fontSize: 13, color: "var(--steel-dim)" }}>{list.length} member{list.length !== 1 ? "s" : ""}</span>
      </div>
      <div style={{ overflowX: "auto" }}>
        <table className="wsc-table" style={{ tableLayout: "fixed", width: "100%" }}>
          <thead><tr>
            {selectMode && <th style={{ width: 30 }}><input type="checkbox" className="wsc-checkbox" checked={allSelected} onChange={() => onToggleAllInGroup(list)} /></th>}
            <th style={{ width: 200 }}>Name</th>
            {ROSTER_COLS.map((c) => <th key={c.key} style={{ width: c.width, textAlign: c.align }}>{c.label}</th>)}
            <th>Notes</th>
            <th style={{ width: 36 }}></th>
          </tr></thead>
          <tbody>
            {list.map((m) => {
              const last = lastActivityByMember[m.id];
              const power = powerByMember[m.id];
              return (
                <tr key={m.id}>
                  {selectMode && <td onClick={(e) => e.stopPropagation()}><input type="checkbox" className="wsc-checkbox" checked={selected.has(m.id)} onChange={() => onToggleOne(m.id)} /></td>}
                  <td style={{ fontWeight: 600, cursor: "pointer" }} onClick={() => onEdit(m)}>{m.name}</td>
                  <td style={{ textAlign: "center" }}><RoleBadge label={m.customRole} /></td>
                  <td style={{ fontFamily: "var(--font-mono)", color: "var(--steel)", textAlign: "right" }}>{power !== undefined && power !== "" ? fmtNum(power) : "—"}</td>
                  <td style={{ color: "var(--steel)", textAlign: "center" }}>{m.joinDate ? fmtDate(m.joinDate) : "Unknown"}</td>
                  <td style={{ color: "var(--steel)", fontFamily: "var(--font-mono)", textAlign: "center" }}><TenureLabel joinDate={m.joinDate} /></td>
                  <td style={{ color: "var(--steel)", textAlign: "center" }}>{last ? fmtDate(last) : "No data"}</td>
                  <td style={{ color: "var(--steel-dim)", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }} title={m.notes || ""}>{m.notes || "—"}</td>
                  <td style={{ textAlign: "right", cursor: "pointer" }} onClick={() => onEdit(m)}><Pencil size={13} color="var(--steel-dim)" /></td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}
function RosterTab({ members, growth, lastActivityByMember, rankLabels, onEdit, onBulkLeave }) {
  const [query, setQuery] = useState("");
  const [rankFilter, setRankFilter] = useState("all");
  const [roleFilter, setRoleFilter] = useState("all");
  const [selectMode, setSelectMode] = useState(false);
  const [selected, setSelected] = useState(new Set());
  const [showBulk, setShowBulk] = useState(false);
  const roster = members.filter((m) => m.status !== "left");
  const powerByMember = useMemo(() => {
    const map = {};
    growth.forEach((g) => { map[g.memberId] = g.power; });
    return map;
  }, [growth]);
  const roleOptions = useMemo(() => [...new Set(roster.map((m) => m.customRole).filter(Boolean))].sort(), [roster]);
  const filtered = roster.filter((m) => {
    const matchesQuery = m.name.toLowerCase().includes(query.toLowerCase()) || (m.gameId || "").includes(query);
    const matchesRank = rankFilter === "all" || m.rank === rankFilter;
    const matchesRole = roleFilter === "all" || m.customRole === roleFilter;
    return matchesQuery && matchesRank && matchesRole;
  });
  const groups = useMemo(() => RANKS_DESC.map((rank) => {
    const list = filtered.filter((m) => m.rank === rank).slice().sort((a, b) => {
      const pa = powerByMember[a.id], pb = powerByMember[b.id];
      const na = pa !== undefined && pa !== "" ? Number(pa) : -1;
      const nb = pb !== undefined && pb !== "" ? Number(pb) : -1;
      return nb - na;
    });
    return { rank, list };
  }).filter((g) => g.list.length > 0), [filtered, powerByMember]);
  const toggleOne = (id) => setSelected((s) => { const n = new Set(s); n.has(id) ? n.delete(id) : n.add(id); return n; });
  const toggleAllInGroup = (list) => setSelected((s) => {
    const n = new Set(s);
    const allIn = list.every((m) => n.has(m.id));
    list.forEach((m) => allIn ? n.delete(m.id) : n.add(m.id));
    return n;
  });
  const exitSelectMode = () => { setSelectMode(false); setSelected(new Set()); };
  return (
    <div>
      <div style={{ display: "flex", gap: 10, marginBottom: 14, flexWrap: "wrap" }}>
        <div className="wsc-search" style={{ flex: 1, minWidth: 180 }}><Search size={14} color="var(--steel-dim)" /><input placeholder="Search name or ID" value={query} onChange={(e) => setQuery(e.target.value)} /></div>
        <select className="wsc-select" style={{ width: 140 }} value={roleFilter} onChange={(e) => setRoleFilter(e.target.value)}>
          <option value="all">All roles</option>{roleOptions.map((r) => <option key={r} value={r}>{r}</option>)}
        </select>
        <select className="wsc-select" style={{ width: 120 }} value={rankFilter} onChange={(e) => setRankFilter(e.target.value)}>
          <option value="all">All ranks</option>{RANKS.map((r) => <option key={r} value={r}>{r}</option>)}
        </select>
        {!selectMode ? (
          <button className="wsc-btn wsc-btn-sm" onClick={() => setSelectMode(true)}><CheckSquare size={13} /> Select</button>
        ) : (
          <button className="wsc-btn wsc-btn-sm" onClick={exitSelectMode}><X size={13} /> Cancel</button>
        )}
      </div>
      {selectMode && selected.size > 0 && (
        <div className="wsc-bulk-bar">
          <span>{selected.size} selected</span>
          <button className="wsc-btn wsc-btn-sm" onClick={() => setShowBulk(true)}><UserX size={12} /> Mark as left</button>
          <button className="wsc-btn wsc-btn-sm" onClick={() => setSelected(new Set())}>Clear</button>
        </div>
      )}
      {roster.length === 0 ? (
        <div className="wsc-card"><EmptyState title="No members yet" body='Add your first chief with "Add member" above to start the roster.' /></div>
      ) : groups.length === 0 ? (
        <div className="wsc-card"><EmptyState title="No matches" body="Try a different search or filter." /></div>
      ) : (
        groups.map((g) => (
          <RankGroup key={g.rank} rank={g.rank} rankLabel={rankLabels?.[g.rank]} list={g.list} selectMode={selectMode} selected={selected}
            onToggleOne={toggleOne} onToggleAllInGroup={toggleAllInGroup} onEdit={onEdit} lastActivityByMember={lastActivityByMember} powerByMember={powerByMember} />
        ))
      )}
      {showBulk && <BulkLeaveModal count={selected.size} onClose={() => setShowBulk(false)}
        onConfirm={(date) => { onBulkLeave([...selected], date); setSelected(new Set()); setShowBulk(false); exitSelectMode(); }} />}
    </div>
  );
}
function LeaversTab({ members, retentionDays, rankLabels, onReactivate, onPurgeNow, onEdit }) {
  const leavers = members.filter((m) => m.status === "left").sort((a, b) => (b.leftDate || "").localeCompare(a.leftDate || ""));
  return (
    <div>
      <div className="wsc-card" style={{ marginBottom: 14, fontSize: 13, color: "var(--steel)" }}>
        Members marked "Left alliance" land here. Their history stays for {retentionDays} days after leaving, then is wiped automatically the next time this app is opened.
      </div>
      <div className="wsc-card" style={{ padding: 0, overflow: "hidden" }}>
        {leavers.length === 0 ? <EmptyState title="No former members on file" body='Mark a member as "Left alliance" from the Roster tab to move them here.' /> : (
          <div style={{ overflowX: "auto" }}>
            <table className="wsc-table">
              <thead><tr><th>Name</th><th>Game ID</th><th>Rank</th><th>Role</th><th>Time in alliance</th><th>Left on</th><th>Purges in</th><th></th></tr></thead>
              <tbody>
                {leavers.map((m) => {
                  const elapsed = daysAgo(m.leftDate);
                  const remaining = Math.max(0, retentionDays - elapsed);
                  return (
                    <tr key={m.id}>
                      <td style={{ fontWeight: 600, cursor: "pointer" }} onClick={() => onEdit(m)}>{m.name}</td>
                      <td style={{ color: "var(--steel-dim)" }}>{m.gameId || "—"}</td>
                      <td><RankBadge rank={m.rank} /></td>
                      <td><RoleBadge label={m.customRole} /></td>
                      <td style={{ fontFamily: "var(--font-mono)" }}><TenureLabel joinDate={m.joinDate} endDate={m.leftDate} /></td>
                      <td style={{ color: "var(--steel)" }}>{fmtDate(m.leftDate)}</td>
                      <td style={{ color: remaining <= 7 ? "var(--amber)" : "var(--steel)", fontFamily: "var(--font-mono)" }}>{remaining} days</td>
                      <td style={{ textAlign: "right" }}>
                        <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                          <button className="wsc-btn wsc-btn-sm" onClick={() => onReactivate(m)}><RotateCcw size={12} /> Reactivate</button>
                          <button className="wsc-btn wsc-btn-sm wsc-btn-danger" onClick={() => onPurgeNow(m.id)}><Trash2 size={12} /> Wipe now</button>
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}
function ClassFields({ label, value, onChange }) {
  const v = value || EMPTY_CLASS;
  const set = (k, val) => onChange({ ...v, [k]: val });
  return (
    <div className="wsc-class-block">
      <div className="wsc-class-title">{label}</div>
      <div style={{ display: "flex", gap: 8 }}>
        <div className="wsc-field" style={{ flex: 1, marginBottom: v.troopTier === "T12" ? 8 : 0 }}>
          <label className="wsc-label">Troop tier</label>
          <select className="wsc-select" value={v.troopTier} onChange={(e) => set("troopTier", e.target.value)}>
            {TROOP_TIERS.map((t) => <option key={t} value={t}>{t || "Not set"}</option>)}
          </select>
        </div>
        <div className="wsc-field" style={{ flex: 1, marginBottom: 0 }}>
          <label className="wsc-label">FC troop level</label>
          <select className="wsc-select" value={v.fcTroopLevel} onChange={(e) => set("fcTroopLevel", e.target.value)}>
            {FC_TROOP_LEVELS.map((t) => <option key={t} value={t}>{t || "Not set"}</option>)}
          </select>
        </div>
      </div>
      {v.troopTier === "T12" && (
        <div className="wsc-field" style={{ marginTop: 8, marginBottom: 0 }}>
          <label className="wsc-label">T12 skills unlocked</label>
          <select className="wsc-select" value={v.t12Skills} onChange={(e) => set("t12Skills", Number(e.target.value))}>
            {[0, 1, 2, 3].map((n) => <option key={n} value={n}>{n} of 3</option>)}
          </select>
        </div>
      )}
    </div>
  );
}
function LogGrowthModal({ members, profiles, initialMemberId, onClose, onSave }) {
  const [memberId, setMemberId] = useState(initialMemberId || members[0]?.id || "");
  const existing = profiles.find((p) => p.memberId === memberId);
  const [power, setPower] = useState(existing?.power ?? "");
  const [furnaceLevel, setFurnaceLevel] = useState(existing?.furnaceLevel ?? "");
  const [classes, setClasses] = useState(existing?.classes ?? { infantry: { ...EMPTY_CLASS }, marksman: { ...EMPTY_CLASS }, lancer: { ...EMPTY_CLASS } });

  const selectMember = (id) => {
    setMemberId(id);
    const p = profiles.find((x) => x.memberId === id);
    setPower(p?.power ?? ""); setFurnaceLevel(p?.furnaceLevel ?? "");
    setClasses(p?.classes ?? { infantry: { ...EMPTY_CLASS }, marksman: { ...EMPTY_CLASS }, lancer: { ...EMPTY_CLASS } });
  };

  return (
    <Modal title="Log player data" onClose={onClose} wide>
      <div className="wsc-field"><label className="wsc-label">Member</label>
        <select className="wsc-select" value={memberId} onChange={(e) => selectMember(e.target.value)}>{members.map((m) => <option key={m.id} value={m.id}>{m.name}</option>)}</select></div>
      <div style={{ display: "flex", gap: 10 }}>
        <div className="wsc-field" style={{ flex: 1 }}><label className="wsc-label">Power (optional)</label>
          <input className="wsc-input" type="number" placeholder="42000000" value={power} onChange={(e) => setPower(e.target.value)} /></div>
        <div className="wsc-field" style={{ flex: 1 }}><label className="wsc-label">Furnace level</label>
          <select className="wsc-select" value={furnaceLevel} onChange={(e) => setFurnaceLevel(e.target.value)}>{FURNACE_LEVELS.map((f) => <option key={f} value={f}>{f || "Not set"}</option>)}</select></div>
      </div>
      {CLASSES.map((c) => <ClassFields key={c.key} label={c.label} value={classes[c.key]} onChange={(v) => setClasses((cs) => ({ ...cs, [c.key]: v }))} />)}
      <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 6 }}>
        <button className="wsc-btn" onClick={onClose}>Cancel</button>
        <button className="wsc-btn wsc-btn-primary" disabled={!memberId} style={{ opacity: memberId ? 1 : 0.5 }}
          onClick={() => memberId && onSave({
            memberId,
            power: power === "" ? "" : Number(power),
            previousPower: existing?.power ?? "",
            furnaceLevel, classes, updatedDate: todayStr(),
          })}><Save size={13} /> Save profile</button>
      </div>
    </Modal>
  );
}
function DeltaTag({ delta }) {
  if (delta === null || delta === undefined) return <span style={{ color: "var(--steel-dim)" }}>—</span>;
  if (delta === 0) return <span style={{ color: "var(--steel)", display: "inline-flex", alignItems: "center", gap: 3 }}><Minus size={11} />0</span>;
  const up = delta > 0;
  return <span style={{ color: up ? "var(--success)" : "var(--danger)", display: "inline-flex", alignItems: "center", gap: 3, fontFamily: "var(--font-mono)" }}>{up ? <ArrowUp size={11} /> : <ArrowDown size={11} />}{fmtNum(Math.abs(delta))}</span>;
}
const SKILL_SUFFIX = ["", "I", "II", "III"];
const tierLabel = (cls) => {
  if (!cls || !cls.troopTier) return "—";
  if (cls.troopTier === "T12") {
    const n = cls.t12Skills || 0;
    return n > 0 ? `T12 - ${SKILL_SUFFIX[n]}` : "T12";
  }
  return cls.troopTier;
};
const fcLabel = (cls) => cls?.fcTroopLevel || "—";
function MergedClassCell({ classes, getLabel }) {
  const order = [["infantry", "Infantry"], ["lancer", "Lancer"], ["marksman", "Marksman"]];
  return (
    <span style={{ fontFamily: "var(--font-mono)" }}>
      {order.map(([key], i) => (
        <span key={key}>
          <span style={{ color: CLASS_COLORS[key] }}>{getLabel(classes?.[key])}</span>
          {i < order.length - 1 && <span style={{ color: "var(--steel-dim)" }}> / </span>}
        </span>
      ))}
    </span>
  );
}
function GrowthTab({ members, growth, onEditMember }) {
  const byMember = useMemo(() => { const map = {}; growth.forEach((g) => { map[g.memberId] = g; }); return map; }, [growth]);
  return (
    <div>
      {members.length === 0 ? <div className="wsc-card"><EmptyState title="Add members first" body="Growth tracking needs a roster — head to the Roster tab." /></div> : (
        <div className="wsc-card" style={{ padding: 0, overflow: "hidden" }}>
          <div style={{ padding: "10px 18px 0", fontSize: 12.5, color: "var(--steel-dim)" }}>Click a row to update that member's profile. Troops shown as Infantry / Lancer / Marksman.</div>
          <div style={{ overflowX: "auto" }}>
            <table className="wsc-table">
              <thead><tr><th>Member</th><th>Power</th><th>Change</th><th>Furnace</th><th>Troop tier</th><th>FC level</th><th>Updated</th></tr></thead>
              <tbody>
                {members.map((m) => {
                  const g = byMember[m.id];
                  const delta = g && g.power !== "" && g.previousPower !== "" && g.previousPower !== undefined ? Number(g.power) - Number(g.previousPower) : null;
                  return (
                    <tr key={m.id} style={{ cursor: "pointer" }} onClick={() => onEditMember(m.id)}>
                      <td style={{ fontWeight: 600 }}>{m.name}</td>
                      <td style={{ fontFamily: "var(--font-mono)" }}>{g && g.power !== "" ? fmtNum(g.power) : "—"}</td>
                      <td><DeltaTag delta={delta} /></td>
                      <td style={{ fontFamily: "var(--font-mono)" }}>{g?.furnaceLevel || "—"}</td>
                      <td><MergedClassCell classes={g?.classes} getLabel={tierLabel} /></td>
                      <td><MergedClassCell classes={g?.classes} getLabel={fcLabel} /></td>
                      <td style={{ color: "var(--steel-dim)" }}>{g ? fmtDate(g.updatedDate) : "Never"}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}
function EventModal({ onClose, onSave }) {
  const [type, setType] = useState(EVENT_TYPES[0]);
  const [name, setName] = useState("");
  const [session, setSession] = useState("");
  const [date, setDate] = useState(todayStr());
  const [mode, setMode] = useState("score");
  const canSave = type !== "Custom" || name.trim().length > 0;
  return (
    <Modal title="Add event" onClose={onClose}>
      <div className="wsc-field"><label className="wsc-label">Event type</label>
        <select className="wsc-select" value={type} onChange={(e) => setType(e.target.value)}>{EVENT_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}</select></div>
      {type === "Custom" && <div className="wsc-field"><label className="wsc-label">Event name</label>
        <input className="wsc-input" value={name} onChange={(e) => setName(e.target.value)} placeholder="e.g. Fishing tournament" /></div>}
      <div className="wsc-field"><label className="wsc-label">Session label (optional)</label>
        <input className="wsc-input" value={session} onChange={(e) => setSession(e.target.value)} placeholder={sessionPlaceholder(type)} /></div>
      <div className="wsc-field"><label className="wsc-label">Date</label><input className="wsc-input" type="date" value={date} onChange={(e) => setDate(e.target.value)} /></div>
      <div className="wsc-field">
        <label className="wsc-label">Track by</label>
        <div style={{ display: "flex", gap: 8 }}>
          <button type="button" className="wsc-btn" style={{ flex: 1, background: mode === "score" ? "var(--frost)" : "var(--panel-2)", color: mode === "score" ? "#08202C" : "var(--white)", borderColor: mode === "score" ? "var(--frost)" : "var(--border)" }} onClick={() => setMode("score")}>Score</button>
          <button type="button" className="wsc-btn" style={{ flex: 1, background: mode === "strategy" ? "var(--frost)" : "var(--panel-2)", color: mode === "strategy" ? "#08202C" : "var(--white)", borderColor: mode === "strategy" ? "var(--frost)" : "var(--border)" }} onClick={() => setMode("strategy")}>Strategy compliance</button>
        </div>
        <div style={{ fontSize: 12, color: "var(--steel-dim)", marginTop: 6 }}>
          {mode === "score" ? "Log a numeric score per member — good for events like Foundry or Canyon Clash." : "Track whether each member followed the called strategy — good for coordinated events like Tyrant or SvS. Produces a ranked leaderboard."}
        </div>
      </div>
      <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
        <button className="wsc-btn" onClick={onClose}>Cancel</button>
        <button className="wsc-btn wsc-btn-primary" disabled={!canSave} style={{ opacity: canSave ? 1 : 0.5 }}
          onClick={() => canSave && onSave({ id: uid(), name: type === "Custom" ? name : type, type, session, date, mode })}><Save size={13} /> Create event</button>
      </div>
    </Modal>
  );
}
const STRATEGY_LABELS = { followed: "Followed strategy", partial: "Partially followed", none: "Did not follow" };
const strategyPoints = (p) => {
  if (!p?.attended) return -1;
  let pts = 2;
  if (!p.partial) pts += 1;
  if (p.strategy === "followed") pts += 2;
  else if (p.strategy === "partial") pts += 1;
  return pts;
};
function ModeBadge({ mode }) {
  const isStrategy = mode === "strategy";
  const c = isStrategy ? "#B14EFF" : "#6FCBEA";
  return (
    <span className="wsc-role-badge" style={{ color: c, borderColor: c, boxShadow: `0 0 6px ${c}66, 0 0 1px ${c}`, marginLeft: 8 }}>
      {isStrategy ? "Strategy" : "Score"}
    </span>
  );
}
function AddParticipantPicker({ members, excludeIds, onAdd }) {
  const [query, setQuery] = useState("");
  const candidates = query
    ? members.filter((m) => m.status !== "left" && !excludeIds.has(m.id) && m.name.toLowerCase().includes(query.toLowerCase())).slice(0, 20)
    : [];
  return (
    <div style={{ position: "relative", marginBottom: 12 }}>
      <div className="wsc-search">
        <Search size={13} color="var(--steel-dim)" />
        <input placeholder="Search a member to add to this event…" value={query} onChange={(e) => setQuery(e.target.value)} />
      </div>
      {query && (
        <div style={{ position: "absolute", top: "100%", left: 0, right: 0, background: "var(--panel-2)", border: "1px solid var(--border)", borderRadius: 8, marginTop: 4, maxHeight: 180, overflowY: "auto", zIndex: 10 }}>
          {candidates.length === 0 ? (
            <div style={{ padding: 10, fontSize: 12.5, color: "var(--steel-dim)" }}>No matches</div>
          ) : candidates.map((m) => (
            <div key={m.id} style={{ padding: "8px 12px", cursor: "pointer", fontSize: 13, borderBottom: "1px solid var(--border-soft)" }}
              onClick={() => { onAdd(m.id); setQuery(""); }}>
              {m.name}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
function EventDetail({ event, members, participation, onClose, onDelete, onToggleSignUp, onToggleAttend, onTogglePartial, onScore, onNote, onSetStrategy }) {
  const isStrategy = event.mode === "strategy";
  const partMap = {};
  participation.filter((p) => p.eventId === event.id).forEach((p) => { partMap[p.memberId] = p; });
  const participantIds = new Set(Object.entries(partMap).filter(([, p]) => p.signedUp).map(([id]) => id));
  let addedMembers = members.filter((m) => participantIds.has(m.id));
  if (isStrategy) {
    addedMembers = addedMembers.slice().sort((a, b) => strategyPoints(partMap[b.id]) - strategyPoints(partMap[a.id]));
  }
  const attendedCount = addedMembers.filter((m) => partMap[m.id]?.attended).length;
  let rankCounter = 0;
  const handleDelete = () => {
    if (window.confirm(`Delete "${event.name}${event.session ? ` · ${event.session}` : ""}" and all its participation data? This can't be undone.`)) onDelete(event.id);
  };
  return (
    <Modal title={`${event.name}${event.session ? ` · ${event.session}` : ""} — ${fmtDate(event.date)}`} onClose={onClose} wide>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
        <div style={{ fontSize: 13, color: "var(--steel-dim)" }}>
          {addedMembers.length} added · {attendedCount} attended
          <ModeBadge mode={event.mode} />
        </div>
        <button className="wsc-btn wsc-btn-sm wsc-btn-danger" onClick={handleDelete}><Trash2 size={12} /> Delete event</button>
      </div>
      <AddParticipantPicker members={members} excludeIds={participantIds} onAdd={(id) => onToggleSignUp(event.id, id, true)} />
      <div className="wsc-scroll" style={{ maxHeight: 380, overflowY: "auto", overflowX: "auto" }}>
        <table className="wsc-table" style={{ minWidth: 560 }}>
          <thead><tr>
            {isStrategy && <th>Rank</th>}
            <th>Member</th>
            {isStrategy ? <><th>Attended</th><th>Full duration?</th><th>Strategy</th></> : <th>Participated</th>}
            {!isStrategy && <th>Score</th>}
            <th>Notes</th>
            <th></th>
          </tr></thead>
          <tbody>
            {addedMembers.length === 0 ? (
              <tr><td colSpan={8} style={{ padding: 20 }}><EmptyState title="No one added yet" body="Search above to add members who are taking part in this event." /></td></tr>
            ) : addedMembers.map((m) => {
              const p = partMap[m.id];
              const noShow = !p?.attended;
              const showRank = isStrategy && p?.attended;
              if (showRank) rankCounter += 1;
              return (
                <tr key={m.id}>
                  {isStrategy && <td style={{ fontFamily: "var(--font-mono)", color: showRank ? "var(--frost)" : "var(--steel-dim)" }}>{showRank ? `#${rankCounter}` : "—"}</td>}
                  <td>{m.name}{isStrategy && noShow && <span className="wsc-pill" style={{ background: "#E2604F22", color: "var(--danger)", marginLeft: 8 }}>No-show</span>}</td>
                  {isStrategy ? (
                    <>
                      <td><button className="wsc-btn wsc-btn-icon" style={{ background: p?.attended ? "#5FBF8C22" : "transparent", borderColor: p?.attended ? "var(--success)" : "var(--border)" }}
                        onClick={() => onToggleAttend(event.id, m.id, !p?.attended)} aria-label="Toggle attended"><Check size={13} color={p?.attended ? "var(--success)" : "var(--steel-dim)"} /></button></td>
                      <td>
                        {p?.attended ? (
                          <button className="wsc-btn wsc-btn-sm" style={{ background: p?.partial ? "#E8A33D22" : "#5FBF8C22", borderColor: p?.partial ? "var(--amber)" : "var(--success)", color: p?.partial ? "var(--amber)" : "var(--success)" }}
                            onClick={() => onTogglePartial(event.id, m.id, !p?.partial)}>{p?.partial ? "Left early" : "Full"}</button>
                        ) : <span style={{ color: "var(--steel-dim)" }}>—</span>}
                      </td>
                      <td>
                        <select className="wsc-select" style={{ width: 150 }} value={p?.strategy || ""} onChange={(e) => onSetStrategy(event.id, m.id, e.target.value)}>
                          <option value="">Not set</option>
                          <option value="followed">{STRATEGY_LABELS.followed}</option>
                          <option value="partial">{STRATEGY_LABELS.partial}</option>
                          <option value="none">{STRATEGY_LABELS.none}</option>
                        </select>
                      </td>
                    </>
                  ) : (
                    <td>
                      <button className="wsc-btn wsc-btn-icon" style={{ background: p?.attended ? "#6FCBEA22" : "transparent", borderColor: p?.attended ? "var(--frost)" : "var(--border)" }}
                        onClick={() => onToggleAttend(event.id, m.id, !p?.attended)} aria-label="Toggle participated"><Check size={13} color={p?.attended ? "var(--frost)" : "var(--steel-dim)"} /></button>
                    </td>
                  )}
                  {!isStrategy && <td><input className="wsc-input" style={{ width: 90 }} type="number" placeholder="—" defaultValue={p?.score ?? ""} onBlur={(e) => onScore(event.id, m.id, e.target.value)} /></td>}
                  <td><input className="wsc-input" style={{ width: 150 }} placeholder="e.g. left after 20 min" defaultValue={p?.note ?? ""} onBlur={(e) => onNote(event.id, m.id, e.target.value)} /></td>
                  <td><button className="wsc-btn wsc-btn-icon" onClick={() => onToggleSignUp(event.id, m.id, false)} aria-label="Remove from event" title="Remove from event"><X size={12} color="var(--steel-dim)" /></button></td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </Modal>
  );
}
function EventsTab({ events, members, participation, onOpenEvent }) {
  const [filter, setFilter] = useState("all");
  const activeMembers = members.filter((m) => m.status !== "left");
  const sorted = [...events].filter((e) => filter === "all" || e.type === filter).sort((a, b) => b.date.localeCompare(a.date));
  return (
    <div>
      <div style={{ display: "flex", justifyContent: "flex-end", alignItems: "center", marginBottom: 14, gap: 8 }}>
        <span style={{ fontSize: 12.5, color: "var(--steel-dim)" }}>Show</span>
        <select className="wsc-select" style={{ width: 190 }} value={filter} onChange={(e) => setFilter(e.target.value)}>
          <option value="all">All types</option>{EVENT_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
        </select>
      </div>
      <div className="wsc-card" style={{ padding: 0, overflow: "hidden" }}>
        {sorted.length === 0 ? <EmptyState title="No events logged" body='Use "Add event" up top to create an occurrence for any event type.' /> : (
          <div style={{ overflowX: "auto" }}>
            <table className="wsc-table">
              <thead><tr><th>Date</th><th>Type</th><th>Session</th><th>Signed up</th><th>Attended</th><th>Left early</th><th>No-shows</th><th>Turnout</th></tr></thead>
              <tbody>
                {sorted.map((ev) => {
                  const rows = participation.filter((p) => p.eventId === ev.id);
                  const signed = rows.filter((p) => p.signedUp).length, attended = rows.filter((p) => p.attended).length;
                  const partial = rows.filter((p) => p.attended && p.partial).length;
                  const noShows = rows.filter((p) => p.signedUp && !p.attended).length;
                  const rate = activeMembers.length > 0 ? Math.round((attended / activeMembers.length) * 100) : 0;
                  return (
                    <tr key={ev.id} style={{ cursor: "pointer" }} onClick={() => onOpenEvent(ev)}>
                      <td style={{ color: "var(--steel)" }}>{fmtDate(ev.date)}</td>
                      <td style={{ fontWeight: 600 }}>{ev.name}<ModeBadge mode={ev.mode} /></td>
                      <td style={{ color: "var(--steel-dim)" }}>{ev.session || "—"}</td>
                      <td style={{ fontFamily: "var(--font-mono)" }}>{signed}</td>
                      <td style={{ fontFamily: "var(--font-mono)", color: "var(--success)" }}>{attended}</td>
                      <td style={{ fontFamily: "var(--font-mono)", color: partial > 0 ? "var(--amber)" : "var(--steel-dim)" }}>{partial}</td>
                      <td style={{ fontFamily: "var(--font-mono)", color: noShows > 0 ? "var(--danger)" : "var(--steel-dim)" }}>{noShows}</td>
                      <td style={{ fontFamily: "var(--font-mono)" }}>{rate}%</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}
function AssignmentModal({ onClose, onSave }) {
  const [name, setName] = useState("");
  const [date, setDate] = useState(todayStr());
  const canSave = name.trim().length > 0;
  return (
    <Modal title="New Canyon Clash plan" onClose={onClose}>
      <div className="wsc-field"><label className="wsc-label">Name</label>
        <input className="wsc-input" value={name} onChange={(e) => setName(e.target.value)} placeholder='e.g. "Canyon Clash — Aug 28"' /></div>
      <div className="wsc-field"><label className="wsc-label">Date</label>
        <input className="wsc-input" type="date" value={date} onChange={(e) => setDate(e.target.value)} /></div>
      <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
        <button className="wsc-btn" onClick={onClose}>Cancel</button>
        <button className="wsc-btn wsc-btn-primary" disabled={!canSave} style={{ opacity: canSave ? 1 : 0.5 }}
          onClick={() => canSave && onSave({ id: uid(), name, date, seats: {} })}><Save size={13} /> Create</button>
      </div>
    </Modal>
  );
}
function CanyonEditor({ assignment, members, growth, onChangeSeats, onExport, onDelete, onBack }) {
  const roster = members.filter((m) => m.status !== "left");
  const powerByMember = useMemo(() => { const map = {}; growth.forEach((g) => { map[g.memberId] = g.power; }); return map; }, [growth]);
  const seats = assignment.seats || {};
  const usedIds = new Set(Object.values(seats).flat().filter(Boolean));
  const setSeat = (teamKey, idx, memberId) => {
    const arr = [...(seats[teamKey] || [])];
    arr[idx] = memberId || null;
    onChangeSeats({ ...seats, [teamKey]: arr });
  };
  const filledCount = [...usedIds].length;
  const totalSeats = CANYON_TEAMS.reduce((s, t) => s + t.seats, 0);
  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexWrap: "wrap", gap: 8 }}>
        <button className="wsc-btn wsc-btn-sm" onClick={onBack}>&larr; Back to plans</button>
        <div style={{ display: "flex", gap: 8 }}>
          <span style={{ fontSize: 12.5, color: "var(--steel-dim)", alignSelf: "center" }}>{filledCount} / {totalSeats} seats filled</span>
          <button className="wsc-btn wsc-btn-sm wsc-btn-danger" onClick={onDelete}><Trash2 size={12} /> Delete</button>
          <button className="wsc-btn wsc-btn-primary" onClick={onExport}><Download size={13} /> Export to Excel</button>
        </div>
      </div>
      {CANYON_TEAMS.map((team) => (
        <div key={team.key} className="wsc-card" style={{ marginBottom: 12 }}>
          <div className="wsc-stat-label" style={{ marginBottom: 10 }}>{team.label} <span style={{ color: "var(--steel-dim)", fontWeight: 400 }}>({team.seats} seats)</span></div>
          {Array.from({ length: team.seats }).map((_, i) => {
            const currentId = seats[team.key]?.[i] || "";
            return (
              <div key={i} style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 6 }}>
                <span style={{ width: 22, textAlign: "right", color: "var(--steel-dim)", fontFamily: "var(--font-mono)", fontSize: 12 }}>{i + 1}</span>
                <select className="wsc-select" value={currentId} onChange={(e) => setSeat(team.key, i, e.target.value)}>
                  <option value="">— empty —</option>
                  {roster.filter((m) => !usedIds.has(m.id) || m.id === currentId).map((m) => (
                    <option key={m.id} value={m.id}>{m.name}{powerByMember[m.id] ? ` (${fmtNum(powerByMember[m.id])})` : ""}</option>
                  ))}
                </select>
              </div>
            );
          })}
        </div>
      ))}
    </div>
  );
}
function AssignmentsTab({ canyonAssignments, members, growth, onCreate, onChangeSeats, onExport, onDelete }) {
  const [showNew, setShowNew] = useState(false);
  const [openId, setOpenId] = useState(null);
  const open = canyonAssignments.find((a) => a.id === openId);
  if (open) {
    return (
      <CanyonEditor assignment={open} members={members} growth={growth}
        onChangeSeats={(seats) => onChangeSeats(open.id, seats)}
        onExport={() => onExport(open)}
        onDelete={() => { onDelete(open.id); setOpenId(null); }}
        onBack={() => setOpenId(null)} />
    );
  }
  return (
    <div>
      <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 14 }}>
        <button className="wsc-btn wsc-btn-primary" onClick={() => setShowNew(true)}><Plus size={13} /> New Canyon Clash plan</button>
      </div>
      <div style={{ fontSize: 12.5, color: "var(--steel-dim)", marginBottom: 14 }}>Foundry Battle and Custom event planning are coming soon — Canyon Clash is ready now.</div>
      {canyonAssignments.length === 0 ? (
        <div className="wsc-card"><EmptyState title="No Canyon Clash plans yet" body='Click "New Canyon Clash plan" to start assigning members to teams.' /></div>
      ) : (
        <div className="wsc-grid" style={{ gridTemplateColumns: "repeat(auto-fill, minmax(220px,1fr))" }}>
          {canyonAssignments.map((a) => {
            const filled = Object.values(a.seats || {}).flat().filter(Boolean).length;
            const total = CANYON_TEAMS.reduce((s, t) => s + t.seats, 0);
            return (
              <div key={a.id} className="wsc-card" style={{ cursor: "pointer" }} onClick={() => setOpenId(a.id)}>
                <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 3 }}>{a.name}</div>
                <div style={{ fontSize: 12, color: "var(--steel-dim)", marginBottom: 10 }}>{a.date ? fmtDate(a.date) : "No date set"}</div>
                <div style={{ fontFamily: "var(--font-mono)", fontSize: 13, color: "var(--frost)" }}>{filled} / {total} seats filled</div>
              </div>
            );
          })}
        </div>
      )}
      {showNew && <AssignmentModal onClose={() => setShowNew(false)} onSave={(a) => { onCreate(a); setShowNew(false); }} />}
    </div>
  );
}

export default function App() {
  const [session, setSession] = useState(undefined); // undefined = checking, null = signed out
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState("");
  const [tab, setTab] = useState("dashboard");
  const [config, setConfig] = useState({ allianceName: "", leaderName: "", inactivityDays: 10, leaverRetentionDays: 90 });
  const [members, setMembers] = useState([]);
  const [growth, setGrowth] = useState([]);
  const [events, setEvents] = useState([]);
  const [participation, setParticipation] = useState([]);
  const [canyonAssignments, setCanyonAssignments] = useState([]);

  const [showConfig, setShowConfig] = useState(false);
  const [memberModal, setMemberModal] = useState(null);
  const [showAddMember, setShowAddMember] = useState(false);
  const [showLogGrowth, setShowLogGrowth] = useState(false);
  const [showAddEvent, setShowAddEvent] = useState(false);
  const [openEvent, setOpenEvent] = useState(null);
  const [growthPreset, setGrowthPreset] = useState(null);

  const openGrowthFor = useCallback((memberId) => { setGrowthPreset(memberId); setShowLogGrowth(true); }, []);

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: sub } = supabase.auth.onAuthStateChange((_event, s) => setSession(s));
    return () => sub.subscription.unsubscribe();
  }, []);

  useEffect(() => {
    // Recharts' auto-sizing containers can get stuck thinking they have zero
    // width/height after a tab has been backgrounded, since nothing tells them
    // to recalculate. Nudging a resize event on return fixes it.
    const onVisible = () => {
      if (document.visibilityState === "visible") {
        window.dispatchEvent(new Event("resize"));
      }
    };
    document.addEventListener("visibilitychange", onVisible);
    return () => document.removeEventListener("visibilitychange", onVisible);
  }, []);

  const userId = session?.user?.id || null;

  useEffect(() => {
    if (!userId) return;
    let cancelled = false;
    (async () => {
      setLoading(true);
      setLoadError("");
      try {
        const { config: cfg, members: mem, growth: gr, events: ev, participation: part, canyonAssignments: canyon } = await fetchAllData();

        const retention = cfg.leaverRetentionDays ?? 90;
        const purgedIds = mem.filter((m) => m.status === "left" && daysAgo(m.leftDate) > retention).map((m) => m.id);
        let finalMembers = mem, finalGrowth = gr, finalPart = part;
        if (purgedIds.length > 0) {
          // FK cascade on members removes their growth/participation rows automatically.
          await supabase.from("members").delete().in("id", purgedIds);
          finalMembers = mem.filter((m) => !purgedIds.includes(m.id));
          finalGrowth = gr.filter((g) => !purgedIds.includes(g.memberId));
          finalPart = part.filter((p) => !purgedIds.includes(p.memberId));
        }
        if (cancelled) return;
        setConfig(cfg); setMembers(finalMembers); setGrowth(finalGrowth); setEvents(ev); setParticipation(finalPart); setCanyonAssignments(canyon);
      } catch (err) {
        console.error("Failed to load alliance data:", err);
        if (!cancelled) setLoadError(err?.message || "Failed to load data. Check your connection and try refreshing.");
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => { cancelled = true; };
    // Only re-run when the signed-in user actually changes (sign in/out),
    // not on every background token refresh (e.g. from switching browser tabs).
  }, [userId]);

  const lastActivityByMember = useMemo(() => {
    const map = {};
    members.forEach((m) => { map[m.id] = null; });
    const consider = (memberId, date) => { if (!date) return; if (!map[memberId] || date > map[memberId]) map[memberId] = date; };
    growth.forEach((g) => consider(g.memberId, g.updatedDate));
    participation.forEach((p) => { if (!p.attended) return; const ev = events.find((e) => e.id === p.eventId); if (ev) consider(p.memberId, ev.date); });
    return map;
  }, [members, growth, participation, events]);

  const saveConfig = useCallback(async (next) => {
    const prev = config;
    setConfig(next); setShowConfig(false);
    const { error } = await supabase.from("settings").update({
      alliance_name: next.allianceName, leader_name: next.leaderName,
      inactivity_days: next.inactivityDays, leaver_retention_days: next.leaverRetentionDays,
      rank_labels: next.rankLabels || {},
    }).eq("id", 1);
    if (error) {
      console.error("Failed to save settings:", error);
      setConfig(prev);
      window.alert(`Couldn't save settings: ${error.message}`);
    }
  }, [config]);

  const saveMember = useCallback(async (m) => {
    if (memberModal) {
      await supabase.from("members").update(memberToRow(m)).eq("id", m.id);
      setMembers((prev) => prev.map((x) => x.id === m.id ? m : x));
    } else {
      const { data, error } = await supabase.from("members").insert(memberToRow(m)).select().single();
      if (!error && data) setMembers((prev) => [...prev, rowToMember(data)]);
    }
    setMemberModal(null); setShowAddMember(false);
  }, [memberModal]);

  const deleteMember = useCallback(async (id) => {
    await supabase.from("members").delete().eq("id", id); // cascades to growth + participation
    setMembers((prev) => prev.filter((m) => m.id !== id));
    setGrowth((prev) => prev.filter((g) => g.memberId !== id));
    setParticipation((prev) => prev.filter((p) => p.memberId !== id));
    setMemberModal(null);
  }, []);

  const reactivateMember = useCallback(async (m) => {
    await supabase.from("members").update({ status: "active", left_date: null }).eq("id", m.id);
    setMembers((prev) => prev.map((x) => x.id === m.id ? { ...x, status: "active", leftDate: "" } : x));
  }, []);

  const bulkMarkLeft = useCallback(async (ids, leftDate) => {
    await supabase.from("members").update({ status: "left", left_date: leftDate }).in("id", ids);
    setMembers((prev) => prev.map((m) => ids.includes(m.id) ? { ...m, status: "left", leftDate } : m));
  }, []);

  const saveGrowth = useCallback(async (profile) => {
    const { data, error } = await supabase.from("growth").upsert(growthToRow(profile), { onConflict: "member_id" }).select().single();
    const saved = !error && data ? rowToGrowth(data) : profile;
    setGrowth((prev) => prev.some((g) => g.memberId === saved.memberId) ? prev.map((g) => g.memberId === saved.memberId ? saved : g) : [...prev, saved]);
    setShowLogGrowth(false);
  }, []);

  const addEvent = useCallback(async (ev) => {
    const { data, error } = await supabase.from("events").insert(eventToRow(ev)).select().single();
    if (!error && data) setEvents((prev) => [rowToEvent(data), ...prev]);
    setShowAddEvent(false);
  }, []);

  const deleteEvent = useCallback(async (id) => {
    await supabase.from("events").delete().eq("id", id); // cascades to participation rows
    setEvents((prev) => prev.filter((e) => e.id !== id));
    setParticipation((prev) => prev.filter((p) => p.eventId !== id));
    setOpenEvent(null);
  }, []);

  const createCanyonAssignment = useCallback(async (a) => {
    const { data, error } = await supabase.from("canyon_assignments").insert(canyonToRow(a)).select().single();
    if (!error && data) setCanyonAssignments((prev) => [rowToCanyon(data), ...prev]);
    else if (error) window.alert(`Couldn't create plan: ${error.message}`);
  }, []);

  const changeCanyonSeats = useCallback(async (id, seats) => {
    setCanyonAssignments((prev) => prev.map((a) => a.id === id ? { ...a, seats } : a)); // optimistic
    const { error } = await supabase.from("canyon_assignments").update({ seats }).eq("id", id);
    if (error) window.alert(`Couldn't save seat change: ${error.message}`);
  }, []);

  const deleteCanyonAssignment = useCallback(async (id) => {
    await supabase.from("canyon_assignments").delete().eq("id", id);
    setCanyonAssignments((prev) => prev.filter((a) => a.id !== id));
  }, []);

  const exportCanyonAssignment = useCallback(async (assignment) => {
    try {
      const res = await fetch("templates/NFG_Canyon_Template.xlsx");
      if (!res.ok) throw new Error("Template not found at public/templates/NFG_Canyon_Template.xlsx");
      const buf = await res.arrayBuffer();
      const zip = await JSZip.loadAsync(buf);
      const sheetFile = zip.file("xl/worksheets/sheet1.xml");
      if (!sheetFile) throw new Error("Unexpected template structure — sheet1.xml not found");
      let sheetXml = await sheetFile.async("string");

      const escapeXml = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
      const seats = assignment.seats || {};

      for (const team of CANYON_TEAMS) {
        const arr = seats[team.key] || [];
        for (let i = 0; i < team.seats; i++) {
          const memberId = arr[i];
          if (!memberId) continue;
          const member = members.find((m) => m.id === memberId);
          if (!member) continue;
          const row = team.startRow + i;
          const cellRef = `C${row}`;
          const nameXml = `<is><t xml:space="preserve">${escapeXml(member.name)}</t></is>`;
          const replaceWith = (attrs) => {
            const styleMatch = attrs.match(/s="(\d+)"/);
            const s = styleMatch ? ` s="${styleMatch[1]}"` : "";
            return `<c r="${cellRef}"${s} t="inlineStr">${nameXml}</c>`;
          };
          const selfClose = new RegExp(`<c r="${cellRef}"([^>]*)/>`);
          const withBody = new RegExp(`<c r="${cellRef}"([^>]*)>.*?</c>`);
          if (selfClose.test(sheetXml)) sheetXml = sheetXml.replace(selfClose, (_, attrs) => replaceWith(attrs));
          else if (withBody.test(sheetXml)) sheetXml = sheetXml.replace(withBody, (_, attrs) => replaceWith(attrs));
        }
      }

      zip.file("xl/worksheets/sheet1.xml", sheetXml);
      const blob = await zip.generateAsync({ type: "blob", compression: "DEFLATE" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${(assignment.name || "canyon-clash").replace(/[^a-z0-9]+/gi, "-")}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error("Canyon export failed:", err);
      window.alert(`Couldn't export: ${err.message}`);
    }
  }, [members]);

  const upsertParticipation = useCallback(async (eventId, memberId, patch) => {
    const existing = participation.find((p) => p.eventId === eventId && p.memberId === memberId);
    const merged = existing ? { ...existing, ...patch } : { eventId, memberId, signedUp: false, attended: false, partial: false, score: "", note: "", ...patch };
    const { data, error } = await supabase.from("participation").upsert(partToRow(merged), { onConflict: "event_id,member_id" }).select().single();
    const saved = !error && data ? rowToPart(data) : merged;
    setParticipation((prev) => existing ? prev.map((p) => p === existing ? saved : p) : [...prev, saved]);
  }, [participation]);
  const toggleSignUp = useCallback((eventId, memberId, v) => upsertParticipation(eventId, memberId, { signedUp: v }), [upsertParticipation]);
  const toggleAttend = useCallback((eventId, memberId, v) => upsertParticipation(eventId, memberId, { attended: v, partial: v ? undefined : false }), [upsertParticipation]);
  const togglePartial = useCallback((eventId, memberId, v) => upsertParticipation(eventId, memberId, { partial: v }), [upsertParticipation]);
  const setScore = useCallback((eventId, memberId, score) => upsertParticipation(eventId, memberId, { score }), [upsertParticipation]);
  const setNote = useCallback((eventId, memberId, note) => upsertParticipation(eventId, memberId, { note }), [upsertParticipation]);
  const setStrategy = useCallback((eventId, memberId, strategy) => upsertParticipation(eventId, memberId, { strategy }), [upsertParticipation]);

  const roster = members.filter((m) => m.status !== "left");
  const leaverCount = members.filter((m) => m.status === "left").length;

  const exportExcel = useCallback(() => {
    const wsMembers = XLSX.utils.json_to_sheet(members.map((m) => ({
      Name: m.name, "Game ID": m.gameId || "", Rank: m.rank, "Custom role": m.customRole || "", Status: m.status,
      "Join date": m.joinDate || "", "Left date": m.leftDate || "",
      Tenure: formatTenure(m.joinDate, m.leftDate || null), Notes: m.notes || "",
    })));
    const growthRows = growth.map((g) => {
      const m = members.find((x) => x.id === g.memberId);
      const c = (k) => g.classes?.[k] || EMPTY_CLASS;
      return {
        Member: m ? m.name : g.memberId, Power: g.power, "Previous power": g.previousPower, "Furnace level": g.furnaceLevel,
        "Infantry tier": c("infantry").troopTier, "Infantry FC level": c("infantry").fcTroopLevel, "Infantry T12 skills": c("infantry").t12Skills,
        "Marksman tier": c("marksman").troopTier, "Marksman FC level": c("marksman").fcTroopLevel, "Marksman T12 skills": c("marksman").t12Skills,
        "Lancer tier": c("lancer").troopTier, "Lancer FC level": c("lancer").fcTroopLevel, "Lancer T12 skills": c("lancer").t12Skills,
        Updated: g.updatedDate,
      };
    });
    const wsGrowth = XLSX.utils.json_to_sheet(growthRows);
    const wsEvents = XLSX.utils.json_to_sheet(events.map((e) => ({ Date: e.date, Type: e.type, Name: e.name, Session: e.session || "", Mode: e.mode === "strategy" ? "Strategy compliance" : "Score" })));
    const partRows = participation.map((p) => {
      const ev = events.find((e) => e.id === p.eventId), m = members.find((x) => x.id === p.memberId);
      return {
        Event: ev ? ev.name : p.eventId, Date: ev ? ev.date : "", Member: m ? m.name : p.memberId,
        "Signed up": p.signedUp ? "Yes" : "No", Attended: p.attended ? "Yes" : "No",
        "Full duration": p.attended ? (p.partial ? "No — left early" : "Yes") : "",
        Score: p.score, Strategy: p.strategy ? (STRATEGY_LABELS[p.strategy] || p.strategy) : "", Notes: p.note || "",
      };
    });
    const wsPart = XLSX.utils.json_to_sheet(partRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, wsMembers, "Roster");
    XLSX.utils.book_append_sheet(wb, wsGrowth, "Growth");
    XLSX.utils.book_append_sheet(wb, wsEvents, "Events");
    XLSX.utils.book_append_sheet(wb, wsPart, "Participation");
    XLSX.writeFile(wb, `${(config.allianceName || "alliance").replace(/[^a-z0-9]+/gi, "-")}-export.xlsx`);
  }, [members, growth, events, participation, config]);

  const exportSummary = useCallback(() => {
    const activeMembers = members.filter((m) => m.status !== "left");
    const leaverCount = members.length - activeMembers.length;
    const growthByMember = {}; growth.forEach((g) => { growthByMember[g.memberId] = g; });
    const roleCounts = {};
    activeMembers.forEach((m) => { if (m.customRole) roleCounts[m.customRole] = (roleCounts[m.customRole] || 0) + 1; });
    const roleRows = Object.entries(roleCounts).sort((a, b) => b[1] - a[1]);

    const attendCounts = {}, noShowCounts = {};
    participation.forEach((p) => {
      if (p.attended) attendCounts[p.memberId] = (attendCounts[p.memberId] || 0) + 1;
      if (p.signedUp && !p.attended) noShowCounts[p.memberId] = (noShowCounts[p.memberId] || 0) + 1;
    });
    const nameOf = (id) => members.find((m) => m.id === id)?.name || id;
    const topAttendees = Object.entries(attendCounts).map(([id, n]) => ({ name: nameOf(id), n })).sort((a, b) => b.n - a.n).slice(0, 10);
    const frequentNoShows = Object.entries(noShowCounts).map(([id, n]) => ({ name: nameOf(id), n })).filter((r) => r.n >= 2).sort((a, b) => b.n - a.n).slice(0, 10);

    const bucketOf = (cls) => {
      if (!cls) return "Not set";
      if (cls.fcTroopLevel) return "FC troops";
      if (cls.troopTier === "T12") return "T12";
      if (["T7", "T8", "T9", "T10", "T11"].includes(cls.troopTier)) return "T7\u2013T11";
      if (["T1", "T2", "T3", "T4", "T5", "T6"].includes(cls.troopTier)) return "T1\u2013T6";
      return "Not set";
    };
    const bucketOrder = ["Not set", "T1\u2013T6", "T7\u2013T11", "T12", "FC troops"];
    const tierRows = bucketOrder.map((b) => {
      const row = { Bucket: b, Infantry: 0, Marksman: 0, Lancer: 0 };
      activeMembers.forEach((m) => {
        const g = growthByMember[m.id];
        row.Infantry += bucketOf(g?.classes?.infantry) === b ? 1 : 0;
        row.Marksman += bucketOf(g?.classes?.marksman) === b ? 1 : 0;
        row.Lancer += bucketOf(g?.classes?.lancer) === b ? 1 : 0;
      });
      return row;
    });

    const recentEvents = [...events].sort((a, b) => b.date.localeCompare(a.date)).slice(0, 10).map((ev) => {
      const rows = participation.filter((p) => p.eventId === ev.id);
      const attended = rows.filter((p) => p.attended).length;
      const rate = activeMembers.length > 0 ? Math.round((attended / activeMembers.length) * 100) : 0;
      return { Date: ev.date, Event: ev.name + (ev.session ? ` (${ev.session})` : ""), "Signed up": rows.filter((p) => p.signedUp).length, Attended: attended, "Turnout %": rate };
    });

    const aoa = [
      [`${config.allianceName || "Alliance"} — Summary Report`],
      [`Generated ${todayStr()}${config.leaderName ? ` · Led by ${config.leaderName}` : ""}`],
      [],
      ["Alliance snapshot"],
      ["Total members", activeMembers.length],
      ["Total events logged", events.length],
      ["Leavers on file", leaverCount],
      [],
      ["Custom roles"],
      ...(roleRows.length ? roleRows.map(([role, n]) => [role, n]) : [["No custom roles assigned yet"]]),
      [],
      ["Most reliable attendees (events attended)"],
      ...(topAttendees.length ? topAttendees.map((r) => [r.name, r.n]) : [["No attendance logged yet"]]),
      [],
      ["Frequent no-shows (signed up, didn't attend, 2+ times)"],
      ...(frequentNoShows.length ? frequentNoShows.map((r) => [r.name, r.n]) : [["No repeat no-shows"]]),
      [],
      ["Troop tier spread (active roster)"],
      ["Bucket", "Infantry", "Marksman", "Lancer"],
      ...tierRows.map((r) => [r.Bucket, r.Infantry, r.Marksman, r.Lancer]),
      [],
      ["Recent event turnout"],
      ["Date", "Event", "Signed up", "Attended", "Turnout %"],
      ...(recentEvents.length ? recentEvents.map((r) => [r.Date, r.Event, r["Signed up"], r.Attended, r["Turnout %"]]) : [["No events logged yet"]]),
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 30 }, { wch: 20 }, { wch: 14 }, { wch: 12 }, { wch: 12 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Summary");
    XLSX.writeFile(wb, `${(config.allianceName || "alliance").replace(/[^a-z0-9]+/gi, "-")}-summary.xlsx`);
  }, [members, growth, events, participation, config]);

  const exportBackup = useCallback(() => {
    const payload = { config, members, growth, events, participation, exportedAt: new Date().toISOString(), version: 1 };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `${(config.allianceName || "alliance").replace(/[^a-z0-9]+/gi, "-")}-backup.json`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, [config, members, growth, events, participation]);

  const importBackup = useCallback(async (file) => {
    const text = await file.text();
    const data = JSON.parse(text);
    if (!data || !Array.isArray(data.members)) throw new Error("Invalid backup file");
    const oldMembers = data.members || [], oldGrowth = data.growth || [], oldEvents = data.events || [], oldPart = data.participation || [];
    const cfg = data.config || config;

    // Wipe existing data (participation/growth first, they reference members & events).
    await supabase.from("participation").delete().not("id", "is", null);
    await supabase.from("growth").delete().not("member_id", "is", null);
    await supabase.from("events").delete().not("id", "is", null);
    await supabase.from("members").delete().not("id", "is", null);

    // Re-insert members and events one at a time so we can map their old
    // (Claude-artifact-generated) string ids to fresh Postgres UUIDs.
    const memberIdMap = {}, eventIdMap = {};
    for (const m of oldMembers) {
      const { data: row } = await supabase.from("members").insert(memberToRow(m)).select().single();
      if (row) memberIdMap[m.id] = row.id;
    }
    for (const e of oldEvents) {
      const { data: row } = await supabase.from("events").insert(eventToRow(e)).select().single();
      if (row) eventIdMap[e.id] = row.id;
    }
    for (const g of oldGrowth) {
      const newMemberId = memberIdMap[g.memberId];
      if (!newMemberId) continue;
      await supabase.from("growth").insert(growthToRow({ ...g, memberId: newMemberId }));
    }
    for (const p of oldPart) {
      const newEventId = eventIdMap[p.eventId], newMemberId = memberIdMap[p.memberId];
      if (!newEventId || !newMemberId) continue;
      await supabase.from("participation").insert(partToRow({ ...p, eventId: newEventId, memberId: newMemberId }));
    }
    await supabase.from("settings").update({
      alliance_name: cfg.allianceName, leader_name: cfg.leaderName,
      inactivity_days: cfg.inactivityDays, leaver_retention_days: cfg.leaverRetentionDays,
    }).eq("id", 1);

    const fresh = await fetchAllData();
    setConfig(fresh.config); setMembers(fresh.members); setGrowth(fresh.growth); setEvents(fresh.events); setParticipation(fresh.participation);
  }, [config]);

  if (session === undefined) {
    return <div className="wsc" style={{ alignItems: "center", justifyContent: "center", minHeight: 300 }}><style>{STYLE}</style><div style={{ color: "var(--steel)", fontSize: 14 }}>Loading…</div></div>;
  }
  if (!session) return <Login />;
  if (loading) return <div className="wsc" style={{ alignItems: "center", justifyContent: "center", minHeight: 300 }}><style>{STYLE}</style><div style={{ color: "var(--steel)", fontSize: 14 }}>Loading alliance data…</div></div>;
  if (loadError) return (
    <div className="wsc" style={{ alignItems: "center", justifyContent: "center", minHeight: 300 }}>
      <style>{STYLE}</style>
      <div style={{ textAlign: "center", maxWidth: 320 }}>
        <div style={{ color: "var(--danger)", fontSize: 14, marginBottom: 12 }}>{loadError}</div>
        <button className="wsc-btn wsc-btn-primary" onClick={() => window.location.reload()}>Retry</button>
      </div>
    </div>
  );

  return (
    <div className="wsc">
      <style>{STYLE}</style>
      <Sidebar tab={tab} setTab={setTab} allianceName={config.allianceName} leaderName={config.leaderName} leaverCount={leaverCount} />
      <div className="wsc-main">
        <div className="wsc-topbar">
          <div>
            <div className="wsc-brand-mobile"><Snowflake size={12} />{config.allianceName || "Alliance"}</div>
            <div className="wsc-title">
              {tab === "dashboard" && "Dashboard"}{tab === "roster" && "Roster"}{tab === "leavers" && "Leavers"}
              {tab === "growth" && "Growth"}{tab === "events" && "Events"}{tab === "assignments" && "Assignments"}
            </div>
            <div className="wsc-title-sub">{roster.length} members tracked{leaverCount > 0 ? ` · ${leaverCount} former` : ""}</div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            {tab === "dashboard" && <button className="wsc-btn wsc-btn-icon" onClick={() => setShowConfig(true)} aria-label="Settings"><Settings size={15} /></button>}
            {tab === "roster" && <button className="wsc-btn wsc-btn-primary" onClick={() => setShowAddMember(true)}><Plus size={13} /> Add member</button>}
            {tab === "growth" && <button className="wsc-btn wsc-btn-primary" onClick={() => openGrowthFor(null)} disabled={roster.length === 0} style={{ opacity: roster.length === 0 ? 0.5 : 1 }}><Plus size={13} /> Log data</button>}
            {tab === "events" && <button className="wsc-btn wsc-btn-primary" onClick={() => setShowAddEvent(true)}><Plus size={13} /> Add event</button>}
            <button className="wsc-btn wsc-btn-icon" onClick={() => supabase.auth.signOut()} aria-label="Sign out" title="Sign out"><LogOut size={15} /></button>
          </div>
        </div>
        <div className="wsc-body wsc-scroll">
          {tab === "dashboard" && <Dashboard members={members} growth={growth} events={events} participation={participation} config={config} />}
          {tab === "roster" && <RosterTab members={members} growth={growth} lastActivityByMember={lastActivityByMember} rankLabels={config.rankLabels} onEdit={(m) => setMemberModal(m)} onBulkLeave={bulkMarkLeft} />}
          {tab === "leavers" && <LeaversTab members={members} retentionDays={config.leaverRetentionDays ?? 90} rankLabels={config.rankLabels} onReactivate={reactivateMember} onPurgeNow={deleteMember} onEdit={(m) => setMemberModal(m)} />}
          {tab === "growth" && <GrowthTab members={roster} growth={growth} onEditMember={openGrowthFor} />}
          {tab === "events" && <EventsTab events={events} members={members} participation={participation} onOpenEvent={(ev) => setOpenEvent(ev)} />}
          {tab === "assignments" && <AssignmentsTab canyonAssignments={canyonAssignments} members={members} growth={growth}
            onCreate={createCanyonAssignment} onChangeSeats={changeCanyonSeats} onExport={exportCanyonAssignment} onDelete={deleteCanyonAssignment} />}
        </div>
      </div>
      <BottomNav tab={tab} setTab={setTab} leaverCount={leaverCount} />
      {showConfig && <ConfigModal config={config} onClose={() => setShowConfig(false)} onSave={saveConfig}
        onExportExcel={exportExcel} onExportSummary={exportSummary} onExportBackup={exportBackup} onImportBackup={importBackup} />}
      {(showAddMember || memberModal) && <MemberModal member={memberModal} onClose={() => { setShowAddMember(false); setMemberModal(null); }} onSave={saveMember} onDelete={deleteMember} />}
      {showLogGrowth && <LogGrowthModal members={roster} profiles={growth} initialMemberId={growthPreset} onClose={() => { setShowLogGrowth(false); setGrowthPreset(null); }} onSave={saveGrowth} />}
      {showAddEvent && <EventModal onClose={() => setShowAddEvent(false)} onSave={addEvent} />}
      {openEvent && <EventDetail event={openEvent} members={members} participation={participation} onClose={() => setOpenEvent(null)} onDelete={deleteEvent} onToggleSignUp={toggleSignUp} onToggleAttend={toggleAttend} onTogglePartial={togglePartial} onScore={setScore} onNote={setNote} onSetStrategy={setStrategy} />}
    </div>
  );
}

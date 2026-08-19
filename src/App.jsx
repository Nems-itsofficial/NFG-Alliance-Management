import { useState, useEffect, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer } from "recharts";
import {
  LayoutDashboard, Users, TrendingUp, CalendarDays, Plus, X, Pencil,
  Trash2, Snowflake, TriangleAlert, Search, Settings, Save,
  ArrowUp, ArrowDown, Minus, Check, UserX, RotateCcw, Ban, Download, Upload, LogOut
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
const classSummary = (cls) => {
  if (!cls) return "—";
  if (cls.fcTroopLevel) return cls.fcTroopLevel;
  if (cls.troopTier === "T12") return `T12 · ${cls.t12Skills || 0}/3`;
  return cls.troopTier || "—";
};

// ---------------------------------------------------------------- data layer (Supabase)
const rowToMember = (r) => ({ id: r.id, name: r.name, gameId: r.game_id || "", rank: r.rank, status: r.status, joinDate: r.join_date || "", leftDate: r.left_date || "", notes: r.notes || "" });
const memberToRow = (m) => ({ name: m.name, game_id: m.gameId || "", rank: m.rank, status: m.status, join_date: m.joinDate || null, left_date: m.leftDate || null, notes: m.notes || "" });
const rowToGrowth = (r) => ({ memberId: r.member_id, power: r.power ?? "", previousPower: r.previous_power ?? "", furnaceLevel: r.furnace_level || "", classes: r.classes || {}, updatedDate: r.updated_date || "" });
const growthToRow = (g) => ({ member_id: g.memberId, power: g.power === "" ? null : g.power, previous_power: g.previousPower === "" ? null : g.previousPower, furnace_level: g.furnaceLevel || "", classes: g.classes || {}, updated_date: g.updatedDate || null });
const rowToEvent = (r) => ({ id: r.id, date: r.date, type: r.type, name: r.name, session: r.session || "" });
const eventToRow = (e) => ({ date: e.date, type: e.type, name: e.name, session: e.session || "" });
const rowToPart = (r) => ({ id: r.id, eventId: r.event_id, memberId: r.member_id, signedUp: !!r.signed_up, attended: !!r.attended, partial: !!r.partial, score: r.score ?? "", note: r.note || "" });
const partToRow = (p) => ({ event_id: p.eventId, member_id: p.memberId, signed_up: !!p.signedUp, attended: !!p.attended, partial: !!p.partial, score: p.score === "" || p.score === undefined ? null : p.score, note: p.note || "" });

async function fetchAllData() {
  const [settingsRes, membersRes, growthRes, eventsRes, partRes] = await Promise.all([
    supabase.from("settings").select("*").eq("id", 1).single(),
    supabase.from("members").select("*").order("created_at", { ascending: true }),
    supabase.from("growth").select("*"),
    supabase.from("events").select("*").order("date", { ascending: false }),
    supabase.from("participation").select("*"),
  ]);
  const config = settingsRes.data
    ? { allianceName: settingsRes.data.alliance_name || "", leaderName: settingsRes.data.leader_name || "", inactivityDays: settingsRes.data.inactivity_days ?? 10, leaverRetentionDays: settingsRes.data.leaver_retention_days ?? 90 }
    : { allianceName: "", leaderName: "", inactivityDays: 10, leaverRetentionDays: 90 };
  return {
    config,
    members: (membersRes.data || []).map(rowToMember),
    growth: (growthRes.data || []).map(rowToGrowth),
    events: (eventsRes.data || []).map(rowToEvent),
    participation: (partRes.data || []).map(rowToPart),
  };
}


const STYLE = `
.wsc { --bg:#0B1420; --bg-elev:#0F1C2C; --panel:#132234; --panel-2:#18293C;
  --border:#24384C; --border-soft:#1B2C3E; --frost:#6FCBEA; --frost-dim:#3E6E85;
  --white:#E9F3F7; --steel:#8397AA; --steel-dim:#5C7086; --amber:#E8A33D;
  --danger:#E2604F; --success:#5FBF8C; --font-display: Arial, "Helvetica Neue", sans-serif;
  --font-body: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
  --font-mono: "SFMono-Regular", Consolas, "Liberation Mono", monospace;
  background:var(--bg); color:var(--white); font-family:var(--font-body);
  min-height:100vh; overflow:hidden; border:none; border-radius:0;
  display:flex; }
.wsc * { box-sizing:border-box; }
.wsc-side { width:200px; flex-shrink:0; background:var(--bg-elev); border-right:1px solid var(--border-soft);
  padding:18px 12px; display:flex; flex-direction:column; gap:4px; }
.wsc-brand { display:flex; align-items:center; gap:8px; padding:4px 8px 18px 8px; }
.wsc-brand-mark { width:26px; height:26px; border-radius:7px; background:linear-gradient(180deg,var(--frost),var(--frost-dim));
  display:flex; align-items:center; justify-content:center; flex-shrink:0; }
.wsc-brand-text { font-family:var(--font-display); font-weight:700; font-size:13px; letter-spacing:0.06em;
  text-transform:uppercase; line-height:1.2; }
.wsc-brand-sub { font-size:10px; color:var(--steel-dim); letter-spacing:0.04em; }
.wsc-brand-mobile { display:none; align-items:center; gap:6px; font-family:var(--font-display); font-weight:700;
  font-size:11px; color:var(--frost); letter-spacing:0.04em; text-transform:uppercase; margin-bottom:2px; }
.wsc-nav-btn { display:flex; align-items:center; gap:10px; padding:9px 10px; border-radius:8px; border:none;
  background:transparent; color:var(--steel); font-size:13px; font-weight:500; cursor:pointer; text-align:left;
  font-family:var(--font-body); transition:background .12s,color .12s; }
.wsc-nav-btn:hover { background:var(--panel-2); color:var(--white); }
.wsc-nav-btn.active { background:var(--panel-2); color:var(--frost); box-shadow:inset 2px 0 0 var(--frost); }
.wsc-nav-count { margin-left:auto; font-family:var(--font-mono); font-size:10.5px; color:var(--steel-dim); }
.wsc-main { flex:1; min-width:0; display:flex; flex-direction:column; }
.wsc-topbar { padding:14px 22px; border-bottom:1px solid var(--border-soft); display:flex; align-items:center;
  justify-content:space-between; gap:12px; flex-shrink:0; }
.wsc-title { font-family:var(--font-display); font-size:16px; font-weight:700; letter-spacing:0.02em; }
.wsc-title-sub { font-size:11px; color:var(--steel-dim); margin-top:2px; }
.wsc-body { flex:1; overflow-y:auto; padding:20px 22px 28px; }
.wsc-card { background:var(--panel); border:1px solid var(--border-soft); border-radius:11px; padding:16px 18px; }
.wsc-grid { display:grid; gap:12px; }
.wsc-stat-label { font-size:11px; color:var(--steel-dim); text-transform:uppercase; letter-spacing:0.06em; margin-bottom:6px; }
.wsc-stat-val { font-family:var(--font-mono); font-size:24px; font-weight:600; color:var(--white); }
.wsc-btn { display:inline-flex; align-items:center; gap:6px; padding:8px 13px; border-radius:8px;
  border:1px solid var(--border); background:var(--panel-2); color:var(--white); font-size:12.5px; font-weight:600;
  cursor:pointer; font-family:var(--font-body); transition:border-color .12s,background .12s; white-space:nowrap; }
.wsc-btn:hover { border-color:var(--frost-dim); }
.wsc-btn-primary { background:var(--frost); color:#08202C; border-color:var(--frost); }
.wsc-btn-primary:hover { background:#8AD6EE; }
.wsc-btn-danger { background:transparent; color:var(--danger); border-color:#4A2A26; }
.wsc-btn-danger:hover { background:#251616; border-color:var(--danger); }
.wsc-btn-sm { padding:5px 9px; font-size:11.5px; }
.wsc-btn-icon { padding:7px; }
.wsc-input, .wsc-select, .wsc-textarea { width:100%; background:var(--bg-elev); border:1px solid var(--border);
  color:var(--white); border-radius:7px; padding:8px 10px; font-size:13px; font-family:var(--font-body); }
.wsc-input:focus, .wsc-select:focus, .wsc-textarea:focus { outline:none; border-color:var(--frost-dim); }
.wsc-label { font-size:11px; color:var(--steel); margin-bottom:5px; display:block; font-weight:600;
  text-transform:uppercase; letter-spacing:0.04em; }
.wsc-field { margin-bottom:12px; }
.wsc-table { width:100%; border-collapse:collapse; font-size:13px; }
.wsc-table th { text-align:left; font-size:10.5px; text-transform:uppercase; letter-spacing:0.05em;
  color:var(--steel-dim); font-weight:600; padding:8px 10px; border-bottom:1px solid var(--border-soft); }
.wsc-table td { padding:9px 10px; border-bottom:1px solid var(--border-soft); vertical-align:middle; }
.wsc-table tr:last-child td { border-bottom:none; }
.wsc-table tr:hover td { background:var(--panel-2); }
.wsc-badge { display:inline-flex; align-items:center; gap:5px; padding:3px 9px; border-radius:20px;
  font-size:11px; font-weight:700; font-family:var(--font-mono); }
.wsc-modal-wrap { position:fixed; inset:0; background:rgba(4,10,16,0.72); display:flex; align-items:center;
  justify-content:center; z-index:50; padding:20px; }
.wsc-modal { background:var(--panel); border:1px solid var(--border); border-radius:13px; padding:22px;
  width:100%; max-width:460px; max-height:85vh; overflow-y:auto; }
.wsc-modal.wide { max-width:640px; }
.wsc-modal h3 { font-family:var(--font-display); font-size:15px; margin:0 0 16px; letter-spacing:0.02em; }
.wsc-empty { text-align:center; padding:40px 20px; color:var(--steel-dim); }
.wsc-empty-title { color:var(--steel); font-weight:600; margin-bottom:4px; font-size:14px; }
.wsc-search { display:flex; align-items:center; gap:8px; background:var(--bg-elev); border:1px solid var(--border);
  border-radius:8px; padding:0 10px; }
.wsc-search input { border:none; background:transparent; padding:8px 0; font-size:13px; color:var(--white); flex:1; }
.wsc-search input:focus { outline:none; }
.wsc-pill { font-size:10.5px; padding:2px 8px; border-radius:20px; font-weight:700; letter-spacing:0.02em; }
.wsc-checkbox { width:15px; height:15px; accent-color:var(--frost); cursor:pointer; }
.wsc-bulk-bar { display:flex; align-items:center; gap:10px; background:#6FCBEA14; border:1px solid var(--frost-dim);
  border-radius:8px; padding:8px 12px; margin-bottom:12px; font-size:12.5px; color:var(--frost); }
.wsc-class-block { border:1px solid var(--border-soft); border-radius:9px; padding:10px 12px; margin-bottom:10px; }
.wsc-class-title { font-size:12px; font-weight:700; color:var(--frost); margin-bottom:8px; }
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
    border:none; color:var(--steel); font-size:10px; font-family:var(--font-body); padding:4px 8px; cursor:pointer;
    position:relative; flex:1; }
  .wsc-bottomnav-btn.active { color:var(--frost); }
  .wsc-bottomnav-count { position:absolute; top:-2px; right:14px; background:var(--danger); color:#fff;
    font-size:9px; font-weight:700; border-radius:20px; padding:1px 4px; font-family:var(--font-mono); }
}
`;

function RankBadge({ rank }) {
  const colors = { R5: "#F2C94C", R4: "#6FCBEA", R3: "#8397AA", R2: "#5C7086", R1: "#3E4E5E" };
  const c = colors[rank] || "#5C7086";
  return <span className="wsc-badge" style={{ background: `${c}22`, color: c }}>{rank}</span>;
}
function StatusPill({ status }) {
  const map = { active: { c: "#5FBF8C", l: "Active" }, inactive: { c: "#E2604F", l: "Inactive" }, left: { c: "#8397AA", l: "Left" } };
  const s = map[status] || map.active;
  return <span className="wsc-pill" style={{ background: `${s.c}22`, color: s.c }}>{s.l}</span>;
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
  return <div className="wsc-empty"><div className="wsc-empty-title">{title}</div><div style={{ fontSize: 12.5 }}>{body}</div></div>;
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
    { id: "roster", label: "Roster", icon: Users },
    { id: "leavers", label: "Leavers", icon: UserX, count: leaverCount },
    { id: "growth", label: "Growth", icon: TrendingUp },
    { id: "events", label: "Events", icon: CalendarDays },
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
  const [threshold, setThreshold] = useState(config.inactivityDays ?? 10);
  const [retention, setRetention] = useState(config.leaverRetentionDays ?? 90);
  const [importMsg, setImportMsg] = useState("");
  const fileRef = useRef(null);

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
      <div className="wsc-field"><label className="wsc-label">Inactivity flag threshold (days)</label>
        <input className="wsc-input" type="number" min="1" value={threshold} onChange={(e) => setThreshold(Number(e.target.value))} />
        <div style={{ fontSize: 11.5, color: "var(--steel-dim)", marginTop: 5 }}>Active members with no logged activity in this many days show up as at-risk on the dashboard.</div></div>
      <div className="wsc-field"><label className="wsc-label">Leaver data retention (days)</label>
        <input className="wsc-input" type="number" min="1" value={retention} onChange={(e) => setRetention(Number(e.target.value))} />
        <div style={{ fontSize: 11.5, color: "var(--steel-dim)", marginTop: 5 }}>Members marked "left" stay on the Leavers tab this long, then are wiped on next load.</div></div>
      <div style={{ display: "flex", gap: 8, marginTop: 18, justifyContent: "flex-end" }}>
        <button className="wsc-btn" onClick={onClose}>Cancel</button>
        <button className="wsc-btn wsc-btn-primary" onClick={() => onSave({ allianceName: name, leaderName: leader, inactivityDays: threshold, leaverRetentionDays: retention })}><Save size={13} /> Save</button>
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
        <div style={{ fontSize: 11.5, color: "var(--steel-dim)" }}>
          Full data is every row, raw. Summary is a one-page digest — snapshot stats, who's reliable, who's not, troop spread, recent turnout — meant for sharing with officers who don't need the raw tables. The .json backup is a full, exact copy of everything here — use it to restore into this app later, or as a record if you ever move off Claude.
          {importMsg && <div style={{ color: "var(--frost)", marginTop: 6 }}>{importMsg}</div>}
        </div>
      </div>
    </Modal>
  );
}
function MemberModal({ member, onClose, onSave, onDelete }) {
  const [form, setForm] = useState(member || { name: "", gameId: "", rank: "R1", status: "active", joinDate: "", leftDate: "", notes: "" });
  const isEdit = !!member;
  const set = (k, v) => setForm((f) => ({ ...f, [k]: v }));
  const canSave = form.name.trim().length > 0;
  const onStatusChange = (status) => { if (status === "left" && !form.leftDate) set("leftDate", todayStr()); setForm((f) => ({ ...f, status })); };
  return (
    <Modal title={isEdit ? "Edit member" : "Add member"} onClose={onClose}>
      <div className="wsc-field"><label className="wsc-label">In-game name</label>
        <input className="wsc-input" value={form.name} onChange={(e) => set("name", e.target.value)} placeholder="Chief name" /></div>
      <div className="wsc-field"><label className="wsc-label">Game ID (optional)</label>
        <input className="wsc-input" value={form.gameId} onChange={(e) => set("gameId", e.target.value)} placeholder="123456789" /></div>
      <div style={{ display: "flex", gap: 10 }}>
        <div className="wsc-field" style={{ flex: 1 }}><label className="wsc-label">Rank</label>
          <select className="wsc-select" value={form.rank} onChange={(e) => set("rank", e.target.value)}>{RANKS.map((r) => <option key={r} value={r}>{r}</option>)}</select></div>
        <div className="wsc-field" style={{ flex: 1 }}><label className="wsc-label">Status</label>
          <select className="wsc-select" value={form.status} onChange={(e) => onStatusChange(e.target.value)}>
            <option value="active">Active</option><option value="inactive">Inactive</option><option value="left">Left alliance</option>
          </select></div>
      </div>
      <div style={{ display: "flex", gap: 10 }}>
        <div className="wsc-field" style={{ flex: 1 }}><label className="wsc-label">Join date</label>
          <input className="wsc-input" type="date" value={form.joinDate} onChange={(e) => set("joinDate", e.target.value)} />
          <div style={{ fontSize: 11, color: "var(--steel-dim)", marginTop: 4 }}>Blank by default — set it if you know it, otherwise tenure just shows "Unknown" instead of guessing.</div></div>
        {form.status === "left" && (
          <div className="wsc-field" style={{ flex: 1 }}><label className="wsc-label">Left date</label>
            <input className="wsc-input" type="date" value={form.leftDate} onChange={(e) => set("leftDate", e.target.value)} /></div>
        )}
      </div>
      <div className="wsc-field"><label className="wsc-label">Notes</label>
        <textarea className="wsc-textarea" rows={2} value={form.notes} onChange={(e) => set("notes", e.target.value)} placeholder="Timezone, role, anything worth remembering" /></div>
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
      <div style={{ fontSize: 11.5, color: "var(--steel-dim)", marginBottom: 10 }}>They'll move to the Leavers tab and drop off active tracking.</div>
      <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
        <button className="wsc-btn" onClick={onClose}>Cancel</button>
        <button className="wsc-btn wsc-btn-primary" onClick={() => onConfirm(date)}><Save size={13} /> Confirm</button>
      </div>
    </Modal>
  );
}
function TierDistributionChart({ members, growth }) {
  const [cls, setCls] = useState("infantry");
  const byMember = useMemo(() => { const map = {}; growth.forEach((g) => { map[g.memberId] = g; }); return map; }, [growth]);
  const chartData = useMemo(() => {
    const tierBuckets = TROOP_TIERS.filter(Boolean).map((t) => ({ label: t, count: 0 }));
    let fcCount = 0, untracked = 0;
    members.forEach((m) => {
      const c = byMember[m.id]?.classes?.[cls];
      if (!c || (!c.troopTier && !c.fcTroopLevel)) { untracked++; return; }
      if (c.fcTroopLevel) { fcCount++; return; }
      const bucket = tierBuckets.find((b) => b.label === c.troopTier);
      if (bucket) bucket.count++;
    });
    return [...tierBuckets, { label: "FC troops", count: fcCount }, { label: "Not set", count: untracked }];
  }, [members, byMember, cls]);
  return (
    <div className="wsc-card">
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10, flexWrap: "wrap", gap: 8 }}>
        <div className="wsc-stat-label" style={{ margin: 0 }}>Troop tier spread</div>
        <select className="wsc-select" style={{ width: 140 }} value={cls} onChange={(e) => setCls(e.target.value)}>
          {CLASSES.map((c) => <option key={c.key} value={c.key}>{c.label}</option>)}
        </select>
      </div>
      {members.length === 0 ? <EmptyState title="No members yet" body="Add members to see their tier spread here." /> : (
        <ResponsiveContainer width="100%" height={190}>
          <BarChart data={chartData} margin={{ top: 4, right: 8, left: -18, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#1B2C3E" />
            <XAxis dataKey="label" tick={{ fill: "#8397AA", fontSize: 10 }} axisLine={{ stroke: "#24384C" }} tickLine={false} interval={0} angle={-35} textAnchor="end" height={50} />
            <YAxis tick={{ fill: "#8397AA", fontSize: 11 }} axisLine={{ stroke: "#24384C" }} tickLine={false} allowDecimals={false} />
            <Tooltip contentStyle={{ background: "#132234", border: "1px solid #24384C", borderRadius: 8, fontSize: 12 }} labelStyle={{ color: "#E9F3F7" }} />
            <Bar dataKey="count" radius={[4, 4, 0, 0]} fill="#6FCBEA" />
          </BarChart>
        </ResponsiveContainer>
      )}
    </div>
  );
}

function Dashboard({ members, growth, events, participation, config }) {
  const threshold = config.inactivityDays ?? 10;
  const activeMembers = members.filter((m) => m.status !== "left");
  const lastActivityByMember = useMemo(() => {
    const map = {};
    activeMembers.forEach((m) => { map[m.id] = null; });
    const consider = (memberId, date) => { if (!date) return; if (!map[memberId] || date > map[memberId]) map[memberId] = date; };
    growth.forEach((g) => consider(g.memberId, g.updatedDate));
    participation.forEach((p) => { if (!p.attended) return; const ev = events.find((e) => e.id === p.eventId); if (ev) consider(p.memberId, ev.date); });
    return map;
  }, [activeMembers, growth, participation, events]);
  const atRisk = useMemo(() => activeMembers.filter((m) => m.status === "active")
    .map((m) => ({ m, days: daysAgo(lastActivityByMember[m.id]) })).filter((x) => x.days >= threshold).sort((a, b) => b.days - a.days), [activeMembers, lastActivityByMember, threshold]);
  const noShows = useMemo(() => {
    const counts = {};
    participation.forEach((p) => { if (p.signedUp && !p.attended) counts[p.memberId] = (counts[p.memberId] || 0) + 1; });
    return Object.entries(counts).map(([memberId, count]) => ({ member: activeMembers.find((m) => m.id === memberId), count }))
      .filter((r) => r.member && r.count >= 2).sort((a, b) => b.count - a.count).slice(0, 8);
  }, [participation, activeMembers]);
  const activeCount = activeMembers.filter((m) => m.status === "active").length;
  const activityRatio = activeCount > 0 ? activeMembers.filter((m) => m.status === "active").filter((m) => daysAgo(lastActivityByMember[m.id]) < threshold).length / activeCount : 0;
  const latestEvent = events.length ? [...events].sort((a, b) => b.date.localeCompare(a.date))[0] : null;
  const latestEventAttendance = latestEvent && activeCount > 0 ? participation.filter((p) => p.eventId === latestEvent.id && p.attended).length / activeCount : null;
  const readiness = Math.round(latestEventAttendance !== null ? (activityRatio * 0.6 + latestEventAttendance * 0.4) * 100 : activityRatio * 100);
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

  return (
    <div>
      <div className="wsc-grid" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(150px,1fr))", marginBottom: 14 }}>
        <div className="wsc-card"><div className="wsc-stat-label">Total members</div><div className="wsc-stat-val">{activeMembers.length}</div></div>
        <div className="wsc-card"><div className="wsc-stat-label">Active</div><div className="wsc-stat-val" style={{ color: "var(--success)" }}>{activeCount}</div></div>
        <div className="wsc-card"><div className="wsc-stat-label">At-risk</div><div className="wsc-stat-val" style={{ color: atRisk.length > 0 ? "var(--danger)" : "var(--white)" }}>{atRisk.length}</div></div>
        <div className="wsc-card"><div className="wsc-stat-label">Avg turnout</div><div className="wsc-stat-val">{events.length > 0 ? `${avgTurnout}%` : "—"}</div></div>
      </div>
      <div className="wsc-grid-hero">
        <div className="wsc-card" style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 6 }}>
          <ReadinessGauge score={isNaN(readiness) ? 0 : readiness} />
          <div style={{ fontSize: 11, color: "var(--steel-dim)", textAlign: "center" }}>Blends recent-activity rate {latestEvent ? "and last event turnout" : "(log an event to add turnout)"}.</div>
        </div>
        <TierDistributionChart members={activeMembers} growth={growth} />
      </div>
      <div className="wsc-grid-pair">
        <div className="wsc-card">
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}><TriangleAlert size={15} color="var(--amber)" /><div className="wsc-stat-label" style={{ margin: 0 }}>Members to check on</div></div>
          {atRisk.length === 0 ? <EmptyState title="Everyone's accounted for" body="No active member has gone quiet past your threshold." /> : (
            <table className="wsc-table"><thead><tr><th>Member</th><th>Days silent</th></tr></thead>
              <tbody>{atRisk.slice(0, 10).map(({ m, days }) => <tr key={m.id}><td>{m.name}</td><td style={{ color: "var(--danger)", fontFamily: "var(--font-mono)" }}>{isFinite(days) ? days : "—"}</td></tr>)}</tbody></table>
          )}
        </div>
        <div className="wsc-card">
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}><Ban size={15} color="var(--danger)" /><div className="wsc-stat-label" style={{ margin: 0 }}>Frequent no-shows</div></div>
          {noShows.length === 0 ? <EmptyState title="No repeat no-shows" body="Members who sign up but don't attend twice or more will show here." /> : (
            <table className="wsc-table"><thead><tr><th>Member</th><th>No-shows</th></tr></thead>
              <tbody>{noShows.map(({ member, count }) => <tr key={member.id}><td>{member.name}</td><td style={{ color: "var(--danger)", fontFamily: "var(--font-mono)" }}>{count}</td></tr>)}</tbody></table>
          )}
        </div>
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
function RosterTab({ members, lastActivityByMember, onEdit, onBulkLeave }) {
  const [query, setQuery] = useState("");
  const [rankFilter, setRankFilter] = useState("all");
  const [statusFilter, setStatusFilter] = useState("all");
  const [selected, setSelected] = useState(new Set());
  const [showBulk, setShowBulk] = useState(false);
  const roster = members.filter((m) => m.status !== "left");
  const filtered = roster.filter((m) => {
    const matchesQuery = m.name.toLowerCase().includes(query.toLowerCase()) || (m.gameId || "").includes(query);
    const matchesRank = rankFilter === "all" || m.rank === rankFilter;
    const matchesStatus = statusFilter === "all" || m.status === statusFilter;
    return matchesQuery && matchesRank && matchesStatus;
  });
  const toggleOne = (id) => setSelected((s) => { const n = new Set(s); n.has(id) ? n.delete(id) : n.add(id); return n; });
  const toggleAll = () => setSelected((s) => s.size === filtered.length ? new Set() : new Set(filtered.map((m) => m.id)));
  return (
    <div>
      <div style={{ display: "flex", gap: 10, marginBottom: 14, flexWrap: "wrap" }}>
        <div className="wsc-search" style={{ flex: 1, minWidth: 180 }}><Search size={14} color="var(--steel-dim)" /><input placeholder="Search name or ID" value={query} onChange={(e) => setQuery(e.target.value)} /></div>
        <select className="wsc-select" style={{ width: 130 }} value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}>
          <option value="all">All statuses</option><option value="active">Active</option><option value="inactive">Inactive</option>
        </select>
        <select className="wsc-select" style={{ width: 120 }} value={rankFilter} onChange={(e) => setRankFilter(e.target.value)}>
          <option value="all">All ranks</option>{RANKS.map((r) => <option key={r} value={r}>{r}</option>)}
        </select>
      </div>
      {selected.size > 0 && (
        <div className="wsc-bulk-bar">
          <span>{selected.size} selected</span>
          <button className="wsc-btn wsc-btn-sm" onClick={() => setShowBulk(true)}><UserX size={12} /> Mark as left</button>
          <button className="wsc-btn wsc-btn-sm" onClick={() => setSelected(new Set())}>Clear</button>
        </div>
      )}
      <div className="wsc-card" style={{ padding: 0, overflow: "hidden" }}>
        {roster.length === 0 ? <EmptyState title="No members yet" body='Add your first chief with "Add member" above to start the roster.' />
          : filtered.length === 0 ? <EmptyState title="No matches" body="Try a different search or filter." /> : (
          <div style={{ overflowX: "auto" }}>
            <table className="wsc-table">
              <thead><tr><th style={{ width: 30 }}><input type="checkbox" className="wsc-checkbox" checked={selected.size === filtered.length && filtered.length > 0} onChange={toggleAll} /></th>
                <th>Name</th><th>Rank</th><th>Status</th><th>Joined</th><th>Tenure</th><th>Last activity</th><th></th></tr></thead>
              <tbody>
                {filtered.map((m) => {
                  const last = lastActivityByMember[m.id];
                  return (
                    <tr key={m.id}>
                      <td onClick={(e) => e.stopPropagation()}><input type="checkbox" className="wsc-checkbox" checked={selected.has(m.id)} onChange={() => toggleOne(m.id)} /></td>
                      <td style={{ fontWeight: 600, cursor: "pointer" }} onClick={() => onEdit(m)}>{m.name}</td>
                      <td><RankBadge rank={m.rank} /></td>
                      <td><StatusPill status={m.status} /></td>
                      <td style={{ color: "var(--steel)" }}>{m.joinDate ? fmtDate(m.joinDate) : "Unknown"}</td>
                      <td style={{ color: "var(--steel)", fontFamily: "var(--font-mono)" }}><TenureLabel joinDate={m.joinDate} /></td>
                      <td style={{ color: "var(--steel)" }}>{last ? fmtDate(last) : "No data"}</td>
                      <td style={{ textAlign: "right", cursor: "pointer" }} onClick={() => onEdit(m)}><Pencil size={13} color="var(--steel-dim)" /></td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>
      {showBulk && <BulkLeaveModal count={selected.size} onClose={() => setShowBulk(false)}
        onConfirm={(date) => { onBulkLeave([...selected], date); setSelected(new Set()); setShowBulk(false); }} />}
    </div>
  );
}
function LeaversTab({ members, retentionDays, onReactivate, onPurgeNow, onEdit }) {
  const leavers = members.filter((m) => m.status === "left").sort((a, b) => (b.leftDate || "").localeCompare(a.leftDate || ""));
  return (
    <div>
      <div className="wsc-card" style={{ marginBottom: 14, fontSize: 12, color: "var(--steel)" }}>
        Members marked "Left alliance" land here. Their history stays for {retentionDays} days after leaving, then is wiped automatically the next time this app is opened.
      </div>
      <div className="wsc-card" style={{ padding: 0, overflow: "hidden" }}>
        {leavers.length === 0 ? <EmptyState title="No former members on file" body='Mark a member as "Left alliance" from the Roster tab to move them here.' /> : (
          <div style={{ overflowX: "auto" }}>
            <table className="wsc-table">
              <thead><tr><th>Name</th><th>Game ID</th><th>Rank</th><th>Time in alliance</th><th>Left on</th><th>Purges in</th><th></th></tr></thead>
              <tbody>
                {leavers.map((m) => {
                  const elapsed = daysAgo(m.leftDate);
                  const remaining = Math.max(0, retentionDays - elapsed);
                  return (
                    <tr key={m.id}>
                      <td style={{ fontWeight: 600, cursor: "pointer" }} onClick={() => onEdit(m)}>{m.name}</td>
                      <td style={{ color: "var(--steel-dim)" }}>{m.gameId || "—"}</td>
                      <td><RankBadge rank={m.rank} /></td>
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
function GrowthTab({ members, growth, onEditMember }) {
  const byMember = useMemo(() => { const map = {}; growth.forEach((g) => { map[g.memberId] = g; }); return map; }, [growth]);
  return (
    <div>
      {members.length === 0 ? <div className="wsc-card"><EmptyState title="Add members first" body="Growth tracking needs a roster — head to the Roster tab." /></div> : (
        <div className="wsc-card" style={{ padding: 0, overflow: "hidden" }}>
          <div style={{ padding: "10px 18px 0", fontSize: 11.5, color: "var(--steel-dim)" }}>Click a row to update that member's profile.</div>
          <div style={{ overflowX: "auto" }}>
            <table className="wsc-table">
              <thead><tr><th>Member</th><th>Power</th><th>Change</th><th>Furnace</th><th>Infantry</th><th>Marksman</th><th>Lancer</th><th>Updated</th></tr></thead>
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
                      <td>{classSummary(g?.classes?.infantry)}</td>
                      <td>{classSummary(g?.classes?.marksman)}</td>
                      <td>{classSummary(g?.classes?.lancer)}</td>
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
      <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
        <button className="wsc-btn" onClick={onClose}>Cancel</button>
        <button className="wsc-btn wsc-btn-primary" disabled={!canSave} style={{ opacity: canSave ? 1 : 0.5 }}
          onClick={() => canSave && onSave({ id: uid(), name: type === "Custom" ? name : type, type, session, date })}><Save size={13} /> Create event</button>
      </div>
    </Modal>
  );
}
function EventDetail({ event, members, participation, onClose, onToggleSignUp, onToggleAttend, onTogglePartial, onScore, onNote }) {
  const [query, setQuery] = useState("");
  const filtered = members.filter((m) => m.status !== "left" && m.name.toLowerCase().includes(query.toLowerCase()));
  const partMap = {};
  participation.filter((p) => p.eventId === event.id).forEach((p) => { partMap[p.memberId] = p; });
  const attendedCount = Object.values(partMap).filter((p) => p.attended).length;
  const signedCount = Object.values(partMap).filter((p) => p.signedUp).length;
  return (
    <Modal title={`${event.name}${event.session ? ` · ${event.session}` : ""} — ${fmtDate(event.date)}`} onClose={onClose} wide>
      <div style={{ fontSize: 12, color: "var(--steel-dim)", marginBottom: 10 }}>{signedCount} signed up · {attendedCount} attended</div>
      <div className="wsc-search" style={{ marginBottom: 10 }}><Search size={13} color="var(--steel-dim)" /><input placeholder="Search member" value={query} onChange={(e) => setQuery(e.target.value)} /></div>
      <div className="wsc-scroll" style={{ maxHeight: 380, overflowY: "auto", overflowX: "auto" }}>
        <table className="wsc-table" style={{ minWidth: 620 }}>
          <thead><tr><th>Member</th><th>Signed up</th><th>Attended</th><th>Full duration?</th><th>Score</th><th>Notes</th></tr></thead>
          <tbody>
            {filtered.map((m) => {
              const p = partMap[m.id];
              const noShow = p?.signedUp && !p?.attended;
              return (
                <tr key={m.id}>
                  <td>{m.name}{noShow && <span className="wsc-pill" style={{ background: "#E2604F22", color: "var(--danger)", marginLeft: 8 }}>No-show</span>}</td>
                  <td><button className="wsc-btn wsc-btn-icon" style={{ background: p?.signedUp ? "#6FCBEA22" : "transparent", borderColor: p?.signedUp ? "var(--frost)" : "var(--border)" }}
                    onClick={() => onToggleSignUp(event.id, m.id, !p?.signedUp)} aria-label="Toggle signed up"><Check size={13} color={p?.signedUp ? "var(--frost)" : "var(--steel-dim)"} /></button></td>
                  <td><button className="wsc-btn wsc-btn-icon" style={{ background: p?.attended ? "#5FBF8C22" : "transparent", borderColor: p?.attended ? "var(--success)" : "var(--border)" }}
                    onClick={() => onToggleAttend(event.id, m.id, !p?.attended)} aria-label="Toggle attended"><Check size={13} color={p?.attended ? "var(--success)" : "var(--steel-dim)"} /></button></td>
                  <td>
                    {p?.attended ? (
                      <button className="wsc-btn wsc-btn-sm" style={{ background: p?.partial ? "#E8A33D22" : "#5FBF8C22", borderColor: p?.partial ? "var(--amber)" : "var(--success)", color: p?.partial ? "var(--amber)" : "var(--success)" }}
                        onClick={() => onTogglePartial(event.id, m.id, !p?.partial)}>{p?.partial ? "Left early" : "Full"}</button>
                    ) : <span style={{ color: "var(--steel-dim)" }}>—</span>}
                  </td>
                  <td><input className="wsc-input" style={{ width: 90 }} type="number" placeholder="—" defaultValue={p?.score ?? ""} onBlur={(e) => onScore(event.id, m.id, e.target.value)} /></td>
                  <td><input className="wsc-input" style={{ width: 150 }} placeholder="e.g. left after 20 min" defaultValue={p?.note ?? ""} onBlur={(e) => onNote(event.id, m.id, e.target.value)} /></td>
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
        <span style={{ fontSize: 11.5, color: "var(--steel-dim)" }}>Show</span>
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
                      <td style={{ fontWeight: 600 }}>{ev.name}</td>
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
export default function App() {
  const [session, setSession] = useState(undefined); // undefined = checking, null = signed out
  const [loading, setLoading] = useState(true);
  const [tab, setTab] = useState("dashboard");
  const [config, setConfig] = useState({ allianceName: "", leaderName: "", inactivityDays: 10, leaverRetentionDays: 90 });
  const [members, setMembers] = useState([]);
  const [growth, setGrowth] = useState([]);
  const [events, setEvents] = useState([]);
  const [participation, setParticipation] = useState([]);

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
    if (!session) return;
    (async () => {
      setLoading(true);
      const { config: cfg, members: mem, growth: gr, events: ev, participation: part } = await fetchAllData();

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
      setConfig(cfg); setMembers(finalMembers); setGrowth(finalGrowth); setEvents(ev); setParticipation(finalPart);
      setLoading(false);
    })();
  }, [session]);

  const lastActivityByMember = useMemo(() => {
    const map = {};
    members.forEach((m) => { map[m.id] = null; });
    const consider = (memberId, date) => { if (!date) return; if (!map[memberId] || date > map[memberId]) map[memberId] = date; };
    growth.forEach((g) => consider(g.memberId, g.updatedDate));
    participation.forEach((p) => { if (!p.attended) return; const ev = events.find((e) => e.id === p.eventId); if (ev) consider(p.memberId, ev.date); });
    return map;
  }, [members, growth, participation, events]);

  const saveConfig = useCallback(async (next) => {
    setConfig(next); setShowConfig(false);
    await supabase.from("settings").update({
      alliance_name: next.allianceName, leader_name: next.leaderName,
      inactivity_days: next.inactivityDays, leaver_retention_days: next.leaverRetentionDays,
    }).eq("id", 1);
  }, []);

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

  const roster = members.filter((m) => m.status !== "left");
  const leaverCount = members.filter((m) => m.status === "left").length;

  const exportExcel = useCallback(() => {
    const wsMembers = XLSX.utils.json_to_sheet(members.map((m) => ({
      Name: m.name, "Game ID": m.gameId || "", Rank: m.rank, Status: m.status,
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
    const wsEvents = XLSX.utils.json_to_sheet(events.map((e) => ({ Date: e.date, Type: e.type, Name: e.name, Session: e.session || "" })));
    const partRows = participation.map((p) => {
      const ev = events.find((e) => e.id === p.eventId), m = members.find((x) => x.id === p.memberId);
      return { Event: ev ? ev.name : p.eventId, Date: ev ? ev.date : "", Member: m ? m.name : p.memberId, "Signed up": p.signedUp ? "Yes" : "No", Attended: p.attended ? "Yes" : "No", Score: p.score };
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
    const threshold = config.inactivityDays ?? 10;
    const activeMembers = members.filter((m) => m.status !== "left");
    const activeOnly = activeMembers.filter((m) => m.status === "active");
    const growthByMember = {}; growth.forEach((g) => { growthByMember[g.memberId] = g; });
    const lastActivity = {};
    activeMembers.forEach((m) => { lastActivity[m.id] = null; });
    const consider = (id, d) => { if (!d) return; if (!lastActivity[id] || d > lastActivity[id]) lastActivity[id] = d; };
    growth.forEach((g) => consider(g.memberId, g.updatedDate));
    participation.forEach((p) => { if (!p.attended) return; const ev = events.find((e) => e.id === p.eventId); if (ev) consider(p.memberId, ev.date); });
    const atRisk = activeOnly.map((m) => ({ name: m.name, days: daysAgo(lastActivity[m.id]) })).filter((x) => x.days >= threshold).sort((a, b) => b.days - a.days);

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
      const rate = activeOnly.length > 0 ? Math.round((attended / activeOnly.length) * 100) : 0;
      return { Date: ev.date, Event: ev.name + (ev.session ? ` (${ev.session})` : ""), "Signed up": rows.filter((p) => p.signedUp).length, Attended: attended, "Turnout %": rate };
    });

    const aoa = [
      [`${config.allianceName || "Alliance"} — Summary Report`],
      [`Generated ${todayStr()}${config.leaderName ? ` · Led by ${config.leaderName}` : ""}`],
      [],
      ["Alliance snapshot"],
      ["Total members", activeMembers.length],
      ["Active", activeOnly.length],
      ["Inactive", activeMembers.filter((m) => m.status === "inactive").length],
      ["At-risk (quiet " + threshold + "+ days)", atRisk.length],
      ["Total events logged", events.length],
      [],
      ["Members to check on (days silent)"],
      ...(atRisk.length ? atRisk.slice(0, 15).map((r) => [r.name, r.days]) : [["None — everyone's accounted for"]]),
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
    return <div className="wsc" style={{ alignItems: "center", justifyContent: "center", minHeight: 300 }}><style>{STYLE}</style><div style={{ color: "var(--steel)", fontSize: 13 }}>Loading…</div></div>;
  }
  if (!session) return <Login />;
  if (loading) return <div className="wsc" style={{ alignItems: "center", justifyContent: "center", minHeight: 300 }}><style>{STYLE}</style><div style={{ color: "var(--steel)", fontSize: 13 }}>Loading alliance data…</div></div>;

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
              {tab === "growth" && "Growth"}{tab === "events" && "Events"}
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
          {tab === "roster" && <RosterTab members={members} lastActivityByMember={lastActivityByMember} onEdit={(m) => setMemberModal(m)} onBulkLeave={bulkMarkLeft} />}
          {tab === "leavers" && <LeaversTab members={members} retentionDays={config.leaverRetentionDays ?? 90} onReactivate={reactivateMember} onPurgeNow={deleteMember} onEdit={(m) => setMemberModal(m)} />}
          {tab === "growth" && <GrowthTab members={roster} growth={growth} onEditMember={openGrowthFor} />}
          {tab === "events" && <EventsTab events={events} members={members} participation={participation} onOpenEvent={(ev) => setOpenEvent(ev)} />}
        </div>
      </div>
      <BottomNav tab={tab} setTab={setTab} leaverCount={leaverCount} />
      {showConfig && <ConfigModal config={config} onClose={() => setShowConfig(false)} onSave={saveConfig}
        onExportExcel={exportExcel} onExportSummary={exportSummary} onExportBackup={exportBackup} onImportBackup={importBackup} />}
      {(showAddMember || memberModal) && <MemberModal member={memberModal} onClose={() => { setShowAddMember(false); setMemberModal(null); }} onSave={saveMember} onDelete={deleteMember} />}
      {showLogGrowth && <LogGrowthModal members={roster} profiles={growth} initialMemberId={growthPreset} onClose={() => { setShowLogGrowth(false); setGrowthPreset(null); }} onSave={saveGrowth} />}
      {showAddEvent && <EventModal onClose={() => setShowAddEvent(false)} onSave={addEvent} />}
      {openEvent && <EventDetail event={openEvent} members={members} participation={participation} onClose={() => setOpenEvent(null)} onToggleSignUp={toggleSignUp} onToggleAttend={toggleAttend} onTogglePartial={togglePartial} onScore={setScore} onNote={setNote} />}
    </div>
  );
}

import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

// ─── DONNÉES INITIALES ────────────────────────────────────────────────────────

const INITIAL_USERS = [
  { id: "admin", username: "coach", password: "coach123", role: "admin", name: "La Coach" },
  { id: "p1", username: "joueur1", password: "poker123", role: "player", name: "Alex Tremblay" },
  { id: "p2", username: "joueur2", password: "poker123", role: "player", name: "Marie Gagnon" },
];

const DAYS = ["Dim", "Lun", "Mar", "Mer", "Jeu", "Ven", "Sam"];
const DAYS_FULL = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi"];

const INITIAL_TOURNAMENTS = [
  { id: "t1", site: "PokerStars", name: "Sunday Million", buyIn: 215, regTime: "17:45", startTime: "18:00", days: [0], notes: "Gros champ fréquentiel" },
  { id: "t2", site: "GGPoker", name: "WSOP Daily", buyIn: 109, regTime: "19:50", startTime: "20:00", days: [1,2,3,4,5], notes: "" },
  { id: "t3", site: "PokerStars", name: "Turbo Series 6-Max", buyIn: 55, regTime: "20:55", startTime: "21:00", days: [1,3,5], notes: "Turbo, structure rapide" },
  { id: "t4", site: "888poker", name: "Super XL Special", buyIn: 33, regTime: "21:25", startTime: "21:30", days: [2,4,6], notes: "" },
  { id: "t5", site: "GGPoker", name: "GGMasters HR", buyIn: 525, regTime: "17:55", startTime: "18:00", days: [0], notes: "High Roller dimanche" },
  { id: "t6", site: "PokerStars", name: "Daily $109", buyIn: 109, regTime: "14:55", startTime: "15:00", days: [0,1,2,3,4,5,6], notes: "" },
];

// ─── STORAGE HELPERS ──────────────────────────────────────────────────────────

const store = {
  get: (k, def) => { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : def; } catch { return def; } },
  set: (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch {} },
};

// ─── SUPABASE CONFIG ──────────────────────────────────────────────────────────
// Étape 1 : Coller tes clés Supabase ci-dessous
// (Supabase → Settings → API)
const SUPABASE_URL  = "https://mchjihofmuyhkktvfxnw.supabase.co";
const SUPABASE_ANON = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im1jaGppaG9mbXV5aGtrdHZmeG53Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzIyMjE2NDUsImV4cCI6MjA4Nzc5NzY0NX0.Y2yNFTHmfIUGh43SWY_aaf32W-lDCy8Rc3xBU4bAAdo";

// ─── CLIENT SUPABASE LÉGER (REST + Realtime WebSocket) ────────────────────────
const sbH = () => ({
  "apikey": SUPABASE_ANON,
  "Authorization": "Bearer " + SUPABASE_ANON,
  "Content-Type": "application/json",
});

const sb = {
  async select(table, filter = "") {
    try {
      const url = `${SUPABASE_URL}/rest/v1/${table}?${filter}&order=created_at.asc`;
      const r = await fetch(url, { headers: sbH() });
      return r.ok ? r.json() : [];
    } catch { return []; }
  },
  async upsert(table, rows) {
    try {
      await fetch(`${SUPABASE_URL}/rest/v1/${table}`, {
        method: "POST",
        headers: { ...sbH(), "Prefer": "resolution=merge-duplicates" },
        body: JSON.stringify(Array.isArray(rows) ? rows : [rows]),
      });
    } catch {}
  },
  async remove(table, col, val) {
    try {
      await fetch(`${SUPABASE_URL}/rest/v1/${table}?${col}=eq.${encodeURIComponent(val)}`, {
        method: "DELETE", headers: sbH(),
      });
    } catch {}
  },
  realtime(table, cb) {
    const wsUrl = SUPABASE_URL.replace("https://", "wss://")
      + "/realtime/v1/websocket?apikey=" + SUPABASE_ANON + "&vsn=1.0.0";
    let ws, dead = false;
    const connect = () => {
      if (dead) return;
      ws = new WebSocket(wsUrl);
      ws.onopen = () => ws.send(JSON.stringify({
        topic: `realtime:public:${table}`, event: "phx_join",
        payload: { config: { broadcast: { ack: false }, presence: { key: "" } } }, ref: "1"
      }));
      ws.onmessage = e => {
        const m = JSON.parse(e.data);
        if (["INSERT","UPDATE","DELETE"].includes(m.event)) cb(m.event, m.payload?.record, m.payload?.old_record);
      };
      ws.onclose = () => { if (!dead) setTimeout(connect, 3000); }; // reconnect auto
      ws.onerror = () => {};
    };
    connect();
    return () => { dead = true; ws?.close(); };
  },
};

// ─── NORMALISEURS snake_case ↔ camelCase ──────────────────────────────────────
const normT = t => ({ id: t.id, site: t.site, name: t.name, buyIn: +t.buy_in || 0, regTime: t.reg_time, startTime: t.start_time || "", days: t.days || [], notes: t.notes || "" });
const normS = s => ({ key: s.key, date: s.date, tournamentId: s.tournament_id, tournamentName: s.tournament_name, site: s.site, buyIn: +s.buy_in || 0, rebuys: s.rebuys || 0, won: s.won ?? null, userId: s.user_id });
const dbT   = t => ({ id: t.id, site: t.site, name: t.name, buy_in: t.buyIn, reg_time: t.regTime, start_time: t.startTime, days: t.days, notes: t.notes });
const dbS   = s => ({ key: s.key, date: s.date, tournament_id: s.tournamentId, tournament_name: s.tournamentName, site: s.site, buy_in: s.buyIn, rebuys: s.rebuys || 0, won: s.won, user_id: s.userId });


// ─── SOUND (contexte audio partagé — supporte alarmes simultanées) ───────────

let _sharedAudioCtx = null;
function getAudioCtx() {
  try {
    if (!_sharedAudioCtx || _sharedAudioCtx.state === "closed") {
      _sharedAudioCtx = new (window.AudioContext || window.webkitAudioContext)();
    }
    if (_sharedAudioCtx.state === "suspended") _sharedAudioCtx.resume();
    return _sharedAudioCtx;
  } catch { return null; }
}

// Chaque appel s'ajoute au même contexte partagé → plusieurs tournois
// peuvent sonner en même temps sans se couper mutuellement
function playBeep(freq = 880, duration = 0.3, vol = 0.4, count = 1, offset = 0) {
  const ctx = getAudioCtx();
  if (!ctx) return;
  for (let i = 0; i < count; i++) {
    try {
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.connect(gain);
      gain.connect(ctx.destination);
      osc.frequency.value = freq;
      osc.type = "sine";
      const t = ctx.currentTime + offset + i * 0.4;
      gain.gain.setValueAtTime(vol, t);
      gain.gain.exponentialRampToValueAtTime(0.001, t + duration);
      osc.start(t);
      osc.stop(t + duration + 0.05);
    } catch {}
  }
}

// Alarme douce — carillon paisible (3 notes montantes, fondu lent)
function playAlarm() {
  const ctx = getAudioCtx();
  if (!ctx) return;
  const notes = [523.25, 659.25, 783.99]; // Do - Mi - Sol (accord majeur doux)
  notes.forEach((freq, i) => {
    try {
      const osc  = ctx.createOscillator();
      const gain = ctx.createGain();
      // Légère réverbération : on mélange un oscillateur sine + un triangle très doux
      const osc2  = ctx.createOscillator();
      const gain2 = ctx.createGain();
      osc.connect(gain);   gain.connect(ctx.destination);
      osc2.connect(gain2); gain2.connect(ctx.destination);
      osc.frequency.value  = freq;
      osc2.frequency.value = freq * 2; // octave haute très discrète
      osc.type  = "sine";
      osc2.type = "triangle";
      const t = ctx.currentTime + i * 0.55; // espacement généreux entre les notes
      // Attaque douce, longue résonance
      gain.gain.setValueAtTime(0, t);
      gain.gain.linearRampToValueAtTime(0.22, t + 0.08);
      gain.gain.exponentialRampToValueAtTime(0.001, t + 2.2);
      gain2.gain.setValueAtTime(0, t);
      gain2.gain.linearRampToValueAtTime(0.04, t + 0.06);
      gain2.gain.exponentialRampToValueAtTime(0.001, t + 1.8);
      osc.start(t);  osc.stop(t + 2.3);
      osc2.start(t); osc2.stop(t + 2.0);
    } catch {}
  });
}

// ─── UTILS ───────────────────────────────────────────────────────────────────

function getSecondsUntilReg(regTime) {
  const now = new Date();
  const [h, m] = regTime.split(":").map(Number);
  const target = new Date(now);
  target.setHours(h, m, 0, 0);
  if (target <= now) target.setDate(target.getDate() + 1);
  return Math.floor((target - now) / 1000);
}

function formatCountdown(secs) {
  if (secs < 0) return "EXPIRÉ";
  const h = Math.floor(secs / 3600);
  const m = Math.floor((secs % 3600) / 60);
  const s = secs % 60;
  if (h > 0) return `${h}h ${String(m).padStart(2,"0")}m ${String(s).padStart(2,"0")}s`;
  return `${String(m).padStart(2,"0")}:${String(s).padStart(2,"0")}`;
}

function todayTournaments(tournaments) {
  const dow = new Date().getDay();
  return tournaments
    .filter(t => t.days.includes(dow))
    .sort((a, b) => getSecondsUntilReg(a.regTime) - getSecondsUntilReg(b.regTime));
}

function getAlertLevel(secs) {
  if (secs < 0) return "expired";
  if (secs <= 300) return "critical";
  if (secs <= 900) return "warning";
  return "ok";
}

// ─── COMPOSANTS ──────────────────────────────────────────────────────────────

function LoginScreen({ onLogin }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError]       = useState("");
  const [shake, setShake]       = useState(false);
  const [logging, setLogging]   = useState(false);

  const isConfigured = SUPABASE_URL !== "COLLE_TON_URL_ICI";

  const handleSubmit = async () => {
    if (logging) return;
    setLogging(true);
    setError("");
    try {
      let user = null;
      if (isConfigured) {
        // Auth via Supabase users table
        const rows = await sb.select("users", `username=eq.${encodeURIComponent(username)}&password=eq.${encodeURIComponent(password)}`);
        user = rows?.[0] || null;
      } else {
        // Fallback local
        const users = store.get("pk_users", INITIAL_USERS);
        user = users.find(u => u.username === username && u.password === password);
      }
      if (user) {
        store.set("pk_session", user);
        onLogin(user);
      } else {
        setError("Identifiants incorrects");
        setShake(true);
        setTimeout(() => setShake(false), 600);
      }
    } catch {
      setError("Erreur de connexion");
    }
    setLogging(false);
  };

  return (
    <div style={{
      minHeight: "100vh", background: "#0a0a0f",
      display: "flex", alignItems: "center", justifyContent: "center",
      fontFamily: "'Courier New', monospace",
      backgroundImage: "radial-gradient(ellipse at 20% 50%, rgba(139,92,246,0.08) 0%, transparent 60%), radial-gradient(ellipse at 80% 20%, rgba(16,185,129,0.06) 0%, transparent 50%)"
    }}>
      <div style={{
        width: 380, padding: "48px 40px",
        background: "rgba(255,255,255,0.03)",
        border: "1px solid rgba(255,255,255,0.08)",
        borderRadius: 2,
        animation: shake ? "shake 0.5s ease" : "none"
      }}>
        <div style={{ textAlign: "center", marginBottom: 40 }}>
          <div style={{ fontSize: 40, marginBottom: 12 }}>♠</div>
          <div style={{ color: "#e2e8f0", fontSize: 22, fontWeight: 700, letterSpacing: 3, textTransform: "uppercase" }}>POKER COACH</div>
          <div style={{ color: "#64748b", fontSize: 11, letterSpacing: 4, marginTop: 6 }}>TOURNAMENT MANAGER</div>
        </div>
        <input
          placeholder="Nom d'utilisateur"
          value={username}
          onChange={e => setUsername(e.target.value)}
          onKeyDown={e => e.key === "Enter" && handleSubmit()}
          style={inputStyle}
          autoFocus
        />
        <input
          type="password"
          placeholder="Mot de passe"
          value={password}
          onChange={e => setPassword(e.target.value)}
          onKeyDown={e => e.key === "Enter" && handleSubmit()}
          style={{ ...inputStyle, marginTop: 12 }}
        />
        {error && <div style={{ color: "#f87171", fontSize: 12, marginTop: 8, textAlign: "center" }}>{error}</div>}
        <button onClick={handleSubmit} disabled={logging} style={{ ...btnPrimary, opacity: logging ? 0.6 : 1 }}>
          {logging ? "CONNEXION..." : "CONNEXION"}
        </button>
        <div style={{ color: "#334155", fontSize: 11, textAlign: "center", marginTop: 24 }}>
          coach / coach123 &nbsp;·&nbsp; joueur1 / poker123
        </div>
      </div>
      <style>{`
        @keyframes shake { 0%,100%{transform:translateX(0)} 25%{transform:translateX(-8px)} 75%{transform:translateX(8px)} }
        @keyframes fadeIn { from{opacity:0;transform:translateY(12px)} to{opacity:1;transform:translateY(0)} }
        @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.5} }
        * { box-sizing: border-box; }
        ::-webkit-scrollbar { width: 4px; } ::-webkit-scrollbar-track { background: #0a0a0f; } ::-webkit-scrollbar-thumb { background: #1e293b; }
      `}</style>
    </div>
  );
}

const inputStyle = {
  width: "100%", padding: "12px 16px",
  background: "rgba(255,255,255,0.04)", border: "1px solid rgba(255,255,255,0.1)",
  borderRadius: 2, color: "#e2e8f0", fontSize: 14,
  fontFamily: "'Courier New', monospace", outline: "none",
  transition: "border-color 0.2s",
};

const btnPrimary = {
  width: "100%", marginTop: 20, padding: "14px",
  background: "linear-gradient(135deg, #10b981, #059669)",
  border: "none", borderRadius: 2, color: "#fff",
  fontSize: 13, fontWeight: 700, letterSpacing: 3, textTransform: "uppercase",
  cursor: "pointer", fontFamily: "'Courier New', monospace",
};

// ─── COUNTDOWN CELL ──────────────────────────────────────────────────────────

function CountdownCell({ regTime, soundEnabled, tournamentId, onFire }) {
  const [secs, setSecs] = useState(() => getSecondsUntilReg(regTime));
  const firedRef = useRef(false);

  useEffect(() => {
    firedRef.current = false; // reset si regTime change (ex: test)
  }, [regTime]);

  useEffect(() => {
    const tick = () => {
      const s = getSecondsUntilReg(regTime);
      setSecs(s);
      if (s <= 0 && !firedRef.current) {
        firedRef.current = true;
        if (soundEnabled) playAlarm();
        onFire && onFire(tournamentId);
      }
    };
    tick();
    const id = setInterval(tick, 1000);
    return () => clearInterval(id);
  }, [regTime, soundEnabled, tournamentId, onFire]);

  const expired = secs <= 0;
  const critical = secs <= 30 && secs > 0;
  const warning  = secs <= 120 && secs > 30;
  const color = expired ? "#334155" : critical ? "#ef4444" : warning ? "#f59e0b" : "#10b981";

  return (
    <span style={{
      color, fontWeight: 700, fontSize: 15,
      fontFamily: "'Courier New', monospace",
      animation: critical ? "pulse 0.8s infinite" : "none",
      letterSpacing: 1,
    }}>
      {formatCountdown(secs)}
    </span>
  );
}

// ─── SCHEDULE VIEW ────────────────────────────────────────────────────────────

function ScheduleView({ user, tournaments, sessions, setSessions, soundEnabled, alarmsEnabled, setAlarmsEnabled }) {
  const [showFilters, setShowFilters]         = useState(false);
  const [firingMap, setFiringMap]             = useState({});
  const [testCountdown, setTestCountdown]     = useState(null);
  // alarmsEnabled + setAlarmsEnabled viennent du parent App (persistance gérée là-bas)

  const toggleAlarm = (tid) => setAlarmsEnabled(prev => {
    const s = new Set(prev);
    s.has(tid) ? s.delete(tid) : s.add(tid);
    return s;
  });

  const allSites = [...new Set(tournaments.map(t => t.site))].sort();
  const [selectedSites, setSelectedSites] = useState([]);
  const [buyInMin, setBuyInMin]   = useState("");
  const [buyInMax, setBuyInMax]   = useState("");
  const [sortBy, setSortBy]       = useState("regTime");
  const [hideExpired, setHideExpired] = useState(false);

  const today = new Date().toISOString().split("T")[0];

  // Nettoyage auto des entrées firingMap > 2 minutes
  useEffect(() => {
    const interval = setInterval(() => {
      const now = Date.now();
      setFiringMap(prev => {
        const next = { ...prev };
        let changed = false;
        Object.entries(next).forEach(([id, ts]) => {
          if (now - ts > 120_000) { delete next[id]; changed = true; }
        });
        return changed ? next : prev;
      });
    }, 5000);
    return () => clearInterval(interval);
  }, []);

  const startTest = () => {
    if (testCountdown !== null) return;
    setTestCountdown(10);
  };

  useEffect(() => {
    if (testCountdown === null) return;
    if (testCountdown <= 0) {
      if (soundEnabled) playAlarm();
      setFiringMap(prev => ({ ...prev, "__test__": Date.now() }));
      setTimeout(() => {
        setFiringMap(prev => { const n = { ...prev }; delete n["__test__"]; return n; });
        setTestCountdown(null);
      }, 120_000);
      return;
    }
    const id = setTimeout(() => setTestCountdown(c => c - 1), 1000);
    return () => clearTimeout(id);
  }, [testCountdown, soundEnabled]);

  const handleFire = useCallback((tid) => {
    if (!alarmsEnabled.has(tid)) return;
    setFiringMap(prev => ({ ...prev, [tid]: Date.now() }));
  }, [alarmsEnabled]);

  const activeFiltersCount = [
    selectedSites.length > 0, buyInMin !== "", buyInMax !== "", hideExpired, sortBy !== "regTime",
  ].filter(Boolean).length;

  const resetFilters = () => { setSelectedSites([]); setBuyInMin(""); setBuyInMax(""); setSortBy("regTime"); setHideExpired(false); };
  const toggleSite = (site) => setSelectedSites(prev => prev.includes(site) ? prev.filter(s => s !== site) : [...prev, site]);

  const baseTodayList = todayTournaments(tournaments);
  const filteredList = baseTodayList
    .filter(t => selectedSites.length === 0 || selectedSites.includes(t.site))
    .filter(t => buyInMin === "" || t.buyIn >= parseFloat(buyInMin))
    .filter(t => buyInMax === "" || t.buyIn <= parseFloat(buyInMax))
    .filter(t => !hideExpired || getSecondsUntilReg(t.regTime) > 0)
    .sort((a, b) => sortBy === "buyInAsc" ? a.buyIn - b.buyIn : sortBy === "buyInDesc" ? b.buyIn - a.buyIn : getSecondsUntilReg(a.regTime) - getSecondsUntilReg(b.regTime));

  const sortLabels = { regTime: "Heure d\'enreg.", buyInAsc: "Buy-in ↑", buyInDesc: "Buy-in ↓" };

  const togglePlayed = (t) => {
    const key = `${today}_${t.id}`;
    const exists = sessions.find(s => s.key === key);
    if (exists) {
      setSessions(prev => prev.filter(s => s.key !== key));
    } else {
      setSessions(prev => [...prev, { key, date: today, tournamentId: t.id, tournamentName: t.name, site: t.site, buyIn: t.buyIn, rebuys: 0, won: null, userId: user.id }]);
      // Marquer joué éteint l'alarme visuelle
      setFiringMap(prev => { const n = { ...prev }; delete n[t.id]; return n; });
    }
  };

  const addRebuy = (t) => {
    const key = `${today}_${t.id}`;
    setSessions(prev => prev.map(s => s.key === key ? { ...s, rebuys: (s.rebuys || 0) + 1 } : s));
  };

  const setWin = (t, amount) => {
    const key = `${today}_${t.id}`;
    setSessions(prev => prev.map(s => s.key === key ? { ...s, won: parseFloat(amount) || 0 } : s));
  };

  const isPlayed   = (t) => sessions.some(s => s.key === `${today}_${t.id}` && s.userId === user.id);
  const getSession = (t) => sessions.find(s => s.key === `${today}_${t.id}` && s.userId === user.id);

  const testRegTime = (() => {
    if (testCountdown === null) return null;
    const d = new Date(Date.now() + testCountdown * 1000);
    return `${String(d.getHours()).padStart(2,"0")}:${String(d.getMinutes()).padStart(2,"0")}`;
  })();

  const TEST_TOURNAMENT = {
    id: "__test__", site: "TEST", name: "⚡ Tournoi Test Alarme",
    buyIn: 0, regTime: testRegTime || "00:00", days: [0,1,2,3,4,5,6], notes: "",
  };

  const displayList = testCountdown !== null ? [TEST_TOURNAMENT, ...filteredList] : filteredList;

  const TournamentCard = ({ t }) => {
    const played   = isPlayed(t);
    const sess     = getSession(t);
    const firing   = !!firingMap[t.id];
    const isTest   = t.id === "__test__";
    const alarmOn  = isTest || alarmsEnabled.has(t.id);

    const testSecs = isTest ? (testCountdown ?? 0) : null;
    const secs     = isTest ? testSecs : getSecondsUntilReg(t.regTime);
    const expired  = secs <= 0 && !firing;
    const critical = secs <= 30 && secs > 0;
    const warning  = secs <= 120 && secs > 30;

    const alarmElapsed   = firingMap[t.id] ? Math.floor((Date.now() - firingMap[t.id]) / 1000) : 0;
    const alarmRemaining = Math.max(0, 120 - alarmElapsed);

    const rebuys     = sess?.rebuys || 0;
    const totalBuyIn = t.buyIn * (1 + rebuys);

    const borderColor = firing
      ? "#ef4444"
      : played   ? "rgba(16,185,129,0.3)"
      : critical ? "rgba(239,68,68,0.35)"
      : warning  ? "rgba(245,158,11,0.25)"
      : isTest   ? "rgba(167,139,250,0.4)"
      : "rgba(255,255,255,0.06)";

    const bgColor = firing ? "rgba(239,68,68,0.10)"
      : played  ? "rgba(16,185,129,0.06)"
      : isTest  ? "rgba(167,139,250,0.06)"
      : "rgba(255,255,255,0.02)";

    return (
      <div style={{
        background: bgColor, border: `1px solid ${borderColor}`,
        borderRadius: 2, padding: "16px 20px",
        opacity: expired ? 0.35 : 1,
        transition: "background 0.4s ease, border-color 0.4s ease, box-shadow 0.4s ease",
        boxShadow: firing ? "0 0 28px rgba(239,68,68,0.4), inset 0 0 20px rgba(239,68,68,0.05)" : "none",
        animation: firing ? "alarmPulse 0.8s ease infinite" : "none",
        position: "relative", overflow: "hidden",
      }}>

        {/* Barre alarme 2 min */}
        {firing && !isTest && (
          <div style={{
            position: "absolute", top: 0, left: 0, height: 3,
            background: "linear-gradient(90deg, #ef4444, #f97316)",
            width: `${(alarmRemaining / 120) * 100}%`,
            transition: "width 1s linear",
          }} />
        )}

        {/* Barre test */}
        {isTest && testCountdown !== null && testCountdown > 0 && (
          <div style={{
            position: "absolute", top: 0, left: 0, height: 3,
            width: `${(testCountdown / 10) * 100}%`,
            background: "linear-gradient(90deg, #a78bfa, #7c3aed)",
            transition: "width 1s linear",
          }} />
        )}

        {/* Badge alarme */}
        {firing && (
          <div style={{
            position: "absolute", top: 8, right: 8,
            background: "#ef4444", color: "#fff",
            fontSize: 10, fontWeight: 900, letterSpacing: 1,
            padding: "3px 10px", borderRadius: 2,
            display: "flex", alignItems: "center", gap: 6,
          }}>
            🔔 ENREGISTREMENT
            {!isTest && <span style={{ opacity: 0.8, fontWeight: 400 }}>
              {Math.floor(alarmRemaining / 60)}:{String(alarmRemaining % 60).padStart(2,"0")}
            </span>}
          </div>
        )}

        <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>

          {/* Toggle alarme */}
          {!isTest && (
            <button onClick={() => toggleAlarm(t.id)}
              title={alarmOn ? "Désactiver l'alarme" : "Activer l'alarme"}
              style={{
                background: "none", border: "none", cursor: "pointer",
                fontSize: 16, padding: "2px 4px", borderRadius: 2,
                color: alarmOn ? "#fbbf24" : "#1e293b", transition: "all 0.2s",
                filter: alarmOn ? "drop-shadow(0 0 4px rgba(251,191,36,0.6))" : "none",
              }}
            >{alarmOn ? "🔔" : "🔕"}</button>
          )}

          {/* Site badge */}
          <div style={{
            padding: "3px 10px", borderRadius: 2,
            background: isTest ? "rgba(167,139,250,0.15)" : "rgba(255,255,255,0.06)",
            color: isTest ? "#a78bfa" : "#94a3b8",
            fontSize: 11, letterSpacing: 1, textTransform: "uppercase", whiteSpace: "nowrap",
            border: isTest ? "1px solid rgba(167,139,250,0.3)" : "none",
          }}>{t.site}</div>

          {/* Name */}
          <div style={{ flex: 1, color: firing ? "#fca5a5" : "#e2e8f0", fontWeight: 600, fontSize: 15, transition: "color 0.3s", minWidth: 100 }}>{t.name}</div>

          {/* Buy-in */}
          {t.buyIn > 0 && (
            <div style={{ color: "#64748b", fontSize: 13, whiteSpace: "nowrap" }}>
              <span style={{ color: "#475569" }}>Buy-in: </span>
              <span style={{ color: rebuys > 0 ? "#f97316" : "#fbbf24", fontWeight: 700 }}>${totalBuyIn}</span>
              {rebuys > 0 && <span style={{ color: "#f97316", fontSize: 10, marginLeft: 4 }}>×{1 + rebuys}</span>}
            </div>
          )}

          {/* Reg time */}
          <div style={{ color: "#64748b", fontSize: 13, whiteSpace: "nowrap" }}>
            Reg: <span style={{ color: "#94a3b8", fontWeight: 600 }}>{t.regTime}</span>
          </div>

          {/* Countdown */}
          {isTest ? (
            <span style={{
              color: testCountdown === 0 ? "#ef4444" : testCountdown <= 3 ? "#ef4444" : "#a78bfa",
              fontWeight: 700, fontSize: 18, fontFamily: "\'Courier New\', monospace",
              animation: testCountdown !== null && testCountdown <= 3 ? "pulse 0.5s infinite" : "none",
              letterSpacing: 1, minWidth: 40, textAlign: "right",
            }}>
              {testCountdown === 0 || firing ? "🔔 NOW" : `00:${String(testCountdown).padStart(2,"0")}`}
            </span>
          ) : !expired && (
            <CountdownCell
              regTime={t.regTime}
              soundEnabled={soundEnabled && alarmOn}
              tournamentId={t.id}
              onFire={handleFire}
            />
          )}

          {/* Bouton JOUÉ */}
          {!isTest && (
            <button onClick={() => togglePlayed(t)} style={{
              padding: "6px 14px", borderRadius: 2, fontSize: 12, fontWeight: 700,
              letterSpacing: 1, textTransform: "uppercase", cursor: "pointer",
              border: "none", fontFamily: "\'Courier New\', monospace",
              background: played ? "rgba(16,185,129,0.2)" : firing ? "rgba(239,68,68,0.2)" : "rgba(255,255,255,0.06)",
              color: played ? "#10b981" : firing ? "#fca5a5" : "#64748b",
              transition: "all 0.2s",
            }}>{played ? "✓ JOUÉ" : "JOUÉ?"}</button>
          )}
        </div>

        {/* Panneau résultats */}
        {played && !isTest && (
          <div style={{ marginTop: 14, paddingTop: 14, borderTop: "1px solid rgba(255,255,255,0.05)", display: "flex", alignItems: "center", gap: 16, flexWrap: "wrap" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <button onClick={() => addRebuy(t)} style={{
                padding: "5px 14px", borderRadius: 2, fontSize: 11, fontWeight: 700,
                letterSpacing: 1, cursor: "pointer", fontFamily: "\'Courier New\', monospace",
                background: rebuys > 0 ? "rgba(249,115,22,0.2)" : "rgba(255,255,255,0.05)",
                border: `1px solid ${rebuys > 0 ? "rgba(249,115,22,0.5)" : "rgba(255,255,255,0.1)"}`,
                color: rebuys > 0 ? "#f97316" : "#64748b", transition: "all 0.2s",
              }}>+ REBUY</button>
              {rebuys > 0 && (
                <span style={{ color: "#f97316", fontSize: 12 }}>
                  {rebuys}× rebuy · Total investi: <span style={{ fontWeight: 700 }}>${totalBuyIn}</span>
                </span>
              )}
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginLeft: "auto" }}>
              <span style={{ color: "#64748b", fontSize: 12 }}>Gains ($)</span>
              <input
                type="number" placeholder="0.00"
                value={sess?.won ?? ""}
                onChange={e => setWin(t, e.target.value)}
                style={{ width: 110, padding: "6px 10px", background: "rgba(255,255,255,0.04)", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 2, color: "#10b981", fontSize: 14, fontWeight: 700, outline: "none", fontFamily: "\'Courier New\', monospace" }}
              />
              {sess?.won != null && (
                <span style={{ color: sess.won - totalBuyIn >= 0 ? "#10b981" : "#ef4444", fontSize: 13, fontWeight: 700, minWidth: 100 }}>
                  {sess.won - totalBuyIn >= 0 ? "+" : ""}${(sess.won - totalBuyIn).toFixed(2)} profit
                </span>
              )}
            </div>
          </div>
        )}
      </div>
    );
  };

  return (
    <div style={{ animation: "fadeIn 0.4s ease" }}>
      <div style={{ marginBottom: 20, display: "flex", alignItems: "flex-start", justifyContent: "space-between", flexWrap: "wrap", gap: 12 }}>
        <div>
          <div style={{ color: "#64748b", fontSize: 11, letterSpacing: 3, textTransform: "uppercase" }}>
            {DAYS_FULL[new Date().getDay()]} · {new Date().toLocaleDateString("fr-CA")}
          </div>
          <div style={{ color: "#e2e8f0", fontSize: 22, fontWeight: 700, marginTop: 4 }}>
            Horaire d'aujourd'hui
            <span style={{ marginLeft: 12, fontSize: 13, color: "#10b981", fontWeight: 400 }}>
              {filteredList.length}{filteredList.length !== baseTodayList.length ? `/${baseTodayList.length}` : ""} tournois
            </span>
          </div>
          <div style={{ color: "#334155", fontSize: 11, marginTop: 6 }}>
            🔔 <span style={{ color: "#475569" }}>= alarme activée</span> &nbsp;
            🔕 <span style={{ color: "#334155" }}>= alarme désactivée · cliquer pour basculer</span>
          </div>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          <button onClick={startTest} disabled={testCountdown !== null} style={{
            padding: "8px 16px", borderRadius: 2, fontSize: 12, cursor: testCountdown !== null ? "not-allowed" : "pointer",
            fontFamily: "\'Courier New\', monospace", letterSpacing: 1,
            border: "1px solid rgba(167,139,250,0.4)",
            background: testCountdown !== null ? "rgba(167,139,250,0.05)" : "rgba(167,139,250,0.12)",
            color: testCountdown !== null ? "#6d4ccc" : "#a78bfa",
            display: "flex", alignItems: "center", gap: 6,
          }}>
            🔔 {testCountdown !== null ? `TEST — ${testCountdown}s` : "TESTER L'ALARME (10s)"}
          </button>
          <button onClick={() => setShowFilters(p => !p)} style={{
            padding: "8px 18px", borderRadius: 2, fontSize: 12, cursor: "pointer",
            fontFamily: "\'Courier New\', monospace", letterSpacing: 1,
            border: `1px solid ${showFilters || activeFiltersCount > 0 ? "rgba(16,185,129,0.5)" : "rgba(255,255,255,0.1)"}`,
            background: showFilters || activeFiltersCount > 0 ? "rgba(16,185,129,0.1)" : "rgba(255,255,255,0.03)",
            color: showFilters || activeFiltersCount > 0 ? "#10b981" : "#64748b",
            display: "flex", alignItems: "center", gap: 8,
          }}>
            ⚙ FILTRES {activeFiltersCount > 0 && (
              <span style={{ background: "#10b981", color: "#0a0a0f", borderRadius: "50%", width: 18, height: 18, display: "inline-flex", alignItems: "center", justifyContent: "center", fontSize: 10, fontWeight: 700 }}>{activeFiltersCount}</span>
            )}
          </button>
        </div>
      </div>

      {showFilters && (
        <div style={{ background: "rgba(255,255,255,0.025)", border: "1px solid rgba(255,255,255,0.08)", borderRadius: 2, padding: "20px 24px", marginBottom: 20, animation: "fadeIn 0.25s ease" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18 }}>
            <span style={{ color: "#94a3b8", fontSize: 12, letterSpacing: 2, textTransform: "uppercase" }}>Options d'affichage</span>
            {activeFiltersCount > 0 && (
              <button onClick={resetFilters} style={{ padding: "4px 12px", fontSize: 11, cursor: "pointer", background: "rgba(239,68,68,0.1)", border: "1px solid rgba(239,68,68,0.3)", borderRadius: 2, color: "#ef4444", fontFamily: "\'Courier New\', monospace", letterSpacing: 1 }}>✕ Réinitialiser</button>
            )}
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(200px, 1fr))", gap: 20 }}>
            <div>
              <div style={{ color: "#475569", fontSize: 11, letterSpacing: 2, marginBottom: 10, textTransform: "uppercase" }}>Sites</div>
              <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                {allSites.map(site => (
                  <label key={site} style={{ display: "flex", alignItems: "center", gap: 8, cursor: "pointer" }}>
                    <div onClick={() => toggleSite(site)} style={{ width: 16, height: 16, borderRadius: 2, flexShrink: 0, border: `1px solid ${selectedSites.includes(site) ? "#10b981" : "rgba(255,255,255,0.15)"}`, background: selectedSites.includes(site) ? "#10b981" : "transparent", display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", transition: "all 0.15s" }}>
                      {selectedSites.includes(site) && <span style={{ color: "#0a0a0f", fontSize: 10, fontWeight: 900 }}>✓</span>}
                    </div>
                    <span onClick={() => toggleSite(site)} style={{ color: selectedSites.includes(site) ? "#e2e8f0" : "#64748b", fontSize: 13, cursor: "pointer" }}>{site}</span>
                  </label>
                ))}
              </div>
            </div>
            <div>
              <div style={{ color: "#475569", fontSize: 11, letterSpacing: 2, marginBottom: 10, textTransform: "uppercase" }}>Buy-in ($)</div>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <div style={{ flex: 1 }}>
                  <div style={{ color: "#334155", fontSize: 10, marginBottom: 4 }}>MIN</div>
                  <input type="number" placeholder="0" value={buyInMin} onChange={e => setBuyInMin(e.target.value)} style={{ ...inputStyle, padding: "8px 10px", fontSize: 13 }} />
                </div>
                <div style={{ color: "#334155", fontSize: 14, marginTop: 16 }}>—</div>
                <div style={{ flex: 1 }}>
                  <div style={{ color: "#334155", fontSize: 10, marginBottom: 4 }}>MAX</div>
                  <input type="number" placeholder="∞" value={buyInMax} onChange={e => setBuyInMax(e.target.value)} style={{ ...inputStyle, padding: "8px 10px", fontSize: 13 }} />
                </div>
              </div>
              <div style={{ display: "flex", gap: 4, marginTop: 8, flexWrap: "wrap" }}>
                {[["Micro","","33"],["Bas","33","109"],["Mid","109","215"],["High","215",""]].map(([label, min, max]) => (
                  <button key={label} onClick={() => { setBuyInMin(min); setBuyInMax(max); }} style={{ padding: "3px 8px", fontSize: 10, cursor: "pointer", fontFamily: "\'Courier New\', monospace", border: `1px solid ${buyInMin === min && buyInMax === max ? "rgba(251,191,36,0.5)" : "rgba(255,255,255,0.08)"}`, background: buyInMin === min && buyInMax === max ? "rgba(251,191,36,0.1)" : "transparent", borderRadius: 2, color: buyInMin === min && buyInMax === max ? "#fbbf24" : "#475569" }}>{label}</button>
                ))}
              </div>
            </div>
            <div>
              <div style={{ color: "#475569", fontSize: 11, letterSpacing: 2, marginBottom: 10, textTransform: "uppercase" }}>Trier par</div>
              <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                {Object.entries(sortLabels).map(([val, label]) => (
                  <label key={val} style={{ display: "flex", alignItems: "center", gap: 8, cursor: "pointer" }} onClick={() => setSortBy(val)}>
                    <div style={{ width: 16, height: 16, borderRadius: "50%", flexShrink: 0, border: `1px solid ${sortBy === val ? "#a78bfa" : "rgba(255,255,255,0.15)"}`, background: sortBy === val ? "#a78bfa" : "transparent", display: "flex", alignItems: "center", justifyContent: "center", transition: "all 0.15s" }}>
                      {sortBy === val && <div style={{ width: 6, height: 6, borderRadius: "50%", background: "#0a0a0f" }} />}
                    </div>
                    <span style={{ color: sortBy === val ? "#e2e8f0" : "#64748b", fontSize: 13 }}>{label}</span>
                  </label>
                ))}
              </div>
            </div>
            <div>
              <div style={{ color: "#475569", fontSize: 11, letterSpacing: 2, marginBottom: 10, textTransform: "uppercase" }}>Options</div>
              <label style={{ display: "flex", alignItems: "center", gap: 8, cursor: "pointer" }} onClick={() => setHideExpired(p => !p)}>
                <div style={{ width: 16, height: 16, borderRadius: 2, flexShrink: 0, border: `1px solid ${hideExpired ? "#10b981" : "rgba(255,255,255,0.15)"}`, background: hideExpired ? "#10b981" : "transparent", display: "flex", alignItems: "center", justifyContent: "center", transition: "all 0.15s", cursor: "pointer" }}>
                  {hideExpired && <span style={{ color: "#0a0a0f", fontSize: 10, fontWeight: 900 }}>✓</span>}
                </div>
                <span style={{ color: hideExpired ? "#e2e8f0" : "#64748b", fontSize: 13 }}>Masquer les expirés</span>
              </label>
            </div>
          </div>
          {activeFiltersCount > 0 && (
            <div style={{ marginTop: 16, paddingTop: 14, borderTop: "1px solid rgba(255,255,255,0.05)", color: "#475569", fontSize: 11 }}>
              Affichage: <span style={{ color: "#10b981" }}>{filteredList.length}</span> sur {baseTodayList.length}
              {selectedSites.length > 0 && <span> · Sites: <span style={{ color: "#94a3b8" }}>{selectedSites.join(", ")}</span></span>}
              {(buyInMin || buyInMax) && <span> · Buy-in: <span style={{ color: "#fbbf24" }}>${buyInMin || "0"}–${buyInMax || "∞"}</span></span>}
              {sortBy !== "regTime" && <span> · Tri: <span style={{ color: "#a78bfa" }}>{sortLabels[sortBy]}</span></span>}
              {hideExpired && <span> · expirés masqués</span>}
            </div>
          )}
        </div>
      )}

      {displayList.length === 0 ? (
        <div style={{ color: "#334155", textAlign: "center", padding: 60, fontSize: 14 }}>
          {baseTodayList.length === 0 ? "Aucun tournoi prévu aujourd'hui." : "Aucun tournoi ne correspond aux filtres actifs."}
        </div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {displayList.map(t => <TournamentCard key={t.id} t={t} />)}
        </div>
      )}
    </div>
  );
}


// ─── STATS VIEW ───────────────────────────────────────────────────────────────

function StatsView({ user, sessions }) {
  const [period, setPeriod] = useState("week");

  const now = new Date();
  const filtered = sessions.filter(s => {
    if (s.userId !== user.id) return false;
    const d = new Date(s.date);
    if (period === "today") return s.date === now.toISOString().split("T")[0];
    if (period === "week") { const w = new Date(now); w.setDate(now.getDate() - 7); return d >= w; }
    if (period === "month") { const m = new Date(now); m.setDate(now.getDate() - 30); return d >= m; }
    return true;
  });

  const volume = filtered.length;
  const invested = filtered.reduce((sum, s) => sum + s.buyIn, 0);
  const won = filtered.filter(s => s.won != null).reduce((sum, s) => sum + s.won, 0);
  const profit = won - invested;
  const roi = invested > 0 ? ((profit / invested) * 100).toFixed(1) : "—";

  // Daily breakdown for chart
  const dailyMap = {};
  filtered.forEach(s => {
    if (!dailyMap[s.date]) dailyMap[s.date] = { invested: 0, won: 0, count: 0 };
    dailyMap[s.date].invested += s.buyIn;
    dailyMap[s.date].won += s.won ?? 0;
    dailyMap[s.date].count += 1;
  });
  const days = Object.keys(dailyMap).sort().slice(-14);
  const maxProfit = Math.max(...days.map(d => Math.abs(dailyMap[d].won - dailyMap[d].invested)), 1);

  const statCard = (label, value, color = "#e2e8f0", sub = null) => (
    <div style={{
      background: "rgba(255,255,255,0.02)", border: "1px solid rgba(255,255,255,0.06)",
      borderRadius: 2, padding: "20px 24px", flex: 1, minWidth: 140,
    }}>
      <div style={{ color: "#475569", fontSize: 11, letterSpacing: 2, textTransform: "uppercase", marginBottom: 8 }}>{label}</div>
      <div style={{ color, fontSize: 26, fontWeight: 700, fontFamily: "'Courier New', monospace" }}>{value}</div>
      {sub && <div style={{ color: "#334155", fontSize: 11, marginTop: 4 }}>{sub}</div>}
    </div>
  );

  return (
    <div style={{ animation: "fadeIn 0.4s ease" }}>
      <div style={{ marginBottom: 24, display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12 }}>
        <div>
          <div style={{ color: "#64748b", fontSize: 11, letterSpacing: 3, textTransform: "uppercase" }}>Performance</div>
          <div style={{ color: "#e2e8f0", fontSize: 22, fontWeight: 700 }}>Statistiques</div>
        </div>
        <div style={{ display: "flex", gap: 6 }}>
          {[["today","Aujourd'hui"],["week","7 jours"],["month","30 jours"],["all","Tout"]].map(([v, l]) => (
            <button key={v} onClick={() => setPeriod(v)} style={{
              padding: "6px 14px", borderRadius: 2, fontSize: 12, cursor: "pointer",
              fontFamily: "'Courier New', monospace", letterSpacing: 1,
              border: "1px solid",
              borderColor: period === v ? "#10b981" : "rgba(255,255,255,0.08)",
              background: period === v ? "rgba(16,185,129,0.15)" : "transparent",
              color: period === v ? "#10b981" : "#64748b",
            }}>{l}</button>
          ))}
        </div>
      </div>

      {/* Stat cards */}
      <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 24 }}>
        {statCard("Volume", volume, "#e2e8f0", "tournois joués")}
        {statCard("Investi", `$${invested.toFixed(0)}`, "#fbbf24")}
        {statCard("Gains bruts", `$${won.toFixed(0)}`, "#10b981")}
        {statCard("Profit net", `${profit >= 0 ? "+" : ""}$${profit.toFixed(0)}`, profit >= 0 ? "#10b981" : "#ef4444")}
        {statCard("ROI", `${roi}%`, parseFloat(roi) >= 0 ? "#10b981" : "#ef4444")}
      </div>

      {/* Chart */}
      {days.length > 0 && (
        <div style={{ background: "rgba(255,255,255,0.02)", border: "1px solid rgba(255,255,255,0.06)", borderRadius: 2, padding: 24 }}>
          <div style={{ color: "#475569", fontSize: 11, letterSpacing: 2, textTransform: "uppercase", marginBottom: 20 }}>Profit quotidien</div>
          <div style={{ display: "flex", alignItems: "flex-end", gap: 6, height: 100 }}>
            {days.map(d => {
              const p = dailyMap[d].won - dailyMap[d].invested;
              const h = Math.abs(p) / maxProfit * 80;
              return (
                <div key={d} style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", gap: 4 }} title={`${d}: $${p.toFixed(0)}`}>
                  <div style={{
                    width: "100%", height: h, minHeight: 4, borderRadius: "2px 2px 0 0",
                    background: p >= 0 ? "#10b981" : "#ef4444",
                    opacity: 0.8,
                  }} />
                  <div style={{ color: "#334155", fontSize: 9, transform: "rotate(-45deg)", whiteSpace: "nowrap" }}>
                    {d.slice(5)}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* Sessions table */}
      {filtered.length > 0 && (
        <div style={{ marginTop: 24, background: "rgba(255,255,255,0.02)", border: "1px solid rgba(255,255,255,0.06)", borderRadius: 2, overflow: "hidden" }}>
          <div style={{ padding: "16px 20px", borderBottom: "1px solid rgba(255,255,255,0.06)", color: "#475569", fontSize: 11, letterSpacing: 2, textTransform: "uppercase" }}>Historique des sessions</div>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead>
                <tr style={{ color: "#475569" }}>
                  {["Date","Site","Tournoi","Buy-in","Gains","Profit"].map(h => (
                    <th key={h} style={{ padding: "10px 16px", textAlign: "left", fontWeight: 400, letterSpacing: 1 }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {[...filtered].reverse().map((s, i) => {
                  const p = (s.won ?? 0) - s.buyIn;
                  return (
                    <tr key={s.key} style={{ borderTop: "1px solid rgba(255,255,255,0.04)" }}>
                      <td style={{ padding: "10px 16px", color: "#64748b" }}>{s.date}</td>
                      <td style={{ padding: "10px 16px", color: "#64748b" }}>{s.site}</td>
                      <td style={{ padding: "10px 16px", color: "#e2e8f0" }}>{s.tournamentName}</td>
                      <td style={{ padding: "10px 16px", color: "#fbbf24" }}>${s.buyIn}</td>
                      <td style={{ padding: "10px 16px", color: "#10b981" }}>{s.won != null ? `$${s.won}` : "—"}</td>
                      <td style={{ padding: "10px 16px", color: p >= 0 ? "#10b981" : "#ef4444", fontWeight: 700 }}>{s.won != null ? `${p >= 0 ? "+" : ""}$${p.toFixed(2)}` : "—"}</td>
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

// ─── TOURNAMENTS ADMIN ────────────────────────────────────────────────────────

function TournamentsAdmin({ tournaments, setTournaments, alarmsEnabled, setAlarmsEnabled }) {
  const [form, setForm]       = useState({ site: "", name: "", buyIn: "", regTime: "", startTime: "", days: [], notes: "" });
  const [editing, setEditing] = useState(null);
  const [showForm, setShowForm] = useState(false);
  const fileRef = useRef();

  // ── Filtres ──
  const [search, setSearch]         = useState("");
  const [filterSite, setFilterSite] = useState("all");
  const [filterDay, setFilterDay]   = useState("all");
  const [filterBuyMin, setFilterBuyMin] = useState("");
  const [filterBuyMax, setFilterBuyMax] = useState("");
  const [sortCol, setSortCol]       = useState("name");  // name | buyIn | regTime | site
  const [sortDir, setSortDir]       = useState("asc");

  // ── Sélection multiple ──
  const [selected, setSelected] = useState(new Set());
  const [confirmDelete, setConfirmDelete] = useState(false); // confirmation avant suppression masse

  const allSites = [...new Set(tournaments.map(t => t.site))].sort();

  // ── Liste filtrée ──
  const filtered = tournaments
    .filter(t => search === "" || t.name.toLowerCase().includes(search.toLowerCase()) || t.site.toLowerCase().includes(search.toLowerCase()))
    .filter(t => filterSite === "all" || t.site === filterSite)
    .filter(t => filterDay === "all" || t.days.includes(parseInt(filterDay)))
    .filter(t => filterBuyMin === "" || t.buyIn >= parseFloat(filterBuyMin))
    .filter(t => filterBuyMax === "" || t.buyIn <= parseFloat(filterBuyMax))
    .sort((a, b) => {
      let va = a[sortCol], vb = b[sortCol];
      if (sortCol === "buyIn") { va = parseFloat(va); vb = parseFloat(vb); }
      const cmp = typeof va === "number" ? va - vb : String(va).localeCompare(String(vb));
      return sortDir === "asc" ? cmp : -cmp;
    });

  const allFilteredIds  = filtered.map(t => t.id);
  const allSelected     = allFilteredIds.length > 0 && allFilteredIds.every(id => selected.has(id));
  const someSelected    = allFilteredIds.some(id => selected.has(id));
  const selectedCount   = [...selected].filter(id => allFilteredIds.includes(id)).length;

  const toggleSelect    = (id) => setSelected(prev => { const s = new Set(prev); s.has(id) ? s.delete(id) : s.add(id); return s; });
  const toggleSelectAll = () => {
    if (allSelected) setSelected(prev => { const s = new Set(prev); allFilteredIds.forEach(id => s.delete(id)); return s; });
    else setSelected(prev => { const s = new Set(prev); allFilteredIds.forEach(id => s.add(id)); return s; });
  };

  const handleBulkDelete = () => {
    setTournaments(prev => prev.filter(t => !selected.has(t.id)));
    setSelected(new Set());
    setConfirmDelete(false);
  };

  const [timeInputMode, setTimeInputMode] = useState("clock"); // "clock" | "countdown"
  const [cdH, setCdH] = useState("0");
  const [cdM, setCdM] = useState("");
  const [cdS, setCdS] = useState("");

  // Convertit le compte à rebours H/M/S en heure HH:MM absolue
  const countdownToRegTime = () => {
    const totalSecs = (parseInt(cdH) || 0) * 3600 + (parseInt(cdM) || 0) * 60 + (parseInt(cdS) || 0);
    const target = new Date(Date.now() + totalSecs * 1000);
    return `${String(target.getHours()).padStart(2,"0")}:${String(target.getMinutes()).padStart(2,"0")}`;
  };

  const resetForm = () => {
    setForm({ site: "", name: "", buyIn: "", regTime: "", startTime: "", days: [], notes: "" });
    setCdH("0"); setCdM(""); setCdS("");
    setTimeInputMode("clock");
  };

  const handleSave = () => {
    // Seuls site + regTime (ou countdown) sont obligatoires
    if (!form.site) return;
    let regTime = form.regTime;
    if (timeInputMode === "countdown") {
      if (!cdM && !cdS && !cdH) return;
      regTime = countdownToRegTime();
    }
    if (!regTime) return;
    const entry = {
      ...form,
      regTime,
      name: form.name || form.site,          // nom par défaut = site
      buyIn: form.buyIn ? parseFloat(form.buyIn) : 0,
    };
    if (editing) {
      setTournaments(prev => prev.map(t => t.id === editing ? { ...entry, id: editing } : t));
      setEditing(null);
    } else {
      const newId = `t${Date.now()}`;
      setTournaments(prev => [...prev, { ...entry, id: newId }]);
      // Alarme activée automatiquement pour tout nouveau tournoi
      setAlarmsEnabled(prev => { const s = new Set(prev); s.add(newId); return s; });
    }
    resetForm(); setShowForm(false);
  };

  const handleEdit = (t) => {
    setForm({ ...t, buyIn: String(t.buyIn || "") });
    setTimeInputMode("clock");
    setCdH("0"); setCdM(""); setCdS("");
    setEditing(t.id); setShowForm(true);
  };
  const handleDelete = (id) => { setTournaments(prev => prev.filter(t => t.id !== id)); setSelected(prev => { const s = new Set(prev); s.delete(id); return s; }); };

  const handleImport = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const wb = XLSX.read(ev.target.result, { type: "binary" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws);
      const imported = data.map((row, i) => ({
        id: `import_${Date.now()}_${i}`,
        site: row["Site"] || row["site"] || "",
        name: row["Nom"] || row["name"] || row["Tournoi"] || "",
        buyIn: parseFloat(row["Buy-in"] || row["buyin"] || row["BuyIn"] || 0),
        regTime: String(row["Heure enregistrement"] || row["regTime"] || row["Reg"] || "20:00"),
        startTime: String(row["Heure début"] || row["startTime"] || ""),
        days: [0,1,2,3,4,5,6],
        notes: row["Notes"] || row["notes"] || "",
      }));
      setTournaments(prev => [...prev, ...imported]);
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  };

  const toggleDay = (d) => setForm(prev => ({ ...prev, days: prev.days.includes(d) ? prev.days.filter(x => x !== d) : [...prev.days, d] }));

  const SortBtn = ({ col, label }) => (
    <button onClick={() => { if (sortCol === col) setSortDir(d => d === "asc" ? "desc" : "asc"); else { setSortCol(col); setSortDir("asc"); } }}
      style={{ background: "none", border: "none", cursor: "pointer", color: sortCol === col ? "#a78bfa" : "#475569", fontSize: 11, letterSpacing: 1, fontFamily: "'Courier New', monospace", display: "flex", alignItems: "center", gap: 3 }}>
      {label} {sortCol === col ? (sortDir === "asc" ? "↑" : "↓") : "↕"}
    </button>
  );

  const activeFilterCount = [search, filterSite !== "all", filterDay !== "all", filterBuyMin, filterBuyMax].filter(Boolean).length;

  return (
    <div style={{ animation: "fadeIn 0.4s ease" }}>

      {/* Header */}
      <div style={{ marginBottom: 20, display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12 }}>
        <div>
          <div style={{ color: "#64748b", fontSize: 11, letterSpacing: 3, textTransform: "uppercase" }}>Admin</div>
          <div style={{ color: "#e2e8f0", fontSize: 22, fontWeight: 700 }}>
            Gestion des Tournois
            <span style={{ marginLeft: 12, fontSize: 13, color: "#10b981", fontWeight: 400 }}>
              {filtered.length}{filtered.length !== tournaments.length ? `/${tournaments.length}` : ""} tournois
            </span>
          </div>
        </div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <input type="file" accept=".xlsx,.csv,.xls" ref={fileRef} style={{ display: "none" }} onChange={handleImport} />
          <button onClick={() => fileRef.current?.click()} style={btnSecondary}>📥 Importer Excel</button>
          <button onClick={() => { resetForm(); setEditing(null); setShowForm(true); }} style={btnPrimary2}>+ Ajouter</button>
        </div>
      </div>

      {/* ── Barre de filtres ── */}
      <div style={{ background: "rgba(255,255,255,0.02)", border: "1px solid rgba(255,255,255,0.07)", borderRadius: 2, padding: "14px 18px", marginBottom: 16, display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>

        {/* Recherche texte */}
        <div style={{ position: "relative", flex: 1, minWidth: 160 }}>
          <span style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)", color: "#334155", fontSize: 13 }}>🔍</span>
          <input
            placeholder="Rechercher..."
            value={search}
            onChange={e => setSearch(e.target.value)}
            style={{ ...inputStyle, padding: "7px 12px 7px 32px", fontSize: 12 }}
          />
        </div>

        {/* Filtre site */}
        <select value={filterSite} onChange={e => setFilterSite(e.target.value)}
          style={{ ...inputStyle, padding: "7px 12px", fontSize: 12, minWidth: 120, cursor: "pointer" }}>
          <option value="all">Tous les sites</option>
          {allSites.map(s => <option key={s} value={s}>{s}</option>)}
        </select>

        {/* Filtre jour */}
        <select value={filterDay} onChange={e => setFilterDay(e.target.value)}
          style={{ ...inputStyle, padding: "7px 12px", fontSize: 12, minWidth: 110, cursor: "pointer" }}>
          <option value="all">Tous les jours</option>
          {DAYS_FULL.map((d, i) => <option key={i} value={i}>{d}</option>)}
        </select>

        {/* Buy-in min/max */}
        <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
          <input type="number" placeholder="$ min" value={filterBuyMin} onChange={e => setFilterBuyMin(e.target.value)}
            style={{ ...inputStyle, padding: "7px 10px", fontSize: 12, width: 72 }} />
          <span style={{ color: "#334155" }}>—</span>
          <input type="number" placeholder="$ max" value={filterBuyMax} onChange={e => setFilterBuyMax(e.target.value)}
            style={{ ...inputStyle, padding: "7px 10px", fontSize: 12, width: 72 }} />
        </div>

        {/* Reset filtres */}
        {activeFilterCount > 0 && (
          <button onClick={() => { setSearch(""); setFilterSite("all"); setFilterDay("all"); setFilterBuyMin(""); setFilterBuyMax(""); }}
            style={{ padding: "7px 12px", fontSize: 11, cursor: "pointer", background: "rgba(239,68,68,0.08)", border: "1px solid rgba(239,68,68,0.25)", borderRadius: 2, color: "#ef4444", fontFamily: "'Courier New', monospace", letterSpacing: 1, whiteSpace: "nowrap" }}>
            ✕ Réinitialiser
          </button>
        )}
      </div>

      {/* ── Barre sélection / suppression masse ── */}
      {filtered.length > 0 && (
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 10, padding: "8px 4px", flexWrap: "wrap" }}>

          {/* Checkbox tout sélectionner */}
          <label style={{ display: "flex", alignItems: "center", gap: 8, cursor: "pointer" }}>
            <div onClick={toggleSelectAll} style={{
              width: 16, height: 16, borderRadius: 2, flexShrink: 0,
              border: `1px solid ${someSelected ? "#a78bfa" : "rgba(255,255,255,0.15)"}`,
              background: allSelected ? "#a78bfa" : someSelected ? "rgba(167,139,250,0.3)" : "transparent",
              display: "flex", alignItems: "center", justifyContent: "center",
              cursor: "pointer", transition: "all 0.15s",
            }}>
              {allSelected && <span style={{ color: "#0a0a0f", fontSize: 10, fontWeight: 900 }}>✓</span>}
              {!allSelected && someSelected && <span style={{ color: "#a78bfa", fontSize: 12, lineHeight: 1 }}>—</span>}
            </div>
            <span style={{ color: "#475569", fontSize: 12 }}>
              {allSelected ? "Tout désélectionner" : "Tout sélectionner"}
              {someSelected && <span style={{ color: "#a78bfa", marginLeft: 6 }}>({selectedCount} sélectionné{selectedCount > 1 ? "s" : ""})</span>}
            </span>
          </label>

          {/* Bouton suppression masse */}
          {selectedCount > 0 && !confirmDelete && (
            <button onClick={() => setConfirmDelete(true)}
              style={{ padding: "6px 16px", borderRadius: 2, fontSize: 12, cursor: "pointer", fontFamily: "'Courier New', monospace", letterSpacing: 1, background: "rgba(239,68,68,0.1)", border: "1px solid rgba(239,68,68,0.35)", color: "#ef4444" }}>
              🗑 Supprimer {selectedCount} tournoi{selectedCount > 1 ? "s" : ""}
            </button>
          )}

          {/* Confirmation */}
          {confirmDelete && (
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <span style={{ color: "#ef4444", fontSize: 12 }}>Confirmer la suppression de {selectedCount} tournoi{selectedCount > 1 ? "s" : ""} ?</span>
              <button onClick={handleBulkDelete}
                style={{ padding: "5px 14px", borderRadius: 2, fontSize: 11, cursor: "pointer", fontFamily: "'Courier New', monospace", background: "#7f1d1d", border: "1px solid #ef4444", color: "#fca5a5", fontWeight: 700 }}>
                OUI, SUPPRIMER
              </button>
              <button onClick={() => setConfirmDelete(false)}
                style={{ padding: "5px 12px", borderRadius: 2, fontSize: 11, cursor: "pointer", fontFamily: "'Courier New', monospace", background: "transparent", border: "1px solid rgba(255,255,255,0.1)", color: "#64748b" }}>
                Annuler
              </button>
            </div>
          )}
        </div>
      )}

      {/* ── Formulaire ajout/édition ── */}
      {showForm && (
        <div style={{ background: "rgba(255,255,255,0.03)", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 2, padding: 24, marginBottom: 20, animation: "fadeIn 0.3s ease" }}>
          <div style={{ color: "#e2e8f0", fontWeight: 600, marginBottom: 20 }}>{editing ? "Modifier le tournoi" : "Nouveau tournoi"}</div>

          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(180px, 1fr))", gap: 12, marginBottom: 16 }}>

            {/* Site — obligatoire */}
            <div>
              <div style={{ color: "#475569", fontSize: 11, marginBottom: 6, letterSpacing: 1 }}>
                SITE <span style={{ color: "#ef4444" }}>*</span>
              </div>
              <input
                placeholder="PokerStars, GGPoker…"
                value={form.site}
                onChange={e => setForm(p => ({ ...p, site: e.target.value }))}
                style={{ ...inputStyle, padding: "8px 12px", borderColor: !form.site ? "rgba(239,68,68,0.3)" : undefined }}
              />
            </div>

            {/* Nom — optionnel */}
            <div>
              <div style={{ color: "#475569", fontSize: 11, marginBottom: 6, letterSpacing: 1 }}>
                NOM <span style={{ color: "#334155", fontSize: 10 }}>(optionnel)</span>
              </div>
              <input
                placeholder={form.site || "Nom du tournoi"}
                value={form.name}
                onChange={e => setForm(p => ({ ...p, name: e.target.value }))}
                style={{ ...inputStyle, padding: "8px 12px" }}
              />
            </div>

            {/* Buy-in — optionnel */}
            <div>
              <div style={{ color: "#475569", fontSize: 11, marginBottom: 6, letterSpacing: 1 }}>
                BUY-IN ($) <span style={{ color: "#334155", fontSize: 10 }}>(optionnel)</span>
              </div>
              <input
                type="number" placeholder="0"
                value={form.buyIn}
                onChange={e => setForm(p => ({ ...p, buyIn: e.target.value }))}
                style={{ ...inputStyle, padding: "8px 12px" }}
              />
            </div>

            {/* Heure début — optionnel */}
            <div>
              <div style={{ color: "#475569", fontSize: 11, marginBottom: 6, letterSpacing: 1 }}>
                HEURE DÉBUT <span style={{ color: "#334155", fontSize: 10 }}>(optionnel)</span>
              </div>
              <input
                type="time"
                value={form.startTime}
                onChange={e => setForm(p => ({ ...p, startTime: e.target.value }))}
                style={{ ...inputStyle, padding: "8px 12px" }}
              />
            </div>

            {/* Notes — optionnel */}
            <div>
              <div style={{ color: "#475569", fontSize: 11, marginBottom: 6, letterSpacing: 1 }}>
                NOTES <span style={{ color: "#334155", fontSize: 10 }}>(optionnel)</span>
              </div>
              <input
                value={form.notes}
                onChange={e => setForm(p => ({ ...p, notes: e.target.value }))}
                style={{ ...inputStyle, padding: "8px 12px" }}
              />
            </div>
          </div>

          {/* ── Heure d'enregistrement — obligatoire ── */}
          <div style={{ background: "rgba(255,255,255,0.02)", border: "1px solid rgba(255,255,255,0.07)", borderRadius: 2, padding: "16px 18px", marginBottom: 16 }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
              <div style={{ color: "#e2e8f0", fontSize: 12, fontWeight: 600, letterSpacing: 1 }}>
                HEURE D'ENREGISTREMENT <span style={{ color: "#ef4444" }}>*</span>
              </div>
              {/* Toggle mode */}
              <div style={{ display: "flex", background: "rgba(255,255,255,0.05)", borderRadius: 2, padding: 2, gap: 2 }}>
                {[["clock", "🕐 Heure"], ["countdown", "⏱ Compte à rebours"]].map(([mode, label]) => (
                  <button key={mode} onClick={() => setTimeInputMode(mode)} style={{
                    padding: "4px 12px", borderRadius: 2, fontSize: 11, cursor: "pointer",
                    fontFamily: "'Courier New', monospace", letterSpacing: 1, border: "none",
                    background: timeInputMode === mode ? "rgba(167,139,250,0.25)" : "transparent",
                    color: timeInputMode === mode ? "#a78bfa" : "#475569",
                    transition: "all 0.15s",
                  }}>{label}</button>
                ))}
              </div>
            </div>

            {timeInputMode === "clock" ? (
              /* Mode heure classique */
              <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                <input
                  type="time"
                  value={form.regTime}
                  onChange={e => setForm(p => ({ ...p, regTime: e.target.value }))}
                  style={{ ...inputStyle, padding: "10px 14px", fontSize: 18, width: 150,
                    borderColor: !form.regTime ? "rgba(239,68,68,0.3)" : "rgba(255,255,255,0.1)" }}
                />
                <span style={{ color: "#475569", fontSize: 12 }}>Entrer l'heure exacte (HH:MM)</span>
              </div>
            ) : (
              /* Mode compte à rebours */
              <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                {/* Heures */}
                <div style={{ textAlign: "center" }}>
                  <input
                    type="number" min="0" max="23" placeholder="0"
                    value={cdH}
                    onChange={e => setCdH(e.target.value)}
                    style={{ ...inputStyle, padding: "10px 8px", fontSize: 20, width: 70, textAlign: "center" }}
                  />
                  <div style={{ color: "#334155", fontSize: 10, marginTop: 4, letterSpacing: 1 }}>HEURES</div>
                </div>
                <span style={{ color: "#475569", fontSize: 24, paddingBottom: 18 }}>:</span>
                {/* Minutes */}
                <div style={{ textAlign: "center" }}>
                  <input
                    type="number" min="0" max="59" placeholder="30"
                    value={cdM}
                    onChange={e => setCdM(e.target.value)}
                    style={{ ...inputStyle, padding: "10px 8px", fontSize: 20, width: 70, textAlign: "center",
                      borderColor: !cdM && !cdH && !cdS ? "rgba(239,68,68,0.3)" : "rgba(255,255,255,0.1)" }}
                  />
                  <div style={{ color: "#334155", fontSize: 10, marginTop: 4, letterSpacing: 1 }}>MINUTES</div>
                </div>
                <span style={{ color: "#475569", fontSize: 24, paddingBottom: 18 }}>:</span>
                {/* Secondes */}
                <div style={{ textAlign: "center" }}>
                  <input
                    type="number" min="0" max="59" placeholder="00"
                    value={cdS}
                    onChange={e => setCdS(e.target.value)}
                    style={{ ...inputStyle, padding: "10px 8px", fontSize: 20, width: 70, textAlign: "center" }}
                  />
                  <div style={{ color: "#334155", fontSize: 10, marginTop: 4, letterSpacing: 1 }}>SECONDES</div>
                </div>
                {/* Aperçu heure calculée */}
                {(cdH || cdM || cdS) && (
                  <div style={{ marginLeft: 8, padding: "8px 14px", background: "rgba(167,139,250,0.08)", border: "1px solid rgba(167,139,250,0.2)", borderRadius: 2 }}>
                    <div style={{ color: "#475569", fontSize: 10, marginBottom: 2 }}>HEURE D'ENREG.</div>
                    <div style={{ color: "#a78bfa", fontSize: 18, fontWeight: 700 }}>{countdownToRegTime()}</div>
                  </div>
                )}
              </div>
            )}
          </div>

          {/* Jours */}
          <div style={{ marginBottom: 16 }}>
            <div style={{ color: "#475569", fontSize: 11, marginBottom: 8, letterSpacing: 1 }}>JOURS</div>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
              {DAYS.map((d, i) => (
                <button key={i} onClick={() => toggleDay(i)} style={{
                  padding: "6px 10px", borderRadius: 2, fontSize: 12, cursor: "pointer",
                  fontFamily: "'Courier New', monospace", border: "1px solid",
                  borderColor: form.days.includes(i) ? "#10b981" : "rgba(255,255,255,0.08)",
                  background: form.days.includes(i) ? "rgba(16,185,129,0.15)" : "transparent",
                  color: form.days.includes(i) ? "#10b981" : "#475569",
                }}>{d}</button>
              ))}
            </div>
          </div>

          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={handleSave} style={{ ...btnPrimary2, padding: "8px 24px" }}>Sauvegarder</button>
            <button onClick={() => { setShowForm(false); setEditing(null); resetForm(); }} style={{ ...btnSecondary, padding: "8px 20px" }}>Annuler</button>
          </div>
        </div>
      )}

      {/* ── En-tête colonnes avec tri ── */}
      {filtered.length > 0 && (
        <div style={{ display: "flex", alignItems: "center", gap: 14, padding: "6px 18px", marginBottom: 4 }}>
          <div style={{ width: 16 }} />
          <SortBtn col="site"    label="SITE" />
          <SortBtn col="name"    label="NOM" />
          <div style={{ flex: 1 }} />
          <SortBtn col="buyIn"   label="BUY-IN" />
          <SortBtn col="regTime" label="REG" />
          <div style={{ minWidth: 160 }} />
        </div>
      )}

      {/* ── Liste filtrée ── */}
      {filtered.length === 0 ? (
        <div style={{ color: "#334155", textAlign: "center", padding: 48, fontSize: 14 }}>
          {tournaments.length === 0 ? "Aucun tournoi. Ajoutez-en ou importez un fichier Excel." : "Aucun résultat pour ces filtres."}
        </div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
          {filtered.map(t => {
            const isSel = selected.has(t.id);
            return (
              <div key={t.id} style={{
                background: isSel ? "rgba(167,139,250,0.06)" : "rgba(255,255,255,0.02)",
                border: `1px solid ${isSel ? "rgba(167,139,250,0.3)" : "rgba(255,255,255,0.06)"}`,
                borderRadius: 2, padding: "12px 18px",
                display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap",
                transition: "background 0.15s, border-color 0.15s",
              }}>
                {/* Checkbox */}
                <div onClick={() => toggleSelect(t.id)} style={{
                  width: 16, height: 16, borderRadius: 2, flexShrink: 0,
                  border: `1px solid ${isSel ? "#a78bfa" : "rgba(255,255,255,0.15)"}`,
                  background: isSel ? "#a78bfa" : "transparent",
                  display: "flex", alignItems: "center", justifyContent: "center",
                  cursor: "pointer", transition: "all 0.15s",
                }}>
                  {isSel && <span style={{ color: "#0a0a0f", fontSize: 10, fontWeight: 900 }}>✓</span>}
                </div>

                <span style={{ color: "#64748b", fontSize: 11, minWidth: 80, whiteSpace: "nowrap" }}>{t.site}</span>
                <span style={{ flex: 1, color: "#e2e8f0", fontWeight: 600, minWidth: 120 }}>{t.name}</span>
                <span style={{ color: "#fbbf24", fontSize: 13, whiteSpace: "nowrap" }}>${t.buyIn}</span>
                <span style={{ color: "#94a3b8", fontSize: 12, whiteSpace: "nowrap" }}>Reg {t.regTime}</span>
                <span style={{ color: "#334155", fontSize: 11, whiteSpace: "nowrap" }}>{t.days.map(d => DAYS[d]).join(" ")}</span>

                <div style={{ display: "flex", gap: 6, marginLeft: "auto" }}>
                  <button onClick={() => handleEdit(t)} style={{ ...btnSecondary, padding: "4px 12px", fontSize: 11 }}>Modifier</button>
                  <button onClick={() => handleDelete(t.id)} style={{ padding: "4px 12px", fontSize: 11, borderRadius: 2, background: "rgba(239,68,68,0.08)", border: "1px solid rgba(239,68,68,0.25)", color: "#ef4444", cursor: "pointer", fontFamily: "'Courier New', monospace" }}>Supprimer</button>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

// ─── PLAYERS ADMIN ────────────────────────────────────────────────────────────

function PlayersAdmin({ sessions, tournaments }) {
  const users = store.get("pk_users", INITIAL_USERS).filter(u => u.role === "player");

  return (
    <div style={{ animation: "fadeIn 0.4s ease" }}>
      <div style={{ marginBottom: 24 }}>
        <div style={{ color: "#64748b", fontSize: 11, letterSpacing: 3, textTransform: "uppercase" }}>Admin</div>
        <div style={{ color: "#e2e8f0", fontSize: 22, fontWeight: 700 }}>Vue Joueurs</div>
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        {users.map(u => {
          const userSessions = sessions.filter(s => s.userId === u.id);
          const now = new Date();
          const weekAgo = new Date(now); weekAgo.setDate(now.getDate() - 7);
          const weekSessions = userSessions.filter(s => new Date(s.date) >= weekAgo);
          const invested = weekSessions.reduce((sum, s) => sum + s.buyIn, 0);
          const won = weekSessions.filter(s => s.won != null).reduce((sum, s) => sum + s.won, 0);
          const profit = won - invested;
          const roi = invested > 0 ? ((profit / invested) * 100).toFixed(1) : "—";

          return (
            <div key={u.id} style={{
              background: "rgba(255,255,255,0.02)", border: "1px solid rgba(255,255,255,0.06)",
              borderRadius: 2, padding: "20px 24px"
            }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12 }}>
                <div>
                  <div style={{ color: "#e2e8f0", fontWeight: 700, fontSize: 16 }}>{u.name}</div>
                  <div style={{ color: "#475569", fontSize: 12, marginTop: 2 }}>@{u.username}</div>
                </div>
                <div style={{ display: "flex", gap: 20, flexWrap: "wrap" }}>
                  <div style={{ textAlign: "center" }}>
                    <div style={{ color: "#475569", fontSize: 10, letterSpacing: 1 }}>VOLUME 7J</div>
                    <div style={{ color: "#e2e8f0", fontWeight: 700, fontSize: 20 }}>{weekSessions.length}</div>
                  </div>
                  <div style={{ textAlign: "center" }}>
                    <div style={{ color: "#475569", fontSize: 10, letterSpacing: 1 }}>PROFIT 7J</div>
                    <div style={{ color: profit >= 0 ? "#10b981" : "#ef4444", fontWeight: 700, fontSize: 20 }}>{profit >= 0 ? "+" : ""}${profit.toFixed(0)}</div>
                  </div>
                  <div style={{ textAlign: "center" }}>
                    <div style={{ color: "#475569", fontSize: 10, letterSpacing: 1 }}>ROI</div>
                    <div style={{ color: parseFloat(roi) >= 0 ? "#10b981" : "#ef4444", fontWeight: 700, fontSize: 20 }}>{roi}%</div>
                  </div>
                  <div style={{ textAlign: "center" }}>
                    <div style={{ color: "#475569", fontSize: 10, letterSpacing: 1 }}>TOTAL SESSIONS</div>
                    <div style={{ color: "#94a3b8", fontWeight: 700, fontSize: 20 }}>{userSessions.length}</div>
                  </div>
                </div>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

const btnSecondary = {
  padding: "8px 16px", borderRadius: 2, fontSize: 12, cursor: "pointer",
  fontFamily: "'Courier New', monospace", letterSpacing: 1,
  border: "1px solid rgba(255,255,255,0.1)",
  background: "rgba(255,255,255,0.04)", color: "#94a3b8",
};

const btnPrimary2 = {
  padding: "8px 20px", borderRadius: 2, fontSize: 12, fontWeight: 700,
  cursor: "pointer", fontFamily: "'Courier New', monospace", letterSpacing: 1,
  border: "1px solid rgba(16,185,129,0.4)",
  background: "rgba(16,185,129,0.15)", color: "#10b981",
};

// ─── MAIN APP ─────────────────────────────────────────────────────────────────

export default function App() {
  const [currentUser, setCurrentUser] = useState(() => store.get("pk_session", null));
  const [tab, setTab]                 = useState("schedule");
  const [tournaments, setTournamentsRaw] = useState([]);
  const [sessions, setSessionsRaw]    = useState([]);
  const [soundEnabled, setSoundEnabled] = useState(() => store.get("pk_sound", true));
  const [loading, setLoading]         = useState(false);
  const [offline, setOffline]         = useState(false); // true si Supabase non configuré
  const [alarmsEnabled, setAlarmsEnabled] = useState(() => {
    const uid = store.get("pk_session", null)?.id;
    const saved = uid ? store.get("pk_alarms_" + uid, null) : null;
    return saved ? new Set(saved) : new Set();
  });

  // ── Détection config Supabase ──────────────────────────────────
  const isConfigured = SUPABASE_URL !== "COLLE_TON_URL_ICI" && SUPABASE_ANON !== "COLLE_TA_CLE_ANON_ICI";

  // ── Chargement initial ─────────────────────────────────────────
  useEffect(() => {
    if (!currentUser) return;
    if (!isConfigured) {
      setOffline(true);
      setTournamentsRaw(store.get("pk_tournaments", INITIAL_TOURNAMENTS));
      setSessionsRaw(store.get("pk_sessions", []));
      return;
    }
    setLoading(true);
    Promise.all([
      sb.select("tournaments"),
      sb.select("sessions", `user_id=eq.${currentUser.id}`),
    ]).then(([ts, ss]) => {
      setTournamentsRaw(ts.map(normT));
      setSessionsRaw(ss.map(normS));
      setOffline(false);
    }).catch(() => {
      setOffline(true);
      setTournamentsRaw(store.get("pk_tournaments", INITIAL_TOURNAMENTS));
      setSessionsRaw(store.get("pk_sessions", []));
    }).finally(() => setLoading(false));
  }, [currentUser?.id]);

  // ── Realtime : tournois (tout le monde voit les changements) ───
  useEffect(() => {
    if (!currentUser || !isConfigured) return;
    return sb.realtime("tournaments", (event, record, old) => {
      if (event === "DELETE") {
        setTournamentsRaw(prev => prev.filter(x => x.id !== old?.id));
      } else {
        const t = normT(record);
        setTournamentsRaw(prev => {
          const exists = prev.find(x => x.id === t.id);
          if (exists) return prev.map(x => x.id === t.id ? t : x);
          // Nouveau tournoi → activer alarme auto
          setAlarmsEnabled(p => { const s = new Set(p); s.add(t.id); return s; });
          return [...prev, t];
        });
      }
    });
  }, [currentUser?.id, isConfigured]);

  // ── Realtime : sessions du joueur connecté seulement ──────────
  useEffect(() => {
    if (!currentUser || !isConfigured) return;
    return sb.realtime("sessions", (event, record) => {
      if (!record || record.user_id !== currentUser.id) return;
      const s = normS(record);
      if (event === "INSERT") setSessionsRaw(prev => [...prev.filter(x => x.key !== s.key), s]);
      if (event === "UPDATE") setSessionsRaw(prev => prev.map(x => x.key === s.key ? s : x));
      if (event === "DELETE") setSessionsRaw(prev => prev.filter(x => x.key !== s.key));
    });
  }, [currentUser?.id, isConfigured]);

  // ── setTournaments : wrapper qui écrit dans Supabase ──────────
  const setTournaments = useCallback((updater) => {
    setTournamentsRaw(prev => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      if (!isConfigured || offline) { store.set("pk_tournaments", next); return next; }
      // Upsert nouveaux/modifiés
      next.forEach(t => {
        const old = prev.find(x => x.id === t.id);
        if (!old || JSON.stringify(old) !== JSON.stringify(t)) sb.upsert("tournaments", dbT(t));
      });
      // Supprimer retirés
      prev.forEach(t => { if (!next.find(x => x.id === t.id)) sb.remove("tournaments", "id", t.id); });
      return next;
    });
  }, [isConfigured, offline]);

  // ── setSessions : wrapper qui écrit dans Supabase ─────────────
  const setSessions = useCallback((updater) => {
    setSessionsRaw(prev => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      if (!isConfigured || offline) { store.set("pk_sessions", next); return next; }
      next.forEach(s => {
        const old = prev.find(x => x.key === s.key);
        if (!old || JSON.stringify(old) !== JSON.stringify(s)) sb.upsert("sessions", dbS(s));
      });
      prev.forEach(s => { if (!next.find(x => x.key === s.key)) sb.remove("sessions", "key", s.key); });
      return next;
    });
  }, [isConfigured, offline]);

  useEffect(() => { store.set("pk_sound", soundEnabled); }, [soundEnabled]);
  useEffect(() => {
    if (currentUser) store.set("pk_alarms_" + currentUser.id, [...alarmsEnabled]);
  }, [alarmsEnabled, currentUser]);

  if (!currentUser) return <LoginScreen onLogin={setCurrentUser} />;

  if (loading) return (
    <div style={{ minHeight: "100vh", background: "#0a0a0f", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Courier New', monospace" }}>
      <div style={{ textAlign: "center" }}>
        <div style={{ fontSize: 40, marginBottom: 16, animation: "pulse 1.5s infinite" }}>♠</div>
        <div style={{ color: "#475569", fontSize: 13, letterSpacing: 3 }}>CHARGEMENT...</div>
      </div>
      <style>{`@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.3}}*{box-sizing:border-box}`}</style>
    </div>
  );

  const isAdmin = currentUser.role === "admin";

  const navItems = [
    { id: "schedule", label: "Horaire", icon: "◈" },
    { id: "stats", label: "Stats", icon: "◆" },
    ...(isAdmin ? [
      { id: "tournaments", label: "Tournois", icon: "◇" },
      { id: "players", label: "Joueurs", icon: "◉" },
    ] : []),
  ];

  return (
    <div style={{
      minHeight: "100vh", background: "#0a0a0f",
      fontFamily: "'Courier New', monospace",
      backgroundImage: "radial-gradient(ellipse at 10% 0%, rgba(139,92,246,0.05) 0%, transparent 50%), radial-gradient(ellipse at 90% 100%, rgba(16,185,129,0.04) 0%, transparent 50%)",
    }}>
      {/* Sidebar */}
      <div style={{
        position: "fixed", left: 0, top: 0, bottom: 0, width: 200,
        background: "rgba(255,255,255,0.015)", borderRight: "1px solid rgba(255,255,255,0.05)",
        display: "flex", flexDirection: "column", padding: "32px 0", zIndex: 100,
      }}>
        {/* Logo */}
        <div style={{ padding: "0 24px 32px", borderBottom: "1px solid rgba(255,255,255,0.05)" }}>
          <div style={{ fontSize: 24, marginBottom: 4 }}>♠</div>
          <div style={{ color: "#e2e8f0", fontSize: 13, fontWeight: 700, letterSpacing: 2 }}>POKER</div>
          <div style={{ color: "#334155", fontSize: 10, letterSpacing: 3 }}>COACH PRO</div>
          {/* Indicateur connexion */}
          <div style={{ marginTop: 8, display: "flex", alignItems: "center", gap: 5 }}>
            <div style={{ width: 6, height: 6, borderRadius: "50%", background: offline ? "#f59e0b" : "#10b981" }} />
            <span style={{ color: offline ? "#78350f" : "#1e3a2f", fontSize: 9, letterSpacing: 1 }}>
              {offline ? "HORS-LIGNE" : "EN LIGNE"}
            </span>
          </div>
        </div>

        {/* Nav */}
        <nav style={{ flex: 1, padding: "24px 0" }}>
          {navItems.map(item => (
            <button key={item.id} onClick={() => setTab(item.id)} style={{
              width: "100%", padding: "12px 24px",
              display: "flex", alignItems: "center", gap: 12,
              background: tab === item.id ? "rgba(16,185,129,0.08)" : "transparent",
              borderLeft: tab === item.id ? "2px solid #10b981" : "2px solid transparent",
              border: "none", borderRadius: 0, cursor: "pointer",
              color: tab === item.id ? "#10b981" : "#475569",
              fontSize: 13, letterSpacing: 1, fontFamily: "'Courier New', monospace",
              transition: "all 0.2s",
            }}>
              <span style={{ fontSize: 16 }}>{item.icon}</span>
              {item.label}
            </button>
          ))}
        </nav>

        {/* Bottom */}
        <div style={{ padding: "24px", borderTop: "1px solid rgba(255,255,255,0.05)" }}>
          <button onClick={() => setSoundEnabled(p => !p)} style={{
            width: "100%", padding: "8px", marginBottom: 12,
            background: soundEnabled ? "rgba(16,185,129,0.1)" : "rgba(255,255,255,0.03)",
            border: `1px solid ${soundEnabled ? "rgba(16,185,129,0.3)" : "rgba(255,255,255,0.07)"}`,
            borderRadius: 2, color: soundEnabled ? "#10b981" : "#475569",
            cursor: "pointer", fontSize: 11, letterSpacing: 1, fontFamily: "'Courier New', monospace",
          }}>
            {soundEnabled ? "🔔 SON ON" : "🔕 SON OFF"}
          </button>
          <div style={{ color: "#334155", fontSize: 11, marginBottom: 8 }}>{currentUser.name}</div>
          <button onClick={() => { setCurrentUser(null); store.set("pk_session", null); setTab("schedule"); }} style={{
            width: "100%", padding: "8px",
            background: "transparent", border: "1px solid rgba(255,255,255,0.06)",
            borderRadius: 2, color: "#334155", cursor: "pointer",
            fontSize: 11, letterSpacing: 1, fontFamily: "'Courier New', monospace",
          }}>DÉCONNEXION</button>
        </div>
      </div>

      {/* Bannière hors-ligne */}
      {offline && isConfigured && (
        <div style={{ position: "fixed", top: 0, left: 200, right: 0, zIndex: 200, background: "rgba(120,53,15,0.95)", borderBottom: "1px solid #f59e0b", padding: "8px 20px", display: "flex", alignItems: "center", gap: 8 }}>
          <span style={{ color: "#fbbf24", fontSize: 11, letterSpacing: 1 }}>⚠ Connexion Supabase perdue — mode hors-ligne. Les données ne sont pas synchronisées.</span>
        </div>
      )}

      {/* Main content */}
      <div style={{ marginLeft: 200, padding: offline && isConfigured ? "56px 40px 40px 48px" : "40px 40px 40px 48px", minHeight: "100vh" }}>
        {tab === "schedule" && <ScheduleView user={currentUser} tournaments={tournaments} sessions={sessions} setSessions={setSessions} soundEnabled={soundEnabled} alarmsEnabled={alarmsEnabled} setAlarmsEnabled={setAlarmsEnabled} />}
        {tab === "stats" && <StatsView user={currentUser} sessions={sessions} />}
        {tab === "tournaments" && isAdmin && <TournamentsAdmin tournaments={tournaments} setTournaments={setTournaments} alarmsEnabled={alarmsEnabled} setAlarmsEnabled={setAlarmsEnabled} />}
        {tab === "players" && isAdmin && <PlayersAdmin sessions={sessions} tournaments={tournaments} />}
      </div>

      <style>{`
        @keyframes fadeIn { from{opacity:0;transform:translateY(10px)} to{opacity:1;transform:translateY(0)} }
        @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.3} }
        @keyframes alarmPulse { 0%,100%{box-shadow:0 0 24px rgba(239,68,68,0.35)} 50%{box-shadow:0 0 48px rgba(239,68,68,0.7)} }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        input:focus { border-color: rgba(16,185,129,0.5) !important; }
        button:hover { filter: brightness(1.15); }
        select { color-scheme: dark; }
        ::-webkit-scrollbar { width: 4px; }
        ::-webkit-scrollbar-track { background: #0a0a0f; }
        ::-webkit-scrollbar-thumb { background: #1e293b; border-radius: 2px; }
        @media (max-width: 768px) {
          div[style*="marginLeft: 200"] { margin-left: 0 !important; padding: 80px 16px 16px !important; }
          div[style*="position: fixed"][style*="width: 200"] {
            width: 100% !important; height: 60px !important; bottom: auto !important;
            flex-direction: row !important; padding: 0 !important;
            justify-content: space-around !important; align-items: center !important;
          }
        }
      `}</style>
    </div>
  );
}

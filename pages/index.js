import React, { useState, useRef, useEffect } from "react";
import {
  Upload, Users, CheckCircle, Clock,
  Building2, TrendingUp, Calendar, X, ChevronRight,
} from "lucide-react";
import {
  BarChart, Bar, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
} from "recharts";

const EMPTY_DATA = {
  metrics: {
    totalPatients: 0, tcmScheduled: 0, notYetScheduled: 0,
    pending: 0, scheduledRate: 0, missed14DayWindow: 0,
    completedWithinWindow: 0, nurseCounts: {},
  },
  patients: [],
  practiceMetrics: { total: 0, enrolled: 0, pending: 0, declined: 0, tbd: 0 },
  practices: [],
};

export default function NurseNavigatorDashboard() {
  const [activeTab, setActiveTab] = useState("patients");
  const [data, setData]           = useState(EMPTY_DATA);
  const [lastSync, setLastSync]   = useState(null);
  const [syncing, setSyncing]     = useState(false);
  const [syncMessage, setSyncMessage] = useState("");
  const [modal, setModal]         = useState(null);
  const fileInputRef = useRef(null);

  const openModal = (title, columns, rows) => setModal({ title, columns, rows });
  const closeModal = () => setModal(null);

  const clean     = (v) => (v === null || v === undefined ? "" : String(v).trim());
  const normalize = (v) => clean(v).toLowerCase();

  const getCell = (row, names) => {
    const keys = Object.keys(row || {});
    for (const name of names) {
      const found = keys.find((k) => normalize(k) === normalize(name));
      if (found) return row[found];
    }
    for (const name of names) {
      const found = keys.find((k) => normalize(k).includes(normalize(name)));
      if (found) return row[found];
    }
    return "";
  };

  const readSheetWithHeaderRow = (sheet, XLSX, headerRowIndex, maxCols = 25) => {
    const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", blankrows: false });
    const headers = (raw[headerRowIndex] || []).slice(0, maxCols).map((h) => clean(h));
    return raw
      .slice(headerRowIndex + 1)
      .map((row) => {
        const obj = {};
        headers.forEach((h, i) => { if (h) obj[h] = row[i] ?? ""; });
        return obj;
      })
      .filter((row) => Object.values(row).some((v) => clean(v) !== ""));
  };

  const findBestHeaderRow = (sheet, XLSX) => {
    const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", blankrows: false });
    let bestIndex = 0, bestScore = -1;
    raw.slice(0, 10).forEach((row, index) => {
      const joined = row.map((c) => normalize(c)).join(" ");
      let score = 0;
      if (joined.includes("patient"))   score += 3;
      if (joined.includes("nurse"))     score += 3;
      if (joined.includes("navigator")) score += 3;
      if (joined.includes("tcm"))       score += 3;
      if (joined.includes("practice"))  score += 2;
      if (joined.includes("completed")) score += 2;
      score += row.filter((c) => clean(c) !== "").length;
      if (score > bestScore) { bestScore = score; bestIndex = index; }
    });
    return bestIndex;
  };

  const isYes = (val) => {
    const s = normalize(val);
    return ["yes","y","true","1","scheduled","complete","completed","verified","done","x","✓"].includes(s);
  };

  const processPatientRows = (patientRows) => {
    const patients = patientRows.map((p) => ({
      name:     clean(getCell(p, ["Patient Name", "Name", "Patient"])),
      practice: clean(getCell(p, ["Practice", "Practice Name"])),
      location: clean(getCell(p, ["Location", "Site", "Market"])),
      navigator: clean(getCell(p, ["Nurse Navigator", "Navigator", "Assigned Nurse", "Nurse"])) || "N/A",
      tcmScheduled:          isYes(getCell(p, ["TCM Appt Scheduled","TCM Appointment Scheduled","TCM Scheduled","TCM Appt","TCM"])),
      completedWithinWindow: isYes(getCell(p, ["Completed Within Window","Completed w/in Window","Visit Verified","Verified","Completed"])),
      missed14DayWindow:     isYes(getCell(p, ["Missed 14-Day Window","Missed 14 Day Window","Missed Window"])),
    }));

    const totalPatients         = patients.length;
    const tcmScheduled          = patients.filter((p) => p.tcmScheduled).length;
    const notYetScheduled       = totalPatients - tcmScheduled;
    const missed14DayWindow     = patients.filter((p) => p.missed14DayWindow).length;
    const completedWithinWindow = patients.filter((p) => p.completedWithinWindow).length;
    const nurseCounts = {};
    patients.forEach((p) => { nurseCounts[p.navigator] = (nurseCounts[p.navigator] || 0) + 1; });

    return {
      totalPatients, tcmScheduled, notYetScheduled,
      pending: notYetScheduled,
      scheduledRate: totalPatients > 0 ? Number(((tcmScheduled / totalPatients) * 100).toFixed(1)) : 0,
      missed14DayWindow, completedWithinWindow, nurseCounts, patients,
    };
  };

  const processPracticeRows = (practiceRows) => {
    const practices = practiceRows.map((p) => {
      const pdvStatus = clean(getCell(p, ["PDV Forms Completed","PDV Status","PDV"]));
      return {
        name:          clean(getCell(p, ["Practice Participants","Practice"])),
        consultant:    clean(getCell(p, ["Consultant"])),
        location:      clean(getCell(p, ["City","Location"])),
        hospitals:     clean(getCell(p, ["Facility Participants","Hospitals"])),
        pdvStatus,
        emrAccess:     clean(getCell(p, ["Nurse Navigator EMR Access granted","EMR Access"])),
        login:         clean(getCell(p, ["Nurse Navigator EMR Access Login"])),
        contact:       clean(getCell(p, ["Direct Office Contact","Contact"])),
        networkAccess: clean(getCell(p, ["Network Management Access"])),
        notes:         clean(getCell(p, ["Notes"])),
      };
    });

    const enrolled = practices.filter((p) => normalize(p.pdvStatus) === "complete").length;
    const declined = practices.filter((p) => normalize(p.pdvStatus).includes("declined")).length;
    const tbd      = practices.filter((p) => normalize(p.pdvStatus) === "tbd").length;
    const pending  = practices.length - enrolled - declined - tbd;

    return { total: practices.length, enrolled, pending, declined, tbd, practices };
  };

  const handleFileUpload = async (event) => {
    const files = event.target.files;
    if (!files || files.length === 0) return;

    setSyncing(true);
    setSyncMessage("Reading uploaded Excel file...");

    try {
      const XLSX  = await import("xlsx");
      let newData = JSON.parse(JSON.stringify(EMPTY_DATA));

      for (const file of files) {
        const buffer   = await file.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: "array" });
        const fileName = file.name.toLowerCase();
        const sheet    = workbook.Sheets[workbook.SheetNames[0]];

        const isPracticeFile = fileName.includes("ccpaco") || fileName.includes("practice") || fileName.includes("tracking");
        const isPatientFile  = fileName.includes("patient") || fileName.includes("patients") || fileName.includes("enrollment");

        if (isPracticeFile && !isPatientFile) {
          const practiceRows = readSheetWithHeaderRow(sheet, XLSX, 2, 10);
          const result       = processPracticeRows(practiceRows);
          newData.practiceMetrics = {
            total: result.total, enrolled: result.enrolled,
            pending: result.pending, declined: result.declined, tbd: result.tbd,
          };
          newData.practices = result.practices;
        } else {
          const headerRow   = findBestHeaderRow(sheet, XLSX);
          const patientRows = readSheetWithHeaderRow(sheet, XLSX, headerRow, 25);
          const result      = processPatientRows(patientRows);
          newData.metrics   = {
            totalPatients: result.totalPatients, tcmScheduled: result.tcmScheduled,
            notYetScheduled: result.notYetScheduled, pending: result.pending,
            scheduledRate: result.scheduledRate, missed14DayWindow: result.missed14DayWindow,
            completedWithinWindow: result.completedWithinWindow, nurseCounts: result.nurseCounts,
          };
          newData.patients = result.patients;
        }
      }

      setData(newData);
      setLastSync(new Date());
      setSyncMessage("✓ Sync complete. Dashboard updated from uploaded file.");
    } catch (error) {
      console.error(error);
      setSyncMessage("Error reading file: " + error.message);
    }

    setTimeout(() => { setSyncing(false); setSyncMessage(""); }, 3000);
    event.target.value = "";
  };

  return (
    <div style={pageStyle}>
      <input ref={fileInputRef} type="file" accept=".xlsx,.xls" multiple onChange={handleFileUpload} style={{ display: "none" }} />

      {modal && <Modal title={modal.title} columns={modal.columns} rows={modal.rows} onClose={closeModal} />}

      <header style={headerStyle}>
        <div style={headerInnerStyle}>
          <h1 style={titleStyle}>Nurse Navigator Program</h1>
          <div style={{ display: "flex", gap: "12px", alignItems: "center" }}>
            <button onClick={() => fileInputRef.current.click()} disabled={syncing} style={buttonStyle}>
              <Upload size={18} />
              {syncing ? "Syncing..." : "Sync from SharePoint"}
            </button>
            <div style={liveStyle}>
              Live Data<br />
              <span style={{ fontWeight: "400" }}>
                {lastSync ? `Updated ${lastSync.toLocaleTimeString()}` : "No file uploaded"}
              </span>
            </div>
          </div>
        </div>

        {syncMessage && <div style={messageStyle}>{syncMessage}</div>}

        <div style={tabsStyle}>
          <TabButton active={activeTab === "patients"}  onClick={() => setActiveTab("patients")}  icon={<Users size={18} />}    label="Patient Tracking"    />
          <TabButton active={activeTab === "practices"} onClick={() => setActiveTab("practices")} icon={<Building2 size={18} />} label="Practice Enrollment" />
        </div>
      </header>

      <main style={mainStyle}>
        {activeTab === "patients"
          ? <PatientTrackingPage   data={data} openModal={openModal} />
          : <PracticeEnrollmentPage data={data} openModal={openModal} />}
      </main>
    </div>
  );
}

// ─── MODAL ───────────────────────────────────────────────────
function Modal({ title, columns, rows, onClose }) {
  useEffect(() => {
    const handleKey = (e) => { if (e.key === "Escape") onClose(); };
    window.addEventListener("keydown", handleKey);
    return () => window.removeEventListener("keydown", handleKey);
  }, [onClose]);

  return (
    <div
      onClick={onClose}
      style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.55)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: "24px" }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{ background: "white", borderRadius: "20px", width: "100%", maxWidth: "900px", maxHeight: "80vh", display: "flex", flexDirection: "column", boxShadow: "0 25px 50px rgba(0,0,0,0.3)", overflow: "hidden" }}
      >
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "20px 28px", borderBottom: "2px solid #e2e8f0", background: "linear-gradient(135deg, #667eea, #764ba2)" }}>
          <h2 style={{ margin: 0, color: "white", fontSize: "20px", fontWeight: "800" }}>{title}</h2>
          <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
            <span style={{ color: "rgba(255,255,255,0.8)", fontSize: "13px" }}>{rows.length} records</span>
            <button onClick={onClose} style={{ background: "rgba(255,255,255,0.2)", border: "none", borderRadius: "8px", padding: "6px", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", color: "white" }}>
              <X size={18} />
            </button>
          </div>
        </div>

        <div style={{ overflowY: "auto", padding: "8px 0" }}>
          {rows.length === 0 ? (
            <div style={{ padding: "60px", textAlign: "center", color: "#94a3b8" }}>
              <p style={{ fontSize: "16px" }}>No records to show.</p>
            </div>
          ) : (
            <table style={{ ...tableStyle, margin: 0 }}>
              <thead style={{ position: "sticky", top: 0, background: "#f8fafc" }}>
                <tr>{columns.map((c) => <th key={c} style={thStyle}>{c}</th>)}</tr>
              </thead>
              <tbody>
                {rows.map((row, i) => (
                  <tr key={i} style={{ background: i % 2 === 0 ? "white" : "#f8fafc" }}>
                    {row.map((cell, j) => (
                      <td key={j} style={tdStyle}>{typeof cell === "object" ? cell : (cell || "N/A")}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>

        <div style={{ padding: "16px 28px", borderTop: "1px solid #e2e8f0", textAlign: "right", background: "#f8fafc" }}>
          <button onClick={onClose} style={{ padding: "10px 24px", background: "linear-gradient(135deg,#667eea,#764ba2)", color: "white", border: "none", borderRadius: "10px", fontWeight: "700", cursor: "pointer", fontSize: "14px" }}>
            Close
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── TAB BUTTON ──────────────────────────────────────────────
function TabButton({ active, onClick, icon, label }) {
  return (
    <button onClick={onClick} style={{ padding: "12px 24px", background: active ? "linear-gradient(135deg,#667eea,#764ba2)" : "white", color: active ? "white" : "#64748b", border: active ? "none" : "2px solid #e2e8f0", borderRadius: "10px", fontWeight: "700", cursor: "pointer", display: "flex", alignItems: "center", gap: "8px" }}>
      {icon}{label}
    </button>
  );
}

// ─── PATIENT PAGE ─────────────────────────────────────────────
function PatientTrackingPage({ data, openModal }) {
  const m        = data.metrics;
  const patients = data.patients;

  const [search,    setSearch]    = useState("");
  const [tcmFilter, setTcmFilter] = useState("all");
  const [navFilter, setNavFilter] = useState("all");
  const [sortCol,   setSortCol]   = useState(null);
  const [sortDir,   setSortDir]   = useState("asc");

  const PATIENT_COLS = ["Patient Name","Practice","Location","Navigator","TCM Status","Completed","Missed Window"];

  const toRows = (list) => list.map((p) => [
    p.name, p.practice, p.location, p.navigator,
    <Badge key="tcm"  text={p.tcmScheduled ? "✓ Scheduled" : "⏳ Not Yet"} bg={p.tcmScheduled ? "#dcfce7" : "#fef9c3"} color={p.tcmScheduled ? "#15803d" : "#b45309"} />,
    <Badge key="comp" text={p.completedWithinWindow ? "✓ Yes" : "–"} bg={p.completedWithinWindow ? "#ede9fe" : "#f1f5f9"} color={p.completedWithinWindow ? "#7c3aed" : "#94a3b8"} />,
    p.missed14DayWindow ? <Badge key="miss" text="⚠ Missed" bg="#fee2e2" color="#b91c1c" /> : <span style={{ color: "#94a3b8" }}>–</span>,
  ]);

  const navigators = ["all", ...Array.from(new Set(patients.map((p) => p.navigator).filter(Boolean)))];

  const pendingPatients    = patients.filter((p) => !p.tcmScheduled);
  const scheduledPatients  = patients.filter((p) => p.tcmScheduled);
  const missedPatients     = patients.filter((p) => p.missed14DayWindow);

  const filtered = patients
    .filter((p) => {
      const q = search.toLowerCase();
      const matchSearch = !search || p.name.toLowerCase().includes(q) || p.practice.toLowerCase().includes(q) || p.navigator.toLowerCase().includes(q);
      const matchTcm = tcmFilter === "all" || (tcmFilter === "scheduled" && p.tcmScheduled) || (tcmFilter === "not_scheduled" && !p.tcmScheduled);
      const matchNav = navFilter === "all" || p.navigator === navFilter;
      return matchSearch && matchTcm && matchNav;
    })
    .sort((a, b) => {
      if (!sortCol) return 0;
      const va = (a[sortCol] ?? "").toString().toLowerCase();
      const vb = (b[sortCol] ?? "").toString().toLowerCase();
      return sortDir === "asc" ? va.localeCompare(vb) : vb.localeCompare(va);
    });

  const handleSort = (col) => {
    if (sortCol === col) setSortDir((d) => (d === "asc" ? "desc" : "asc"));
    else { setSortCol(col); setSortDir("asc"); }
  };

  const pieData = [
    { name: "Scheduled",         value: m.tcmScheduled,    color: "#10b981" },
    { name: "Not Yet Scheduled", value: m.notYetScheduled, color: "#f59e0b" },
  ];

  const nurseData = Object.entries(m.nurseCounts || {}).map(([name, count]) => ({ name, count }));

  return (
    <div>
      {/* ── KPI CARDS ── */}
      <div style={gridStyle}>
        <MetricCard
          icon={<Users />} title="Total Patients" value={m.totalPatients}
          subtitle="Click to view all" color="#3b82f6"
          onClick={() => openModal(`All Patients (${patients.length})`, PATIENT_COLS, toRows(patients))}
        />
        <MetricCard
          icon={<Calendar />} title="TCM Appt Scheduled" value={m.tcmScheduled}
          subtitle={`${m.scheduledRate}% · Click to view`} color="#10b981"
          onClick={() => openModal(`TCM Scheduled (${scheduledPatients.length})`, PATIENT_COLS, toRows(scheduledPatients))}
        />
        <MetricCard
          icon={<Clock />} title="Not Yet Scheduled" value={m.notYetScheduled}
          subtitle="Click to view" color="#f59e0b"
          onClick={() => openModal(`Not Yet Scheduled (${pendingPatients.length})`, PATIENT_COLS, toRows(pendingPatients))}
        />
        <MetricCard
          icon={<TrendingUp />} title="Missed 14-Day Window" value={m.missed14DayWindow}
          subtitle="Click to view" color="#ef4444"
          onClick={() => openModal(`Missed 14-Day Window (${missedPatients.length})`, PATIENT_COLS, toRows(missedPatients))}
        />
        {/* ── PENDING card (replaces Completed Within Window) ── */}
        <MetricCard
          icon={<Clock />} title="Pending" value={m.pending}
          subtitle="Awaiting action · Click to view" color="#f59e0b"
          onClick={() => openModal(`Pending Patients (${pendingPatients.length})`, PATIENT_COLS, toRows(pendingPatients))}
        />
      </div>

      {/* ── CHARTS ── */}
      <div style={chartGridStyle}>
        <ChartCard title="TCM Scheduling Status">
          <ResponsiveContainer width="100%" height={280}>
            <PieChart>
              <Pie data={pieData} dataKey="value" outerRadius={100} label>
                {pieData.map((e, i) => <Cell key={i} fill={e.color} />)}
              </Pie>
              <Tooltip /><Legend />
            </PieChart>
          </ResponsiveContainer>
        </ChartCard>

        <ChartCard title="Nurse Counts">
          <ResponsiveContainer width="100%" height={280}>
            <BarChart data={nurseData}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="name" /><YAxis allowDecimals={false} />
              <Tooltip /><Legend />
              <Bar dataKey="count" fill="#667eea" name="Patients" />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>
      </div>

      {/* ── INTERACTIVE DRILLDOWN ── */}
      <div style={tableWrapStyle}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "16px", marginBottom: "16px" }}>
          <h2 style={{ margin: 0 }}>
            Patient Drilldown{" "}
            <span style={{ fontSize: "14px", color: "#64748b", fontWeight: "400" }}>
              ({filtered.length} of {patients.length})
            </span>
          </h2>

          <div style={{ display: "flex", gap: "8px", flexWrap: "wrap" }}>
            {[
              { label: "Scheduled",     list: scheduledPatients,                                  color: "#15803d", bg: "#dcfce7" },
              { label: "Not Scheduled", list: pendingPatients,                                    color: "#b45309", bg: "#fef9c3" },
              { label: "Completed",     list: patients.filter((p) => p.completedWithinWindow),    color: "#7c3aed", bg: "#ede9fe" },
              { label: "Missed Window", list: missedPatients,                                     color: "#b91c1c", bg: "#fee2e2" },
            ].map((b) => (
              <span
                key={b.label}
                onClick={() => openModal(`${b.label} (${b.list.length})`, PATIENT_COLS, toRows(b.list))}
                style={{ padding: "4px 14px", borderRadius: "20px", background: b.bg, color: b.color, fontSize: "12px", fontWeight: "700", cursor: "pointer" }}
              >
                {b.label}: {b.list.length}
              </span>
            ))}
          </div>
        </div>

        {/* Filters */}
        <div style={{ display: "flex", gap: "12px", flexWrap: "wrap", marginBottom: "16px" }}>
          <input
            placeholder="Search name, practice, navigator…"
            value={search} onChange={(e) => setSearch(e.target.value)}
            style={{ ...inputStyle, minWidth: "260px", flex: 1 }}
          />
          <select value={tcmFilter} onChange={(e) => setTcmFilter(e.target.value)} style={inputStyle}>
            <option value="all">All TCM Status</option>
            <option value="scheduled">Scheduled</option>
            <option value="not_scheduled">Not Yet Scheduled</option>
          </select>
          <select value={navFilter} onChange={(e) => setNavFilter(e.target.value)} style={inputStyle}>
            {navigators.map((n) => (
              <option key={n} value={n}>{n === "all" ? "All Navigators" : n}</option>
            ))}
          </select>
          {(search || tcmFilter !== "all" || navFilter !== "all") && (
            <button
              onClick={() => { setSearch(""); setTcmFilter("all"); setNavFilter("all"); }}
              style={{ ...inputStyle, background: "#f1f5f9", cursor: "pointer", border: "2px solid #e2e8f0", color: "#64748b", whiteSpace: "nowrap" }}
            >
              ✕ Clear
            </button>
          )}
        </div>

        {/* Table */}
        <div style={{ overflowX: "auto" }}>
          <table style={tableStyle}>
            <thead>
              <tr>
                {[
                  { label: "Patient Name", col: "name" },
                  { label: "Practice",     col: "practice" },
                  { label: "Location",     col: "location" },
                  { label: "Navigator",    col: "navigator" },
                  { label: "TCM Status",   col: null },
                  { label: "Completed",    col: null },
                  { label: "Missed Window",col: null },
                ].map(({ label, col }) => (
                  <th
                    key={label}
                    onClick={() => col && handleSort(col)}
                    style={{ ...thStyle, cursor: col ? "pointer" : "default" }}
                  >
                    {label}{col ? (sortCol === col ? (sortDir === "asc" ? " ↑" : " ↓") : " ↕") : ""}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 ? (
                <tr>
                  <td colSpan={7} style={{ ...tdStyle, textAlign: "center", color: "#94a3b8", padding: "40px" }}>
                    No patients match the current filters.
                  </td>
                </tr>
              ) : (
                filtered.map((p, i) => (
                  <tr
                    key={i}
                    style={{ background: i % 2 === 0 ? "white" : "#f8fafc", cursor: "pointer" }}
                    onClick={() => openModal(`Patient: ${p.name}`, ["Field", "Value"], [
                      ["Name",          p.name],
                      ["Practice",      p.practice],
                      ["Location",      p.location],
                      ["Navigator",     p.navigator],
                      ["TCM Scheduled", p.tcmScheduled ? "✓ Yes" : "✗ No"],
                      ["Completed",     p.completedWithinWindow ? "✓ Yes" : "–"],
                      ["Missed Window", p.missed14DayWindow ? "⚠ Yes" : "–"],
                    ])}
                  >
                    <td style={{ ...tdStyle, fontWeight: "700" }}>{p.name || "N/A"}</td>
                    <td style={tdStyle}>{p.practice || "N/A"}</td>
                    <td style={tdStyle}>{p.location || "N/A"}</td>
                    <td style={tdStyle}>{p.navigator || "N/A"}</td>
                    <td style={tdStyle}>
                      <Badge text={p.tcmScheduled ? "✓ Scheduled" : "⏳ Not Yet"} bg={p.tcmScheduled ? "#dcfce7" : "#fef9c3"} color={p.tcmScheduled ? "#15803d" : "#b45309"} />
                    </td>
                    <td style={tdStyle}>
                      <Badge text={p.completedWithinWindow ? "✓ Yes" : "–"} bg={p.completedWithinWindow ? "#ede9fe" : "#f1f5f9"} color={p.completedWithinWindow ? "#7c3aed" : "#94a3b8"} />
                    </td>
                    <td style={tdStyle}>
                      {p.missed14DayWindow
                        ? <Badge text="⚠ Missed" bg="#fee2e2" color="#b91c1c" />
                        : <span style={{ color: "#94a3b8" }}>–</span>}
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ─── PRACTICE PAGE ────────────────────────────────────────────
function PracticeEnrollmentPage({ data, openModal }) {
  const m         = data.practiceMetrics;
  const practices = data.practices;
  const [filter, setFilter] = useState("all");
  const [search, setSearch] = useState("");

  const PRACTICE_COLS = ["Practice","Consultant","City","Facility","PDV Status","EMR Access","Contact"];
  const toRows = (list) => list.map((p) => [p.name, p.consultant, p.location, p.hospitals, p.pdvStatus, p.emrAccess, p.contact]);

  const enrolledPractices  = practices.filter((p) => normalize(p.pdvStatus) === "complete");
  const declinedPractices  = practices.filter((p) => normalize(p.pdvStatus).includes("declined"));
  const tbdPractices       = practices.filter((p) => normalize(p.pdvStatus) === "tbd");
  const pendingPractices   = practices.filter((p) => {
    const s = normalize(p.pdvStatus);
    return s !== "complete" && !s.includes("declined") && s !== "tbd";
  });

  const normalize = (v) => (v === null || v === undefined ? "" : String(v).trim().toLowerCase());

  const statusData = [
    { name: "Complete", value: m.enrolled,  color: "#10b981" },
    { name: "Pending",  value: m.pending,   color: "#f59e0b" },
    { name: "Declined", value: m.declined,  color: "#ef4444" },
    { name: "TBD",      value: m.tbd,       color: "#8b5cf6" },
  ];

  const consultantMap = {};
  practices.forEach((p) => {
    const c = p.consultant || "N/A";
    if (!consultantMap[c]) consultantMap[c] = { name: c, complete: 0, pending: 0 };
    const s = (p.pdvStatus || "").toLowerCase();
    if (s === "complete") consultantMap[c].complete += 1;
    else consultantMap[c].pending += 1;
  });
  const consultantData = Object.values(consultantMap);

  const filtered = practices.filter((p) => {
    const searchMatch = !search ||
      p.name.toLowerCase().includes(search.toLowerCase()) ||
      p.consultant.toLowerCase().includes(search.toLowerCase());
    if (!searchMatch) return false;
    const s = (p.pdvStatus || "").toLowerCase();
    if (filter === "complete") return s === "complete";
    if (filter === "declined") return s.includes("declined");
    if (filter === "tbd")      return s === "tbd";
    if (filter === "pending")  return s !== "complete" && !s.includes("declined") && s !== "tbd";
    return true;
  });

  return (
    <div>
      {/* ── KPI CARDS ── */}
      <div style={gridStyle}>
        <MetricCard icon={<Building2 />}   title="Total Practices" value={m.total}
          subtitle="Click to view all" color="#3b82f6"
          onClick={() => openModal(`All Practices (${practices.length})`, PRACTICE_COLS, toRows(practices))} />
        <MetricCard icon={<CheckCircle />} title="Enrolled" value={m.enrolled}
          subtitle={`${m.total ? ((m.enrolled/m.total)*100).toFixed(1) : 0}% · Click to view`} color="#10b981"
          onClick={() => openModal(`Enrolled Practices (${enrolledPractices.length})`, PRACTICE_COLS, toRows(enrolledPractices))} />
        <MetricCard icon={<Clock />}       title="Pending" value={m.pending}
          subtitle="Click to view" color="#f59e0b"
          onClick={() => openModal(`Pending Practices (${pendingPractices.length})`, PRACTICE_COLS, toRows(pendingPractices))} />
        <MetricCard icon={<TrendingUp />}  title="Declined" value={m.declined}
          subtitle="Click to view" color="#ef4444"
          onClick={() => openModal(`Declined Practices (${declinedPractices.length})`, PRACTICE_COLS, toRows(declinedPractices))} />
        <MetricCard icon={<Clock />}       title="TBD" value={m.tbd}
          subtitle="Click to view" color="#8b5cf6"
          onClick={() => openModal(`TBD Practices (${tbdPractices.length})`, PRACTICE_COLS, toRows(tbdPractices))} />
      </div>

      {/* ── CHARTS ── */}
      <div style={chartGridStyle}>
        <ChartCard title="Enrollment Status Breakdown">
          <ResponsiveContainer width="100%" height={300}>
            <PieChart>
              <Pie data={statusData} dataKey="value" outerRadius={110} label>
                {statusData.map((e, i) => <Cell key={i} fill={e.color} />)}
              </Pie>
              <Tooltip /><Legend />
            </PieChart>
          </ResponsiveContainer>
        </ChartCard>

        <ChartCard title="Practice Enrolled by Consultant">
          <ResponsiveContainer width="100%" height={300}>
            <BarChart data={consultantData} layout="vertical" margin={{ left: 140 }}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis type="number" allowDecimals={false} />
              <YAxis dataKey="name" type="category" width={130} />
              <Tooltip /><Legend />
              <Bar dataKey="complete" fill="#10b981" name="Complete"
                onClick={(d) => openModal(`${d.name} — Complete`, PRACTICE_COLS, toRows(practices.filter((p) => p.consultant === d.name && (p.pdvStatus || "").toLowerCase() === "complete")))} />
              <Bar dataKey="pending"  fill="#f59e0b" name="Pending"
                onClick={(d) => openModal(`${d.name} — Pending`,  PRACTICE_COLS, toRows(practices.filter((p) => p.consultant === d.name && (p.pdvStatus || "").toLowerCase() !== "complete")))} />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>
      </div>

      {/* ── PRACTICE DRILLDOWN ── */}
      <div style={tableWrapStyle}>
        <div style={tableHeaderStyle}>
          <h2>Practice Drilldown ({filtered.length} of {practices.length})</h2>
          <div style={{ display: "flex", gap: "12px" }}>
            <input
              placeholder="Search practice or consultant..."
              value={search} onChange={(e) => setSearch(e.target.value)}
              style={inputStyle}
            />
            <select value={filter} onChange={(e) => setFilter(e.target.value)} style={inputStyle}>
              <option value="all">All Status</option>
              <option value="complete">Complete</option>
              <option value="pending">Pending</option>
              <option value="declined">Declined</option>
              <option value="tbd">TBD</option>
            </select>
          </div>
        </div>

        <div style={{ overflowX: "auto" }}>
          <table style={tableStyle}>
            <thead>
              <tr>{PRACTICE_COLS.map((c) => <th key={c} style={thStyle}>{c}</th>)}</tr>
            </thead>
            <tbody>
              {filtered.length === 0 ? (
                <tr>
                  <td colSpan={7} style={{ ...tdStyle, textAlign: "center", color: "#94a3b8", padding: "40px" }}>
                    No data. Upload your Excel files.
                  </td>
                </tr>
              ) : (
                filtered.map((p, i) => (
                  <tr
                    key={i}
                    style={{ background: i % 2 === 0 ? "white" : "#f8fafc", cursor: "pointer" }}
                    onClick={() => openModal(`Practice: ${p.name}`, ["Field", "Value"], [
                      ["Practice Name",   p.name],
                      ["Consultant",      p.consultant],
                      ["City",            p.location],
                      ["Facility",        p.hospitals],
                      ["PDV Status",      p.pdvStatus],
                      ["EMR Access",      p.emrAccess],
                      ["Login",           p.login],
                      ["Contact",         p.contact],
                      ["Network Access",  p.networkAccess],
                      ["Notes",           p.notes],
                    ])}
                  >
                    <td style={{ ...tdStyle, fontWeight: "700" }}>{p.name || "N/A"}</td>
                    <td style={tdStyle}>{p.consultant || "N/A"}</td>
                    <td style={tdStyle}>{p.location || "N/A"}</td>
                    <td style={{ ...tdStyle, maxWidth: "180px", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{p.hospitals || "N/A"}</td>
                    <td style={tdStyle}>
                      <Badge
                        text={p.pdvStatus || "N/A"}
                        bg={(p.pdvStatus||"").toLowerCase()==="complete" ? "#dcfce7" : (p.pdvStatus||"").toLowerCase().includes("declined") ? "#fee2e2" : (p.pdvStatus||"").toLowerCase()==="tbd" ? "#ede9fe" : "#fef9c3"}
                        color={(p.pdvStatus||"").toLowerCase()==="complete" ? "#15803d" : (p.pdvStatus||"").toLowerCase().includes("declined") ? "#b91c1c" : (p.pdvStatus||"").toLowerCase()==="tbd" ? "#7c3aed" : "#b45309"}
                      />
                    </td>
                    <td style={tdStyle}>{p.emrAccess || "N/A"}</td>
                    <td style={tdStyle}>{p.contact || "N/A"}</td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ─── SHARED COMPONENTS ────────────────────────────────────────
function MetricCard({ icon, title, value, subtitle, color, onClick }) {
  const [hovered, setHovered] = useState(false);
  return (
    <div
      onClick={onClick}
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
      style={{
        ...cardStyle,
        cursor: onClick ? "pointer" : "default",
        transform: hovered && onClick ? "translateY(-6px)" : "none",
        boxShadow: hovered && onClick ? `0 12px 24px ${color}33` : "0 4px 8px rgba(0,0,0,0.08)",
        transition: "all 0.2s ease",
        border: hovered && onClick ? `2px solid ${color}44` : "2px solid transparent",
      }}
    >
      <div style={{ width: "48px", height: "48px", borderRadius: "12px", background: color, color: "white", display: "flex", alignItems: "center", justifyContent: "center", marginBottom: "16px" }}>
        {icon}
      </div>
      <div style={{ fontSize: "14px", color: "#475569", fontWeight: "700" }}>{title}</div>
      <div style={{ fontSize: "34px", fontWeight: "800", marginTop: "8px" }}>{value}</div>
      <div style={{ fontSize: "12px", color: hovered && onClick ? color : "#64748b", marginTop: "6px", display: "flex", alignItems: "center", gap: "4px" }}>
        {subtitle}{onClick && hovered && <ChevronRight size={12} />}
      </div>
    </div>
  );
}

function Badge({ text, bg, color }) {
  return (
    <span style={{ display: "inline-block", padding: "4px 12px", borderRadius: "20px", fontSize: "12px", fontWeight: "700", background: bg, color }}>
      {text}
    </span>
  );
}

function ChartCard({ title, children }) {
  return (
    <div style={cardStyle}>
      <h3>{title}</h3>
      {children}
    </div>
  );
}

// ─── STYLES ───────────────────────────────────────────────────
const pageStyle        = { minHeight: "100vh", background: "linear-gradient(135deg, #667eea 0%, #764ba2 100%)", fontFamily: "Arial, sans-serif" };
const headerStyle      = { background: "rgba(255,255,255,0.96)", padding: "24px 40px", boxShadow: "0 4px 8px rgba(0,0,0,0.08)" };
const headerInnerStyle = { maxWidth: "1400px", margin: "0 auto", display: "flex", justifyContent: "space-between", alignItems: "center", gap: "16px", flexWrap: "wrap" };
const titleStyle       = { fontSize: "30px", fontWeight: "800", margin: 0, color: "#5b5fc7" };
const buttonStyle      = { padding: "12px 24px", background: "linear-gradient(135deg,#667eea,#764ba2)", color: "white", border: "none", borderRadius: "10px", fontWeight: "700", cursor: "pointer", display: "flex", alignItems: "center", gap: "8px" };
const liveStyle        = { padding: "10px 18px", background: "#f0fdf4", border: "2px solid #86efac", borderRadius: "12px", color: "#166534", fontSize: "12px", fontWeight: "700" };
const messageStyle     = { maxWidth: "1400px", margin: "16px auto 0", padding: "12px", background: "#dbeafe", borderRadius: "10px", color: "#1e40af", textAlign: "center", fontWeight: "700" };
const tabsStyle        = { maxWidth: "1400px", margin: "24px auto 0", display: "flex", gap: "10px" };
const mainStyle        = { maxWidth: "1400px", margin: "0 auto", padding: "40px" };
const gridStyle        = { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: "24px", marginBottom: "32px" };
const chartGridStyle   = { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(420px, 1fr))", gap: "24px", marginBottom: "24px" };
const cardStyle        = { background: "white", borderRadius: "16px", padding: "24px", boxShadow: "0 4px 8px rgba(0,0,0,0.08)" };
const tableWrapStyle   = { background: "white", borderRadius: "16px", padding: "24px", boxShadow: "0 4px 8px rgba(0,0,0,0.08)", marginTop: "24px" };
const tableHeaderStyle = { display: "flex", justifyContent: "space-between", gap: "12px", flexWrap: "wrap", marginBottom: "16px" };
const tableStyle       = { width: "100%", borderCollapse: "collapse", marginTop: "16px" };
const thStyle          = { padding: "12px", textAlign: "left", fontSize: "12px", fontWeight: "800", color: "#475569", textTransform: "uppercase", borderBottom: "2px solid #e2e8f0" };
const tdStyle          = { padding: "14px 12px", fontSize: "14px", color: "#1e293b", borderBottom: "1px solid #f1f5f9" };
const inputStyle       = { padding: "10px 14px", border: "2px solid #e2e8f0", borderRadius: "8px", fontSize: "14px" };

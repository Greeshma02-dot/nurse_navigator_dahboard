import React, { useState, useRef, useEffect } from "react";
import {
  Upload, Users, CheckCircle, Clock,
  Building2, TrendingUp, Calendar, X, ChevronRight, AlertTriangle, Activity,
} from "lucide-react";
import {
  BarChart, Bar, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
} from "recharts";

const EMPTY_DATA = {
  metrics: {
    totalPatients: 0, currentlyAdmitted: 0, discharged: 0,
    tcmScheduled: 0, tcmPending: 0, notYetScheduled: 0,
    visitVerified: 0, missed14DayWindow: 0, scheduledRate: 0, nurseCounts: {},
  },
  patients: [],
  practiceMetrics: { total: 0, enrolled: 0, pending: 0, declined: 0, tbd: 0, emrComplete: 0 },
  practices: [],
};

export default function NurseNavigatorDashboard() {
  const [activeTab, setActiveTab] = useState("patients");
  const [data, setData]           = useState(EMPTY_DATA);
  const [lastSync, setLastSync]   = useState(null);
  const [syncing, setSyncing]     = useState(false);
  const [syncMessage, setSyncMessage] = useState("");
  const [modal, setModal]         = useState(null);
  const [showLagWarning, setShowLagWarning] = useState(false);
  const fileInputRef = useRef(null);

  useEffect(() => {
    if (new Date().getDay() === 1) setShowLagWarning(true);
  }, []);

  const openModal  = (title, columns, rows) => setModal({ title, columns, rows });
  const closeModal = () => setModal(null);

  const clean     = (v) => (v === null || v === undefined ? "" : String(v).trim());
  const normalize = (v) => clean(v).toLowerCase();

  const getCell = (row, names) => {
    const keys = Object.keys(row || {});
    for (const name of names) {
      const found = keys.find((k) => normalize(k) === normalize(name));
      if (found !== undefined) return row[found];
    }
    for (const name of names) {
      const found = keys.find((k) => normalize(k).includes(normalize(name)));
      if (found !== undefined) return row[found];
    }
    return "";
  };

  // ── FIXED DATE PARSER ──────────────────────────────────────
  // Handles: JS Date objects, Excel serial numbers, ISO strings, 
  // "N/A see notes", empty values — all safely
  const parseDate = (val) => {
    if (val === null || val === undefined || val === "") return null;

    // Already a proper JS Date
    if (val instanceof Date) {
      return isNaN(val.getTime()) ? null : val;
    }

    // Excel serial number → JS Date
    // Formula: (serial - 25569) * 86400000 ms
    if (typeof val === "number") {
      if (val <= 0 || val > 200000) return null;       // out of realistic range
      const d = new Date(Math.round((val - 25569) * 86400 * 1000));
      return isNaN(d.getTime()) ? null : d;
    }

    // String value
    if (typeof val === "string") {
      const s = val.trim();
      if (!s) return null;
      const lower = s.toLowerCase();
      // Reject non-date strings
      if (
        lower === "n/a" || lower.includes("see notes") ||
        lower.includes("tbd") || lower.includes("pending")
      ) return null;

      // Try to parse as date string
      const d = new Date(s);
      return isNaN(d.getTime()) ? null : d;
    }

    return null;
  };

  const formatDate = (d) => {
    if (!d) return "N/A";
    return d.toLocaleDateString("en-US", { month: "2-digit", day: "2-digit", year: "numeric" });
  };

  const daysBetween = (d1, d2) => {
    if (!d1 || !d2) return null;
    return Math.round((d2 - d1) / (1000 * 60 * 60 * 24));
  };

  const isYes = (val) => {
    const s = normalize(val);
    return ["yes","y","true","1","complete","completed","verified","done","x","✓"].includes(s);
  };

  const readSheetWithHeaderRow = (sheet, XLSX, headerRowIndex, maxCols = 16) => {
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
      if (joined.includes("patient"))          score += 4;
      if (joined.includes("navigator assign")) score += 4;
      if (joined.includes("tcm"))              score += 4;
      if (joined.includes("discharge"))        score += 3;
      if (joined.includes("practice participants")) score += 5;
      if (joined.includes("pdv"))              score += 5;
      if (joined.includes("practice"))         score += 2;
      if (joined.includes("consultant"))       score += 3;
      if (joined.includes("verified"))         score += 2;
      score += row.filter((c) => clean(c) !== "").length;
      if (score > bestScore) { bestScore = score; bestIndex = index; }
    });
    return bestIndex;
  };

  const processPatientRows = (patientRows) => {
    const today = new Date();

    const patients = patientRows.map((p) => {
      // ── EXACT COLUMN NAMES FROM YOUR EXCEL ──
      const anticipatedDischarge = parseDate(getCell(p, ["Anticipated Discharge","Anticipated Discharg","Anticipated"]));
      const actualDischarge      = parseDate(getCell(p, ["Actual Discharge Date","Actual Discharge Dat","Actual Discharge","Discharge Date","Discharged","Discharge"]));
      const tcmRaw               = clean(getCell(p, ["TCM Appt Scheduled?","TCM Appt Scheduled","TCM Appt Sch","TCM Scheduled","TCM Appt","TCM"]));
      const tcmDate              = parseDate(getCell(p, ["TCM Appt Date","TCM Appt Dat","TCM Date"]));
      const apptType             = clean(getCell(p, ["Appt Type","Appointment Type"]));
      const verifiedRaw          = clean(getCell(p, ["Visit Verified","Verified"]));
      const windowStatus         = clean(getCell(p, ["14-Day Window Status","14-Dag Window St","14 Day Window","14-Day Window","Missed 14-Day Window","Missed Window"]));
      const call2                = clean(getCell(p, ["2-Day Call Attempt","2-Dag Call Atte","2 Day Call"]));
      const call7                = clean(getCell(p, ["7-Day Call Attempt","7-Dag Call Atte","7 Day Call"]));
      const navigator            = clean(getCell(p, ["Navigator Assigned","Navigator Assign","Nurse Navigator","Navigator","Assigned Nurse"]));
      const practice             = clean(getCell(p, ["Practice"]));
      const location             = clean(getCell(p, ["Location","Facility","Site"]));
      const room                 = clean(getCell(p, ["Room #","Room"]));
      const notes                = clean(getCell(p, ["Notes"]));

      const tcmScheduled = isYes(tcmRaw);
      const tcmPending   = normalize(tcmRaw) === "pending";
      const visitVerified = isYes(verifiedRaw);
      const missed14Day  = normalize(windowStatus).includes("missed") || normalize(windowStatus).includes("no");

      // Days since actual discharge
      const daysSinceDischarge = actualDischarge ? daysBetween(actualDischarge, today) : null;

      // Status logic
      let status = "Unknown";
      if (visitVerified)                      status = "Visit Verified";
      else if (tcmScheduled)                  status = "Scheduled";
      else if (tcmPending)                    status = "Pending";
      else if (actualDischarge && !tcmScheduled) status = "Discharged - No TCM";
      else if (!actualDischarge)              status = "Admitted";

      return {
        name:                  clean(getCell(p, ["Patient Name","Name","Patient"])),
        practice,
        location,
        room,
        navigator:             navigator || "N/A",
        anticipatedDischarge,
        actualDischarge,
        anticipatedStr:        formatDate(anticipatedDischarge),
        actualDischargeStr:    formatDate(actualDischarge),
        tcmRaw,
        tcmScheduled,
        tcmPending,
        tcmDate,
        tcmDateStr:            formatDate(tcmDate),
        apptType,
        visitVerified,
        missed14Day,
        windowStatus,
        call2,
        call7,
        daysSinceDischarge,
        status,
        notes,
      };
    }).filter((p) => p.name); // remove empty rows

    const totalPatients     = patients.length;
    const tcmScheduled      = patients.filter((p) => p.tcmScheduled).length;
    const tcmPendingCount   = patients.filter((p) => p.tcmPending).length;
    const visitVerified     = patients.filter((p) => p.visitVerified).length;
    const missed14Day       = patients.filter((p) => p.missed14Day).length;
    const notYetScheduled   = patients.filter((p) => !p.tcmScheduled && !p.tcmPending && !p.visitVerified).length;
    const currentlyAdmitted = patients.filter((p) => !p.actualDischarge).length;
    const discharged        = patients.filter((p) => p.actualDischarge && !p.visitVerified).length;

    const nurseCounts = {};
    patients.forEach((p) => {
      if (p.navigator && p.navigator !== "N/A") {
        nurseCounts[p.navigator] = (nurseCounts[p.navigator] || 0) + 1;
      }
    });

    return {
      totalPatients, tcmScheduled, tcmPending: tcmPendingCount,
      visitVerified, missed14Day, missed14DayWindow: missed14Day,
      notYetScheduled, currentlyAdmitted, discharged,
      scheduledRate: totalPatients > 0 ? Number(((tcmScheduled / totalPatients) * 100).toFixed(1)) : 0,
      nurseCounts, patients,
    };
  };

  const processPracticeRows = (practiceRows) => {
    const practices = practiceRows.map((p) => {
      const pdvStatus  = clean(getCell(p, ["PDV Forms Completed","PDV Status","PDV"]));
      const emrGranted = clean(getCell(p, ["Nurse Navigator EMR Access granted","EMR Access granted","EMR Access","EMR"]));
      return {
        name:          clean(getCell(p, ["Practice Participants","Practice"])),
        consultant:    clean(getCell(p, ["Consultant"])),
        location:      clean(getCell(p, ["City","Location"])),
        hospitals:     clean(getCell(p, ["Facility Participants","Hospitals"])),
        pdvStatus,
        emrGranted,
        emrAccess:     emrGranted,
        login:         clean(getCell(p, ["Nurse Navigator EMR Access Login"])),
        contact:       clean(getCell(p, ["Direct Office Contact","Contact"])),
        networkAccess: clean(getCell(p, ["Network Management Access"])),
        notes:         clean(getCell(p, ["Notes"])),
      };
    });

    // "Complete" or starts with "Complete" = enrolled
    const isPDVComplete  = (v) => normalize(v).startsWith("complete");
    const isEMRComplete  = (v) => normalize(v).startsWith("complete");
    const isDeclined     = (v) => normalize(v).includes("declined");
    const isTBD          = (v) => normalize(v) === "tbd";

    const enrolled   = practices.filter((p) => isPDVComplete(p.pdvStatus)).length;
    const declined   = practices.filter((p) => isDeclined(p.pdvStatus)).length;
    const tbd        = practices.filter((p) => isTBD(p.pdvStatus)).length;
    const pending    = practices.length - enrolled - declined - tbd;
    const emrComplete = practices.filter((p) => isEMRComplete(p.emrGranted)).length;

    return { total: practices.length, enrolled, pending, declined, tbd, emrComplete, practices };
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
        const workbook = XLSX.read(buffer, { type: "array", cellDates: true }); // cellDates: true = parse dates properly
        const fileName = file.name.toLowerCase();
        const sheet    = workbook.Sheets[workbook.SheetNames[0]];

        // More specific detection to avoid filename conflicts
        const isPracticeFile = fileName.includes("ccpaco") || 
                               (fileName.includes("tracking") && !fileName.includes("patient"));
        const isPatientFile  = fileName.includes("patient") || 
                               (fileName.includes("tracker") && !fileName.includes("ccpaco"));

        if (isPracticeFile && !isPatientFile) {
          // Practice file: row 1=title, row 2=headers, row 3=empty, row 4+=data
          // Use range:1 to skip title and use row 2 as headers directly
          const practiceRows = XLSX.utils.sheet_to_json(sheet, {
            defval: "",
            blankrows: false,
            range: 1,
          });
          const result       = processPracticeRows(practiceRows);
          newData.practiceMetrics = { total: result.total, enrolled: result.enrolled, pending: result.pending, declined: result.declined, tbd: result.tbd, emrComplete: result.emrComplete || 0 };
          newData.practices = result.practices;
        } else {
          // Patient tracker: header is on row 3 (index 2) — rows 1-2 are title/section headers
          const headerRow   = findBestHeaderRow(sheet, XLSX);
          const patientRows = readSheetWithHeaderRow(sheet, XLSX, headerRow, 16);
          const result      = processPatientRows(patientRows);
          newData.metrics   = {
            totalPatients: result.totalPatients, currentlyAdmitted: result.currentlyAdmitted,
            discharged: result.discharged, tcmScheduled: result.tcmScheduled,
            tcmPending: result.tcmPending, visitVerified: result.visitVerified,
            notYetScheduled: result.notYetScheduled, missed14DayWindow: result.missed14Day,
            scheduledRate: result.scheduledRate, nurseCounts: result.nurseCounts,
          };
          newData.patients = result.patients;
        }
      }
      setData(newData);
      setLastSync(new Date());
      setSyncMessage(`✓ Sync complete. ${newData.patients.length} patients loaded.`);
    } catch (error) {
      console.error(error);
      setSyncMessage("Error: " + error.message);
    }
    setTimeout(() => { setSyncing(false); setSyncMessage(""); }, 4000);
    event.target.value = "";
  };

  return (
    <div style={pageStyle}>
      <input ref={fileInputRef} type="file" accept=".xlsx,.xls" multiple onChange={handleFileUpload} style={{ display: "none" }} />
      {modal && <Modal title={modal.title} columns={modal.columns} rows={modal.rows} onClose={closeModal} />}

      <header style={headerStyle}>
        <div style={headerInnerStyle}>
          <h1 style={titleStyle}>Nurse Navigator Program</h1>
          <div style={{ display: "flex", gap: "12px", alignItems: "center", flexWrap: "wrap" }}>
            <button onClick={() => fileInputRef.current.click()} disabled={syncing} style={buttonStyle}>
              <Upload size={18} />{syncing ? "Syncing..." : "Sync from SharePoint"}
            </button>
            <div style={liveStyle}>
              Live Data<br />
              <span style={{ fontWeight: "400" }}>{lastSync ? `Updated ${lastSync.toLocaleTimeString()}` : "No file uploaded"}</span>
            </div>
          </div>
        </div>

        {showLagWarning && (
          <div style={{ maxWidth: "1400px", margin: "12px auto 0", padding: "10px 16px", background: "#fef9c3", border: "2px solid #fbbf24", borderRadius: "10px", color: "#92400e", fontSize: "13px", fontWeight: "600", display: "flex", alignItems: "center", gap: "8px" }}>
            <AlertTriangle size={16} />
            Monday reminder: Weekend admissions (Fri–Sun) may be delayed 2–3 days in Care Compass. Verify directly with facilities.
            <button onClick={() => setShowLagWarning(false)} style={{ marginLeft: "auto", background: "none", border: "none", cursor: "pointer", color: "#92400e" }}>✕</button>
          </div>
        )}

        {syncMessage && (
          <div style={{ maxWidth: "1400px", margin: "12px auto 0", padding: "12px", background: "#dbeafe", borderRadius: "10px", color: "#1e40af", textAlign: "center", fontWeight: "700" }}>
            {syncMessage}
          </div>
        )}

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
    const k = (e) => { if (e.key === "Escape") onClose(); };
    window.addEventListener("keydown", k);
    return () => window.removeEventListener("keydown", k);
  }, [onClose]);

  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.55)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: "24px" }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: "white", borderRadius: "20px", width: "100%", maxWidth: "1000px", maxHeight: "82vh", display: "flex", flexDirection: "column", boxShadow: "0 25px 50px rgba(0,0,0,0.3)", overflow: "hidden" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "20px 28px", background: "linear-gradient(135deg,#667eea,#764ba2)" }}>
          <h2 style={{ margin: 0, color: "white", fontSize: "20px", fontWeight: "800" }}>{title}</h2>
          <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
            <span style={{ color: "rgba(255,255,255,0.8)", fontSize: "13px" }}>{rows.length} records</span>
            <button onClick={onClose} style={{ background: "rgba(255,255,255,0.2)", border: "none", borderRadius: "8px", padding: "6px", cursor: "pointer", color: "white", display: "flex" }}><X size={18} /></button>
          </div>
        </div>
        <div style={{ overflowY: "auto" }}>
          <table style={{ ...tableStyle, margin: 0 }}>
            <thead style={{ position: "sticky", top: 0, background: "#f8fafc" }}>
              <tr>{columns.map((c) => <th key={c} style={thStyle}>{c}</th>)}</tr>
            </thead>
            <tbody>
              {rows.length === 0
                ? <tr><td colSpan={columns.length} style={{ ...tdStyle, textAlign: "center", color: "#94a3b8", padding: "40px" }}>No records.</td></tr>
                : rows.map((row, i) => (
                  <tr key={i} style={{ background: i % 2 === 0 ? "white" : "#f8fafc" }}>
                    {row.map((cell, j) => <td key={j} style={tdStyle}>{typeof cell === "object" && cell !== null ? cell : (cell || "N/A")}</td>)}
                  </tr>
                ))}
            </tbody>
          </table>
        </div>
        <div style={{ padding: "16px 28px", borderTop: "1px solid #e2e8f0", textAlign: "right", background: "#f8fafc" }}>
          <button onClick={onClose} style={{ padding: "10px 24px", background: "linear-gradient(135deg,#667eea,#764ba2)", color: "white", border: "none", borderRadius: "10px", fontWeight: "700", cursor: "pointer" }}>Close</button>
        </div>
      </div>
    </div>
  );
}

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

  const [search,         setSearch]         = useState("");
  const [statusFilter,   setStatusFilter]   = useState("all");
  const [navFilter,      setNavFilter]      = useState("all");
  const [locationFilter, setLocationFilter] = useState("all");
  const [sortCol,        setSortCol]        = useState(null);
  const [sortDir,        setSortDir]        = useState("asc");

  const COLS = ["Patient","Practice","Location","Room","Navigator","Anticipated D/C","Actual Discharge","Days Since D/C","Status","TCM Appt","TCM Date","Appt Type","Visit Verified","14-Day Status","Notes"];

  const statusBadge = (p) => {
    const map = {
      "Visit Verified":      { bg: "#ede9fe", color: "#7c3aed" },
      "Scheduled":           { bg: "#dcfce7", color: "#15803d" },
      "Pending":             { bg: "#dbeafe", color: "#1e40af" },
      "Discharged - No TCM": { bg: "#fef9c3", color: "#b45309" },
      "Admitted":            { bg: "#e0f2fe", color: "#0369a1" },
    };
    const s = map[p.status] || { bg: "#f1f5f9", color: "#64748b" };
    return <Badge text={p.status} bg={s.bg} color={s.color} />;
  };

  const toRows = (list) => list.map((p) => [
    p.name,
    p.practice,
    p.location,
    p.room || "–",
    p.navigator,
    p.anticipatedStr,
    p.actualDischargeStr,
    p.daysSinceDischarge !== null
      ? <span style={{ fontWeight: "700", color: p.daysSinceDischarge > 14 ? "#ef4444" : p.daysSinceDischarge > 10 ? "#f59e0b" : "#10b981" }}>{p.daysSinceDischarge}d</span>
      : "–",
    statusBadge(p),
    <Badge key="tcm" text={p.tcmRaw || "–"} bg={p.tcmScheduled ? "#dcfce7" : p.tcmPending ? "#dbeafe" : "#fee2e2"} color={p.tcmScheduled ? "#15803d" : p.tcmPending ? "#1e40af" : "#991b1b"} />,
    p.tcmDateStr,
    p.apptType || "–",
    <Badge key="vv" text={p.visitVerified ? "✓ Yes" : "–"} bg={p.visitVerified ? "#ede9fe" : "#f1f5f9"} color={p.visitVerified ? "#7c3aed" : "#94a3b8"} />,
    p.windowStatus || "–",
    p.notes ? <span title={p.notes} style={{ maxWidth: "120px", display: "inline-block", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{p.notes}</span> : "–",
  ]);

  const navigators = ["all", ...Array.from(new Set(patients.map((p) => p.navigator).filter(Boolean)))];
  const locations  = ["all", ...Array.from(new Set(patients.map((p) => p.location).filter(Boolean)))];

  const scheduledList  = patients.filter((p) => p.tcmScheduled && !p.visitVerified);
  const verifiedList   = patients.filter((p) => p.visitVerified);
  const pendingList    = patients.filter((p) => p.tcmPending);
  const noTcmList      = patients.filter((p) => !p.tcmScheduled && !p.tcmPending && !p.visitVerified);
  const admittedList   = patients.filter((p) => !p.actualDischarge);
  const dischargedList = patients.filter((p) => p.actualDischarge && !p.visitVerified);

  const filtered = patients
    .filter((p) => {
      const q = search.toLowerCase();
      const matchSearch = !search ||
        p.name.toLowerCase().includes(q) ||
        p.practice.toLowerCase().includes(q) ||
        p.navigator.toLowerCase().includes(q) ||
        p.location.toLowerCase().includes(q);

      const matchStatus =
        statusFilter === "all"           ? true :
        statusFilter === "admitted"      ? !p.actualDischarge :
        statusFilter === "discharged"    ? (p.actualDischarge && !p.visitVerified) :
        statusFilter === "scheduled"     ? (p.tcmScheduled && !p.visitVerified) :
        statusFilter === "pending"       ? p.tcmPending :
        statusFilter === "visit_verified"? p.visitVerified :
        statusFilter === "no_tcm"        ? (!p.tcmScheduled && !p.tcmPending && !p.visitVerified) : true;

      const matchNav = navFilter === "all"      || p.navigator === navFilter;
      const matchLoc = locationFilter === "all" || p.location === locationFilter;
      return matchSearch && matchStatus && matchNav && matchLoc;
    })
    .sort((a, b) => {
      if (!sortCol) return 0;
      const va = (a[sortCol] ?? "").toString().toLowerCase();
      const vb = (b[sortCol] ?? "").toString().toLowerCase();
      return sortDir === "asc" ? va.localeCompare(vb) : vb.localeCompare(va);
    });

  const handleSort = (col) => {
    if (sortCol === col) setSortDir((d) => d === "asc" ? "desc" : "asc");
    else { setSortCol(col); setSortDir("asc"); }
  };

  const pieData = [
    { name: "Admitted",       value: admittedList.length,  color: "#3b82f6" },
    { name: "Discharged",     value: dischargedList.length,color: "#f59e0b" },
    { name: "Scheduled",      value: scheduledList.length, color: "#10b981" },
    { name: "Visit Verified", value: verifiedList.length,  color: "#8b5cf6" },
    { name: "Pending",        value: pendingList.length,   color: "#60a5fa" },
  ].filter((d) => d.value > 0);

  const nurseData = Object.entries(m.nurseCounts || {}).map(([name, count]) => ({ name, count }));

  return (
    <div>
      {/* KPI CARDS */}
      <div style={gridStyle}>
        <MetricCard icon={<Users />}         title="Total Patients"      value={m.totalPatients}     subtitle="Click to view all"          color="#3b82f6"
          onClick={() => openModal(`All Patients (${patients.length})`, COLS, toRows(patients))} />
        <MetricCard icon={<Activity />}      title="Currently Admitted"  value={m.currentlyAdmitted} subtitle="In facility · Click"         color="#0369a1"
          onClick={() => openModal(`Currently Admitted (${admittedList.length})`, COLS, toRows(admittedList))} />
        <MetricCard icon={<Calendar />}      title="TCM Scheduled"       value={m.tcmScheduled}      subtitle={`${m.scheduledRate}% rate · Click`} color="#10b981"
          onClick={() => openModal(`TCM Scheduled (${scheduledList.length})`, COLS, toRows(scheduledList))} />
        <MetricCard icon={<CheckCircle />}   title="Visit Verified"      value={m.visitVerified}     subtitle="NN verified · Click"          color="#8b5cf6"
          onClick={() => openModal(`Visit Verified (${verifiedList.length})`, COLS, toRows(verifiedList))} />
        <MetricCard icon={<Clock />}         title="Pending"             value={m.tcmPending}        subtitle="Pending TCM · Click"          color="#60a5fa"
          onClick={() => openModal(`TCM Pending (${pendingList.length})`, COLS, toRows(pendingList))} />
        <MetricCard icon={<TrendingUp />}    title="No TCM Scheduled"    value={m.notYetScheduled}   subtitle="Needs outreach · Click"       color="#ef4444"
          onClick={() => openModal(`No TCM Scheduled (${noTcmList.length})`, COLS, toRows(noTcmList))} />
      </div>

      {/* CHARTS */}
      <div style={chartGridStyle}>
        <ChartCard title="Patient Status Breakdown">
          <ResponsiveContainer width="100%" height={280}>
            <PieChart>
              <Pie data={pieData} dataKey="value" outerRadius={100} label>
                {pieData.map((e, i) => <Cell key={i} fill={e.color} />)}
              </Pie>
              <Tooltip /><Legend />
            </PieChart>
          </ResponsiveContainer>
        </ChartCard>

        <ChartCard title="Nurse Navigator Workload">
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

      {/* DRILLDOWN TABLE */}
      <div style={tableWrapStyle}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "16px", marginBottom: "16px" }}>
          <h2 style={{ margin: 0 }}>
            Patient Drilldown <span style={{ fontSize: "14px", color: "#64748b", fontWeight: "400" }}>({filtered.length} of {patients.length})</span>
          </h2>
          <div style={{ display: "flex", gap: "8px", flexWrap: "wrap" }}>
            {[
              { label: "Admitted",       list: admittedList,  color: "#0369a1", bg: "#e0f2fe" },
              { label: "Discharged",     list: dischargedList,color: "#b45309", bg: "#fef9c3" },
              { label: "Scheduled",      list: scheduledList, color: "#15803d", bg: "#dcfce7" },
              { label: "Visit Verified", list: verifiedList,  color: "#7c3aed", bg: "#ede9fe" },
              { label: "Pending",        list: pendingList,   color: "#1e40af", bg: "#dbeafe" },
            ].map((b) => (
              <span key={b.label} onClick={() => openModal(`${b.label} (${b.list.length})`, COLS, toRows(b.list))}
                style={{ padding: "4px 14px", borderRadius: "20px", background: b.bg, color: b.color, fontSize: "12px", fontWeight: "700", cursor: "pointer" }}>
                {b.label}: {b.list.length}
              </span>
            ))}
          </div>
        </div>

        <div style={{ display: "flex", gap: "12px", flexWrap: "wrap", marginBottom: "16px" }}>
          <input placeholder="Search name, practice, facility, navigator…" value={search} onChange={(e) => setSearch(e.target.value)} style={{ ...inputStyle, minWidth: "240px", flex: 1 }} />
          <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)} style={inputStyle}>
            <option value="all">All Status</option>
            <option value="admitted">Currently Admitted</option>
            <option value="discharged">Discharged</option>
            <option value="scheduled">TCM Scheduled</option>
            <option value="pending">Pending</option>
            <option value="visit_verified">Visit Verified</option>
            <option value="no_tcm">No TCM Scheduled</option>
          </select>
          <select value={navFilter} onChange={(e) => setNavFilter(e.target.value)} style={inputStyle}>
            {navigators.map((n) => <option key={n} value={n}>{n === "all" ? "All Navigators" : n}</option>)}
          </select>
          <select value={locationFilter} onChange={(e) => setLocationFilter(e.target.value)} style={inputStyle}>
            {locations.map((l) => <option key={l} value={l}>{l === "all" ? "All Locations" : l}</option>)}
          </select>
          {(search || statusFilter !== "all" || navFilter !== "all" || locationFilter !== "all") && (
            <button onClick={() => { setSearch(""); setStatusFilter("all"); setNavFilter("all"); setLocationFilter("all"); }}
              style={{ ...inputStyle, background: "#f1f5f9", cursor: "pointer", border: "2px solid #e2e8f0", color: "#64748b", whiteSpace: "nowrap" }}>✕ Clear</button>
          )}
        </div>

        <div style={{ overflowX: "auto" }}>
          <table style={tableStyle}>
            <thead>
              <tr>
                {[
                  { label: "Patient",        col: "name" },
                  { label: "Practice",       col: "practice" },
                  { label: "Location",       col: "location" },
                  { label: "Navigator",      col: "navigator" },
                  { label: "Anticipated D/C",col: null },
                  { label: "Actual Discharge",col: null },
                  { label: "Days Since D/C", col: null },
                  { label: "Status",         col: null },
                  { label: "TCM Appt",       col: null },
                  { label: "Appt Date",      col: null },
                  { label: "Type",           col: null },
                  { label: "Visit Verified", col: null },
                ].map(({ label, col }) => (
                  <th key={label} onClick={() => col && handleSort(col)} style={{ ...thStyle, cursor: col ? "pointer" : "default" }}>
                    {label}{col ? (sortCol === col ? (sortDir === "asc" ? " ↑" : " ↓") : " ↕") : ""}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 ? (
                <tr><td colSpan={12} style={{ ...tdStyle, textAlign: "center", color: "#94a3b8", padding: "40px" }}>No patients match filters. Upload files or adjust filters.</td></tr>
              ) : (
                filtered.map((p, i) => (
                  <tr key={i} style={{ background: i % 2 === 0 ? "white" : "#f8fafc", cursor: "pointer" }}
                    onClick={() => openModal(`Patient: ${p.name}`, ["Field","Value"], [
                      ["Name",               p.name],
                      ["Practice",           p.practice],
                      ["Location/Facility",  p.location],
                      ["Room",               p.room || "N/A"],
                      ["Navigator",          p.navigator],
                      ["Anticipated D/C",    p.anticipatedStr],
                      ["Actual Discharge",   p.actualDischargeStr],
                      ["Days Since D/C",     p.daysSinceDischarge !== null ? `${p.daysSinceDischarge} days` : "Still Admitted"],
                      ["Status",             p.status],
                      ["TCM Appt Scheduled", p.tcmRaw || "N/A"],
                      ["TCM Appt Date",      p.tcmDateStr],
                      ["Appt Type",          p.apptType || "N/A"],
                      ["2-Day Call Attempt", p.call2 || "–"],
                      ["7-Day Call Attempt", p.call7 || "–"],
                      ["14-Day Window",      p.windowStatus || "–"],
                      ["Visit Verified",     p.visitVerified ? "✓ Yes" : "No"],
                      ["Notes",              p.notes || "–"],
                    ])}
                  >
                    <td style={{ ...tdStyle, fontWeight: "700" }}>{p.name || "N/A"}</td>
                    <td style={tdStyle}>{p.practice || "N/A"}</td>
                    <td style={tdStyle}>{p.location || "N/A"}</td>
                    <td style={tdStyle}>{p.navigator}</td>
                    <td style={tdStyle}>{p.anticipatedStr}</td>
                    <td style={tdStyle}>{p.actualDischargeStr}</td>
                    <td style={tdStyle}>
                      {p.daysSinceDischarge !== null
                        ? <span style={{ fontWeight: "700", color: p.daysSinceDischarge > 14 ? "#ef4444" : p.daysSinceDischarge > 10 ? "#f59e0b" : "#10b981" }}>{p.daysSinceDischarge}d</span>
                        : <span style={{ color: "#94a3b8" }}>–</span>}
                    </td>
                    <td style={tdStyle}>{statusBadge(p)}</td>
                    <td style={tdStyle}>
                      <Badge text={p.tcmRaw || "–"} bg={p.tcmScheduled ? "#dcfce7" : p.tcmPending ? "#dbeafe" : "#fee2e2"} color={p.tcmScheduled ? "#15803d" : p.tcmPending ? "#1e40af" : "#991b1b"} />
                    </td>
                    <td style={tdStyle}>{p.tcmDateStr}</td>
                    <td style={tdStyle}>{p.apptType || "–"}</td>
                    <td style={tdStyle}>
                      <Badge text={p.visitVerified ? "✓ Yes" : "–"} bg={p.visitVerified ? "#ede9fe" : "#f1f5f9"} color={p.visitVerified ? "#7c3aed" : "#94a3b8"} />
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

  const normalize = (v) => (v === null || v === undefined ? "" : String(v).trim().toLowerCase());

  const PRACTICE_COLS = ["Practice","Consultant","City","Facility","PDV Status","EMR Access","Login","Contact","Network Access","Notes"];
  const toRows = (list) => list.map((p) => [p.name, p.consultant, p.location, p.hospitals, p.pdvStatus, p.emrAccess, p.login, p.contact, p.networkAccess, p.notes]);

  const enrolledPractices = practices.filter((p) => normalize(p.pdvStatus) === "complete");
  const declinedPractices = practices.filter((p) => normalize(p.pdvStatus).includes("declined"));
  const tbdPractices      = practices.filter((p) => normalize(p.pdvStatus) === "tbd");
  const pendingPractices  = practices.filter((p) => { const s = normalize(p.pdvStatus); return s !== "complete" && !s.includes("declined") && s !== "tbd"; });

  const statusData = [
    { name: "PDV Complete",  value: m.enrolled,  color: "#10b981" },
    { name: "Pending",       value: m.pending,   color: "#f59e0b" },
    { name: "Declined",      value: m.declined,  color: "#ef4444" },
    { name: "TBD",           value: m.tbd,       color: "#8b5cf6" },
  ].filter((d) => d.value > 0);

  const consultantMap = {};
  practices.forEach((p) => {
    const c = p.consultant || "N/A";
    if (!consultantMap[c]) consultantMap[c] = { name: c, complete: 0, pending: 0 };
    if ((p.pdvStatus || "").toLowerCase() === "complete") consultantMap[c].complete += 1;
    else consultantMap[c].pending += 1;
  });
  const consultantData = Object.values(consultantMap);

  const filtered = practices.filter((p) => {
    const searchMatch = !search || p.name.toLowerCase().includes(search.toLowerCase()) || p.consultant.toLowerCase().includes(search.toLowerCase());
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
      <div style={gridStyle}>
        <MetricCard icon={<Building2 />}   title="Total Practices Enrolled" value={m.total}    subtitle="Click to view all"  color="#3b82f6"
          onClick={() => openModal(`All Practices (${practices.length})`, PRACTICE_COLS, toRows(practices))} />
        <MetricCard icon={<CheckCircle />} title="PDV Forms Completed"   value={m.enrolled}    subtitle={`${m.total ? ((m.enrolled/m.total)*100).toFixed(1) : 0}% · Click to view`} color="#10b981"
          onClick={() => openModal(`PDV Complete (${enrolledPractices.length})`, PRACTICE_COLS, toRows(enrolledPractices))} />
        <MetricCard icon={<Activity />}    title="EMR Access Granted"    value={m.emrComplete} subtitle="Click to view"      color="#8b5cf6"
          onClick={() => openModal(`EMR Access Granted (${practices.filter(p=>normalize(p.emrGranted||"").startsWith("complete")).length})`, PRACTICE_COLS, toRows(practices.filter(p=>normalize(p.emrGranted||"").startsWith("complete"))))} />
        <MetricCard icon={<TrendingUp />}  title="Declined"              value={m.declined}    subtitle="Click to view"      color="#ef4444"
          onClick={() => openModal(`Declined (${declinedPractices.length})`, PRACTICE_COLS, toRows(declinedPractices))} />
        <MetricCard icon={<Clock />}       title="Pending / TBD"         value={m.pending + m.tbd} subtitle="Click to view"  color="#f59e0b"
          onClick={() => openModal(`Pending & TBD (${pendingPractices.length + tbdPractices.length})`, PRACTICE_COLS, toRows([...pendingPractices, ...tbdPractices]))} />
      </div>

      <div style={chartGridStyle}>
        <ChartCard title="Enrollment Status Breakdown">
          {statusData.length === 0 ? (
            <div style={{ padding: "60px", textAlign: "center", color: "#94a3b8" }}>
              <Building2 size={40} style={{ margin: "0 auto 12px", opacity: 0.3 }} />
              <p>Upload practice file to see data</p>
            </div>
          ) : (
            <ResponsiveContainer width="100%" height={300}>
              <PieChart>
                <Pie data={statusData} dataKey="value" outerRadius={110} label>
                  {statusData.map((e, i) => <Cell key={i} fill={e.color} />)}
                </Pie>
                <Tooltip /><Legend />
              </PieChart>
            </ResponsiveContainer>
          )}
        </ChartCard>

        <ChartCard title="Practice Enrolled by Consultant">
          {consultantData.length === 0 ? (
            <div style={{ padding: "60px", textAlign: "center", color: "#94a3b8" }}>
              <p>Upload practice file to see data</p>
            </div>
          ) : (
            <ResponsiveContainer width="100%" height={300}>
              <BarChart data={consultantData} layout="vertical" margin={{ left: 140 }}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis type="number" allowDecimals={false} />
                <YAxis dataKey="name" type="category" width={130} />
                <Tooltip /><Legend />
                <Bar dataKey="complete" fill="#10b981" name="Complete"
                  onClick={(d) => openModal(`${d.name} — Complete`, PRACTICE_COLS, toRows(practices.filter((p) => p.consultant === d.name && (p.pdvStatus||"").toLowerCase() === "complete")))} />
                <Bar dataKey="pending"  fill="#f59e0b" name="Pending"
                  onClick={(d) => openModal(`${d.name} — Pending`,  PRACTICE_COLS, toRows(practices.filter((p) => p.consultant === d.name && (p.pdvStatus||"").toLowerCase() !== "complete")))} />
              </BarChart>
            </ResponsiveContainer>
          )}
        </ChartCard>
      </div>

      <div style={tableWrapStyle}>
        <div style={{ display: "flex", justifyContent: "space-between", gap: "12px", flexWrap: "wrap", marginBottom: "16px" }}>
          <h2 style={{ margin: 0 }}>Practice Drilldown ({filtered.length} of {practices.length})</h2>
          <div style={{ display: "flex", gap: "12px" }}>
            <input placeholder="Search practice or consultant..." value={search} onChange={(e) => setSearch(e.target.value)} style={inputStyle} />
            <select value={filter} onChange={(e) => setFilter(e.target.value)} style={inputStyle}>
              <option value="all">All Status</option>
              <option value="complete">Complete (Enrolled)</option>
              <option value="pending">Pending</option>
              <option value="declined">Declined</option>
              <option value="tbd">TBD</option>
            </select>
          </div>
        </div>

        <div style={{ overflowX: "auto" }}>
          <table style={tableStyle}>
            <thead>
              <tr>{["Practice","Consultant","City","Facility","PDV Status","EMR Access","Contact"].map((c) => <th key={c} style={thStyle}>{c}</th>)}</tr>
            </thead>
            <tbody>
              {filtered.length === 0
                ? <tr><td colSpan={7} style={{ ...tdStyle, textAlign: "center", color: "#94a3b8", padding: "40px" }}>No data. Upload Excel files.</td></tr>
                : filtered.map((p, i) => (
                  <tr key={i} style={{ background: i % 2 === 0 ? "white" : "#f8fafc", cursor: "pointer" }}
                    onClick={() => openModal(`Practice: ${p.name}`, ["Field","Value"], [
                      ["Practice Name",  p.name], ["Consultant", p.consultant], ["City", p.location],
                      ["Facility", p.hospitals], ["PDV Status", p.pdvStatus], ["EMR Access", p.emrAccess],
                      ["Login", p.login], ["Contact", p.contact], ["Network Access", p.networkAccess], ["Notes", p.notes],
                    ])}
                  >
                    <td style={{ ...tdStyle, fontWeight: "700" }}>{p.name || "N/A"}</td>
                    <td style={tdStyle}>{p.consultant || "N/A"}</td>
                    <td style={tdStyle}>{p.location || "N/A"}</td>
                    <td style={{ ...tdStyle, maxWidth: "160px", overflow: "hidden", textOverflow: "ellipsis" }}>{p.hospitals || "N/A"}</td>
                    <td style={tdStyle}>
                      <Badge text={p.pdvStatus || "N/A"}
                        bg={(p.pdvStatus||"").toLowerCase()==="complete" ? "#dcfce7" : (p.pdvStatus||"").toLowerCase().includes("declined") ? "#fee2e2" : (p.pdvStatus||"").toLowerCase()==="tbd" ? "#ede9fe" : "#fef9c3"}
                        color={(p.pdvStatus||"").toLowerCase()==="complete" ? "#15803d" : (p.pdvStatus||"").toLowerCase().includes("declined") ? "#b91c1c" : (p.pdvStatus||"").toLowerCase()==="tbd" ? "#7c3aed" : "#b45309"}
                      />
                    </td>
                    <td style={tdStyle}>{p.emrAccess || "N/A"}</td>
                    <td style={tdStyle}>{p.contact || "N/A"}</td>
                  </tr>
                ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ─── SHARED ───────────────────────────────────────────────────
function MetricCard({ icon, title, value, subtitle, color, onClick }) {
  const [hovered, setHovered] = useState(false);
  return (
    <div onClick={onClick} onMouseEnter={() => setHovered(true)} onMouseLeave={() => setHovered(false)}
      style={{ ...cardStyle, cursor: onClick ? "pointer" : "default", transform: hovered && onClick ? "translateY(-6px)" : "none", boxShadow: hovered && onClick ? `0 12px 24px ${color}33` : "0 4px 8px rgba(0,0,0,0.08)", transition: "all 0.2s ease", border: hovered && onClick ? `2px solid ${color}44` : "2px solid transparent" }}>
      <div style={{ width: "48px", height: "48px", borderRadius: "12px", background: color, color: "white", display: "flex", alignItems: "center", justifyContent: "center", marginBottom: "16px" }}>{icon}</div>
      <div style={{ fontSize: "14px", color: "#475569", fontWeight: "700" }}>{title}</div>
      <div style={{ fontSize: "34px", fontWeight: "800", marginTop: "8px" }}>{value}</div>
      <div style={{ fontSize: "12px", color: hovered && onClick ? color : "#64748b", marginTop: "6px", display: "flex", alignItems: "center", gap: "4px" }}>
        {subtitle}{onClick && hovered && <ChevronRight size={12} />}
      </div>
    </div>
  );
}

function Badge({ text, bg, color }) {
  return <span style={{ display: "inline-block", padding: "4px 12px", borderRadius: "20px", fontSize: "12px", fontWeight: "700", background: bg, color }}>{text}</span>;
}

function ChartCard({ title, children }) {
  return <div style={cardStyle}><h3 style={{ marginTop: 0 }}>{title}</h3>{children}</div>;
}

const pageStyle        = { minHeight: "100vh", background: "linear-gradient(135deg,#667eea 0%,#764ba2 100%)", fontFamily: "Arial, sans-serif" };
const headerStyle      = { background: "rgba(255,255,255,0.96)", padding: "24px 40px", boxShadow: "0 4px 8px rgba(0,0,0,0.08)" };
const headerInnerStyle = { maxWidth: "1400px", margin: "0 auto", display: "flex", justifyContent: "space-between", alignItems: "center", gap: "16px", flexWrap: "wrap" };
const titleStyle       = { fontSize: "30px", fontWeight: "800", margin: 0, color: "#5b5fc7" };
const buttonStyle      = { padding: "12px 24px", background: "linear-gradient(135deg,#667eea,#764ba2)", color: "white", border: "none", borderRadius: "10px", fontWeight: "700", cursor: "pointer", display: "flex", alignItems: "center", gap: "8px" };
const liveStyle        = { padding: "10px 18px", background: "#f0fdf4", border: "2px solid #86efac", borderRadius: "12px", color: "#166534", fontSize: "12px", fontWeight: "700" };
const tabsStyle        = { maxWidth: "1400px", margin: "24px auto 0", display: "flex", gap: "10px" };
const mainStyle        = { maxWidth: "1400px", margin: "0 auto", padding: "40px" };
const gridStyle        = { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: "24px", marginBottom: "32px" };
const chartGridStyle   = { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(420px, 1fr))", gap: "24px", marginBottom: "24px" };
const cardStyle        = { background: "white", borderRadius: "16px", padding: "24px", boxShadow: "0 4px 8px rgba(0,0,0,0.08)" };
const tableWrapStyle   = { background: "white", borderRadius: "16px", padding: "24px", boxShadow: "0 4px 8px rgba(0,0,0,0.08)", marginTop: "24px" };
const tableStyle       = { width: "100%", borderCollapse: "collapse", marginTop: "16px" };
const thStyle          = { padding: "12px", textAlign: "left", fontSize: "12px", fontWeight: "800", color: "#475569", textTransform: "uppercase", borderBottom: "2px solid #e2e8f0" };
const tdStyle          = { padding: "14px 12px", fontSize: "14px", color: "#1e293b", borderBottom: "1px solid #f1f5f9" };
const inputStyle       = { padding: "10px 14px", border: "2px solid #e2e8f0", borderRadius: "8px", fontSize: "14px" };

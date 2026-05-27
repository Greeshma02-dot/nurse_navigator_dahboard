import React, { useState, useRef } from "react";
import {
  Upload,
  Users,
  CheckCircle,
  Clock,
  Building2,
  TrendingUp,
  Calendar,
} from "lucide-react";
import {
  BarChart,
  Bar,
  PieChart,
  Pie,
  Cell,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer,
} from "recharts";

const EMPTY_DATA = {
  metrics: {
    totalPatients: 0,
    tcmScheduled: 0,
    notYetScheduled: 0,
    pending: 0,
    scheduledRate: 0,
    missed14DayWindow: 0,
    completedWithinWindow: 0,
    nurseCounts: {},
  },
  patients: [],
  practiceMetrics: {
    total: 0,
    enrolled: 0,
    pending: 0,
    declined: 0,
    tbd: 0,
  },
  practices: [],
};

export default function NurseNavigatorDashboard() {
  const [activeTab, setActiveTab] = useState("patients");
  const [data, setData] = useState(EMPTY_DATA);
  const [lastSync, setLastSync] = useState(null);
  const [syncing, setSyncing] = useState(false);
  const [syncMessage, setSyncMessage] = useState("");
  const fileInputRef = useRef(null);

  const clean = (v) =>
    v === null || v === undefined ? "" : String(v).trim();

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
    const raw = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      blankrows: false,
    });

    const headers = (raw[headerRowIndex] || [])
      .slice(0, maxCols)
      .map((h) => clean(h));

    return raw
      .slice(headerRowIndex + 1)
      .map((row) => {
        const obj = {};
        headers.forEach((h, i) => {
          if (h) obj[h] = row[i] ?? "";
        });
        return obj;
      })
      .filter((row) => Object.values(row).some((v) => clean(v) !== ""));
  };

  const findBestHeaderRow = (sheet, XLSX) => {
    const raw = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      blankrows: false,
    });

    let bestIndex = 0;
    let bestScore = -1;

    raw.slice(0, 10).forEach((row, index) => {
      const joined = row.map((c) => normalize(c)).join(" ");
      let score = 0;

      if (joined.includes("patient")) score += 3;
      if (joined.includes("nurse")) score += 3;
      if (joined.includes("navigator")) score += 3;
      if (joined.includes("tcm")) score += 3;
      if (joined.includes("practice")) score += 2;
      if (joined.includes("completed")) score += 2;

      const nonEmpty = row.filter((c) => clean(c) !== "").length;
      score += nonEmpty;

      if (score > bestScore) {
        bestScore = score;
        bestIndex = index;
      }
    });

    return bestIndex;
  };

  const isYes = (val) => {
    const s = normalize(val);
    return (
      s === "yes" ||
      s === "y" ||
      s === "true" ||
      s === "1" ||
      s === "scheduled" ||
      s === "complete" ||
      s === "completed" ||
      s === "verified" ||
      s === "done" ||
      s === "x" ||
      s === "✓"
    );
  };

  const isNoOrBlank = (val) => {
    const s = normalize(val);
    return s === "" || s === "no" || s === "n" || s === "not yet scheduled";
  };

  const processPatientRows = (patientRows) => {
    const patients = patientRows.map((p) => {
      const navigator = clean(
        getCell(p, ["Nurse Navigator", "Navigator", "Assigned Nurse", "Nurse"])
      );

      const tcmStatus = clean(
        getCell(p, [
          "TCM Appt Scheduled",
          "TCM Appointment Scheduled",
          "TCM Scheduled",
          "TCM Appt",
          "TCM",
        ])
      );

      const completedStatus = clean(
        getCell(p, [
          "Completed Within Window",
          "Completed w/in Window",
          "Visit Verified",
          "Verified",
          "Completed",
        ])
      );

      const missedStatus = clean(
        getCell(p, [
          "Missed 14-Day Window",
          "Missed 14 Day Window",
          "Missed Window",
        ])
      );

      return {
        name: clean(getCell(p, ["Patient Name", "Name", "Patient"])),
        practice: clean(getCell(p, ["Practice", "Practice Name"])),
        location: clean(getCell(p, ["Location", "Site", "Market"])),
        navigator: navigator || "N/A",
        tcmScheduled: isYes(tcmStatus),
        completedWithinWindow: isYes(completedStatus),
        missed14DayWindow: isYes(missedStatus),
        rawTcmStatus: tcmStatus,
      };
    });

    const totalPatients = patients.length;
    const tcmScheduled = patients.filter((p) => p.tcmScheduled).length;
    const notYetScheduled = totalPatients - tcmScheduled;
    const missed14DayWindow = patients.filter((p) => p.missed14DayWindow).length;
    const completedWithinWindow = patients.filter(
      (p) => p.completedWithinWindow
    ).length;

    const nurseCounts = {};
    patients.forEach((p) => {
      nurseCounts[p.navigator] = (nurseCounts[p.navigator] || 0) + 1;
    });

    return {
      totalPatients,
      tcmScheduled,
      notYetScheduled,
      pending: notYetScheduled,
      scheduledRate:
        totalPatients > 0
          ? Number(((tcmScheduled / totalPatients) * 100).toFixed(1))
          : 0,
      missed14DayWindow,
      completedWithinWindow,
      nurseCounts,
      patients,
    };
  };

  const processPracticeRows = (practiceRows) => {
    const practices = practiceRows.map((p) => {
      const pdvStatus = clean(
        getCell(p, ["PDV Forms Completed", "PDV Status", "PDV"])
      );

      return {
        name: clean(getCell(p, ["Practice Participants", "Practice"])),
        consultant: clean(getCell(p, ["Consultant"])),
        location: clean(getCell(p, ["City", "Location"])),
        hospitals: clean(getCell(p, ["Facility Participants", "Hospitals"])),
        pdvStatus,
        emrAccess: clean(
          getCell(p, ["Nurse Navigator EMR Access granted", "EMR Access"])
        ),
        login: clean(getCell(p, ["Nurse Navigator EMR Access Login"])),
        contact: clean(getCell(p, ["Direct Office Contact", "Contact"])),
        networkAccess: clean(getCell(p, ["Network Management Access"])),
        notes: clean(getCell(p, ["Notes"])),
      };
    });

    const enrolled = practices.filter(
      (p) => normalize(p.pdvStatus) === "complete"
    ).length;

    const declined = practices.filter((p) =>
      normalize(p.pdvStatus).includes("declined")
    ).length;

    const tbd = practices.filter((p) => normalize(p.pdvStatus) === "tbd").length;
    const pending = practices.length - enrolled - declined - tbd;

    return {
      total: practices.length,
      enrolled,
      pending,
      declined,
      tbd,
      practices,
    };
  };

  const handleFileUpload = async (event) => {
    const files = event.target.files;
    if (!files || files.length === 0) return;

    setSyncing(true);
    setSyncMessage("Reading uploaded Excel file...");

    try {
      const XLSX = await import("xlsx");
      let newData = JSON.parse(JSON.stringify(EMPTY_DATA));

      for (const file of files) {
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: "array" });
        const fileName = file.name.toLowerCase();
        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        const isPracticeFile =
          fileName.includes("ccpaco") ||
          fileName.includes("practice") ||
          fileName.includes("tracking");

        const isPatientFile =
          fileName.includes("patient") ||
          fileName.includes("patients") ||
          fileName.includes("enrollment");

        if (isPracticeFile && !isPatientFile) {
          const practiceRows = readSheetWithHeaderRow(sheet, XLSX, 2, 10);
          const result = processPracticeRows(practiceRows);

          newData.practiceMetrics = {
            total: result.total,
            enrolled: result.enrolled,
            pending: result.pending,
            declined: result.declined,
            tbd: result.tbd,
          };

          newData.practices = result.practices;
        } else {
          const headerRow = findBestHeaderRow(sheet, XLSX);
          const patientRows = readSheetWithHeaderRow(sheet, XLSX, headerRow, 25);
          const result = processPatientRows(patientRows);

          newData.metrics = {
            totalPatients: result.totalPatients,
            tcmScheduled: result.tcmScheduled,
            notYetScheduled: result.notYetScheduled,
            pending: result.pending,
            scheduledRate: result.scheduledRate,
            missed14DayWindow: result.missed14DayWindow,
            completedWithinWindow: result.completedWithinWindow,
            nurseCounts: result.nurseCounts,
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

    setTimeout(() => {
      setSyncing(false);
      setSyncMessage("");
    }, 3000);

    event.target.value = "";
  };

  return (
    <div style={pageStyle}>
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xls"
        multiple
        onChange={handleFileUpload}
        style={{ display: "none" }}
      />

      <header style={headerStyle}>
        <div style={headerInnerStyle}>
          <h1 style={titleStyle}>Nurse Navigator Program</h1>

          <div style={{ display: "flex", gap: "12px", alignItems: "center" }}>
            <button
              onClick={() => fileInputRef.current.click()}
              disabled={syncing}
              style={buttonStyle}
            >
              <Upload size={18} />
              {syncing ? "Syncing..." : "Sync from SharePoint"}
            </button>

            <div style={liveStyle}>
              Live Data
              <br />
              <span style={{ fontWeight: "400" }}>
                {lastSync
                  ? `Updated ${lastSync.toLocaleTimeString()}`
                  : "No file uploaded"}
              </span>
            </div>
          </div>
        </div>

        {syncMessage && <div style={messageStyle}>{syncMessage}</div>}

        <div style={tabsStyle}>
          <TabButton
            active={activeTab === "patients"}
            onClick={() => setActiveTab("patients")}
            icon={<Users size={18} />}
            label="Patient Tracking"
          />
          <TabButton
            active={activeTab === "practices"}
            onClick={() => setActiveTab("practices")}
            icon={<Building2 size={18} />}
            label="Practice Enrollment"
          />
        </div>
      </header>

      <main style={mainStyle}>
        {activeTab === "patients" ? (
          <PatientTrackingPage data={data} />
        ) : (
          <PracticeEnrollmentPage data={data} />
        )}
      </main>
    </div>
  );
}

function TabButton({ active, onClick, icon, label }) {
  return (
    <button
      onClick={onClick}
      style={{
        padding: "12px 24px",
        background: active ? "linear-gradient(135deg, #667eea, #764ba2)" : "white",
        color: active ? "white" : "#64748b",
        border: active ? "none" : "2px solid #e2e8f0",
        borderRadius: "10px",
        fontWeight: "700",
        cursor: "pointer",
        display: "flex",
        alignItems: "center",
        gap: "8px",
      }}
    >
      {icon}
      {label}
    </button>
  );
}

function PatientTrackingPage({ data }) {
  const m = data.metrics;
  const patients = data.patients;

  const pieData = [
    { name: "Scheduled", value: m.tcmScheduled, color: "#10b981" },
    { name: "Not Yet Scheduled", value: m.notYetScheduled, color: "#f59e0b" },
  ];

  const nurseData = Object.entries(m.nurseCounts || {}).map(([name, count]) => ({
    name,
    count,
  }));

  return (
    <div>
      <div style={gridStyle}>
        <MetricCard
          icon={<Users />}
          title="Total Patients"
          value={m.totalPatients}
          subtitle="From uploaded file"
          color="#3b82f6"
        />
        <MetricCard
          icon={<Calendar />}
          title="TCM Appt Scheduled"
          value={m.tcmScheduled}
          subtitle={`${m.scheduledRate}% scheduled`}
          color="#10b981"
        />
        <MetricCard
          icon={<Clock />}
          title="Not Yet Scheduled"
          value={m.notYetScheduled}
          subtitle="Awaiting scheduling"
          color="#f59e0b"
        />
        <MetricCard
          icon={<TrendingUp />}
          title="Missed 14-Day Window"
          value={m.missed14DayWindow}
          subtitle="Missed window"
          color="#ef4444"
        />
        <MetricCard
          icon={<CheckCircle />}
          title="Completed Within Window"
          value={m.completedWithinWindow}
          subtitle="Completed on time"
          color="#8b5cf6"
        />
      </div>

      <div style={chartGridStyle}>
        <ChartCard title="TCM Scheduling Status">
          <ResponsiveContainer width="100%" height={280}>
            <PieChart>
              <Pie data={pieData} dataKey="value" outerRadius={100} label>
                {pieData.map((e, i) => (
                  <Cell key={i} fill={e.color} />
                ))}
              </Pie>
              <Tooltip />
              <Legend />
            </PieChart>
          </ResponsiveContainer>
        </ChartCard>

        <ChartCard title="Nurse Counts">
          <ResponsiveContainer width="100%" height={280}>
            <BarChart data={nurseData}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="name" />
              <YAxis allowDecimals={false} />
              <Tooltip />
              <Legend />
              <Bar dataKey="count" fill="#667eea" name="Patients" />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>
      </div>

      <DataTable
        title={`Patient Drilldown (${patients.length})`}
        columns={[
          "Patient Name",
          "Practice",
          "Location",
          "Navigator",
          "TCM Scheduled",
          "Completed",
          "Missed 14-Day Window",
        ]}
        rows={patients.map((p) => [
          p.name,
          p.practice,
          p.location,
          p.navigator,
          p.tcmScheduled ? "Scheduled" : "Not Yet Scheduled",
          p.completedWithinWindow ? "Yes" : "No",
          p.missed14DayWindow ? "Yes" : "No",
        ])}
      />
    </div>
  );
}

function PracticeEnrollmentPage({ data }) {
  const m = data.practiceMetrics;
  const practices = data.practices;
  const [filter, setFilter] = useState("all");
  const [search, setSearch] = useState("");

  const statusData = [
    { name: "Complete", value: m.enrolled, color: "#10b981" },
    { name: "Pending", value: m.pending, color: "#f59e0b" },
    { name: "Declined", value: m.declined, color: "#ef4444" },
    { name: "TBD", value: m.tbd, color: "#8b5cf6" },
  ];

  const consultantMap = {};
  practices.forEach((p) => {
    const c = p.consultant || "N/A";
    if (!consultantMap[c]) {
      consultantMap[c] = { name: c, complete: 0, pending: 0 };
    }

    const s = (p.pdvStatus || "").toLowerCase();
    if (s === "complete") consultantMap[c].complete += 1;
    else consultantMap[c].pending += 1;
  });

  const consultantData = Object.values(consultantMap);

  const filtered = practices.filter((p) => {
    const searchMatch =
      !search ||
      p.name.toLowerCase().includes(search.toLowerCase()) ||
      p.consultant.toLowerCase().includes(search.toLowerCase());

    if (!searchMatch) return false;

    const status = (p.pdvStatus || "").toLowerCase();

    if (filter === "complete") return status === "complete";
    if (filter === "declined") return status.includes("declined");
    if (filter === "tbd") return status === "tbd";
    if (filter === "pending") {
      return status !== "complete" && !status.includes("declined") && status !== "tbd";
    }

    return true;
  });

  return (
    <div>
      <div style={gridStyle}>
        <MetricCard
          icon={<Building2 />}
          title="Total Practices"
          value={m.total}
          subtitle="From uploaded file"
          color="#3b82f6"
        />
        <MetricCard
          icon={<CheckCircle />}
          title="Enrolled"
          value={m.enrolled}
          subtitle={`${m.total ? ((m.enrolled / m.total) * 100).toFixed(1) : 0}% success rate`}
          color="#10b981"
        />
        <MetricCard
          icon={<Clock />}
          title="Pending"
          value={m.pending}
          subtitle="In process / blank"
          color="#f59e0b"
        />
        <MetricCard
          icon={<TrendingUp />}
          title="Declined"
          value={m.declined}
          subtitle="Not participating"
          color="#ef4444"
        />
        <MetricCard
          icon={<Clock />}
          title="TBD"
          value={m.tbd}
          subtitle="To be determined"
          color="#8b5cf6"
        />
      </div>

      <div style={chartGridStyle}>
        <ChartCard title="Enrollment Status Breakdown">
          <ResponsiveContainer width="100%" height={300}>
            <PieChart>
              <Pie data={statusData} dataKey="value" outerRadius={110} label>
                {statusData.map((e, i) => (
                  <Cell key={i} fill={e.color} />
                ))}
              </Pie>
              <Tooltip />
              <Legend />
            </PieChart>
          </ResponsiveContainer>
        </ChartCard>

        <ChartCard title="Consultant Performance">
          <ResponsiveContainer width="100%" height={300}>
            <BarChart data={consultantData} layout="vertical" margin={{ left: 140 }}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis type="number" allowDecimals={false} />
              <YAxis dataKey="name" type="category" width={130} />
              <Tooltip />
              <Legend />
              <Bar dataKey="complete" fill="#10b981" name="Complete" />
              <Bar dataKey="pending" fill="#f59e0b" name="Pending" />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>
      </div>

      <div style={tableWrapStyle}>
        <div style={tableHeaderStyle}>
          <h2>
            Practice Drilldown ({filtered.length} of {practices.length})
          </h2>

          <div style={{ display: "flex", gap: "12px" }}>
            <input
              placeholder="Search practice or consultant..."
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              style={inputStyle}
            />

            <select
              value={filter}
              onChange={(e) => setFilter(e.target.value)}
              style={inputStyle}
            >
              <option value="all">All Status</option>
              <option value="complete">Complete</option>
              <option value="pending">Pending</option>
              <option value="declined">Declined</option>
              <option value="tbd">TBD</option>
            </select>
          </div>
        </div>

        <Table
          columns={[
            "Practice",
            "Consultant",
            "City",
            "Facility",
            "PDV Forms Completed",
            "EMR Access",
            "Contact",
          ]}
          rows={filtered.map((p) => [
            p.name,
            p.consultant,
            p.location,
            p.hospitals,
            p.pdvStatus,
            p.emrAccess,
            p.contact,
          ])}
        />
      </div>
    </div>
  );
}

function MetricCard({ icon, title, value, subtitle, color }) {
  return (
    <div style={cardStyle}>
      <div
        style={{
          width: "48px",
          height: "48px",
          borderRadius: "12px",
          background: color,
          color: "white",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          marginBottom: "16px",
        }}
      >
        {icon}
      </div>

      <div style={{ fontSize: "14px", color: "#475569", fontWeight: "700" }}>
        {title}
      </div>

      <div style={{ fontSize: "34px", fontWeight: "800", marginTop: "8px" }}>
        {value}
      </div>

      <div style={{ fontSize: "12px", color: "#64748b", marginTop: "6px" }}>
        {subtitle}
      </div>
    </div>
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

function DataTable({ title, columns, rows }) {
  return (
    <div style={tableWrapStyle}>
      <h2>{title}</h2>
      <Table columns={columns} rows={rows} />
    </div>
  );
}

function Table({ columns, rows }) {
  return (
    <div style={{ overflowX: "auto" }}>
      <table style={tableStyle}>
        <thead>
          <tr>
            {columns.map((c) => (
              <th key={c} style={thStyle}>
                {c}
              </th>
            ))}
          </tr>
        </thead>

        <tbody>
          {rows.length === 0 ? (
            <tr>
              <td colSpan={columns.length} style={tdStyle}>
                No uploaded data yet.
              </td>
            </tr>
          ) : (
            rows.map((row, i) => (
              <tr key={i}>
                {row.map((cell, j) => (
                  <td key={j} style={tdStyle}>
                    {cell || "N/A"}
                  </td>
                ))}
              </tr>
            ))
          )}
        </tbody>
      </table>
    </div>
  );
}

const pageStyle = {
  minHeight: "100vh",
  background: "linear-gradient(135deg, #667eea 0%, #764ba2 100%)",
  fontFamily: "Arial, sans-serif",
};

const headerStyle = {
  background: "rgba(255,255,255,0.96)",
  padding: "24px 40px",
  boxShadow: "0 4px 8px rgba(0,0,0,0.08)",
};

const headerInnerStyle = {
  maxWidth: "1400px",
  margin: "0 auto",
  display: "flex",
  justifyContent: "space-between",
  alignItems: "center",
  gap: "16px",
  flexWrap: "wrap",
};

const titleStyle = {
  fontSize: "30px",
  fontWeight: "800",
  margin: 0,
  color: "#5b5fc7",
};

const buttonStyle = {
  padding: "12px 24px",
  background: "linear-gradient(135deg, #667eea, #764ba2)",
  color: "white",
  border: "none",
  borderRadius: "10px",
  fontWeight: "700",
  cursor: "pointer",
  display: "flex",
  alignItems: "center",
  gap: "8px",
};

const liveStyle = {
  padding: "10px 18px",
  background: "#f0fdf4",
  border: "2px solid #86efac",
  borderRadius: "12px",
  color: "#166534",
  fontSize: "12px",
  fontWeight: "700",
};

const messageStyle = {
  maxWidth: "1400px",
  margin: "16px auto 0",
  padding: "12px",
  background: "#dbeafe",
  borderRadius: "10px",
  color: "#1e40af",
  textAlign: "center",
  fontWeight: "700",
};

const tabsStyle = {
  maxWidth: "1400px",
  margin: "24px auto 0",
  display: "flex",
  gap: "10px",
};

const mainStyle = {
  maxWidth: "1400px",
  margin: "0 auto",
  padding: "40px",
};

const gridStyle = {
  display: "grid",
  gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))",
  gap: "24px",
  marginBottom: "32px",
};

const chartGridStyle = {
  display: "grid",
  gridTemplateColumns: "repeat(auto-fit, minmax(420px, 1fr))",
  gap: "24px",
  marginBottom: "24px",
};

const cardStyle = {
  background: "white",
  borderRadius: "16px",
  padding: "24px",
  boxShadow: "0 4px 8px rgba(0,0,0,0.08)",
};

const tableWrapStyle = {
  background: "white",
  borderRadius: "16px",
  padding: "24px",
  boxShadow: "0 4px 8px rgba(0,0,0,0.08)",
  marginTop: "24px",
};

const tableHeaderStyle = {
  display: "flex",
  justifyContent: "space-between",
  gap: "12px",
  flexWrap: "wrap",
};

const tableStyle = {
  width: "100%",
  borderCollapse: "collapse",
  marginTop: "16px",
};

const thStyle = {
  padding: "12px",
  textAlign: "left",
  fontSize: "12px",
  fontWeight: "800",
  color: "#475569",
  textTransform: "uppercase",
  borderBottom: "2px solid #e2e8f0",
};

const tdStyle = {
  padding: "14px 12px",
  fontSize: "14px",
  color: "#1e293b",
  borderBottom: "1px solid #f1f5f9",
};

const inputStyle = {
  padding: "10px 14px",
  border: "2px solid #e2e8f0",
  borderRadius: "8px",
  fontSize: "14px",
};

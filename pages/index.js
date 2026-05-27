// ============================================================
// REPLACE the entire PatientTrackingPage function with this:
// ============================================================

function PatientTrackingPage({ data }) {
  const m = data.metrics;
  const patients = data.patients;

  // ---- interactive drilldown state ----
  const [search, setSearch] = useState("");
  const [tcmFilter, setTcmFilter] = useState("all");
  const [navFilter, setNavFilter] = useState("all");
  const [sortCol, setSortCol] = useState(null);
  const [sortDir, setSortDir] = useState("asc");

  const pieData = [
    { name: "Scheduled", value: m.tcmScheduled, color: "#10b981" },
    { name: "Not Yet Scheduled", value: m.notYetScheduled, color: "#f59e0b" },
  ];

  const nurseData = Object.entries(m.nurseCounts || {}).map(([name, count]) => ({
    name,
    count,
  }));

  // unique navigator names for filter dropdown
  const navigators = ["all", ...Array.from(new Set(patients.map((p) => p.navigator).filter(Boolean)))];

  // filter + sort
  const filtered = patients
    .filter((p) => {
      const matchSearch =
        !search ||
        p.name.toLowerCase().includes(search.toLowerCase()) ||
        p.practice.toLowerCase().includes(search.toLowerCase()) ||
        p.navigator.toLowerCase().includes(search.toLowerCase());

      const matchTcm =
        tcmFilter === "all" ||
        (tcmFilter === "scheduled" && p.tcmScheduled) ||
        (tcmFilter === "not_scheduled" && !p.tcmScheduled);

      const matchNav = navFilter === "all" || p.navigator === navFilter;

      return matchSearch && matchTcm && matchNav;
    })
    .sort((a, b) => {
      if (!sortCol) return 0;
      const valA = (a[sortCol] ?? "").toString().toLowerCase();
      const valB = (b[sortCol] ?? "").toString().toLowerCase();
      return sortDir === "asc" ? valA.localeCompare(valB) : valB.localeCompare(valA);
    });

  const handleSort = (col) => {
    if (sortCol === col) {
      setSortDir((d) => (d === "asc" ? "desc" : "asc"));
    } else {
      setSortCol(col);
      setSortDir("asc");
    }
  };

  const sortIcon = (col) => {
    if (sortCol !== col) return " ↕";
    return sortDir === "asc" ? " ↑" : " ↓";
  };

  return (
    <div>
      {/* ---- KPI cards ---- */}
      <div style={gridStyle}>
        <MetricCard icon={<Users />}        title="Total Patients"          value={m.totalPatients}          subtitle="From uploaded file"        color="#3b82f6" />
        <MetricCard icon={<Calendar />}     title="TCM Appt Scheduled"      value={m.tcmScheduled}           subtitle={`${m.scheduledRate}% scheduled`} color="#10b981" />
        <MetricCard icon={<Clock />}        title="Not Yet Scheduled"       value={m.notYetScheduled}        subtitle="Awaiting scheduling"       color="#f59e0b" />
        <MetricCard icon={<TrendingUp />}   title="Missed 14-Day Window"    value={m.missed14DayWindow}      subtitle="Missed window"             color="#ef4444" />
        <MetricCard icon={<CheckCircle />}  title="Completed Within Window" value={m.completedWithinWindow}  subtitle="Completed on time"         color="#8b5cf6" />
      </div>

      {/* ---- charts ---- */}
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

      {/* ---- interactive patient drilldown ---- */}
      <div style={tableWrapStyle}>
        {/* header row */}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "16px", marginBottom: "16px" }}>
          <h2 style={{ margin: 0 }}>
            Patient Drilldown{" "}
            <span style={{ fontSize: "14px", color: "#64748b", fontWeight: "400" }}>
              ({filtered.length} of {patients.length})
            </span>
          </h2>

          {/* count badges */}
          <div style={{ display: "flex", gap: "8px", flexWrap: "wrap" }}>
            {[
              { label: "Scheduled",    count: patients.filter((p) => p.tcmScheduled).length,  color: "#10b981", bg: "#dcfce7" },
              { label: "Not Scheduled", count: patients.filter((p) => !p.tcmScheduled).length, color: "#b45309", bg: "#fef9c3" },
              { label: "Completed",    count: patients.filter((p) => p.completedWithinWindow).length, color: "#7c3aed", bg: "#ede9fe" },
              { label: "Missed Window", count: patients.filter((p) => p.missed14DayWindow).length, color: "#b91c1c", bg: "#fee2e2" },
            ].map((b) => (
              <span key={b.label} style={{ padding: "4px 12px", borderRadius: "20px", background: b.bg, color: b.color, fontSize: "12px", fontWeight: "700" }}>
                {b.label}: {b.count}
              </span>
            ))}
          </div>
        </div>

        {/* filter controls */}
        <div style={{ display: "flex", gap: "12px", flexWrap: "wrap", marginBottom: "16px" }}>
          <input
            placeholder="Search name, practice, navigator..."
            value={search}
            onChange={(e) => setSearch(e.target.value)}
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
              ✕ Clear filters
            </button>
          )}
        </div>

        {/* table */}
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
                    style={{ ...thStyle, cursor: col ? "pointer" : "default", userSelect: "none" }}
                    onClick={() => col && handleSort(col)}
                  >
                    {label}{col ? sortIcon(col) : ""}
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
                  <tr key={i} style={{ background: i % 2 === 0 ? "white" : "#f8fafc" }}>
                    <td style={{ ...tdStyle, fontWeight: "700" }}>{p.name || "N/A"}</td>
                    <td style={tdStyle}>{p.practice || "N/A"}</td>
                    <td style={tdStyle}>{p.location || "N/A"}</td>
                    <td style={tdStyle}>{p.navigator || "N/A"}</td>

                    {/* TCM status badge */}
                    <td style={tdStyle}>
                      <span style={{
                        display: "inline-block",
                        padding: "4px 12px",
                        borderRadius: "20px",
                        fontSize: "12px",
                        fontWeight: "700",
                        background: p.tcmScheduled ? "#dcfce7" : "#fef9c3",
                        color: p.tcmScheduled ? "#15803d" : "#b45309",
                      }}>
                        {p.tcmScheduled ? "✓ Scheduled" : "⏳ Not Yet"}
                      </span>
                    </td>

                    {/* Completed badge */}
                    <td style={tdStyle}>
                      <span style={{
                        display: "inline-block",
                        padding: "4px 12px",
                        borderRadius: "20px",
                        fontSize: "12px",
                        fontWeight: "700",
                        background: p.completedWithinWindow ? "#ede9fe" : "#f1f5f9",
                        color: p.completedWithinWindow ? "#7c3aed" : "#94a3b8",
                      }}>
                        {p.completedWithinWindow ? "✓ Yes" : "–"}
                      </span>
                    </td>

                    {/* Missed window badge */}
                    <td style={tdStyle}>
                      {p.missed14DayWindow ? (
                        <span style={{ display: "inline-block", padding: "4px 12px", borderRadius: "20px", fontSize: "12px", fontWeight: "700", background: "#fee2e2", color: "#b91c1c" }}>
                          ⚠ Missed
                        </span>
                      ) : (
                        <span style={{ color: "#94a3b8", fontSize: "12px" }}>–</span>
                      )}
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

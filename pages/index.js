import React, { useState, useRef, useEffect } from "react";
 
export default function NurseNavigatorDashboard() {
 
  const [lastSync, setLastSync] = useState(null);

  const [syncing, setSyncing] = useState(false);

  const [syncMessage, setSyncMessage] = useState("");

  const [data, setData] = useState({});

  const fileInputRef = useRef(null);
 
  // ===============================

  // LOAD LIVE DATA FROM API

  // ===============================

  const loadLiveData = async () => {

    try {

      setSyncing(true);

      setSyncMessage("Refreshing dashboard...");
 
      const res = await fetch("/api/nurse-navigator-data");

      const liveData = await res.json();
 
      setData(liveData);
 
      setLastSync(new Date());
 
      setSyncMessage("✓ Dashboard refreshed successfully");

    } catch (error) {

      console.error("Refresh failed:", error);

      setSyncMessage("Refresh failed");

    }
 
    setTimeout(() => {

      setSyncing(false);

      setSyncMessage("");

    }, 2000);

  };
 
  // ===============================

  // AUTO REFRESH EVERY 5 MINUTES

  // ===============================

  useEffect(() => {
 
    // first load

    loadLiveData();
 
    // auto refresh

    const interval = setInterval(() => {

      loadLiveData();

    }, 300000);
 
    return () => clearInterval(interval);
 
  }, []);
 
  return (
<div>
 
      {/* FILE INPUT */}
<input

        ref={fileInputRef}

        type="file"

        accept=".xlsx,.xls"

        style={{ display: "none" }}

      />
 
      {/* HEADER */}
<div

        style={{

          display: "flex",

          gap: "12px",

          alignItems: "center",

          padding: "20px",

        }}
>
 
        {/* EXISTING BUTTON */}
<button

          onClick={() => fileInputRef.current.click()}

          disabled={syncing}

          style={{

            padding: "12px 24px",

            background: "#667eea",

            color: "white",

            border: "none",

            borderRadius: "10px",

            cursor: "pointer",

            fontWeight: "700",

          }}
>

          {syncing ? "Syncing..." : "Sync from SharePoint"}
</button>
 
        {/* NEW REFRESH BUTTON */}
<button

          onClick={loadLiveData}

          disabled={syncing}

          style={{

            padding: "12px 24px",

            background: "#10b981",

            color: "white",

            border: "none",

            borderRadius: "10px",

            cursor: "pointer",

            fontWeight: "700",

          }}
>

          Refresh Dashboard
</button>
 
        {/* LIVE STATUS */}
<div

          style={{

            padding: "10px 18px",

            background: "#f0fdf4",

            border: "2px solid #86efac",

            borderRadius: "12px",

            color: "#166534",

            fontSize: "12px",

            fontWeight: "700",

          }}
>

          Live Data
<br />
 
          <span style={{ fontWeight: "400" }}>

            {lastSync

              ? `Updated ${lastSync.toLocaleTimeString()}`

              : "No data loaded"}
</span>
</div>
 
      </div>
 
      {/* MESSAGE */}

      {syncMessage && (
<div

          style={{

            margin: "20px",

            padding: "12px",

            background: "#dbeafe",

            borderRadius: "10px",

          }}
>

          {syncMessage}
</div>

      )}
 
      {/* DASHBOARD CONTENT */}
<div style={{ padding: "20px" }}>
 
        <h2>Dashboard Data</h2>
 
        <pre

          style={{

            background: "#f8fafc",

            padding: "20px",

            borderRadius: "10px",

            overflow: "auto",

          }}
>

          {JSON.stringify(data, null, 2)}
</pre>
 
      </div>
 
    </div>

  );

}
 

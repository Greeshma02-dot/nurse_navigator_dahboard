// pages/api/nurse-navigator-data.js
 
let storedData = global.storedData || {
  patients: [],
  practices: [],
  metrics: {
    totalPatients: 0,
    tcmScheduled: 0,
    visitVerified: 0,
    tcmSchedulingRate: 0,
    verificationRate: 0,
    devanPatients: 0,
    sunniePatients: 0,
    tvhPatients: 0,
    ivmPatients: 0,
  },
  practiceMetrics: {
    total: 0,
    enrolled: 0,
    pending: 0,
    declined: 0,
  },
  lastUpdated: null,
};
 
global.storedData = storedData;
 
export default function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
 
  // =========================
  // OPTIONS
  // =========================
  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }
 
  // =========================
  // POST
  // =========================
  if (req.method === "POST") {
    try {
      const incomingData = req.body || {};
 
      // =========================
      // PATIENTS
      // =========================
      if (
        incomingData.patients &&
        Array.isArray(incomingData.patients)
      ) {
        storedData.patients = incomingData.patients.map((p) => ({
          name: p["Patient Name"] || p.name || "",
          practice: p["Practice"] || p.practice || "",
          location: p["Location"] || p.location || "",
          navigator:
            p["Navigator Assigned"] || p.navigator || "",
 
          tcmScheduled:
            p["TCM Appt Scheduled?"] === "Yes" ||
            p.tcmScheduled === true,
 
          visitVerified:
            p["Visit Verified"] === "Yes" ||
            p.visitVerified === true,
        }));
 
        storedData.metrics = calculatePatientMetrics(
          storedData.patients
        );
      }
 
      // =========================
      // PRACTICES
      // =========================
      if (
        incomingData.practices &&
        Array.isArray(incomingData.practices)
      ) {
        storedData.practices = incomingData.practices;
 
        storedData.practiceMetrics =
          calculatePracticeMetrics(
            incomingData.practices
          );
      }
 
      // =========================
      // LAST UPDATED
      // =========================
      storedData.lastUpdated =
        new Date().toISOString();
 
      global.storedData = storedData;
 
      return res.status(200).json({
        success: true,
        message: "Data updated successfully",
        patientsProcessed:
          storedData.patients.length,
        practicesProcessed:
          storedData.practices.length,
        lastUpdated: storedData.lastUpdated,
      });
 
    } catch (error) {
      console.error("POST Error:", error);
 
      return res.status(500).json({
        error: "Failed to process data",
      });
    }
  }
 
  // =========================
  // GET
  // =========================
  else if (req.method === "GET") {
 
    // RETURN LIVE DATA ONLY
    return res.status(200).json(storedData);
  }
 
  // =========================
  // INVALID METHOD
  // =========================
  else {
    res.setHeader("Allow", ["GET", "POST"]);
 
    return res.status(405).json({
      error: `Method ${req.method} not allowed`,
    });
  }
}
 
// ======================================
// PATIENT METRICS
// ======================================
function calculatePatientMetrics(patients) {
 
  const total = patients.length;
 
  const tcmScheduled = patients.filter(
    (p) => p.tcmScheduled
  ).length;
 
  const visitVerified = patients.filter(
    (p) => p.visitVerified
  ).length;
 
  const devanPatients = patients.filter(
    (p) =>
      (p.navigator || "")
        .toLowerCase()
        .includes("devan")
  ).length;
 
  const sunniePatients = patients.filter(
    (p) =>
      (p.navigator || "")
        .toLowerCase()
        .includes("sunnie")
  ).length;
 
  const tvhPatients = patients.filter(
    (p) => p.location === "TVH"
  ).length;
 
  const ivmPatients = patients.filter(
    (p) => p.location === "IVM"
  ).length;
 
  return {
    totalPatients: total,
 
    tcmScheduled,
 
    visitVerified,
 
    tcmSchedulingRate:
      total > 0
        ? Number(
            (
              (tcmScheduled / total) *
              100
            ).toFixed(1)
          )
        : 0,
 
    verificationRate:
      tcmScheduled > 0
        ? Number(
            (
              (visitVerified /
                tcmScheduled) *
              100
            ).toFixed(1)
          )
        : 0,
 
    devanPatients,
 
    sunniePatients,
 
    tvhPatients,
 
    ivmPatients,
  };
}
 
// ======================================
// PRACTICE METRICS
// ======================================
function calculatePracticeMetrics(practices) {
 
  const total = practices.length;
 
  const enrolled = practices.filter(
    (p) =>
      p["PDV Forms Completed"] ===
        "Complete" ||
      p.status === "enrolled"
  ).length;
 
  const declined = practices.filter(
    (p) =>
      p["PDV Forms Completed"] ===
        "Declined" ||
      p.status === "declined"
  ).length;
 
  const pending =
    total - enrolled - declined;
 
  return {
    total,
    enrolled,
    pending,
    declined,
  };
}

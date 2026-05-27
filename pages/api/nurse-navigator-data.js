let storedData = global.storedData || {
  patients: [],
  practices: [],
  metrics: {},
  practiceMetrics: {},
  lastUpdated: null,
};
 
global.storedData = storedData;
 
export default function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
 
  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }
 
  if (req.method === "POST") {
    try {
      const incomingData = req.body || {};
 
      storedData = {
        patients: incomingData.patients || [],
        practices: incomingData.practices || [],
        metrics: incomingData.metrics || calculatePatientMetrics(incomingData.patients || []),
        practiceMetrics:
          incomingData.practiceMetrics || calculatePracticeMetrics(incomingData.practices || []),
        lastUpdated: new Date().toISOString(),
      };
 
      global.storedData = storedData;
 
      return res.status(200).json({
        success: true,
        message: "Data updated successfully",
        patientsProcessed: storedData.patients.length,
        practicesProcessed: storedData.practices.length,
        lastUpdated: storedData.lastUpdated,
      });
    } catch (error) {
      console.error("Error processing data:", error);
      return res.status(500).json({ error: "Failed to process data" });
    }
  }
 
  if (req.method === "GET") {
    return res.status(200).json(storedData);
  }
 
  res.setHeader("Allow", ["GET", "POST"]);
  return res.status(405).json({ error: `Method ${req.method} not allowed` });
}
 
function calculatePatientMetrics(patients) {
  const total = patients.length;
 
  const tcmScheduled = patients.filter(
    (p) => p["TCM Appt Scheduled?"] === "Yes" || p.tcmScheduled === true
  ).length;
 
  const visitVerified = patients.filter(
    (p) => p["Visit Verified"] === "Yes" || p.visitVerified === true
  ).length;
 
  return {
    totalPatients: total,
    tcmScheduled,
    visitVerified,
    tcmSchedulingRate: total > 0 ? Number(((tcmScheduled / total) * 100).toFixed(1)) : 0,
    verificationRate:
      tcmScheduled > 0 ? Number(((visitVerified / tcmScheduled) * 100).toFixed(1)) : 0,
  };
}
 
function calculatePracticeMetrics(practices) {
  const total = practices.length;
 
  const enrolled = practices.filter(
    (p) => p["PDV Forms Completed"] === "Complete" || p.status === "enrolled"
  ).length;
 
  const declined = practices.filter(
    (p) => p["PDV Forms Completed"] === "Declined" || p.status === "declined"
  ).length;
 
  const pending = total - enrolled - declined;
 
  return {
    total,
    enrolled,
    pending,
    declined,
  };
}

// API endpoint to receive and serve nurse navigator data
// This receives data from Power Automate and stores it

let storedData = {
  patients: [],
  metrics: {
    totalPatients: 0,
    tcmScheduled: 0,
    visitVerified: 0
  },
  lastUpdated: null
};

export default function handler(req, res) {
  // Enable CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  // POST - Receive data from Power Automate
  if (req.method === 'POST') {
    try {
      const incomingData = req.body;
      
      // Process and store the data
      if (incomingData.patients && Array.isArray(incomingData.patients)) {
        storedData = {
          patients: incomingData.patients.map(p => ({
            name: p['Patient Name'] || p.name,
            practice: p['Practice'] || p.practice,
            location: p['Location'] || p.location,
            navigator: p['Navigator Assigned'] || p.navigator,
            tcmScheduled: p['TCM Appt Scheduled?'] === 'Yes' || p.tcmScheduled,
            visitVerified: p['Visit Verified'] === 'Yes' || p.visitVerified
          })),
          metrics: calculateMetrics(incomingData.patients),
          lastUpdated: new Date().toISOString()
        };

        return res.status(200).json({
          success: true,
          message: 'Data updated successfully',
          patientsProcessed: storedData.patients.length
        });
      }

      return res.status(400).json({ error: 'Invalid data format' });
      
    } catch (error) {
      console.error('Error processing data:', error);
      return res.status(500).json({ error: 'Failed to process data' });
    }
  }
  
  // GET - Return current data to dashboard
  else if (req.method === 'GET') {
    if (storedData.patients.length === 0) {
      // Return sample data if no real data yet
      return res.status(200).json({
        patients: [
          {
            name: 'Sample Patient 1',
            practice: 'Sample Practice',
            location: 'TVH',
            navigator: 'Devan Lambruno',
            tcmScheduled: true,
            visitVerified: true
          },
          {
            name: 'Sample Patient 2',
            practice: 'Sample Practice 2',
            location: 'IVM',
            navigator: 'Sunnie Emberson',
            tcmScheduled: true,
            visitVerified: false
          }
        ],
        metrics: {
          totalPatients: 2,
          tcmScheduled: 2,
          visitVerified: 1
        },
        lastUpdated: new Date().toISOString(),
        note: 'Sample data - waiting for Power Automate connection'
      });
    }

    return res.status(200).json(storedData);
  }
  
  else {
    res.setHeader('Allow', ['GET', 'POST']);
    return res.status(405).json({ error: `Method ${req.method} not allowed` });
  }
}

function calculateMetrics(patients) {
  const total = patients.length;
  const tcmScheduled = patients.filter(p => 
    p['TCM Appt Scheduled?'] === 'Yes' || p.tcmScheduled
  ).length;
  const visitVerified = patients.filter(p => 
    p['Visit Verified'] === 'Yes' || p.visitVerified
  ).length;

  return {
    totalPatients: total,
    tcmScheduled: tcmScheduled,
    visitVerified: visitVerified,
    tcmSchedulingRate: total > 0 ? ((tcmScheduled / total) * 100).toFixed(1) : 0,
    verificationRate: tcmScheduled > 0 ? ((visitVerified / tcmScheduled) * 100).toFixed(1) : 0
  };
}

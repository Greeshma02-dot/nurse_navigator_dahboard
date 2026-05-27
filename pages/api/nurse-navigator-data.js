// Enhanced API endpoint for both patient and practice data
let storedData = {
  patients: [],
  practices: [],
  metrics: {
    totalPatients: 0,
    tcmScheduled: 0,
    visitVerified: 0,
    devanPatients: 0,
    sunniePatients: 0,
    tvhPatients: 0,
    ivmPatients: 0
  },
  practiceMetrics: {
    total: 33,
    enrolled: 18,
    pending: 8,
    declined: 7
  },
  lastUpdated: null
};

export default function handler(req, res) {
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
      
      // Process patient data
      if (incomingData.patients && Array.isArray(incomingData.patients)) {
        storedData.patients = incomingData.patients.map(p => ({
          name: p['Patient Name'] || p.name,
          practice: p['Practice'] || p.practice,
          location: p['Location'] || p.location,
          navigator: p['Navigator Assigned'] || p.navigator,
          tcmScheduled: p['TCM Appt Scheduled?'] === 'Yes' || p.tcmScheduled,
          visitVerified: p['Visit Verified'] === 'Yes' || p.visitVerified
        }));
        
        storedData.metrics = calculatePatientMetrics(incomingData.patients);
      }

      // Process practice data
      if (incomingData.practices && Array.isArray(incomingData.practices)) {
        storedData.practices = incomingData.practices;
        storedData.practiceMetrics = calculatePracticeMetrics(incomingData.practices);
      }

      storedData.lastUpdated = new Date().toISOString();

      return res.status(200).json({
        success: true,
        message: 'Data updated successfully',
        patientsProcessed: storedData.patients.length,
        practicesProcessed: storedData.practices.length
      });
      
    } catch (error) {
      console.error('Error processing data:', error);
      return res.status(500).json({ error: 'Failed to process data' });
    }
  }
  
  // GET - Return current data to dashboard
  else if (req.method === 'GET') {
    // Return sample data if no real data yet
    if (storedData.patients.length === 0) {
      return res.status(200).json({
        patients: [
          {
            name: 'Kellie Woodis',
            practice: 'Brian Nguyen',
            location: 'TVH',
            navigator: 'Devan Lambruno',
            tcmScheduled: true,
            visitVerified: true
          },
          {
            name: 'Sally Borra',
            practice: 'David H Nguyen',
            location: 'IVM',
            navigator: 'Sunnie Emberson',
            tcmScheduled: true,
            visitVerified: true
          },
          {
            name: 'Jamal Safa',
            practice: 'Temecula Valley PCP',
            location: 'IVM',
            navigator: 'Sunnie Emberson',
            tcmScheduled: true,
            visitVerified: false
          },
          {
            name: 'Mary Rouse',
            practice: 'Jeremy V Gomer',
            location: 'TVH',
            navigator: 'Devan Lambruno',
            tcmScheduled: true,
            visitVerified: true
          },
          {
            name: 'Penelope Litonjua',
            practice: 'Ocampo',
            location: 'TVH',
            navigator: 'Devan Lambruno',
            tcmScheduled: false,
            visitVerified: false
          }
        ],
        practices: [],
        metrics: {
          totalPatients: 15,
          tcmScheduled: 4,
          visitVerified: 3,
          tcmSchedulingRate: 26.7,
          verificationRate: 75.0,
          devanPatients: 11,
          sunniePatients: 2,
          tvhPatients: 13,
          ivmPatients: 2
        },
        practiceMetrics: {
          total: 33,
          enrolled: 18,
          pending: 8,
          declined: 7
        },
        lastUpdated: new Date().toISOString(),
        note: 'Sample data - Connect Power Automate for real data'
      });
    }

    return res.status(200).json(storedData);
  }
  
  else {
    res.setHeader('Allow', ['GET', 'POST']);
    return res.status(405).json({ error: `Method ${req.method} not allowed` });
  }
}

function calculatePatientMetrics(patients) {
  const total = patients.length;
  const tcmScheduled = patients.filter(p => 
    p['TCM Appt Scheduled?'] === 'Yes' || p.tcmScheduled
  ).length;
  const visitVerified = patients.filter(p => 
    p['Visit Verified'] === 'Yes' || p.visitVerified
  ).length;
  const devanPatients = patients.filter(p => 
    (p['Navigator Assigned'] || p.navigator || '').includes('Devan')
  ).length;
  const sunniePatients = patients.filter(p => 
    (p['Navigator Assigned'] || p.navigator || '').includes('Sunnie')
  ).length;
  const tvhPatients = patients.filter(p => 
    (p['Location'] || p.location) === 'TVH'
  ).length;
  const ivmPatients = patients.filter(p => 
    (p['Location'] || p.location) === 'IVM'
  ).length;

  return {
    totalPatients: total,
    tcmScheduled: tcmScheduled,
    visitVerified: visitVerified,
    tcmSchedulingRate: total > 0 ? ((tcmScheduled / total) * 100).toFixed(1) : 0,
    verificationRate: tcmScheduled > 0 ? ((visitVerified / tcmScheduled) * 100).toFixed(1) : 0,
    devanPatients: devanPatients,
    sunniePatients: sunniePatients,
    tvhPatients: tvhPatients,
    ivmPatients: ivmPatients
  };
}

function calculatePracticeMetrics(practices) {
  const total = practices.length;
  const enrolled = practices.filter(p => 
    p['PDV Forms Completed'] === 'Complete' || p.status === 'enrolled'
  ).length;
  const declined = practices.filter(p => 
    p['PDV Forms Completed'] === 'Declined' || p.status === 'declined'
  ).length;
  const pending = total - enrolled - declined;

  return {
    total: total,
    enrolled: enrolled,
    pending: pending,
    declined: declined
  };
}

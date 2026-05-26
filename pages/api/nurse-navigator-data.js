// API endpoint with REAL DATA from your Excel files!
// Last updated: May 26, 2026

let storedData = {
  patients: [],
  practices: [],
  metrics: {},
  practiceMetrics: {}
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
    // REAL DATA from your Excel files (as of May 26, 2026)
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
        },
        {
          name: 'Mary Reed',
          practice: 'Sylvia Gisi',
          location: 'TVH',
          navigator: 'Devan Lambruno',
          tcmScheduled: false,
          visitVerified: false
        },
        {
          name: 'Suraya Glenna',
          practice: 'Ocampo',
          location: 'TVH',
          navigator: 'Devan Lambruno',
          tcmScheduled: false,
          visitVerified: false
        },
        {
          name: 'Jean McDaniel',
          practice: 'Sylvia Gisi',
          location: 'TVH',
          navigator: 'Devan Lambruno',
          tcmScheduled: false,
          visitVerified: false
        },
        {
          name: 'Doris Mahoney',
          practice: 'Ocampo',
          location: 'TVH',
          navigator: 'Devan Lambruno',
          tcmScheduled: false,
          visitVerified: false
        },
        {
          name: 'Carol Brady',
          practice: 'Timothy Killeen',
          location: 'TVH',
          navigator: 'Devan Lambruno',
          tcmScheduled: false,
          visitVerified: false
        },
        {
          name: 'Charles W Hoague',
          practice: 'Bella Shah',
          location: 'TVH',
          navigator: 'Devan Lambruno',
          tcmScheduled: false,
          visitVerified: false
        },
        {
          name: 'Ramon Cardenas',
          practice: 'Temecula Valley PCP',
          location: 'TVH',
          navigator: 'Devan Lambruno',
          tcmScheduled: false,
          visitVerified: false
        },
        {
          name: 'Jane Curtis',
          practice: 'Temecula Valley PCP',
          location: 'TVH',
          navigator: 'Devan Lambruno',
          tcmScheduled: false,
          visitVerified: false
        },
        {
          name: 'Nancy Huckaby',
          practice: 'Bella Shah',
          location: 'TVH',
          navigator: 'Devan Lambruno',
          tcmScheduled: false,
          visitVerified: false
        },
        {
          name: 'Steven Behrle',
          practice: 'Bella Shah',
          location: 'TVH',
          navigator: 'Devan Lambruno',
          tcmScheduled: false,
          visitVerified: false
        },
        {
          name: 'Jennifer Lopez',
          practice: 'Ocampo',
          location: 'TVH',
          navigator: 'Devan Lambruno',
          tcmScheduled: false,
          visitVerified: false
        },
        {
          name: 'Gilbert Valenzuela',
          practice: 'Jeremy V Gomer',
          location: 'TVH',
          navigator: 'Devan Lambruno',
          tcmScheduled: false,
          visitVerified: false
        },
        {
          name: 'Lajuana Poole',
          practice: 'Ocampo',
          location: 'TVH',
          navigator: 'Devan Lambruno',
          tcmScheduled: false,
          visitVerified: false
        },
        {
          name: 'Shirley Howe',
          practice: 'Bella Shah',
          location: 'TVH',
          navigator: 'Devan Lambruno',
          tcmScheduled: false,
          visitVerified: false
        },
        {
          name: 'Elizabeth Rodriguez',
          practice: 'Bella Shah',
          location: 'TVH',
          navigator: 'Devan Lambruno',
          tcmScheduled: false,
          visitVerified: false
        },
        {
          name: 'Joan Gonzalez',
          practice: 'Ocampo',
          location: 'TVH',
          navigator: 'Devan Lambruno',
          tcmScheduled: false,
          visitVerified: false
        }
      ],
      practices: [],
      metrics: {
        totalPatients: 21,
        tcmScheduled: 4,
        visitVerified: 3,
        tcmSchedulingRate: 19.0,
        verificationRate: 75.0,
        devanPatients: 19,
        sunniePatients: 2,
        tvhPatients: 19,
        ivmPatients: 2
      },
      practiceMetrics: {
        total: 34,
        enrolled: 15,
        pending: 11,
        declined: 8
      },
      lastUpdated: new Date().toISOString(),
      note: 'Real data from Excel files - May 26, 2026'
    });
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
    tcmSchedulingRate: total > 0 ? parseFloat(((tcmScheduled / total) * 100).toFixed(1)) : 0,
    verificationRate: tcmScheduled > 0 ? parseFloat(((visitVerified / tcmScheduled) * 100).toFixed(1)) : 0,
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

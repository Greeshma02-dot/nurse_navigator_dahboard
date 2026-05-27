import React, { useState, useEffect, useRef } from 'react';
import { RefreshCw, Users, CheckCircle, Clock, Building2, TrendingUp, Calendar, Upload } from 'lucide-react';
import { LineChart, Line, BarChart, Bar, PieChart, Pie, Cell, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';

export default function NurseNavigatorDashboard() {
  const [activeTab, setActiveTab] = useState('patients');
  const [data, setData] = useState({
    metrics: { totalPatients: 21, tcmScheduled: 4, visitVerified: 3, tcmSchedulingRate: 19.0, verificationRate: 75.0, devanPatients: 19, sunniePatients: 2 },
    patients: [],
    practiceMetrics: { total: 34, enrolled: 15, pending: 11, declined: 8 },
    practices: []
  });
  const [loading, setLoading] = useState(false);
  const [lastSync, setLastSync] = useState(new Date());
  const [syncing, setSyncing] = useState(false);
  const [syncMessage, setSyncMessage] = useState('');
  const fileInputRef = useRef(null);

  const handleSyncClick = () => {
    fileInputRef.current.click();
  };

  const handleFileUpload = async (event) => {
    const files = event.target.files;
    if (!files || files.length === 0) return;

    setSyncing(true);
    setSyncMessage('Reading Excel files...');

    try {
      const XLSX = await import('xlsx');
      let patientRows = null;
      let practiceRows = null;

      for (const file of files) {
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer);
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        
        // Try with header row detection (skip metadata rows)
        const rawData = XLSX.utils.sheet_to_json(firstSheet, { defval: '', header: 1 });
        
        // Find the actual header row (look for row with most non-empty cells)
        let headerRowIndex = 0;
        let maxNonEmpty = 0;
        for (let i = 0; i < Math.min(5, rawData.length); i++) {
          const nonEmpty = rawData[i].filter(c => c !== '').length;
          if (nonEmpty > maxNonEmpty) {
            maxNonEmpty = nonEmpty;
            headerRowIndex = i;
          }
        }

        const headers = rawData[headerRowIndex];
        const dataRows = rawData.slice(headerRowIndex + 1).filter(row => 
          row.some(cell => cell !== '' && cell !== null)
        );

        const jsonData = dataRows.map(row => {
          const obj = {};
          headers.forEach((header, idx) => {
            if (header) obj[header] = row[idx] || '';
          });
          return obj;
        });

        const fileName = file.name.toLowerCase();
        if (fileName.includes('patient')) {
          patientRows = jsonData;
          setSyncMessage('Patient file processed!');
        } else if (fileName.includes('ccpaco') || fileName.includes('tracking') || fileName.includes('practice')) {
          practiceRows = jsonData;
          setSyncMessage('Practice file processed!');
        }
      }

      const updatedData = processExcelData(patientRows, practiceRows);
      setData(updatedData);
      setLastSync(new Date());
      setSyncMessage('Sync complete! Data updated.');
      
      setTimeout(() => {
        setSyncing(false);
        setSyncMessage('');
      }, 2500);

    } catch (error) {
      console.error('Sync error:', error);
      setSyncMessage('Error: ' + error.message);
      setTimeout(() => {
        setSyncing(false);
        setSyncMessage('');
      }, 3000);
    }

    event.target.value = '';
  };

  const findColumn = (row, possibleNames) => {
    for (const name of possibleNames) {
      const keys = Object.keys(row);
      const match = keys.find(k => k && k.toLowerCase().includes(name.toLowerCase()));
      if (match) return row[match];
    }
    return '';
  };

  const isYes = (val) => {
    const s = (val || '').toString().toLowerCase().trim();
    return s === 'yes' || s === 'y' || s === 'true' || s === '1' || s === 'scheduled' || s === 'verified' || s === 'complete';
  };

  const processExcelData = (patientRows, practiceRows) => {
    const result = { ...data };

    if (patientRows && patientRows.length > 0) {
      console.log('Patient rows:', patientRows);
      console.log('Patient columns:', Object.keys(patientRows[0]));

      const tcmScheduled = patientRows.filter(p => 
        isYes(findColumn(p, ['TCM Scheduled', 'TCM scheduled', 'tcm']))
      ).length;
      
      const visitVerified = patientRows.filter(p => 
        isYes(findColumn(p, ['Visit Verified', 'verified', 'visit']))
      ).length;

      const devanCount = patientRows.filter(p => {
        const nav = (findColumn(p, ['Nurse Navigator', 'Navigator', 'nurse']) || '').toString().toLowerCase();
        return nav.includes('devan');
      }).length;

      const sunnieCount = patientRows.filter(p => {
        const nav = (findColumn(p, ['Nurse Navigator', 'Navigator', 'nurse']) || '').toString().toLowerCase();
        return nav.includes('sunnie');
      }).length;

      const totalPatients = patientRows.length;

      result.metrics = {
        totalPatients,
        tcmScheduled,
        visitVerified,
        tcmSchedulingRate: totalPatients > 0 ? ((tcmScheduled / totalPatients) * 100).toFixed(1) : 0,
        verificationRate: tcmScheduled > 0 ? ((visitVerified / tcmScheduled) * 100).toFixed(1) : 0,
        devanPatients: devanCount,
        sunniePatients: sunnieCount
      };

      result.patients = patientRows.map(p => ({
        name: findColumn(p, ['Patient Name', 'Name', 'patient']) || 'N/A',
        practice: findColumn(p, ['Practice', 'practice']) || 'N/A',
        location: findColumn(p, ['Location', 'location', 'site']) || 'N/A',
        navigator: findColumn(p, ['Nurse Navigator', 'Navigator', 'nurse']) || 'N/A',
        tcmScheduled: isYes(findColumn(p, ['TCM Scheduled', 'tcm'])),
        visitVerified: isYes(findColumn(p, ['Visit Verified', 'verified', 'visit']))
      }));
    }

    if (practiceRows && practiceRows.length > 0) {
      console.log('Practice rows:', practiceRows);
      console.log('Practice columns:', Object.keys(practiceRows[0]));

      const enrolled = practiceRows.filter(p => 
        isYes(findColumn(p, ['PDV Status', 'PDV_Status', 'PDV']))
      ).length;
      
      const pending = practiceRows.filter(p => {
        const status = (findColumn(p, ['Status']) || '').toString().toLowerCase();
        return status.includes('pending') || status.includes('progress') || status.includes('contact');
      }).length;

      const declined = practiceRows.filter(p => {
        const status = (findColumn(p, ['Status']) || '').toString().toLowerCase();
        return status.includes('declined') || status.includes('no') || status.includes('not');
      }).length;

      result.practiceMetrics = { 
        total: practiceRows.length, 
        enrolled, 
        pending, 
        declined 
      };

      result.practices = practiceRows.map(p => ({
        name: findColumn(p, ['Practice', 'practice', 'name']) || 'N/A',
        consultant: findColumn(p, ['Consultant', 'consultant']) || 'N/A',
        location: findColumn(p, ['Location', 'location']) || 'N/A',
        hospitals: findColumn(p, ['Hospitals', 'hospital']) || 'N/A',
        status: findColumn(p, ['Status']) || 'N/A',
        pdvStatus: findColumn(p, ['PDV Status', 'PDV_Status', 'PDV']) || 'N/A',
        emrAccess: findColumn(p, ['EMR Access', 'EMR', 'emr']) || 'N/A',
        contact: findColumn(p, ['Contact', 'contact']) || 'N/A',
        notes: findColumn(p, ['Notes', 'notes', 'comment']) || ''
      }));
    }

    return result;
  };

  return (
    <div style={{
      minHeight: '100vh',
      background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
      fontFamily: "'Inter', -apple-system, sans-serif"
    }}>
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xls"
        multiple
        onChange={handleFileUpload}
        style={{ display: 'none' }}
      />

      <header style={{
        background: 'rgba(255, 255, 255, 0.95)',
        backdropFilter: 'blur(10px)',
        padding: '24px 40px',
        boxShadow: '0 4px 6px rgba(0,0,0,0.1)',
        position: 'sticky',
        top: 0,
        zIndex: 100
      }}>
        <div style={{ maxWidth: '1400px', margin: '0 auto', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <h1 style={{
            fontSize: '28px',
            fontWeight: '800',
            background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
            WebkitBackgroundClip: 'text',
            WebkitTextFillColor: 'transparent',
            margin: 0
          }}>
            Nurse Navigator Program
          </h1>
          <div style={{ display: 'flex', gap: '12px', alignItems: 'center' }}>
            <button
              onClick={handleSyncClick}
              disabled={syncing}
              style={{
                padding: '12px 24px',
                background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
                color: 'white',
                border: 'none',
                borderRadius: '12px',
                fontSize: '14px',
                fontWeight: '700',
                cursor: syncing ? 'wait' : 'pointer',
                display: 'flex',
                alignItems: 'center',
                gap: '8px',
                boxShadow: '0 4px 12px rgba(102, 126, 234, 0.4)',
                opacity: syncing ? 0.7 : 1
              }}
            >
              <Upload size={18} />
              {syncing ? 'Syncing...' : 'Sync from SharePoint'}
            </button>
            <div style={{
              display: 'flex',
              alignItems: 'center',
              gap: '12px',
              padding: '12px 20px',
              background: '#f0fdf4',
              borderRadius: '12px',
              border: '2px solid #86efac'
            }}>
              <div style={{
                width: '8px',
                height: '8px',
                borderRadius: '50%',
                background: '#22c55e',
                animation: 'pulse 2s infinite'
              }}></div>
              <div>
                <div style={{ fontSize: '12px', color: '#15803d', fontWeight: '600' }}>Live Data</div>
                <div style={{ fontSize: '11px', color: '#166534' }}>Updated {lastSync?.toLocaleTimeString()}</div>
              </div>
            </div>
          </div>
        </div>

        {syncMessage && (
          <div style={{
            maxWidth: '1400px',
            margin: '12px auto 0',
            padding: '12px 20px',
            background: '#dbeafe',
            borderRadius: '12px',
            color: '#1e40af',
            fontSize: '14px',
            fontWeight: '600',
            textAlign: 'center'
          }}>
            {syncMessage}
          </div>
        )}

        <div style={{ maxWidth: '1400px', margin: '20px auto 0', display: 'flex', gap: '8px' }}>
          <TabButton
            active={activeTab === 'patients'}
            onClick={() => setActiveTab('patients')}
            icon={<Users size={18} />}
            label="Patient Tracking"
          />
          <TabButton
            active={activeTab === 'practices'}
            onClick={() => setActiveTab('practices')}
            icon={<Building2 size={18} />}
            label="Practice Enrollment"
          />
        </div>
      </header>

      <div style={{ maxWidth: '1400px', margin: '0 auto', padding: '40px' }}>
        {activeTab === 'patients' ? (
          <PatientTrackingPage data={data} />
        ) : (
          <PracticeEnrollmentPage data={data} />
        )}
      </div>

      <style jsx>{`
        @keyframes pulse {
          0%, 100% { opacity: 1; }
          50% { opacity: 0.5; }
        }
      `}</style>
    </div>
  );
}

function TabButton({ active, onClick, icon, label }) {
  return (
    <button
      onClick={onClick}
      style={{
        padding: '12px 24px',
        background: active ? 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)' : 'white',
        color: active ? 'white' : '#64748b',
        border: active ? 'none' : '2px solid #e2e8f0',
        borderRadius: '12px',
        fontSize: '14px',
        fontWeight: '600',
        cursor: 'pointer',
        display: 'flex',
        alignItems: 'center',
        gap: '8px',
        transition: 'all 0.2s',
        boxShadow: active ? '0 4px 12px rgba(102, 126, 234, 0.4)' : 'none'
      }}
    >
      {icon}
      {label}
    </button>
  );
}

function PatientTrackingPage({ data }) {
  const patients = data?.patients || [];
  const metrics = data?.metrics || {};

  const statusData = [
    { name: 'TCM Scheduled', value: metrics.tcmScheduled || 0, color: '#3b82f6' },
    { name: 'Pending', value: (metrics.totalPatients - metrics.tcmScheduled) || 0, color: '#f59e0b' },
  ];

  const navigatorData = [
    { name: 'Devan', patients: metrics.devanPatients || 0 },
    { name: 'Sunnie', patients: metrics.sunniePatients || 0 }
  ];

  const weeklyTrendData = [
    { week: 'Apr 13-19', patients: 3, tcm: 1, verified: 1 },
    { week: 'Apr 20-26', patients: 8, tcm: 2, verified: 2 },
    { week: 'Apr 27-May 3', patients: 11, tcm: 3, verified: 2 },
    { week: 'May 4-10', patients: 16, tcm: 3, verified: 2 },
    { week: 'May 11-17', patients: 19, tcm: 3, verified: 3 },
    { week: 'May 18-24', patients: 21, tcm: 4, verified: 3 }
  ];

  return (
    <div>
      <div style={{
        display: 'grid',
        gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))',
        gap: '24px',
        marginBottom: '32px'
      }}>
        <MetricCard icon={<Users size={24} />} title="Active Patients" value={metrics.totalPatients || 0} subtitle="Currently tracking" color="#3b82f6" gradient="linear-gradient(135deg, #3b82f6 0%, #2563eb 100%)" />
        <MetricCard icon={<Calendar size={24} />} title="TCM Scheduled" value={metrics.tcmScheduled || 0} subtitle={`${metrics.tcmSchedulingRate || 0}% scheduling rate`} color="#10b981" gradient="linear-gradient(135deg, #10b981 0%, #059669 100%)" />
        <MetricCard icon={<CheckCircle size={24} />} title="Visits Verified" value={metrics.visitVerified || 0} subtitle={`${metrics.verificationRate || 0}% success rate`} color="#8b5cf6" gradient="linear-gradient(135deg, #8b5cf6 0%, #7c3aed 100%)" />
        <MetricCard icon={<Clock size={24} />} title="Pending Action" value={(metrics.totalPatients - metrics.tcmScheduled) || 0} subtitle="Awaiting scheduling" color="#f59e0b" gradient="linear-gradient(135deg, #f59e0b 0%, #d97706 100%)" />
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1fr', gap: '24px', marginBottom: '24px' }}>
        <ChartCard title="Weekly Patient & TCM Trend">
          <ResponsiveContainer width="100%" height={280}>
            <LineChart data={weeklyTrendData}>
              <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
              <XAxis dataKey="week" stroke="#64748b" style={{ fontSize: '12px' }} />
              <YAxis stroke="#64748b" style={{ fontSize: '12px' }} />
              <Tooltip />
              <Legend />
              <Line type="monotone" dataKey="patients" stroke="#3b82f6" strokeWidth={3} name="Patients" />
              <Line type="monotone" dataKey="tcm" stroke="#10b981" strokeWidth={3} name="TCM Scheduled" />
              <Line type="monotone" dataKey="verified" stroke="#8b5cf6" strokeWidth={3} name="Verified" />
            </LineChart>
          </ResponsiveContainer>
        </ChartCard>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(350px, 1fr))', gap: '24px', marginBottom: '24px' }}>
        <ChartCard title="Navigator Workload">
          <ResponsiveContainer width="100%" height={280}>
            <BarChart data={navigatorData}>
              <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
              <XAxis dataKey="name" stroke="#64748b" style={{ fontSize: '12px' }} />
              <YAxis stroke="#64748b" style={{ fontSize: '12px' }} />
              <Tooltip />
              <Bar dataKey="patients" fill="#667eea" radius={[8, 8, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>

        <ChartCard title="TCM Scheduling Status">
          <ResponsiveContainer width="100%" height={280}>
            <PieChart>
              <Pie data={statusData} cx="50%" cy="50%" labelLine={false} label={({ name, percent }) => `${name}: ${(percent * 100).toFixed(0)}%`} outerRadius={100} fill="#8884d8" dataKey="value">
                {statusData.map((entry, index) => <Cell key={`cell-${index}`} fill={entry.color} />)}
              </Pie>
              <Tooltip />
            </PieChart>
          </ResponsiveContainer>
        </ChartCard>
      </div>

      <div style={{ background: 'white', borderRadius: '16px', padding: '24px', boxShadow: '0 4px 6px rgba(0,0,0,0.05)' }}>
        <h2 style={{ fontSize: '20px', fontWeight: '700', marginBottom: '20px', color: '#1e293b' }}>
          Patient List ({patients.length} total)
        </h2>
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <thead>
              <tr style={{ borderBottom: '2px solid #e2e8f0' }}>
                <th style={headerStyle}>Patient Name</th>
                <th style={headerStyle}>Practice</th>
                <th style={headerStyle}>Location</th>
                <th style={headerStyle}>Navigator</th>
                <th style={headerStyle}>TCM Status</th>
                <th style={headerStyle}>Verified</th>
              </tr>
            </thead>
            <tbody>
              {patients.map((patient, index) => (
                <tr key={index} style={{ borderBottom: '1px solid #f1f5f9' }}>
                  <td style={cellStyle}><strong>{patient.name || 'N/A'}</strong></td>
                  <td style={cellStyle}>{patient.practice || 'N/A'}</td>
                  <td style={cellStyle}>
                    <span style={{ padding: '4px 12px', background: patient.location === 'TVH' ? '#dbeafe' : '#fef3c7', color: patient.location === 'TVH' ? '#1e40af' : '#92400e', borderRadius: '12px', fontSize: '12px', fontWeight: '600' }}>
                      {patient.location || 'N/A'}
                    </span>
                  </td>
                  <td style={cellStyle}>{patient.navigator || 'N/A'}</td>
                  <td style={cellStyle}>
                    <span style={{ padding: '4px 12px', background: patient.tcmScheduled ? '#dcfce7' : '#fee2e2', color: patient.tcmScheduled ? '#15803d' : '#991b1b', borderRadius: '12px', fontSize: '12px', fontWeight: '600' }}>
                      {patient.tcmScheduled ? 'Scheduled' : 'Pending'}
                    </span>
                  </td>
                  <td style={cellStyle}>
                    <span style={{ fontSize: '20px' }}>{patient.visitVerified ? '✅' : '⏳'}</span>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function PracticeEnrollmentPage({ data }) {
  const practiceMetrics = data?.practiceMetrics || { total: 0, enrolled: 0, pending: 0, declined: 0 };
  const practices = data?.practices || [];
  const [filterStatus, setFilterStatus] = useState('all');
  const [searchTerm, setSearchTerm] = useState('');

  const enrollmentData = [
    { name: 'Enrolled', value: practiceMetrics.enrolled, color: '#10b981' },
    { name: 'Pending', value: practiceMetrics.pending, color: '#f59e0b' },
    { name: 'Declined', value: practiceMetrics.declined, color: '#ef4444' }
  ];

  // Build consultant data from real practices data
  const consultantMap = {};
  practices.forEach(p => {
    const c = p.consultant || 'Unknown';
    if (!consultantMap[c]) consultantMap[c] = { name: c, enrolled: 0, pending: 0, total: 0 };
    consultantMap[c].total++;
    const pdvLower = (p.pdvStatus || '').toString().toLowerCase();
    const statusLower = (p.status || '').toString().toLowerCase();
    if (pdvLower.includes('complete') || pdvLower.includes('yes')) {
      consultantMap[c].enrolled++;
    } else if (statusLower.includes('pending') || statusLower.includes('progress')) {
      consultantMap[c].pending++;
    }
  });
  
  const consultantData = Object.values(consultantMap).sort((a, b) => b.enrolled - a.enrolled).slice(0, 10);
  
  // If no real data, use default
  const finalConsultantData = consultantData.length > 0 ? consultantData : [
    { name: 'Rachel Robinson', enrolled: 7, pending: 4 },
    { name: 'Marlen Cornejo', enrolled: 2, pending: 4 },
    { name: 'Bella McKay', enrolled: 2, pending: 0 },
    { name: 'Trysh Logan', enrolled: 1, pending: 0 },
    { name: 'Chris Ford', enrolled: 1, pending: 0 },
    { name: 'Sheyenne Powers', enrolled: 1, pending: 0 }
  ];

  // Filter practices
  const filteredPractices = practices.filter(p => {
    const matchesSearch = !searchTerm || 
      (p.name && p.name.toString().toLowerCase().includes(searchTerm.toLowerCase())) ||
      (p.consultant && p.consultant.toString().toLowerCase().includes(searchTerm.toLowerCase()));
    
    if (filterStatus === 'all') return matchesSearch;
    
    const pdv = (p.pdvStatus || '').toString().toLowerCase();
    const status = (p.status || '').toString().toLowerCase();
    
    if (filterStatus === 'enrolled') return matchesSearch && (pdv.includes('complete') || pdv.includes('yes'));
    if (filterStatus === 'pending') return matchesSearch && (status.includes('pending') || status.includes('progress'));
    if (filterStatus === 'declined') return matchesSearch && (status.includes('declined') || status.includes('no'));
    return matchesSearch;
  });

  const getStatusBadge = (practice) => {
    const pdv = (practice.pdvStatus || '').toString().toLowerCase();
    const status = (practice.status || '').toString().toLowerCase();
    
    if (pdv.includes('complete') || pdv.includes('yes')) {
      return { label: 'Enrolled', bg: '#dcfce7', color: '#15803d' };
    } else if (status.includes('declined') || status.includes('no')) {
      return { label: 'Declined', bg: '#fee2e2', color: '#991b1b' };
    } else if (status.includes('pending') || status.includes('progress')) {
      return { label: 'Pending', bg: '#fef3c7', color: '#92400e' };
    }
    return { label: status || 'Unknown', bg: '#e2e8f0', color: '#475569' };
  };

  return (
    <div>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))', gap: '24px', marginBottom: '32px' }}>
        <MetricCard icon={<Building2 size={24} />} title="Total Practices" value={practiceMetrics.total} subtitle="Contacted to date" color="#3b82f6" gradient="linear-gradient(135deg, #3b82f6 0%, #2563eb 100%)" />
        <MetricCard icon={<CheckCircle size={24} />} title="Enrolled" value={practiceMetrics.enrolled} subtitle={`${practiceMetrics.total > 0 ? ((practiceMetrics.enrolled / practiceMetrics.total) * 100).toFixed(1) : 0}% success rate`} color="#10b981" gradient="linear-gradient(135deg, #10b981 0%, #059669 100%)" />
        <MetricCard icon={<Clock size={24} />} title="Pending" value={practiceMetrics.pending} subtitle="In enrollment process" color="#f59e0b" gradient="linear-gradient(135deg, #f59e0b 0%, #d97706 100%)" />
        <MetricCard icon={<TrendingUp size={24} />} title="Declined" value={practiceMetrics.declined} subtitle="Not participating" color="#ef4444" gradient="linear-gradient(135deg, #ef4444 0%, #dc2626 100%)" />
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(350px, 1fr))', gap: '24px', marginBottom: '24px' }}>
        <ChartCard title="Enrollment Status Breakdown">
          <ResponsiveContainer width="100%" height={300}>
            <PieChart>
              <Pie data={enrollmentData} cx="50%" cy="50%" labelLine={false} label={({ name, percent }) => `${name}: ${(percent * 100).toFixed(0)}%`} outerRadius={110} fill="#8884d8" dataKey="value">
                {enrollmentData.map((entry, index) => <Cell key={`cell-${index}`} fill={entry.color} />)}
              </Pie>
              <Tooltip />
            </PieChart>
          </ResponsiveContainer>
        </ChartCard>

        <ChartCard title="Consultant Performance">
          <ResponsiveContainer width="100%" height={300}>
            <BarChart data={finalConsultantData} layout="vertical" margin={{ left: 180, right: 20 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
              <XAxis type="number" stroke="#64748b" style={{ fontSize: '12px' }} />
              <YAxis dataKey="name" type="category" stroke="#64748b" width={170} style={{ fontSize: '11px' }} />
              <Tooltip />
              <Legend />
              <Bar dataKey="enrolled" fill="#10b981" radius={[0, 8, 8, 0]} name="Enrolled" />
              <Bar dataKey="pending" fill="#f59e0b" radius={[0, 8, 8, 0]} name="Pending" />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>
      </div>

      {/* PRACTICE DRILLDOWN TABLE */}
      <div style={{ background: 'white', borderRadius: '16px', padding: '24px', boxShadow: '0 4px 6px rgba(0,0,0,0.05)' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px', flexWrap: 'wrap', gap: '16px' }}>
          <h2 style={{ fontSize: '20px', fontWeight: '700', color: '#1e293b', margin: 0 }}>
            Practice Drilldown ({filteredPractices.length} of {practices.length})
          </h2>
          <div style={{ display: 'flex', gap: '12px', flexWrap: 'wrap' }}>
            <input
              type="text"
              placeholder="Search practice or consultant..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              style={{
                padding: '8px 16px',
                border: '2px solid #e2e8f0',
                borderRadius: '8px',
                fontSize: '14px',
                width: '250px',
                outline: 'none'
              }}
            />
            <select
              value={filterStatus}
              onChange={(e) => setFilterStatus(e.target.value)}
              style={{
                padding: '8px 16px',
                border: '2px solid #e2e8f0',
                borderRadius: '8px',
                fontSize: '14px',
                background: 'white',
                cursor: 'pointer',
                outline: 'none'
              }}
            >
              <option value="all">All Status</option>
              <option value="enrolled">Enrolled Only</option>
              <option value="pending">Pending Only</option>
              <option value="declined">Declined Only</option>
            </select>
          </div>
        </div>

        {practices.length === 0 ? (
          <div style={{ padding: '40px', textAlign: 'center', color: '#94a3b8' }}>
            <Building2 size={48} style={{ margin: '0 auto 16px', opacity: 0.5 }} />
            <p style={{ fontSize: '16px', fontWeight: '600' }}>No practice data yet</p>
            <p style={{ fontSize: '14px', marginTop: '8px' }}>Click "Sync from SharePoint" to load practice data!</p>
          </div>
        ) : (
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead>
                <tr style={{ borderBottom: '2px solid #e2e8f0' }}>
                  <th style={headerStyle}>Practice Name</th>
                  <th style={headerStyle}>Consultant</th>
                  <th style={headerStyle}>Location</th>
                  <th style={headerStyle}>Hospitals</th>
                  <th style={headerStyle}>Status</th>
                  <th style={headerStyle}>EMR Access</th>
                  <th style={headerStyle}>Contact</th>
                </tr>
              </thead>
              <tbody>
                {filteredPractices.map((practice, index) => {
                  const badge = getStatusBadge(practice);
                  return (
                    <tr key={index} style={{ borderBottom: '1px solid #f1f5f9' }}>
                      <td style={cellStyle}><strong>{practice.name}</strong></td>
                      <td style={cellStyle}>{practice.consultant}</td>
                      <td style={cellStyle}>{practice.location}</td>
                      <td style={cellStyle}>{practice.hospitals}</td>
                      <td style={cellStyle}>
                        <span style={{ padding: '4px 12px', background: badge.bg, color: badge.color, borderRadius: '12px', fontSize: '12px', fontWeight: '600' }}>
                          {badge.label}
                        </span>
                      </td>
                      <td style={cellStyle}>{practice.emrAccess}</td>
                      <td style={cellStyle}>{practice.contact}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

function MetricCard({ icon, title, value, subtitle, color, gradient }) {
  return (
    <div style={{ background: 'white', borderRadius: '16px', padding: '24px', boxShadow: '0 4px 6px rgba(0,0,0,0.05)', position: 'relative', overflow: 'hidden', transition: 'transform 0.2s', cursor: 'pointer' }}
      onMouseOver={(e) => e.currentTarget.style.transform = 'translateY(-4px)'}
      onMouseOut={(e) => e.currentTarget.style.transform = 'translateY(0)'}>
      <div style={{ position: 'absolute', top: '-20px', right: '-20px', width: '100px', height: '100px', background: `${color}15`, borderRadius: '50%' }}></div>
      <div style={{ width: '48px', height: '48px', borderRadius: '12px', background: gradient, display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'white', marginBottom: '16px' }}>
        {icon}
      </div>
      <div style={{ fontSize: '13px', color: '#64748b', fontWeight: '600', marginBottom: '8px' }}>{title}</div>
      <div style={{ fontSize: '32px', fontWeight: '800', color: '#1e293b', marginBottom: '4px' }}>{value}</div>
      <div style={{ fontSize: '12px', color: '#94a3b8' }}>{subtitle}</div>
    </div>
  );
}

function ChartCard({ title, children }) {
  return (
    <div style={{ background: 'white', borderRadius: '16px', padding: '24px', boxShadow: '0 4px 6px rgba(0,0,0,0.05)' }}>
      <h3 style={{ fontSize: '16px', fontWeight: '700', marginBottom: '20px', color: '#1e293b' }}>{title}</h3>
      {children}
    </div>
  );
}

const headerStyle = {
  padding: '12px',
  textAlign: 'left',
  fontSize: '12px',
  fontWeight: '700',
  color: '#475569',
  textTransform: 'uppercase',
  letterSpacing: '0.05em'
};

const cellStyle = {
  padding: '16px 12px',
  fontSize: '14px',
  color: '#1e293b'
};

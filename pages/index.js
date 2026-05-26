import React, { useState, useEffect } from 'react';
import { RefreshCw, Users, CheckCircle, Clock, Building2, TrendingUp, Hospital, Calendar } from 'lucide-react';
import { LineChart, Line, BarChart, Bar, PieChart, Pie, Cell, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';

export default function NurseNavigatorDashboard() {
  const [activeTab, setActiveTab] = useState('patients');
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [lastSync, setLastSync] = useState(null);

  useEffect(() => {
    fetchData();
    const interval = setInterval(fetchData, 120000);
    return () => clearInterval(interval);
  }, []);

  const fetchData = async () => {
    try {
      const response = await fetch('/api/nurse-navigator-data');
      if (response.ok) {
        const jsonData = await response.json();
        setData(jsonData);
        setLastSync(new Date());
        setLoading(false);
      }
    } catch (error) {
      console.error('Error:', error);
      setLoading(false);
    }
  };

  if (loading) {
    return <LoadingScreen />;
  }

  return (
    <div style={{
      minHeight: '100vh',
      background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
      fontFamily: "'Inter', -apple-system, sans-serif",
      padding: '0'
    }}>
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
          <div>
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
          </div>
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
        <MetricCard
          icon={<Users size={24} />}
          title="Active Patients"
          value={metrics.totalPatients || 0}
          subtitle="Currently tracking"
          color="#3b82f6"
          gradient="linear-gradient(135deg, #3b82f6 0%, #2563eb 100%)"
        />
        <MetricCard
          icon={<Calendar size={24} />}
          title="TCM Scheduled"
          value={metrics.tcmScheduled || 0}
          subtitle={`${metrics.tcmSchedulingRate || 0}% scheduling rate`}
          color="#10b981"
          gradient="linear-gradient(135deg, #10b981 0%, #059669 100%)"
        />
        <MetricCard
          icon={<CheckCircle size={24} />}
          title="Visits Verified"
          value={metrics.visitVerified || 0}
          subtitle={`${metrics.verificationRate || 0}% success rate`}
          color="#8b5cf6"
          gradient="linear-gradient(135deg, #8b5cf6 0%, #7c3aed 100%)"
        />
        <MetricCard
          icon={<Clock size={24} />}
          title="Pending Action"
          value={(metrics.totalPatients - metrics.tcmScheduled) || 0}
          subtitle="Awaiting scheduling"
          color="#f59e0b"
          gradient="linear-gradient(135deg, #f59e0b 0%, #d97706 100%)"
        />
      </div>

      <div style={{
        display: 'grid',
        gridTemplateColumns: '1fr',
        gap: '24px',
        marginBottom: '24px'
      }}>
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

      <div style={{
        display: 'grid',
        gridTemplateColumns: 'repeat(auto-fit, minmax(350px, 1fr))',
        gap: '24px',
        marginBottom: '24px'
      }}>
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
              <Pie
                data={statusData}
                cx="50%"
                cy="50%"
                labelLine={false}
                label={({ name, percent }) => `${name}: ${(percent * 100).toFixed(0)}%`}
                outerRadius={100}
                fill="#8884d8"
                dataKey="value"
              >
                {statusData.map((entry, index) => (
                  <Cell key={`cell-${index}`} fill={entry.color} />
                ))}
              </Pie>
              <Tooltip />
            </PieChart>
          </ResponsiveContainer>
        </ChartCard>
      </div>

      <div style={{
        background: 'white',
        borderRadius: '16px',
        padding: '24px',
        boxShadow: '0 4px 6px rgba(0,0,0,0.05)'
      }}>
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
                    <span style={{
                      padding: '4px 12px',
                      background: patient.location === 'TVH' ? '#dbeafe' : '#fef3c7',
                      color: patient.location === 'TVH' ? '#1e40af' : '#92400e',
                      borderRadius: '12px',
                      fontSize: '12px',
                      fontWeight: '600'
                    }}>
                      {patient.location || 'N/A'}
                    </span>
                  </td>
                  <td style={cellStyle}>{patient.navigator || 'N/A'}</td>
                  <td style={cellStyle}>
                    <span style={{
                      padding: '4px 12px',
                      background: patient.tcmScheduled ? '#dcfce7' : '#fee2e2',
                      color: patient.tcmScheduled ? '#15803d' : '#991b1b',
                      borderRadius: '12px',
                      fontSize: '12px',
                      fontWeight: '600'
                    }}>
                      {patient.tcmScheduled ? 'Scheduled' : 'Pending'}
                    </span>
                  </td>
                  <td style={cellStyle}>
                    <span style={{ fontSize: '20px' }}>
                      {patient.visitVerified ? '✅' : '⏳'}
                    </span>
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
  const practiceMetrics = data?.practiceMetrics || {
    total: 34,
    enrolled: 15,
    pending: 11,
    declined: 8
  };

  const enrollmentData = [
    { name: 'Enrolled', value: practiceMetrics.enrolled, color: '#10b981' },
    { name: 'Pending', value: practiceMetrics.pending, color: '#f59e0b' },
    { name: 'Declined', value: practiceMetrics.declined, color: '#ef4444' }
  ];

  const consultantData = [
    { name: 'Rachel Robinson', enrolled: 7, pending: 4 },
    { name: 'Marlen Cornejo', enrolled: 2, pending: 4 },
    { name: 'Bella McKay', enrolled: 2, pending: 0 },
    { name: 'Trysh Logan', enrolled: 1, pending: 0 },
    { name: 'Chris Ford', enrolled: 1, pending: 0 },
    { name: 'Sheyenne Powers', enrolled: 1, pending: 0 },
    { name: 'Rachel Robinson/Menifee', enrolled: 1, pending: 0 }
  ];

  return (
    <div>
      <div style={{
        display: 'grid',
        gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))',
        gap: '24px',
        marginBottom: '32px'
      }}>
        <MetricCard
          icon={<Building2 size={24} />}
          title="Total Practices"
          value={practiceMetrics.total}
          subtitle="Contacted to date"
          color="#3b82f6"
          gradient="linear-gradient(135deg, #3b82f6 0%, #2563eb 100%)"
        />
        <MetricCard
          icon={<CheckCircle size={24} />}
          title="Enrolled"
          value={practiceMetrics.enrolled}
          subtitle={`${((practiceMetrics.enrolled / practiceMetrics.total) * 100).toFixed(1)}% success rate`}
          color="#10b981"
          gradient="linear-gradient(135deg, #10b981 0%, #059669 100%)"
        />
        <MetricCard
          icon={<Clock size={24} />}
          title="Pending"
          value={practiceMetrics.pending}
          subtitle="In enrollment process"
          color="#f59e0b"
          gradient="linear-gradient(135deg, #f59e0b 0%, #d97706 100%)"
        />
        <MetricCard
          icon={<TrendingUp size={24} />}
          title="Declined"
          value={practiceMetrics.declined}
          subtitle="Not participating"
          color="#ef4444"
          gradient="linear-gradient(135deg, #ef4444 0%, #dc2626 100%)"
        />
      </div>

      <div style={{
        display: 'grid',
        gridTemplateColumns: 'repeat(auto-fit, minmax(350px, 1fr))',
        gap: '24px',
        marginBottom: '24px'
      }}>
        <ChartCard title="Enrollment Status Breakdown">
          <ResponsiveContainer width="100%" height={300}>
            <PieChart>
              <Pie
                data={enrollmentData}
                cx="50%"
                cy="50%"
                labelLine={false}
                label={({ name, percent }) => `${name}: ${(percent * 100).toFixed(0)}%`}
                outerRadius={110}
                fill="#8884d8"
                dataKey="value"
              >
                {enrollmentData.map((entry, index) => (
                  <Cell key={`cell-${index}`} fill={entry.color} />
                ))}
              </Pie>
              <Tooltip />
            </PieChart>
          </ResponsiveContainer>
        </ChartCard>

        <ChartCard title="Consultant Performance">
          <ResponsiveContainer width="100%" height={300}>
            <BarChart 
              data={consultantData} 
              layout="vertical"
              margin={{ left: 180, right: 20 }}
            >
              <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
              <XAxis type="number" stroke="#64748b" style={{ fontSize: '12px' }} />
              <YAxis 
                dataKey="name" 
                type="category"
                stroke="#64748b" 
                width={170}
                style={{ fontSize: '11px' }} 
              />
              <Tooltip />
              <Legend />
              <Bar dataKey="enrolled" fill="#10b981" radius={[0, 8, 8, 0]} name="Enrolled" />
              <Bar dataKey="pending" fill="#f59e0b" radius={[0, 8, 8, 0]} name="Pending" />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>
      </div>
    </div>
  );
}

function LoadingScreen() {
  return (
    <div style={{
      minHeight: '100vh',
      background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center'
    }}>
      <div style={{ textAlign: 'center' }}>
        <RefreshCw size={48} color="white" style={{ animation: 'spin 2s linear infinite', margin: '0 auto 20px' }} />
        <h2 style={{ fontSize: '24px', fontWeight: '700', color: 'white' }}>Loading Dashboard...</h2>
      </div>
      <style jsx>{`
        @keyframes spin {
          from { transform: rotate(0deg); }
          to { transform: rotate(360deg); }
        }
      `}</style>
    </div>
  );
}

function MetricCard({ icon, title, value, subtitle, color, gradient }) {
  return (
    <div style={{
      background: 'white',
      borderRadius: '16px',
      padding: '24px',
      boxShadow: '0 4px 6px rgba(0,0,0,0.05)',
      position: 'relative',
      overflow: 'hidden',
      transition: 'transform 0.2s',
      cursor: 'pointer'
    }}
    onMouseOver={(e) => e.currentTarget.style.transform = 'translateY(-4px)'}
    onMouseOut={(e) => e.currentTarget.style.transform = 'translateY(0)'}>
      <div style={{
        position: 'absolute',
        top: '-20px',
        right: '-20px',
        width: '100px',
        height: '100px',
        background: `${color}15`,
        borderRadius: '50%'
      }}></div>
      <div style={{
        width: '48px',
        height: '48px',
        borderRadius: '12px',
        background: gradient,
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        color: 'white',
        marginBottom: '16px'
      }}>
        {icon}
      </div>
      <div style={{ fontSize: '13px', color: '#64748b', fontWeight: '600', marginBottom: '8px' }}>
        {title}
      </div>
      <div style={{ fontSize: '32px', fontWeight: '800', color: '#1e293b', marginBottom: '4px' }}>
        {value}
      </div>
      <div style={{ fontSize: '12px', color: '#94a3b8' }}>
        {subtitle}
      </div>
    </div>
  );
}

function ChartCard({ title, children }) {
  return (
    <div style={{
      background: 'white',
      borderRadius: '16px',
      padding: '24px',
      boxShadow: '0 4px 6px rgba(0,0,0,0.05)'
    }}>
      <h3 style={{ fontSize: '16px', fontWeight: '700', marginBottom: '20px', color: '#1e293b' }}>
        {title}
      </h3>
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

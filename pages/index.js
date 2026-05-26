import React, { useState, useEffect } from 'react';
import { RefreshCw, CheckCircle, AlertCircle, Clock, Zap } from 'lucide-react';

export default function AutoSyncDashboard() {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [lastSync, setLastSync] = useState(null);
  const [syncStatus, setSyncStatus] = useState('checking');

  useEffect(() => {
    fetchLatestData();
    const interval = setInterval(fetchLatestData, 120000); // Check every 2 minutes
    return () => clearInterval(interval);
  }, []);

  const fetchLatestData = async () => {
    try {
      setSyncStatus('checking');
      const response = await fetch('/api/nurse-navigator-data');
      
      if (response.ok) {
        const jsonData = await response.json();
        setData(jsonData);
        setLastSync(new Date());
        setSyncStatus('synced');
        setLoading(false);
      } else {
        throw new Error('Failed to fetch');
      }
    } catch (error) {
      console.error('Error:', error);
      setSyncStatus('error');
      setLoading(false);
    }
  };

  const manualSync = () => {
    setSyncStatus('syncing');
    fetchLatestData();
  };

  if (loading && !data) {
    return (
      <div style={{
        minHeight: '100vh',
        background: 'linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 50%, #f0fdf4 100%)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        fontFamily: "'Inter', sans-serif"
      }}>
        <div style={{ textAlign: 'center' }}>
          <RefreshCw size={48} color="#0891b2" style={{ animation: 'spin 2s linear infinite', margin: '0 auto 20px' }} />
          <h2 style={{ fontSize: '24px', fontWeight: '700', color: '#1e293b' }}>
            Loading Dashboard...
          </h2>
          <p style={{ fontSize: '14px', color: '#64748b' }}>
            Fetching latest data from Excel
          </p>
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

  return (
    <div style={{
      minHeight: '100vh',
      background: 'linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 50%, #f0fdf4 100%)',
      fontFamily: "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
      padding: '40px 24px'
    }}>
      <div style={{ maxWidth: '1200px', margin: '0 auto' }}>
        <header style={{ marginBottom: '32px' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
            <div>
              <div style={{
                display: 'inline-block',
                padding: '8px 16px',
                background: '#dcfce7',
                borderRadius: '20px',
                fontSize: '12px',
                fontWeight: '700',
                color: '#15803d',
                marginBottom: '12px'
              }}>
                ⚡ AUTO-SYNC ENABLED
              </div>
              <h1 style={{
                fontSize: '36px',
                fontWeight: '800',
                margin: '0 0 12px 0',
                background: 'linear-gradient(135deg, #0891b2 0%, #06b6d4 50%, #10b981 100%)',
                WebkitBackgroundClip: 'text',
                WebkitTextFillColor: 'transparent'
              }}>
                Nurse Navigator Dashboard
              </h1>
              <p style={{ fontSize: '16px', color: '#64748b', margin: 0 }}>
                Auto-syncs from Excel in OneDrive
              </p>
            </div>
            
            <div style={{
              background: 'white',
              borderRadius: '12px',
              padding: '16px 24px',
              border: '2px solid #e2e8f0'
            }}>
              <SyncStatusBadge status={syncStatus} />
              {lastSync && (
                <div style={{ fontSize: '13px', color: '#64748b', marginTop: '8px' }}>
                  Last sync: {lastSync.toLocaleTimeString()}
                </div>
              )}
              <button
                onClick={manualSync}
                disabled={syncStatus === 'syncing'}
                style={{
                  marginTop: '8px',
                  padding: '8px 16px',
                  background: syncStatus === 'syncing' ? '#f1f5f9' : '#0891b2',
                  color: syncStatus === 'syncing' ? '#94a3b8' : 'white',
                  border: 'none',
                  borderRadius: '8px',
                  fontSize: '13px',
                  fontWeight: '600',
                  cursor: syncStatus === 'syncing' ? 'not-allowed' : 'pointer',
                  width: '100%'
                }}
              >
                {syncStatus === 'syncing' ? 'Syncing...' : 'Sync Now'}
              </button>
            </div>
          </div>

          <div style={{
            background: 'linear-gradient(135deg, #ecfdf5 0%, #d1fae5 100%)',
            border: '2px solid #86efac',
            borderRadius: '16px',
            padding: '20px',
            display: 'flex',
            gap: '16px',
            alignItems: 'center'
          }}>
            <Zap size={32} color="#15803d" />
            <div>
              <h3 style={{ fontSize: '16px', fontWeight: '700', color: '#15803d', margin: '0 0 6px 0' }}>
                Automatic Sync Active
              </h3>
              <p style={{ fontSize: '14px', color: '#166534', margin: 0 }}>
                Nurses edit Excel → Power Automate detects changes → Dashboard updates automatically
              </p>
            </div>
          </div>
        </header>

        {data && data.patients && data.patients.length > 0 ? (
          <>
            <div style={{
              display: 'grid',
              gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))',
              gap: '20px',
              marginBottom: '32px'
            }}>
              <MetricCard title="Active Patients" value={data.metrics?.totalPatients || data.patients.length} color="#10b981" />
              <MetricCard title="TCM Scheduled" value={data.metrics?.tcmScheduled || 0} color="#3b82f6" />
              <MetricCard title="Visits Verified" value={data.metrics?.visitVerified || 0} color="#0891b2" />
            </div>

            <div style={{
              background: 'white',
              borderRadius: '16px',
              padding: '24px',
              boxShadow: '0 4px 6px -1px rgba(0, 0, 0, 0.05)',
              border: '1px solid #e2e8f0'
            }}>
              <h2 style={{ fontSize: '20px', fontWeight: '700', marginBottom: '20px', color: '#1e293b' }}>
                Patient List ({data.patients.length} total)
              </h2>
              <div style={{ overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                  <thead>
                    <tr style={{ borderBottom: '2px solid #e2e8f0' }}>
                      <th style={headerStyle}>Patient Name</th>
                      <th style={headerStyle}>Practice</th>
                      <th style={headerStyle}>Location</th>
                      <th style={headerStyle}>Navigator</th>
                    </tr>
                  </thead>
                  <tbody>
                    {data.patients.map((patient, index) => (
                      <tr key={index} style={{ borderBottom: '1px solid #f1f5f9' }}>
                        <td style={cellStyle}>{patient.name || 'N/A'}</td>
                        <td style={cellStyle}>{patient.practice || 'N/A'}</td>
                        <td style={cellStyle}>{patient.location || 'N/A'}</td>
                        <td style={cellStyle}>{patient.navigator || 'N/A'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        ) : (
          <div style={{
            background: 'white',
            borderRadius: '16px',
            padding: '48px',
            textAlign: 'center',
            border: '2px dashed #e2e8f0'
          }}>
            <AlertCircle size={48} color="#f59e0b" style={{ margin: '0 auto 16px' }} />
            <h3 style={{ fontSize: '20px', fontWeight: '700', color: '#1e293b', marginBottom: '8px' }}>
              No Data Yet
            </h3>
            <p style={{ fontSize: '14px', color: '#64748b', marginBottom: '16px' }}>
              Waiting for Excel data from Power Automate...
            </p>
            <button
              onClick={manualSync}
              style={{
                padding: '12px 24px',
                background: '#0891b2',
                color: 'white',
                border: 'none',
                borderRadius: '10px',
                fontSize: '14px',
                fontWeight: '600',
                cursor: 'pointer'
              }}
            >
              Check for Data Now
            </button>
          </div>
        )}
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

function SyncStatusBadge({ status }) {
  const configs = {
    checking: { icon: <Clock size={16} />, text: 'Checking...', color: '#94a3b8', bg: '#f1f5f9' },
    synced: { icon: <CheckCircle size={16} />, text: 'Up to date', color: '#15803d', bg: '#dcfce7' },
    syncing: { icon: <RefreshCw size={16} />, text: 'Syncing', color: '#0891b2', bg: '#cffafe' },
    error: { icon: <AlertCircle size={16} />, text: 'Error', color: '#dc2626', bg: '#fee2e2' }
  };
  const config = configs[status] || configs.checking;

  return (
    <div style={{
      display: 'flex',
      alignItems: 'center',
      gap: '8px',
      padding: '8px 12px',
      background: config.bg,
      color: config.color,
      borderRadius: '8px',
      fontSize: '13px',
      fontWeight: '600'
    }}>
      {config.icon}
      {config.text}
    </div>
  );
}

function MetricCard({ title, value, color }) {
  return (
    <div style={{
      background: 'white',
      borderRadius: '12px',
      padding: '20px',
      border: `2px solid ${color}20`,
      boxShadow: '0 2px 4px rgba(0,0,0,0.05)'
    }}>
      <div style={{ fontSize: '13px', color: '#64748b', fontWeight: '600', marginBottom: '8px' }}>
        {title}
      </div>
      <div style={{ fontSize: '32px', fontWeight: '800', color: color }}>
        {value}
      </div>
    </div>
  );
}

const headerStyle = {
  padding: '12px',
  textAlign: 'left',
  fontSize: '13px',
  fontWeight: '700',
  color: '#475569',
  textTransform: 'uppercase'
};

const cellStyle = {
  padding: '12px',
  fontSize: '14px',
  color: '#1e293b'
};

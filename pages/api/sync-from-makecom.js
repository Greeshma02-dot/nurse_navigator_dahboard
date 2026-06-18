// FILE: pages/api/sync-from-makecom.js
// This endpoint receives Excel files from Make.com and updates the dashboard

import formidable from 'formidable';
import fs from 'fs';
import XLSX from 'xlsx';

export const config = {
  api: {
    bodyParser: false, // Required for file uploads
  },
};

// Store data in memory (persists between requests)
let cachedData = null;
let lastUpdated = null;

export default async function handler(req, res) {
  // Allow CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  // GET - return cached data to dashboard
  if (req.method === 'GET') {
    if (cachedData) {
      return res.status(200).json({
        ...cachedData,
        lastUpdated,
        source: 'make.com auto-sync'
      });
    }
    return res.status(200).json({ 
      patients: [], 
      metrics: { totalPatients: 0, tcmScheduled: 0, notYetScheduled: 0, visitVerified: 0, missed14DayWindow: 0, scheduledRate: 0, nurseCounts: {} },
      practiceMetrics: { total: 0, enrolled: 0, pending: 0, declined: 0, tbd: 0, emrComplete: 0 },
      practices: [],
      lastUpdated: null 
    });
  }

  // POST - receive files from Make.com
  if (req.method === 'POST') {
    try {
      const form = formidable({ multiples: true, maxFileSize: 50 * 1024 * 1024 });
      
      const [fields, files] = await new Promise((resolve, reject) => {
        form.parse(req, (err, fields, files) => {
          if (err) reject(err);
          else resolve([fields, files]);
        });
      });

      let patientData = null;
      let practiceData = null;

      // Process each uploaded file
      const fileList = Array.isArray(files.file) ? files.file : 
                       files.file ? [files.file] :
                       Object.values(files).flat();

      for (const file of fileList) {
        const buffer = fs.readFileSync(file.filepath || file.path);
        const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const fileName = (file.originalFilename || file.name || '').toLowerCase();

        const isPractice = fileName.includes('ccpaco') || 
                           (fileName.includes('tracking') && !fileName.includes('patient'));
        
        if (isPractice) {
          // Practice file: find header row then read data by column index
          const rawSheet = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
          const headerIdx = rawSheet.findIndex(row =>
            String(row[0] || '').toLowerCase().includes('practice participants') ||
            String(row[0] || '').toLowerCase().includes('practice')
          );
          const dataStart = rawSheet.findIndex((row, idx) => {
            if (idx <= headerIdx) return false;
            const first = String(row[0] || '').trim();
            return first && !first.toLowerCase().includes('practice participants') && !first.toLowerCase().startsWith('column');
          });

          const rows = rawSheet
            .slice(dataStart >= 0 ? dataStart : headerIdx + 1)
            .filter(row => row[0] && String(row[0]).trim())
            .map(row => ({
              name: String(row[0] || '').trim(),
              consultant: String(row[1] || '').trim(),
              location: String(row[2] || '').trim(),
              hospitals: String(row[3] || '').trim(),
              pdvStatus: String(row[4] || '').trim(),
              emrGranted: String(row[5] || '').trim(),
              emrAccess: String(row[5] || '').trim(),
              login: String(row[6] || '').trim(),
              contact: String(row[7] || '').trim(),
              networkAccess: String(row[8] || '').trim(),
              notes: String(row[9] || '').trim(),
            }));

          const isPDVComplete = v => String(v || '').toLowerCase().startsWith('complete');
          const isDeclined = v => String(v || '').toLowerCase().includes('declined');
          const isTBD = v => String(v || '').toLowerCase() === 'tbd';

          const enrolled = rows.filter(p => isPDVComplete(p.pdvStatus)).length;
          const declined = rows.filter(p => isDeclined(p.pdvStatus)).length;
          const tbd = rows.filter(p => isTBD(p.pdvStatus)).length;
          const pending = rows.length - enrolled - declined - tbd;
          const emrComplete = rows.filter(p => isPDVComplete(p.emrGranted)).length;

          practiceData = {
            practiceMetrics: { total: rows.length, enrolled, pending, declined, tbd, emrComplete },
            practices: rows
          };

        } else {
          // Patient file
          const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });
          let bestIdx = 0, bestScore = -1;
          raw.slice(0, 10).forEach((row, idx) => {
            const joined = row.map(c => String(c || '').toLowerCase()).join(' ');
            let score = 0;
            if (joined.includes('patient')) score += 4;
            if (joined.includes('navigator')) score += 3;
            if (joined.includes('tcm')) score += 4;
            if (joined.includes('discharge')) score += 3;
            score += row.filter(c => String(c || '').trim()).length;
            if (score > bestScore) { bestScore = score; bestIdx = idx; }
          });

          const headers = (raw[bestIdx] || []).slice(0, 25).map(h => String(h || '').trim());
          const dataRows = raw.slice(bestIdx + 1)
            .map(row => {
              const obj = {};
              headers.forEach((h, i) => { if (h) obj[h] = row[i] ?? ''; });
              return obj;
            })
            .filter(row => Object.values(row).some(v => String(v || '').trim()));

          const getCell = (row, names) => {
            const keys = Object.keys(row || {});
            for (const name of names) {
              const found = keys.find(k => k.toLowerCase() === name.toLowerCase());
              if (found !== undefined) return row[found];
            }
            for (const name of names) {
              const found = keys.find(k => k.toLowerCase().includes(name.toLowerCase()));
              if (found !== undefined) return row[found];
            }
            return '';
          };

          const isYes = v => ['yes','y','true','1','scheduled','complete','completed','verified','done','x','✓'].includes(String(v || '').toLowerCase().trim());
          const parseDate = v => {
            if (!v || v === '' || v === 'N/A') return null;
            if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
            if (typeof v === 'number' && v > 1 && v < 200000) return new Date(Math.round((v - 25569) * 86400 * 1000));
            if (typeof v === 'string') { const d = new Date(v); return isNaN(d.getTime()) ? null : d; }
            return null;
          };
          const fmt = d => d ? d.toLocaleDateString('en-US') : 'N/A';

          const today = new Date();
          const patients = dataRows.map(p => {
            const actualDischarge = parseDate(getCell(p, ['Actual Discharge Date','Actual Discharge','Discharge Date']));
            const anticipated = parseDate(getCell(p, ['Anticipated Discharge','Anticipated']));
            const tcmRaw = String(getCell(p, ['TCM Appt Scheduled?','TCM Appt Scheduled','TCM Scheduled','TCM']) || '');
            const tcmScheduled = isYes(tcmRaw);
            const tcmPending = tcmRaw.toLowerCase() === 'pending';
            const visitVerified = isYes(getCell(p, ['Visit Verified','Verified']));
            const missed = isYes(getCell(p, ['14-Day Window Status','Missed 14-Day Window','Missed Window']));
            const nav = String(getCell(p, ['Navigator Assigned','Nurse Navigator','Navigator']) || '');
            const daysSince = actualDischarge ? Math.round((today - actualDischarge) / 86400000) : null;
            let status = 'Unknown';
            if (visitVerified) status = 'Visit Verified';
            else if (tcmScheduled) status = 'Scheduled';
            else if (tcmPending) status = 'Pending';
            else if (actualDischarge) status = 'Discharged - No TCM';
            else status = 'Admitted';
            return {
              name: String(getCell(p, ['Patient Name','Name']) || ''),
              practice: String(getCell(p, ['Practice']) || ''),
              location: String(getCell(p, ['Location','Site']) || ''),
              room: String(getCell(p, ['Room #','Room']) || ''),
              navigator: nav || 'N/A',
              anticipatedStr: fmt(anticipated),
              actualDischargeStr: fmt(actualDischarge),
              daysSinceDischarge: daysSince,
              tcmRaw, tcmScheduled, tcmPending,
              tcmDateStr: fmt(parseDate(getCell(p, ['TCM Appt Date']))),
              apptType: String(getCell(p, ['Appt Type']) || ''),
              visitVerified, missed14Day: missed,
              windowStatus: String(getCell(p, ['14-Day Window Status','14-Day Window']) || ''),
              call2: String(getCell(p, ['2-Day Call Attempt','2-Day Call']) || ''),
              call7: String(getCell(p, ['7-Day Call Attempt','7-Day Call']) || ''),
              status,
              notes: String(getCell(p, ['Notes']) || ''),
            };
          }).filter(p => p.name);

          const tcmScheduled = patients.filter(p => p.tcmScheduled).length;
          const visitVerified = patients.filter(p => p.visitVerified).length;
          const missed = patients.filter(p => p.missed14Day).length;
          const notYet = patients.filter(p => !p.tcmScheduled && !p.tcmPending && !p.visitVerified).length;
          const nurseCounts = {};
          patients.forEach(p => { if (p.navigator && p.navigator !== 'N/A') nurseCounts[p.navigator] = (nurseCounts[p.navigator] || 0) + 1; });

          patientData = {
            patients,
            metrics: {
              totalPatients: patients.length,
              currentlyAdmitted: patients.filter(p => !p.actualDischargeStr || p.actualDischargeStr === 'N/A').length,
              tcmScheduled, visitVerified, missed14DayWindow: missed,
              notYetScheduled: notYet, tcmPending: patients.filter(p => p.tcmPending).length,
              scheduledRate: patients.length > 0 ? Number(((tcmScheduled / patients.length) * 100).toFixed(1)) : 0,
              nurseCounts
            }
          };
        }
      }

      // Merge and cache
      cachedData = {
        ...(patientData || {}),
        ...(practiceData || {}),
      };
      lastUpdated = new Date().toISOString();

      return res.status(200).json({ 
        success: true, 
        message: `Synced successfully`,
        patients: cachedData.patients?.length || 0,
        practices: cachedData.practices?.length || 0,
        lastUpdated 
      });

    } catch (error) {
      console.error('Sync error:', error);
      return res.status(500).json({ error: error.message });
    }
  }

  return res.status(405).json({ error: 'Method not allowed' });
}

// ✅ SHEETS.JS - All functions exported to window
(function() {
  'use strict';

  // Google Sheets URLs
  const SHEETS = {
    evaluasi1: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vT6w6f01zKeCm8Xzx2nFGk1qGXVQbeOTHK8G6MoJLrjhM-XfGjgYE-Vq2eKMtOh6VboifRXZvSrW0R_/pub?output=csv',
    evaluasi2: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgocuvj3VAPt3n7xYgzHxWcECqxmu26KBO-rJQqAg6XlI_AKutOuZlP3PNznjeRyrBbtgzkAksL8mn/pub?gid=916136123&single=true&output=csv' // ⚠️ GANTI DENGAN LINK SHEET EVALUASI 2
  };

  // Parse CSV
  const parseCSV = function(csvText) {
    const lines = csvText.trim().split(/\r?\n/);
    if (lines.length < 2) return [];
    const rows = lines.slice(1).filter(line => line.trim() !== '');
    return rows.map(line => {
      const fields = [];
      let inQuotes = false;
      let current = '';
      for (let i = 0; i < line.length; i++) {
        const char = line[i];
        if (char === '"') {
          if (inQuotes && line[i + 1] === '"') {
            current += '"';
            i++;
          } else {
            inQuotes = !inQuotes;
          }
        } else if (char === ',' && !inQuotes) {
          fields.push(current.trim());
          current = '';
        } else {
          current += char;
        }
      }
      fields.push(current.trim());
      return {
        timestamp: fields[0] || '',
        nama: fields[2] || '',
        departemen: fields[3] || ''
      };
    });
  };

  // Fetch Google Sheet
  const fetchSheet = async function(type) {
    if (!SHEETS[type]) {
      console.error(`❌ Unknown sheet type: ${type}`);
      console.error('Available types:', Object.keys(SHEETS).join(', '));
      throw new Error(`Unknown sheet type: ${type}. Available: ${Object.keys(SHEETS).join(', ')}`);
    }
    try {
      console.log(`📊 Fetching sheet: ${type}...`);
      const res = await fetch(SHEETS[type], { method: 'GET', headers: { 'Accept': 'text/csv' } });
      console.log(`📊 Response status: ${res.status}`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const csv = await res.text();
      console.log(`📊 CSV fetched: ${csv.length} chars`);
      const data = parseCSV(csv);
      console.log(`📊 Parsed: ${data.length} rows`);
      const filtered = data.filter(d => d.nama && d.departemen);
      console.log(`📊 Filtered: ${filtered.length} complete rows`);
      return filtered;
    } catch (error) {
      console.error('❌ Error in fetchSheet:', error);
      throw error;
    }
  };

  // Get unique names
  const uniqueByName = function(data) {
    const seen = new Map();
    const result = [];
    for (const item of data) {
      const key = item.nama.trim().toLowerCase();
      if (!seen.has(key)) {
        seen.set(key, true);
        result.push(item);
      }
    }
    return result;
  };

  // Group by department
  const groupByDepartmentFiltered = function(data) {
    const uniqueData = uniqueByName(data);
    const map = {};
    for (const item of uniqueData) {
      const dept = item.departemen;
      if (!map[dept]) map[dept] = [];
      map[dept].push(item);
    }
    return map;
  };

  // Get department summary
  const getDepartmentSummary = function(data) {
    console.log('🔍 getDepartmentSummary called with', data.length, 'rows');
    if (!data || data.length === 0) return [];
    const grouped = groupByDepartmentFiltered(data);
    return Object.entries(grouped).map(([departemen, items]) => ({
      departemen,
      jumlah: items.length
    }));
  };

  // Find duplicate names
  const findDuplicateNames = function(data) {
    const count = {};
    const normalized = data.map(d => {
      const key = d.nama.trim().toLowerCase();
      count[key] = (count[key] || 0) + 1;
      return { ...d, _key: key };
    });
    return normalized.filter(d => count[d._key] > 1);
  };

  // Export to CSV
  const exportToCSV = function(filename, rows) {
    if (!rows || rows.length === 0) {
      console.warn('⚠️ No data to export');
      return;
    }
    const headers = Object.keys(rows[0]).join(',');
    const escapedRows = rows.map(row => Object.values(row).map(val => {
      if (val == null) return '""';
      let str = String(val).replace(/"/g, '""');
      return `"${str}"`;
    }).join(','));
    const csv = [headers, ...escapedRows].join('\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    console.log(`✅ Exported ${rows.length} rows to ${filename}`);
  };

  // Export department detail
  const exportDepartemenDetail = async function(type) {
    if (!SHEETS[type]) throw new Error(`Unknown sheet type: ${type}`);
    const data = await fetchSheet(type);
    const grouped = groupByDepartmentFiltered(data);
    const rows = [];
    for (const [dept, items] of Object.entries(grouped)) {
      for (const item of items) {
        rows.push({ departemen: dept, nama: item.nama, timestamp: item.timestamp });
      }
    }
    const filename = `detail_departemen_unik_${type}.csv`;
    exportToCSV(filename, rows);
  };

  // Export double input
  const exportDoubleInput = async function(type) {
    if (!SHEETS[type]) throw new Error(`Unknown sheet type: ${type}`);
    const data = await fetchSheet(type);
    const doubles = findDuplicateNames(data);
    const clean = doubles.map(({ _key, ...rest }) => rest);
    const filename = `double_input_${type}.csv`;
    exportToCSV(filename, clean);
  };

  // Export to window - this is critical!
  window.parseCSV = parseCSV;
  window.fetchSheet = fetchSheet;
  window.uniqueByName = uniqueByName;
  window.groupByDepartmentFiltered = groupByDepartmentFiltered;
  window.getDepartmentSummary = getDepartmentSummary;
  window.findDuplicateNames = findDuplicateNames;
  window.exportToCSV = exportToCSV;
  window.exportDepartemenDetail = exportDepartemenDetail;
  window.exportDoubleInput = exportDoubleInput;

  console.log('%c✅ sheets.js loaded - All functions ready!', 'color: green; font-weight: bold; font-size: 13px');
  console.log('  ✓ Available sheets:', Object.keys(SHEETS).join(', '));
  console.log('  ✓ fetchSheet:', typeof window.fetchSheet);
  console.log('  ✓ uniqueByName:', typeof window.uniqueByName);
  console.log('  ✓ getDepartmentSummary:', typeof window.getDepartmentSummary);

})();
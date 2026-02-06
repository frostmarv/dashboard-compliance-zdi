// assets/js/sheets.js
console.log('✅ sheets.js loaded successfully');

const SHEETS = {
  evaluasi1:
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vT6w6f01zKeCm8Xzx2nFGk1qGXVQbeOTHK8G6MoJLrjhM-XfGjgYE-Vq2eKMtOh6VboifRXZvSrW0R_/pub?gid=1248578848&single=true&output=csv'
};

/**
 * Parse CSV baris demi baris secara aman (handle quoted fields)
 */
function parseCSV(csvText) {
  const lines = csvText.trim().split(/\r?\n/);
  if (lines.length < 2) return [];

  // Ambil header (opsional, kita skip karena format tetap)
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
}

/**
 * Fetch Google Sheet (CSV publish)
 */
async function fetchSheet(type) {
  if (!SHEETS[type]) {
    throw new Error(`Unknown sheet type: ${type}`);
  }

  try {
    console.log(`📊 Fetching sheet: ${type}...`);
    const res = await fetch(SHEETS[type], {
      method: 'GET',
      headers: {
        'Accept': 'text/csv'
      }
    });
    console.log(`📊 Response status: ${res.status}`);
    
    if (!res.ok) {
      throw new Error(`HTTP ${res.status}: Gagal mengambil data`);
    }
    const csv = await res.text();
    console.log(`📊 CSV fetched, length: ${csv.length}`);
    
    const data = parseCSV(csv);
    console.log(`📊 Parsed ${data.length} rows`);

    // Filter hanya data lengkap
    const filtered = data.filter(d => d.nama && d.departemen);
    console.log(`📊 Filtered to ${filtered.length} complete rows`);
    
    return filtered;
  } catch (error) {
    console.error('❌ Error in fetchSheet:', error);
    throw error;
  }
}

/**
 * Ambil entri pertama untuk setiap nama (case-insensitive, trim)
 */
function uniqueByName(data) {
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
}

/**
 * Group by departemen (setelah deduplikasi nama)
 */
function groupByDepartmentFiltered(data) {
  const uniqueData = uniqueByName(data);
  const map = {};

  for (const item of uniqueData) {
    const dept = item.departemen;
    if (!map[dept]) map[dept] = [];
    map[dept].push(item);
  }

  return map;
}

/**
 * Rekap jumlah unik per departemen
 */
function getDepartmentSummary(data) {
  const grouped = groupByDepartmentFiltered(data);
  return Object.entries(grouped).map(([departemen, items]) => ({
    departemen,
    jumlah: items.length
  }));
}

/**
 * Cari semua entri dengan nama ganda (untuk audit)
 */
function findDuplicateNames(data) {
  const count = {};
  const normalized = data.map(d => {
    const key = d.nama.trim().toLowerCase();
    count[key] = (count[key] || 0) + 1;
    return { ...d, _key: key };
  });

  return normalized.filter(d => count[d._key] > 1);
}

/**
 * Export ke CSV
 */
function exportToCSV(filename, rows) {
  if (!rows || rows.length === 0) return;

  const headers = Object.keys(rows[0]).join(',');
  const escapedRows = rows.map(row =>
    Object.values(row)
      .map(val => {
        if (val == null) return '""';
        let str = String(val).replace(/"/g, '""');
        return `"${str}"`;
      })
      .join(',')
  );

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
}

/**
 * Export detail per departemen (nama unik saja)
 */
async function exportDepartemenDetail(type) {
  const data = await fetchSheet(type);
  const grouped = groupByDepartmentFiltered(data);
  const rows = [];

  for (const [dept, items] of Object.entries(grouped)) {
    for (const item of items) {
      rows.push({
        departemen: dept,
        nama: item.nama,
        timestamp: item.timestamp
      });
    }
  }

  exportToCSV('detail_departemen_unik.csv', rows);
}

/**
 * Export semua data double input
 */
async function exportDoubleInput(type) {
  const data = await fetchSheet(type);
  const doubles = findDuplicateNames(data);
  // Hapus properti internal `_key`
  const clean = doubles.map(({ _key, ...rest }) => rest);
  exportToCSV('double_input.csv', clean);
}

// ✅ Verify all functions are exported
console.log('✅ Functions available:');
console.log('  - parseCSV:', typeof parseCSV);
console.log('  - fetchSheet:', typeof fetchSheet);
console.log('  - uniqueByName:', typeof uniqueByName);
console.log('  - groupByDepartmentFiltered:', typeof groupByDepartmentFiltered);
console.log('  - getDepartmentSummary:', typeof getDepartmentSummary);
console.log('  - findDuplicateNames:', typeof findDuplicateNames);
console.log('  - exportToCSV:', typeof exportToCSV);
console.log('  - exportDepartemenDetail:', typeof exportDepartemenDetail);
console.log('  - exportDoubleInput:', typeof exportDoubleInput);
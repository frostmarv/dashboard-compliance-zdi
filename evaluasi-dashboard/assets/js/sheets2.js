// ✅ SHEETS2.JS - Connector Khusus Evaluasi 2 (FINAL - FIXED)
(function() {
  'use strict';

  // ⚠️ URL Web App untuk Evaluasi 2 - TRIM SPASI!
  const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbz9sqkxCI1MLGyw57HBrn7Adr0ZcWMgSAUSiycWyifD7YG5jjwmSwrWDkAejSURlQRS/exec'.trim();

  // Helper: Fetch dan unwrap response {status, data}
  const fetchUnwrapped = async (action) => {
    try {
      const res = await fetch(`${SCRIPT_URL}?action=${action}`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const result = await res.json();
      if (result && result.status === 'success' && Array.isArray(result.data)) {
        return result.data;
      }
      return Array.isArray(result) ? result : (result.data || []);
    } catch (error) {
      console.error(`❌ Error fetching ${action}:`, error);
      throw error;
    }
  };

  const fetchAllEmployees = async function() {
    console.log('📊 [Evaluasi 2] Fetching all employees...');
    const data = await fetchUnwrapped('getAllEmployees');
    console.log(`📊 [Evaluasi 2] Employees fetched: ${Array.isArray(data) ? data.length : 0} rows`);
    return Array.isArray(data) ? data : [];
  };

  const fetchAllResponses = async function() {
    console.log('📊 [Evaluasi 2] Fetching all responses...');
    const data = await fetchUnwrapped('getAllResponses');
    console.log(`📊 [Evaluasi 2] Responses fetched: ${Array.isArray(data) ? data.length : 0} rows`);
    return Array.isArray(data) ? data : [];
  };

  const getDepartmentSummary = async function() {
    const employees = await fetchAllEmployees();
    const responses = await fetchAllResponses();
    const responseMap = {};
    responses.forEach(r => { if (r?.nik) responseMap[String(r.nik).trim()] = r; });
    
    const deptMap = {};
    employees.forEach(emp => {
      const dept = emp?.departemen;
      if (!dept) return;
      if (!deptMap[dept]) deptMap[dept] = { departemen: dept, total: 0, done: 0, pending: 0 };
      deptMap[dept].total++;
      const nikKey = emp?.nik ? String(emp.nik).trim() : '';
      if (responseMap[nikKey]) deptMap[dept].done++;
      else deptMap[dept].pending++;
    });
    return Object.values(deptMap).sort((a, b) => (a.departemen || '').localeCompare(b.departemen || ''));
  };

  const getDepartmentDetail = async function(deptFilter) {
    const employees = await fetchAllEmployees();
    const responses = await fetchAllResponses();
    const responseMap = {};
    responses.forEach(r => { if (r?.nik) responseMap[String(r.nik).trim()] = r; });
    
    let filteredEmployees = employees;
    if (deptFilter) filteredEmployees = employees.filter(emp => emp?.departemen === deptFilter);
    
    return filteredEmployees.map(emp => {
      const nikKey = emp?.nik ? String(emp.nik).trim() : '';
      const response = responseMap[nikKey];
      return {
        nik: emp?.nik || '-',
        nama: emp?.nama || '-',
        departemen: emp?.departemen || '-',
        status: response ? 'done' : 'pending',
        nilai: response?.nilai ?? '-',
        waktu: response?.timestamp ?? '-'  // ✅ Ambil dari timestamp
      };
    });
  };

  // ✅✅✅ EXPORT TO CSV - FIXED: Handle uppercase & lowercase keys + UTF-8 BOM
  const exportToCSV = function(filename, rows) {
    if (!rows || rows.length === 0) {
      console.warn('⚠️ [Export] No data to export');
      alert('Tidak ada data untuk diekspor!');
      return;
    }

    console.log(`📦 [Export] Exporting ${rows.length} rows...`);
    console.log('📦 [Export] Sample row:', rows[0]);

    // ✅ UTF-8 BOM untuk Excel
    const BOM = '\uFEFF';
    
    // ✅ Header sesuai permintaan
    const headers = ['NIK', 'Nama Karyawan', 'Departemen', 'Status', 'Waktu Input'];
    
    // ✅ Helper: Ambil nilai dari object (handle UPPERCASE & lowercase keys)
    const getVal = (row, ...keys) => {
      for (let key of keys) {
        if (row[key] !== undefined && row[key] !== null) return row[key];
      }
      return '-';
    };
    
    // ✅ Helper: Escape nilai untuk CSV
    const escapeCSV = (val) => {
      if (val == null || val === '' || val === '-') return '-';
      const str = String(val).trim();
      if (str.includes(',') || str.includes('\n') || str.includes('"')) {
        return `"${str.replace(/"/g, '""')}"`;
      }
      return str;
    };
    
    // ✅ Format rows
    const escapedRows = rows.map(row => {
      // ✅ Ambil data dengan flexible keys (handle NIK/nik, Nama/nama, dll)
      const nik = getVal(row, 'NIK', 'nik');
      const nama = getVal(row, 'Nama', 'nama');
      const departemen = getVal(row, 'Departemen', 'departemen');
      const statusRaw = getVal(row, 'Status', 'status');
      const waktu = getVal(row, 'Waktu', 'waktu', 'timestamp');
      
      // ✅ Format status dengan simbol
      const statusDisplay = (statusRaw === 'done' || statusRaw === 'Sudah') 
        ? '✓ Sudah' : '✗ Belum';
      
      // ✅ Format waktu
      const waktuDisplay = (waktu && waktu !== '-' && waktu !== 'undefined') ? waktu : '-';
      
      return [
        escapeCSV(nik),
        escapeCSV(nama),
        escapeCSV(departemen),
        escapeCSV(statusDisplay),
        escapeCSV(waktuDisplay)
      ].join(','); // ✅ Standard CSV: comma separator
    });
    
    // ✅ Gabungkan: BOM + header + rows
    const csvContent = BOM + [headers.join(','), ...escapedRows].join('\n');
    
    // ✅ Buat blob UTF-8
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    
    // ✅ Trigger download
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    
    console.log(`✅ [Export] Success: ${filename}`);
  };

  const exportDepartemenDetail = async function(dept) {
    const status = await getDepartmentDetail(dept);
    const filename = `detail_evaluasi2_${dept || 'semua'}_${new Date().toISOString().split('T')[0]}.csv`;
    exportToCSV(filename, status);
  };

  const exportPendingList = async function(dept) {
    const status = await getDepartmentDetail(dept);
    const pending = status.filter(s => s.status === 'pending');
    const filename = `belum_mengisi_evaluasi2_${dept || 'semua'}_${new Date().toISOString().split('T')[0]}.csv`;
    exportToCSV(filename, pending);
  };

  // ✅ Export functions ke window
  window.ev2_fetchAllEmployees = fetchAllEmployees;
  window.ev2_fetchAllResponses = fetchAllResponses;
  window.ev2_getDepartmentSummary = getDepartmentSummary;
  window.ev2_getDepartmentDetail = getDepartmentDetail;
  window.ev2_exportToCSV = exportToCSV;
  window.ev2_exportDepartemenDetail = exportDepartemenDetail;
  window.ev2_exportPendingList = exportPendingList;

  console.log('%c✅ sheets2.js loaded - Evaluasi 2 Connected!', 'color: #10b981; font-weight: bold; font-size: 13px');
  console.log('  ✓ URL:', SCRIPT_URL.substring(0, 50) + '...');

})();
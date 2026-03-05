// ✅ SHEETS2.JS - Connector Khusus Evaluasi 2 (FINAL)
(function() {
  'use strict';

  // ⚠️ URL Web App untuk Evaluasi 2 (Responses 2) - TRIM SPASI!
  const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbz9sqkxCI1MLGyw57HBrn7Adr0ZcWMgSAUSiycWyifD7YG5jjwmSwrWDkAejSURlQRS/exec'.trim();

  // Helper: Fetch dan unwrap response {status, data}
  const fetchUnwrapped = async (action) => {
    try {
      const res = await fetch(`${SCRIPT_URL}?action=${action}`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const result = await res.json();
      
      // ✅ Unwrap: ambil .data jika ada, atau kembalikan result langsung
      if (result && result.status === 'success' && Array.isArray(result.data)) {
        return result.data;
      }
      return result.data || result;
    } catch (error) {
      console.error(`❌ Error fetching ${action}:`, error);
      throw error;
    }
  };

  // Fetch All Employees
  const fetchAllEmployees = async function() {
    console.log('📊 [Evaluasi 2] Fetching all employees...');
    const data = await fetchUnwrapped('getAllEmployees');
    console.log(`📊 [Evaluasi 2] Employees fetched: ${Array.isArray(data) ? data.length : 0} rows`);
    return Array.isArray(data) ? data : [];
  };

  // Fetch All Responses
  const fetchAllResponses = async function() {
    console.log('📊 [Evaluasi 2] Fetching all responses...');
    const data = await fetchUnwrapped('getAllResponses');
    console.log(`📊 [Evaluasi 2] Responses fetched: ${Array.isArray(data) ? data.length : 0} rows`);
    return Array.isArray(data) ? data : [];
  };

  // Get Department Summary
  const getDepartmentSummary = async function() {
    const employees = await fetchAllEmployees();
    const responses = await fetchAllResponses();
    
    const responseMap = {};
    responses.forEach(r => { 
      if (r?.nik) responseMap[String(r.nik).trim()] = r; 
    });
    
    const deptMap = {};
    
    employees.forEach(emp => {
      const dept = emp?.departemen;
      if (!dept) return;
      
      if (!deptMap[dept]) {
        deptMap[dept] = { departemen: dept, total: 0, done: 0, pending: 0 };
      }
      deptMap[dept].total++;
      
      const nikKey = emp?.nik ? String(emp.nik).trim() : '';
      if (responseMap[nikKey]) {
        deptMap[dept].done++;
      } else {
        deptMap[dept].pending++;
      }
    });
    
    return Object.values(deptMap).sort((a, b) => 
      (a.departemen || '').localeCompare(b.departemen || '')
    );
  };

  // Get Department Detail
  const getDepartmentDetail = async function(deptFilter) {
    const employees = await fetchAllEmployees();
    const responses = await fetchAllResponses();
    
    console.log(`🔍 [Detail] Employees: ${employees.length}, Responses: ${responses.length}`);
    
    const responseMap = {};
    responses.forEach(r => { 
      if (r?.nik) responseMap[String(r.nik).trim()] = r; 
    });
    
    let filteredEmployees = employees;
    if (deptFilter) {
      filteredEmployees = employees.filter(emp => 
        emp?.departemen === deptFilter
      );
    }
    
    console.log(`🔍 [Detail] Filtered: ${filteredEmployees.length} employees`);
    
    return filteredEmployees.map(emp => {
      const nikKey = emp?.nik ? String(emp.nik).trim() : '';
      const response = responseMap[nikKey];
      
      return {
        nik: emp?.nik || '-',
        nama: emp?.nama || '-',
        departemen: emp?.departemen || '-',
        status: response ? 'done' : 'pending',
        nilai: response?.nilai ?? '-',
        // ✅ FIX: Ambil dari response.timestamp (bukan response.waktu)
        waktu: response?.timestamp ?? '-'
      };
    });
  };

  // ✅ Export to CSV - FINAL VERSION (UTF-8 BOM + Clean Format)
  const exportToCSV = function(filename, rows) {
    if (!rows || rows.length === 0) {
      console.warn('⚠️ [Export] No data to export');
      alert('Tidak ada data untuk diekspor!');
      return;
    }

    console.log(`📦 [Export] Exporting ${rows.length} rows...`);
    console.log('📦 [Export] Sample row:', rows[0]);

    // ✅ UTF-8 BOM - WAJIB agar Excel baca karakter ✓✗ dengan benar
    const BOM = '\uFEFF';
    
    // ✅ Header sesuai permintaan
    const headers = ['NIK', 'Nama Karyawan', 'Departemen', 'Status', 'Waktu Input'];
    
    // ✅ Helper: Escape nilai untuk CSV (handle koma & newline)
    const escapeCSV = (val) => {
      if (val == null || val === '') return '-';
      const str = String(val).trim();
      // Jika mengandung koma, newline, atau tanda kutip → wrap dengan kutip
      if (str.includes(',') || str.includes('\n') || str.includes('"')) {
        return `"${str.replace(/"/g, '""')}"`;
      }
      return str;
    };
    
    // ✅ Format rows
    const escapedRows = rows.map(row => {
      // Format status dengan simbol
      const statusDisplay = row.status === 'done' ? '✓ Sudah' : '✗ Belum';
      // Format waktu (handle undefined/null)
      const waktuDisplay = (row.waktu && row.waktu !== '-' && row.waktu !== 'undefined') 
        ? row.waktu 
        : '-';
      
      return [
        escapeCSV(row.nik),
        escapeCSV(row.nama),
        escapeCSV(row.departemen),
        escapeCSV(statusDisplay),  // ✓ Sudah / ✗ Belum
        escapeCSV(waktuDisplay)
      ].join(','); // ✅ Standard CSV: comma separator
    });
    
    // ✅ Gabungkan: BOM + header + rows
    const csvContent = BOM + [headers.join(','), ...escapedRows].join('\n');
    
    // ✅ Buat blob dengan encoding UTF-8
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

  // ✅ Export functions ke window dengan prefix ev2_
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
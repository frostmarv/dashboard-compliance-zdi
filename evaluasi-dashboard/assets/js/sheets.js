// ✅ SHEETS.JS - Connector to Google Apps Script
(function() {
  'use strict';

  // ⚠️ GANTI URL INI DENGAN URL WEB APP ANDA
  const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwubxaFb1MuDGLLxreh2O7LLNtP54fuihjaVmjUxDnnNP-osEWmMa5DhkUe4Gr_Yvn9/exec';

  // Fetch All Employees from Database
  const fetchAllEmployees = async function() {
    try {
      console.log('📊 Fetching all employees...');
      const res = await fetch(`${SCRIPT_URL}?action=getAllEmployees`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();
      console.log(`📊 Employees fetched: ${data.length} rows`);
      return data;
    } catch (error) {
      console.error('❌ Error in fetchAllEmployees:', error);
      throw error;
    }
  };

  // Fetch All Responses from Responses Sheet
  const fetchAllResponses = async function() {
    try {
      console.log('📊 Fetching all responses...');
      const res = await fetch(`${SCRIPT_URL}?action=getAllResponses`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();
      console.log(`📊 Responses fetched: ${data.length} rows`);
      return data;
    } catch (error) {
      console.error('❌ Error in fetchAllResponses:', error);
      throw error;
    }
  };

  // Get All Unique Departments from Database
  const getAllDepartments = async function() {
    const employees = await fetchAllEmployees();
    const deptSet = new Set();
    
    employees.forEach(emp => {
      if (emp.departemen) {
        deptSet.add(emp.departemen);
      }
    });
    
    return Array.from(deptSet).sort();
  };

  // Get Department Summary (Total Employees per Department from Database)
  const getDepartmentSummary = async function() {
    const employees = await fetchAllEmployees();
    const responses = await fetchAllResponses();
    
    // Map responses by NIK
    const responseMap = {};
    responses.forEach(r => {
      responseMap[r.nik] = r;
    });
    
    // Group by department
    const deptMap = {};
    
    employees.forEach(emp => {
      const dept = emp.departemen;
      if (!deptMap[dept]) {
        deptMap[dept] = {
          departemen: dept,
          total: 0,
          done: 0,
          pending: 0
        };
      }
      
      deptMap[dept].total++;
      
      if (responseMap[emp.nik]) {
        deptMap[dept].done++;
      } else {
        deptMap[dept].pending++;
      }
    });
    
    return Object.values(deptMap).sort((a, b) => a.departemen.localeCompare(b.departemen));
  };

  // Get Completion Status for Specific Department
  const getDepartmentDetail = async function(deptFilter) {
    const employees = await fetchAllEmployees();
    const responses = await fetchAllResponses();
    
    // Map responses by NIK
    const responseMap = {};
    responses.forEach(r => {
      responseMap[r.nik] = r;
    });
    
    // Filter by department if specified
    let filteredEmployees = employees;
    if (deptFilter) {
      filteredEmployees = employees.filter(emp => emp.departemen === deptFilter);
    }
    
    // Merge data with status
    const status = filteredEmployees.map(emp => {
      const response = responseMap[emp.nik];
      return {
        nik: emp.nik,
        nama: emp.nama,
        departemen: emp.departemen,
        status: response ? 'done' : 'pending',
        nilai: response ? response.nilai : '-',
        waktu: response ? response.waktu : '-'
      };
    });
    
    return status;
  };

  // Get Completion Status (All Employees)
  const getCompletionStatus = async function() {
    return await getDepartmentDetail(null);
  };

  // Get Summary Stats
  const getSummaryStats = function(statusData) {
    const total = statusData.length;
    const done = statusData.filter(d => d.status === 'done').length;
    const pending = total - done;
    const percent = total > 0 ? Math.round((done / total) * 100) : 0;
    
    return { total, done, pending, percent };
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

  // Export Department Detail
  const exportDepartemenDetail = async function(dept) {
    const status = await getDepartmentDetail(dept);
    const filename = `detail_evaluasi1_${dept || 'semua'}_${new Date().toISOString().split('T')[0]}.csv`;
    exportToCSV(filename, status);
  };

  // Export Pending List
  const exportPendingList = async function(dept) {
    const status = await getDepartmentDetail(dept);
    const pending = status.filter(s => s.status === 'pending');
    const filename = `belum_mengisi_${dept || 'semua'}_${new Date().toISOString().split('T')[0]}.csv`;
    exportToCSV(filename, pending);
  };

  // Export to window
  window.fetchAllEmployees = fetchAllEmployees;
  window.fetchAllResponses = fetchAllResponses;
  window.getAllDepartments = getAllDepartments;
  window.getDepartmentSummary = getDepartmentSummary;
  window.getDepartmentDetail = getDepartmentDetail;
  window.getCompletionStatus = getCompletionStatus;
  window.getSummaryStats = getSummaryStats;
  window.exportToCSV = exportToCSV;
  window.exportDepartemenDetail = exportDepartemenDetail;
  window.exportPendingList = exportPendingList;

  console.log('%c✅ sheets.js loaded - Connected to Google Apps Script!', 'color: green; font-weight: bold; font-size: 13px');
  console.log('  ✓ SCRIPT_URL:', SCRIPT_URL.substring(0, 50) + '...');
  console.log('  ✓ getDepartmentSummary:', typeof window.getDepartmentSummary);
  console.log('  ✓ getDepartmentDetail:', typeof window.getDepartmentDetail);

})();
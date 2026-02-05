// assets/js/sheets.js

const SHEETS = {
  evaluasi1:
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vT6w6f01zKeCm8Xzx2nFGk1qGXVQbeOTHK8G6MoJLrjhM-XfGjgYE-Vq2eKMtOh6VboifRXZvSrW0R_/pub?gid=1248578848&single=true&output=csv'
};

/**
 * Fetch Google Sheet (CSV publish)
 * Mapping kolom:
 * A = timestamp
 * C = nama
 * D = departemen
 */
async function fetchSheet(type) {
  const res = await fetch(SHEETS[type]);
  const csv = await res.text();

  // split per baris, buang header
  const rows = csv.split('\n').slice(1);

  return rows
    .map(row => {
      // NOTE: CSV sederhana (aman selama field tidak ada koma)
      const c = row.split(',');

      return {
        timestamp: c[0]?.trim() || '', // 🔥 KOLOM A
        nama: c[2]?.trim() || '',      // 🔥 KOLOM C
        departemen: c[3]?.trim() || '' // 🔥 KOLOM D
      };
    })
    .filter(d => d.nama && d.departemen); // buang row kosong
}

/**
 * Group data by departemen
 */
function groupByDepartment(data) {
  const map = {};

  data.forEach(d => {
    if (!map[d.departemen]) {
      map[d.departemen] = [];
    }
    map[d.departemen].push(d);
  });

  return map;
}

/**
 * Deteksi double input berdasarkan nama
 * (opsional, buat page double-input.html)
 */
function findDuplicateNames(data) {
  const counter = {};
  const duplicates = [];

  data.forEach(d => {
    counter[d.nama] = (counter[d.nama] || 0) + 1;
  });

  data.forEach(d => {
    if (counter[d.nama] > 1) {
      duplicates.push(d);
    }
  });

  return duplicates;
}

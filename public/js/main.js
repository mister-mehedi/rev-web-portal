// public/js/main.js
document.addEventListener('DOMContentLoaded', () => {
  // tab handling
  const tabBtns = document.querySelectorAll('.tab-btn');
  const panels = document.querySelectorAll('.tab-panel');
  const resultsArea = document.getElementById('results-area');
  let currentRows = []; // currently displayed data
  let currentColumns = [];

  tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      tabBtns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      const tab = btn.getAttribute('data-tab');
      panels.forEach(p => p.classList.toggle('hidden', p.id !== tab));
      resultsArea.classList.add('hidden');
    });
  });

  // init flatpickr for any .date-input
  flatpickr('.date-input', { dateFormat: 'd-M-y' });

  // helper: render table and DataTable
  let dataTable = null;
  function renderTable(rows) {
    const thead = document.getElementById('results-thead');
    const tbody = document.getElementById('results-tbody');
    thead.innerHTML = '';
    tbody.innerHTML = '';

    if (!rows || rows.length === 0) {
      thead.innerHTML = '<tr><th>No results</th></tr>';
      if (dataTable) { dataTable.destroy(); dataTable = null; }
      return;
    }

    // columns
    const cols = Object.keys(rows[0]);
    currentColumns = cols;
    const headerRow = document.createElement('tr');
    cols.forEach(c => {
      const th = document.createElement('th');
      th.textContent = c;
      headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);

    // body
    rows.forEach(r => {
      const tr = document.createElement('tr');
      cols.forEach(c => {
        const td = document.createElement('td');
        let val = r[c];
        if (val === null || val === undefined) val = '';
        td.textContent = val;
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });

    // DataTables (destroy previous)
    if (dataTable) { dataTable.destroy(); dataTable = null; }
    dataTable = $('#results-table').DataTable({
      paging: true,
      searching: true,
      ordering: true,
      pageLength: 25
    });

    currentRows = rows;
  }

  // helper draw chart (simple numeric trend)
  let chart = null;
  function drawChart(rows) {
    const ctx = document.getElementById('results-chart').getContext('2d');

    // choose first numeric column (except keys that are text)
    if (!rows || rows.length === 0) {
      if (chart) { chart.destroy(); chart = null; }
      return;
    }

    const cols = Object.keys(rows[0]);
    // try to find a numeric column to plot (Total_Revenue*, DATA_REV_TOT, etc)
    let numericCol = cols.find(c => /TOTAL_REVENUE|DATA_REV|TOTAL_DATA|AVG/i.test(c));
    if (!numericCol) {
      numericCol = cols.find(c => typeof rows[0][c] === 'number');
    }
    if (!numericCol) {
      if (chart) { chart.destroy(); chart = null; }
      return;
    }

    const labels = rows.map(r => {
      // pick label column: chunk or month or officer name
      if (r.CHUNK_NO !== undefined) return String(r.CHUNK_NO);
      if (r.MONTH_KEY !== undefined) return String(r.MONTH_KEY);
      if (r.FIELD_OFFICER_NAME !== undefined) return String(r.FIELD_OFFICER_NAME).slice(0,15);
      return '';
    });

    const dataVals = rows.map(r => Number(r[numericCol]) || 0);

    if (chart) chart.destroy();
    chart = new Chart(ctx, {
      type: 'bar',
      data: {
        labels,
        datasets: [{ label: numericCol, data: dataVals }]
      },
      options: { responsive: true }
    });
  }

  // generic AJAX function to call /api/query
  async function runQueryAjax(queryName, params) {
    try {
      const res = await fetch('/api/query', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ queryName, params })
      });
      const payload = await res.json();
      if (res.ok) {
        const rows = payload.rows || [];
        document.getElementById('results-area').classList.remove('hidden');
        renderTable(rows);
        drawChart(rows);
      } else {
        alert('Query failed: ' + (payload.error || 'unknown error'));
      }
    } catch (err) {
      console.error(err);
      alert('Request error: ' + err.message);
    }
  }

  // wire forms
  document.getElementById('form-day-chunk').addEventListener('submit', (e) => {
    e.preventDefault();
    const fd = new FormData(e.target);
    runQueryAjax('day-chunk', {
      baseDate: fd.get('baseDate'),
      chunkDays: fd.get('chunkDays'),
      numChunks: fd.get('numChunks')
    });
  });

  document.getElementById('form-month-chunk').addEventListener('submit', (e) => {
    e.preventDefault();
    const fd = new FormData(e.target);
    runQueryAjax('month-chunk', {
      baseMonth: fd.get('baseMonth'),
      chunkMonths: fd.get('chunkMonths')
    });
  });

  document.getElementById('form-fo-day').addEventListener('submit', (e) => {
    e.preventDefault();
    const fd = new FormData(e.target);
    runQueryAjax('fo-day', {
      baseDate: fd.get('baseDate'),
      chunkDays: fd.get('chunkDays'),
      numChunks: fd.get('numChunks')
    });
  });

  document.getElementById('form-fo-month').addEventListener('submit', (e) => {
    e.preventDefault();
    const fd = new FormData(e.target);
    runQueryAjax('fo-month', {
      baseMonth: fd.get('baseMonth'),
      chunkMonths: fd.get('chunkMonths')
    });
  });

  // Export buttons
  document.getElementById('btn-export-csv').addEventListener('click', async () => {
    if (!currentRows || currentRows.length === 0) return alert('No data to export');
    try {
      const res = await fetch('/export/csv', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ filename: 'export', rows: currentRows })
      });
      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = 'export.csv'; document.body.appendChild(a); a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (err) { console.error(err); alert('Export failed: ' + err.message); }
  });

  document.getElementById('btn-export-excel').addEventListener('click', async () => {
    if (!currentRows || currentRows.length === 0) return alert('No data to export');
    try {
      const res = await fetch('/export/excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ filename: 'export', rows: currentRows })
      });
      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = 'export.xlsx'; document.body.appendChild(a); a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (err) { console.error(err); alert('Export failed: ' + err.message); }
  });

  // copy table to clipboard
  document.getElementById('btn-copy').addEventListener('click', async () => {
    if (!currentRows || currentRows.length === 0) return alert('No data to copy');
    // CSV text
    const cols = currentColumns;
    const lines = [cols.join('\t')].concat(currentRows.map(r => cols.map(c => r[c] ?? '').join('\t')));
    const text = lines.join('\n');
    try {
      await navigator.clipboard.writeText(text);
      alert('Copied to clipboard');
    } catch (err) {
      alert('Copy failed: ' + err.message);
    }
  });

  // print
  document.getElementById('btn-print').addEventListener('click', () => {
    window.print();
  });

});

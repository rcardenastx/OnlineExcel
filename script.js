async function loadExcel() {
  const url = 'https://raw.githubusercontent.com/rcardenastx/OnlineExcel/main/data.xlsx';
  try {
    const response = await fetch(url);
    if (!response.ok) throw new Error('Network response was not OK');
    
    const data = await response.arrayBuffer();
    const workbook = XLSX.read(data, { type: 'array' });

    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    renderTable(json);
  } catch (error) {
    document.getElementById('table-container').innerText = 'Failed to load Excel file.';
    console.error('Error loading Excel:', error);
  }
}

function renderTable(data) {
  const container = document.getElementById('table-container');
  if (!data || data.length === 0) {
    container.innerHTML = 'No data found.';
    return;
  }

  let html = '<table><thead><tr>';
  data[0].forEach(header => {
    html += `<th>${header}</th>`;
  });
  html += '</tr></thead><tbody>';

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    html += '<tr>';
    row.forEach(cell => {
      html += `<td>${formatCell(cell)}</td>`;
    });
    html += '</tr>';
  }

  html += '</tbody></table>';
  container.innerHTML = html;
}

function formatCell(cell) {
  if (typeof cell === 'number' && cell > 0 && cell < 1) {
    const totalMinutes = Math.round(cell * 24 * 60);
    const hours = Math.floor(totalMinutes / 60);
    const minutes = totalMinutes % 60;

    const hour12 = hours % 12 || 12;
    const ampm = hours >= 12 ? 'PM' : 'AM';

    return `${hour12}:${minutes.toString().padStart(2, '0')} ${ampm}`;
  }

 
  if (typeof cell === 'number' && cell > 30000) {
    const date = XLSX.SSF.parse_date_code(cell);
    if (date) {
      return `${date.y}-${pad(date.m)}-${pad(date.d)}`;
    }
  }

  return cell ?? '';
}

function pad(n) {
  return n.toString().padStart(2, '0');
}

// Start loading
loadExcel();

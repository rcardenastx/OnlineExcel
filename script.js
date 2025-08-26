async function loadExcel() {
  const url = 'https://raw.githubusercontent.com/rcardenastx/OnlineExcel/main/data.xlsx'; // 
  const response = await fetch(url);
  const data = await response.arrayBuffer();

  const workbook = XLSX.read(data, { type: 'array' });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 }); // Get 2D array

  renderTable(json);
}

function renderTable(data) {
  const container = document.getElementById('table-container');
  if (data.length === 0) {
    container.innerHTML = 'No data found.';
    return;
  }

  let html = '<table><thead><tr>';
  data[0].forEach(header => {
    html += `<th>${header}</th>`;
  });
  html += '</tr></thead><tbody>';

  data.slice(1).forEach(row => {
    html += '<tr>';
    row.forEach(cell => {
      html += `<td>${formatCell(cell)}</td>`;
    });
    html += '</tr>';
  });

  html += '</tbody></table>';
  container.innerHTML = html;
}

function formatCell(cell) {
  // Format Excel-style decimal times to readable time
  if (typeof cell === 'number' && cell > 0 && cell < 1) {
    const totalMinutes = Math.round(cell * 24 * 60);
    const hours = Math.floor(totalMinutes / 60);
    const minutes = totalMinutes % 60;

    const hour12 = hours % 12 || 12;
    const ampm = hours >= 12 ? 'PM' : 'AM';

    return `${hour12}:${minutes.toString().padStart(2, '0')} ${ampm}`;
  }
  return cell ?? '';
}


  let html = '<table><thead><tr>';
  data[0].forEach(header => {
    html += `<th>${header}</th>`;
  });
  html += '</tr></thead><tbody>';

  data.slice(1).forEach(row => {
    html += '<tr>';
    row.forEach(cell => {
      html += `<td>${cell ?? ''}</td>`;
    });
    html += '</tr>';
  });

  html += '</tbody></table>';
  container.innerHTML = html;
}

loadExcel();

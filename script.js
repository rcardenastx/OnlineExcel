// script.js
fetch('https://raw.githubusercontent.com/rcardenastx/OnlineExcel/main/data.xlsx')
  .then(res => res.arrayBuffer())
  .then(data => {
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet);
    console.table(json);
});

let allSheetsData = {};
let currentSheet = "";
let allRows = [];
let filteredRows = [];
let selectedColumns = [];
let gradeColumnIndex = -1;

document.getElementById('file-input').addEventListener('change', handleFile, false);
document.getElementById('filter-grade').addEventListener('click', filterByGrade);
document.getElementById('download-excel').addEventListener('click', downloadExcel);
document.getElementById('sheet-select').addEventListener('change', changeSheet);

function handleFile(event) {
  const file = event.target.files[0];
  const reader = new FileReader();

  reader.onload = function(e) {
    const data = e.target.result;
    const workbook = XLSX.read(data, { type: 'binary' });

    allSheetsData = {};  
    workbook.SheetNames.forEach(sheetName => {
      allSheetsData[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
    });

    populateSheetSelector(workbook.SheetNames);
  };

  reader.readAsBinaryString(file);
}

function populateSheetSelector(sheetNames) {
  const sheetSelect = document.getElementById('sheet-select');
  sheetSelect.innerHTML = "";  

  sheetNames.forEach(sheet => {
    let option = document.createElement('option');
    option.value = sheet;
    option.textContent = sheet;
    sheetSelect.appendChild(option);
  });

  currentSheet = sheetNames[0];
  updateSheetData();
}

function changeSheet() {
  currentSheet = document.getElementById('sheet-select').value;
  updateSheetData();
}

// function updateSheetData() {
//   allRows = allSheetsData[currentSheet] || [];
//   if (allRows.length === 0) {
//     alert("La hoja seleccionada está vacía.");
//     return;
//   }

//   gradeColumnIndex = allRows[0].indexOf("Grado al que se matricula");
//   filteredRows = [...allRows];
//   displayTable(filteredRows);
// }
function cleanString(str) {
  return str
    .trim() // Eliminar espacios al inicio y al final
    .replace(/\s+/g, ' ') // Eliminar espacios múltiples y reemplazarlos por un solo espacio
    .normalize("NFD") // Eliminar tildes
    .replace(/[\u0300-\u036f]/g, "") // Eliminar caracteres de acentuación
    .replace(/[^\x20-\x7E]/g, '') // Eliminar caracteres no imprimibles
    .toLowerCase(); // Convertir a minúsculas
}

function findGradeColumn() {
  const headers = allRows.flat(); // Aplanar todas las filas
  const normalizedHeaders = headers.map(header => cleanString(header));

  // Intentar encontrar el índice de la columna "grado al que se matricula"
  return normalizedHeaders.indexOf(cleanString("grado al que se matricula"));
}

function cleanString(str) {
  return str
    .trim() // Eliminar espacios al inicio y al final
    .replace(/\s+/g, ' ') // Eliminar espacios múltiples y reemplazarlos por un solo espacio
    .normalize("NFD") // Eliminar tildes
    .replace(/[\u0300-\u036f]/g, "") // Eliminar caracteres de acentuación
    .replace(/[^\x20-\x7E]/g, '') // Eliminar caracteres no imprimibles
    .toLowerCase(); // Convertir a minúsculas
}

function updateSheetData() {
  allRows = allSheetsData[currentSheet] || [];
  if (allRows.length === 0) {
    alert("La hoja seleccionada está vacía.");
    return;
  }

  const headers = allRows[0].map(header => header ? cleanString(header) : "");
  console.log("Encabezados de la hoja: ", headers);  // Ver los encabezados reales

  // Intentar buscar la columna "grado al que se matricula" en cada fila
  for (let i = 0; i < allRows.length; i++) {
    let row = allRows[i];
    let rowHeaders = row.map(cell => cleanString(cell));
    console.log("Encabezados fila", i, rowHeaders);  // Ver los datos de cada fila

    gradeColumnIndex = rowHeaders.indexOf(cleanString("grado al que se matricula"));
    if (gradeColumnIndex !== -1) {
      console.log(`Columna encontrada en la fila ${i}`);
      break;
    }
  }

  if (gradeColumnIndex === -1) {
    alert(`⚠️ Advertencia: No se encontró la columna "Grado al que se matricula" en la hoja "${currentSheet}".`);
  }

  filteredRows = [...allRows];
  displayTable(filteredRows);
}


function displayTable(rows) {
  const tableBody = document.querySelector('#excel-table tbody');
  const tableHead = document.querySelector('#excel-table thead');
  tableBody.innerHTML = '';
  tableHead.innerHTML = '';

  if (rows.length === 0) return;

  // Crear la fila de encabezado, independientemente si es <td> o <th>
  const headerRow = document.createElement('tr');
  rows[0].forEach((col, index) => {
    const cell = document.createElement(index === 0 ? 'th' : 'td'); // Asegura que la primera columna siempre sea <th>
    cell.textContent = col;
    headerRow.appendChild(cell);
  });
  tableHead.appendChild(headerRow);

  // Agregar las filas de datos
  rows.slice(1).forEach(row => {
    const tr = document.createElement('tr');
    row.forEach(cell => {
      const td = document.createElement('td');
      td.textContent = cell || '';
      tr.appendChild(td);
    });
    tableBody.appendChild(tr);
  });
}

// function displayTable(rows) {
//   const tableBody = document.querySelector('#excel-table tbody');
//   const tableHead = document.querySelector('#excel-table thead');
//   tableBody.innerHTML = '';
//   tableHead.innerHTML = '';

//   if (rows.length === 0) return;

//   const headerRow = document.createElement('tr');
//   rows[0].forEach(col => {
//     const th = document.createElement('th');
//     th.textContent = col;
//     headerRow.appendChild(th);
//   });
//   tableHead.appendChild(headerRow);

//   rows.slice(1).forEach(row => {
//     const tr = document.createElement('tr');
//     row.forEach(cell => {
//       const td = document.createElement('td');
//       td.textContent = cell || '';
//       tr.appendChild(td);
//     });
//     tableBody.appendChild(tr);
//   });
// }

function filterByGrade() {
  if (gradeColumnIndex === -1) {
    alert("No se encontró la columna 'Grado al que se matricula'.");
    return;
  }

  const uniqueGrades = [...new Set(allRows.slice(1).map(row => row[gradeColumnIndex]).filter(Boolean))];

  if (uniqueGrades.length === 0) {
    alert("No hay datos disponibles para filtrar.");
    return;
  }

  const selectedGrade = prompt(`Selecciona un grado para filtrar:\n${uniqueGrades.join(", ")}`);

  if (selectedGrade && uniqueGrades.includes(selectedGrade.trim())) {
    filteredRows = [allRows[0]];
    filteredRows.push(...allRows.slice(1).filter(row => row[gradeColumnIndex] === selectedGrade.trim()));

    displayTable(filteredRows);
  } else {
    alert("Grado no válido o no seleccionado.");
  }
}

document.querySelector('#excel-table').addEventListener('click', function(event) {
  const columnIndex = event.target.cellIndex;
  if (columnIndex === undefined) return;

  if (selectedColumns.includes(columnIndex)) {
    selectedColumns = selectedColumns.filter(index => index !== columnIndex);
    event.target.style.backgroundColor = '';
  } else {
    selectedColumns.push(columnIndex);
    event.target.style.backgroundColor = 'yellow';
  }

  updateFilteredTable();
});

function updateFilteredTable() {
  const resultTableBody = document.querySelector('#result-table tbody');
  const resultTableHead = document.querySelector('#result-table thead');

  resultTableBody.innerHTML = '';
  resultTableHead.innerHTML = '';

  if (selectedColumns.length === 0) return;

  const headerRow = document.createElement('tr');
  selectedColumns.forEach(index => {
    const th = document.createElement('th');
    th.textContent = filteredRows[0][index];
    headerRow.appendChild(th);
  });
  resultTableHead.appendChild(headerRow);

  filteredRows.slice(1).forEach(row => {
    const tr = document.createElement('tr');
    selectedColumns.forEach(index => {
      const td = document.createElement('td');
      td.textContent = row[index] || '';
      tr.appendChild(td);
    });
    resultTableBody.appendChild(tr);
  });
}

function downloadExcel() {
  if (selectedColumns.length === 0) {
    alert("Selecciona al menos una columna.");
    return;
  }

  const worksheetData = [];
  worksheetData.push(selectedColumns.map(index => filteredRows[0][index]));

  filteredRows.slice(1).forEach(row => {
    worksheetData.push(selectedColumns.map(index => row[index]));
  });

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(worksheetData);
  XLSX.utils.book_append_sheet(wb, ws, "Datos Filtrados");
  XLSX.writeFile(wb, "Datos_Filtrados.xlsx");
}

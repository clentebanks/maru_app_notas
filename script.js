let allRows = []; // Almacena todas las filas del archivo
let filteredRows = []; // Almacena las filas después del filtro por grado
let selectedColumns = []; // Almacena los índices de columnas seleccionadas
let gradeColumnIndex = -1; // Índice de la columna "Grado al que se matricula"

document.getElementById('file-input').addEventListener('change', handleFile, false);

function handleFile(event) {
  const file = event.target.files[0];
  const reader = new FileReader();

  reader.onload = function(e) {
    const data = e.target.result;
    const workbook = XLSX.read(data, { type: 'binary' });

    // Obtener la primera hoja
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    allRows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

    if (allRows.length === 0) {
      alert("El archivo no contiene datos.");
      return;
    }

    // Buscar el índice de la columna "Grado al que se matricula"
    gradeColumnIndex = allRows[0].indexOf("Grado al que se matricula");

    if (gradeColumnIndex === -1) {
      alert("No se encontró la columna 'Grado al que se matricula'.");
      return;
    }

    // Inicialmente, mostrar todos los datos
    filteredRows = [...allRows];
    displayTable(filteredRows);
  };

  reader.readAsBinaryString(file);
}

function displayTable(rows) {
  const tableBody = document.querySelector('#excel-table tbody');
  const tableHead = document.querySelector('#excel-table thead');

  tableBody.innerHTML = '';
  tableHead.innerHTML = '';

  // Agregar encabezado
  const headerRow = document.createElement('tr');
  rows[0].forEach(col => {
    const th = document.createElement('th');
    th.textContent = col;
    headerRow.appendChild(th);
  });
  tableHead.appendChild(headerRow);

  // Agregar filas de datos
  rows.slice(1).forEach(row => {
    const tr = document.createElement('tr');
    row.forEach(cell => {
      const td = document.createElement('td');
      td.textContent = cell || ''; // Evitar celdas vacías
      tr.appendChild(td);
    });
    tableBody.appendChild(tr);
  });
}

// FILTRAR POR GRADO
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

  if (selectedGrade && uniqueGrades.includes(selectedGrade)) {
    filteredRows = [allRows[0]]; // Mantener encabezado
    filteredRows.push(...allRows.slice(1).filter(row => row[gradeColumnIndex] === selectedGrade));

    displayTable(filteredRows);
  } else {
    alert("Grado no válido o no seleccionado.");
  }
}

// SELECCIONAR COLUMNAS
document.querySelector('#excel-table').addEventListener('click', function(event) {
  const columnIndex = event.target.cellIndex;
  if (columnIndex === undefined) return;

  // Alternar selección de columna
  if (selectedColumns.includes(columnIndex)) {
    selectedColumns = selectedColumns.filter(index => index !== columnIndex);
    event.target.style.backgroundColor = '';
  } else {
    selectedColumns.push(columnIndex);
    event.target.style.backgroundColor = 'yellow';
  }

  updateFilteredTable();
});

// ACTUALIZAR TABLA CON COLUMNAS SELECCIONADAS Y FILTRO
function updateFilteredTable() {
  const resultTableBody = document.querySelector('#result-table tbody');
  const resultTableHead = document.querySelector('#result-table thead');

  resultTableBody.innerHTML = '';
  resultTableHead.innerHTML = '';

  if (selectedColumns.length === 0) return;

  // Agregar encabezado con columnas seleccionadas
  const headerRow = document.createElement('tr');
  selectedColumns.forEach(index => {
    const th = document.createElement('th');
    th.textContent = filteredRows[0][index];
    headerRow.appendChild(th);
  });
  resultTableHead.appendChild(headerRow);

  // Agregar datos filtrados con columnas seleccionadas
  filteredRows.slice(1).forEach(row => {
    const tr = document.createElement('tr');
    selectedColumns.forEach(index => {
      const td = document.createElement('td');
      td.textContent = row[index] || ''; // Evitar celdas vacías
      tr.appendChild(td);
    });
    resultTableBody.appendChild(tr);
  });
}

// DESCARGAR FILTRO SELECCIONADO
document.getElementById('download-excel').addEventListener('click', () => {
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
});
document.getElementById('filter-grade').addEventListener('click', filterByGrade);

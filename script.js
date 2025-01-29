
    let allRows = [];
    let filteredRows = [];
    let selectedColumns = [];
    let gradeColumnIndex = -1;

    document.getElementById('file-input').addEventListener('change', handleFile, false);
    document.getElementById('filter-grade').addEventListener('click', filterByGrade);
    document.getElementById('download-excel').addEventListener('click', downloadExcel);

    function handleFile(event) {
      const file = event.target.files[0];
      const reader = new FileReader();

      reader.onload = function(e) {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        allRows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

        if (allRows.length === 0) {
          alert("El archivo no contiene datos.");
          return;
        }

        gradeColumnIndex = allRows[0].indexOf("Grado al que se matricula");
        if (gradeColumnIndex === -1) {
          alert("No se encontró la columna 'Grado al que se matricula'.");
        }

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

      if (rows.length === 0) return;

      const headerRow = document.createElement('tr');
      rows[0].forEach(col => {
        const th = document.createElement('th');
        th.textContent = col;
        headerRow.appendChild(th);
      });
      tableHead.appendChild(headerRow);

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
 

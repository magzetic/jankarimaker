// Listen to the configuration form submission
document.getElementById('configForm').addEventListener('submit', function (e) {
    e.preventDefault();
  
    // Get configuration values
    const sheetTitle = document.getElementById('sheetTitle').value;
    const numColumns = parseInt(document.getElementById('numColumns').value, 10);
    const numRows = parseInt(document.getElementById('numRows').value, 10);
  
    // Clear any previous forms
    document.getElementById('columnTitlesDiv').innerHTML = '';
    document.getElementById('dataEntryDiv').innerHTML = '';
  
    // Generate inputs for column headers
    const colTitlesDiv = document.getElementById('columnTitlesDiv');
    const colForm = document.createElement('form');
    colForm.id = 'colTitlesForm';
  
    const header = document.createElement('h2');
    header.textContent = 'Enter Column Titles';
    colForm.appendChild(header);
  
    const inputContainer = document.createElement('div');
    inputContainer.className = 'form-row';
  
    for (let i = 0; i < numColumns; i++) {
      const input = document.createElement('input');
      input.type = 'text';
      input.placeholder = `Column ${i + 1} Title`;
      input.required = true;
      input.name = `col${i}`;
      inputContainer.appendChild(input);
    }
    colForm.appendChild(inputContainer);
  
    const colSubmit = document.createElement('button');
    colSubmit.type = 'submit';
    colSubmit.textContent = 'Next';
    colForm.appendChild(colSubmit);
    colTitlesDiv.appendChild(colForm);
  
    colForm.addEventListener('submit', function (e) {
      e.preventDefault();
  
      // Get column titles from the form
      const columnTitles = [];
      for (let i = 0; i < numColumns; i++) {
        columnTitles.push(e.target.elements[`col${i}`].value);
      }
  
      // Generate the data entry form
      const dataDiv = document.getElementById('dataEntryDiv');
      dataDiv.innerHTML = '';  // Clear any existing content
  
      const dataForm = document.createElement('form');
      dataForm.id = 'dataForm';
  
      const dataHeader = document.createElement('h2');
      dataHeader.textContent = 'Enter Data';
      dataForm.appendChild(dataHeader);
  
      // Create a row of input fields for each data record
      for (let row = 0; row < numRows; row++) {
        const rowDiv = document.createElement('div');
        rowDiv.className = 'form-row';
        for (let col = 0; col < numColumns; col++) {
          const input = document.createElement('input');
          input.type = 'text';
          input.placeholder = `${columnTitles[col]} (Row ${row + 1})`;
          input.required = true;
          input.name = `cell_${row}_${col}`;
          rowDiv.appendChild(input);
        }
        dataForm.appendChild(rowDiv);
      }
  
      const dataSubmit = document.createElement('button');
      dataSubmit.type = 'submit';
      dataSubmit.textContent = 'Generate Excel';
      dataForm.appendChild(dataSubmit);
      dataDiv.appendChild(dataForm);
  
      // Listen to the data form submission
      dataForm.addEventListener('submit', function (e) {
        e.preventDefault();
  
        // Collect the data into a 2D array
        const dataRows = [];
        for (let row = 0; row < numRows; row++) {
          const rowData = [];
          for (let col = 0; col < numColumns; col++) {
            rowData.push(e.target.elements[`cell_${row}_${col}`].value);
          }
          dataRows.push(rowData);
        }
  
        // Generate the Excel file using ExcelJS with dynamic column width, borders, and center alignment
        generateExcelFile(sheetTitle, columnTitles, dataRows);
      });
    });
  });
  
  // Function to generate and download the Excel file using ExcelJS with dynamic column widths,
  // visible cell borders, and centered cell content
  function generateExcelFile(sheetTitle, columnTitles, dataRows) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');
  
    // Add a top heading (merged across the columns)
    const titleRow = sheet.addRow([sheetTitle]);
    titleRow.font = { size: 16, bold: true };
    sheet.mergeCells(1, 1, 1, columnTitles.length);
  
    // Optional: add an empty row for spacing
    sheet.addRow([]);
  
    // Add header row with column titles
    const headerRow = sheet.addRow(columnTitles);
    headerRow.font = { bold: true };
  
    // Add data rows
    dataRows.forEach((rowData) => {
      sheet.addRow(rowData);
    });
  
    // Set dynamic column widths based on the maximum length of header and cell content.
    const columnWidths = columnTitles.map((title, i) => {
      let maxLength = title.toString().length;
      dataRows.forEach(row => {
        const cellValue = row[i] ? row[i].toString() : '';
        if (cellValue.length > maxLength) {
          maxLength = cellValue.length;
        }
      });
      // Add some extra padding to the width
      return { width: maxLength + 2 };
    });
    sheet.columns = columnWidths;
  
    // Apply borders and center alignment to every cell
    sheet.eachRow({ includeEmpty: true }, (row) => {
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });
  
    // Create a downloadable Excel file
    workbook.xlsx.writeBuffer().then((buffer) => {
      const blob = new Blob(
        [buffer],
        { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
      );
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = `${sheetTitle}.xlsx`;
      link.click();
    });
  }
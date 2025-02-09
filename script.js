document.getElementById('configForm').addEventListener('submit', function (e) {
  e.preventDefault();

  const sheetTitle = document.getElementById('sheetTitle').value;
  const numColumns = parseInt(document.getElementById('numColumns').value, 10);
  const numRows = parseInt(document.getElementById('numRows').value, 10);

  // Clear previous forms and previews
  document.getElementById('columnTitlesDiv').innerHTML = '';
  document.getElementById('dataEntryDiv').innerHTML = '';
  document.getElementById('tablePreviewDiv').innerHTML = '';
  document.getElementById('buttonsDiv').style.display = 'none';

  // Generate Column Titles Form
  const colTitlesDiv = document.getElementById('columnTitlesDiv');
  const colForm = document.createElement('form');
  colForm.id = 'colTitlesForm';

  const header = document.createElement('h2');
  header.textContent = 'जानकारी के कॉलम भरें';
  colForm.appendChild(header);

  for (let i = 0; i < numColumns; i++) {
    const input = document.createElement('input');
    input.type = 'text';
    input.placeholder = `कॉलम ${i + 1} `;
    input.required = true;
    input.name = `col${i}`;
    colForm.appendChild(input);
  }

  const colSubmit = document.createElement('button');
  colSubmit.type = 'submit';
  colSubmit.textContent = 'Next';
  colForm.appendChild(colSubmit);
  colTitlesDiv.appendChild(colForm);

  colForm.addEventListener('submit', function (e) {
    e.preventDefault();

    const columnTitles = [];
    for (let i = 0; i < numColumns; i++) {
      columnTitles.push(e.target.elements[`col${i}`].value);
    }

    generateDataForm(columnTitles, numRows, sheetTitle);
  });
});

function generateDataForm(columnTitles, numRows, sheetTitle) {
  const dataDiv = document.getElementById('dataEntryDiv');
  dataDiv.innerHTML = '';

  const dataForm = document.createElement('form');
  dataForm.id = 'dataForm';

  const dataHeader = document.createElement('h2');
  dataHeader.textContent = 'जानकारी भरें';
  dataForm.appendChild(dataHeader);

  for (let row = 0; row < numRows; row++) {
      // Create a heading for the row
      const rowHeading = document.createElement('h3');
      rowHeading.textContent = `क्र ${row + 1}`;
      rowHeading.style = "font-size: 16px; font-weight: bold; color: #007bff; margin-top: 15px; margin-bottom: 5px;";
      dataForm.appendChild(rowHeading);

      // Create row container
      const rowDiv = document.createElement('div');
      rowDiv.className = 'form-row'; // CSS class for spacing
      rowDiv.style = "display: flex; flex-wrap: wrap; gap: 10px; margin-bottom: 15px;"; // Inline styling

      for (let col = 0; col < columnTitles.length; col++) {
          const fieldContainer = document.createElement('div');
          fieldContainer.style = "display: flex; flex-direction: column;"; // Align label above input

          // Create label
          const label = document.createElement('label');
          label.textContent = columnTitles[col];
          label.style = "font-size: 14px; font-weight: bold; margin-bottom: 5px;";

          // Create input field
          const input = document.createElement('input');
          input.type = 'text';
          input.placeholder = `${columnTitles[col]} (क्र  ${row + 1})`;
          input.required = true;
          input.name = `cell_${row}_${col}`;
          input.style = "padding: 8px; width: 250px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px;";

          // Append label and input
          fieldContainer.appendChild(label);
          fieldContainer.appendChild(input);
          rowDiv.appendChild(fieldContainer);
      }

      dataForm.appendChild(rowDiv);
  }

  const generateTableBtn = document.createElement('button');
  generateTableBtn.type = 'submit';
  generateTableBtn.textContent = 'जानकारी / सूची बनाऐं';
  dataForm.appendChild(generateTableBtn);
  dataDiv.appendChild(dataForm);

  dataForm.addEventListener('submit', function (e) {
      e.preventDefault();

      const dataRows = [];
      for (let row = 0; row < numRows; row++) {
          const rowData = [];
          for (let col = 0; col < columnTitles.length; col++) {
              rowData.push(e.target.elements[`cell_${row}_${col}`].value);
          }
          dataRows.push(rowData);
      }

      generateTablePreview(sheetTitle, columnTitles, dataRows);
  });
}


function generateTablePreview(sheetTitle, columnTitles, dataRows) {
  const tableDiv = document.getElementById('tablePreviewDiv');
  tableDiv.innerHTML = '';

  // Add heading for the sheet title
  const heading = document.createElement('h2');
  heading.textContent = sheetTitle;
  tableDiv.appendChild(heading);

  const table = document.createElement('table');
  table.border = '1';

  const headerRow = document.createElement('tr');
  columnTitles.forEach(title => {
    const th = document.createElement('th');
    th.textContent = title;
    headerRow.appendChild(th);
  });
  table.appendChild(headerRow);

  dataRows.forEach(rowData => {
    const row = document.createElement('tr');
    rowData.forEach(cellData => {
      const td = document.createElement('td');
      td.textContent = cellData;
      row.appendChild(td);
    });
    table.appendChild(row);
  });

  tableDiv.appendChild(table);
  document.getElementById('buttonsDiv').style.display = 'flex';

  // Update button event listeners to use the new PDF conversion tool
  document.getElementById('generatePDF').onclick = () => generatePDF(sheetTitle);
  document.getElementById('generateExcel').onclick = () => generateExcel(sheetTitle, columnTitles, dataRows);
}

function generatePDF(sheetTitle) {
  const element = document.getElementById('tablePreviewDiv');

  if (!element || element.innerHTML.trim() === "") {
      alert("Error: No table found to generate PDF!");
      return;
  }

  // Ensure the table is fully visible before capturing
  element.style.display = "block";

  // Wait a short time before capturing to ensure rendering
  setTimeout(() => {
      const opt = {
          margin: [0,0,0,0],
          filename: sheetTitle + '.pdf',
          image: { type: 'jpeg', quality: 0.80 },
          html2canvas: {
              scale: 2, // Higher scale for better quality
              useCORS: true, // Fixes cross-origin issues
              scrollX: 0,
              scrollY: 0
          },
          jsPDF: { unit: 'mm', format: 'a3', orientation: 'landscape' } // LANDSCAPE MODE
      };

      html2pdf().set(opt).from(element).save();
  }, 1000); // Give time to render before capturing
}



function generateExcel(sheetTitle, columnTitles, dataRows) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');

  // Add a top heading (merged across the columns)
  const titleRow = sheet.addRow([sheetTitle]);
  titleRow.font = { size: 16, bold: true };
  sheet.mergeCells(1, 1, 1, columnTitles.length);

  // Optional: add an empty row for spacing
  sheet.addRow([]);
  const headerRow = sheet.addRow(columnTitles);
  headerRow.font = { bold: true };

  dataRows.forEach(rowData => {
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

  workbook.xlsx.writeBuffer().then(buffer => {
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `${sheetTitle}.xlsx`;
    link.click();
  });
}
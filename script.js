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

  // --- Add Header Row with "क्रमांक" Column ---
  const headerRow = document.createElement('tr');

  // Add "क्रमांक" as the first column
  const srNoTh = document.createElement('th');
  srNoTh.textContent = 'क्रमांक';
  headerRow.appendChild(srNoTh);

  columnTitles.forEach(title => {
    const th = document.createElement('th');
    th.textContent = title;
    headerRow.appendChild(th);
  });

  // Add blank column header (rightmost)
  const blankTh = document.createElement('th');
  blankTh.className = 'blank-column';
  headerRow.appendChild(blankTh);
  table.appendChild(headerRow);

  // --- Add Data Rows with Incremental "क्रमांक" Column ---
  dataRows.forEach((rowData, index) => {
    const row = document.createElement('tr');

    // Add Serial Number Column
    const srNoTd = document.createElement('td');
    srNoTd.textContent = index + 1; // Incremental Serial Number
    row.appendChild(srNoTd);

    // Add Data Columns
    rowData.forEach(cellData => {
      const td = document.createElement('td');
      td.textContent = cellData;
      row.appendChild(td);
    });

    // Add Blank Column (rightmost)
    const blankTd = document.createElement('td');
    blankTd.className = 'blank-column';
    row.appendChild(blankTd);
    table.appendChild(row);
  });

  tableDiv.appendChild(table);
  document.getElementById('buttonsDiv').style.display = 'flex';

  // Update button event listeners
  document.getElementById('generatePDF').onclick = () => generatePDF(sheetTitle);
  document.getElementById('generateExcel').onclick = () => generateExcel(sheetTitle, ['क्रमांक', ...columnTitles], dataRows);
}

function generatePDF(sheetTitle) {
  const originalElement = document.getElementById('tablePreviewDiv');
  if (!originalElement || originalElement.innerHTML.trim() === "") {
    alert("Error: No table found to generate PDF!");
    return;
  }
  
  // Clone the table preview so the original page isn't modified
  const clonedElement = originalElement.cloneNode(true);
  
  // Create a container for the PDF content
  const pdfContainer = document.createElement('div');
  pdfContainer.style.background = '#fff';
  pdfContainer.style.padding = '20px';
  
  // Append the cloned table preview into the container
  pdfContainer.appendChild(clonedElement);
  
  // Create the credit element that will appear below the table
  const creditDiv = document.createElement('div');
  creditDiv.innerHTML = 'यह जानकारी <a href="https://jankarimaker.in" target="_blank">jankarimaker.in</a> की सहायता से बनाइ गइ है';
  creditDiv.style.fontSize = '10px';
  creditDiv.style.color = '#333';
  creditDiv.style.textAlign = 'right';
  creditDiv.style.marginTop = '240px'; // Adds space between the table and the credit
  
  // Append the credit block after the table preview
  pdfContainer.appendChild(creditDiv);
  
  // PDF generation options
  const opt = {
    margin: [10, 10, 10, 10],
    filename: sheetTitle + '.pdf',
    image: { type: 'jpeg', quality: 0.98 },
    html2canvas: {
      scale: 2,
      useCORS: true,
      scrollX: 0,
      scrollY: 0,
      logging: true,
    },
    jsPDF: { unit: 'mm', format: 'a3', orientation: 'landscape' }
  };
  
  // Generate and download the PDF from the container
  html2pdf().set(opt).from(pdfContainer).save();
}

function generateExcel(sheetTitle, columnTitles, dataRows) {
  // If the first column title is already "क्रमांक", remove it to avoid duplication
  if (columnTitles.length > 0 && columnTitles[0].trim() === 'क्रमांक') {
    columnTitles = columnTitles.slice(1);
  }

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');

  // 1. Title Row (merged across all columns)
  const titleRow = sheet.addRow([sheetTitle]);
  titleRow.font = { size: 16, bold: true };
  // Merge from column 1 to the number of data columns (1 for our serial + remaining columns)
  sheet.mergeCells(1, 1, 1, columnTitles.length + 1);

  // 2. Header Row (prepend our "क्रमांक" header)
  const headerRow = sheet.addRow(['क्रमांक', ...columnTitles]);
  headerRow.font = { bold: true };
  headerRow.alignment = { horizontal: 'center', vertical: 'middle' };

  // 3. Data Rows (each gets a serial number in the first column)
  dataRows.forEach((rowData, index) => {
    sheet.addRow([index + 1, ...rowData]);
  });

  // 4. Set Column Widths:
  //    - The first column ("क्रमांक") gets a fixed narrow width.
  //    - Other columns get a width based on the longest text (header or cell data)
  sheet.columns = [
    { header: 'क्रमांक', key: 'serial', width: 8 },
    ...columnTitles.map((title, i) => {
      let maxLength = title.length;
      dataRows.forEach(row => {
        const cellValue = row[i] ? row[i].toString() : '';
        if (cellValue.length > maxLength) {
          maxLength = cellValue.length;
        }
      });
      return { header: title, key: title, width: maxLength + 5 };
    })
  ];

  // 5. Apply borders and center alignment to all cells
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

  // 6. Generate and download the Excel file
  workbook.xlsx.writeBuffer().then((buffer) => {
    const blob = new Blob([buffer], { 
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${sheetTitle}.xlsx`;
    link.click();
  });
}



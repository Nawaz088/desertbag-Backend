const express = require('express');
const multer = require('multer');
const excel = require('exceljs');
const cors = require('cors');
const path = require('path');
const xlsxToJson = require('xlsx-to-json');

const app = express();
const port = 3001;

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(cors());

app.use('/uploads', express.static(path.join(__dirname, 'uploads')));
app.use('/public', express.static(path.join(__dirname, 'public')));

const storage = multer.diskStorage({
  destination: 'uploads/',
  filename: function (req, file, cb) {
    cb(null, Date.now() + path.extname(file.originalname));
  },
});

const upload = multer({ storage: storage });

app.post('/submit', upload.array('images', 5), (req, res) => {
  const { name, mobile, description } = req.body;

  // Process uploaded image files
  const imagePaths = req.files.map((file) => {
    return path.join(__dirname, 'uploads', file.filename);
  });

  // Update the Excel file with the form data and clickable image paths
  updateExcel(name, mobile, description, imagePaths);

  res.status(200).send('Form submitted successfully');
});

app.get('/view-excel', async (req, res) => {
  try {
    const workbook = new excel.Workbook();
    await workbook.xlsx.readFile('public/data.xlsx');
    const sheet = workbook.getWorksheet('Sheet1');

    // Convert sheet data to JSON
    const jsonData = [];
    sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      const rowData = {};
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        if (colNumber === 4) {
          // Handle the image column (assuming it's at column 4)
          rowData[`Image${rowNumber}`] = cell.value ? cell.value.hyperlink : '';
        } else {
          rowData[`Column${colNumber}`] = cell.value;
        }
      });
      jsonData.push(rowData);
    });

    const htmlTable = renderHtmlTable(jsonData);

    res.send(htmlTable);
  } catch (error) {
    console.error('Error reading Excel file:', error.message);
    res.status(500).send('Internal server error');
  }
});



app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});

function renderHtmlTable(data) {
  const headerRow = `<tr>${Object.keys(data[0]).map((key) => `<th>${key}</th>`).join('')}</tr>`;
  const bodyRows = data.map((row) => `<tr>${Object.values(row).map((value) => `<td>${value}</td>`).join('')}</tr>`);

  return `<table>${headerRow}${bodyRows.join('')}</table>`;
}

// Function to update the Excel file
async function updateExcel(name, mobile, description, imagePaths) {
  const excel = require('exceljs');
  const path = require('path');

  const workbook = new excel.Workbook();
  try {
    await workbook.xlsx.readFile('public/data.xlsx');
    const sheet = workbook.getWorksheet('Sheet1');

    // Add a new row with the provided data
    const newRow = sheet.addRow([name, mobile, description]);

    // Find the starting column index for the images
    let startColumnIndex = 4; // Start at column D for images

    // Set the headers for the image columns
    imagePaths.forEach((imagePath, index) => {
      const columnIndex = startColumnIndex + index;
      sheet.getColumn(columnIndex).header = `Image ${index + 1}`;
      newRow.getCell(columnIndex).value = { text: 'View Image', hyperlink: path.resolve(__dirname, imagePath) };
    });

    await workbook.xlsx.writeFile('public/data.xlsx');
    console.log('Excel file updated successfully');
  } catch (error) {
    console.error('Error updating Excel file:', error.message);
  }
}

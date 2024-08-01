const express = require('express');
const bodyParser = require('body-parser');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;

const EXCEL_FILE = path.join(__dirname, 'data', 'employee_details.xlsx');

app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

const upload = multer();

// Function to ensure the directory exists
function ensureDirectoryExists(filePath) {
  const directory = path.dirname(filePath);
  if (!fs.existsSync(directory)) {
    fs.mkdirSync(directory, { recursive: true });
  }
}

// Function to save data to Excel file
function saveToExcel(data, filePath) {
  const workbook = fs.existsSync(filePath) ? xlsx.readFile(filePath) : xlsx.utils.book_new();
  const worksheet = workbook.Sheets['Sheet1'] || xlsx.utils.aoa_to_sheet([]);
  const sheetData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
  sheetData.push(data);
  const newWorksheet = xlsx.utils.aoa_to_sheet(sheetData);
  workbook.SheetNames = ['Sheet1'];
  workbook.Sheets = { 'Sheet1': newWorksheet };
  xlsx.writeFile(workbook, filePath);
}

// Route to handle form submission
app.post('/submit', upload.none(), (req, res) => {
  const { name, department, dob, email, mobile } = req.body;
  ensureDirectoryExists(EXCEL_FILE);
  saveToExcel([name, department, dob, email, mobile], EXCEL_FILE);
  res.redirect('/');
});

// Route to handle file download
app.get('/download', (req, res) => {
  if (fs.existsSync(EXCEL_FILE)) {
    res.download(EXCEL_FILE, 'employee_details.xlsx');
  } else {
    res.status(404).send('File not found');
  }
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});

const express = require("express");
const bodyParser = require("body-parser");
const multer = require("multer");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const app = express();
const PORT = 3000;

// Middleware
app.use(bodyParser.json());
app.use(express.static("public")); // Serve static files (frontend)
app.use(express.urlencoded({ extended: true }));

// Paths
const FILES_DIR = path.join(__dirname, "files");
const UPDATED_DIR = path.join(__dirname, "updated-files");

// Endpoint to list available files
app.get("/get-files", (req, res) => {
  const files = fs.readdirSync(FILES_DIR).filter(file => file.endsWith(".xlsx"));
  res.json({ files });
});

// Endpoint to get dynamic fields for a selected file
app.get("/get-fields", (req, res) => {
  const fileName = req.query.file;
  const filePath = path.join(FILES_DIR, fileName);

  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: "File not found" });
  }

  // Read Excel file and extract headers
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0]; // Use the first sheet
  const sheet = workbook.Sheets[sheetName];
  const headers = Object.keys(sheet).filter(key => !key.startsWith("!"));

  // Send fields to the frontend
  const fields = headers.map(header => ({
    name: header,
    placeholder: `Enter value for cell ${header}`,
  }));
  res.json({ fields });
});

// Endpoint to process and update Excel file
app.post("/submit", (req, res) => {
  const fileName = req.body.file;
  const filePath = path.join(FILES_DIR, fileName);

  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: "File not found" });
  }

  // Read the Excel file
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  // Update the cells based on user input
  Object.keys(req.body).forEach(key => {
    if (key !== "file") {
      sheet[key] = { v: req.body[key] };
    }
  });

  // Save the updated file
  if (!fs.existsSync(UPDATED_DIR)) {
    fs.mkdirSync(UPDATED_DIR);
}
  const updatedFilePath = path.join(UPDATED_DIR, `updated_${fileName}`);
  XLSX.writeFile(workbook, updatedFilePath);

  // Send the updated file for download
  res.download(updatedFilePath, err => {
    if (err) {
      console.error(err);
    }
  });
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});

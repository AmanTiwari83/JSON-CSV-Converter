const express = require("express");
const multer = require("multer");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

const app = express();

// Configure multer (upload folder)
const upload = multer({ dest: "uploads/" });

// Serve HTML file
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "views", "index.html"));
});


// Convert Excel to JSON (streaming)
app.post("/upload-json", upload.single("excelFile"), async (req, res) => {
  const filePath = req.file.path;
  const outputFile = path.join(__dirname, "output.json");
  const writeStream = fs.createWriteStream(outputFile);

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    const worksheet = workbook.worksheets[0];

    worksheet.eachRow({ includeEmpty: false }, (row) => {
      const rowObject = {};
      row.eachCell((cell, colNumber) => {
        rowObject[`col${colNumber}`] = cell.value;
      });
      writeStream.write(JSON.stringify(rowObject) + "\n");
    });

    writeStream.end(() => {
      res.download(outputFile, "converted.json", () => {
        fs.unlinkSync(filePath);
        fs.unlinkSync(outputFile);
      });
    });
  } catch (err) {
    console.error(err);
    res.status(500).send("Error processing file");
  }
});

// Convert Excel to CSV (streaming, memory efficient)
app.post("/upload-csv", upload.single("excelFile"), async (req, res) => {
  const filePath = req.file.path;
  const outputFile = path.join(__dirname, "output.csv");
  const writeStream = fs.createWriteStream(outputFile);

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.worksheets[0];

    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      const rowValues = row.values.slice(1); // row.values[0] is null
      const csvLine = rowValues
        .map((val) =>
          val === null || val === undefined
            ? ""
            : `"${val.toString().replace(/"/g, '""')}"`
        )
        .join(",");
      writeStream.write(csvLine + "\n");
    });

    writeStream.end(() => {
      res.download(outputFile, "converted.csv", () => {
        fs.unlinkSync(filePath);
        fs.unlinkSync(outputFile);
      });
    });
  } catch (err) {
    console.error(err);
    res.status(500).send("Error processing file");
  }
});

const port = process.env.PORT || 3879;
app.listen(port, () => {
  console.log(`🚀 Server running at http://localhost:${port}`);
});


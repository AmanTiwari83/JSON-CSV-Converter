const express = require("express");
const multer = require("multer");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

const app = express();
const upload = multer({ dest: "uploads/" });

// Serve HTML file
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "views", "index.html"));
});

// Convert Excel → JSON (streaming)
app.post("/upload-json", upload.single("excelFile"), async (req, res) => {
  const filePath = req.file.path;
  const outputFile = path.join(__dirname, "output.json");
  const writeStream = fs.createWriteStream(outputFile);

  try {
    const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(filePath);

    for await (const worksheetReader of workbookReader) {
      for await (const row of worksheetReader) {
        const rowObject = {};
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          rowObject[`col${colNumber}`] = cell.value;
        });
        writeStream.write(JSON.stringify(rowObject) + "\n");
      }
    }

    writeStream.end(() => {
      res.download(outputFile, "converted.json", () => {
        fs.unlinkSync(filePath);
        fs.unlinkSync(outputFile);
      });
    });
  } catch (err) {
    console.error(err);
    res.status(500).send("Error processing JSON");
  }
});

// Convert Excel → CSV (streaming)
app.post("/upload-csv", upload.single("excelFile"), async (req, res) => {
  const filePath = req.file.path;
  const outputFile = path.join(__dirname, "output.csv");
  const writeStream = fs.createWriteStream(outputFile);

  try {
    const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(filePath);

    for await (const worksheetReader of workbookReader) {
      for await (const row of worksheetReader) {
        const rowValues = [];
        row.eachCell({ includeEmpty: true }, (cell) => {
          rowValues.push(
            cell.value === null || cell.value === undefined
              ? ""
              : `"${cell.value.toString().replace(/"/g, '""')}"`
          );
        });
        writeStream.write(rowValues.join(",") + "\n");
      }
    }

    writeStream.end(() => {
      res.download(outputFile, "converted.csv", () => {
        fs.unlinkSync(filePath);
        fs.unlinkSync(outputFile);
      });
    });
  } catch (err) {
    console.error(err);
    res.status(500).send("Error processing CSV");
  }
});

const port = process.env.PORT || 3879;
app.listen(port, () => {
  console.log(`🚀 Server running at http://localhost:${port}`);
});

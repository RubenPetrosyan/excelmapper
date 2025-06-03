// api/processFile.js

import { IncomingForm } from "formidable";
import XLSX from "xlsx";
import fs from "fs";
import path from "path";

/**
 * Disable Next.js’s built-in bodyParser so Formidable can handle multipart/form-data.
 */
export const config = {
  api: {
    bodyParser: false,
  },
};

/**
 * Fixed headers for columns A1 → AO1 in the output workbook.
 * Each string here corresponds to a column, starting at A (index 0).
 */
const OUTPUT_HEADERS = [
  "Veh #",
  "Company vehicle #",
  "Insured ID",
  "Plate #",
  "Year",
  "Make",
  "Model",
  "Body type",
  "Body, if \"Other\"",
  "VIN",
  "Default account address for garaging",
  "Garaging Address 1",
  "Garaging Address 2",
  "Garaging Address 3",
  "Garaging City",
  "Garaging State",
  "Garaging Zip/Postal",
  "Garaging County",
  "Garaging Country",
  "Vehicle type",
  "Symbol/age group",
  "Cost new",
  "Licensed state",
  "Territory",
  "GVW/GCW",
  "Class",
  "Special industry class",
  "Factor Liability",
  "Factor Secondary",
  "Factor Physical damage",
  "Seating capacity",
  "Radius",
  "Farthest terminal",
  "Use",
  "Special use",
  "Days driven per week",
  "Date purchased",
  "Purchased N or U",
  "Leased",
  "Included in fleet",
  "Rate class code"
];

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).send("Method not allowed");
  }

  const form = new IncomingForm({ multiples: false });

  form.parse(req, async (err, fields, files) => {
    if (err) {
      console.error("Formidable parsing error:", err);
      return res.status(500).send("Error parsing uploaded file");
    }

    // 1) Retrieve the uploaded file under field name "file"
    let fileObj = files.file;
    if (!fileObj) {
      console.error("No `files.file` present. Keys were:", Object.keys(files));
      return res.status(400).send("Please upload under the form field name 'file'.");
    }
    if (Array.isArray(fileObj)) {
      fileObj = fileObj[0];
    }

    // 2) Locate the actual temp path on disk where Formidable saved it
    let uploadedPath = null;
    for (const value of Object.values(fileObj)) {
      if (typeof value === "string") {
        try {
          if (fs.existsSync(value)) {
            uploadedPath = value;
            break;
          }
        } catch {
          // ignore non-path values
        }
      }
    }
    if (!uploadedPath) {
      console.error(
        "Could not locate temp path on fileObj. Available keys were:",
        Object.keys(fileObj)
      );
      return res.status(400).send("Could not locate the uploaded file on disk.");
    }

    // 3) Ensure the uploaded file has content
    if (!fileObj.size || fileObj.size === 0) {
      return res.status(400).send("Uploaded file was empty.");
    }

    // 4) Read and parse the workbook
    let workbook;
    try {
      const fileBuffer = fs.readFileSync(uploadedPath);
      workbook = XLSX.read(fileBuffer, { type: "buffer" });
    } catch (xlsxErr) {
      console.error("XLSX.read (buffer) error:", xlsxErr);
      return res.status(400).send("Invalid or unreadable Excel file.");
    }

    // 5) Grab the first worksheet
    const firstSheetName = workbook.SheetNames[0];
    if (!firstSheetName) {
      return res.status(400).send("Uploaded workbook has no sheets.");
    }
    const sheet = workbook.Sheets[firstSheetName];

    // 6) Convert sheet to a 2D array so we can inspect every row
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      blankrows: false
    });
    if (!Array.isArray(rows) || rows.length === 0) {
      return res.status(400).send("The first sheet in the workbook contains no data.");
    }

    // 7) Filter out any row where any cell (lowercased) contains
    //    "year", "vin", "tractor", "truck", "trailer", or "total"
    const dataRows = rows.filter(row => {
      for (const cell of row) {
        const text = String(cell).toLowerCase();
        if (
          text.includes("year") ||
          text.includes("vin") ||
          text.includes("tractor") ||
          text.includes("truck") ||
          text.includes("trailer") ||
          text.includes("total")
        ) {
          return false;
        }
      }
      return true;
    });

    if (dataRows.length === 0) {
      return res.status(400).send("No valid data rows found after filtering out junk rows.");
    }

    // 8) Identify which column is Year/Make/VIN/Cost via pattern heuristics
    //    Sample up to first 5 rows
    const sampleSize = Math.min(5, dataRows.length);
    const columnScores = [];

    // Initialize counters for each column index
    for (let c = 0; c < dataRows[0].length; c++) {
      columnScores[c] = {
        yearCount: 0,
        makeCount: 0,
        vinCount: 0,
        costCount: 0,
        total: 0
      };
    }

    // Sample the first `sampleSize` rows
    for (let r = 0; r < sampleSize; r++) {
      const row = dataRows[r];
      for (let c = 0; c < row.length; c++) {
        const s = String(row[c] || "").trim();

        // Year heuristic: exactly 4 digits between 1900–2100
        if (/^\d{4}$/.test(s)) {
          const y = parseInt(s, 10);
          if (y >= 1900 && y <= 2100) {
            columnScores[c].yearCount++;
          }
        }

        // VIN heuristic: 16+ alphanumeric characters
        if (/^[A-Za-z0-9]{16,}$/.test(s)) {
          columnScores[c].vinCount++;
        }

        // Make heuristic: all letters, at least 2 characters
        if (/^[A-Za-z]{2,}$/.test(s)) {
          columnScores[c].makeCount++;
        }

        // Cost heuristic: money-like ($ or comma/decimal) or pure digits length > 4
        if (
          /^\$?[\d,]+(\.\d+)?$/.test(s) ||
          (/^\d+$/.test(s) && s.length > 4)
        ) {
          columnScores[c].costCount++;
        }

        columnScores[c].total++;
      }
    }

    // Pick best column index for each field
    let yearColIdx = -1, bestYear = 0;
    let makeColIdx = -1, bestMake = 0;
    let vinColIdx  = -1, bestVin = 0;
    let costColIdx = -1, bestCost = 0;

    for (let c = 0; c < columnScores.length; c++) {
      const scores = columnScores[c];
      if (scores.yearCount > bestYear) {
        bestYear = scores.yearCount;
        yearColIdx = c;
      }
      if (scores.makeCount > bestMake) {
        bestMake = scores.makeCount;
        makeColIdx = c;
      }
      if (scores.vinCount > bestVin) {
        bestVin = scores.vinCount;
        vinColIdx = c;
      }
      if (scores.costCount > bestCost) {
        bestCost = scores.costCount;
        costColIdx = c;
      }
    }

    if (bestYear < 2) {
      return res.status(400).send("Cannot reliably detect the Year column.");
    }
    if (bestMake < 2) {
      return res.status(400).send("Cannot reliably detect the Make column.");
    }
    if (bestVin < 1) {
      return res.status(400).send("Cannot reliably detect the VIN column.");
    }
    if (bestCost < 1) {
      return res.status(400).send("Cannot reliably detect the Cost column.");
    }

    // 9) Build a new worksheet and write fixed headers A1→AO1
    const newSheet = {};
    OUTPUT_HEADERS.forEach((headerText, colIndex) => {
      const address = XLSX.utils.encode_cell({ r: 0, c: colIndex });
      newSheet[address] = { v: headerText };
    });

    // 10) Iterate all dataRows, extract and clean fields, and write them
    let outputRowCount = 0;
    for (let r = 0; r < dataRows.length; r++) {
      const row = dataRows[r];

      const rawYear = row[yearColIdx] || "";
      const rawMake = row[makeColIdx] || "";
      const rawVin  = row[vinColIdx]  || "";
      const rawCost = row[costColIdx] || "";

      const yearStr = String(rawYear).trim();
      const makeStr = String(rawMake).trim();
      const vinStr  = String(rawVin).trim();
      const costStr = String(rawCost).trim();

      // Skip if all four fields are blank
      if (!yearStr && !makeStr && !vinStr && !costStr) {
        continue;
      }

      // Clean cost by removing all non-digits (e.g. "$20,000.00" → "20000")
      const costVal = costStr.replace(/\D/g, "");

      outputRowCount++;
      const outRow = outputRowCount + 1; // because row=1 (r=0) is header

      // Year → column E (index 4)
      newSheet[`E${outRow}`] = { v: yearStr };
      // Make → column F (index 5)
      newSheet[`F${outRow}`] = { v: makeStr };
      // VIN → column J (index 9)
      newSheet[`J${outRow}`] = { v: vinStr };
      // Cost new → column V (index 21)
      newSheet[`V${outRow}`] = { v: costVal };
    }

    // 11) If nothing was written, return an error
    if (outputRowCount === 0) {
      return res
        .status(400)
        .send("No valid data rows (Year/Make/VIN/Cost) were found after filtering.");
    }

    // 12) Define the worksheet range (“!ref”) from A1 → AO(lastDataRow)
    const lastRowIndex = outputRowCount;            // data rows occupy 1..outputRowCount
    const lastColIndex = OUTPUT_HEADERS.length - 1; // 40 (A=0…AO=40)
    newSheet["!ref"] = XLSX.utils.encode_range(
      { r: 0, c: 0 },
      { r: lastRowIndex, c: lastColIndex }
    );

    // 13) Create a new workbook, append this sheet, and write it to /tmp
    const newBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newBook, newSheet, "Standardized");

    const timestamp = Date.now();
    const tempPath = path.join("/tmp", `processed-${timestamp}.xlsx`);
    try {
      XLSX.writeFile(newBook, tempPath);
    } catch (writeErr) {
      console.error("Error writing new workbook:", writeErr);
      return res.status(500).send("Failed to create the processed file.");
    }

    // 14) Verify the /tmp file exists and is non-empty
    let stats;
    try {
      stats = fs.statSync(tempPath);
    } catch (statErr) {
      console.error("Temp-file check error:", statErr);
      return res.status(500).send("Processed file was not created correctly.");
    }
    if (stats.size === 0) {
      console.error("Temp file is empty after writing");
      return res.status(500).send("Processed file is empty.");
    }

    // 15) Read the buffer and send it back to the client
    try {
      const fileBuffer = fs.readFileSync(tempPath);
      const originalName = fileObj.originalFilename || `upload-${timestamp}.xlsx`;
      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );
      res.setHeader(
        "Content-Disposition",
        `attachment; filename=Processed_${path.basename(originalName)}`
      );
      return res.send(fileBuffer);
    } catch (sendErr) {
      console.error("Error sending file buffer:", sendErr);
      return res.status(500).send("Failed to send processed file.");
    }
  });
}

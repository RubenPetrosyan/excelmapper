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
 * These are the exact headers for columns A1 → AO1 in the output workbook.
 * Each string in this array corresponds to a column, starting at A (index 0).
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

    // 1) Grab the uploaded file under field name "file"
    let fileObj = files.file;
    if (!fileObj) {
      console.error("No `files.file` present. Keys were:", Object.keys(files));
      return res
        .status(400)
        .send("Please upload under the form field name 'file'.");
    }

    // 2) If Formidable returned an array (because <input multiple>), pick the first element
    if (Array.isArray(fileObj)) {
      fileObj = fileObj[0];
    }

    // 3) Attempt to find the real temp path on disk
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
        "Could not locate a temp path on fileObj. Available keys were:",
        Object.keys(fileObj)
      );
      return res
        .status(400)
        .send(
          "Could not locate the uploaded file on disk. Check Function Logs for details."
        );
    }

    // 4) Ensure the uploaded file has content
    if (!fileObj.size || fileObj.size === 0) {
      return res.status(400).send("Uploaded file was empty.");
    }

    // 5) Read the uploaded file into a Buffer and parse as XLSX
    let workbook;
    try {
      const fileBuffer = fs.readFileSync(uploadedPath);
      workbook = XLSX.read(fileBuffer, { type: "buffer" });
    } catch (xlsxErr) {
      console.error("XLSX.read (buffer) error:", xlsxErr);
      return res.status(400).send("Invalid or unreadable Excel file.");
    }

    // 6) Take the first worksheet
    const firstSheetName = workbook.SheetNames[0];
    if (!firstSheetName) {
      return res.status(400).send("Uploaded workbook has no sheets.");
    }
    const sheet = workbook.Sheets[firstSheetName];

    // 7) Convert the sheet into a 2D array (header:1) so we can detect the actual header row
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      blankrows: false
    });
    if (!Array.isArray(rows) || rows.length === 0) {
      return res
        .status(400)
        .send("The first sheet in the workbook contains no data.");
    }

    // 8) Find the single header row index by searching for a row containing
    //    "year", "make", "vin", and ("cost" or "stated") (case-insensitive)
    let headerRowIdx = -1;
    for (let r = 0; r < rows.length; r++) {
      const rowLower = rows[r].map(cell =>
        typeof cell === "string" ? cell.toLowerCase() : ""
      );
      const hasYear = rowLower.some(c => c.includes("year"));
      const hasMake = rowLower.some(c => c.includes("make"));
      const hasVin = rowLower.some(c => c.includes("vin"));
      const hasCostOrStated =
        rowLower.some(c => c.includes("cost")) ||
        rowLower.some(c => c.includes("stated"));
      if (hasYear && hasMake && hasVin && hasCostOrStated) {
        headerRowIdx = r;
        break;
      }
    }

    if (headerRowIdx < 0) {
      return res
        .status(400)
        .send(
          "Input sheet is missing a header row containing Year, Make, VIN, and Cost or Stated Value (case-insensitive)."
        );
    }

    // 9) All rows after headerRowIdx are data rows
    const dataRows = rows.slice(headerRowIdx + 1);
    if (dataRows.length === 0) {
      return res.status(400).send("No data found after the header row.");
    }

    // 10) From the first few dataRows, compute heuristics to identify column indexes
    const sampleSize = Math.min(5, dataRows.length);
    const columnScores = [];

    // Initialize counters for each column
    for (let c = 0; c < dataRows[0].length; c++) {
      columnScores[c] = {
        yearCount: 0,
        makeCount: 0,
        vinCount: 0,
        costCount: 0,
        total: 0
      };
    }

    // Sample up to sampleSize rows
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

        // Cost heuristic: contains $ or comma or decimal, or is digits length > 4
        if (
          /^\$?[\d,]+(\.\d+)?$/.test(s) ||
          (/^\d+$/.test(s) && s.length > 4)
        ) {
          columnScores[c].costCount++;
        }

        columnScores[c].total++;
      }
    }

    // 11) Select column index with highest score for each field (with minimal thresholds)
    let yearColIdx = -1;
    let bestYear = 0;
    let makeColIdx = -1;
    let bestMake = 0;
    let vinColIdx = -1;
    let bestVin = 0;
    let costColIdx = -1;
    let bestCost = 0;

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
      return res.status(400).send("Cannot reliably find the Year column.");
    }
    if (bestMake < 2) {
      return res.status(400).send("Cannot reliably find the Make column.");
    }
    if (bestVin < 1) {
      return res.status(400).send("Cannot reliably find the VIN column.");
    }
    if (bestCost < 1) {
      return res.status(400).send("Cannot reliably find the Cost column.");
    }

    // 12) Build a new worksheet and write fixed headers A1→AO1
    const newSheet = {};
    OUTPUT_HEADERS.forEach((headerText, colIndex) => {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: colIndex });
      newSheet[cellAddress] = { v: headerText };
    });

    // 13) Loop over all dataRows, extract and clean fields, and write to output
    let outputRowCount = 0;
    for (let r = 0; r < dataRows.length; r++) {
      const row = dataRows[r];
      const rawYear = row[yearColIdx] || "";
      const rawMake = row[makeColIdx] || "";
      const rawVin = row[vinColIdx] || "";
      const rawCost = row[costColIdx] || "";

      // Skip entirely blank rows
      if (
        String(rawYear).trim() === "" &&
        String(rawMake).trim() === "" &&
        String(rawVin).trim() === "" &&
        String(rawCost).trim() === ""
      ) {
        continue;
      }

      outputRowCount++;
      const outRow = outputRowCount + 1; // header is at row 1 (r=0)

      const yearVal = String(rawYear).trim();
      const makeVal = String(rawMake).trim();
      const vinVal = String(rawVin).trim();
      const costVal = String(rawCost).replace(/\D/g, ""); // digits only

      newSheet[`E${outRow}`] = { v: yearVal };
      newSheet[`F${outRow}`] = { v: makeVal };
      newSheet[`J${outRow}`] = { v: vinVal };
      newSheet[`V${outRow}`] = { v: costVal };
    }

    // 14) If no rows were written, error out
    if (outputRowCount === 0) {
      return res
        .status(400)
        .send("No valid data rows (Year/Make/VIN/Cost) were found in the file.");
    }

    // 15) Define the sheet range "!ref" from A1 to AO(lastDataRow)
    const lastRowIndex = outputRowCount; // data runs 1..outputRowCount
    const lastColIndex = OUTPUT_HEADERS.length - 1; // 40
    newSheet["!ref"] = XLSX.utils.encode_range(
      { r: 0, c: 0 },
      { r: lastRowIndex, c: lastColIndex }
    );

    // 16) Create a new workbook, append this sheet, and write it to /tmp
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

    // 17) Verify the /tmp file exists and is non-empty
    let stats;
    try {
      stats = fs.statSync(tempPath);
    } catch (statErr) {
      console.error("Temp-file check error:", statErr);
      return res
        .status(500)
        .send("Processed file was not created correctly.");
    }
    if (stats.size === 0) {
      console.error("Temp file is empty after writing");
      return res.status(500).send("Processed file is empty.");
    }

    // 18) Read the file buffer and send it back to the client
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

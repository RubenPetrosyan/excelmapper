// api/processFile.js

import { IncomingForm } from "formidable";
import XLSX from "xlsx";
import fs from "fs";
import path from "path";

/**
 * Disable Next.js’s built‐in bodyParser so Formidable can handle multipart/form‐data.
 */
export const config = {
  api: {
    bodyParser: false,
  },
};

/**
 * These are the exact headers for columns A1 → AO1 in the output workbook.
 * Each string in this array corresponds to a column, starting at A.
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

    // 8) Find the header row index by searching for a row containing "year", "make", "vin", and "cost"
    let headerRowIdx = -1;
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r].map((cell) =>
        typeof cell === "string" ? cell.toLowerCase() : ""
      );
      if (
        row.some((c) => c.includes("year")) &&
        row.some((c) => c.includes("make")) &&
        row.some((c) => c.includes("vin")) &&
        row.some((c) => c.includes("cost"))
      ) {
        headerRowIdx = r;
        break;
      }
    }

    if (headerRowIdx < 0) {
      return res
        .status(400)
        .send(
          "Input sheet is missing a header row containing Year, Make, VIN, and Cost (case-insensitive)."
        );
    }

    // 9) Extract the actual header row (normalized) and find the column indexes
    const inputHeaderRow = rows[headerRowIdx].map((h) =>
      typeof h === "string" ? h.toLowerCase() : ""
    );
    let yearColIdx = -1;
    let makeColIdx = -1;
    let vinColIdx = -1;
    let costColIdx = -1;

    inputHeaderRow.forEach((cellText, idx) => {
      if (cellText.includes("year") && yearColIdx === -1) {
        yearColIdx = idx;
      }
      if (cellText.includes("make") && makeColIdx === -1) {
        makeColIdx = idx;
      }
      if (cellText.includes("vin") && vinColIdx === -1) {
        vinColIdx = idx;
      }
      if (cellText.includes("cost") && costColIdx === -1) {
        costColIdx = idx;
      }
    });

    if (
      yearColIdx < 0 ||
      makeColIdx < 0 ||
      vinColIdx < 0 ||
      costColIdx < 0
    ) {
      return res
        .status(400)
        .send(
          "Unable to detect all required columns (Year, Make, VIN, Cost) in the header row."
        );
    }

    // 10) Build a new worksheet object and place the fixed headers into A1→AO1
    const newSheet = {};
    OUTPUT_HEADERS.forEach((headerText, colIndex) => {
      const address = XLSX.utils.encode_cell({ r: 0, c: colIndex });
      newSheet[address] = { v: headerText };
    });

    // 11) For each data row after the headerRowIdx, extract Year/Make/VIN/Cost and place into output
    let outputRowCount = 0;
    for (let i = headerRowIdx + 1; i < rows.length; i++) {
      const inputRow = rows[i];
      // Stop if we hit an empty row or a different section (e.g. "TRAILERS")
      // We check: if all four key columns are empty, skip
      const yearCell = (inputRow[yearColIdx] || "").toString().trim();
      const makeCell = (inputRow[makeColIdx] || "").toString().trim();
      const vinCell = (inputRow[vinColIdx] || "").toString().trim();
      const costCell = (inputRow[costColIdx] || "").toString().trim();
      if (!yearCell && !makeCell && !vinCell && !costCell) {
        continue;
      }

      // Parse and clean each field
      const yearVal = yearCell;
      const makeVal = makeCell;
      const vinVal = vinCell;
      // Remove any non-digit characters from cost (e.g. "$20,000" → "20000")
      const cleanedCost = costCell.replace(/\D/g, "");

      outputRowCount++;
      const outputRowNumber = outputRowCount + 1; // row 2,3,4...

      // Place Year → column E (index 4)
      newSheet[`E${outputRowNumber}`] = { v: yearVal };

      // Place Make → column F (index 5)
      newSheet[`F${outputRowNumber}`] = { v: makeVal };

      // Place VIN → column J (index 9)
      newSheet[`J${outputRowNumber}`] = { v: vinVal };

      // Place Cost new → column V (index 21)
      newSheet[`V${outputRowNumber}`] = { v: cleanedCost };
    }

    // If no data rows were processed, return an error
    if (outputRowCount === 0) {
      return res
        .status(400)
        .send("No data rows with Year, Make, VIN, and Cost were found.");
    }

    // 12) Define the sheet range from A1 to AO(lastDataRow)
    const lastRowIndex = outputRowCount; // since header is at row 1 (r=0), data rows 2..(outputRowCount+1)
    const lastColIndex = OUTPUT_HEADERS.length - 1; // 40 (A=0, …, AO=40)
    const startCell = { r: 0, c: 0 };
    const endCell = { r: lastRowIndex, c: lastColIndex };
    newSheet["!ref"] = XLSX.utils.encode_range(startCell, endCell);

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
      return res
        .status(500)
        .send("Processed file was not created correctly.");
    }
    if (stats.size === 0) {
      console.error("Temp file is empty after writing");
      return res.status(500).send("Processed file is empty.");
    }

    // 15) Read the file buffer and send it back to the client
    try {
      const fileBuffer = fs.readFileSync(tempPath);
      const originalName =
        fileObj.originalFilename || `upload-${timestamp}.xlsx`;
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

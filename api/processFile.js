// api/processFile.js

import { IncomingForm } from "formidable";
import XLSX from "xlsx";
import fs from "fs";
import path from "path";

export const config = {
  api: {
    bodyParser: false,
  },
};

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).send("Method not allowed");
  }

  const form = new IncomingForm({ multiples: false });
  form.parse(req, async (err, fields, files) => {
    if (err) {
      console.error("Formidable error:", err);
      return res.status(500).send("Error parsing uploaded file");
    }

    const file = files.file;
    if (!file) {
      return res
        .status(400)
        .send("Please upload under the form field name 'file'.");
    }

    // *** REPLACE `YOUR_KEY_HERE` WITH THE ACTUAL PROPERTY YOU SAW IN THE LOGS ***
    const uploadedPath = file.YOUR_KEY_HERE;

    console.log("Resolved uploadedPath:", uploadedPath);
    console.log("Incoming file metadata:", {
      originalFilename: file.originalFilename,
      size: file.size,
      mimeType: file.mimetype || "n/a",
    });

    if (!uploadedPath) {
      console.error(
        "Still no temp path found. Available keys on `file` were:",
        Object.keys(file)
      );
      return res
        .status(400)
        .send(
          "Could not locate the uploaded file on disk. Check Function Logs for the full 'file' object."
        );
    }

    if (file.size === 0) {
      return res.status(400).send("Uploaded file was empty.");
    }

    // Read the file buffer and parse as XLSX
    let workbook;
    try {
      const fileBuffer = fs.readFileSync(uploadedPath);
      workbook = XLSX.read(fileBuffer, { type: "buffer" });
    } catch (xlsxErr) {
      console.error("XLSX.read (buffer) error:", xlsxErr);
      return res.status(400).send("Invalid or unreadable Excel file.");
    }

    // (…then continue with your generic “find ≥3 non-empty cells per row” logic…)
    // Build a 2D array of all rows:
    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      blankrows: false,
    });

    // Collect any row with ≥3 non-empty cells:
    const output = [];
    rows.forEach((rowArr) => {
      const nonEmpty = rowArr.filter(
        (cell) => cell !== null && cell !== undefined && cell !== ""
      );
      if (nonEmpty.length >= 3) {
        const [makeVal, yearVal, vinVal] = nonEmpty;
        output.push({
          Make: makeVal,
          Year: yearVal,
          "VIN Number": vinVal,
        });
      }
    });

    if (output.length === 0) {
      return res
        .status(400)
        .send("No rows with at least three non-empty cells were found.");
    }

    // Create a new workbook with just those 3 columns
    const newSheet = XLSX.utils.json_to_sheet(output, {
      header: ["Make", "Year", "VIN Number"],
    });
    const newBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newBook, newSheet, "Standardized");

    // Write to /tmp in Vercel
    const timestamp = Date.now();
    const tempPath = path.join("/tmp", `processed-${timestamp}.xlsx`);
    try {
      XLSX.writeFile(newBook, tempPath);
    } catch (writeErr) {
      console.error("Error writing new workbook:", writeErr);
      return res.status(500).send("Failed to create the processed file.");
    }

    // Ensure the new file exists and is non-empty
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

    // Stream the result back to the client
    try {
      const fileBuffer = fs.readFileSync(tempPath);
      const originalName = file.originalFilename || `upload-${timestamp}.xlsx`;
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

// api/processFile.js

import { IncomingForm } from "formidable";
import XLSX from "xlsx";
import fs from "fs";
import path from "path";

// Disable Next.js body parsing so Formidable can handle multipart/form-data
export const config = {
  api: {
    bodyParser: false,
  },
};

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).send("Method not allowed");
  }

  // Set up Formidable to parse exactly one file under field name "file"
  const form = new IncomingForm({ multiples: false });

  form.parse(req, async (err, fields, files) => {
    if (err) {
      console.error("Formidable error:", err);
      return res.status(500).send("Error parsing uploaded file");
    }

    // Ensure there is a file under the exact field name "file"
    const file = files.file;
    if (!file) {
      return res
        .status(400)
        .send("Please upload under the form field name 'file'.");
    }
    if (file.size === 0) {
      return res.status(400).send("Uploaded file was empty.");
    }

    // Try reading the uploaded .xlsx (or .xls) from its temp filepath
    let workbook;
    try {
      workbook = XLSX.readFile(file.filepath);
    } catch (xlsxErr) {
      console.error("XLSX.readFile error:", xlsxErr);
      return res.status(400).send("Invalid or unreadable Excel file.");
    }

    // Take the very first worksheet
    const firstSheetName = workbook.SheetNames[0];
    if (!firstSheetName) {
      return res.status(400).send("Uploaded workbook has no sheets.");
    }
    const sheet = workbook.Sheets[firstSheetName];

    // Convert the entire sheet into a 2D array (header:1) so we can see every row/column
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      blankrows: false,
    });
    if (!Array.isArray(rows) || rows.length === 0) {
      return res
        .status(400)
        .send("The first sheet in the workbook contains no data.");
    }

    // Build an array of output objects: each must have Make, Year, VIN Number
    const output = [];
    rows.forEach((rowArr) => {
      // Collect all non-empty cells in this row (in left-to-right order)
      const nonEmpty = rowArr.filter((cell) => {
        return cell !== null && cell !== undefined && cell !== "";
      });

      // If this row has at least three non-empty values, assume they map to [Make, Year, VIN]
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

    // Create a new workbook with a single sheet named “Standardized”
    const newSheet = XLSX.utils.json_to_sheet(output, {
      header: ["Make", "Year", "VIN Number"],
    });
    const newBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newBook, newSheet, "Standardized");

    // Write that workbook to /tmp so Vercel can read it back
    const timestamp = Date.now();
    const tempPath = path.join("/tmp", `processed-${timestamp}.xlsx`);
    try {
      XLSX.writeFile(newBook, tempPath);
    } catch (writeErr) {
      console.error("Error writing new workbook:", writeErr);
      return res.status(500).send("Failed to create the processed file.");
    }

    // Verify that the file was actually written and is non-empty
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

    // Read the result and stream it back
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

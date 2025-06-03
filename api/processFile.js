// api/processFile.js

import { IncomingForm } from "formidable";
import XLSX from "xlsx";
import fs from "fs";
import path from "path";

// Disable Next.js’s built-in bodyParser so we can let Formidable handle the multipart upload.
export const config = {
  api: {
    bodyParser: false,
  },
};

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).send("Method not allowed");
  }

  // Create a new IncomingForm; we only expect one file under the field name "file"
  const form = new IncomingForm({ multiples: false });

  form.parse(req, async (err, fields, files) => {
    if (err) {
      console.error("Formidable error:", err);
      return res.status(500).send("Error parsing uploaded file");
    }

    // Look for the uploaded file under the name “file”
    const file = files.file;
    if (!file) {
      return res
        .status(400)
        .send("Please upload under the form field name 'file'.");
    }

    // *** LOG THE ENTIRE file OBJECT to see which property actually holds the temp path. ***
    console.log("Full file object from Formidable:", file);

    // Formidable v3+ uses `file.filepath`; older versions use `file.path`
    const uploadedPath = file.filepath || file.path;
    if (!uploadedPath) {
      console.error(
        "Unable to find a temp path. Did Formidable save this file under an unexpected property?"
      );
      return res
        .status(400)
        .send("Could not locate the uploaded file on disk.");
    }

    console.log("Resolved temp path:", uploadedPath);
    console.log("Incoming file metadata:", {
      originalFilename: file.originalFilename,
      size: file.size,
      mimeType: file.mimetype || "n/a",
    });

    if (file.size === 0) {
      return res.status(400).send("Uploaded file was empty.");
    }

    // Attempt to read the uploaded file into a Buffer, then parse as XLSX
    let workbook;
    try {
      const fileBuffer = fs.readFileSync(uploadedPath);
      workbook = XLSX.read(fileBuffer, { type: "buffer" });
    } catch (xlsxErr) {
      console.error("XLSX.read (buffer) error:", xlsxErr);
      return res.status(400).send("Invalid or unreadable Excel file.");
    }

    // Take the very first sheet in the workbook
    const firstSheetName = workbook.SheetNames[0];
    if (!firstSheetName) {
      return res.status(400).send("Uploaded workbook has no sheets.");
    }
    const sheet = workbook.Sheets[firstSheetName];

    // Convert the entire sheet to a 2D array so we can find data anywhere
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

    // Build the output array: every row with ≥3 non-empty cells → [Make, Year, VIN Number]
    const output = [];
    rows.forEach((rowArr) => {
      // Filter out truly blank cells
      const nonEmpty = rowArr.filter(
        (cell) => cell !== null && cell !== undefined && cell !== ""
      );

      // If at least three real values remain, assume they map to [Make, Year, VIN]
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

    // Create a new workbook with those three columns (Make, Year, VIN Number)
    const newSheet = XLSX.utils.json_to_sheet(output, {
      header: ["Make", "Year", "VIN Number"],
    });
    const newBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newBook, newSheet, "Standardized");

    // Write to /tmp so Vercel can read it back
    const timestamp = Date.now();
    const tempPath = path.join("/tmp", `processed-${timestamp}.xlsx`);
    try {
      XLSX.writeFile(newBook, tempPath);
    } catch (writeErr) {
      console.error("Error writing new workbook:", writeErr);
      return res.status(500).send("Failed to create the processed file.");
    }

    // Verify that the temp file was created and is non-empty
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

    // Finally, read the buffer and send it back to the client
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

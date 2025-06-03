// api/processFile.js

import { IncomingForm } from "formidable";
import XLSX from "xlsx";
import fs from "fs";
import path from "path";

/**
 * We disable Next.js’s built‐in bodyParser so Formidable can handle multipart/form‐data.
 */
export const config = {
  api: {
    bodyParser: false,
  },
};

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).send("Method not allowed");
  }

  // 1) Create a new Formidable instance that expects exactly one file under field name "file"
  const form = new IncomingForm({ multiples: false });

  form.parse(req, async (err, fields, files) => {
    if (err) {
      console.error("Formidable parsing error:", err);
      return res.status(500).send("Error parsing uploaded file");
    }

    // 2) Ensure a file arrived under files.file
    const file = files.file;
    let fileObj = files.file;
if (Array.isArray(fileObj)) {
  // Formidable gave us an array; pick the first file
  fileObj = fileObj[0];
}

    if (!file) {
      console.error("No `files.file` present. Keys were:", Object.keys(files));
      return res
        .status(400)
        .send("Please upload under the form field name 'file'.");
    }

    // 3) Log the full file object so you can inspect it in the Vercel logs if needed
    console.log("Full file object from Formidable:", file);

    // 4) We now scan *all* string‐valued properties on `file` looking for one
    //    that points to an existing temp file. In most environments, Formidable
    //    has either `file.filepath`, `file.path`, or another string that begins
    //    with `/tmp` or `/vercel/tmp`. We’ll pick the first value that `fs.existsSync` returns true for.
    let uploadedPath = null;
    for (const candidateValue of Object.values(file)) {
      if (typeof candidateValue === "string") {
        try {
          // If the path actually exists on disk, we assume it’s our temp file.
          if (fs.existsSync(candidateValue)) {
            uploadedPath = candidateValue;
            break;
          }
        } catch (_) {
          // ignore any fs errors for non‐paths
        }
      }
    }

    console.log("Resolved uploadedPath:", uploadedPath);

    if (!uploadedPath) {
      console.error(
        "Could not locate a temp path on `file`. Available keys were:",
        Object.keys(file)
      );
      return res
        .status(400)
        .send(
          "Could not locate the uploaded file on disk. Check Function Logs for the full 'file' object."
        );
    }

    // 5) Confirm we actually received bytes
    console.log("Incoming file metadata:", {
      originalFilename: file.originalFilename,
      size: file.size,
      mimeType: file.mimetype || "n/a",
    });
    if (file.size === 0) {
      return res.status(400).send("Uploaded file was empty.");
    }

    // 6) Read the file from disk into a Buffer, then parse with XLSX
    let workbook;
    try {
      const fileBuffer = fs.readFileSync(uploadedPath);
      workbook = XLSX.read(fileBuffer, { type: "buffer" });
    } catch (xlsxErr) {
      console.error("XLSX.read (buffer) error:", xlsxErr);
      return res.status(400).send("Invalid or unreadable Excel file.");
    }

    // 7) Grab the very first sheet
    const firstSheetName = workbook.SheetNames[0];
    if (!firstSheetName) {
      return res.status(400).send("Uploaded workbook has no sheets.");
    }
    const sheet = workbook.Sheets[firstSheetName];

    // 8) Convert the entire sheet to a 2D array so we can find data in ANY column
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

    // 9) Build the output: each row that has ≥3 non-empty cells (in left-to-right order)
    //    becomes an object { Make, Year, VIN Number }.
    const output = [];
    rows.forEach((rowArr) => {
      // Filter out truly blank cells
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

    // 10) Create a new workbook with exactly those three columns (Make, Year, VIN Number)
    const newSheet = XLSX.utils.json_to_sheet(output, {
      header: ["Make", "Year", "VIN Number"],
    });
    const newBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newBook, newSheet, "Standardized");

    // 11) Write that new workbook to a file in /tmp (the only writable dir on Vercel)
    const timestamp = Date.now();
    const tempPath = path.join("/tmp", `processed-${timestamp}.xlsx`);
    try {
      XLSX.writeFile(newBook, tempPath);
    } catch (writeErr) {
      console.error("Error writing new workbook:", writeErr);
      return res.status(500).send("Failed to create the processed file.");
    }

    // 12) Verify that the file exists and is non-empty
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

    // 13) Finally, read the buffer and send it back to the client
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

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

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).send("Method not allowed");
  }

  // 1) Create a new Formidable instance that expects at most one file under field name "file"
  const form = new IncomingForm({ multiples: false });

  form.parse(req, async (err, fields, files) => {
    if (err) {
      console.error("Formidable parsing error:", err);
      return res.status(500).send("Error parsing uploaded file");
    }

    // 2) Grab whatever Formidable placed under files.file
    let fileObj = files.file;
    if (!fileObj) {
      console.error("No `files.file` present. Keys were:", Object.keys(files));
      return res
        .status(400)
        .send("Please upload under the form field name 'file'.");
    }

    // 3) If Formidable returned an array (because the client used `multiple`), pick the first element
    if (Array.isArray(fileObj)) {
      fileObj = fileObj[0];
    }

    // 4) Log the full file object so you can inspect its shape in Vercel logs if needed
    console.log("Full file object from Formidable:", fileObj);

    // 5) Scan all string‐valued properties of fileObj to find a real path on disk.
    //    Formidable v3+ uses `fileObj.filepath`, older versions use `fileObj.path`,
    //    and occasionally it might be under some other string property that begins with "/tmp".
    let uploadedPath = null;
    for (const candidateValue of Object.values(fileObj)) {
      if (typeof candidateValue === "string") {
        try {
          if (fs.existsSync(candidateValue)) {
            uploadedPath = candidateValue;
            break;
          }
        } catch (_) {
          // ignore errors for non‐path values
        }
      }
    }

    console.log("Resolved uploadedPath:", uploadedPath);

    if (!uploadedPath) {
      console.error(
        "Could not locate a temp path on fileObj. Available keys were:",
        Object.keys(fileObj)
      );
      return res
        .status(400)
        .send(
          "Could not locate the uploaded file on disk. Check Function Logs for the full 'file' object."
        );
    }

    // 6) Confirm the client actually sent bytes
    console.log("Incoming file metadata:", {
      originalFilename: fileObj.originalFilename,
      size: fileObj.size,
      mimeType: fileObj.mimetype || "n/a",
    });
    if (fileObj.size === 0) {
      return res.status(400).send("Uploaded file was empty.");
    }

    // 7) Read the uploaded file into a Buffer, then parse as XLSX
    let workbook;
    try {
      const fileBuffer = fs.readFileSync(uploadedPath);
      workbook = XLSX.read(fileBuffer, { type: "buffer" });
    } catch (xlsxErr) {
      console.error("XLSX.read (buffer) error:", xlsxErr);
      return res.status(400).send("Invalid or unreadable Excel file.");
    }

    // 8) Take the very first sheet from the workbook
    const firstSheetName = workbook.SheetNames[0];
    if (!firstSheetName) {
      return res.status(400).send("Uploaded workbook has no sheets.");
    }
    const sheet = workbook.Sheets[firstSheetName];

    // 9) Convert the entire sheet into a 2D array so we can locate data in ANY column
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

    // 10) Build the output rows: any row with ≥3 non‐empty cells → {Make, Year, VIN Number}
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

    // 11) Create a new workbook with exactly those three columns
    const newSheet = XLSX.utils.json_to_sheet(output, {
      header: ["Make", "Year", "VIN Number"],
    });
    const newBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newBook, newSheet, "Standardized");

    // 12) Write the new workbook to /tmp (the only writable directory on Vercel)
    const timestamp = Date.now();
    const tempPath = path.join("/tmp", `processed-${timestamp}.xlsx`);
    try {
      XLSX.writeFile(newBook, tempPath);
    } catch (writeErr) {
      console.error("Error writing new workbook:", writeErr);
      return res.status(500).send("Failed to create the processed file.");
    }

    // 13) Verify the /tmp file exists and is non‐empty
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

    // 14) Finally, read the buffer and send it back to the client
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

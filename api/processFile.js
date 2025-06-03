// api/processFile.js
import formidable from "formidable";
import XLSX from "xlsx";
import fs from "fs";
import path from "path";

export const config = { api: { bodyParser: false } };

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).send("Method not allowed");
  }

  const form = new formidable.IncomingForm({ multiples: false });

  form.parse(req, async (err, fields, files) => {
    if (err) {
      console.error("Form parsing error:", err);
      return res.status(500).send("Error parsing uploaded file");
    }

    console.log("Parsed fields:", fields);
    console.log("Parsed files:", files);

    const file = files.file;
    if (!file) {
      console.error("No file.part named 'file' found.");
      return res
        .status(400)
        .send("No file uploaded under the field name 'file'");
    }

    console.log("Incoming file metadata:", {
      originalFilename: file.originalFilename,
      size: file.size,
      path: file.filepath,
    });

    // If the uploaded file is empty:
    if (file.size === 0) {
      console.error("Uploaded file is zero bytes.");
      return res.status(400).send("Uploaded file was empty");
    }

    // Now try reading it as an Excel workbook:
    let workbook;
    try {
      workbook = XLSX.readFile(file.filepath);
    } catch (xlsxErr) {
      console.error("XLSX.readFile error:", xlsxErr);
      return res.status(500).send("Invalid or unreadable Excel file");
    }

    const firstSheetName = workbook.SheetNames[0];
    if (!firstSheetName) {
      console.error("No sheets found in uploaded workbook");
      return res.status(500).send("Uploaded workbook has no sheets");
    }

    const sheetData = XLSX.utils.sheet_to_json(
      workbook.Sheets[firstSheetName],
      { defval: "" }
    );

    // Transform rows
    let output;
    try {
      output = sheetData.map((row) => {
        const newRow = {};
        for (const key in row) {
          const lowerKey = key.toLowerCase().trim();
          if (lowerKey.includes("make")) newRow["Make"] = row[key];
          if (lowerKey.includes("year")) newRow["Year"] = row[key];
          if (lowerKey.includes("vin")) newRow["VIN Number"] = row[key];
        }
        return newRow;
      });
    } catch (transformErr) {
      console.error("Error transforming sheet data:", transformErr);
      return res.status(500).send("Error processing sheet data");
    }

    // Convert back to a new workbook
    let tempPath;
    try {
      const newSheet = XLSX.utils.json_to_sheet(output);
      const newBook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(newBook, newSheet, "Standardized");

      tempPath = path.join("/tmp", `processed-${Date.now()}.xlsx`);
      XLSX.writeFile(newBook, tempPath);
      console.log("Wrote new workbook to", tempPath);
    } catch (writeErr) {
      console.error("Error writing new workbook:", writeErr);
      return res.status(500).send("Failed to generate processed file");
    }

    // Verify temp file exists and is nonempty
    let stats;
    try {
      stats = fs.statSync(tempPath);
      console.log("Temp file size:", stats.size);
    } catch (statErr) {
      console.error("Temp file not found or unreadable:", statErr);
      return res.status(500).send("Processed file was not created correctly");
    }
    if (stats.size === 0) {
      console.error("Temp file is empty after writing");
      return res.status(500).send("Processed file is empty");
    }

    // Send the file back
    try {
      const fileBuffer = fs.readFileSync(tempPath);
      const originalName = file.originalFilename || `upload-${Date.now()}.xlsx`;
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
      return res.status(500).send("Failed to send processed file");
    }
  });
}

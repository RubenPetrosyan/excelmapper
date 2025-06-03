// api/processFile.js

import { IncomingForm } from "formidable";
import XLSX from "xlsx";
import fs from "fs";
import path from "path";

// Disable Next.js default body parsing so we can use Formidable
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
      console.error("Form parsing error:", err);
      return res.status(500).send("Error parsing uploaded file");
    }

    const file = files.file;
    if (!file) {
      return res
        .status(400)
        .send("Please upload a file under the field name 'file'.");
    }
    if (file.size === 0) {
      return res.status(400).send("Uploaded file was empty.");
    }

    let workbook;
    try {
      workbook = XLSX.readFile(file.filepath);
    } catch (xlsxErr) {
      console.error("XLSX.readFile error:", xlsxErr);
      return res.status(400).send("Invalid Excel file format.");
    }

    const firstSheetName = workbook.SheetNames[0];
    if (!firstSheetName) {
      return res.status(400).send("Uploaded workbook has no sheets.");
    }

    const sheetData = XLSX.utils.sheet_to_json(
      workbook.Sheets[firstSheetName],
      { defval: "" }
    );
    if (sheetData.length === 0) {
      return res
        .status(400)
        .send("The first sheet in the workbook contains no data.");
    }

    const output = sheetData.map((row) => {
      const newRow = {};
      for (const key in row) {
        const lowerKey = key.toLowerCase().trim();
        if (lowerKey.includes("make")) newRow["Make"] = row[key];
        if (lowerKey.includes("year")) newRow["Year"] = row[key];
        if (lowerKey.includes("vin")) newRow["VIN Number"] = row[key];
      }
      return newRow;
    });

    const newSheet = XLSX.utils.json_to_sheet(output);
    const newBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newBook, newSheet, "Standardized");

    const tempPath = path.join("/tmp", `processed-${Date.now()}.xlsx`);
    try {
      XLSX.writeFile(newBook, tempPath);
    } catch (writeErr) {
      console.error("Error writing new workbook:", writeErr);
      return res.status(500).send("Failed to create the processed file.");
    }

    let stats;
    try {
      stats = fs.statSync(tempPath);
    } catch (statErr) {
      console.error("Temp file not found or unreadable:", statErr);
      return res
        .status(500)
        .send("Processed file was not created correctly.");
    }
    if (stats.size === 0) {
      return res.status(500).send("Processed file is empty.");
    }

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
      return res.status(500).send("Failed to send processed file.");
    }
  });
}

// api/processFile.js

import formidable from "formidable";
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

  const form = new formidable.IncomingForm({ multiples: false });

  form.parse(req, async (err, fields, files) => {
    if (err) {
      console.error("Form parsing error:", err);
      return res.status(500).send("Error parsing uploaded file");
    }

    const file = files.file;
    if (!file) {
      return res.status(400).send("No file uploaded under field name 'file'");
    }

    try {
      // Read the uploaded Excel file from its temporary filepath
      const workbook = XLSX.readFile(file.filepath);
      const firstSheetName = workbook.SheetNames[0];
      const sheetData = XLSX.utils.sheet_to_json(
        workbook.Sheets[firstSheetName],
        { defval: "" }
      );

      // Transform rows by extracting only Make, Year, and VIN Number
      const output = sheetData.map((row) => {
        const newRow = {};
        for (const key in row) {
          const lowerKey = key.toLowerCase().trim();
          if (lowerKey.includes("make")) {
            newRow["Make"] = row[key];
          }
          if (lowerKey.includes("year")) {
            newRow["Year"] = row[key];
          }
          if (lowerKey.includes("vin")) {
            newRow["VIN Number"] = row[key];
          }
        }
        return newRow;
      });

      // Convert the transformed data back into a new workbook
      const newSheet = XLSX.utils.json_to_sheet(output);
      const newBook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(newBook, newSheet, "Standardized");

      // Write the new workbook to a temporary file in the /tmp directory
      const timestamp = Date.now();
      const tempPath = path.join("/tmp", `processed-${timestamp}.xlsx`);
      XLSX.writeFile(newBook, tempPath);

      // Verify that the file was created and is not empty
      let stats;
      try {
        stats = fs.statSync(tempPath);
      } catch (statErr) {
        console.error("Temporary file not created:", statErr);
        return res.status(500).send("Failed to create output file");
      }
      if (stats.size === 0) {
        console.error("Temporary file is empty");
        return res.status(500).send("Output file is empty");
      }

      // Read the temporary file into a buffer and send it back
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
      res.send(fileBuffer);
    } catch (processingError) {
      console.error("Processing error:", processingError);
      return res.status(500).send("Failed to process Excel file");
    }
  });
}

// api/processFile.js
import formidable from "formidable";
import XLSX from "xlsx";
import fs from "fs";
import path from "path";

// Tell Vercel not to parse the body (we’re using formidable instead)
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
      return res.status(500).send("Error parsing file");
    }

    const file = files.file;
    if (!file) return res.status(400).send("No file uploaded");

    try {
      // Read the uploaded Excel
      const workbook = XLSX.readFile(file.filepath);
      const firstSheetName = workbook.SheetNames[0];
      const sheetData = XLSX.utils.sheet_to_json(
        workbook.Sheets[firstSheetName],
        { defval: "" }
      );

      // Transform rows
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

      // Convert back to a new workbook
      const newSheet = XLSX.utils.json_to_sheet(output);
      const newBook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(newBook, newSheet, "Standardized");

      // Write to a temporary file in /tmp
      const tempPath = path.join("/tmp", `processed-${Date.now()}.xlsx`);
      XLSX.writeFile(newBook, tempPath);

      // Read the final file buffer
      const fileBuffer = fs.readFileSync(tempPath);
      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );
      res.setHeader(
        "Content-Disposition",
        `attachment; filename=Processed_${path.basename(file.originalFilename)}`
      );
      res.send(fileBuffer);
    } catch (error) {
      console.error("Processing error:", error);
      res.status(500).send("Failed to process Excel file");
    }
  });
}

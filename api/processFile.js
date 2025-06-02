const formidable = require("formidable");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

export const config = {
  api: {
    bodyParser: false, // Required for formidable
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
      const workbook = XLSX.readFile(file.filepath);
      const firstSheetName = workbook.SheetNames[0];
      const sheet = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], {
        defval: "",
      });

      // Auto-map known fields (flexible structure)
      const columns = ["make", "year", "vin"];
      const output = sheet.map((row) => {
        const newRow = {};
        for (let key in row) {
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
      XLSX.writeFile(newBook, tempPath);

      const fileBuffer = fs.readFileSync(tempPath);
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", "attachment; filename=Processed.xlsx");
      res.send(fileBuffer);
    } catch (error) {
      console.error("Processing error:", error);
      res.status(500).send("Failed to process Excel file");
    }
  });
}

// 1) Convert sheet → 2D array
const rows = XLSX.utils.sheet_to_json(sheet, {
  header: 1,
  defval: "",
  blankrows: false
});

if (!Array.isArray(rows) || rows.length === 0) {
  return res.status(400).send("Uploaded sheet is empty.");
}

// 2) Find the single header row index
let headerRowIdx = -1;
for (let r = 0; r < rows.length; r++) {
  // Lowercase all cells to ease substring checks
  const rowLower = rows[r].map(cell =>
    typeof cell === "string" ? cell.toLowerCase() : ""
  );
  const hasYear = rowLower.some(c => c.includes("year"));
  const hasMake = rowLower.some(c => c.includes("make"));
  const hasVin  = rowLower.some(c => c.includes("vin"));
  const hasCostOrStated =
    rowLower.some(c => c.includes("cost")) ||
    rowLower.some(c => c.includes("stated"));
  if (hasYear && hasMake && hasVin && hasCostOrStated) {
    headerRowIdx = r;
    break;
  }
}

if (headerRowIdx < 0) {
  return res.status(400).send(
    "Input sheet is missing a header row containing Year, Make, VIN, and Cost or Stated Value."
  );
}

// 3) All rows after headerRowIdx are “dataRows”
const dataRows = rows.slice(headerRowIdx + 1);
if (dataRows.length === 0) {
  return res.status(400).send("No data found after the header row.");
}
// 4) Look at up to N = 5 of those dataRows to score each column
const sampleSize = Math.min(5, dataRows.length);
const columnScores = [];

// Initialize an object for each column index
for (let c = 0; c < dataRows[0].length; c++) {
  columnScores[c] = {
    yearCount: 0,
    makeCount: 0,
    vinCount:  0,
    costCount: 0,
    total:     0
  };
}

// For each of the first sampleSize rows, update the four counters
for (let r = 0; r < sampleSize; r++) {
  const row = dataRows[r];
  for (let c = 0; c < row.length; c++) {
    const s = String(row[c] || "").trim();

    // Year: exactly 4 digits between 1900–2100
    if (/^\d{4}$/.test(s) && parseInt(s, 10) >= 1900 && parseInt(s, 10) <= 2100) {
      columnScores[c].yearCount++;
    }

    // VIN: 16+ alphanumeric (no spaces)
    if (/^[A-Za-z0-9]{16,}$/.test(s)) {
      columnScores[c].vinCount++;
    }

    // Make: all letters at least 2 chars (VOLVO, FORD, etc.)
    if (/^[A-Za-z]{2,}$/.test(s)) {
      columnScores[c].makeCount++;
    }

    // Cost: contains $ or comma or decimal, or is pure digits length > 4
    if (
      /^\$?[\d,]+(\.\d+)?$/.test(s) ||
      (/^\d+$/.test(s) && s.length > 4)
    ) {
      columnScores[c].costCount++;
    }

    columnScores[c].total++;
  }
}

// 5) Pick the column index with the highest “yearCount” (and at least 2 matches)
let yearColIdx = -1;
let bestYear = 0;
for (let c = 0; c < columnScores.length; c++) {
  if (columnScores[c].yearCount > bestYear) {
    bestYear = columnScores[c].yearCount;
    yearColIdx = c;
  }
}
if (bestYear < 2) {
  return res.status(400).send("Cannot reliably find the Year column.");
}

// 6) Pick “Make” similarly
let makeColIdx = -1;
let bestMake = 0;
for (let c = 0; c < columnScores.length; c++) {
  if (columnScores[c].makeCount > bestMake) {
    bestMake = columnScores[c].makeCount;
    makeColIdx = c;
  }
}
if (bestMake < 2) {
  return res.status(400).send("Cannot reliably find the Make column.");
}

// 7) Pick “VIN”
let vinColIdx = -1;
let bestVin = 0;
for (let c = 0; c < columnScores.length; c++) {
  if (columnScores[c].vinCount > bestVin) {
    bestVin = columnScores[c].vinCount;
    vinColIdx = c;
  }
}
if (bestVin < 1) {
  return res.status(400).send("Cannot reliably find the VIN column.");
}

// 8) Pick “Cost” (or “Stated”)
let costColIdx = -1;
let bestCost = 0;
for (let c = 0; c < columnScores.length; c++) {
  if (columnScores[c].costCount > bestCost) {
    bestCost = columnScores[c].costCount;
    costColIdx = c;
  }
}
if (bestCost < 1) {
  return res.status(400).send("Cannot reliably find the Cost column.");
}
// 9) Prepare a fresh “newSheet” and write A1→AO1
const newSheet = {};
OUTPUT_HEADERS.forEach((headerText, colIndex) => {
  const address = XLSX.utils.encode_cell({ r: 0, c: colIndex });
  newSheet[address] = { v: headerText };
});

// 10) Loop over all dataRows (not skipping any—they’re all “real” rows)
let outputRowCount = 0;
for (let r = 0; r < dataRows.length; r++) {
  const row = dataRows[r];
  const rawYear = row[yearColIdx]  || "";
  const rawMake = row[makeColIdx]  || "";
  const rawVin  = row[vinColIdx]   || "";
  const rawCost = row[costColIdx]  || "";

  // If all four are empty, we can skip (completely blank line)
  if (
    String(rawYear).trim() === "" &&
    String(rawMake).trim() === "" &&
    String(rawVin).trim() === "" &&
    String(rawCost).trim() === ""
  ) {
    continue;
  }

  outputRowCount++;
  const outRow = outputRowCount + 1; // data starts at row 2 (r=1)
  
  // Clean each field
  const yearVal  = String(rawYear).trim();
  const makeVal  = String(rawMake).trim();
  const vinVal   = String(rawVin).trim();
  const costVal  = String(rawCost).replace(/\D/g, ""); // digits only

  // Place them into E/F/J/V
  newSheet[`E${outRow}`] = { v: yearVal };
  newSheet[`F${outRow}`] = { v: makeVal };
  newSheet[`J${outRow}`] = { v: vinVal };
  newSheet[`V${outRow}`] = { v: costVal };
}

// 11) If we didn’t actually process any rows, error out
if (outputRowCount === 0) {
  return res
    .status(400)
    .send("No valid data rows (Year/Make/VIN/Cost) were found in the file.");
}

// 12) Define the range !ref from A1 → AO(lastDataRow)
const lastRowIndex = outputRowCount; // header=0, and data runs 1..outputRowCount
const lastColIndex = OUTPUT_HEADERS.length - 1; // 40 (A=0..AO=40)
const range = XLSX.utils.encode_range(
  { r: 0, c: 0 },
  { r: lastRowIndex, c: lastColIndex }
);
newSheet["!ref"] = range;

// 13) Append to a workbook, write to /tmp, and stream back exactly as before...
const newBook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newBook, newSheet, "Standardized");

const timestamp = Date.now();
const tempPath = path.join("/tmp", `processed-${timestamp}.xlsx`);
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
  console.error("Temp-file check error:", statErr);
  return res
    .status(500)
    .send("Processed file was not created correctly.");
}
if (stats.size === 0) {
  console.error("Temp file is empty after writing");
  return res.status(500).send("Processed file is empty.");
}

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

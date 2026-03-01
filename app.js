Office.onReady(() => {
  document.getElementById("cleanBtn").onclick = runClean;
});

function cleanName(value) {
  if (!value || !value.trim()) return value;
  return value
    .trim()
    .toLowerCase()
    .split(/\s+/)
    .map(w => w.charAt(0).toUpperCase() + w.slice(1))
    .join(" ");
}

function cleanEmail(value) {
  if (!value || !value.trim()) return value;
  return value.trim().toLowerCase();
}

function cleanDate(value) {
  const date = new Date(value);
  if (isNaN(date.getTime())) return value;
  return date.toISOString().split("T")[0];
}

function trimValue(value) {
  if (typeof value !== "string") return value;
  return value.trim();
}

function detectColumnType(header) {
  const h = header.toLowerCase();
  if (h.includes("name")) return "name";
  if (h.includes("email")) return "email";
  if (h.includes("date")) return "date";
  return "other";
}

function cleanValue(value, type) {
  const trimmed = trimValue(value);
  switch (type) {
    case "name": return cleanName(trimmed);
    case "email": return cleanEmail(trimmed);
    case "date": return cleanDate(trimmed);
    default: return trimmed;
  }
}

async function runClean() {
  const status = document.getElementById("status");
  status.innerText = "Processing...";

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    usedRange.load("values");
    await context.sync();

    const values = usedRange.values;
    if (!values || values.length < 2) {
      status.innerText = "No data found.";
      return;
    }

    const headers = values[0];
    const columnTypes = headers.map(detectColumnType);
    const dataRows = values.slice(1);

    const cleanedRows = dataRows.map(row =>
      row.map((cell, c) => cleanValue(String(cell || ""), columnTypes[c]))
    );

    const finalData = [headers, ...cleanedRows];
    usedRange.values = finalData;

    await context.sync();
    status.innerText = "Cleaning complete!";
  });
}
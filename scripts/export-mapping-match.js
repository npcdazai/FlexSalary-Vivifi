import "dotenv/config";
import fs from "node:fs";
import ExcelJS from "exceljs";

const SCRATCH = "/tmp/claude-1000/-home-ubuntu-FlexSalary/ce1b4704-a401-4295-9c0f-c75e7966915a/scratchpad";
const MAPPING_FILE = "/home/ubuntu/FlexSalary/mapping mis - vivifi.xlsx";

function loadJsonl(path) {
  return fs
    .readFileSync(path, "utf8")
    .split("\n")
    .filter(Boolean)
    .map((line) => JSON.parse(line));
}

function normalizePhone(p) {
  return p == null ? null : String(p).trim();
}

async function main() {
  const appLookup = new Map();
  for (const row of loadJsonl(`${SCRATCH}/applications_lookup.jsonl`)) {
    appLookup.set(normalizePhone(row.phone), row);
  }

  const webhookPhones = new Set(
    loadJsonl(`${SCRATCH}/webhook_events_phones2.jsonl`).map((r) => normalizePhone(r.phone))
  );

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(MAPPING_FILE);
  const sheet = wb.getWorksheet("RawData");

  const header = sheet.getRow(1).values.slice(1); // ExcelJS values array is 1-indexed with a leading empty slot
  const numCols = header.length;
  const phoneColIdx = header.indexOf("Phonenumber") + 1;

  const outHeader = [
    ...header,
    "foundInWebhookEvents",
    "foundInApplications",
    "ourLeadId",
    "ourName",
    "ourStatus",
    "ourMatchCount",
  ];

  const outWb = new ExcelJS.Workbook();
  const outSheet = outWb.addWorksheet("Mapping vs Our Data");
  outSheet.addRow(outHeader);
  outSheet.getRow(1).font = { bold: true };

  let totalRows = 0;
  let matchedInWebhookEvents = 0;
  let matchedInApplications = 0;

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // header
    // row.values.slice(1) is variable-length for sparse rows (trailing empty cells get
    // dropped, not padded) — reading each column explicitly by fixed index guarantees the
    // appended columns always land in the same absolute position for every row.
    const values = [];
    for (let col = 1; col <= numCols; col++) {
      values.push(row.getCell(col).value ?? null);
    }
    if (values.every((v) => v === null)) return; // truly blank formatted row

    totalRows++;
    const rawPhone = values[phoneColIdx - 1];
    const phone = normalizePhone(rawPhone);

    const inWebhook = phone ? webhookPhones.has(phone) : false;
    const appMatch = phone ? appLookup.get(phone) : undefined;
    const inApplications = Boolean(appMatch);

    if (inWebhook) matchedInWebhookEvents++;
    if (inApplications) matchedInApplications++;

    outSheet.addRow([
      ...values,
      inWebhook,
      inApplications,
      appMatch?.matchedLeadId ?? null,
      appMatch?.matchedName ?? null,
      appMatch?.matchedStatus ?? null,
      appMatch?.matchCount ?? 0,
    ]);
  });

  const outPath = "/home/ubuntu/FlexSalary/reports/Mapping_vs_Our_Data.xlsx";
  await outWb.xlsx.writeFile(outPath);

  console.log(`Source rows read: ${totalRows}`);
  console.log(`Output rows written: ${outSheet.rowCount - 1}`);
  console.log(`Matched in webhook_events: ${matchedInWebhookEvents}`);
  console.log(`Matched in applications: ${matchedInApplications}`);
  console.log(`Wrote: ${outPath}`);

  if (outSheet.rowCount - 1 !== totalRows) {
    throw new Error(
      `Row count mismatch: read ${totalRows} source rows but wrote ${outSheet.rowCount - 1} — possible data loss!`
    );
  }
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});

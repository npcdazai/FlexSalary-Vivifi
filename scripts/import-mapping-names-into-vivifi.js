import "dotenv/config";
import fs from "node:fs";
import ExcelJS from "exceljs";
import { clickhouse } from "../src/lib/clickhouse.js";

const SCRATCH = "/tmp/claude-1000/-home-ubuntu-FlexSalary/ce1b4704-a401-4295-9c0f-c75e7966915a/scratchpad";
const MAPPING_FILE = "/home/ubuntu/FlexSalary/mapping mis - vivifi.xlsx";

function loadPhoneSet(path) {
  return new Set(
    fs
      .readFileSync(path, "utf8")
      .split("\n")
      .filter(Boolean)
      .map((line) => String(JSON.parse(line).phone).trim())
  );
}

function toDateOnly(d) {
  if (!(d instanceof Date)) return null;
  return d.toISOString().slice(0, 10);
}

async function main() {
  const webhookPhones = loadPhoneSet(`${SCRATCH}/webhook_events_phones2.jsonl`);
  const existingVivifiPhones = loadPhoneSet(`${SCRATCH}/existing_vivifi_phones.jsonl`);

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(MAPPING_FILE);
  const sheet = wb.getWorksheet("RawData");

  const header = sheet.getRow(1).values.slice(1);
  const numCols = header.length;
  const idx = Object.fromEntries(header.map((h, i) => [h, i]));

  const byPhone = new Map(); // phone -> row (first occurrence; names never conflict, verified separately)
  let totalRead = 0;

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const values = [];
    for (let col = 1; col <= numCols; col++) values.push(row.getCell(col).value ?? null);
    if (values.every((v) => v === null)) return;
    totalRead++;

    const rawPhone = values[idx["Phonenumber"]];
    if (!rawPhone) return;
    const phone = String(rawPhone).trim();

    if (!webhookPhones.has(phone)) return; // only phones that can actually join to something
    if (existingVivifiPhones.has(phone)) return; // already has a proper ViviFi-sourced entry
    if (byPhone.has(phone)) return; // dedupe: keep first occurrence

    byPhone.set(phone, {
      name: values[idx["Names"]] ?? null,
      phone,
      message: "Imported from mapping mis - vivifi.xlsx",
      customerUniqueId: null,
      referenceId: null,
      email: values[idx["Email"]] ?? null,
      dob: null,
      panNumber: null,
      rawData: JSON.stringify(Object.fromEntries(header.map((h, i) => [h, values[i]]))),
      sourceFile: "mapping mis - vivifi.xlsx",
      sourceFileDate: toDateOnly(values[idx["LeadDate"]]) ?? "1970-01-01",
      rowNumber,
    });
  });

  const toInsert = Array.from(byPhone.values());
  console.log(`Mapping file rows read: ${totalRead}`);
  console.log(`Rows to insert into vivifi_lead_summary: ${toInsert.length}`);

  if (toInsert.length > 0) {
    await clickhouse.insert({ table: "vivifi_lead_summary", values: toInsert, format: "JSONEachRow" });
  }

  const verify = await clickhouse.query({
    query: `SELECT count() AS c FROM vivifi_lead_summary WHERE sourceFile = 'mapping mis - vivifi.xlsx'`,
    format: "JSONEachRow",
  });
  const inserted = Number((await verify.json())[0]?.c ?? 0);
  console.log(`Verified rows in ClickHouse with sourceFile = mapping file: ${inserted}`);

  if (inserted !== toInsert.length) {
    throw new Error(`Mismatch: tried to insert ${toInsert.length}, but found ${inserted} in ClickHouse`);
  }
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});

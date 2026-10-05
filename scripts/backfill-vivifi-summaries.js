import "dotenv/config";
import { S3Client, ListObjectsV2Command, GetObjectCommand } from "@aws-sdk/client-s3";
import ExcelJS from "exceljs";
import { clickhouse } from "../src/lib/clickhouse.js";

const BUCKET = process.env.AWS_BUCKET_NAME;
const PREFIXES = ["ViviFi/Jun-2026/summary-files/", "ViviFi/Jul-2026/summary-files/"];
const START_DATE = "2026-06-06";
const END_DATE = "2026-07-06";
const CONCURRENCY = 6;

const s3 = new S3Client({ region: process.env.AWS_REGION });

const MONTHS = { Jan: 1, Feb: 2, Mar: 3, Apr: 4, May: 5, Jun: 6, Jul: 7, Aug: 8, Sep: 9, Oct: 10, Nov: 11, Dec: 12 };

function fileDate(key) {
  const m = key.match(/(\d{2})-([A-Za-z]{3})-(\d{4})_/);
  if (!m) return null;
  const [, day, mon, year] = m;
  const month = String(MONTHS[mon]).padStart(2, "0");
  return `${year}-${month}-${day.padStart(2, "0")}`;
}

async function listFilesInRange() {
  const keys = [];
  for (const prefix of PREFIXES) {
    let token;
    do {
      const resp = await s3.send(
        new ListObjectsV2Command({ Bucket: BUCKET, Prefix: prefix, ContinuationToken: token })
      );
      for (const obj of resp.Contents ?? []) {
        const d = fileDate(obj.Key);
        if (d && d >= START_DATE && d <= END_DATE) keys.push({ key: obj.Key, date: d });
      }
      token = resp.NextContinuationToken;
    } while (token);
  }
  return keys;
}

async function alreadyIngested(key) {
  const result = await clickhouse.query({
    query: `SELECT count() AS c FROM vivifi_lead_summary WHERE sourceFile = {key:String}`,
    query_params: { key },
    format: "JSONEachRow",
  });
  const rows = await result.json();
  return Number(rows[0]?.c ?? 0) > 0;
}

function extractReferenceId(redirectUrl) {
  if (!redirectUrl) return null;
  const m = redirectUrl.match(/leadReferenceId=([0-9a-fA-F-]{36})/);
  return m ? m[1] : null;
}

async function downloadBuffer(key, attempt = 1) {
  try {
    const resp = await s3.send(new GetObjectCommand({ Bucket: BUCKET, Key: key }));
    const chunks = [];
    for await (const chunk of resp.Body) chunks.push(chunk);
    return Buffer.concat(chunks);
  } catch (err) {
    if (attempt >= 3) throw err;
    await new Promise((r) => setTimeout(r, 1000 * attempt));
    return downloadBuffer(key, attempt + 1);
  }
}

async function processFile({ key, date }, stats) {
  if (await alreadyIngested(key)) {
    stats.filesSkippedAlreadyDone++;
    return;
  }

  const buffer = await downloadBuffer(key);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const sheet = workbook.getWorksheet("Summary") ?? workbook.worksheets[0];

  const batch = [];
  let rowsRead = 0;
  let rowsSkippedEmpty = 0;
  let rowsSkippedNotAccepted = 0;

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // header
    rowsRead++;

    const name = row.getCell(1).text?.trim() ?? "";
    const phone = row.getCell(2).text?.trim() ?? "";
    const message = row.getCell(3).text?.trim() ?? "";
    const rawData = row.getCell(4).text?.trim() ?? "";

    if (!name && !phone && !rawData) {
      rowsSkippedEmpty++;
      return;
    }

    let parsed = null;
    try {
      parsed = rawData ? JSON.parse(rawData) : null;
    } catch {
      parsed = null;
    }

    // Only rows where FlexSalary actually accepted the application carry a
    // customerUniqueId — that's the same id FlexSalary calls leadId in its
    // webhooks, so this is the only case where a join to webhook_events makes
    // sense. Failure rows (email/phone already exists, etc.) are intentionally
    // skipped, not stored.
    if (!parsed?.customerUniqueId) {
      rowsSkippedNotAccepted++;
      return;
    }

    const personerDetails = parsed?.reqBody?.PersonerDetails ?? null;

    batch.push({
      name,
      phone,
      message,
      customerUniqueId: parsed.customerUniqueId,
      referenceId: extractReferenceId(parsed.redirectUrl),
      email: personerDetails?.Email ?? null,
      dob: personerDetails?.DateOfBirth ?? null,
      panNumber: personerDetails?.PanNumber ?? null,
      rawData,
      sourceFile: key,
      sourceFileDate: date,
      rowNumber,
    });
  });

  if (batch.length > 0) {
    await clickhouse.insert({ table: "vivifi_lead_summary", values: batch, format: "JSONEachRow" });
  }

  const verify = await clickhouse.query({
    query: `SELECT count() AS c FROM vivifi_lead_summary WHERE sourceFile = {key:String}`,
    query_params: { key },
    format: "JSONEachRow",
  });
  const inserted = Number((await verify.json())[0]?.c ?? 0);

  if (inserted !== batch.length) {
    throw new Error(`Row count mismatch for ${key}: expected ${batch.length}, got ${inserted}`);
  }

  stats.filesProcessed++;
  stats.rowsRead += rowsRead;
  stats.rowsInserted += batch.length;
  stats.rowsSkippedEmpty += rowsSkippedEmpty;
  stats.rowsSkippedNotAccepted += rowsSkippedNotAccepted;

  console.log(
    `[${stats.filesProcessed + stats.filesSkippedAlreadyDone}/${stats.totalFiles}] ${key} — read ${rowsRead}, inserted ${batch.length}, not-accepted ${rowsSkippedNotAccepted}, empty ${rowsSkippedEmpty}`
  );
}

async function runWithConcurrency(items, limit, worker) {
  let index = 0;
  const errors = [];
  async function next() {
    while (index < items.length) {
      const i = index++;
      try {
        await worker(items[i]);
      } catch (err) {
        errors.push({ item: items[i], error: err });
        console.error(`FAILED: ${items[i].key} — ${err.message}`);
      }
    }
  }
  await Promise.all(Array.from({ length: limit }, next));
  return errors;
}

async function main() {
  console.log(`Listing files in ${PREFIXES.join(", ")} between ${START_DATE} and ${END_DATE}...`);
  let files = await listFilesInRange();
  console.log(`Found ${files.length} files to process.`);

  if (process.env.BACKFILL_LIMIT) {
    files = files.slice(0, Number(process.env.BACKFILL_LIMIT));
    console.log(`BACKFILL_LIMIT set — restricting to first ${files.length} files for this run.`);
  }

  const stats = {
    totalFiles: files.length,
    filesProcessed: 0,
    filesSkippedAlreadyDone: 0,
    rowsRead: 0,
    rowsInserted: 0,
    rowsSkippedEmpty: 0,
    rowsSkippedNotAccepted: 0,
  };

  const errors = await runWithConcurrency(files, CONCURRENCY, (f) => processFile(f, stats));

  console.log("\n=== BACKFILL SUMMARY ===");
  console.log(JSON.stringify(stats, null, 2));
  console.log(`Failed files: ${errors.length}`);
  if (errors.length > 0) {
    console.log(errors.map((e) => e.item.key).join("\n"));
  }

  const finalCount = await clickhouse.query({
    query: `SELECT count() AS c FROM vivifi_lead_summary WHERE sourceFileDate BETWEEN {start:Date} AND {end:Date}`,
    query_params: { start: START_DATE, end: END_DATE },
    format: "JSONEachRow",
  });
  console.log("Rows currently in ClickHouse for this date range:", (await finalCount.json())[0]);

  if (errors.length > 0) {
    process.exitCode = 1;
  }
}

main().catch((err) => {
  console.error("Fatal error:", err);
  process.exitCode = 1;
});

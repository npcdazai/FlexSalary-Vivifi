import "dotenv/config";
import ExcelJS from "exceljs";
import { clickhouse } from "../src/lib/clickhouse.js";

// Leads in the 3 target statuses where webhook_events never carried an EligibleAmount
// (elig.eligibleAmount IS NULL) but the mapping-file import into vivifi_lead_summary did
// (v.eligibleAmount IS NOT NULL) — i.e. exactly the ones the mapping file resolved.
const query = `
  SELECT
      app.leadId AS leadId,
      nullIf(v.name, '') AS name,
      nullIf(v.phone, '') AS phone,
      app.phoneNumber AS phoneNumber,
      app.status AS status,
      v.eligibleAmount AS eligibleAmount,
      app.updatedAt AS updatedAt
  FROM
  (
      SELECT
          leadId,
          argMax(phoneNumber, receivedAt) AS phoneNumber,
          argMax(status, receivedAt) AS status,
          max(receivedAt) AS updatedAt
      FROM webhook_events
      GROUP BY leadId
  ) AS app
  LEFT JOIN vivifi_lead_summary AS v
      ON (app.leadId = v.customerUniqueId) OR (app.leadId = v.referenceId) OR (app.phoneNumber = v.phone)
  LEFT JOIN
  (
      SELECT leadId, argMax(eligibleAmount, receivedAt) AS eligibleAmount
      FROM
      (
          SELECT
              leadId,
              receivedAt,
              JSONExtract(data, 'EligibleAmount', 'Nullable(Decimal(14,2))') AS eligibleAmount
          FROM webhook_events
          WHERE JSONExtract(data, 'EligibleAmount', 'Nullable(Decimal(14,2))') IS NOT NULL
      ) AS raw
      GROUP BY leadId
  ) AS elig
      ON app.leadId = elig.leadId
  WHERE app.status IN ('Awaiting VKYC', 'Awaiting Esign', 'Awaiting EMandate')
    AND elig.eligibleAmount IS NULL
    AND v.eligibleAmount IS NOT NULL
  ORDER BY app.status, app.updatedAt
`;

const result = await clickhouse.query({ query, format: "JSONEachRow" });
const rows = await result.json();

const wb = new ExcelJS.Workbook();
const sheet = wb.addWorksheet("Resolved via Mapping File");
sheet.columns = [
  { header: "Lead ID", key: "leadId", width: 45 },
  { header: "Name", key: "name", width: 25 },
  { header: "Phone (ViviFi)", key: "phone", width: 16 },
  { header: "Phone (Webhook)", key: "phoneNumber", width: 16 },
  { header: "Status", key: "status", width: 20 },
  { header: "Eligible Amount (from mapping file)", key: "eligibleAmount", width: 28 },
  { header: "Updated At", key: "updatedAt", width: 28 },
];
sheet.getRow(1).font = { bold: true };
rows.forEach((r) => sheet.addRow(r));

const outPath = "/home/ubuntu/FlexSalary/reports/Eligible_Amount_Resolved_via_Mapping_File.xlsx";
await wb.xlsx.writeFile(outPath);
console.log(`Wrote ${rows.length} resolved rows to ${outPath}`);

const byStatus = {};
for (const r of rows) byStatus[r.status] = (byStatus[r.status] ?? 0) + 1;
console.log("By status:", byStatus);

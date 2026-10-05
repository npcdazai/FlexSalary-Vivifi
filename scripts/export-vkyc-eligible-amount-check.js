import "dotenv/config";
import ExcelJS from "exceljs";
import { clickhouse } from "../src/lib/clickhouse.js";

// Pulls every lead currently at "Awaiting VKYC" and, for each, shows the eligibleAmount
// plus exactly where it came from: FlexSalary's own webhook data, the mapping mis - vivifi
// file (via vivifi_lead_summary), or nowhere at all.
const query = `
  SELECT
      app.leadId AS leadId,
      nullIf(v.name, '') AS name,
      nullIf(v.phone, '') AS phone,
      app.phoneNumber AS phoneNumber,
      app.status AS status,
      coalesce(elig.eligibleAmount, v.eligibleAmount) AS eligibleAmount,
      elig.eligibleAmount AS webhookEligibleAmount,
      v.eligibleAmount AS misEligibleAmount,
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
  WHERE app.status = 'Awaiting VKYC'
  ORDER BY app.updatedAt
`;

const result = await clickhouse.query({ query, format: "JSONEachRow" });
const rows = await result.json();

function source(row) {
  if (row.webhookEligibleAmount != null) return "Webhook Data";
  if (row.misEligibleAmount != null) return "MIS File";
  return "Not Found";
}

const wb = new ExcelJS.Workbook();
const sheet = wb.addWorksheet("Awaiting VKYC - Eligible Amount");
sheet.columns = [
  { header: "Lead ID", key: "leadId", width: 45 },
  { header: "Name", key: "name", width: 25 },
  { header: "Phone (ViviFi)", key: "phone", width: 16 },
  { header: "Phone (Webhook)", key: "phoneNumber", width: 16 },
  { header: "Status", key: "status", width: 16 },
  { header: "Eligible Amount", key: "eligibleAmount", width: 18 },
  { header: "Source", key: "source", width: 16 },
  { header: "Updated At", key: "updatedAt", width: 28 },
];
sheet.getRow(1).font = { bold: true };

const summary = { "Webhook Data": 0, "MIS File": 0, "Not Found": 0 };
for (const r of rows) {
  const src = source(r);
  summary[src]++;
  sheet.addRow({
    leadId: r.leadId,
    name: r.name,
    phone: r.phone,
    phoneNumber: r.phoneNumber,
    status: r.status,
    eligibleAmount: r.eligibleAmount,
    source: src,
    updatedAt: r.updatedAt,
  });
}

const outPath = "/home/ubuntu/FlexSalary/reports/Awaiting_VKYC_Eligible_Amount_Check.xlsx";
await wb.xlsx.writeFile(outPath);

console.log(`Total Awaiting VKYC leads: ${rows.length}`);
console.log("Breakdown by source:", summary);
console.log(`Wrote: ${outPath}`);

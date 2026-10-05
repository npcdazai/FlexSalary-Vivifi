import "dotenv/config";
import ExcelJS from "exceljs";
import { clickhouse } from "../src/lib/clickhouse.js";

const STATUSES = ["Awaiting VKYC", "Awaiting EMandate", "Awaiting Esign"];

// Pulls every lead currently at any of the 3 statuses and, for each, shows the
// eligibleAmount plus exactly where it came from: FlexSalary's own webhook data, the
// mapping mis - vivifi file (via vivifi_lead_summary), or nowhere at all.
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
  WHERE app.status IN ({statuses:Array(String)})
  ORDER BY app.status, app.updatedAt
`;

const result = await clickhouse.query({
  query,
  query_params: { statuses: STATUSES },
  format: "JSONEachRow",
});
const rows = await result.json();

function source(row) {
  if (row.webhookEligibleAmount != null) return "Webhook Data";
  if (row.misEligibleAmount != null) return "MIS File";
  return "Not Found";
}

const wb = new ExcelJS.Workbook();
const sheet = wb.addWorksheet("Eligible Amount Check");
sheet.columns = [
  { header: "Lead ID", key: "leadId", width: 45 },
  { header: "Name", key: "name", width: 25 },
  { header: "Phone (ViviFi)", key: "phone", width: 16 },
  { header: "Phone (Webhook)", key: "phoneNumber", width: 16 },
  { header: "Status", key: "status", width: 20 },
  { header: "Eligible Amount", key: "eligibleAmount", width: 18 },
  { header: "Source", key: "source", width: 16 },
  { header: "Updated At", key: "updatedAt", width: 28 },
];
sheet.getRow(1).font = { bold: true };

const summary = {};
for (const r of rows) {
  const src = source(r);
  summary[r.status] ??= { "Webhook Data": 0, "MIS File": 0, "Not Found": 0 };
  summary[r.status][src]++;
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

const outPath = "/home/ubuntu/FlexSalary/reports/Eligible_Amount_Check_VKYC_Esign_EMandate.xlsx";
await wb.xlsx.writeFile(outPath);

console.log(`Total leads across all 3 statuses: ${rows.length}`);
console.log("Breakdown by status/source:", JSON.stringify(summary, null, 2));
console.log(`Wrote: ${outPath}`);

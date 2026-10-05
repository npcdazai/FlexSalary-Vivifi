import "dotenv/config";
import ExcelJS from "exceljs";
import { clickhouse } from "../src/lib/clickhouse.js";

const STATUS = process.argv[2];
if (!STATUS) {
  console.error('Usage: node scripts/export-applications-by-status.js "Status Name"');
  process.exit(1);
}

const result = await clickhouse.query({
  query: `
    SELECT leadId, name, phone, phoneNumber, eventType, status, rejectionReason, eligibleAmount, createdAt, updatedAt
    FROM applications
    WHERE status = {status:String}
    ORDER BY updatedAt
  `,
  query_params: { status: STATUS },
  format: "JSONEachRow",
});
const rows = await result.json();

const wb = new ExcelJS.Workbook();
const sheet = wb.addWorksheet(STATUS.slice(0, 31)); // Excel sheet name limit
sheet.columns = [
  { header: "Lead ID", key: "leadId", width: 45 },
  { header: "Name", key: "name", width: 25 },
  { header: "Phone (ViviFi)", key: "phone", width: 16 },
  { header: "Phone (Webhook)", key: "phoneNumber", width: 16 },
  { header: "Latest Event Type", key: "eventType", width: 20 },
  { header: "Status", key: "status", width: 22 },
  { header: "Rejection Reason", key: "rejectionReason", width: 30 },
  { header: "Eligible Amount", key: "eligibleAmount", width: 16 },
  { header: "Created At", key: "createdAt", width: 28 },
  { header: "Updated At", key: "updatedAt", width: 28 },
];
sheet.getRow(1).font = { bold: true };
rows.forEach((r) => sheet.addRow(r));

const safeName = STATUS.replace(/[^a-z0-9]+/gi, "_");
const outPath = `/home/ubuntu/FlexSalary/reports/Applications_${safeName}.xlsx`;
await wb.xlsx.writeFile(outPath);
console.log(`Wrote ${rows.length} rows to ${outPath}`);

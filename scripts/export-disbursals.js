import "dotenv/config";
import ExcelJS from "exceljs";
import { clickhouse } from "../src/lib/clickhouse.js";

const DATE = process.argv[2];
if (!DATE) {
  console.error("Usage: node scripts/export-disbursals.js YYYY-MM-DD");
  process.exit(1);
}

const result = await clickhouse.query({
  query: `
    SELECT leadId, phoneNumber, disbursalAmount, disbursalDate
    FROM loans
    WHERE toDate(disbursalDate) = {date:Date}
    ORDER BY disbursalDate
  `,
  query_params: { date: DATE },
  format: "JSONEachRow",
});
const rows = await result.json();

const wb = new ExcelJS.Workbook();
const sheet = wb.addWorksheet(`Disbursals ${DATE}`);
sheet.columns = [
  { header: "Lead ID", key: "leadId", width: 45 },
  { header: "Phone Number", key: "phoneNumber", width: 18 },
  { header: "Disbursal Amount", key: "disbursalAmount", width: 18 },
  { header: "Disbursal Date", key: "disbursalDate", width: 28 },
];
sheet.getRow(1).font = { bold: true };

let total = 0;
for (const r of rows) {
  sheet.addRow(r);
  total += Number(r.disbursalAmount);
}
const totalRow = sheet.addRow({ leadId: "", phoneNumber: "TOTAL", disbursalAmount: total });
totalRow.font = { bold: true };

const outPath = `/home/ubuntu/FlexSalary/reports/Disbursals_${DATE}.xlsx`;
await wb.xlsx.writeFile(outPath);
console.log(`Wrote ${rows.length} rows, total = ${total}, to ${outPath}`);

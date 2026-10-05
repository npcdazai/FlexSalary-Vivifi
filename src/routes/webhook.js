import { Router } from "express";
import { clickhouse } from "../lib/clickhouse.js";

export const webhookRouter = Router();

const VALID_EVENT_TYPES = new Set(["APPLICATION_STATUS", "LOAN_STATUS"]);

function isValidPayload(body) {
  return (
    body &&
    typeof body === "object" &&
    VALID_EVENT_TYPES.has(body.eventType) &&
    typeof body.status === "string" &&
    typeof body.leadId === "string" &&
    typeof body.timestamp === "string" &&
    !Number.isNaN(Date.parse(body.timestamp))
  );
}

// The eventTimestamp column is DateTime64(7, 'Asia/Kolkata'): ClickHouse interprets a
// naive 'YYYY-MM-DD HH:MM:SS.ffffff' string as already being in that timezone, so this
// converts any incoming offset to true IST wall-clock time rather than UTC (India has no
// DST, so a fixed +5:30 is safe). Sub-second digits are taken from the original string,
// since Date only carries millisecond precision and FlexSalary sends up to 7 digits.
function toClickHouseDateTime(isoString) {
  const fractionMatch = isoString.match(/\.(\d+)/);
  const fraction = (fractionMatch ? fractionMatch[1] : "0").padEnd(7, "0").slice(0, 7);

  const kolkataMs = new Date(isoString).getTime() + 5.5 * 60 * 60 * 1000;
  const kolkata = new Date(kolkataMs);
  const pad = (n) => String(n).padStart(2, "0");

  const datePart = `${kolkata.getUTCFullYear()}-${pad(kolkata.getUTCMonth() + 1)}-${pad(kolkata.getUTCDate())}`;
  const timePart = `${pad(kolkata.getUTCHours())}:${pad(kolkata.getUTCMinutes())}:${pad(kolkata.getUTCSeconds())}`;
  return `${datePart} ${timePart}.${fraction}`;
}

async function isDuplicate({ leadId, eventType, status, eventTimestamp }) {
  const result = await clickhouse.query({
    query: `
      SELECT count() AS count
      FROM webhook_events
      WHERE leadId = {leadId:String}
        AND eventType = {eventType:String}
        AND status = {status:String}
        AND eventTimestamp = {eventTimestamp:DateTime64(7)}
    `,
    query_params: { leadId, eventType, status, eventTimestamp },
    format: "JSONEachRow",
  });
  const rows = await result.json();
  return Number(rows[0]?.count ?? 0) > 0;
}

webhookRouter.post("/flexsalary", async (req, res) => {
  const payload = req.body;

  if (!isValidPayload(payload)) {
    return res.status(400).json({ status: 400, message: "invalid payload" });
  }

  const { eventType, status, leadId, phoneNumber, timestamp, data } = payload;
  const eventTimestamp = toClickHouseDateTime(timestamp);

  try {
    // FlexSalary may retry delivery; skip re-inserting an event we've already recorded.
    if (await isDuplicate({ leadId, eventType, status, eventTimestamp })) {
      return res.status(200).json({ status: 200, message: "success" });
    }

    await clickhouse.insert({
      table: "webhook_events",
      values: [
        {
          eventType,
          status,
          leadId,
          phoneNumber: phoneNumber ?? null,
          eventTimestamp,
          data: data ? JSON.stringify(data) : "",
          rawPayload: JSON.stringify(payload),
        },
      ],
      format: "JSONEachRow",
    });
  } catch (err) {
    console.error("Failed to persist webhook event", err);
    return res.status(500).json({ status: 500, message: "failed to record event" });
  }

  // applications/loans are kept in sync automatically by ClickHouse materialized
  // views (mv_applications, mv_loans) defined on top of webhook_events.
  res.status(200).json({ status: 200, message: "success" });
});

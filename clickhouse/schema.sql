-- =====================================================================
-- webhook_events -> append-only fact table (source of truth for everything)
-- =====================================================================
CREATE TABLE IF NOT EXISTS webhook_events
(
    eventType       LowCardinality(String),          -- APPLICATION_STATUS | LOAN_STATUS
    status          LowCardinality(String),
    leadId          String,
    phoneNumber     Nullable(String),
    eventTimestamp  DateTime64(7, 'Asia/Kolkata'),    -- timestamp FlexSalary sent us
    data            String DEFAULT '',                -- raw `data` object as JSON text
    rawPayload      String,                            -- full raw envelope as JSON text, for replay/audit
    receivedAt      DateTime64(3, 'Asia/Kolkata') DEFAULT now64(3),
    createdAt       DateTime64(3, 'Asia/Kolkata') DEFAULT receivedAt
)
ENGINE = MergeTree
PARTITION BY toYYYYMM(eventTimestamp)
ORDER BY (leadId, eventTimestamp, eventType, status)
-- Dedupe of partner retries is done at the application layer before insert (see
-- src/routes/webhook.js), since ClickHouse has no unique constraint.
;

-- =====================================================================
-- vivifi_lead_summary -> one row per lead, sourced from two places:
--   1. ViviFi's own S3 summary-file exports (customerUniqueId/referenceId populated,
--      backfilled historically — see scripts/backfill-vivifi-summaries.js)
--   2. The "mapping mis - vivifi.xlsx" MIS file (customerUniqueId/referenceId NULL,
--      phone/name/eligibleAmount populated instead — see
--      scripts/import-mapping-names-into-vivifi.js). Tagged via sourceFile.
-- `phone`, `customerUniqueId`, and `referenceId` are each independently unique across
-- the whole table — applications' join below relies on that to avoid row multiplication.
-- =====================================================================
-- (schema omitted here — created via the backfill/import scripts; see those for the
-- exact column list: name, phone, message, customerUniqueId, referenceId, email, dob,
-- panNumber, eligibleAmount, rawData, sourceFile, sourceFileDate, rowNumber, createdAt)

-- =====================================================================
-- applications / loans -> latest status per lead.
--
-- These are REAL physical MergeTree tables (not views), each kept in sync by a
-- Refreshable Materialized View that fully recomputes and atomically replaces their
-- contents every 30 seconds (REFRESH EVERY 30 SECOND ... TO <table>). This gives the
-- same "always exactly one row per leadId, no FINAL/merge-timing issues" guarantee a
-- live view had, but backed by physical storage — reads never re-run the underlying
-- join, and BI tools that expect a plain table (not a view) work unmodified.
--
-- Tradeoff vs. the live view this replaced: up to ~30s of staleness after a new
-- webhook, instead of always being instantaneously live. Adjust REFRESH EVERY if a
-- different staleness/cost tradeoff is needed — check system.view_refreshes for
-- current status/timing.
--
-- name/phone/eligibleAmount are enriched from vivifi_lead_summary. FlexSalary is
-- inconsistent about what it sends as leadId — sometimes ViviFi's customerUniqueId
-- (hex string), sometimes its referenceId (UUID) — and phone is also matched as a
-- third fallback (for MIS-file-sourced rows that have no customerUniqueId/referenceId
-- at all). All three join keys are independently unique in vivifi_lead_summary, so
-- this can't multiply rows; LEFT JOIN means leads with no match simply get NULL,
-- never dropped.
-- =====================================================================

CREATE TABLE IF NOT EXISTS applications
(
    leadId          String,
    name            Nullable(String),
    phone           Nullable(String),
    phoneNumber     Nullable(String),
    eventType       String,
    status          String,
    rejectionReason Nullable(String),
    eligibleAmount  Nullable(Decimal(14, 2)),
    createdAt       DateTime64(3, 'Asia/Kolkata'),
    updatedAt       DateTime64(3, 'Asia/Kolkata')
)
ENGINE = MergeTree
ORDER BY leadId;

CREATE MATERIALIZED VIEW IF NOT EXISTS mv_refresh_applications
REFRESH EVERY 30 SECOND
TO applications
AS
SELECT
    app.leadId AS leadId,
    nullIf(v.name, '') AS name,
    nullIf(v.phone, '') AS phone,
    app.phoneNumber AS phoneNumber,
    app.eventType AS eventType,
    app.status AS status,
    app.rejectionReason AS rejectionReason,
    coalesce(elig.eligibleAmount, v.eligibleAmount) AS eligibleAmount,
    app.createdAt AS createdAt,
    app.updatedAt AS updatedAt
FROM
(
    SELECT
        leadId,
        argMax(phoneNumber, receivedAt) AS phoneNumber,
        argMax(eventType, receivedAt) AS eventType,
        argMax(status, receivedAt) AS status,
        argMax(nullIf(JSONExtractString(data, 'rejectionReason'), ''), receivedAt) AS rejectionReason,
        min(createdAt) AS createdAt,
        max(receivedAt) AS updatedAt
    FROM webhook_events
    GROUP BY leadId
) AS app
LEFT JOIN vivifi_lead_summary AS v
    ON (app.leadId = v.customerUniqueId) OR (app.leadId = v.referenceId) OR (app.phoneNumber = v.phone)
LEFT JOIN
(
    -- EligibleAmount can arrive on either event type and not every event carries it —
    -- take the most recent event that DID carry a value, not just the latest event
    -- overall (which would wrongly show NULL whenever that event lacked it).
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
    ON app.leadId = elig.leadId;

CREATE TABLE IF NOT EXISTS loans
(
    leadId          String,
    phoneNumber     Nullable(String),
    status          String,
    amount          Nullable(Decimal(14, 2)),
    disbursalDate   Nullable(DateTime64(7, 'Asia/Kolkata')),
    disbursalAmount Nullable(Decimal(14, 2)),
    createdAt       DateTime64(3, 'Asia/Kolkata'),
    updatedAt       DateTime64(3, 'Asia/Kolkata')
)
ENGINE = MergeTree
ORDER BY leadId;

CREATE MATERIALIZED VIEW IF NOT EXISTS mv_refresh_loans
REFRESH EVERY 30 SECOND
TO loans
AS
SELECT
    leadId,
    argMax(phoneNumber, receivedAt) AS phoneNumber,
    argMax(status, receivedAt) AS status,
    argMax(JSONExtract(data, 'amount', 'Nullable(Decimal(14,2))'), receivedAt) AS amount,
    -- JSONExtract(..., 'DateTime64') can't parse the offset-suffixed ISO string FlexSalary
    -- sends (returns NULL); parseDateTime64BestEffort handles the offset correctly.
    argMax(parseDateTime64BestEffortOrNull(JSONExtractString(data, 'disbursalDate'), 7, 'Asia/Kolkata'), receivedAt) AS disbursalDate,
    argMax(JSONExtract(data, 'disbursalAmount', 'Nullable(Decimal(14,2))'), receivedAt) AS disbursalAmount,
    min(createdAt) AS createdAt,
    max(receivedAt) AS updatedAt
FROM webhook_events
WHERE eventType = 'LOAN_STATUS'
GROUP BY leadId
-- Only these 2 statuses are wanted in `loans` — leads still at Awaiting Esign/EMandate
-- are intentionally excluded here (full history is still in webhook_events).
HAVING status IN ('Disbursed', 'Pending For Disbursal');

-- =====================================================================
-- vivifi_leads_status -> convenience view joining vivifi_lead_summary to
-- applications/loans on customerUniqueId = leadId, for a single combined read.
-- =====================================================================
CREATE VIEW IF NOT EXISTS vivifi_leads_status AS
SELECT
    v.customerUniqueId AS leadId,
    v.name,
    v.phone,
    v.email,
    v.dob,
    v.panNumber,
    v.referenceId,
    v.sourceFile,
    v.sourceFileDate,
    v.createdAt AS vivifiAcceptedAt,
    a.eventType AS latestEventType,
    a.status AS applicationStatus,
    l.status AS loanStatus,
    l.amount,
    l.disbursalDate,
    l.disbursalAmount
FROM vivifi_lead_summary v
LEFT JOIN applications a ON v.customerUniqueId = a.leadId
LEFT JOIN loans l ON v.customerUniqueId = l.leadId;

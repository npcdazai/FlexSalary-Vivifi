-- CreateEnum
CREATE TYPE "EventType" AS ENUM ('APPLICATION_STATUS', 'LOAN_STATUS');

-- CreateTable
CREATE TABLE "WebhookEvent" (
    "id" BIGSERIAL NOT NULL,
    "eventType" "EventType" NOT NULL,
    "status" TEXT NOT NULL,
    "leadId" TEXT NOT NULL,
    "phoneNumber" TEXT,
    "eventTimestamp" TIMESTAMP(3) NOT NULL,
    "data" JSONB,
    "rawPayload" JSONB NOT NULL,
    "receivedAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "WebhookEvent_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "Application" (
    "leadId" TEXT NOT NULL,
    "phoneNumber" TEXT,
    "status" TEXT NOT NULL,
    "rejectionReason" TEXT,
    "updatedAt" TIMESTAMP(3) NOT NULL,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "Application_pkey" PRIMARY KEY ("leadId")
);

-- CreateTable
CREATE TABLE "Loan" (
    "leadId" TEXT NOT NULL,
    "phoneNumber" TEXT,
    "status" TEXT NOT NULL,
    "amount" DECIMAL(14,2),
    "disbursalDate" TIMESTAMP(3),
    "disbursalAmount" DECIMAL(14,2),
    "updatedAt" TIMESTAMP(3) NOT NULL,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "Loan_pkey" PRIMARY KEY ("leadId")
);

-- CreateIndex
CREATE INDEX "WebhookEvent_leadId_idx" ON "WebhookEvent"("leadId");

-- CreateIndex
CREATE INDEX "WebhookEvent_eventType_status_idx" ON "WebhookEvent"("eventType", "status");

-- CreateIndex
CREATE UNIQUE INDEX "WebhookEvent_leadId_eventType_status_eventTimestamp_key" ON "WebhookEvent"("leadId", "eventType", "status", "eventTimestamp");

import crypto from "node:crypto";

const HEADER_NAME = process.env.WEBHOOK_AUTH_HEADER || "apikey";

function safeEqual(a, b) {
  const bufA = Buffer.from(String(a));
  const bufB = Buffer.from(String(b));
  if (bufA.length !== bufB.length) return false;
  return crypto.timingSafeEqual(bufA, bufB);
}

export function verifyApiKey(req, res, next) {
  const expected = process.env.FLEXSALARY_API_KEY;
  if (!expected) {
    // Misconfiguration, not a client error — fail closed either way.
    return res.status(500).json({ status: 500, message: "server not configured" });
  }

  const provided = req.get(HEADER_NAME);
  if (!provided || !safeEqual(provided, expected)) {
    return res.status(401).json({ status: 401, message: "unauthorized" });
  }

  next();
}

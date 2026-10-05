import "dotenv/config";
import express from "express";
import helmet from "helmet";
import morgan from "morgan";
import { webhookRouter } from "./routes/webhook.js";
import { verifyApiKey } from "./middleware/verifyApiKey.js";

const app = express();

app.use(helmet());
app.use(morgan("tiny"));
app.use(express.json({ limit: "1mb" }));

app.get("/health", (req, res) => res.json({ status: "ok" }));

app.use("/webhooks", verifyApiKey, webhookRouter);

// Malformed JSON body, etc.
app.use((err, req, res, next) => {
  if (err.type === "entity.parse.failed") {
    return res.status(400).json({ status: 400, message: "invalid JSON" });
  }
  console.error(err);
  res.status(500).json({ status: 500, message: "internal error" });
});

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`FlexSalary webhook service listening on :${port}`);
});

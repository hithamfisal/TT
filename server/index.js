import express from "express";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { handleDownloadWorkbook } from "./microsoftProxy.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, "..");
const distDir = path.join(projectRoot, "dist");
const app = express();
const port = Number(process.env.PORT ?? 3000);

app.use(express.json({ limit: "1mb" }));
app.post("/api/download-workbook", handleDownloadWorkbook);
app.use(express.static(distDir));
app.get("*", (_req, res) => {
  res.sendFile(path.join(distDir, "index.html"));
});

app.listen(port, () => {
  console.log(`Dashboard server running at http://localhost:${port}`);
});

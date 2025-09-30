// export_and_send.js
import { execSync } from "node:child_process";
import { writeFileSync, readFileSync } from "node:fs";
import { JWT } from "google-auth-library";

const {
  SA_JSON_BASE64,
  SHEET_ID, GID, RANGE_A1,
  SEA_URL,

  PNG_NAME = "Report.png",
  PAPER_SIZE = "letter",
  PORTRAIT = "true",
  FITW = "true",
  GRIDLINES = "false",
  MARGIN_INCH = "0.30",
  MAX_BYTES_MB = "5",
  SCALE_TO_PX = "1600"
} = process.env;

if (!SA_JSON_BASE64 || !SHEET_ID || !GID || !RANGE_A1 || !SEA_URL) {
  console.error("Missing env: SA_JSON_BASE64, SHEET_ID, GID, RANGE_A1, SEA_URL");
  process.exit(1);
}

// URL export PDF (được bảo vệ bởi OAuth nhưng cho phép tham số gid + range)
const exportUrl =
  `https://docs.google.com/spreadsheets/d/${encodeURIComponent(SHEET_ID)}/export` +
  `?exportFormat=pdf` +
  `&gid=${encodeURIComponent(GID)}` +
  `&range=${encodeURIComponent(RANGE_A1)}` +
  `&size=${PAPER_SIZE}&portrait=${PORTRAIT}&fitw=${FITW}&gridlines=${GRIDLINES}` +
  `&fzr=FALSE&top_margin=${MARGIN_INCH}&bottom_margin=${MARGIN_INCH}` +
  `&left_margin=${MARGIN_INCH}&right_margin=${MARGIN_INCH}`;

// Lấy access token từ Service Account
const sa = JSON.parse(Buffer.from(SA_JSON_BASE64, "base64").toString("utf8"));
const jwt = new JWT({
  email: sa.client_email,
  key: sa.private_key,
  scopes: ["https://www.googleapis.com/auth/drive.readonly"],
});

// Node 20 có sẵn fetch toàn cục
const accessToken = (await jwt.getAccessToken()).token;
if (!accessToken) {
  console.error("Failed to obtain access token");
  process.exit(1);
}

// Tải PDF
const pdfResp = await fetch(exportUrl, {
  headers: { Authorization: `Bearer ${accessToken}` },
});
if (!pdfResp.ok) {
  console.error("Export PDF failed:", pdfResp.status, await pdfResp.text());
  process.exit(1);
}
const pdfBuf = Buffer.from(await pdfResp.arrayBuffer());
writeFileSync("report.pdf", pdfBuf);
console.log("PDF bytes:", pdfBuf.length);

// Convert PDF -> PNG (trang đầu) bằng pdftoppm
const scale = Number(SCALE_TO_PX) || 1600;
execSync(`pdftoppm -png -singlefile -scale-to ${scale} report.pdf report`, { stdio: "inherit" });

let png = readFileSync("report.png");
const maxBytes = (Number(MAX_BYTES_MB) || 5) * 1024 * 1024;
if (png.length > maxBytes) {
  const scale2 = Math.max(800, Math.floor(scale * 0.75));
  console.log(`PNG too big; retry with scale-to=${scale2}`);
  execSync(`pdftoppm -png -singlefile -scale-to ${scale2} report.pdf report`, { stdio: "inherit" });
  png = readFileSync("report.png");
}

// Gửi PNG vào SeaTalk (tag=file để có preview inline)
const payload = {
  tag: "file",
  file: { filename: PNG_NAME, content: Buffer.from(png).toString("base64") }
};

const sea = await fetch(SEA_URL, {
  method: "POST",
  headers: { "Content-Type": "application/json" },
  body: JSON.stringify(payload)
});
console.log("SeaTalk:", sea.status, await sea.text());

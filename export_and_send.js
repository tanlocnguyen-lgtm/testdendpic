// export_and_send.js (CommonJS)
const { execSync } = require("node:child_process");
const { writeFileSync, readFileSync } = require("node:fs");
const { JWT } = require("google-auth-library");

// Lấy biến môi trường từ workflow
const {
  SA_JSON_BASE64, SHEET_ID, GID, RANGE_A1, SEA_URL,
  PNG_NAME = "Report.png",
  PAPER_SIZE = "letter",
  PORTRAIT = "true",
  FITW = "true",
  GRIDLINES = "false",
  MARGIN_INCH = "0.30",
  MAX_BYTES_MB = "5",
  SCALE_TO_PX = "1600"
} = process.env;

function need(v, name) { if (!v) { console.error(`Missing env: ${name}`); process.exit(1); } }
need(SA_JSON_BASE64,'SA_JSON_BASE64');
need(SHEET_ID,'SHEET_ID');
need(GID,'GID');
need(RANGE_A1,'RANGE_A1');
need(SEA_URL,'SEA_URL');

(async () => {
  try {
const exportUrl =
  `https://docs.google.com/spreadsheets/d/${encodeURIComponent(SHEET_ID)}/export` +
  `?exportFormat=pdf&gid=${encodeURIComponent(GID)}` +
  `&range=${encodeURIComponent(RANGE_A1)}` +
  `&size=${PAPER_SIZE}&portrait=${PORTRAIT}&fitw=${FITW}&gridlines=${GRIDLINES}` +
  `&fzr=FALSE&top_margin=${MARGIN_INCH}&bottom_margin=${MARGIN_INCH}` +
  `&left_margin=${MARGIN_INCH}&right_margin=${MARGIN_INCH}`;


    // Giải mã SA JSON & xin access token
    const sa = JSON.parse(Buffer.from(SA_JSON_BASE64, "base64").toString("utf8"));
    const jwt = new JWT({
      email: sa.client_email,
      key: sa.private_key,
      scopes: ["https://www.googleapis.com/auth/drive.readonly"],
    });

    const token = (await jwt.getAccessToken()).token;
    if (!token) { console.error("Failed to obtain access token"); process.exit(1); }

    console.log("Export URL:", exportUrl);

    // Export PDF
    const pdfResp = await fetch(exportUrl, { headers: { Authorization: `Bearer ${token}` }});
    console.log("PDF export status:", pdfResp.status);
    if (!pdfResp.ok) { console.error("Export PDF failed:", await pdfResp.text()); process.exit(1); }

    const pdfBuf = Buffer.from(await pdfResp.arrayBuffer());
    writeFileSync("report.pdf", pdfBuf);
    console.log("PDF bytes:", pdfBuf.length);

    // Convert PDF -> PNG (pdftoppm)
    const scale = Number(SCALE_TO_PX) || 1600;
    execSync(`pdftoppm -png -singlefile -scale-to ${scale} report.pdf report`, { stdio: "inherit" });

    let png = readFileSync("report.png");
    console.log("PNG bytes:", png.length);

    // Nếu PNG > 5MB (SeaTalk limit) thì giảm size và convert lại
    const maxBytes = (Number(MAX_BYTES_MB)||5) * 1024 * 1024;
    if (png.length > maxBytes) {
      const scale2 = Math.max(800, Math.floor(scale * 0.75));
      console.log(`PNG too big; retry with scale-to=${scale2}`);
      execSync(`pdftoppm -png -singlefile -scale-to ${scale2} report.pdf report`, { stdio: "inherit" });
      png = readFileSync("report.png");
      console.log("PNG bytes after shrink:", png.length);
    }

    // Gửi file PNG lên SeaTalk (tag: file)
    const payload = { tag: "file", file: { filename: PNG_NAME, content: Buffer.from(png).toString("base64") } };
    const sea = await fetch(SEA_URL, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
    console.log("SeaTalk status:", sea.status);
    console.log("SeaTalk body:", await sea.text());
  } catch (e) {
    console.error(e);
    process.exit(1);
  }
})();

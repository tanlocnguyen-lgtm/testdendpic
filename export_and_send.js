// export_and_send.js (CommonJS)
// Clean + validated version

const { execSync } = require("node:child_process");
const { writeFileSync, readFileSync } = require("node:fs");
const { JWT } = require("google-auth-library");

// ENV
const {
  SA_JSON_BASE64, SHEET_ID, GID, RANGE_A1, SEA_URL,
  PNG_NAME = "Report.png",
  PORTRAIT = "true",
  FITW = "true",
  GRIDLINES = "false",
  MAX_BYTES_MB = "5",
  SCALE_TO_PX = "1600",
} = process.env;

// Validate env
for (const k of ["SA_JSON_BASE64", "SHEET_ID", "GID", "RANGE_A1", "SEA_URL"]) {
  if (!process.env[k]) {
    console.error("Missing env:", k);
    process.exit(1);
  }
}

// Helpers
function colLetterToIndex(letter) {
  let n = 0;
  for (let i = 0; i < letter.length; i++)
    n = n * 26 + (letter.charCodeAt(i) - 64);
  return n;
}

function parseA1Range(a1) {
  const [a, b] = a1.split(":");
  function parseCell(c) {
    const m = c.match(/^([A-Z]+)(\d+)$/i);
    return { col: colLetterToIndex(m[1]), row: Number(m[2]) };
  }
  if (!b) {
    const c = parseCell(a);
    return { startRow: c.row, endRow: c.row, startCol: c.col, endCol: c.col };
  }
  const c1 = parseCell(a), c2 = parseCell(b);
  return {
    startRow: Math.min(c1.row, c2.row),
    endRow: Math.max(c1.row, c2.row),
    startCol: Math.min(c1.col, c2.col),
    endCol: Math.max(c1.col, c2.col),
  };
}

// MAIN
(async () => {
  try {
    // AUTH --------------------------------------------------------
    const sa = JSON.parse(Buffer.from(SA_JSON_BASE64, "base64").toString());
    const jwt = new JWT({
      email: sa.client_email,
      key: sa.private_key,
      scopes: [
        "https://www.googleapis.com/auth/drive.readonly",
        "https://www.googleapis.com/auth/spreadsheets",
      ],
    });

    const token = (await jwt.getAccessToken()).token;
    if (!token) throw new Error("Could not fetch access token");

    const parsed = parseA1Range(RANGE_A1);
    console.log("Parsed range:", parsed);

    // 1) DUPLICATE SHEET ------------------------------------------
    const dupName = "tmp_export_" + Date.now();
    let resp = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}:batchUpdate`,
      {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify({
          requests: [
            {
              duplicateSheet: {
                sourceSheetId: Number(GID),
                insertSheetIndex: 0,
                newSheetName: dupName,
              },
            },
          ],
        }),
      }
    );

    if (!resp.ok) {
      console.error(await resp.text());
      process.exit(1);
    }

    const dup = await resp.json();
    const tempSheetId = dup.replies[0].duplicateSheet.properties.sheetId;
    const gridRows = dup.replies[0].duplicateSheet.properties.gridProperties.rowCount;
    const gridCols = dup.replies[0].duplicateSheet.properties.gridProperties.columnCount;

    // 2) CROP SHEET -----------------------------------------------
    const reqs = [];
    const sr = parsed.startRow - 1;
    const er = parsed.endRow;
    const sc = parsed.startCol - 1;
    const ec = parsed.endCol;

    if (sr > 0) reqs.push({ deleteDimension: { range: { sheetId: tempSheetId, dimension: "ROWS", startIndex: 0, endIndex: sr } } });
    if (er < gridRows) reqs.push({ deleteDimension: { range: { sheetId: tempSheetId, dimension: "ROWS", startIndex: er, endIndex: gridRows } } });
    if (sc > 0) reqs.push({ deleteDimension: { range: { sheetId: tempSheetId, dimension: "COLUMNS", startIndex: 0, endIndex: sc } } });
    if (ec < gridCols) reqs.push({ deleteDimension: { range: { sheetId: tempSheetId, dimension: "COLUMNS", startIndex: ec, endIndex: gridCols } } });

    if (reqs.length > 0) {
      await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}:batchUpdate`,
        {
          method: "POST",
          headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
          body: JSON.stringify({ requests: reqs }),
        }
      );
    }

    console.log("Cropping done");

    // 3) EXPORT PDF -----------------------------------------------
    const exportUrl =
      `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export` +
      `?exportFormat=pdf&gid=${tempSheetId}` +
      `&portrait=${PORTRAIT}` +
      `&fitw=${FITW}` +
      `&gridlines=${GRIDLINES}` +
      `&top_margin=0&bottom_margin=0&left_margin=0&right_margin=0`;

    console.log("Export URL:", exportUrl);

    const pdfResp = await fetch(exportUrl, {
      headers: { Authorization: `Bearer ${token}` },
    });

    if (!pdfResp.ok) {
      console.error(await pdfResp.text());
      process.exit(1);
    }

    const pdfBuf = Buffer.from(await pdfResp.arrayBuffer());
    writeFileSync("report.pdf", pdfBuf);

    // 4) PDF → PNG ------------------------------------------------
    const scale = Number(SCALE_TO_PX) || 1600;
    execSync(`pdftoppm -png -singlefile -scale-to ${scale} report.pdf report`, {
      stdio: "inherit",
    });

    // 5) TRIM WHITESPACE -------------------------------------------
    let png;
    try {
      execSync(`convert report.png -fuzz 4% -trim +repage report_trim.png`, {
        stdio: "inherit",
      });
      png = readFileSync("report_trim.png");
    } catch {
      png = readFileSync("report.png");
    }

    // Shrink if too big
    const maxBytes = Number(MAX_BYTES_MB) * 1024 * 1024;
    if (png.length > maxBytes) {
      const scale2 = Math.max(600, Math.floor(scale * 0.75));
      execSync(`pdftoppm -png -singlefile -scale-to ${scale2} report.pdf report_small`, {
        stdio: "inherit",
      });

      try {
        execSync(`convert report_small.png -fuzz 4% -trim +repage report_small_trim.png`);
        png = readFileSync("report_small_trim.png");
      } catch {
        png = readFileSync("report_small.png");
      }
    }

    // 6) DELETE TEMP SHEET ----------------------------------------
    await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}:batchUpdate`,
      {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify({
          requests: [{ deleteSheet: { sheetId: tempSheetId } }],
        }),
      }
    ).catch(() => {});

    // 7) SEND TO SEATALK -------------------------------------------
    const payload = {
      tag: "file",
      file: { filename: PNG_NAME, content: png.toString("base64") },
    };

    const sea = await fetch(SEA_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    console.log("SeaTalk status:", sea.status);
    console.log("SeaTalk body:", await sea.text());

  } catch (err) {
    console.error("ERROR:", err);
    process.exit(1);
  }
})();

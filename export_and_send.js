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

function colLetterToIndex(letter) {
  let n = 0;
  for (let i = 0; i < letter.length; i++) {
    n = n * 26 + (letter.charCodeAt(i) - 64);
  }
  return n;
}

function parseA1Range(a1) {
  const parts = a1.split(":");
  function parseCell(cell) {
    const m = cell.match(/^([A-Z]+)(\d+)$/i);
    if (!m) throw new Error("Invalid A1 cell: " + cell);
    return { col: colLetterToIndex(m[1].toUpperCase()), row: Number(m[2]) };
  }
  if (parts.length === 1) {
    const c = parseCell(parts[0]);
    return { startRow: c.row, endRow: c.row, startCol: c.col, endCol: c.col };
  } else {
    const a = parseCell(parts[0]);
    const b = parseCell(parts[1]);
    return {
      startRow: Math.min(a.row, b.row),
      endRow: Math.max(a.row, b.row),
      startCol: Math.min(a.col, b.col),
      endCol: Math.max(a.col, b.col)
    };
  }
}

(async () => {
  try {
    // Giải mã SA JSON & xin access token
    const sa = JSON.parse(Buffer.from(SA_JSON_BASE64, "base64").toString("utf8"));
    const jwt = new JWT({
      email: sa.client_email,
      key: sa.private_key,
      scopes: [
        "https://www.googleapis.com/auth/drive.readonly",
        "https://www.googleapis.com/auth/spreadsheets"
      ],
    });

    const token = (await jwt.getAccessToken()).token;
    if (!token) { console.error("Failed to obtain access token"); process.exit(1); }

    const parsed = parseA1Range(RANGE_A1);
    console.log("Range parsed:", parsed);

    // 1) Duplicate sheet
    const dupName = `tmp_export_${Date.now()}`;
    const batchDupBody = {
      requests: [
        {
          duplicateSheet: {
            sourceSheetId: Number(GID),
            insertSheetIndex: 0,
            newSheetName: dupName
          }
        }
      ]
    };

    let resp = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}:batchUpdate`,
      {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify(batchDupBody)
      }
    );

    if (!resp.ok) {
      console.error("Failed to duplicate sheet:", resp.status, await resp.text());
      process.exit(1);
    }

    const dupData = await resp.json();
    const tempSheetId = dupData.replies[0].duplicateSheet.properties.sheetId;
    const gridRows = dupData.replies[0].duplicateSheet.properties.gridProperties.rowCount;
    const gridCols = dupData.replies[0].duplicateSheet.properties.gridProperties.columnCount;

    console.log("Temp sheet created:", tempSheetId);

    // 2) Crop dimension
    const requests = [];

    const startIndexRow = parsed.startRow - 1;
    const endIndexRowExclusive = parsed.endRow;

    if (startIndexRow > 0) {
      requests.push({
        deleteDimension: {
          range: {
            sheetId: tempSheetId,
            dimension: "ROWS",
            startIndex: 0,
            endIndex: startIndexRow
          }
        }
      });
    }

    if (endIndexRowExclusive < gridRows) {
      requests.push({
        deleteDimension: {
          range: {
            sheetId: tempSheetId,
            dimension: "ROWS",
            startIndex: endIndexRowExclusive,
            endIndex: gridRows
          }
        }
      });
    }

    const startIndexCol = parsed.startCol - 1;
    const endIndexColExclusive = parsed.endCol;

    if (startIndexCol > 0) {
      requests.push({
        deleteDimension: {
          range: {
            sheetId: tempSheetId,
            dimension: "COLUMNS",
            startIndex: 0,
            endIndex: startIndexCol
          }
        }
      });
    }

    if (endIndexColExclusive < gridCols) {
      requests.push({
        deleteDimension: {
          range: {
            sheetId: tempSheetId,
            dimension: "COLUMNS",
            startIndex: endIndexColExclusive,
            endIndex: gridCols
          }
        }
      });
    }

    if (requests.length > 0) {
      resp = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}:batchUpdate`,
        {
          method: "POST",
          headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
          body: JSON.stringify({ requests })
        }
      );

      if (!resp.ok) {
        console.error("Failed to crop temp sheet:", resp.status, await resp.text());
        // cleanup
        await fetch(
          `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}:batchUpdate`,
          {
            method: "POST",
            headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
            body: JSON.stringify({ requests: [{ deleteSheet: { sheetId: tempSheetId } }] })
          }
        ).catch(()=>{});
        process.exit(1);
      }
    }

    console.log("Temp sheet cropped.");

    // 3) Export PDF from temp sheet
    const exportUrl =
      `https://docs.google.com/spreadsheets/d/${encodeURIComponent(SHEET_ID)}/export` +
      `?exportFormat=pdf&gid=${encodeURIComponent(tempSheetId)}` +
      `&portrait=${PORTRAIT}` +
      `&fitw=${FITW}` +
      `&gridlines=${GRIDLINES}` +
      `&top_margin=0&bottom_margin=0&left_margin=0&right_margin=0`;

    console.log("Export URL:", exportUrl);

    const pdfResp = await fetch(exportUrl, {
      headers: { Authorization: `Bearer ${token}` }
    });

    console.log("PDF export status:", pdfResp.status);

    if (!pdfResp.ok) {
      console.error("Export PDF failed:", await pdfResp.text());
      process.exit(1);
    }

    const pdfBuf = Buffer.from(await pdfResp.arrayBuffer());
    writeFileSync("report.pdf", pdfBuf);

    // Convert PDF -> PNG
    const scale = Number(SCALE_TO_PX) || 1600;
    execSync(`pdftoppm -png -singlefile -scale-to ${scale} report.pdf report`, {
      stdio: "inherit"
    });

    let png = readFileSync("report.png");

    // Shrink if too big
    const maxBytes = (Number(MAX_BYTES_MB)||5) * 1024 * 1024;
    if (png.length > maxBytes) {
      const scale2 = Math.max(800, Math.floor(scale * 0.75));
      console.log(`PNG too big; retry with scale-to=${scale2}`);
      execSync(`pdftoppm -png -singlefile -scale-to ${scale2} report.pdf report`, {
        stdio: "inherit"
      });
      png = readFileSync("report.png");
    }

    // Cleanup temp sheet
    await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}:batchUpdate`,
      {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify({ requests: [{ deleteSheet: { sheetId: tempSheetId } }] })
      }
    ).catch(err => {
      console.warn("Failed to delete temp sheet:", err);
    });

    // Send PNG to SeaTalk
    const payload = {
      tag: "file",
      file: { filename: PNG_NAME, content: Buffer.from(png).toString("base64") }
    };

    const sea = await fetch(SEA_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    console.log("SeaTalk status:", sea.status);
    console.log("SeaTalk body:", await sea.text());

  } catch (e) {
    console.error(e);
    process.exit(1);
  }
})();

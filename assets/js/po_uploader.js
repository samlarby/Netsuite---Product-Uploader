document.addEventListener("DOMContentLoaded", () => {
  console.log("PO Upload script loaded (v2)");

  // ---- Defaults for the 2nd (header) CSV ----
  // Change these if you need:
  const DEFAULT_SITE = "PrimarySite";

  // Always today (Europe/London) in YYYY-MM-DD
  function londonTodayISO() {
    const parts = new Intl.DateTimeFormat("en-GB", {
      timeZone: "Europe/London",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
    }).formatToParts(new Date());

    const get = (type) => parts.find((p) => p.type === type)?.value || "";
    return `${get("year")}-${get("month")}-${get("day")}`;
  }

  const generateBtn = document.getElementById("generateBtn");
  generateBtn.addEventListener("click", () => {
    console.log("Generate button clicked");
    const statusEl = document.getElementById("status");
    const poTextEl = document.getElementById("poText");

    statusEl.textContent = "";

    const poText = (poTextEl.value || "").trim();
    if (!poText) {
      statusEl.innerHTML =
        "<strong>Error:</strong> Please paste PO data into the box.";
      return;
    }

    try {
      statusEl.textContent = "Parsing pasted data...";

      const poRows = parsePastedTable(poText);
      if (!poRows.length) {
        statusEl.innerHTML =
          "<strong>Error:</strong> Could not find any rows in the pasted data.";
        return;
      }

      const uploadRows = buildUploadRows(poRows);
      if (!uploadRows.length) {
        statusEl.innerHTML =
          "<strong>Done:</strong> No valid rows to export (or all cancelled / zero qty).";
        return;
      }

      const headerRows = buildHeaderRows(poRows);

      downloadCSV(toCSV(uploadRows), "Upload_Lines.csv");
      downloadCSV(toCSV(headerRows), "Upload_Headers.csv");

      statusEl.innerHTML =
        "<strong>Success:</strong> Generated Upload_Lines.csv (" +
        uploadRows.length +
        " rows) and Upload_Headers.csv (" +
        headerRows.length +
        " rows).";
    } catch (err) {
      console.error(err);
      statusEl.innerHTML =
        "<strong>Error:</strong> " +
        (err.message || "Something went wrong while processing the data.");
    }
  });

  // ---------- Parse pasted table ----------
  // Supports:
  // 1) Tab-separated (Excel paste)
  // 2) Proper CSV with quoted fields
  function parsePastedTable(text) {
    const rawLines = text.split(/\r?\n/).filter((l) => l.trim() !== "");
    if (!rawLines.length) return [];

    const headerLine = rawLines[0];
    const isTSV = headerLine.includes("\t");

    let headers = [];
    const dataLines = rawLines.slice(1);

    if (isTSV) {
      headers = headerLine.split("\t").map((h) => h.trim());
      return dataLines.map((line) => {
        const parts = line.split("\t");
        const rowObj = {};
        headers.forEach((h, idx) => (rowObj[h] = (parts[idx] || "").trim()));
        return rowObj;
      });
    }

    headers = parseCSVLine(headerLine).map((h) => h.trim());
    return dataLines.map((line) => {
      const parts = parseCSVLine(line);
      const rowObj = {};
      headers.forEach((h, idx) => (rowObj[h] = (parts[idx] ?? "").trim()));
      return rowObj;
    });
  }

  // Minimal single-line CSV parser (handles commas inside quotes and escaped quotes "")
  function parseCSVLine(line) {
    const out = [];
    let cur = "";
    let inQuotes = false;

    for (let i = 0; i < line.length; i++) {
      const ch = line[i];

      if (inQuotes) {
        if (ch === '"') {
          if (line[i + 1] === '"') {
            cur += '"';
            i++;
          } else {
            inQuotes = false;
          }
        } else {
          cur += ch;
        }
      } else {
        if (ch === '"') {
          inQuotes = true;
        } else if (ch === ",") {
          out.push(cur);
          cur = "";
        } else {
          cur += ch;
        }
      }
    }
    out.push(cur);
    return out;
  }

  // ---------- Helpers ----------
  function norm(str) {
    return String(str || "").trim();
  }
  function normUpper(str) {
    return norm(str).toUpperCase();
  }

  function toProper(str) {
    str = String(str || "").toLowerCase();
    return str.replace(/\b\w+/g, (w) => w.charAt(0).toUpperCase() + w.slice(1));
  }

  function parseNumber(val) {
    const s = String(val ?? "")
      .replace(/,/g, "")
      .trim();
    const n = parseFloat(s);
    return Number.isFinite(n) ? n : NaN;
  }

  function formatExpectedDate(val) {
    return val ? String(val).trim() : "";
  }

  function expectedDateWithCurrentYear(val) {
    if (!val || !String(val).trim()) return "";

    const raw = String(val).trim();

    // Match formats like: 30-Dec, 1-Jan, 05-Feb
    const match = raw.match(/^(\d{1,2})[-\s]?([A-Za-z]{3})$/);
    if (!match) return "";

    const day = match[1].padStart(2, "0");
    const monthText = match[2].toLowerCase();

    const monthMap = {
      jan: "01",
      feb: "02",
      mar: "03",
      apr: "04",
      may: "05",
      jun: "06",
      jul: "07",
      aug: "08",
      sep: "09",
      oct: "10",
      nov: "11",
      dec: "12",
    };

    const month = monthMap[monthText];
    if (!month) return "";

    const year = new Intl.DateTimeFormat("en-GB", {
      timeZone: "Europe/London",
      year: "numeric",
    }).format(new Date());

    return `${day}/${month}/${year}`;
  }

  function supplierLabel(raw) {
    const u = normUpper(raw);
    if (u === "AMPLEBOX" || u === "AMPLEBOX LIMITED") return "SUP00030";
    if (u === "SJA FASHION" || "SJA") return "SUP00243";
    if (u === "DP") return "SUP00355";
    if (u === "GRAND APPARELS") return "SUP00130";
    if (u === "RAGTEKS") return "SUP00354";
    if (u === "ERSIN") return "SUP00361";
    if (u === "FLOMAK") return "SUP00363";
    if (u === "LI & FUNG") return "Li & Fung";
    if (u === "LUCKY MONDAY") return "SUP00403";
    if (u === "WETEX") return "SUP00302";
    if (u === "SKYLAND") return "SUP00356";
    if (u === "WELLSUCCEED") return "SUP00300";
    if (u === "ELEANOLA") return "SUP00328";
    return toProper(raw);
  }

  // ---------- Header Mapping ----------
  function guessHeaders(sampleRow) {
    const map = {};
    const entries = Object.keys(sampleRow || {});

    function findKey(possible) {
      const target = possible.map((p) => p.toLowerCase());
      return entries.find((k) => target.includes(k.toLowerCase())) || "";
    }

    map.PO = findKey([
      "PO",
      "PO NUMBER",
      "PONUMBER",
      "PURCHASE ORDER",
      "PURCHASE ORDER NUMBER",
    ]);
    map.DESCRIPTION = findKey([
      "DESCRIPTION",
      "PRODUCT DESCRIPTION",
      "ITEM DESCRIPTION",
    ]);
    map.SUPPLIER = findKey(["SUPPLIER", "VENDOR", "FACTORY"]);
    map.STATUS = findKey(["STATUS", "PO STATUS"]);

    // New file fields
    map.UNIT_COST = findKey([
      "UNIT COST PRICE (GBP)",
      "UNIT COST GBP",
      "UNIT COST",
      "COST",
      "UNIT PRICE",
      "PRICE",
    ]);
    map.SKU_VAR = findKey(["SKU VAR", "SKU_VAR", "VARIANT"]);
    map.SIZE = findKey(["SIZE", " SIZE "]);
    map.UNITS = findKey(["UNITS", " UNITS "]);
    map.ITEM_CODE = findKey(["ITEM CODE", "ITEMCODE", "ITEM_CODE"]);

    map.DELIVERY_DATE = findKey(["DELIVERY DATE (ACTUAL)", "DELIVERY DATE"]);
    map.HANDOVER_DATE = findKey(["HANDOVER DATE (ACTUAL)", "HANDOVER DATE"]);

    // Optional (if you add these columns later)
    map.SITE = findKey(["SITE"]);
    map.SUBMITTED_DATE = findKey([
      "SUBMITTED DATE",
      "SUBMITTEDDATE",
      "DATE SUBMITTED",
    ]);

    return map;
  }

  function validateRequiredHeaders(map) {
    const required = ["PO", "SUPPLIER", "STATUS", "SIZE", "UNITS"];
    const missing = required.filter((k) => !map[k]);
    if (missing.length) {
      throw new Error(
        "Missing required column(s): " +
          missing.join(", ") +
          ". Check the pasted header row matches expected names.",
      );
    }
  }

  // ---------- Main Transform: Lines (one row per pasted line / size) ----------
  function buildUploadRows(poRows) {
    const headersMap = guessHeaders(poRows[0] || {});
    validateRequiredHeaders(headersMap);

    const out = [];

    for (const row of poRows) {
      const po = norm(row[headersMap.PO]);
      if (!po) continue;

      const status = normUpper(row[headersMap.STATUS]);
      if (status.startsWith("CANCEL")) continue;

      const qty = parseNumber(row[headersMap.UNITS]);
      if (!Number.isFinite(qty) || qty <= 0) continue;

      const rate = parseNumber(row[headersMap.UNIT_COST]);
      const amount = Number.isFinite(rate) ? +(qty * rate).toFixed(2) : "";

      const itemCode = norm(row[headersMap.ITEM_CODE]);

      out.push({
        PurchaseOrderNumber: po, // PO ONLY (no description appended)
        Status: "Submitted",
        ItemCode: itemCode || "*** Missing Item Code ***",
        Quantity: qty,
        CostPrice: Number.isFinite(rate) ? rate : "",
        Amount: amount !== "" ? amount : "",
        Taxcode: "VAT:20% - S-GB",
      });
    }

    return out;
  }

  // ---------- Second CSV: Headers (one row per PO) ----------
  function buildHeaderRows(poRows) {
    const headersMap = guessHeaders(poRows[0] || {});
    validateRequiredHeaders(headersMap);

    const byPO = new Map();

    for (const row of poRows) {
      const po = norm(row[headersMap.PO]);
      if (!po) continue;

      const status = normUpper(row[headersMap.STATUS]);
      if (status.startsWith("CANCEL")) continue;

      const supplierRaw = norm(row[headersMap.SUPPLIER]);
      const supplier = supplierLabel(supplierRaw);

      const expectedDate = expectedDateWithCurrentYear(
        row[headersMap.DELIVERY_DATE] || row[headersMap.HANDOVER_DATE],
      );

      const site = headersMap.SITE ? norm(row[headersMap.SITE]) : "";
      const submittedDate = londonTodayISO();

      if (!byPO.has(po)) {
        byPO.set(po, {
          PurchaseOrderNumber: po,
          Supplier: supplier,
          ExpectedDeliveryDate: expectedDate,
          Site: site || DEFAULT_SITE,
          SubmittedDate: londonTodayISO(),
        });
      } else {
        // If we later see a non-empty ExpectedDeliveryDate, keep it
        const existing = byPO.get(po);
        if (!existing.ExpectedDeliveryDate && expectedDate)
          existing.ExpectedDeliveryDate = expectedDate;

        // Fill site/submitted if missing
        if (!existing.Site && site) existing.Site = site;
        if (!existing.SubmittedDate && submittedDate)
          existing.SubmittedDate = submittedDate;
      }
    }

    return Array.from(byPO.values());
  }

  // ---------- CSV ----------
  function toCSV(rows) {
    if (!rows.length) return "";
    const headers = Object.keys(rows[0]);
    const escape = (v) => {
      if (v === null || v === undefined) return "";
      v = String(v);
      if (v.includes('"') || v.includes(",") || v.includes("\n")) {
        return `"${v.replace(/"/g, '""')}"`;
      }
      return v;
    };
    return [
      headers.join(","),
      ...rows.map((row) => headers.map((h) => escape(row[h])).join(",")),
    ].join("\r\n");
  }

  function downloadCSV(csv, filename) {
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  }
});
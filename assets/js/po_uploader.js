document.addEventListener("DOMContentLoaded", () => {
  console.log("PO Upload script loaded");

  const sizeColumns = ["ONE SIZE", "XXS", "XS", "S", "M", "L", "XL", "2XL", "3XL"];

  const generateBtn = document.getElementById("generateBtn");
  generateBtn.addEventListener("click", async () => {
    console.log("Generate button clicked");
    const poFile = document.getElementById("poFile").files[0];
    const statusEl = document.getElementById("status");
    statusEl.textContent = "";

    if (!poFile) {
      statusEl.innerHTML = "<strong>Error:</strong> Please select a PO planning file.";
      return;
    }

    try {
      statusEl.textContent = "Reading file...";
      const poRows = await readWorkbook(poFile);

      const uploadRows = buildUploadRows(poRows);

      if (!uploadRows.length) {
        statusEl.innerHTML = "<strong>Done:</strong> No valid rows to export (or all cancelled / zero qty).";
        return;
      }

      const csv = toCSV(uploadRows);
      downloadCSV(csv, "Upload.csv");
      statusEl.innerHTML =
        "<strong>Success:</strong> Generated Upload.csv with " + uploadRows.length + " rows.";
    } catch (err) {
      console.error(err);
      statusEl.innerHTML =
        "<strong>Error:</strong> " + (err.message || "Something went wrong while processing the file.");
    }
  });

  // ---------- File Reading ----------
  function readWorkbook(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const wb = XLSX.read(data, { type: "array" });
          const sheetName = wb.SheetNames[0];
          const sheet = wb.Sheets[sheetName];
          const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
          resolve(json);
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = (e) => reject(e);
      reader.readAsArrayBuffer(file);
    });
  }

  // ---------- Helpers ----------
  function norm(str) { return String(str || "").trim(); }
  function normUpper(str) { return norm(str).toUpperCase(); }
  function toProper(str) {
    str = String(str || "").toLowerCase();
    return str.replace(/\b\w+/g, w => w.charAt(0).toUpperCase() + w.slice(1));
  }
  function variantLabel(raw) {
    const v = normUpper(raw);
    if (v === "" || v === "REG") return "Regular";
    if (v === "TALL") return "Tall";
    if (v === "PETITE") return "Petite";
    return norm(raw);
  }
  function displaySize(sizeName) {
    const v = normUpper(sizeName);
    if (v === "XXS") return "2XS";
    return sizeName;
  }
  function supplierLabel(raw) {
    const u = normUpper(raw);
    if (u === "AMPLEBOX" || u === "AMPLEBOX LIMITED") {
      return "SUP00030 Amplebox Limited";
    }
    else if (u === "SJA FASHION") {
      return "SUP00243 Sja Fashion";
    }
    else if (u === "DP") {
      return "	SUP00355 DP Limited";
    }
    else if (u === "GRAND APPARELS") {
      return "SUP00130 Grand Apparels Designs";
    }
    else if (u === "RAGTEKS") {
      return "SUP00354 Ragteks";
    }
    else if (u === "ERSIN") {
      return "SUP00361 Ersin Limited"
    }
    else if (u === "FLOMAK") {
      return "SUP00363 Flomak Group"
    }
    else if (u === "LI & FUNG") {
      return "*******"
    }
    else if (u === "LUCKY MONDAY") {
      return "*******"
    }
    else if (u === "WETEX") {
      return "SUP00302 Wetex Fashion UK"
    }
    else if (u === "SKYLAND") {
      return "SUP00356 Skyland"
    }
    else if (u === "WELLSUCCEED") {
      return "SUP00300 Wellsucceed Embroidery Limited"
    }
    else if (u === "ELEANOLA") {
      return "SUP00328 Eleanola Inc"
    }
    return toProper(raw);
  }
  function formatToday() {
    const d = new Date();
    return `${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}/${d.getFullYear()}`;
  }
  function formatExpectedDate(val) {
    if (!val) return "";
    if (typeof val === "string") return val.trim();
    if (typeof val === "number") {
      const jsDate = XLSX.SSF.parse_date_code(val);
      if (!jsDate) return "";
      const d = new Date(jsDate.y, jsDate.m - 1, jsDate.d);
      return `${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}/${d.getFullYear()}`;
    }
    return "";
  }

  // ---------- Main Transform ----------
  function buildUploadRows(poRows) {
    const headersMap = guessHeaders(poRows[0] || {});
    const todayStr = formatToday();
    const out = [];
    let lastExtBase = "";
    let orderline = 0;

    for (const row of poRows) {
      const po = norm(row[headersMap.PO]);
      if (!po) continue;

      const status = normUpper(row[headersMap.STATUS]);
      if (status.startsWith("CANCEL")) continue;

      const descRaw = norm(row[headersMap.DESCRIPTION]);
      const styleCode = norm(row[headersMap.STYLE_CODE]);
      const supplierRaw = norm(row[headersMap.SUPPLIER]);
      const skuVar = norm(row[headersMap.SKU_VAR]);
      const currRaw = norm(row[headersMap.CURRENCY]);
      const rateRaw = row[headersMap.UNIT_COST];
      const planDateRaw = row[headersMap.PLANNED_WC];

      const currency = currRaw || "GBP";
      const rate = parseFloat(rateRaw) || 0;
      const supplier = supplierLabel(supplierRaw);

      let extBase = `${po} ${styleCode}`;
      if (skuVar && skuVar !== "0") {
        extBase += ` (${variantLabel(skuVar)})`;
      }

      let memo = toProper(descRaw);
      if (skuVar && skuVar !== "0") {
        memo += ` (${variantLabel(skuVar)})`;
      }

      const expectedDate = formatExpectedDate(planDateRaw);

      for (const sizeCol of sizeColumns) {
        const qtyRaw = row[sizeCol] ?? row[sizeCol.toUpperCase()] ?? row[sizeCol.toLowerCase()];
        const qty = parseFloat(qtyRaw);
        if (!qty || isNaN(qty) || qty <= 0) continue;

        if (extBase !== lastExtBase) {
          orderline = 1;
          lastExtBase = extBase;
        } else {
          orderline += 1;
        }

        const dispSize = displaySize(sizeCol);
        const partNum = `${styleCode} : ${skuVar || "REG"}_${dispSize}`;

        let partDesc = toProper(descRaw);
        if (skuVar && skuVar !== "0") partDesc += ` (${variantLabel(skuVar)})`;
        partDesc += sizeCol.toUpperCase() !== "ONE SIZE" ? ` ${dispSize}` : ` One Size`;

        const amount = rate ? qty * rate : "";

        out.push({
          "ExternalID": extBase,
          "Orderline": orderline,
          "PartNumber": "*** Add parent and child item ***",
          "PartDescription": partDesc,
          "Quantity": qty,
          "PO": po,
          "Date": todayStr,
          "Supplier": supplier,
          "Subsidiary": "8",
          "Department": "Product",
          "Currency": currency,
          "Exchange Rate": 1,
          "Rate": rate || "",
          "Amount": amount !== "" ? amount : "",
          "Taxcode": "VAT:20% - S-GB",
          "Expected Date": expectedDate,
          "Memo": memo
        });
      }
    }

    return out;
  }

  // ---------- Header Mapping ----------
  function guessHeaders(sampleRow) {
    const map = {};
    const entries = Object.keys(sampleRow || {});

    function findKey(possible) {
      const target = possible.map(p => p.toLowerCase());
      return entries.find(k => target.includes(k.toLowerCase())) || "";
    }

    map.DESCRIPTION = findKey(["DESCRIPTION"]);
    map.STYLE_CODE = findKey(["STYLE CODE", "STYLE_CODE", "STYLE"]);
    map.PO = findKey(["PO", "PO NUMBER", "PONUMBER"]);
    map.PLANNED_WC = findKey(["PLANNED DELIVERY WC", "PLANNED DELIVERY", "DELIVERY WEEK"]);
    map.SUPPLIER = findKey(["SUPPLIER"]);
    map.CURRENCY = findKey(["CURRENCY"]);
    map.UNIT_COST = findKey(["UNIT COST GBP", "UNIT COST", "COST"]);
    map.SKU_VAR = findKey(["SKU VAR", "SKU_VAR", "VARIANT"]);
    map.STATUS = findKey(["STATUS"]);

    return map;
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
    return [headers.join(","), ...rows.map(row => headers.map(h => escape(row[h])).join(","))].join("\r\n");
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

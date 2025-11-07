const sizeColumns = ["ONE SIZE", "XXS", "XS", "S", "M", "L", "XL", "2XL", "3XL"];

document.getElementById("generateBtn").addEventListener("click", async () => {
  const poFile = document.getElementById("poFile").files[0];
  const partFile = document.getElementById("partFile").files[0];
  const statusEl = document.getElementById("status");
  statusEl.textContent = "";

  if (!poFile) {
    statusEl.innerHTML = "<strong>Error:</strong> Please select a PO planning file.";
    return;
  }

  try {
    statusEl.textContent = "Reading files...";
    const poRows = await readWorkbook(poFile);

    let partDict = {};
    if (partFile) {
      const partRows = await readWorkbook(partFile);
      partDict = buildPartDictionary(partRows);
    }

    const uploadRows = buildUploadRows(poRows, partDict);

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
      "<strong>Error:</strong> " + (err.message || "Something went wrong while processing the files.");
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

// ---------- Helpers: Normalization & Formatting ----------

function norm(str) {
  return String(str || "").trim();
}

function normUpper(str) {
  return norm(str).toUpperCase();
}

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
  const nm = toProper(raw);
  return nm;
}

function formatToday() {
  const d = new Date();
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yyyy = d.getFullYear();
  return `${dd}/${mm}/${yyyy}`; // UK style to match your examples
}

function formatExpectedDate(val) {
  if (val === null || val === undefined || val === "") return "";
  // If already a string, return as-is
  if (typeof val === "string") return val.trim();
  // If numeric (Excel date)
  if (typeof val === "number") {
    const jsDate = XLSX.SSF.parse_date_code(val);
    if (!jsDate) return "";
    const d = new Date(jsDate.y, jsDate.m - 1, jsDate.d);
    const dd = String(d.getDate()).padStart(2, "0");
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const yyyy = d.getFullYear();
    return `${dd}/${mm}/${yyyy}`;
  }
  return "";
}

// ---------- PartNumber Dictionary ----------

function buildKey(styleCode, skuVar, sizeName) {
  const style = normUpper(styleCode);
  const variant = normUpper(variantLabel(skuVar));
  const size = normUpper(sizeName);
  return `${style}|${variant}|${size}`;
}

function buildPartDictionary(rows) {
  const dict = {};
  rows.forEach(row => {
    // Flexible header matching
    const style = row["STYLE CODE"] ?? row["Style Code"] ?? row["stylecode"] ?? row["style_code"];
    const skuVar = row["SKU VAR"] ?? row["Sku Var"] ?? row["skuvar"] ?? row["SKU"] ?? "";
    const size = row["SIZE"] ?? row["Size"] ?? "";
    const pn = row["PARTNUMBER"] ?? row["PartNumber"] ?? row["Part Number"] ?? "";

    const styleN = norm(style);
    const sizeN = norm(size);
    const pnN = norm(pn);
    if (!styleN || !sizeN || !pnN) return;

    const key = buildKey(styleN, skuVar, sizeN);
    if (!dict[key]) {
      dict[key] = pnN;
    }
  });
  return dict;
}

function getPartNumberFromDict(dict, styleCode, skuVar, sizeName) {
  if (!dict) return "";
  const key = buildKey(styleCode, skuVar, sizeName);
  return dict[key] || "";
}

function buildFallbackPartNumber(styleCode, skuVar, sizeName) {
  const base = norm(styleCode);
  const sizeDisp = displaySize(sizeName);
  const v = norm(skuVar);
  if (v && v !== "0") {
    return `${base} : ${v}_${sizeDisp}`;
  }
  return `${base} : ${sizeDisp}`;
}

// ---------- Core Transform ----------

function buildUploadRows(poRows, partDict) {
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

    // ExternalID: PO + Style Code + (Variant)
    let extBase = `${po} ${styleCode}`;
    if (skuVar && skuVar !== "0") {
      extBase += ` (${variantLabel(skuVar)})`;
    }

    // Memo: proper case desc + (Variant)
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
      let partNum = getPartNumberFromDict(partDict, styleCode, skuVar, sizeCol);
      if (!partNum) {
        partNum = buildFallbackPartNumber(styleCode, skuVar, dispSize);
      }

      const properDesc = toProper(descRaw);
      let partDesc = properDesc;
      if (skuVar && skuVar !== "0") {
        partDesc += ` (${variantLabel(skuVar)})`;
      }
      if (sizeCol.toUpperCase() !== "ONE SIZE") {
        partDesc += ` ${dispSize}`;
      } else {
        partDesc += ` One Size`;
      }

      const amount = rate ? (qty * rate) : "";

      out.push({
        "ExternalID": extBase,
        "Orderline": orderline,
        "PartNumber": partNum,
        "PartDescription": partDesc,
        "Quantity": qty,
        "PO": po,
        "Date": todayStr,
        "Supplier": supplier,
        "Subsidiary": "8",
        "Department": "Product",
        "Location": "SSUK Warehouse : SSUK - Ecomm",
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

function guessHeaders(sampleRow) {
  // Map logical names to actual keys with some flexibility:
  const map = {};
  const entries = Object.keys(sampleRow || {});

  function findKey(possible) {
    const target = possible.map(p => p.toLowerCase());
    return (
      entries.find(k => target.includes(k.toLowerCase())) || ""
    );
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

// ---------- CSV Download ----------

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
  const lines = [];
  lines.push(headers.join(","));
  for (const row of rows) {
    lines.push(headers.map(h => escape(row[h])).join(","));
  }
  return lines.join("\r\n");
}

function downloadCSV(csv, filename) {
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
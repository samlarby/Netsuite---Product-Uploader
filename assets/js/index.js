/* ---------- utils ---------- */
const byId = id => document.getElementById(id);
const clean = s => (s??"").toString().replace(/\u00A0/g,' ').normalize("NFKC").trim();
const cleanKey = s => clean(s).replace(/[‐–—−]/g,'-').replace(/\s+/g,' ').toUpperCase();

function parentFromSku(sku){
  return clean(sku)
    .replace(/^[PRT]/i, '')   // remove Petite/Regular/Tall prefix
    .replace(/\d+$/, '')      // remove trailing number index
    .replace(/[ _-]$/, '');   // tidy trailing separator
}

/* find a header by candidate names (case-insensitive/loose) */
function findHeader(rows, candidates){
  if(!rows.length) return null;
  const headers = Object.keys(rows[0]);
  const norm = s => s.toLowerCase().replace(/[^a-z0-9]/g,'');
  for(const cand of candidates){
    const want = norm(cand);
    for(const h of headers){ const n=norm(h); if(n===want || n.includes(want)) return h; }
  }
  return null;
}

// Remove one trailing parenthetical e.g. "NAME (PETITE)" -> "NAME"
function stripVariantSuffix(name){
  return clean(name).replace(/\s*\([^)]+\)\s*$/,'').trim();
}

/* forward-fill a column (handles merged-cell exports) */
function ffillColumn(rows, col){
  if(!col) return;
  let last="";
  for(const r of rows){
    const v = clean(r[col]);
    if(v){ last=v; r[col]=v; } else { r[col]=last; }
  }
}

/* robust CSV/TSV parser (header-based delimiter + quote aware) */
function parseTable(text){
  text = (text||"").trim();
  if (!text) return [];
  const firstLine = text.split(/\r?\n/, 1)[0];
  let delim = '\t';
  const tabCount = (firstLine.match(/\t/g)||[]).length;
  const commaCount = (firstLine.match(/,/g)||[]).length;
  const semiCount = (firstLine.match(/;/g)||[]).length;
  if (tabCount > 0) delim = '\t';
  else if (commaCount >= semiCount) delim = ',';
  else delim = ';';

  const rows = [];
  let row = [], field = "", i = 0, inQ = false;
  const n = text.length;
  while (i < n){
    const c = text[i++];
    if (c === '"'){
      if (inQ && text[i] === '"'){ field += '"'; i++; }
      else inQ = !inQ;
    } else if (c === delim && !inQ){
      row.push(field); field = "";
    } else if ((c === '\n' || c === '\r') && !inQ){
      if (c === '\r' && text[i] === '\n') i++;
      row.push(field); field = "";
      if (row.length) rows.push(row);
      row = [];
    } else {
      field += c;
    }
  }
  if (field.length || row.length){ row.push(field); rows.push(row); }

  if (!rows.length) return [];
  const header = rows.shift().map(h => clean(h));
  return rows
    .filter(r => r.some(x => clean(x) !== ""))
    .map(r => {
      const o = {};
      header.forEach((h, idx) => { o[h] = clean(r[idx] ?? ""); });
      return o;
    });
}

function readFileAsText(file){
  return new Promise((res,rej)=>{
    const fr=new FileReader();
    fr.onerror=()=>rej(fr.error);
    fr.onload=()=>res(fr.result);
    fr.readAsText(file);
  });
}

function toCSV(rows){
  if(!rows.length) return '';
  const headers=Object.keys(rows[0]);
  const esc=v=>{const s=(v==null?'':String(v)); return /[",\n]/.test(s)?`"${s.replace(/"/g,'""')}"`:s;};
  const lines=[headers.map(esc).join(',')];
  for(const r of rows) lines.push(headers.map(h=>esc(r[h])).join(','));
  return '\uFEFF'+lines.join('\n');
}

function categoryFrom(t, brand = "Sisters and Seekers"){
  t = (t||'').toLowerCase();
  
  const b = String(brand||"").trim().toUpperCase();
  brand = (b === "BK" || /BROTHER|KIN/.test(b))
    ? "Brother & Kin"
    : "SistersandSeekers";
  // Safety: ensure correct brand formatting

  // ACCESSORIES
  if(/bag|purse|tote/.test(t))
    return `${brand} : Accessories : Accessories : Bags`;
  if(/hat|beanie|cap/.test(t))
    return `${brand} : Accessories : Accessories : Hats`;
  if(/hair|scrunchie|clip|claw/.test(t))
    return `${brand} : Accessories : Accessories : Hair Accessories`;
  if(/tight|hosiery|stocking/.test(t))
    return `${brand} : Accessories : Accessories : Hosiery`;
  if(/jewel|necklace|bracelet|earring|ring|chain/.test(t))
    return `${brand} : Accessories : Accessories : Jewellery`;
  if(/sunglass|shades/.test(t))
    return `${brand} : Accessories : Accessories : Sunglasses`;
  if(/sock/.test(t))
    return `${brand} : Accessories : Accessories : Socks`;
  if(/scarf|glove|snood|balaclava/.test(t))
    return `${brand} : Accessories : Accessories : Winter Accessories`;

  // BOTTOMS
  if(/skirt/.test(t))
    return `${brand} : Clothing : Bottoms : Skirts`;
  if(/shorts?/.test(t))
    return `${brand} : Clothing : Bottoms : Shorts`;
  if(/sweatpant|jogger|track pant|wide leg/.test(t))
    return `${brand} : Clothing : Bottoms : Sweatpants`;
  if(/legging/.test(t))
    return `${brand} : Clothing : Bottoms : Leggings`;
  if(/trouser|pant|(cargo)|chino/.test(t))
    return `${brand} : Clothing : Bottoms : Trousers`;

  // DRESSES & JEANS (note: Brother & Kin has no dresses, fallback later still safe)
  if(/dress/.test(t))
    return `${brand} : Clothing : Dresses : Dresses`;
  if(/jean|denim/.test(t))
    return `${brand} : Clothing : Jeans : Jeans`;

  // KNITWEAR
  if(/cardi/.test(t))
    return `${brand} : Clothing : Knitwear : Cardigans`;
  if(/knitted hoodie|knit hoodie/.test(t))
    return `${brand} : Clothing : Knitwear : Knitted Hoodies`;
  if(/jumper|sweater|knit/.test(t))
    return `${brand} : Clothing : Knitwear : Jumpers`;

  // NIGHTWEAR
  if(/pyjama set/.test(t))
    return `${brand} : Clothing : Nightwear : Pyjama Set`;
  if(/pyjama top|pj tee|pj top/.test(t))
    return `${brand} : Clothing : Nightwear : Pyjama Top`;
  if(/pyjama bottom|pj pants/.test(t))
    return `${brand} : Clothing : Nightwear : Pyjama Bottoms`;
  if(/robe|dressing gown/.test(t))
    return `${brand} : Clothing : Nightwear : Robe`;

  // OUTERWEAR
  if(/blazer/.test(t))
    return `${brand} : Clothing : Outerwear : Blazers`;
  if(/coat|parka|trench|mac/.test(t))
    return `${brand} : Clothing : Outerwear : Coats`;
  if(/gilet|vest/.test(t))
    return `${brand} : Clothing : Outerwear : Gilets`;
  if(/jacket|bomber|puffer|windbreaker/.test(t))
    return `${brand} : Clothing : Outerwear : Jackets`;

  // OUTFITS
  if(/all in one|jumpsuit|playsuit|unitard|romper/.test(t))
    return `${brand} : Clothing : Outfits : All in Ones`;

  // SWIMWEAR
  if(brand === "Brother & Kin" && /swim short|boardshort/.test(t))
    return `Brother & Kin : Clothing : Swimwear : Swim Shorts`;
  if(/bikini set|swim set/.test(t))
    return `${brand} : Clothing : Swimwear : Sets`;
  if(/bikini bottom|swim brief/.test(t))
    return `${brand} : Clothing : Swimwear : Bikini Bottom`;
  if(/bikini top|swim top/.test(t))
    return `${brand} : Clothing : Swimwear : Bikini Top`;
  if(/swimsuit|one piece/.test(t))
    return `${brand} : Clothing : Swimwear : Swimsuit`;

  // TOPS
  if(/hoodie|zip up/.test(t))
    return `${brand} : Clothing : Tops : Hoodies`;
  if(/polo/.test(t))
    return `${brand} : Clothing : Tops : Polos`;
  if(/shirt|button down|oxford/.test(t))
    return `${brand} : Clothing : Tops : Shirts`;
  if(/sweater|jumper|crew neck|crewneck/.test(t))
    return `${brand} : Clothing : Tops : Sweaters`;
  if(/tee|t-shirt|longsleeve|top|ls/.test(t))
    return `${brand} : Clothing : Tops : T-Shirts`;

  // UNDERWEAR
  if(/boxer/.test(t))
    return `${brand} : Clothing : Underwear : Boxers`;
  if(/bra|bralette/.test(t))
    return `${brand} : Clothing : Underwear : Bra`;
  if(/knicker|panty|thong|brief/.test(t))
    return `${brand} : Clothing : Underwear : Knicker`;

  // FOOTWEAR
  if(/boot/.test(t))
    return `${brand} : Footwear : Footwear : Boots`;
  if(/sandal|slider|flip flop/.test(t))
    return `${brand} : Footwear : Footwear : Sandals`;
  if(/trainer|sneaker|runner/.test(t))
    return `${brand} : Footwear : Footwear : Trainers`;
  if(/shoe|loafer|flat|heel/.test(t))
    return `${brand} : Footwear : Footwear : Shoes`;

  // FALLBACK
  return `***Add hierarchy***`;
}

function formatCOO(raw) {
  if (!raw) return "";
  let s = String(raw).trim().toLowerCase();

  // normalise punctuation and spacing
  s = s.replace(/\./g, "").replace(/\s+/g, " ");

  // handle Turkey special case
  if (s === "turkey") return "Türkiye";
  if (s === "uk") return "United Kingdom"; 
  

  // capitalize only the first letter
  return s.charAt(0).toUpperCase() + s.slice(1);
}

function formatName(s) {
  if (!s) return "";
  s = String(s).trim().toLowerCase();
  // Capitalize at start of string, after space, slash, or hyphen
  s = s.replace(/(^|[ \/\-])([a-z])/g, (m, p1, p2) => p1 + p2.toUpperCase());
  return s;
}

function upcaseSizeTokens(s){
  return String(s||"").replace(/\b(2xs|xs|s|m|l|xl|2xl|3xl)\b/gi, m => m.toUpperCase());
}

function isOneSize(size){
  const s = String(size||"").trim().toUpperCase();
  return s === "ONE SIZE" || s === "OS";
}


/* ---------- state ---------- */
let namesRows=[], trackerRows=[];

byId('namesFile').addEventListener('change', async e=>{
  const f=e.target.files[0]; if(!f) return;
  const txt=await readFileAsText(f);
  namesRows=parseTable(txt);

  // forward-fill NAME & HS so each size row carries them
  const NAMECOL = findHeader(namesRows, ["NAME"]);
  const HSCOL   = findHeader(namesRows, ["HS Code","HSCode"]);
  ffillColumn(namesRows, NAMECOL);
  ffillColumn(namesRows, HSCOL);

  byId('namesStatus').innerHTML=`Loaded <span class="ok">${namesRows.length}</span> rows (NAME & HS forward-filled)`;
});

byId('trackerFile').addEventListener('change', async e=>{
  const f=e.target.files[0]; if(!f) return;
  const txt=await readFileAsText(f);
  trackerRows=parseTable(txt);
  byId('trackerStatus').innerHTML=`Loaded <span class="ok">${trackerRows.length}</span> rows`;
});

/* ---------- main generate ---------- */
byId('genBtn').addEventListener('click', ()=>{
  byId('dlLink').style.display='none'; byId('previewCard').style.display='none';
  if(!namesRows.length || !trackerRows.length){
    byId('log').innerHTML='<span class="err">Please upload both CSV files first.</span>';
    return;
  }

  // map flexible headers
  const NAME = findHeader(namesRows, ["NAME"]);
  const SKU  = findHeader(namesRows, ["SKU"]);
  const SIZE = findHeader(namesRows, ["SIZE"]);
  const HS   = findHeader(namesRows, ["HS Code","HSCode"]);
  const EAN  = findHeader(namesRows, ["EAN","UPC"]);
  const COO1 = findHeader(namesRows, ["COO"]);
  const RRP1 = findHeader(namesRows, ["RRP","Sell Price","Retail"]);
  const DESC1= findHeader(namesRows, ["DESCRIPTION","DETAILS","NAME"]);
  
  const BRAND2 = findHeader(trackerRows, ["Brand","Label","Collection","BRAND"]);
  const STYLE2 = findHeader(trackerRows, ["STYLE CODE","STYLE"]);
  const DESC2  = findHeader(trackerRows, ["Line Description","Description","Name"]);
  const COO2   = findHeader(trackerRows, ["COO","Country of Origin"]);
  const PRICE2 = findHeader(trackerRows, ["Sell Price","RRP","Retail","Retail Price","Price"]);
  const COLOR2 = findHeader(trackerRows, ["Colour","Color","COLOUR"]);
  
  if (BRAND2) ffillColumn(trackerRows, BRAND2);
  

  if(!NAME || !SKU){ byId('log').innerHTML=`<span class="err">Missing NAME or SKU in Names CSV.</span>`; return; }
  if(!STYLE2){ byId('log').innerHTML=`<span class="err">Missing STYLE CODE in Tracker CSV.</span>`; return; }

  // index tracker by STYLE CODE (raw + base-without-variant)
  const tIndex = new Map(), rightKeys = new Set();
  for(const r of trackerRows){
    const raw  = cleanKey(r[STYLE2]);
    const base = cleanKey(stripVariantSuffix(r[STYLE2]));
    if(raw){ tIndex.set(raw, r); rightKeys.add(raw); }
    if(base){ tIndex.set(base, r); rightKeys.add(base); }
  }

  // diagnostics
  const leftKeys=new Set();
  for(const r of namesRows){ const k=cleanKey(r[NAME]); if(k) leftKeys.add(k); }
  const leftOnly=[...leftKeys].filter(k=>!rightKeys.has(k));
  const rightOnly=[...rightKeys].filter(k=>!leftKeys.has(k));
  byId('diag').innerHTML =
    `Left keys: ${leftKeys.size} · Right keys: ${rightKeys.size}<br>`+
    `Left-only: ${leftOnly.length} ${leftOnly.length? 'e.g. '+leftOnly.slice(0,5).join(', ') : ''}<br>`+
    `Right-only: ${rightOnly.length} ${rightOnly.length? 'e.g. '+rightOnly.slice(0,5).join(', ') : ''}`;

  // build output (strict inner join; skip rows lacking SKU or bad SKU)
  const outChildren = [];
  const outParents = [];
  const outStandalone = [];
  let kept=0, skipped=0, missing=0;

  for(const r of namesRows){
    const sku  = clean(r[SKU]);
    const name = clean(r[NAME]);
    if(!sku || !name || /\s/.test(sku)){ skipped++; continue; } // guard: skip obviously bad SKU (with spaces)

    // exact match, then base match
    let t = tIndex.get(cleanKey(name));
    if(!t){
      const baseName = stripVariantSuffix(name);
      t = tIndex.get(cleanKey(baseName));
    }
    if(!t){ missing++; continue; }

    const size   = clean(r[SIZE]);
    const ean    = clean(r[EAN]);
    const hs     = clean(r[HS]) || clean(t["HS Code"] || t["HSCode"]);
    const coo    = formatCOO(t[COO2] || r[COO1]);
    const price  = clean(t[PRICE2] || r[RRP1]);
    const baseName = stripVariantSuffix(name);
    const colour = COLOR2 ? clean(t[COLOR2]) : "";
    const desc     = clean(t[DESC2] || r[DESC1] || baseName);


    const isOS = isOneSize(size);
    const partDesc = (!isOS && size) ? `${desc}` : desc;
    const parentMatrix = (!isOS && size) ? parentFromSku(sku) : "";
    // Cleaned brand cell from tracker, then fallbacks
    const brandCell = BRAND2 ? clean(t[BRAND2]) : "";
    let brandVal =
      (/^BK$/i.test(brandCell) ? "BK" :
       /^SS$/i.test(brandCell) ? "SS" :
       /^BK/i.test(sku)        ? "BK" :
       /^SS/i.test(sku)        ? "SS" :
       /brother|kin/i.test(name) ? "BK" :
       /sisters|seekers/i.test(name) ? "SS" :
       "SS");

    const category = categoryFrom(desc, brandVal);

    const displayName = upcaseSizeTokens(formatName(desc));
    const partDescription = upcaseSizeTokens(formatName(partDesc));

    const baseRow = {
      "externalid": sku,
      "PartNumber": sku,
      "PartDescription": partDescription,
      "DisplayName": displayName,
      "Merchandise Hierarchy": category,
      "HSCode": hs,
      "CountryOfManufacture": coo,
      "UPC": ean,
      "unit type": "Each",
      "stock unit - pair/pack/": "",
      "purchase unit pair/pack": "",
      "Sales unit - pair/pack/": "",
      "subsid": "Sisters & Seekers Limited",
      "Class": "",
      "Department": "Product",
      "Location": "SSUK Warehouse : SSUK - Ecomm", 
      "children true/false": "",
      "salesDescription": "",
      "Sale Price": price,
      "Item - Pricing 1 : Currency (Req)": "GBP",
      "Item - Pricing 1 : Price Level (Req)": "RRP",
      "Item - Pricing 1 : Quantity (Req)" : "",
      "costing method": "FIFO",
      "Cost": "",
      "Drop Ship": "",
      "spec order true/false": "",
      "OrderUp toLevel": "",
      "re-order multi": "",
      "reordeer": "",
      "Vendor": "",
      "Prefered Vendor": "",
      "Purchase price": "",
      "cogsaccount": "50000 Cost of Goods Sold",
      "incomeaccount": "40001 Sales : Shopify Sales",
      "assetaccount": "12030 Inventory : Finished Goods",
      "purchasetaxcode": "",
      "Matrix Type": size ? "Child Matrix Item" : "Standalone Item",
      "Parent Matrix item": parentMatrix,
      "Matrix - Colour": colour,   // <-- colour ONLY from tracker
      "Matrix Size": size
    };
    
    if (!isOS && size) {
      // true child matrix item
      outChildren.push({ ...baseRow, "Matrix Type": "Child Matrix Item" });
    } else {
      // standalone (no size or ONE SIZE/OS)
      outStandalone.push({ ...baseRow, "Matrix Type": "Standalone Item" });
    }
    kept++;
  }
    // add Matrix Parent rows if requested
  // --- Always add Matrix Parent rows when children exist ---
    const groups = new Map();
    for(const r of outChildren){
      const parent = r["Parent Matrix item"];
      if(!parent) continue;
      if(!groups.has(parent)) groups.set(parent, []);
      groups.get(parent).push(r);
    }
  
    for(const [parentCode, kids] of groups.entries()){
      const k0 = kids[0];
      outParents.push({
        "externalid": parentCode,
        "PartNumber": parentCode,
        "PartDescription": k0["DisplayName"],
        "DisplayName": k0["DisplayName"],
        "Merchandise Hierarchy": k0["Merchandise Hierarchy"],
        "HSCode": k0["HSCode"],
        "CountryOfManufacture": k0["CountryOfManufacture"],
        "UPC": "",
        "unit type": "Each",
        "stock unit - pair/pack/": "",
        "purchase unit pair/pack": "",
        "Sales unit - pair/pack/": "",
        "subsid": "Sisters & Seekers Limited",
        "Class": "",
        "Department": "Product",
        "Location": "SSUK Warehouse : SSUK - Ecomm",
        "children true/false": "",
        "salesDescription": "",
        "Sale Price": k0["Sale Price"],
        "Item - Pricing 1 : Currency (Req)": "GBP",
        "Item - Pricing 1 : Price Level (Req)": "RRP",
        "Item - Pricing 1 : Quantity (Req)" : "",
        "costing method": "FIFO",
        "Cost": "",
        "Drop Ship": "",
        "spec order true/false": "",
        "OrderUp toLevel": "",
        "re-order multi": "",
        "reordeer": "",
        "Vendor": "",
        "Prefered Vendor": "",
        "Purchase price": "",
        "cogsaccount": "50000 Cost of Goods Sold",
        "incomeaccount": "40001 Sales : Shopify Sales",
        "assetaccount": "12030 Inventory : Finished Goods",
        "purchasetaxcode": "",
        "Matrix Type": "Parent Matrix Item",
        "Parent Matrix item": "",
        "Matrix - Colour": "",
        "Matrix Size": ""
      });
    }
 
  // --- Remove unnecessary columns from Standalone sheet ---
  const columnsToRemove = ["Matrix Type", "Parent Matrix item", "Matrix - Colour", "Matrix Size"];
  outStandalone.forEach(row => {
    columnsToRemove.forEach(col => delete row[col]);
  });

  
  // show preview counts
  byId('log').innerHTML =
    `✅ ${outChildren.length} children · ${outParents.length} parents · ${outStandalone.length} standalone (kept ${kept}, skipped ${skipped}, no match ${missing}).`;

  if (typeof XLSX === "undefined") {
    byId('log').innerHTML += `<br><span class="err">XLSX not found. Add the SheetJS script tag to your HTML.</span>`;
    return;
  }

  const wb  = XLSX.utils.book_new();
  const wsChildren   = XLSX.utils.json_to_sheet(outChildren);
  const wsParents    = XLSX.utils.json_to_sheet(outParents);
  const wsStandalone = XLSX.utils.json_to_sheet(outStandalone);

  XLSX.utils.book_append_sheet(wb, wsChildren,   "Children");
  XLSX.utils.book_append_sheet(wb, wsParents,    "Parents");
  XLSX.utils.book_append_sheet(wb, wsStandalone, "None Matrix Items");

  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const xlsBlob = new Blob([wbout], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const url = URL.createObjectURL(xlsBlob);

  const dl = byId('dlLink');
  dl.href = url;
  dl.download = "items_export.xlsx";
  dl.textContent = "Download Excel (Children + Parents + None Matrix Items)";
  dl.style.display = 'inline-block';

  // keep your preview (e.g., show Children or Standalone—your call)

  // Optional: quick preview table (children only)
  const preview = outChildren.slice(0,20);
  const tbl = byId('previewTable');
  tbl.innerHTML = "";
  if (preview.length){
    const headers = Object.keys(preview[0]);
    const thead = document.createElement('thead');
    const trh = document.createElement('tr');
    headers.forEach(h=>{ const th=document.createElement('th'); th.textContent=h; trh.appendChild(th); });
    thead.appendChild(trh);
    tbl.appendChild(thead);
    const tbody = document.createElement('tbody');
    preview.forEach(r=>{
      const tr=document.createElement('tr');
      headers.forEach(h=>{ const td=document.createElement('td'); td.textContent=r[h]||""; tr.appendChild(td); });
      tbody.appendChild(tr);
    });
    tbl.appendChild(tbody);
    byId('previewCard').style.display='block';
  }
}); // <-- this closes the click handler properly

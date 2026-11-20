
(() => {
  "use strict";
  const loadScript = (src) => new Promise((resolve, reject) => {
    const existing = [...document.scripts].find(s => s.src === src);
    if (existing) { if (existing.dataset.loaded === "yes") resolve(); else { existing.addEventListener("load", resolve, {once:true}); existing.addEventListener("error", reject, {once:true}); } return; }
    const s = document.createElement("script"); s.src = src; s.async = true;
    s.onload = () => { s.dataset.loaded = "yes"; resolve(); };
    s.onerror = () => reject(new Error("Could not load " + src));
    document.head.appendChild(s);
  });
  const deps = {
    xlsx: "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js",
    mammoth: "https://cdn.jsdelivr.net/npm/mammoth@1.9.0/mammoth.browser.min.js",
    jszip: "https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js",
    exceljs: "https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js"
  };
  const ensure = async (names) => {
    for (const name of names) {
      const present = name === "xlsx" ? window.XLSX :
        name === "mammoth" ? window.mammoth :
        name === "jszip" ? window.JSZip :
        name === "exceljs" ? window.ExcelJS : null;
      if (!deps[name]) throw new Error("Unknown Office library: " + name);
      if (!present) await loadScript(deps[name]);
    }
  };
  const button = (id, label, handler, primary) => {
    const b = document.createElement("button"); b.id = id; b.textContent = label;
    if (primary) b.className = "primary";
    b.addEventListener("click", handler); return b;
  };
  const fileInput = (id, accept, handler) => {
    const input = document.createElement("input"); input.type = "file"; input.id = id;
    input.accept = accept; input.className = "hidden";
    input.addEventListener("change", handler); return input;
  };
  const safeName = (s, ext) => (String(s || "opencreative").replace(/[^\w-]+/g, "-").replace(/^-+|-+$/g, "").slice(0, 60) || "opencreative") + ext;
  const MAX_DOCX_BYTES = 15 * 1024 * 1024;
  const MAX_XLSX_BYTES = 25 * 1024 * 1024;
  const MAX_XLSX_CELLS = 250000;
  const MAX_XLSX_ROWS = 5000;
  const MAX_XLSX_COLS = 256;
  const MAX_XLSX_SHEETS = 50;
  const checkFileSize = (file, limit, label) => {
    if (file && Number(file.size) > limit) throw new Error(label + " exceeds the " + Math.round(limit / 1024 / 1024) + " MB import limit.");
  };
  const xlsxColor = value => {
    const raw = typeof value === "string" ? value : value && (value.rgb || value.argb);
    if (!raw) return "";
    let hex = String(raw).replace(/^#/, "").toUpperCase();
    if (/^[0-9A-F]{8}$/.test(hex)) hex = hex.slice(-6);
    return /^[0-9A-F]{6}$/.test(hex) ? "#" + hex : "";
  };
  const importedStyle = cell => {
    const source = cell && cell.s || {}, style = {};
    if (source.font && source.font.bold) style.bold = true;
    if (source.font && source.font.italic) style.italic = true;
    const color = xlsxColor(source.font && source.font.color);
    const fill = xlsxColor(source.fill && source.fill.fgColor) || xlsxColor(source.fill && source.fill.bgColor);
    const align = source.alignment && source.alignment.horizontal;
    if (color) style.color = color;
    if (fill) style.background = fill;
    if (["left", "center", "right"].includes(align)) style.align = align;
    const numberFormat = window.OpenCreativeSheets && window.OpenCreativeSheets.normalizeNumberFormat(cell.z || "");
    if (numberFormat && numberFormat !== "general") style.numberFormat = numberFormat;
    return style;
  };
  const safeSheetName = (value, index, used) => {
    const base = String(value || ("Sheet" + (index + 1))).replace(/[\\/:?*\[\]]/g, " ").replace(/^'+|'+$/g, "").trim().slice(0, 31) || ("Sheet" + (index + 1));
    let name = base, suffix = 2;
    while (used.has(name.toLowerCase())) {
      const tail = " (" + suffix++ + ")";
      name = base.slice(0, 31 - tail.length) + tail;
    }
    used.add(name.toLowerCase());
    return name;
  };
  const downloadBlob = (blob, name) => {
    const url = URL.createObjectURL(blob), a = document.createElement("a");
    a.href = url; a.download = name; a.click(); setTimeout(() => URL.revokeObjectURL(url), 1500);
  };
  const status = (text) => {
    let el = document.getElementById("interopStatus");
    if (!el) { el = document.createElement("span"); el.id = "interopStatus"; el.className = "muted"; document.querySelector("#sheets .toolbar").append(el); }
    el.textContent = text;
  };
  const docStatus = (text) => { const el = document.getElementById("docStatus"); if (el) el.textContent = text; };
  const xmlEscape = (s) => String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&apos;");
  const wordInline = (node, props = {}) => {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = node.nodeValue || "";
      if (!text) return "";
      const p = [];
      if (props.bold) p.push("<w:b/>");
      if (props.italic) p.push("<w:i/>");
      if (props.underline) p.push('<w:u w:val="single"/>');
      return '<w:r>' + (p.length ? "<w:rPr>" + p.join("") + "</w:rPr>" : "") + '<w:t xml:space="preserve">' + xmlEscape(text) + "</w:t></w:r>";
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return "";
    const tag = node.tagName.toLowerCase();
    if (tag === "br") return "<w:r><w:br/></w:r>";
    if (tag === "img" || tag === "script" || tag === "style") return "";
    const next = {...props};
    if (tag === "b" || tag === "strong") next.bold = true;
    if (tag === "i" || tag === "em") next.italic = true;
    if (tag === "u") next.underline = true;
    return [...node.childNodes].map(child => wordInline(child, next)).join("");
  };
  const cellNodeBlocks = (node, isHeader) => {
    if (node.nodeType === Node.TEXT_NODE) {
      return (node.nodeValue || "").trim() ? "<w:p>" + wordInline(node, isHeader ? {bold:true} : {}) + "</w:p>" : "";
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return "";
    const tag = node.tagName.toLowerCase();
    if (tag === "table") return wordTableXml(node);
    if (tag === "p") {
      const runs = [...node.childNodes].map(child => wordInline(child, isHeader ? {bold:true} : {})).join("");
      return "<w:p>" + runs + "</w:p>";
    }
    if (/^h[1-6]$/.test(tag) || tag === "blockquote" || tag === "ul" || tag === "ol" || tag === "li") return blockXml(node);
    const hasBlocks = [...node.childNodes].some(child => child.nodeType === Node.ELEMENT_NODE && ["P","DIV","UL","OL","TABLE","BLOCKQUOTE"].includes(child.tagName));
    if (hasBlocks) return [...node.childNodes].map(child => cellNodeBlocks(child, isHeader)).join("");
    return "<w:p>" + wordInline(node, isHeader ? {bold:true} : {}) + "</w:p>";
  };
  const tableCellXml = cell => {
    const isHeader = cell.tagName.toLowerCase() === "th";
    const content = [...cell.childNodes].map(child => cellNodeBlocks(child, isHeader)).join("") || "<w:p/>";
    const properties = isHeader ? '<w:tcPr><w:shd w:val="clear" w:fill="E9EDF5"/></w:tcPr>' : "";
    return "<w:tc>" + properties + content + "</w:tc>";
  };
  const wordTableXml = table => {
    const rows = [...table.children].flatMap(section => {
      const tag = section.tagName.toLowerCase();
      if (tag === "tr") return [section];
      if (["thead","tbody","tfoot"].includes(tag)) return [...section.children].filter(row => row.tagName.toLowerCase() === "tr");
      return [];
    });
    if (!rows.length) return "<w:p/>";
    const columnCount = Math.max(1, ...rows.map(row => [...row.children].filter(cell => ["td","th"].includes(cell.tagName.toLowerCase())).length));
    const columnWidth = Math.max(1, Math.floor(9360 / columnCount));
    const grid = Array.from({length:columnCount}, () => '<w:gridCol w:w="' + columnWidth + '"/>').join("");
    const borders = '<w:tblBorders><w:top w:val="single" w:sz="4" w:color="B7C0CE"/><w:left w:val="single" w:sz="4" w:color="B7C0CE"/><w:bottom w:val="single" w:sz="4" w:color="B7C0CE"/><w:right w:val="single" w:sz="4" w:color="B7C0CE"/><w:insideH w:val="single" w:sz="4" w:color="B7C0CE"/><w:insideV w:val="single" w:sz="4" w:color="B7C0CE"/></w:tblBorders>';
    const body = rows.map(row => "<w:tr>" + [...row.children].filter(cell => ["td","th"].includes(cell.tagName.toLowerCase())).map(tableCellXml).join("") + "</w:tr>").join("");
    return '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>' + borders + '</w:tblPr><w:tblGrid>' + grid + '</w:tblGrid>' + body + '</w:tbl>';
  };
  const blockXml = (node) => {
    const tag = node.tagName.toLowerCase();
    if (tag === "table") return wordTableXml(node);
    if (tag === "ul" || tag === "ol") return [...node.children].map((li, index) => {
      const marker = tag === "ol" ? (index + 1) + ". " : "• ";
      return '<w:p><w:pPr><w:ind w:left="360"/></w:pPr><w:r><w:t xml:space="preserve">' + xmlEscape(marker) + "</w:t></w:r>" + [...li.childNodes].map(n => wordInline(n)).join("") + "</w:p>";
    }).join("");
    if (tag === "li") return '<w:p><w:r><w:t>• </w:t></w:r>' + [...node.childNodes].map(n => wordInline(n)).join("") + "</w:p>";
    const runs = [...node.childNodes].map(n => wordInline(n)).join("");
    let pPr = "";
    if (/^h[1-6]$/.test(tag)) pPr = '<w:pPr><w:pStyle w:val="Heading' + tag.slice(1) + '"/></w:pPr>';
    if (tag === "blockquote") pPr = '<w:pPr><w:ind w:left="480"/></w:pPr>';
    return "<w:p>" + pPr + runs + "</w:p>";
  };
  const buildDocxBlob = async html => {
    await ensure(["jszip"]);
    const parsed = new DOMParser().parseFromString(String(html || ""), "text/html");
    const body = [...parsed.body.childNodes].map(n => {
      if (n.nodeType === Node.TEXT_NODE) return "<w:p><w:r><w:t>" + xmlEscape(n.nodeValue) + "</w:t></w:r></w:p>";
      if (n.nodeType === Node.ELEMENT_NODE) return blockXml(n);
      return "";
    }).join("") || "<w:p/>";
    const documentXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + body +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>';
    const headingSizes = [32, 28, 24, 22, 20, 18];
    const headingStyles = headingSizes.map((size, index) => '<w:style w:type="paragraph" w:styleId="Heading' + (index + 1) + '"><w:name w:val="heading ' + (index + 1) + '"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/></w:pPr><w:rPr><w:b/><w:sz w:val="' + size + '"/></w:rPr></w:style>').join("");
    const stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:sz w:val="22"/></w:rPr></w:style>' + headingStyles + "</w:styles>";
    const zip = new JSZip();
    zip.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>');
    zip.file("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
    zip.file("word/_rels/document.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
    zip.file("word/document.xml", documentXml);
    zip.file("word/styles.xml", stylesXml);
    return zip.generateAsync({type:"blob", mimeType:"application/vnd.openxmlformats-officedocument.wordprocessingml.document", compression:"DEFLATE", compressionOptions:{level:6}});
  };
  const exportDocx = async () => {
    try {
      const title = document.getElementById("docTitle").value || "Document";
      const blob = await buildDocxBlob(document.getElementById("docEditor").innerHTML);
      downloadBlob(blob, safeName(title, ".docx")); docStatus("DOCX exported");
    } catch (e) { docStatus("DOCX export unavailable"); alert("DOCX export failed: " + e.message); }
  };
  const cleanImportedHtml = (html) => {
    const parsed = new DOMParser().parseFromString(html, "text/html");
    const allowed = new Set(["P","H1","H2","H3","H4","H5","H6","BR","STRONG","B","EM","I","U","UL","OL","LI","BLOCKQUOTE","TABLE","THEAD","TBODY","TFOOT","TR","TD","TH"]);
    const cleanNode = (node) => {
      if (node.nodeType === Node.TEXT_NODE) return document.createTextNode(node.nodeValue || "");
      if (node.nodeType !== Node.ELEMENT_NODE) return null;
      if (!allowed.has(node.tagName)) {
        const fragment = document.createDocumentFragment();
        [...node.childNodes].forEach(child => { const c = cleanNode(child); if (c) fragment.append(c); });
        return fragment;
      }
      const el = document.createElement(node.tagName.toLowerCase());
      [...node.childNodes].forEach(child => { const c = cleanNode(child); if (c) el.append(c); });
      return el;
    };
    const out = document.createElement("div");
    [...parsed.body.childNodes].forEach(n => { const c = cleanNode(n); if (c) out.append(c); });
    return out.innerHTML;
  };
  const convertDocxBuffer = async arrayBuffer => {
    await ensure(["mammoth"]);
    const result = await mammoth.convertToHtml({arrayBuffer});
    return {html:cleanImportedHtml(result.value),messages:result.messages || []};
  };
  const importDocx = async (file) => {
    try {
      checkFileSize(file, MAX_DOCX_BYTES, "DOCX file");
      const result = await convertDocxBuffer(await file.arrayBuffer());
      const clean = result.html;
      if (!clean.trim()) throw new Error("No readable document text was found.");
      document.getElementById("docEditor").innerHTML = clean;
      document.getElementById("docTitle").value = file.name.replace(/\.docx$/i, "") || "Imported document";
      document.getElementById("docEditor").dispatchEvent(new Event("input", {bubbles:true}));
      docStatus("DOCX imported");
      if (result.messages && result.messages.length) alert("Imported with notes: some Word formatting or unsupported content may not be preserved.");
    } catch (e) { alert("DOCX import failed: " + e.message + ". Check your internet connection and try again."); }
  };
  const styleArgb = value => {
    const hex = String(value || "").replace(/^#/, "").toUpperCase();
    return /^[0-9A-F]{6}$/.test(hex) ? "FF" + hex : "";
  };
  const hasExportStyle = style => !!(style && (
    style.bold === true || style.italic === true ||
    styleArgb(style.color) || styleArgb(style.background) ||
    ["left", "center", "right"].includes(style.align) ||
    ["number", "integer", "percent", "currency"].includes(style.numberFormat)
  ));
  const applyExcelJsStyle = (cell, style) => {
    if (!style || typeof style !== "object") return;
    const font = {};
    if (style.bold === true) font.bold = true;
    if (style.italic === true) font.italic = true;
    const color = styleArgb(style.color);
    if (color) font.color = {argb:color};
    if (Object.keys(font).length) cell.font = font;
    const fill = styleArgb(style.background);
    if (fill) cell.fill = {type:"pattern", pattern:"solid", fgColor:{argb:fill}};
    if (["left", "center", "right"].includes(style.align)) {
      cell.alignment = {horizontal:style.align};
    }
    const numberFormat = window.OpenCreativeSheets && window.OpenCreativeSheets.numberFormats[style.numberFormat];
    if (numberFormat) cell.numFmt = numberFormat;
  };
  const cellAddress = (r, c) => {
    let n = c + 1, label = "";
    while (n > 0) {
      const remainder = (n - 1) % 26;
      label = String.fromCharCode(65 + remainder) + label;
      n = Math.floor((n - 1) / 26);
    }
    return label + (r + 1);
  };
  const verifyExcelJsRoundTrip = async (bytes, sheets, names) => {
    const reopened = new ExcelJS.Workbook();
    await reopened.xlsx.load(bytes);
    if (reopened.worksheets.length !== sheets.length) {
      throw new Error("The generated workbook did not reopen with the expected worksheet count.");
    }
    let checked = 0, styled = 0;
    sheets.forEach((sheet, index) => {
      const ws = reopened.getWorksheet(names[index]);
      if (!ws) throw new Error("Worksheet verification failed for " + names[index] + ".");
      sheet.cells.forEach((row, r) => row.forEach((value, c) => {
        const expected = String(value ?? ""), cell = ws.getCell(r + 1, c + 1);
        const style = sheet.styles && sheet.styles[r + ":" + c] || {};
        if (expected.startsWith("=")) {
          const formula = cell.value && typeof cell.value === "object" ? cell.value.formula : "";
          if (String(formula || "") !== expected.slice(1)) {
            throw new Error("Formula verification failed at " + names[index] + "!" + cellAddress(r,c) + ".");
          }
        } else if (expected !== "" && String(cell.value ?? "") !== expected) {
          throw new Error("Value verification failed at " + names[index] + "!" + cellAddress(r,c) + ".");
        }
        if (style.bold === true && cell.font.bold !== true) {
          throw new Error("Bold formatting verification failed at " + names[index] + "!" + cellAddress(r,c) + ".");
        }
        if (style.italic === true && cell.font.italic !== true) {
          throw new Error("Italic formatting verification failed at " + names[index] + "!" + cellAddress(r,c) + ".");
        }
        const expectedColor = styleArgb(style.color);
        if (expectedColor && String(cell.font.color && cell.font.color.argb || "").toUpperCase() !== expectedColor) {
          throw new Error("Font color verification failed at " + names[index] + "!" + cellAddress(r,c) + ".");
        }
        const expectedFill = styleArgb(style.background);
        if (expectedFill && String(cell.fill && cell.fill.fgColor && cell.fill.fgColor.argb || "").toUpperCase() !== expectedFill) {
          throw new Error("Fill color verification failed at " + names[index] + "!" + cellAddress(r,c) + ".");
        }
        if (["left", "center", "right"].includes(style.align) &&
            (!cell.alignment || cell.alignment.horizontal !== style.align)) {
          throw new Error("Alignment verification failed at " + names[index] + "!" + cellAddress(r,c) + ".");
        }
        const expectedNumberFormat = window.OpenCreativeSheets && window.OpenCreativeSheets.numberFormats[style.numberFormat];
        if (expectedNumberFormat && cell.numFmt !== expectedNumberFormat) {
          throw new Error("Number format verification failed at " + names[index] + "!" + cellAddress(r,c) + ".");
        }
        if (expected !== "") checked++;
        if (hasExportStyle(style)) styled++;
      }));
    });
    return {checked, styled};
  };
  const buildXlsxBuffer = async sheets => {
    if (!sheets.length || sheets.length > MAX_XLSX_SHEETS) throw new Error("A workbook must contain 1 to " + MAX_XLSX_SHEETS + " worksheets.");
    let totalCells = 0;
    sheets.forEach(sheet => {
      if (!sheet || !Array.isArray(sheet.cells) || sheet.cells.some(row => !Array.isArray(row))) throw new Error("A worksheet contains invalid cell data.");
      const rowCount = sheet.cells.length, colCount = Math.max(1, ...sheet.cells.map(row => row.length));
      if (rowCount > MAX_XLSX_ROWS || colCount > MAX_XLSX_COLS) throw new Error("Workbook exceeds the supported export size of " + MAX_XLSX_ROWS + " rows by " + MAX_XLSX_COLS + " columns per sheet.");
      totalCells += rowCount * colCount;
    });
    if (totalCells > MAX_XLSX_CELLS) throw new Error("Workbook exceeds the " + MAX_XLSX_CELLS.toLocaleString() + " cell export limit.");
    const workbook = new ExcelJS.Workbook(), usedNames = new Set(), names = [];
    workbook.creator = "OpenCreative";
    workbook.calcProperties = {fullCalcOnLoad:true, forceFullCalc:true};
    sheets.forEach((sheet, index) => {
      const name = safeSheetName(sheet.name, index, usedNames);
      names.push(name);
      const worksheet = workbook.addWorksheet(name);
      sheet.cells.forEach((row, r) => row.forEach((value, c) => {
        const raw = String(value ?? ""), style = sheet.styles && sheet.styles[r + ":" + c] || {};
        if (raw === "" && !hasExportStyle(style)) return;
        const cell = worksheet.getCell(r + 1, c + 1);
        if (raw.startsWith("=")) cell.value = {formula:raw.slice(1)};
        else if (raw !== "") {
          const numericFormat = ["number", "integer", "percent", "currency"].includes(style.numberFormat);
          const numericValue = Number(raw);
          cell.value = numericFormat && raw.trim() !== "" && Number.isFinite(numericValue) ? numericValue : raw;
        }
        applyExcelJsStyle(cell, style);
      }));
    });
    const bytes = await workbook.xlsx.writeBuffer();
    const verification = await verifyExcelJsRoundTrip(bytes, sheets, names);
    return {bytes,names,verification};
  };
  const exportXlsx = async () => {
    try {
      await ensure(["exceljs"]);
      const sheets = window.OpenCreativeSheets ? window.OpenCreativeSheets.getWorkbook() : [{name:"Sheet1",cells:data.sheet,styles:{}}];
      const built = await buildXlsxBuffer(sheets);
      downloadBlob(new Blob([built.bytes], {type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}), "opencreative-workbook.xlsx");
      const renamed = sheets.some((sheet, index) => String(sheet.name || ("Sheet" + (index + 1))) !== built.names[index]);
      status("XLSX verified: " + sheets.length + " sheet(s), " + built.verification.checked + " value/formula cell(s), " + built.verification.styled + " styled cell(s)." +
        (renamed ? " Invalid or duplicate sheet names were made unique in the export." : "") +
        " Formulas are set to recalculate when opened in a compatible spreadsheet app.");
    } catch (e) {
      status("XLSX export unavailable");
      alert("XLSX export failed: " + e.message + (e.message.toLowerCase().includes("exceljs") ? " Check your internet connection and try again." : ""));
    }
  };
  const importXlsx = async (file, apply = true) => {
    try {
      checkFileSize(file, MAX_XLSX_BYTES, "XLSX file");
      await ensure(["xlsx"]);
      const wb = XLSX.read(await file.arrayBuffer(), {type:"array", cellFormula:true, cellStyles:true});
      if (!wb.SheetNames.length) throw new Error("The workbook has no worksheets.");
      if (wb.SheetNames.length > MAX_XLSX_SHEETS) throw new Error("Workbook contains more than " + MAX_XLSX_SHEETS + " worksheets.");
      let totalCells = 0;
      const imported = wb.SheetNames.map(name => {
        const ws = wb.Sheets[name];
        const ref = ws["!ref"] || "A1", range = XLSX.utils.decode_range(ref);
        const rowCount = range.e.r + 1, colCount = range.e.c + 1;
        if (rowCount > MAX_XLSX_ROWS || colCount > MAX_XLSX_COLS) throw new Error("Worksheet '" + name + "' exceeds the supported import size of " + MAX_XLSX_ROWS + " rows by " + MAX_XLSX_COLS + " columns.");
        totalCells += rowCount * colCount;
        if (totalCells > MAX_XLSX_CELLS) throw new Error("Workbook exceeds the " + MAX_XLSX_CELLS.toLocaleString() + " cell import limit.");
        const cells = Array.from({length:rowCount}, () => Array(colCount).fill(""));
        const styles = {};
        for (let r=range.s.r; r<=range.e.r; r++) {
          for (let c=range.s.c; c<=range.e.c; c++) {
            const cell = ws[cellAddress(r,c)];
            if (!cell) continue;
            cells[r][c] = cell.f ? "=" + cell.f : String(cell.v ?? cell.w ?? "");
            const style = importedStyle(cell);
            if (Object.keys(style).length) styles[r + ":" + c] = style;
          }
        }
        return {name, cells, styles};
      });
      if (apply) {
        if (window.OpenCreativeSheets) window.OpenCreativeSheets.setWorkbook(imported);
        else {
          data.sheet = imported[0].cells; rows = data.sheet.length; cols = Math.max(1,...imported[0].cells.map(r=>r.length)); persist(); renderSheet();
        }
        const styled = imported.reduce((count, sheet) => count + Object.keys(sheet.styles).length, 0);
        status("Imported " + imported.length + " worksheet(s), " + styled + " styled cell(s) where supported, from " + file.name);
      }
      return imported;
    } catch (e) {
      if (apply) alert("XLSX import failed: " + e.message + ". The current workbook was not changed.");
      else throw e;
    }
  };
  const runOfficeRegressionTests = async () => {
    const checks = [];
    const assert = (condition, message) => { if (!condition) throw new Error(message); checks.push(message); };
    await ensure(["exceljs", "xlsx", "jszip", "mammoth"]);
    assert(window.OpenCreativeSheets.normalizeNumberFormat("#,##0.00") === "number" &&
      window.OpenCreativeSheets.normalizeNumberFormat("0.00%") === "percent" &&
      window.OpenCreativeSheets.normalizeNumberFormat("₹#,##0.00") === "currency",
      "Recognizes the supported imported and exported number format codes.");
    assert(window.OpenCreativeSheets.formatNumber(0.125, "percent") === "12.50%" &&
      window.OpenCreativeSheets.formatNumber(1250.5, "currency").includes("₹"),
      "Displays percentage and Indian rupee formats in the sheet grid.");
    const sample = [{
      name:"Office regression",
      cells:[["Metric","Value","Formula"],["Revenue","1250.5","=SUM(B2:B2)"],["Rate","0.125","=B3*100"]],
      styles:{"1:1":{bold:true,numberFormat:"currency"},"2:1":{numberFormat:"percent"},"1:0":{italic:true}}
    }];
    const built = await buildXlsxBuffer(sample);
    assert(built.verification.checked === 9, "XLSX export reopens with all nine populated values/formulas.");
    assert(built.verification.styled === 3, "XLSX export verifies bold, italic and number-format styles.");
    const parsed = await importXlsx(new File([built.bytes], "office-regression.xlsx", {type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}), false);
    assert(parsed.length === 1 && parsed[0].name === "Office regression", "XLSX import restores the worksheet name.");
    assert(parsed[0].cells[1][2] === "=SUM(B2:B2)" && parsed[0].cells[2][2] === "=B3*100", "XLSX import preserves formula expressions.");
    assert(parsed[0].styles["1:1"].numberFormat === "currency" && parsed[0].styles["2:1"].numberFormat === "percent", "XLSX import restores supported number formats.");
    assert(parsed[0].cells[1][1] === "1250.5" && parsed[0].cells[2][1] === "0.125", "XLSX import preserves raw numeric values.");
    const sampleHtml = '<h1>Regression heading</h1><p><strong>Bold text</strong> and <em>italic text</em>.</p><table><thead><tr><th>Item</th><th>Qty</th></tr></thead><tbody><tr><td>Apples</td><td>3</td></tr></tbody></table>';
    const docx = await buildDocxBlob(sampleHtml);
    const converted = await convertDocxBuffer(await docx.arrayBuffer());
    const doc = new DOMParser().parseFromString(converted.html, "text/html");
    assert(!!doc.querySelector("h1"), "DOCX export/import preserves a heading.");
    assert(!!doc.querySelector("strong") && !!doc.querySelector("em"), "DOCX export/import preserves basic emphasis.");
    assert(doc.querySelectorAll("table tr").length === 2 && doc.querySelectorAll("th").length === 2, "DOCX export/import preserves a simple table and header cells.");
    return {passed:checks.length,checks};
  };
  window.OpenCreativeInteropTests = {run:runOfficeRegressionTests};

  const init = () => {
    const docToolbar = document.querySelector("#docs .toolbar");
    const docInput = fileInput("docxImportInput", ".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document", e => {
      const f = e.target.files[0]; if (f) importDocx(f); e.target.value = "";
    });
    const insertTable = () => {
      const editor = document.getElementById("docEditor"), table = document.createElement("table");
      const head = document.createElement("thead"), headerRow = document.createElement("tr");
      for (let c = 1; c <= 3; c++) { const cell = document.createElement("th"); cell.textContent = "Column " + c; headerRow.append(cell); }
      head.append(headerRow); table.append(head);
      const body = document.createElement("tbody");
      for (let r = 0; r < 2; r++) {
        const row = document.createElement("tr");
        for (let c = 0; c < 3; c++) row.append(document.createElement("td"));
        body.append(row);
      }
      table.append(body); editor.append(table);
      editor.dispatchEvent(new Event("input", {bubbles:true}));
      const firstCell = table.querySelector("th"); if (firstCell) firstCell.focus();
      docStatus("3 × 3 table inserted. Edit the cells and export as DOCX.");
    };
    docToolbar.append(button("docxImport", "Import DOCX", () => docInput.click()), button("docxExport", "Export DOCX", exportDocx), button("docInsertTable", "Insert 3×3 table", insertTable), docInput);
    const sheetToolbar = document.querySelector("#sheets .toolbar");
    const xlsxInput = fileInput("xlsxImportInput", ".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", e => {
      const f = e.target.files[0]; if (f) importXlsx(f); e.target.value = "";
    });
    sheetToolbar.append(button("xlsxImport", "Import XLSX", () => xlsxInput.click()), button("xlsxExport", "Export XLSX", exportXlsx, true), xlsxInput);
    status("Office import/export runs in your browser; libraries load on demand.");
    docStatus("Saved locally");
  };
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init, {once:true}); else init();
})();

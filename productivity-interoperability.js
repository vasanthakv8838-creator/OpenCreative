
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
    xlsx: "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js",
    mammoth: "https://cdn.jsdelivr.net/npm/mammoth@1.9.0/mammoth.browser.min.js",
    jszip: "https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js"
  };
  const ensure = async (names) => {
    for (const name of names) {
      const present = name === "xlsx" ? window.XLSX : name === "mammoth" ? window.mammoth : window.JSZip;
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
  const blockXml = (node) => {
    const tag = node.tagName.toLowerCase();
    if (tag === "ul" || tag === "ol") return [...node.children].map(li => '<w:p><w:pPr><w:ind w:left="360"/></w:pPr><w:r><w:t>• </w:t></w:r>' + [...li.childNodes].map(n => wordInline(n)).join("") + "</w:p>").join("");
    if (tag === "li") return '<w:p><w:r><w:t>• </w:t></w:r>' + [...node.childNodes].map(n => wordInline(n)).join("") + "</w:p>";
    const runs = [...node.childNodes].map(wordRun).join("");
    let pPr = "";
    if (/^h[1-6]$/.test(tag)) pPr = '<w:pPr><w:pStyle w:val="Heading' + tag.slice(1) + '"/></w:pPr>';
    if (tag === "blockquote") pPr = '<w:pPr><w:ind w:left="480"/></w:pPr>';
    return "<w:p>" + pPr + runs + "</w:p>";
  };
  const exportDocx = async () => {
    try {
      await ensure(["jszip"]);
      const title = document.getElementById("docTitle").value || "Document";
      const editor = document.getElementById("docEditor");
      const parsed = new DOMParser().parseFromString(editor.innerHTML, "text/html");
      const body = [...parsed.body.childNodes].map(n => {
        if (n.nodeType === Node.TEXT_NODE) return "<w:p><w:r><w:t>" + xmlEscape(n.nodeValue) + "</w:t></w:r></w:p>";
        if (n.nodeType === Node.ELEMENT_NODE) return blockXml(n);
        return "";
      }).join("") || "<w:p/>";
      const documentXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + body +
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>';
      const zip = new JSZip();
      zip.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
      zip.file("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
      zip.file("word/document.xml", documentXml);
      const blob = await zip.generateAsync({type:"blob", mimeType:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"});
      downloadBlob(blob, safeName(title, ".docx")); docStatus("DOCX exported");
    } catch (e) { docStatus("DOCX export unavailable"); alert("DOCX export failed: " + e.message); }
  };
  const cleanImportedHtml = (html) => {
    const parsed = new DOMParser().parseFromString(html, "text/html");
    const allowed = new Set(["P","H1","H2","H3","H4","H5","H6","BR","STRONG","B","EM","I","U","UL","OL","LI","BLOCKQUOTE"]);
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
  const importDocx = async (file) => {
    try {
      await ensure(["mammoth"]);
      const result = await mammoth.convertToHtml({arrayBuffer: await file.arrayBuffer()});
      const clean = cleanImportedHtml(result.value);
      if (!clean.trim()) throw new Error("No readable document text was found.");
      document.getElementById("docEditor").innerHTML = clean;
      document.getElementById("docTitle").value = file.name.replace(/\.docx$/i, "") || "Imported document";
      document.getElementById("docEditor").dispatchEvent(new Event("input", {bubbles:true}));
      docStatus("DOCX imported");
      if (result.messages && result.messages.length) alert("Imported with notes: some Word formatting or unsupported content may not be preserved.");
    } catch (e) { alert("DOCX import failed: " + e.message + ". Check your internet connection and try again."); }
  };
  const toAoA = () => data.sheet.map((row, r) => row.map((value, c) => {
    const raw = String(value ?? "");
    if (raw.startsWith("=")) return {f:raw.slice(1)};
    return raw;
  }));
  const exportXlsx = async () => {
    try {
      await ensure(["xlsx"]);
      const ws = XLSX.utils.aoa_to_sheet(toAoA());
      const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
      const bytes = XLSX.write(wb, {bookType:"xlsx", type:"array", cellFormula:true});
      downloadBlob(new Blob([bytes], {type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}), "opencreative-sheet.xlsx");
      status("XLSX exported");
    } catch (e) { status("XLSX export unavailable"); alert("XLSX export failed: " + e.message); }
  };
  const importXlsx = async (file) => {
    try {
      await ensure(["xlsx"]);
      const wb = XLSX.read(await file.arrayBuffer(), {type:"array", cellFormula:true});
      const ws = wb.Sheets[wb.SheetNames[0]];
      if (!ws) throw new Error("The workbook has no visible worksheet.");
      const ref = ws["!ref"] || "A1";
      const range = XLSX.utils.decode_range(ref);
      const next = [];
      for (let r=range.s.r; r<=range.e.r; r++) {
        const row = [];
        for (let c=range.s.c; c<=range.e.c; c++) {
          const cell = ws[XLSX.utils.encode_cell({r,c})];
          row.push(!cell ? "" : cell.f ? "=" + cell.f : cell.v ?? cell.w ?? "");
        }
        next.push(row);
      }
      if (!next.length) throw new Error("The workbook is empty.");
      data.sheet = next; rows = next.length; cols = Math.max(1, ...next.map(r => r.length));
      data.sheet.forEach(r => { while (r.length < cols) r.push(""); });
      persist(); renderSheet(); status("Imported " + file.name);
    } catch (e) { alert("XLSX import failed: " + e.message + ". The current sheet was not changed."); }
  };
  const init = () => {
    const docToolbar = document.querySelector("#docs .toolbar");
    const docInput = fileInput("docxImportInput", ".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document", e => {
      const f = e.target.files[0]; if (f) importDocx(f); e.target.value = "";
    });
    docToolbar.append(button("docxImport", "Import DOCX", () => docInput.click()), button("docxExport", "Export DOCX", exportDocx), docInput);
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

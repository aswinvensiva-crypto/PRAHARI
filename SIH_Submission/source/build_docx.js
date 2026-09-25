// Build PRAHARI_SIH_Idea_Submission.docx from the Markdown sources in SIH_Submission/.
// Usage (from SIH_Submission/): node source/build_docx.js   — requires `npm install docx`.
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, LevelFormat, Footer, PageNumber, PageBreak,
} = require("docx");

const ROOT = path.resolve(__dirname, "..");
const NAVY = "0F2A44", TEAL = "0E7C7B", SAFFRON = "E8811F", MUTED = "5B6B7A";
const CONTENT_W = 9026; // A4 width 11906 minus two 1" margins

const SECTIONS = [
  { file: "01_Idea_Title.md", heading: "Field 1: Idea Title" },
  { file: "02_Idea_Description.md", heading: "Field 2: Idea Description" },
  { file: "03_Abstract.md", heading: "Field 3: Abstract / Summary" },
  { file: "04_Business_Model_Canvas.md", heading: "Field 4: Additional Document, Business Model Canvas" },
  { file: "05_YouTube_Video_Script_and_Storyboard.md", heading: "Field 5: YouTube Video Script & Storyboard" },
  { file: "README.md", heading: "Appendix: Evaluation Alignment & Submission Checklist" },
];

// ---------- inline markdown → TextRuns ----------
function runs(text, base = {}) {
  const out = [];
  const re = /(\*\*[^*]+\*\*|`[^`]+`|\*[^*\s][^*]*\*)/g;
  let last = 0, m;
  while ((m = re.exec(text))) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), ...base }));
    const tok = m[0];
    if (tok.startsWith("**")) out.push(new TextRun({ text: tok.slice(2, -2), ...base, bold: true }));
    else if (tok.startsWith("`")) out.push(new TextRun({ text: tok.slice(1, -1), ...base, font: "Consolas", size: 19 }));
    else out.push(new TextRun({ text: tok.slice(1, -1), ...base, italics: true }));
    last = m.index + tok.length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), ...base }));
  return out;
}

let numInstance = 0;
function table(rows) {
  const cells = rows.map((r) => r.replace(/^\s*\|/, "").replace(/\|\s*$/, "").split("|").map((c) => c.trim()));
  const header = cells[0];
  const body = cells.slice(2);
  const n = header.length;
  const weights = header.map((_, i) => Math.max(8, Math.min(60, ...[header, ...body].map((r) => (r[i] || "").length))));
  const sum = weights.reduce((a, b) => a + b, 0);
  const widths = weights.map((w) => Math.floor((w / sum) * CONTENT_W));
  widths[n - 1] += CONTENT_W - widths.reduce((a, b) => a + b, 0);
  const border = { style: BorderStyle.SINGLE, size: 4, color: "C9D3DD" };
  const borders = { top: border, bottom: border, left: border, right: border };
  const mk = (r, isHead) => new TableRow({
    tableHeader: isHead,
    children: widths.map((w, i) => new TableCell({
      width: { size: w, type: WidthType.DXA },
      borders,
      shading: isHead ? { fill: NAVY, type: ShadingType.CLEAR, color: "auto" } : undefined,
      margins: { top: 60, bottom: 60, left: 90, right: 90 },
      children: [new Paragraph({ children: runs(r[i] || "", isHead ? { bold: true, color: "FFFFFF", size: 18 } : { size: 18 }) })],
    })),
  });
  return new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: widths, rows: [mk(header, true), ...body.map((r) => mk(r, false))] });
}

function convert(md, skipH1) {
  const lines = md.split("\n");
  const out = [];
  let i = 0, inNumbered = false;
  while (i < lines.length) {
    const line = lines[i];
    if (/^```/.test(line)) {
      i++;
      while (i < lines.length && !/^```/.test(lines[i])) {
        out.push(new Paragraph({ shading: { fill: "F2F5F8", type: ShadingType.CLEAR, color: "auto" }, spacing: { after: 0 }, children: [new TextRun({ text: lines[i] || " ", font: "Consolas", size: 18 })] }));
        i++;
      }
      i++; continue;
    }
    if (/^\s*\|/.test(line)) {
      const rows = [];
      while (i < lines.length && /^\s*\|/.test(lines[i])) rows.push(lines[i++]);
      out.push(table(rows));
      out.push(new Paragraph({ spacing: { after: 80 }, children: [] }));
      continue;
    }
    const h = line.match(/^(#{1,6})\s+(.*)$/);
    if (h) {
      inNumbered = false;
      const lvl = h[1].length;
      if (lvl === 1 && skipH1) { i++; continue; }
      const map = { 1: HeadingLevel.HEADING_2, 2: HeadingLevel.HEADING_2, 3: HeadingLevel.HEADING_3 };
      out.push(new Paragraph({ heading: map[lvl] || HeadingLevel.HEADING_3, children: runs(h[2]) }));
      i++; continue;
    }
    if (/^---\s*$/.test(line)) {
      out.push(new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "C9D3DD", space: 1 } }, spacing: { after: 120 }, children: [] }));
      i++; continue;
    }
    const q = line.match(/^>\s?(.*)$/);
    if (q) {
      out.push(new Paragraph({ shading: { fill: "FDF0E2", type: ShadingType.CLEAR, color: "auto" }, border: { left: { style: BorderStyle.SINGLE, size: 18, color: SAFFRON, space: 6 } }, children: runs(q[1]) }));
      i++; continue;
    }
    const b = line.match(/^(\s*)[-*]\s+(.*)$/);
    if (b) {
      inNumbered = false;
      const level = Math.min(2, Math.floor(b[1].length / 2));
      let text = b[2];
      const cb = text.match(/^\[( |x)\]\s+(.*)$/i);
      if (cb) text = (cb[1].toLowerCase() === "x" ? "☑ " : "☐ ") + cb[2];
      out.push(new Paragraph({ numbering: { reference: "bullets", level }, spacing: { after: 60 }, children: runs(text) }));
      i++; continue;
    }
    const nm = line.match(/^(\s*)(\d+)\.\s+(.*)$/);
    if (nm) {
      if (!inNumbered) { numInstance++; inNumbered = true; }
      const level = Math.min(1, Math.floor(nm[1].length / 3));
      out.push(new Paragraph({ numbering: { reference: "numbers", level, instance: numInstance }, spacing: { after: 60 }, children: runs(nm[3]) }));
      i++; continue;
    }
    if (line.trim() === "") { i++; continue; }
    inNumbered = inNumbered && /^\s{2,}/.test(line);
    out.push(new Paragraph({ spacing: { after: 120 }, children: runs(line.trim()) }));
    i++;
  }
  return out;
}

// ---------- cover ----------
const cover = [
  new Paragraph({ spacing: { before: 2400, after: 120 }, children: [new TextRun({ text: "SMART INDIA HACKATHON · STUDENT INNOVATION · TRAVEL & TOURISM", color: TEAL, bold: true, size: 20 })] }),
  new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: "PRAHARI", bold: true, color: NAVY, size: 96, font: "Arial" })] }),
  new Paragraph({ spacing: { after: 360 }, children: [new TextRun({ text: "Predictive Routing & AI for Hotels, Attractions, Regions & Itineraries", color: MUTED, size: 22 })] }),
  new Paragraph({ spacing: { after: 240 }, children: [new TextRun({ text: "AI Demand Forecasting & Crowd Rebalancing to Boost Hotels, Homestays and Travel in India", bold: true, color: NAVY, size: 34 })] }),
  new Paragraph({ spacing: { after: 600 }, children: [new TextRun({ text: "PREDICT · PROTECT · REDISTRIBUTE", bold: true, color: SAFFRON, size: 26 })] }),
  new Paragraph({ border: { top: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 6 } }, spacing: { after: 80 }, children: [new TextRun({ text: "Complete Idea Submission Package", bold: true, size: 24, color: NAVY })] }),
  ...SECTIONS.map((s) => new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 40 }, children: [new TextRun({ text: s.heading, size: 21 })] })),
  new Paragraph({ spacing: { before: 240 }, children: [new TextRun({ text: "Problem Statement: Student Innovation. A solution or idea that can boost the current situation of the tourism industry, including hotels, travel and others.", italics: true, color: MUTED, size: 19 })] }),
  new Paragraph({ children: [new TextRun({ text: "Upload-ready PDF for Field 4: PRAHARI_Additional_Document.pdf · Copy-paste text for Fields 1–3: portal_text/", italics: true, color: MUTED, size: 19 })] }),
];

const body = [];
SECTIONS.forEach((s) => {
  body.push(new Paragraph({ children: [new PageBreak()] }));
  body.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(s.heading)] }));
  body.push(...convert(fs.readFileSync(path.join(ROOT, s.file), "utf8"), true));
});

const doc = new Document({
  creator: "Team PRAHARI",
  title: "PRAHARI: SIH Idea Submission",
  styles: {
    default: { document: { run: { font: "Calibri", size: 21 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 34, bold: true, color: NAVY, font: "Arial" }, paragraph: { spacing: { before: 120, after: 240 }, outlineLevel: 0, border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: SAFFRON, space: 4 } } } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 27, bold: true, color: NAVY, font: "Arial" }, paragraph: { spacing: { before: 300, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 23, bold: true, color: TEAL, font: "Arial" }, paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2 } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [0, 1, 2].map((l) => ({ level: l, format: LevelFormat.BULLET, text: ["•", "◦", "▪"][l], alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360 * (l + 1), hanging: 260 } } } })) },
      { reference: "numbers", levels: [0, 1].map((l) => ({ level: l, format: LevelFormat.DECIMAL, text: `%${l + 1}.`, alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400 * (l + 1), hanging: 300 } } } })) },
    ],
  },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "PRAHARI · SIH Idea Submission · Page ", color: MUTED, size: 16 }), new TextRun({ children: [PageNumber.CURRENT], color: MUTED, size: 16 })] })] }) },
    children: [...cover, ...body],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  const out = path.join(ROOT, "PRAHARI_SIH_Idea_Submission.docx");
  fs.writeFileSync(out, buf);
  console.log("wrote", out, buf.length, "bytes");
});

// AltaCV-style modern resume layout helpers.
// Ported from build_resume_modern.js — only the inline data has moved to YAML.

const {
  Document, Paragraph, TextRun, ExternalHyperlink, AlignmentType, LevelFormat,
  Table, TableRow, TableCell, WidthType, BorderStyle,
  TabStopType, VerticalAlign, UnderlineType,
} = require("docx");

const FONT = "Calibri";
const FONT_HEADER = "Calibri";

const C = {
  black:    "000000",
  light:    "9A9A9A",
  textGray: "555555",
  bullet:   "2E2E2E",
};

const MARGIN = 1008;                       // 0.7" in DXA
const PAGE_W = 12240;                      // US Letter width in DXA
const PAGE_H = 15840;
const CW = PAGE_W - 2 * MARGIN;            // content width 10224 DXA

const T  = (text, opts = {}) => new TextRun({ text, font: FONT, size: 22, color: C.bullet, ...opts });
const TB = (text, opts = {}) => new TextRun({ text, font: FONT, size: 22, color: C.bullet, bold: true, ...opts });

const NB = {
  top:    { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left:   { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right:  { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
};

// Section heading with horizontal rule extending to the right edge.
// 2-cell borderless table — cell 1 holds the heading, cell 2 has a bottom border.
function sectionHeading(text) {
  const headW = Math.max(1400, text.length * 200 + 400);
  const restW = CW - headW;
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [headW, restW],
    borders: {
      top: NB.top, bottom: NB.bottom, left: NB.left, right: NB.right,
      insideHorizontal: NB.top, insideVertical: NB.left,
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: NB,
            width: { size: headW, type: WidthType.DXA },
            margins: { top: 140, bottom: 0, left: 0, right: 100 },
            verticalAlign: VerticalAlign.BOTTOM,
            children: [
              new Paragraph({
                spacing: { before: 0, after: 0 },
                children: [
                  new TextRun({ text, font: FONT_HEADER, size: 36, bold: true, color: C.black }),
                ],
              }),
            ],
          }),
          new TableCell({
            borders: {
              top: NB.top,
              bottom: { style: BorderStyle.SINGLE, size: 8, color: C.black, space: 1 },
              left: NB.left, right: NB.right,
            },
            width: { size: restW, type: WidthType.DXA },
            margins: { top: 0, bottom: 0, left: 0, right: 0 },
            verticalAlign: VerticalAlign.BOTTOM,
            children: [
              new Paragraph({
                spacing: { before: 0, after: 0 },
                children: [new TextRun({ text: " ", font: FONT, size: 16 })],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

function entryLine1(left, right) {
  return new Paragraph({
    spacing: { before: 80, after: 0 },
    tabStops: [{ type: TabStopType.RIGHT, position: CW }],
    children: [
      new TextRun({ text: left, font: FONT, size: 24, bold: true, color: C.black }),
      new TextRun({ text: "\t", font: FONT, size: 24 }),
      new TextRun({ text: right, font: FONT, size: 22, italics: true, color: C.textGray }),
    ],
  });
}

function entryLine2(left, right) {
  return new Paragraph({
    spacing: { before: 0, after: 30 },
    tabStops: [{ type: TabStopType.RIGHT, position: CW }],
    children: [
      new TextRun({ text: left, font: FONT, size: 20, smallCaps: true, color: C.textGray }),
      new TextRun({ text: "\t", font: FONT, size: 20 }),
      new TextRun({ text: right, font: FONT, size: 20, italics: true, color: C.textGray }),
    ],
  });
}

function bullet(children) {
  return new Paragraph({
    numbering: { reference: "b", level: 0 },
    spacing: { before: 10, after: 10 },
    children,
  });
}

// Render an array of {text, bold?} runs (or plain strings) into TextRun[] for a bullet.
function renderRuns(runs) {
  return runs.map((r) => {
    if (typeof r === "string") return T(r);
    if (r.bold) return TB(r.text);
    return T(r.text);
  });
}

function headerName(nameFirst, nameLast) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 40 },
    children: [
      new TextRun({ text: `${nameFirst} `, font: FONT_HEADER, size: 72, color: C.light }),
      new TextRun({ text: nameLast, font: FONT_HEADER, size: 72, bold: true, color: C.black }),
    ],
  });
}

function headerContact(parts, releaseUrl) {
  const children = [];
  parts.forEach((p, i) => {
    if (i > 0) children.push(new TextRun({ text: "  |  ", font: FONT, size: 22, color: C.light }));
    children.push(new TextRun({ text: p, font: FONT, size: 22, color: C.bullet }));
  });
  if (releaseUrl) {
    children.push(new TextRun({ text: "  |  ", font: FONT, size: 22, color: C.light }));
    children.push(new ExternalHyperlink({
      link: releaseUrl,
      children: [
        new TextRun({
          text: "↗ Latest version",
          font: FONT,
          size: 22,
          color: "1F4E79",
          underline: { type: UnderlineType.SINGLE, color: "1F4E79" },
        }),
      ],
    }));
  }
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 60 },
    children,
  });
}

function skillLine(label, value, isLast) {
  return new Paragraph({
    spacing: { before: 0, after: 20 },
    children: [TB(`${label}:   `), T(value)],
  });
}

function buildDocument(body) {
  return new Document({
    styles: { default: { document: { run: { font: FONT, size: 22, color: C.bullet } } } },
    numbering: {
      config: [{
        reference: "b",
        levels: [{
          level: 0,
          format: LevelFormat.BULLET,
          text: "•",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 360, hanging: 240 } } },
        }],
      }],
    },
    sections: [{
      properties: {
        page: {
          size: { width: PAGE_W, height: PAGE_H },
          margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
        },
      },
      children: body,
    }],
  });
}

module.exports = {
  sectionHeading,
  entryLine1,
  entryLine2,
  bullet,
  renderRuns,
  headerName,
  headerContact,
  skillLine,
  buildDocument,
};

const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  AlignmentType, BorderStyle, WidthType, ShadingType, HeadingLevel, PageBreak
} = require("docx");

const dir = __dirname;

const img8356 = fs.readFileSync(path.join(dir, "IMG_8356_doc.jpg"));
const img8357 = fs.readFileSync(path.join(dir, "IMG_8357_doc.jpg"));
const img8358 = fs.readFileSync(path.join(dir, "IMG_8358_doc.jpg"));

const border = { style: BorderStyle.SINGLE, size: 1, color: "999999" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0 };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

const cellMargins = { top: 60, bottom: 60, left: 100, right: 100 };

function headerCell(text, width) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: "D9E2F3", type: ShadingType.CLEAR },
    margins: cellMargins,
    verticalAlign: "center",
    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text, bold: true, font: "游ゴシック", size: 20 })] })],
  });
}

function dataCell(text, width, align) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    margins: cellMargins,
    children: [new Paragraph({ alignment: align || AlignmentType.LEFT, children: [new TextRun({ text, font: "游ゴシック", size: 20 })] })],
  });
}

function amountCell(amount, width) {
  return dataCell(amount.toLocaleString(), width, AlignmentType.RIGHT);
}

function totalRow(label, amount, colWidths) {
  return new TableRow({
    children: [
      new TableCell({
        borders,
        width: { size: colWidths[0], type: WidthType.DXA },
        shading: { fill: "F2F2F2", type: ShadingType.CLEAR },
        margins: cellMargins,
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: label, bold: true, font: "游ゴシック", size: 20 })] })],
      }),
      new TableCell({
        borders,
        width: { size: colWidths[1], type: WidthType.DXA },
        columnSpan: 2,
        shading: { fill: "F2F2F2", type: ShadingType.CLEAR },
        margins: cellMargins,
        children: [new Paragraph({ children: [] })],
      }),
      new TableCell({
        borders,
        width: { size: colWidths[3], type: WidthType.DXA },
        shading: { fill: "F2F2F2", type: ShadingType.CLEAR },
        margins: cellMargins,
        children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "¥" + amount.toLocaleString(), bold: true, font: "游ゴシック", size: 20 })] })],
      }),
    ],
  });
}

const p = (text, options) => new Paragraph({ spacing: { after: 120 }, ...options, children: [new TextRun({ font: "游ゴシック", size: 22, ...( typeof text === "string" ? { text } : text ) })] });
const heading = (text, level) => new Paragraph({ heading: level, spacing: { before: 240, after: 120 }, children: [new TextRun({ text, font: "游ゴシック", bold: true, size: level === HeadingLevel.HEADING_1 ? 32 : 26 })] });

// Income table
const incomeColWidths = [2000, 3500, 1800, 1700];
const incomeTable = new Table({
  width: { size: 9000, type: WidthType.DXA },
  columnWidths: incomeColWidths,
  rows: [
    new TableRow({ children: [headerCell("項目", 2000), headerCell("摘要", 3500), headerCell("支払方法", 1800), headerCell("金額", 1700)] }),
    new TableRow({ children: [dataCell("参加費", 2000, AlignmentType.CENTER), dataCell("2,000円 × 17名", 3500), dataCell("現金", 1800, AlignmentType.CENTER), amountCell(34000, 1700)] }),
    new TableRow({ children: [dataCell("補助金", 2000, AlignmentType.CENTER), dataCell("近畿北陸支部", 3500), dataCell("現金", 1800, AlignmentType.CENTER), amountCell(12000, 1700)] }),
    totalRow("収入合計", 46000, incomeColWidths),
  ],
});

// Expense table
const expColWidths = [1800, 2800, 1400, 1400, 1600];
const expTableWidth = expColWidths.reduce((a, b) => a + b, 0);

function expHeaderRow() {
  return new TableRow({ children: [headerCell("項目", 1800), headerCell("摘要", 2800), headerCell("支払方法", 1400), headerCell("状態", 1400), headerCell("金額", 1600)] });
}
function expRow(item, desc, method, status, amount) {
  return new TableRow({
    children: [
      dataCell(item, 1800, AlignmentType.CENTER),
      dataCell(desc, 2800),
      dataCell(method, 1400, AlignmentType.CENTER),
      dataCell(status, 1400, AlignmentType.CENTER),
      new TableCell({
        borders, width: { size: 1600, type: WidthType.DXA }, margins: cellMargins,
        children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "¥" + amount.toLocaleString(), font: "游ゴシック", size: 20 })] })],
      }),
    ],
  });
}
function expTotalRow(label, amount) {
  return new TableRow({
    children: [
      new TableCell({ borders, width: { size: 1800, type: WidthType.DXA }, shading: { fill: "F2F2F2", type: ShadingType.CLEAR }, margins: cellMargins, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: label, bold: true, font: "游ゴシック", size: 20 })] })] }),
      new TableCell({ borders, width: { size: 2800, type: WidthType.DXA }, columnSpan: 3, shading: { fill: "F2F2F2", type: ShadingType.CLEAR }, margins: cellMargins, children: [new Paragraph({ children: [] })] }),
      new TableCell({ borders, width: { size: 1600, type: WidthType.DXA }, shading: { fill: "F2F2F2", type: ShadingType.CLEAR }, margins: cellMargins, children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "¥" + amount.toLocaleString(), bold: true, font: "游ゴシック", size: 20 })] })] }),
    ],
  });
}

const expenseTable = new Table({
  width: { size: expTableWidth, type: WidthType.DXA },
  columnWidths: expColWidths,
  rows: [
    expHeaderRow(),
    expRow("交通費補助", "5,000円 × 4名", "現金", "済", 20000),
    expRow("交通費補助", "3,000円 × 1名", "現金", "済", 3000),
    expRow("交通費補助", "2,000円 × 2名", "現金", "済", 4000),
    expRow("交通費補助", "1,000円 × 1名", "現金", "済", 1000),
    expRow("お土産代", "印刷レシート", "現金", "済", 3240),
    expRow("お土産代", "領収書", "現金", "済", 1296),
    expRow("備品", "Seria", "現金", "済", 1433),
    expRow("飲食", "ボンとらや", "現金", "済", 4545),
    expRow("飲食", "Waltz / Just Coffee", "現金", "済", 2400),
    expRow("飲食", "APITA", "現金", "済", 2890),
    expRow("会議室", "蒲郡商工会議所", "請求書", "精算済", 5690),
    expRow("交通費補助", "田中さん 1,000円 × 1名", "現金", "未精算", 1000),
    new TableRow({
      children: [
        dataCell("会議室", 1800, AlignmentType.CENTER),
        dataCell("初日ワークショップ専用", 2800),
        dataCell("未定", 1400, AlignmentType.CENTER),
        dataCell("未精算", 1400, AlignmentType.CENTER),
        new TableCell({ borders, width: { size: 1600, type: WidthType.DXA }, margins: cellMargins,
          children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "4〜5万円見込", font: "游ゴシック", size: 18 })] })] }),
      ],
    }),
    expTotalRow("支出合計（精算済分）", 49494),
  ],
});

// Summary table
const sumColWidths = [5000, 4000];
const sumTable = new Table({
  width: { size: 9000, type: WidthType.DXA },
  columnWidths: sumColWidths,
  rows: [
    new TableRow({ children: [
      new TableCell({ borders, width: { size: 5000, type: WidthType.DXA }, shading: { fill: "D9E2F3", type: ShadingType.CLEAR }, margins: cellMargins, children: [new Paragraph({ children: [new TextRun({ text: "出金額（元手）", bold: true, font: "游ゴシック", size: 22 })] })] }),
      new TableCell({ borders, width: { size: 4000, type: WidthType.DXA }, margins: cellMargins, children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "¥95,000", font: "游ゴシック", size: 22 })] })] }),
    ] }),
    new TableRow({ children: [
      new TableCell({ borders, width: { size: 5000, type: WidthType.DXA }, shading: { fill: "D9E2F3", type: ShadingType.CLEAR }, margins: cellMargins, children: [new Paragraph({ children: [new TextRun({ text: "現金収入合計", bold: true, font: "游ゴシック", size: 22 })] })] }),
      new TableCell({ borders, width: { size: 4000, type: WidthType.DXA }, margins: cellMargins, children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "+¥46,000", font: "游ゴシック", size: 22 })] })] }),
    ] }),
    new TableRow({ children: [
      new TableCell({ borders, width: { size: 5000, type: WidthType.DXA }, shading: { fill: "D9E2F3", type: ShadingType.CLEAR }, margins: cellMargins, children: [new Paragraph({ children: [new TextRun({ text: "現金支出合計（精算済分）", bold: true, font: "游ゴシック", size: 22 })] })] }),
      new TableCell({ borders, width: { size: 4000, type: WidthType.DXA }, margins: cellMargins, children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "-¥43,804", font: "游ゴシック", size: 22 })] })] }),
    ] }),
    new TableRow({ children: [
      new TableCell({ borders, width: { size: 5000, type: WidthType.DXA }, shading: { fill: "E2EFDA", type: ShadingType.CLEAR }, margins: cellMargins, children: [new Paragraph({ children: [new TextRun({ text: "現金残高", bold: true, font: "游ゴシック", size: 24 })] })] }),
      new TableCell({ borders, width: { size: 4000, type: WidthType.DXA }, shading: { fill: "E2EFDA", type: ShadingType.CLEAR }, margins: cellMargins, children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "¥97,196", bold: true, font: "游ゴシック", size: 24 })] })] }),
    ] }),
  ],
});

// Image helper
function imgParagraph(data, w, h) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 200 },
    children: [new ImageRun({ type: "jpg", data, transformation: { width: w, height: h }, altText: { title: "証憑", description: "レシート・領収書画像", name: "receipt" } })],
  });
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: "游ゴシック", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 32, bold: true, font: "游ゴシック" }, paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 26, bold: true, font: "游ゴシック" }, paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1260, bottom: 1440, left: 1260 },
      },
    },
    children: [
      // Title
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 }, children: [new TextRun({ text: "SAAJ 3支部合同研究会", font: "游ゴシック", size: 36, bold: true })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [new TextRun({ text: "会計報告書（簡易版）", font: "游ゴシック", size: 30, bold: true })] }),

      // Date & author
      new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 80 }, children: [new TextRun({ text: "2026年6月21日", font: "游ゴシック", size: 22 })] }),
      new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 300 }, children: [new TextRun({ text: "作成者: 加藤 智康", font: "游ゴシック", size: 22 })] }),

      // Overview
      heading("1. 収支概要", HeadingLevel.HEADING_1),
      sumTable,
      new Paragraph({ spacing: { before: 120, after: 40 }, children: [new TextRun({ text: "※ 蒲郡商工会議所（¥5,690 請求書払い）は現金残高に含まず", font: "游ゴシック", size: 18, italics: true })] }),
      new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "※ 田中さん交通費（¥1,000）および初日ワークショップ会議室代（4〜5万円見込み）は未精算", font: "游ゴシック", size: 18, italics: true })] }),
      new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "※ JISTA中部支部からの入金 ¥12,000 は未着", font: "游ゴシック", size: 18, italics: true })] }),

      // Income
      heading("2. 収入明細", HeadingLevel.HEADING_1),
      incomeTable,
      new Paragraph({ spacing: { after: 200 }, children: [] }),

      // Expense
      heading("3. 支出明細", HeadingLevel.HEADING_1),
      expenseTable,
      new Paragraph({ spacing: { after: 200 }, children: [] }),

      // Outstanding
      heading("4. 交通費補助の基準", HeadingLevel.HEADING_1),
      p("今回の交通費補助は、以下の基準で支給いたしました。"),
      p("・遠方または宿泊を伴い、交通費1万円以上の見込みの方: 5,000円"),
      p("・岐阜県等から参加、交通費5,000円程度の見込みの方: 3,000円"),
      p("・2日間参加で片道1,000円程度（名古屋市内想定）の方: 2,000円"),
      p("・1日のみ参加で往復2,000円程度の方: 1,000円"),
      new Paragraph({ spacing: { after: 200 }, children: [] }),

      heading("5. 未精算・未入金事項", HeadingLevel.HEADING_1),
      p("以下の項目は精算・入金が完了次第、追って報告いたします。"),
      p({ text: "【未精算（支出）】", bold: true }),
      p("・田中さん交通費: ¥1,000（現金未払い）"),
      p("・初日ワークショップ専用会議室代: 金額未確定（4〜5万円の見込み）"),
      new Paragraph({ spacing: { after: 80 }, children: [] }),
      p({ text: "【未入金（収入）】", bold: true }),
      p("・JISTA中部支部からの入金: ¥12,000（未入金）"),
      new Paragraph({ spacing: { after: 200 }, children: [] }),

      // Receipts
      new Paragraph({ children: [new PageBreak()] }),
      heading("6. 証憑画像", HeadingLevel.HEADING_1),
      p("レシート・請求書類", { spacing: { after: 100 } }),
      imgParagraph(img8356, 340, 453),

      p("領収書（交通費等）", { spacing: { before: 200, after: 100 } }),
      imgParagraph(img8357, 340, 453),

      p("領収書（交通費支給）", { spacing: { before: 200, after: 100 } }),
      imgParagraph(img8358, 340, 453),
    ],
  }],
});

const outPath = path.join(dir, "20260621_SAAJ3支部合同研究会_会計報告_v4.docx");
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(outPath, buf);
  console.log("Created: " + outPath);
});

const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, BorderStyle, WidthType, ShadingType,
        LevelFormat, Header, Footer, PageNumber } = require('docx');
const fs = require('fs');

const ACCENT    = "2E5FA3";
const WHITE     = "FFFFFF";
const LIGHT_BG  = "EDF2FB";
const BORDER_C  = "ADBFDF";
const TEXT_DARK = "1A1A1A";
const TEXT_MUTED= "666666";

const PAGE_W   = 11906;
const PAGE_H   = 16838;
const MARGIN   = 1440;
const CONTENT_W= PAGE_W - MARGIN * 2; // 9026

const border = (color) => ({ style: BorderStyle.SINGLE, size: 1, color });
const cellBorders = {
  top: border(BORDER_C), bottom: border(BORDER_C),
  left: border(BORDER_C), right: border(BORDER_C)
};

function sectionTitle(text) {
  return new Paragraph({
    spacing: { before: 360, after: 140 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 4 } },
    children: [new TextRun({ text, font: "Meiryo", size: 24, bold: true, color: ACCENT })]
  });
}

function body(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 80, after: 80 },
    children: [new TextRun({ text, font: "Meiryo", size: 20, color: TEXT_DARK, ...opts })]
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, font: "Meiryo", size: 20, color: TEXT_DARK })]
  });
}

function space() {
  return new Paragraph({ spacing: { before: 80, after: 80 }, children: [new TextRun({ text: "" })] });
}

function hCell(text, width) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: { fill: ACCENT, type: ShadingType.CLEAR },
    borders: cellBorders,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      children: [new TextRun({ text, font: "Meiryo", size: 18, bold: true, color: WHITE })]
    })]
  });
}

function dCell(text, width, even = false) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: { fill: even ? LIGHT_BG : WHITE, type: ShadingType.CLEAR },
    borders: cellBorders,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      children: [new TextRun({ text, font: "Meiryo", size: 18, color: TEXT_DARK })]
    })]
  });
}

const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{
        level: 0,
        format: LevelFormat.BULLET,
        text: "\u25CF",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 560, hanging: 280 } } }
      }]
    }]
  },
  styles: {
    default: { document: { run: { font: "Meiryo", size: 20 } } }
  },
  sections: [{
    properties: {
      page: {
        size: { width: PAGE_W, height: PAGE_H },
        margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: ACCENT, space: 4 } },
          children: [new TextRun({ text: "WPC2029 学生アドボカシープロジェクト", font: "Meiryo", size: 16, color: TEXT_MUTED })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: ACCENT, space: 4 } },
          children: [
            new TextRun({ text: "NPO法人てんびん  |  ", font: "Meiryo", size: 16, color: TEXT_MUTED }),
            new TextRun({ children: [PageNumber.CURRENT], font: "Meiryo", size: 16, color: TEXT_MUTED }),
          ]
        })]
      })
    },
    children: [

      // ===== タイトル =====
      new Paragraph({
        spacing: { before: 0, after: 180 },
        children: [new TextRun({ text: "WPC2029 学生アドボカシープロジェクト", font: "Meiryo", size: 40, bold: true, color: ACCENT })]
      }),
      new Paragraph({
        spacing: { before: 0, after: 80 },
        children: [new TextRun({ text: "パーキンソン病の課題に、自分たちの専門性で向き合う仲間を探しています", font: "Meiryo", size: 22, color: TEXT_DARK })]
      }),
      new Paragraph({
        spacing: { before: 0, after: 360 },
        children: [new TextRun({ text: "主催：NPO法人てんびん　2026年", font: "Meiryo", size: 18, color: TEXT_MUTED })]
      }),

      // ===== 1. このプロジェクトでやること =====
      sectionTitle("1. このプロジェクトでやること"),
      body("NPO法人てんびんと連携した高校・大学が、パーキンソン病に関するテーマを「半年サイクル」で探究・発表するプロジェクトです。"),
      space(),
      body("【半年サイクル制】", { bold: true }),
      bullet("各期の参加チームは、半年間テーマを探究し、期末に国内発表の場で成果を発表します"),
      bullet("発表で終わりではなく、次の期のチームがそのテーマを掘り下げるか、実践へ発展させるかを引き継ぎます"),
      bullet("メンバーは期ごとに変わってよい。「人」ではなく「テーマと知見」が積み重なる構造です"),
      space(),
      body("【WPC2029への発信】", { bold: true }),
      bullet("てんびんが複数期にわたる成果を整理・統合し、世界パーキンソン病学会（WPC2029）で発信します"),
      bullet("発表の形式・分野は問いません（映像・アート・学術研究・政策提言・ワークショップなど）"),
      space(),

      // ===== 2. 参加メンバーと役割 =====
      sectionTitle("2. 参加メンバーと役割"),
      new Table({
        width: { size: CONTENT_W, type: WidthType.DXA },
        columnWidths: [2800, 6226],
        rows: [
          new TableRow({ children: [hCell("役割", 2800), hCell("内容", 6226)] }),
          new TableRow({ children: [
            dCell("コーディネーター（NPO）", 2800, false),
            dCell("場の設計・当事者へのアクセス提供・各期の進行管理・WPC2029への蓄積整理", 6226, false)
          ]}),
          new TableRow({ children: [
            dCell("参加チーム（高校・大学）", 2800, true),
            dCell("半年間のテーマ探究・期末発表。複数校混合チームも可", 6226, true)
          ]}),
          new TableRow({ children: [
            dCell("次期参加チーム", 2800, false),
            dCell("前期の成果を受け取り、テーマをどう継承・発展させるかを自ら設計する", 6226, false)
          ]}),
          new TableRow({ children: [
            dCell("パーキンソン病当事者・家族", 2800, true),
            dCell("活動への協力・意見提供・各期の発表に参加", 6226, true)
          ]}),
        ]
      }),
      space(),

      // ===== 3. ステークホルダーそれぞれの意義 =====
      sectionTitle("3. ステークホルダーそれぞれの意義"),
      new Table({
        width: { size: CONTENT_W, type: WidthType.DXA },
        columnWidths: [2800, 6226],
        rows: [
          new TableRow({ children: [hCell("ステークホルダー", 2800), hCell("この企画から得られる意義", 6226)] }),
          new TableRow({ children: [
            dCell("NPO法人てんびん", 2800, false),
            dCell("若い世代との継続的な連携・テーマの蓄積・WPC2029での国際発信力強化", 6226, false)
          ]}),
          new TableRow({ children: [
            dCell("参加チームの学生（共通）", 2800, true),
            dCell("半年間の実践経験・国内発表の実績・本物の社会課題への参加・異分野連携体験", 6226, true)
          ]}),
          new TableRow({ children: [
            dCell("高校生", 2800, false),
            dCell("総合型選抜での活動実績・社会へ提言する体験（半年で完結するため参加しやすい）", 6226, false)
          ]}),
          new TableRow({ children: [
            dCell("大学生", 2800, true),
            dCell("研究・就活における実績・社会実装の経験（ゼミ・研究室単位での参加も可）", 6226, true)
          ]}),
          new TableRow({ children: [
            dCell("パーキンソン病当事者・家族", 2800, false),
            dCell("多様な視点からの継続的なアドボカシー・次世代との継続的な対話", 6226, false)
          ]}),
        ]
      }),
      space(),

      // ===== 4. サイクルとテーマ継承のイメージ =====
      sectionTitle("4. サイクルとテーマ継承のイメージ"),
      body("以下はテーマが「スティグマと学生」から発展する場合の例です。実際のテーマは参加チームが設定します。"),
      space(),
      new Table({
        width: { size: CONTENT_W, type: WidthType.DXA },
        columnWidths: [1400, 2200, 5426],
        rows: [
          new TableRow({ children: [hCell("期", 1400), hCell("テーマの方向", 2200), hCell("活動・発表のイメージ（例）", 5426)] }),
          new TableRow({ children: [
            dCell("第1期（半年）", 1400, false),
            dCell("探究・提言", 2200, false),
            dCell("「パーキンソン病のスティグマと学生にできること」を調査し、医療者・行政への提言を発表", 5426, false)
          ]}),
          new TableRow({ children: [
            dCell("第2期（半年）", 1400, true),
            dCell("掘り下げ・発展", 2200, true),
            dCell("前期の提言を受け、「スティグマを減らす映像表現」をテーマにアニメーション制作・上映", 5426, true)
          ]}),
          new TableRow({ children: [
            dCell("とりまとめ（約1年）", 1400, false),
            dCell("てんびんが統合・整理", 2200, false),
            dCell("てんびんが2期分の成果を整理・編集し、WPC2029の発表コンテンツを仕上げる", 5426, false)
          ]}),
          new TableRow({ children: [
            dCell("WPC2029", 1400, true),
            dCell("世界へ発信", 2200, true),
            dCell("「日本の学生×パーキンソン病」として国際学会で発表。次の世代へ引き継ぐ", 5426, true)
          ]}),
        ]
      }),
      space(),
      body("※ 各期のテーマは次期チームが自由に設定します。前期を掘り下げるか、全く新しい角度で入り直してもかまいません。", { color: TEXT_MUTED }),
      space(),

      // ===== 5. 懸念への回答 =====
      sectionTitle("5. 懸念への回答"),
      body("Q：学生がお客様になってしまうのでは？", { bold: true }),
      body("A：半年サイクル制にすることで、この問題を構造的に解決します。"),
      bullet("参加は半年間の完結型。「3年間続けなければ」というプレッシャーがない"),
      bullet("期末発表が必須なので、参加するだけで終わる構造になっていない"),
      bullet("前期の成果をどう使うかを自分たちで設計するため、次期チームも主体的に関わる"),
      space(),
      body("それでも以下の3条件を参加の前提として明示し、自己選別を促します："),
      bullet("自分たちの視点・テーマを持っていること"),
      bullet("主体的に動けること（「参加させてもらう」スタンスは不可）"),
      bullet("定期連携（月1〜2回程度）に参加できること"),
      space(),

      // ===== 6. 次のステップ =====
      sectionTitle("6. 次のステップ"),
      bullet("第1期の参加チーム募集開始（2026年内）"),
      bullet("キックオフミーティングの設計・日程調整"),
      bullet("国内発表の場の設計（開催時期・場所・形式）"),
      bullet("テーマ継承の仕組みづくり（引き継ぎ資料・記録のフォーマット）"),
      bullet("WPC2029の発表申し込みに関する情報収集"),
      space(),

      new Paragraph({
        spacing: { before: 200, after: 0 },
        alignment: AlignmentType.RIGHT,
        children: [new TextRun({ text: "2026年6月　NPO法人てんびん", font: "Meiryo", size: 18, color: TEXT_MUTED })]
      }),
    ]
  }]
});

const outputPath = "C:/Users/KATOTO/Downloads/資料作成/shiryo/10bin/20260618_10bin_WPC2029学生企画_企画書_v3.docx";
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outputPath, buffer);
  console.log("保存完了: " + outputPath);
}).catch(err => { console.error(err); process.exit(1); });

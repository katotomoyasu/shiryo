const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "WPC2029学生アドボカシープロジェクト 運営体制 課題メモ";

// カラーパレット
const C = {
  dark:    "1E2761",
  mid:     "3D5A99",
  light:   "CADCFC",
  white:   "FFFFFF",
  accent:  "F96167",
  gray:    "64748B",
  bgLight: "F4F7FC",
};

const makeShadow = () => ({ type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.12 });

// ─── スライド1：タイトル ───────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.dark };

  // 背景装飾
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 4.2, w: 10, h: 1.425, fill: { color: C.mid }, line: { color: C.mid } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: C.accent }, line: { color: C.accent } });

  s.addText("WPC2029 学生アドボカシープロジェクト", {
    x: 0.5, y: 1.1, w: 9, h: 0.7,
    fontSize: 20, fontFace: "Meiryo", color: C.light, bold: false,
  });
  s.addText("運営体制 課題メモ", {
    x: 0.5, y: 1.9, w: 9, h: 1.1,
    fontSize: 40, fontFace: "Meiryo", color: C.white, bold: true,
  });
  s.addText("2026年6月18日　NPO法人てんびん", {
    x: 0.5, y: 4.4, w: 9, h: 0.5,
    fontSize: 14, fontFace: "Meiryo", color: C.light,
  });
}

// ─── スライド2：現状 ──────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.bgLight };

  // ヘッダー帯
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.dark }, line: { color: C.dark } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: C.accent }, line: { color: C.accent } });
  s.addText("現状", { x: 0.4, y: 0.15, w: 9, h: 0.65, fontSize: 24, fontFace: "Meiryo", color: C.white, bold: true, margin: 0 });

  // 現状カード
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.3, w: 9, h: 1.5, fill: { color: C.white }, line: { color: C.light }, shadow: makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.3, w: 0.12, h: 1.5, fill: { color: C.accent }, line: { color: C.accent } });
  s.addText([
    { text: "先生（指導教員・アドバイザー）は未確保", options: { bold: true, breakLine: true } },
    { text: "これから探す段階", options: { breakLine: false } },
  ], { x: 0.8, y: 1.4, w: 8.5, h: 1.2, fontSize: 18, fontFace: "Meiryo", color: C.dark });

  // 必要な理由
  s.addText("なぜ先生が必要か", { x: 0.5, y: 3.1, w: 9, h: 0.45, fontSize: 16, fontFace: "Meiryo", color: C.gray, bold: true });
  s.addText([
    { text: "高校生の参加を「学校公認」にするために指導教員が不可欠（特に総合型選抜実績）", options: { bullet: true, breakLine: true } },
    { text: "大学生はゼミ・研究室単位での参加が想定されており、担当教員の後ろ盾が参加ハードルを下げる", options: { bullet: true, breakLine: true } },
    { text: "半年ごとにメンバーが変わる構造でも、担当教員がいることで継続性が保ちやすくなる", options: { bullet: true } },
  ], { x: 0.6, y: 3.55, w: 9, h: 1.8, fontSize: 13, fontFace: "Meiryo", color: C.dark });
}

// ─── スライド3：必要な先生の種類 ─────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.bgLight };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.dark }, line: { color: C.dark } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: C.accent }, line: { color: C.accent } });
  s.addText("必要な先生の種類", { x: 0.4, y: 0.15, w: 9, h: 0.65, fontSize: 24, fontFace: "Meiryo", color: C.white, bold: true, margin: 0 });

  // カード①
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.15, w: 4.3, h: 3.8, fill: { color: C.white }, line: { color: C.light }, shadow: makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.15, w: 4.3, h: 0.55, fill: { color: C.mid }, line: { color: C.mid } });
  s.addText("① チーム伴走型", { x: 0.5, y: 1.18, w: 4.1, h: 0.5, fontSize: 15, fontFace: "Meiryo", color: C.white, bold: true, margin: 0 });
  s.addText([
    { text: "各参加校ごとに1人", options: { bold: true, breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "高校：総合探究担当教員 or 顧問", options: { bullet: true, breakLine: true } },
    { text: "大学：ゼミ担当教員・指導教授", options: { bullet: true, breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "パーキンソン病の専門知識は不要", options: { breakLine: false } },
  ], { x: 0.55, y: 1.8, w: 4.0, h: 3.0, fontSize: 13, fontFace: "Meiryo", color: C.dark });

  // カード②
  s.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 1.15, w: 4.3, h: 3.8, fill: { color: C.white }, line: { color: C.light }, shadow: makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 1.15, w: 4.3, h: 0.55, fill: { color: C.accent }, line: { color: C.accent } });
  s.addText("② 専門アドバイザー型", { x: 5.4, y: 1.18, w: 4.1, h: 0.5, fontSize: 15, fontFace: "Meiryo", color: C.white, bold: true, margin: 0 });
  s.addText([
    { text: "企画全体に1〜2人", options: { bold: true, breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "医療・福祉・社会課題系の研究者", options: { bullet: true, breakLine: true } },
    { text: "WPC2029での発信に説得力をもたせるために必要", options: { bullet: true, breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "てんびんの医療系人脈から探せるか要確認", options: { breakLine: false } },
  ], { x: 5.45, y: 1.8, w: 4.0, h: 3.0, fontSize: 13, fontFace: "Meiryo", color: C.dark });
}

// ─── スライド4：優先アクション・探し方 ──────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.bgLight };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.dark }, line: { color: C.dark } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: C.accent }, line: { color: C.accent } });
  s.addText("優先アクションと探し方", { x: 0.4, y: 0.15, w: 9, h: 0.65, fontSize: 24, fontFace: "Meiryo", color: C.white, bold: true, margin: 0 });

  // 優先アクション
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.15, w: 9.2, h: 1.3, fill: { color: "FFF3F3" }, line: { color: C.accent }, shadow: makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.15, w: 0.12, h: 1.3, fill: { color: C.accent }, line: { color: C.accent } });
  s.addText("優先アクション", { x: 0.65, y: 1.2, w: 3, h: 0.4, fontSize: 13, fontFace: "Meiryo", color: C.accent, bold: true, margin: 0 });
  s.addText("まず専門アドバイザーを1人確保する", {
    x: 0.65, y: 1.58, w: 8.7, h: 0.6,
    fontSize: 17, fontFace: "Meiryo", color: C.dark, bold: true,
  });

  s.addText("→「○○先生が関わっている」という信頼性が、学校へのアプローチを一気に楽にする", {
    x: 0.6, y: 2.55, w: 9.1, h: 0.5,
    fontSize: 13, fontFace: "Meiryo", color: C.gray, italic: true,
  });

  // テーブル
  s.addText("探し方の候補", { x: 0.4, y: 3.05, w: 9, h: 0.4, fontSize: 14, fontFace: "Meiryo", color: C.dark, bold: true });

  const rows = [
    [
      { text: "対象", options: { bold: true, color: C.white, fill: { color: C.mid }, fontFace: "Meiryo", fontSize: 12 } },
      { text: "アプローチ方法", options: { bold: true, color: C.white, fill: { color: C.mid }, fontFace: "Meiryo", fontSize: 12 } },
    ],
    ["高校の先生", "探究実践校リストを活用して直接アプローチ"],
    ["大学の先生", "社会課題系ゼミ・医療福祉系学部に当たる"],
    ["アドバイザー", "WPC参加経験のある研究者・医師"],
  ];
  s.addTable(rows, {
    x: 0.4, y: 3.5, w: 9.2,
    colW: [2.5, 6.7],
    fontSize: 12, fontFace: "Meiryo", color: C.dark,
    border: { pt: 1, color: C.light },
    fill: { color: C.white },
  });
}

// ─── スライド5：未確認事項 ────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.bgLight };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.dark }, line: { color: C.dark } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: C.accent }, line: { color: C.accent } });
  s.addText("未確認事項（TODO）", { x: 0.4, y: 0.15, w: 9, h: 0.65, fontSize: 24, fontFace: "Meiryo", color: C.white, bold: true, margin: 0 });

  const todos = [
    "てんびんの医療系人脈でアドバイザー候補がいるか確認する",
    "声をかける高校・大学の候補リストを作成する",
  ];

  todos.forEach((todo, i) => {
    const y = 1.4 + i * 1.3;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y, w: 9.2, h: 1.0, fill: { color: C.white }, line: { color: C.light }, shadow: makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y, w: 0.55, h: 1.0, fill: { color: C.mid }, line: { color: C.mid } });
    s.addText(String(i + 1), { x: 0.4, y, w: 0.55, h: 1.0, fontSize: 20, fontFace: "Meiryo", color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(todo, { x: 1.1, y: y + 0.15, w: 8.3, h: 0.7, fontSize: 15, fontFace: "Meiryo", color: C.dark });
  });

  // フッター
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.2, w: 10, h: 0.425, fill: { color: C.dark }, line: { color: C.dark } });
  s.addText("2026年6月　NPO法人てんびん　WPC2029学生アドボカシープロジェクト", {
    x: 0.3, y: 5.22, w: 9.4, h: 0.35, fontSize: 10, fontFace: "Meiryo", color: C.light,
  });
}

pres.writeFile({ fileName: "C:\\Users\\katop\\Downloads\\shiryo\\10bin\\20260618_10bin_運営体制課題_v1.pptx" })
  .then(() => console.log("Done"))
  .catch(e => { console.error(e); process.exit(1); });

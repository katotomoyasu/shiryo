const pptxgen = require("pptxgenjs");

// === カラーパレット ===
const C = {
  bgWhite:   "FFFFFF",
  bgLight:   "F5F7FA",
  bgNavy:    "0B1F3F",
  textBlack: "1A1A1A",
  textDark:  "333333",
  textMuted: "888888",
  textWhite: "FFFFFF",
  royalBlue: "1B3A6B",
  blue2:     "3D6098",
  blue3:     "7A9CC6",
  blue4:     "B8CCE4",
  blue5:     "D9E5F2",
  accentRed: "C8102E",
  gridLine:  "D0D0D0",
  divider:   "1B3A6B",
};

// === レイアウト定数 ===
const L = {
  mx: 0.5, my: 0.5,
  contentX: 0.5, contentY: 1.2,
  contentW: 9.0, contentH: 3.6,
  titleX: 0.5, titleY: 0.42,
  titleW: 9.0, titleH: 0.55,
  dividerY: 1.02,
  sourceX: 0.5, sourceY: 5.15,
  pageNumX: 9.0, pageNumY: 5.15,
  col2W: 4.3, col2Gap: 0.4, col2RightX: 5.2,
  col3W: 2.73, col3Gap: 0.4,
};

// === フォントファクトリ ===
const font = {
  actionTitle: () => ({ fontFace: "Yu Mincho", fontSize: 17, bold: true, color: C.bgNavy, margin: 0 }),
  sectionHeader: () => ({ fontFace: "Yu Mincho", fontSize: 13, bold: true, color: C.bgNavy, margin: 0 }),
  body: () => ({ fontFace: "Yu Gothic", fontSize: 11, color: C.textDark, margin: 0 }),
  source: () => ({ fontFace: "Yu Gothic", fontSize: 8, color: C.textMuted, margin: 0 }),
  pageNum: () => ({ fontFace: "Calibri", fontSize: 8, color: C.textMuted, margin: 0 }),
};

function addDivider(slide, pres) {
  slide.addShape(pres.shapes.LINE, {
    x: L.mx, y: L.dividerY, w: L.contentW, h: 0,
    line: { color: C.divider, width: 0.5 }
  });
}

function addPageNum(slide, num) {
  slide.addText(String(num), {
    x: L.pageNumX, y: L.pageNumY, w: 0.5, h: 0.3,
    ...font.pageNum(), align: "right"
  });
}

async function buildDeck() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = "WPC2029 学生アドボカシープロジェクト企画書";

  // ===== スライド1: タイトル =====
  {
    const slide = pres.addSlide();
    slide.background = { color: C.bgNavy };

    slide.addText("WPC2029\n学生アドボカシープロジェクト", {
      x: 0.8, y: 1.1, w: 8.4, h: 1.8,
      fontFace: "Yu Mincho", fontSize: 34, bold: true,
      color: C.textWhite, align: "left", valign: "middle", margin: 0
    });

    slide.addText("同じ課題意識を持つ仲間を探しています", {
      x: 0.8, y: 3.1, w: 8.4, h: 0.5,
      fontFace: "Yu Gothic", fontSize: 15, color: C.blue3,
      align: "left", valign: "top", margin: 0
    });

    slide.addText("NPO法人てんびん  |  2026年", {
      x: 0.8, y: 4.65, w: 8.4, h: 0.3,
      fontFace: "Calibri", fontSize: 10, color: C.blue4,
      align: "left", margin: 0
    });
  }

  // ===== スライド2: WPC2026が示した世界の課題 =====
  {
    const slide = pres.addSlide();
    slide.background = { color: C.bgWhite };

    slide.addText("WPC2026リーダーズ会議が示した世界のパーキンソン病アドボカシー5大課題", {
      x: L.titleX, y: L.titleY, w: L.titleW, h: L.titleH,
      ...font.actionTitle(), align: "left", valign: "top"
    });
    addDivider(slide, pres);

    const issues = [
      { num: "①", title: "スティグマ・社会的孤立", desc: "偏見が当事者を社会から遠ざける。年齢を問わず包摂が必要。" },
      { num: "②", title: "移動・公共空間の障壁", desc: "信号時間・ベンチ・交通機関が自立した生活を制限する。" },
      { num: "③", title: "診断時支援の不足", desc: "診断直後に必要な情報・つながりが届いていない。" },
      { num: "④", title: "医療アクセスの地域格差", desc: "専門医・リハビリへのアクセスに国・地域間で大きな差がある。" },
      { num: "⑤", title: "国際連携体の必要性", desc: "各国単独では声が届かない。連携で政府・WHOへの働きかけを強化する。" },
    ];

    // 3+2 レイアウト
    issues.forEach((issue, i) => {
      const row = Math.floor(i / 3);
      const col = i % 3;
      const bw = row === 0 ? 2.85 : L.col2W;
      const gap = row === 0 ? 0.225 : L.col2Gap;
      const startX = row === 0 ? L.contentX : (L.contentX + (L.contentW - (L.col2W * 2 + L.col2Gap)) / 2);
      const bx = startX + col * (bw + gap);
      const by = L.contentY + row * 1.9;
      const bh = 1.72;

      slide.addShape(pres.shapes.RECTANGLE, {
        x: bx, y: by, w: bw, h: bh,
        fill: { color: i % 2 === 0 ? C.blue5 : C.bgLight },
        line: { color: C.gridLine, width: 0.5 }
      });

      slide.addText(issue.num, {
        x: bx + 0.12, y: by + 0.1, w: 0.42, h: 0.42,
        fontFace: "Georgia", fontSize: 16, bold: true,
        color: C.royalBlue, align: "center", valign: "middle", margin: 0
      });

      slide.addText(issue.title, {
        x: bx + 0.12, y: by + 0.56, w: bw - 0.24, h: 0.38,
        fontFace: "Yu Mincho", fontSize: 11, bold: true,
        color: C.bgNavy, align: "left", valign: "top", margin: 0
      });

      slide.addText(issue.desc, {
        x: bx + 0.12, y: by + 0.98, w: bw - 0.24, h: 0.62,
        ...font.body(), fontSize: 10, align: "left", valign: "top"
      });
    });

    slide.addText("出所: WPC2026 Leaders Meeting 議事録", {
      x: L.sourceX, y: L.sourceY, w: 6, h: 0.3,
      ...font.source(), align: "left"
    });
    addPageNum(slide, 2);
  }

  // ===== スライド3: 日本から何ができるか =====
  {
    const slide = pres.addSlide();
    slide.background = { color: C.bgWhite };

    slide.addText("「医療」だけでなく「社会・文化・教育」からの発信こそ、日本が世界に貢献できる領域だ", {
      x: L.titleX, y: L.titleY, w: L.titleW, h: L.titleH,
      ...font.actionTitle(), align: "left", valign: "top"
    });
    addDivider(slide, pres);

    const cols = [
      { label: "社会", color: C.royalBlue, items: ["公共空間・交通のデザイン提言", "スティグマ解消キャンペーン", "行政・企業への働きかけ"] },
      { label: "文化・アート", color: C.blue2, items: ["映像・アニメーションによる啓発", "アートセラピーの普及", "当事者の表現の場づくり"] },
      { label: "教育", color: C.blue3, items: ["学生による社会提言", "異分野連携プロジェクト", "次世代アドボカシーの育成"] },
    ];

    cols.forEach((col, i) => {
      const cx = L.contentX + i * (L.col3W + L.col3Gap);

      slide.addShape(pres.shapes.RECTANGLE, {
        x: cx, y: L.contentY, w: L.col3W, h: 0.42,
        fill: { color: col.color }, line: { color: col.color, width: 0 }
      });
      slide.addText(col.label, {
        x: cx, y: L.contentY, w: L.col3W, h: 0.42,
        fontFace: "Yu Mincho", fontSize: 14, bold: true,
        color: C.textWhite, align: "center", valign: "middle", margin: 0
      });

      slide.addShape(pres.shapes.RECTANGLE, {
        x: cx, y: L.contentY + 0.42, w: L.col3W, h: 2.3,
        fill: { color: C.bgLight }, line: { color: C.gridLine, width: 0.5 }
      });

      col.items.forEach((item, j) => {
        slide.addText("● " + item, {
          x: cx + 0.12, y: L.contentY + 0.55 + j * 0.62, w: L.col3W - 0.24, h: 0.55,
          ...font.body(), fontSize: 11, align: "left", valign: "top"
        });
      });
    });

    slide.addShape(pres.shapes.RECTANGLE, {
      x: L.contentX, y: 4.35, w: L.contentW, h: 0.55,
      fill: { color: C.bgNavy }, line: { color: C.bgNavy, width: 0 }
    });
    slide.addText("次世代の声が、世界を動かす。", {
      x: L.contentX, y: 4.35, w: L.contentW, h: 0.55,
      fontFace: "Yu Mincho", fontSize: 16, bold: true,
      color: C.textWhite, align: "center", valign: "middle", margin: 0
    });

    addPageNum(slide, 3);
  }

  // ===== スライド4: てんびんとは／灘高校との連携実績 =====
  {
    const slide = pres.addSlide();
    slide.background = { color: C.bgWhite };

    slide.addText("てんびんは、学生と共に社会課題に向き合い続けてきた実績を持つNPOだ", {
      x: L.titleX, y: L.titleY, w: L.titleW, h: L.titleH,
      ...font.actionTitle(), align: "left", valign: "top"
    });
    addDivider(slide, pres);

    // 左: てんびんとは
    slide.addShape(pres.shapes.RECTANGLE, {
      x: L.contentX, y: L.contentY, w: L.col2W, h: 0.38,
      fill: { color: C.royalBlue }, line: { color: C.royalBlue, width: 0 }
    });
    slide.addText("NPO法人てんびんとは", {
      x: L.contentX, y: L.contentY, w: L.col2W, h: 0.38,
      fontFace: "Yu Mincho", fontSize: 12, bold: true,
      color: C.textWhite, align: "center", valign: "middle", margin: 0
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: L.contentX, y: L.contentY + 0.38, w: L.col2W, h: 2.9,
      fill: { color: C.bgLight }, line: { color: C.gridLine, width: 0.5 }
    });

    const aboutItems = [
      "パーキンソン病の当事者・家族を支援するNPO",
      "PDヨガ・アートセラピー・無料相談会を主催",
      "「誰もが彩り豊かに生き合う未来」を理念に活動",
      "医療・文化・教育をつなぐハブとして機能",
    ];
    aboutItems.forEach((item, i) => {
      slide.addText("● " + item, {
        x: L.contentX + 0.15, y: L.contentY + 0.52 + i * 0.55, w: L.col2W - 0.3, h: 0.48,
        ...font.body(), fontSize: 11, align: "left", valign: "top"
      });
    });

    // 右: 灘高校との連携実績
    slide.addShape(pres.shapes.RECTANGLE, {
      x: L.col2RightX, y: L.contentY, w: L.col2W, h: 0.38,
      fill: { color: C.blue2 }, line: { color: C.blue2, width: 0 }
    });
    slide.addText("灘高校との連携実績", {
      x: L.col2RightX, y: L.contentY, w: L.col2W, h: 0.38,
      fontFace: "Yu Mincho", fontSize: 12, bold: true,
      color: C.textWhite, align: "center", valign: "middle", margin: 0
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: L.col2RightX, y: L.contentY + 0.38, w: L.col2W, h: 2.9,
      fill: { color: C.blue5 }, line: { color: C.gridLine, width: 0.5 }
    });

    slide.addText("2026年4月12日（日）", {
      x: L.col2RightX + 0.15, y: L.contentY + 0.52, w: L.col2W - 0.3, h: 0.32,
      fontFace: "Yu Gothic", fontSize: 11, bold: true,
      color: C.royalBlue, align: "left", valign: "top", margin: 0
    });
    slide.addText("兵庫県立美術館 KOBELCOミュージアムホール", {
      x: L.col2RightX + 0.15, y: L.contentY + 0.85, w: L.col2W - 0.3, h: 0.32,
      fontFace: "Yu Gothic", fontSize: 10, color: C.textDark,
      align: "left", valign: "top", margin: 0
    });
    slide.addText("「学生たちの発表会\n〜僕たち私たちがあなたのために出来ること〜」", {
      x: L.col2RightX + 0.15, y: L.contentY + 1.22, w: L.col2W - 0.3, h: 0.55,
      fontFace: "Yu Gothic", fontSize: 10, bold: true,
      color: C.bgNavy, align: "left", valign: "top", margin: 0
    });
    slide.addText("灘中学生が「パーキンソン病を取り巻く世界を良くするには」をテーマに、医療従事者・行政・当事者団体へ社会提言を実施。同志社女子大学生によるアニメーション啓発作品も同日発表。", {
      x: L.col2RightX + 0.15, y: L.contentY + 1.85, w: L.col2W - 0.3, h: 1.3,
      ...font.body(), fontSize: 10, align: "left", valign: "top"
    });

    addPageNum(slide, 4);
  }

  // ===== スライド5: コンセプト（Before/After） =====
  {
    const slide = pres.addSlide();
    slide.background = { color: C.bgWhite };

    slide.addText("この企画は「参加させてあげる」のではなく「一緒に世界を動かす仲間」を集めるプロジェクトだ", {
      x: L.titleX, y: L.titleY, w: L.titleW, h: L.titleH,
      ...font.actionTitle(), align: "left", valign: "top"
    });
    addDivider(slide, pres);

    // 左: よくある参加型
    slide.addShape(pres.shapes.RECTANGLE, {
      x: L.contentX, y: L.contentY, w: L.col2W, h: 0.38,
      fill: { color: C.blue4 }, line: { color: C.blue4, width: 0 }
    });
    slide.addText("よくある「参加型」企画", {
      x: L.contentX, y: L.contentY, w: L.col2W, h: 0.38,
      fontFace: "Yu Mincho", fontSize: 12, bold: true,
      color: C.bgNavy, align: "center", valign: "middle", margin: 0
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: L.contentX, y: L.contentY + 0.38, w: L.col2W, h: 2.9,
      fill: { color: C.bgLight }, line: { color: C.gridLine, width: 0.5 }
    });

    const oldItems = [
      "テーマは主催者が決める",
      "参加者はコンテンツを「受け取る」",
      "単発イベントで終わる",
      "お客様として扱われる",
    ];
    oldItems.forEach((item, i) => {
      slide.addText("x  " + item, {
        x: L.contentX + 0.15, y: L.contentY + 0.55 + i * 0.58, w: L.col2W - 0.3, h: 0.5,
        fontFace: "Yu Gothic", fontSize: 11, color: C.textMuted,
        align: "left", valign: "top", margin: 0
      });
    });

    // 矢印
    slide.addText("->", {
      x: L.contentX + L.col2W, y: L.contentY + 1.5, w: L.col2Gap, h: 0.5,
      fontFace: "Calibri", fontSize: 22, bold: true,
      color: C.royalBlue, align: "center", valign: "middle", margin: 0
    });

    // 右: この企画
    slide.addShape(pres.shapes.RECTANGLE, {
      x: L.col2RightX, y: L.contentY, w: L.col2W, h: 0.38,
      fill: { color: C.royalBlue }, line: { color: C.royalBlue, width: 0 }
    });
    slide.addText("この企画のスタンス", {
      x: L.col2RightX, y: L.contentY, w: L.col2W, h: 0.38,
      fontFace: "Yu Mincho", fontSize: 12, bold: true,
      color: C.textWhite, align: "center", valign: "middle", margin: 0
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: L.col2RightX, y: L.contentY + 0.38, w: L.col2W, h: 2.9,
      fill: { color: C.blue5 }, line: { color: C.gridLine, width: 0.5 }
    });

    const newItems = [
      "各校が自分たちのテーマを持ち込む",
      "学生が「つくる側」として主体的に参加",
      "WPC2029まで継続して連携する",
      "仲間として対等に動く",
    ];
    newItems.forEach((item, i) => {
      slide.addText("◎  " + item, {
        x: L.col2RightX + 0.15, y: L.contentY + 0.55 + i * 0.58, w: L.col2W - 0.3, h: 0.5,
        fontFace: "Yu Gothic", fontSize: 11, bold: true, color: C.royalBlue,
        align: "left", valign: "top", margin: 0
      });
    });

    addPageNum(slide, 5);
  }

  // ===== スライド6: 参加校に期待すること =====
  {
    const slide = pres.addSlide();
    slide.background = { color: C.bgWhite };

    slide.addText("参加校には「テーマと主体性」だけを求める—発表形式・分野・規模は一切問わない", {
      x: L.titleX, y: L.titleY, w: L.titleW, h: L.titleH,
      ...font.actionTitle(), align: "left", valign: "top"
    });
    addDivider(slide, pres);

    slide.addText("発表形式の例（これに限らない）", {
      x: L.contentX, y: L.contentY, w: L.contentW, h: 0.32,
      ...font.sectionHeader(), align: "left", valign: "top"
    });

    const formats = [
      "映像・ドキュメンタリー", "アート・デザイン作品",
      "学術研究・論文発表", "政策提言・ロビー活動",
      "ワークショップ企画", "アニメーション・SNS発信",
    ];
    formats.forEach((fmt, i) => {
      const col = i % 3;
      const row = Math.floor(i / 3);
      const fx = L.contentX + col * 3.05;
      const fy = L.contentY + 0.4 + row * 0.55;

      slide.addShape(pres.shapes.RECTANGLE, {
        x: fx, y: fy, w: 2.85, h: 0.42,
        fill: { color: C.blue5 }, line: { color: C.gridLine, width: 0.5 }
      });
      slide.addText(fmt, {
        x: fx + 0.1, y: fy, w: 2.65, h: 0.42,
        fontFace: "Yu Gothic", fontSize: 11, bold: true, color: C.royalBlue,
        align: "center", valign: "middle", margin: 0
      });
    });

    slide.addText("お願いすること", {
      x: L.contentX, y: L.contentY + 1.6, w: L.contentW, h: 0.32,
      ...font.sectionHeader(), align: "left", valign: "top"
    });

    const asks = [
      "①  自分たちの視点・テーマを持って参加すること",
      "②  てんびんとの定期連携（年数回の進捗共有・現場へのアクセス）に参加できること",
      "③  WPC2029での発表に向けて、自律的に活動を進められること",
    ];
    asks.forEach((ask, i) => {
      slide.addText(ask, {
        x: L.contentX + 0.1, y: L.contentY + 2.02 + i * 0.52, w: L.contentW - 0.2, h: 0.45,
        ...font.body(), fontSize: 12, align: "left", valign: "top"
      });
    });

    addPageNum(slide, 6);
  }

  // ===== スライド7: 参加の価値 =====
  {
    const slide = pres.addSlide();
    slide.background = { color: C.bgWhite };

    slide.addText("参加校・参加学生には国際発表から実践経験まで、複数の具体的な価値がある", {
      x: L.titleX, y: L.titleY, w: L.titleW, h: L.titleH,
      ...font.actionTitle(), align: "left", valign: "top"
    });
    addDivider(slide, pres);

    const values = [
      { num: "①", color: C.royalBlue, textColor: C.textWhite, bg: C.blue5,
        title: "国際学会（WPC2029）での発表実績",
        desc: "世界パーキンソン病学会という国際舞台に、学生として立てる" },
      { num: "②", color: C.blue2, textColor: C.textWhite, bg: C.bgLight,
        title: "異分野・異校種との本物の連携体験",
        desc: "NPO・医療者・他校学生と共に動く実践的な協働経験" },
      { num: "③", color: C.blue4, textColor: C.bgNavy, bg: C.blue5,
        title: "高校生：総合型選抜でのアピール",
        desc: "社会課題への実践参加・国際学会発表は強力な活動実績になる" },
      { num: "④", color: C.blue4, textColor: C.bgNavy, bg: C.bgLight,
        title: "大学生：研究・就職活動の実績",
        desc: "国際学会発表・社会実装経験は研究者・社会人として強みになる" },
    ];

    values.forEach((v, i) => {
      const col = i % 2;
      const row = Math.floor(i / 2);
      const vx = col === 0 ? L.contentX : L.col2RightX;
      const vy = L.contentY + row * 1.9;
      const bh = 1.72;

      slide.addShape(pres.shapes.RECTANGLE, {
        x: vx, y: vy, w: L.col2W, h: 0.38,
        fill: { color: v.color }, line: { color: v.color, width: 0 }
      });
      slide.addText(v.num + "  " + v.title, {
        x: vx + 0.12, y: vy, w: L.col2W - 0.24, h: 0.38,
        fontFace: "Yu Gothic", fontSize: 11, bold: true,
        color: v.textColor, align: "left", valign: "middle", margin: 0
      });
      slide.addShape(pres.shapes.RECTANGLE, {
        x: vx, y: vy + 0.38, w: L.col2W, h: bh - 0.38,
        fill: { color: v.bg }, line: { color: C.gridLine, width: 0.5 }
      });
      slide.addText(v.desc, {
        x: vx + 0.15, y: vy + 0.55, w: L.col2W - 0.3, h: bh - 0.7,
        ...font.body(), fontSize: 11, align: "left", valign: "top"
      });
    });

    slide.addShape(pres.shapes.RECTANGLE, {
      x: L.contentX, y: L.contentY + 3.88, w: L.contentW, h: 0.42,
      fill: { color: C.bgNavy }, line: { color: C.bgNavy, width: 0 }
    });
    slide.addText("⑤  「本物の社会課題」への実践参加——授業・演習を超えた、当事者と共に動く経験", {
      x: L.contentX + 0.2, y: L.contentY + 3.88, w: L.contentW - 0.4, h: 0.42,
      fontFace: "Yu Gothic", fontSize: 11, bold: true,
      color: C.textWhite, align: "left", valign: "middle", margin: 0
    });

    addPageNum(slide, 7);
  }

  // ===== スライド8: スケジュール =====
  {
    const slide = pres.addSlide();
    slide.background = { color: C.bgWhite };

    slide.addText("2026年の参加校募集から始まり、WPC2029での発表まで3年間の共同プロジェクトとして動く", {
      x: L.titleX, y: L.titleY, w: L.titleW, h: L.titleH,
      ...font.actionTitle(), align: "left", valign: "top"
    });
    addDivider(slide, pres);

    const steps = [
      { year: "2026", phase: "始動", color: C.blue4, textColor: C.bgNavy,
        items: ["参加校・参加学生の募集", "キックオフミーティング", "てんびん活動への参加"] },
      { year: "2027", phase: "活動", color: C.blue3, textColor: C.bgNavy,
        items: ["各校テーマ設定・調査", "当事者・医療者との交流", "中間連携ミーティング"] },
      { year: "2028", phase: "制作", color: C.blue2, textColor: C.textWhite,
        items: ["発表コンテンツ制作", "中間発表・フィードバック", "WPC2029 発表準備"] },
      { year: "2029", phase: "発表", color: C.royalBlue, textColor: C.textWhite,
        items: ["WPC2029にて世界発信", "国内外への成果共有", "次期プロジェクトへ継続"] },
    ];

    const sw = 2.15;
    const sg = 0.1;

    steps.forEach((step, i) => {
      const sx = L.contentX + i * (sw + sg);
      const sy = L.contentY;

      slide.addShape(pres.shapes.RECTANGLE, {
        x: sx, y: sy, w: sw, h: 0.62,
        fill: { color: step.color }, line: { color: step.color, width: 0 }
      });
      slide.addText(step.year, {
        x: sx, y: sy + 0.02, w: sw, h: 0.36,
        fontFace: "Georgia", fontSize: 20, bold: true,
        color: step.textColor, align: "center", valign: "middle", margin: 0
      });
      slide.addText(step.phase, {
        x: sx, y: sy + 0.38, w: sw, h: 0.24,
        fontFace: "Yu Gothic", fontSize: 9, bold: true,
        color: step.textColor, align: "center", valign: "middle", margin: 0
      });

      slide.addShape(pres.shapes.RECTANGLE, {
        x: sx, y: sy + 0.62, w: sw, h: 2.78,
        fill: { color: i === 3 ? C.blue5 : C.bgLight },
        line: { color: C.gridLine, width: 0.5 }
      });

      step.items.forEach((item, j) => {
        slide.addText("● " + item, {
          x: sx + 0.12, y: sy + 0.78 + j * 0.72, w: sw - 0.24, h: 0.65,
          ...font.body(), fontSize: 10, align: "left", valign: "top"
        });
      });
    });

    slide.addShape(pres.shapes.RECTANGLE, {
      x: L.contentX, y: 4.42, w: L.contentW, h: 0.48,
      fill: { color: C.bgNavy }, line: { color: C.bgNavy, width: 0 }
    });
    slide.addText("最終ゴール: WPC2029（世界パーキンソン病学会）での学生連合による国際発表", {
      x: L.contentX + 0.2, y: 4.42, w: L.contentW - 0.4, h: 0.48,
      fontFace: "Yu Gothic", fontSize: 12, bold: true,
      color: C.textWhite, align: "center", valign: "middle", margin: 0
    });

    addPageNum(slide, 8);
  }

  // ===== スライド9: 参加条件・問い合わせ（ダーク） =====
  {
    const slide = pres.addSlide();
    slide.background = { color: C.bgNavy };

    slide.addText("一緒に動ける仲間を、待っています", {
      x: 0.8, y: 0.85, w: 8.4, h: 0.75,
      fontFace: "Yu Mincho", fontSize: 26, bold: true,
      color: C.textWhite, align: "center", valign: "middle", margin: 0
    });

    slide.addShape(pres.shapes.LINE, {
      x: 2.5, y: 1.72, w: 5.0, h: 0,
      line: { color: C.blue3, width: 1.0 }
    });

    // 左: 参加の条件
    slide.addText("参加の条件", {
      x: 0.8, y: 1.92, w: 4.1, h: 0.38,
      fontFace: "Yu Mincho", fontSize: 14, bold: true,
      color: C.blue3, align: "left", valign: "middle", margin: 0
    });

    const conditions = [
      "自分たちの視点・テーマを持っていること",
      "主体的に動けること（お客様参加はお断り）",
      "定期連携（年数回）に参加できること",
    ];
    conditions.forEach((cond, i) => {
      slide.addText("●  " + cond, {
        x: 0.8, y: 2.42 + i * 0.48, w: 4.0, h: 0.42,
        fontFace: "Yu Gothic", fontSize: 11, color: C.textWhite,
        align: "left", valign: "top", margin: 0
      });
    });

    // 縦線
    slide.addShape(pres.shapes.LINE, {
      x: 5.05, y: 1.92, w: 0, h: 1.9,
      line: { color: C.blue3, width: 0.5 }
    });

    // 右: 問い合わせ
    slide.addText("問い合わせ先", {
      x: 5.3, y: 1.92, w: 4.0, h: 0.38,
      fontFace: "Yu Mincho", fontSize: 14, bold: true,
      color: C.blue3, align: "left", valign: "middle", margin: 0
    });
    slide.addText("NPO法人てんびん", {
      x: 5.3, y: 2.42, w: 4.0, h: 0.38,
      fontFace: "Yu Gothic", fontSize: 13, bold: true,
      color: C.textWhite, align: "left", valign: "top", margin: 0
    });
    slide.addText("Web: 10bin.jp", {
      x: 5.3, y: 2.85, w: 4.0, h: 0.35,
      fontFace: "Calibri", fontSize: 12, color: C.blue3,
      align: "left", valign: "top", margin: 0
    });

    // 締めメッセージ
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 1.0, y: 4.05, w: 8.0, h: 0.78,
      fill: { color: C.royalBlue }, line: { color: C.royalBlue, width: 0 }
    });
    slide.addText("「パーキンソン病を取り巻く世界を良くしたい」\nその思いがあれば、分野も規模も問いません。", {
      x: 1.0, y: 4.05, w: 8.0, h: 0.78,
      fontFace: "Yu Gothic", fontSize: 12, color: C.textWhite,
      align: "center", valign: "middle", margin: 0
    });
  }

  const outputPath = "C:/Users/KATOTO/Downloads/資料作成/shiryo/10bin/20260618_10bin_WPC2029学生企画書_v1.pptx";
  await pres.writeFile({ fileName: outputPath });
  console.log("保存完了: " + outputPath);
}

buildDeck().catch(console.error);

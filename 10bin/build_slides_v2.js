const pptxgen = require("pptxgenjs");

const C = {
  bgWhite:   "FFFFFF",
  bgLight:   "F5F7FA",
  bgNavy:    "0B1F3F",
  textDark:  "333333",
  textMuted: "888888",
  textWhite: "FFFFFF",
  royalBlue: "1B3A6B",
  blue2:     "3D6098",
  blue3:     "7A9CC6",
  blue4:     "B8CCE4",
  blue5:     "D9E5F2",
  gridLine:  "D0D0D0",
};

const L = {
  mx: 0.5, contentX: 0.5, contentY: 1.2,
  contentW: 9.0, titleX: 0.5, titleY: 0.42,
  titleW: 9.0, titleH: 0.55, dividerY: 1.02,
  sourceX: 0.5, sourceY: 5.15,
  pageNumX: 9.0, pageNumY: 5.15,
  col2W: 4.3, col2Gap: 0.4, col2RightX: 5.2,
  col3W: 2.73, col3Gap: 0.4,
};

const font = {
  actionTitle: () => ({ fontFace: "Yu Mincho", fontSize: 17, bold: true, color: C.bgNavy, margin: 0 }),
  sectionHeader: () => ({ fontFace: "Yu Mincho", fontSize: 13, bold: true, color: C.bgNavy, margin: 0 }),
  body: () => ({ fontFace: "Yu Gothic", fontSize: 11, color: C.textDark, margin: 0 }),
  source: () => ({ fontFace: "Yu Gothic", fontSize: 8, color: C.textMuted, margin: 0 }),
  pageNum: () => ({ fontFace: "Calibri", fontSize: 8, color: C.textMuted, margin: 0 }),
};

function divider(s, pres) {
  s.addShape(pres.shapes.LINE, {
    x: L.mx, y: L.dividerY, w: L.contentW, h: 0,
    line: { color: C.royalBlue, width: 0.5 }
  });
}

function pgNum(s, n) {
  s.addText(String(n), {
    x: L.pageNumX, y: L.pageNumY, w: 0.5, h: 0.3,
    ...font.pageNum(), align: "right"
  });
}

function rect(s, pres, x, y, w, h, fill, lineColor) {
  s.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: fill },
    line: { color: lineColor || fill, width: lineColor ? 0.5 : 0 }
  });
}

async function buildDeck() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = "WPC2029 学生アドボカシープロジェクト 参加チーム募集";

  // ===== スライド1: タイトル =====
  {
    const s = pres.addSlide();
    s.background = { color: C.bgNavy };
    s.addText("WPC2029\n学生アドボカシープロジェクト", {
      x: 0.8, y: 1.0, w: 8.4, h: 1.9,
      fontFace: "Yu Mincho", fontSize: 34, bold: true,
      color: C.textWhite, align: "left", valign: "middle", margin: 0
    });
    s.addText("パーキンソン病の課題に、自分たちの専門性で向き合う仲間を探しています", {
      x: 0.8, y: 3.1, w: 8.4, h: 0.5,
      fontFace: "Yu Gothic", fontSize: 13, color: C.blue3,
      align: "left", valign: "top", margin: 0
    });
    s.addText("参加チーム募集　|　NPO法人てんびん　|　2026年", {
      x: 0.8, y: 4.65, w: 8.4, h: 0.3,
      fontFace: "Calibri", fontSize: 10, color: C.blue4,
      align: "left", margin: 0
    });
  }

  // ===== スライド2: WPC2026の5大課題 =====
  {
    const s = pres.addSlide();
    s.background = { color: C.bgWhite };
    s.addText("WPC2026リーダーズ会議が示した世界のパーキンソン病アドボカシー5大課題", {
      x: L.titleX, y: L.titleY, w: L.titleW, h: L.titleH,
      ...font.actionTitle(), align: "left", valign: "top"
    });
    divider(s, pres);

    const issues = [
      { n: "①", t: "スティグマ・社会的孤立",  d: "偏見が当事者を社会から遠ざける。年齢を問わず包摂が必要。" },
      { n: "②", t: "移動・公共空間の障壁",     d: "信号・ベンチ・交通機関が自立した生活を制限する。" },
      { n: "③", t: "診断時支援の不足",         d: "診断直後に必要な情報・つながりが届いていない。" },
      { n: "④", t: "医療アクセスの地域格差",   d: "専門医・リハビリへのアクセスに国・地域間で大きな差がある。" },
      { n: "⑤", t: "国際連携体の必要性",       d: "各国単独では声が届かない。連携で政府・WHOへの働きかけを強化。" },
    ];
    issues.forEach((iss, i) => {
      const row = Math.floor(i / 3), col = i % 3;
      const bw = row === 0 ? 2.85 : L.col2W;
      const gap = row === 0 ? 0.225 : L.col2Gap;
      const startX = row === 0 ? L.contentX : L.contentX + (L.contentW - (L.col2W * 2 + L.col2Gap)) / 2;
      const bx = startX + col * (bw + gap);
      const by = L.contentY + row * 1.9;
      rect(s, pres, bx, by, bw, 1.72, i % 2 === 0 ? C.blue5 : C.bgLight, C.gridLine);
      s.addText(iss.n, { x: bx+0.12, y: by+0.1, w: 0.42, h: 0.42, fontFace: "Georgia", fontSize: 16, bold: true, color: C.royalBlue, align: "center", valign: "middle", margin: 0 });
      s.addText(iss.t, { x: bx+0.12, y: by+0.56, w: bw-0.24, h: 0.38, fontFace: "Yu Mincho", fontSize: 11, bold: true, color: C.bgNavy, align: "left", valign: "top", margin: 0 });
      s.addText(iss.d, { x: bx+0.12, y: by+0.98, w: bw-0.24, h: 0.62, ...font.body(), fontSize: 10, align: "left", valign: "top" });
    });
    s.addText("出所: WPC2026 Leaders Meeting 議事録", { x: L.sourceX, y: L.sourceY, w: 6, h: 0.3, ...font.source(), align: "left" });
    pgNum(s, 2);
  }

  // ===== スライド3: てんびんとは / 灘高校実績 =====
  {
    const s = pres.addSlide();
    s.background = { color: C.bgWhite };
    s.addText("てんびんは、学生と共に社会課題に向き合い続けてきた実績を持つNPOだ", {
      x: L.titleX, y: L.titleY, w: L.titleW, h: L.titleH,
      ...font.actionTitle(), align: "left", valign: "top"
    });
    divider(s, pres);

    rect(s, pres, L.contentX, L.contentY, L.col2W, 0.38, C.royalBlue);
    s.addText("NPO法人てんびんとは", { x: L.contentX, y: L.contentY, w: L.col2W, h: 0.38, fontFace: "Yu Mincho", fontSize: 12, bold: true, color: C.textWhite, align: "center", valign: "middle", margin: 0 });
    rect(s, pres, L.contentX, L.contentY+0.38, L.col2W, 2.9, C.bgLight, C.gridLine);
    ["パーキンソン病の当事者・家族を支援するNPO",
     "PDヨガ・アートセラピー・無料相談会を主催",
     "「誰もが彩り豊かに生き合う未来」を理念に活動",
     "医療・文化・教育をつなぐハブとして機能",
    ].forEach((txt, i) => {
      s.addText("● " + txt, { x: L.contentX+0.15, y: L.contentY+0.52+i*0.55, w: L.col2W-0.3, h: 0.48, ...font.body(), fontSize: 11, align: "left", valign: "top" });
    });

    rect(s, pres, L.col2RightX, L.contentY, L.col2W, 0.38, C.blue2);
    s.addText("灘高校との連携実績（2026年4月）", { x: L.col2RightX, y: L.contentY, w: L.col2W, h: 0.38, fontFace: "Yu Mincho", fontSize: 12, bold: true, color: C.textWhite, align: "center", valign: "middle", margin: 0 });
    rect(s, pres, L.col2RightX, L.contentY+0.38, L.col2W, 2.9, C.blue5, C.gridLine);
    s.addText("2026年4月12日　兵庫県立美術館", { x: L.col2RightX+0.15, y: L.contentY+0.52, w: L.col2W-0.3, h: 0.32, fontFace: "Yu Gothic", fontSize: 11, bold: true, color: C.royalBlue, align: "left", valign: "top", margin: 0 });
    s.addText("「学生たちの発表会〜僕たち私たちがあなたのために出来ること〜」", { x: L.col2RightX+0.15, y: L.contentY+0.88, w: L.col2W-0.3, h: 0.5, fontFace: "Yu Gothic", fontSize: 10, bold: true, color: C.bgNavy, align: "left", valign: "top", margin: 0 });
    s.addText("灘中学生が「パーキンソン病を取り巻く世界を良くするには」をテーマに医療従事者・行政・当事者団体へ社会提言を実施。", { x: L.col2RightX+0.15, y: L.contentY+1.44, w: L.col2W-0.3, h: 1.7, ...font.body(), fontSize: 10, align: "left", valign: "top" });
    pgNum(s, 3);
  }

  // ===== スライド4: 半年サイクル制のコンセプト =====
  {
    const s = pres.addSlide();
    s.background = { color: C.bgWhite };
    s.addText("「人が続く」のではなく「テーマが育つ」——半年サイクル制で学生の主体性と継続性を両立する", {
      x: L.titleX, y: L.titleY, w: L.titleW, h: L.titleH,
      ...font.actionTitle(), align: "left", valign: "top"
    });
    divider(s, pres);

    // サイクル図：横並び4ブロック + 矢印
    const blocks = [
      { label: "第1期\n（半年）", color: C.blue4, textC: C.bgNavy },
      { label: "第2期\n（半年）", color: C.blue3, textC: C.bgNavy },
      { label: "第3期\n（半年）", color: C.blue2, textC: C.textWhite },
      { label: "WPC2029\n（てんびん）", color: C.royalBlue, textC: C.textWhite },
    ];

    const bw = 1.8, bh = 0.9, by = L.contentY + 0.1, gap = 0.38;
    const totalW = bw * 4 + gap * 3;
    const startX = L.contentX + (L.contentW - totalW) / 2;

    blocks.forEach((b, i) => {
      const bx = startX + i * (bw + gap);
      rect(s, pres, bx, by, bw, bh, b.color);
      s.addText(b.label, { x: bx, y: by, w: bw, h: bh, fontFace: "Yu Mincho", fontSize: 11, bold: true, color: b.textC, align: "center", valign: "middle", margin: 0 });
      if (i < 3) {
        s.addText("→", { x: bx + bw, y: by + 0.2, w: gap, h: 0.5, fontFace: "Calibri", fontSize: 18, bold: true, color: C.royalBlue, align: "center", valign: "middle", margin: 0 });
      }
    });

    // 「国内発表」ラベル
    for (let i = 0; i < 3; i++) {
      const bx = startX + i * (bw + gap);
      s.addText("↓ 国内発表", { x: bx, y: by + bh + 0.05, w: bw, h: 0.32, fontFace: "Yu Gothic", fontSize: 10, bold: true, color: C.blue2, align: "center", valign: "top", margin: 0 });
    }
    s.addText("↓ 世界へ発信", { x: startX + 3 * (bw + gap), y: by + bh + 0.05, w: bw, h: 0.32, fontFace: "Yu Gothic", fontSize: 10, bold: true, color: C.royalBlue, align: "center", valign: "top", margin: 0 });

    // 3つの特徴
    const feats = [
      { title: "学生にとって",   body: "半年間コミット → 発表 → 完結\n（次期参加も任意）" },
      { title: "テーマにとって", body: "前期の成果が次期の\nスタートラインになる" },
      { title: "てんびんとして", body: "各期の蓄積を整理し\nWPC2029でまとめて発信" },
    ];
    feats.forEach((f, i) => {
      const cx = L.contentX + i * (L.col3W + L.col3Gap);
      const cy = L.contentY + 1.55;
      rect(s, pres, cx, cy, L.col3W, 0.38, i === 2 ? C.royalBlue : C.blue2);
      s.addText(f.title, { x: cx, y: cy, w: L.col3W, h: 0.38, fontFace: "Yu Mincho", fontSize: 12, bold: true, color: C.textWhite, align: "center", valign: "middle", margin: 0 });
      rect(s, pres, cx, cy + 0.38, L.col3W, 1.0, C.bgLight, C.gridLine);
      s.addText(f.body, { x: cx + 0.12, y: cy + 0.5, w: L.col3W - 0.24, h: 0.8, ...font.body(), fontSize: 11, align: "left", valign: "top" });
    });

    pgNum(s, 4);
  }

  // ===== スライド5: 参加チームに期待すること（3条件）=====
  {
    const s = pres.addSlide();
    s.background = { color: C.bgWhite };
    s.addText("参加チームには3つの条件だけを求める——形式・分野・規模は一切問わない", {
      x: L.titleX, y: L.titleY, w: L.titleW, h: L.titleH,
      ...font.actionTitle(), align: "left", valign: "top"
    });
    divider(s, pres);

    s.addText("発表形式の例（これに限らない）", { x: L.contentX, y: L.contentY, w: L.contentW, h: 0.32, ...font.sectionHeader(), fontSize: 11, align: "left", valign: "top" });
    const fmts = ["映像・ドキュメンタリー", "アート・デザイン作品", "学術研究・論文発表",
                  "政策提言・ロビー活動", "ワークショップ企画", "アニメーション・SNS発信"];
    fmts.forEach((fmt, i) => {
      const col = i % 3, row = Math.floor(i / 3);
      const fx = L.contentX + col * 3.05, fy = L.contentY + 0.4 + row * 0.55;
      rect(s, pres, fx, fy, 2.85, 0.42, C.blue5, C.gridLine);
      s.addText(fmt, { x: fx+0.1, y: fy, w: 2.65, h: 0.42, fontFace: "Yu Gothic", fontSize: 11, bold: true, color: C.royalBlue, align: "center", valign: "middle", margin: 0 });
    });

    s.addText("参加の3条件", { x: L.contentX, y: L.contentY+1.62, w: L.contentW, h: 0.32, ...font.sectionHeader(), fontSize: 11, align: "left", valign: "top" });
    const conds = [
      { n: "①", t: "自分たちの視点・テーマを持っていること",              c: C.royalBlue },
      { n: "②", t: "主体的に動けること\n（「参加させてもらう」スタンスは不可）", c: C.blue2 },
      { n: "③", t: "月1〜2回程度の定期連携に参加できること",              c: C.blue3 },
    ];
    conds.forEach((cond, i) => {
      const cx = L.contentX + i * 3.05, cy = L.contentY + 2.02;
      rect(s, pres, cx, cy, 2.85, 0.38, cond.c);
      s.addText(cond.n, { x: cx+0.1, y: cy, w: 0.35, h: 0.38, fontFace: "Georgia", fontSize: 14, bold: true, color: C.textWhite, align: "center", valign: "middle", margin: 0 });
      rect(s, pres, cx, cy+0.38, 2.85, 0.98, C.bgLight, C.gridLine);
      s.addText(cond.t, { x: cx+0.12, y: cy+0.48, w: 2.61, h: 0.8, ...font.body(), fontSize: 10, align: "left", valign: "top" });
    });

    s.addText("条件を満たさない場合はお断りします。前向きな仲間だけを迎えるために、入口で正直に伝えます。", {
      x: L.contentX, y: L.contentY + 3.5, w: L.contentW, h: 0.35,
      fontFace: "Yu Gothic", fontSize: 10, color: C.textMuted,
      align: "left", valign: "middle", margin: 0
    });
    pgNum(s, 5);
  }

  // ===== スライド6: ステークホルダーの意義 =====
  {
    const s = pres.addSlide();
    s.background = { color: C.bgWhite };
    s.addText("この企画は全員にとって意義がある——関わるすべてのステークホルダーに価値を設計した", {
      x: L.titleX, y: L.titleY, w: L.titleW, h: L.titleH,
      ...font.actionTitle(), align: "left", valign: "top"
    });
    divider(s, pres);

    const COL1 = 2.6, COL2 = L.contentW - COL1 - 0.1;
    const rows = [
      { who: "NPO法人てんびん",         val: "若い世代との継続的な連携・テーマの蓄積・WPC2029での国際発信力強化" },
      { who: "参加チームの学生（共通）", val: "半年間の実践経験・国内発表実績・本物の社会課題への参加・異分野連携体験" },
      { who: "高校生",                  val: "総合型選抜での活動実績・社会へ提言する体験（半年完結で参加しやすい）" },
      { who: "大学生",                  val: "研究・就活における実績・社会実装の経験（ゼミ・研究室単位での参加も可）" },
      { who: "パーキンソン病当事者・家族", val: "多様な視点からの継続的なアドボカシー・次世代との継続的な対話" },
    ];

    rect(s, pres, L.contentX, L.contentY, COL1, 0.38, C.royalBlue);
    s.addText("ステークホルダー", { x: L.contentX, y: L.contentY, w: COL1, h: 0.38, fontFace: "Yu Gothic", fontSize: 11, bold: true, color: C.textWhite, align: "center", valign: "middle", margin: 0 });
    rect(s, pres, L.contentX+COL1+0.1, L.contentY, COL2, 0.38, C.royalBlue);
    s.addText("この企画から得られる意義", { x: L.contentX+COL1+0.1, y: L.contentY, w: COL2, h: 0.38, fontFace: "Yu Gothic", fontSize: 11, bold: true, color: C.textWhite, align: "center", valign: "middle", margin: 0 });

    rows.forEach((row, i) => {
      const ry = L.contentY + 0.38 + i * 0.62;
      const fill = i % 2 === 0 ? C.bgLight : C.blue5;
      rect(s, pres, L.contentX, ry, COL1, 0.58, fill, C.gridLine);
      s.addText(row.who, { x: L.contentX+0.1, y: ry, w: COL1-0.2, h: 0.58, fontFace: "Yu Gothic", fontSize: 10, bold: true, color: C.bgNavy, align: "left", valign: "middle", margin: 0 });
      rect(s, pres, L.contentX+COL1+0.1, ry, COL2, 0.58, fill, C.gridLine);
      s.addText(row.val, { x: L.contentX+COL1+0.2, y: ry, w: COL2-0.2, h: 0.58, ...font.body(), fontSize: 10, align: "left", valign: "middle" });
    });
    pgNum(s, 6);
  }

  // ===== スライド7: テーマ継承のイメージ =====
  {
    const s = pres.addSlide();
    s.background = { color: C.bgWhite };
    s.addText("「人」ではなく「テーマ」が育つ——各期の成果が次の探究のスタートラインになる", {
      x: L.titleX, y: L.titleY, w: L.titleW, h: L.titleH,
      ...font.actionTitle(), align: "left", valign: "top"
    });
    divider(s, pres);

    const stages = [
      { label: "第1期", sub: "探究・提言",   color: C.blue4, textC: C.bgNavy,
        body: "「スティグマと\n学生にできること」\n調査・提言発表（例）" },
      { label: "第2期", sub: "掘り下げ",     color: C.blue3, textC: C.bgNavy,
        body: "前期の提言を受け\n映像表現に発展\nアニメーション制作（例）" },
      { label: "第3期", sub: "実践・展開",   color: C.blue2, textC: C.textWhite,
        body: "映像を使った\n出前授業を設計・実施\n啓発プログラムへ（例）" },
      { label: "WPC2029", sub: "てんびんが統合", color: C.royalBlue, textC: C.textWhite,
        body: "3期分の成果を\nてんびんが整理し\n国際学会で発表" },
    ];

    const bw = 1.95, bh = 2.8, by = L.contentY, gap = 0.15;
    stages.forEach((st, i) => {
      const bx = L.contentX + i * (bw + gap);
      rect(s, pres, bx, by, bw, 0.55, st.color);
      s.addText(st.label, { x: bx, y: by+0.02, w: bw, h: 0.3, fontFace: "Georgia", fontSize: 14, bold: true, color: st.textC, align: "center", valign: "middle", margin: 0 });
      s.addText(st.sub, { x: bx, y: by+0.32, w: bw, h: 0.22, fontFace: "Yu Gothic", fontSize: 9, bold: true, color: st.textC, align: "center", valign: "middle", margin: 0 });
      rect(s, pres, bx, by+0.55, bw, bh-0.55, i === 3 ? C.blue5 : C.bgLight, C.gridLine);
      s.addText(st.body, { x: bx+0.12, y: by+0.7, w: bw-0.24, h: bh-0.85, ...font.body(), fontSize: 10, align: "left", valign: "top" });

      if (i < 3) {
        s.addText("→", { x: bx+bw, y: by+0.9, w: gap+0.05, h: 0.5, fontFace: "Calibri", fontSize: 16, bold: true, color: C.royalBlue, align: "center", valign: "middle", margin: 0 });
        s.addText("継承", { x: bx+bw, y: by+1.45, w: gap+0.05, h: 0.3, fontFace: "Yu Gothic", fontSize: 7, color: C.textMuted, align: "center", valign: "top", margin: 0 });
      }
    });

    s.addText("※ テーマは次期チームが自由に設定します。掘り下げるか、全く新しい角度で入り直しても構いません。", {
      x: L.contentX, y: L.contentY + 3.0, w: L.contentW, h: 0.35,
      fontFace: "Yu Gothic", fontSize: 9, color: C.textMuted,
      align: "left", valign: "middle", margin: 0
    });
    pgNum(s, 7);
  }

  // ===== スライド8: スケジュール =====
  {
    const s = pres.addSlide();
    s.background = { color: C.bgWhite };
    s.addText("半年サイクルを積み重ね、WPC2029にてんびんが蓄積を持ち寄る", {
      x: L.titleX, y: L.titleY, w: L.titleW, h: L.titleH,
      ...font.actionTitle(), align: "left", valign: "top"
    });
    divider(s, pres);

    // 横軸タイムライン：2026〜2029
    const years = ["2026年", "2027年", "2028年", "2029年"];
    const yw = 2.15, yg = 0.1, yy = L.contentY;
    years.forEach((yr, i) => {
      const yx = L.contentX + i * (yw + yg);
      const isLast = i === 3;
      rect(s, pres, yx, yy, yw, 0.45, isLast ? C.royalBlue : C.blue3);
      s.addText(yr, { x: yx, y: yy, w: yw, h: 0.45, fontFace: "Georgia", fontSize: 16, bold: true, color: C.textWhite, align: "center", valign: "middle", margin: 0 });
    });

    // 各年の内容
    const yearContent = [
      { items: ["第1期スタート（前半）", "第1期 国内発表（後半）", "第2期スタート"] },
      { items: ["第2期 国内発表", "第3期スタート", "第3期 国内発表"] },
      { items: ["第4期スタート", "第4期 国内発表", "WPC2029 発表準備"] },
      { items: ["WPC2029 開催", "てんびんが全期\nの成果を統合発表", ""] },
    ];
    yearContent.forEach((yc, i) => {
      const yx = L.contentX + i * (yw + yg);
      rect(s, pres, yx, yy+0.45, yw, 2.95, i === 3 ? C.blue5 : C.bgLight, C.gridLine);
      yc.items.forEach((item, j) => {
        if (!item) return;
        s.addText("● " + item, { x: yx+0.12, y: yy+0.6+j*0.72, w: yw-0.24, h: 0.65, ...font.body(), fontSize: 9, align: "left", valign: "top" });
      });
    });

    rect(s, pres, L.contentX, 4.42, L.contentW, 0.48, C.bgNavy);
    s.addText("てんびんが複数期の蓄積を整理・統合し、WPC2029（世界パーキンソン病学会）で発信する", {
      x: L.contentX+0.2, y: 4.42, w: L.contentW-0.4, h: 0.48,
      fontFace: "Yu Gothic", fontSize: 12, bold: true,
      color: C.textWhite, align: "center", valign: "middle", margin: 0
    });
    pgNum(s, 8);
  }

  // ===== スライド9: 参加条件・問い合わせ =====
  {
    const s = pres.addSlide();
    s.background = { color: C.bgNavy };
    s.addText("一緒に動ける仲間を、待っています", {
      x: 0.8, y: 0.85, w: 8.4, h: 0.75,
      fontFace: "Yu Mincho", fontSize: 26, bold: true,
      color: C.textWhite, align: "center", valign: "middle", margin: 0
    });
    s.addShape(pres.shapes.LINE, { x: 2.5, y: 1.72, w: 5.0, h: 0, line: { color: C.blue3, width: 1.0 } });

    s.addText("参加の3条件", { x: 0.8, y: 1.92, w: 4.1, h: 0.38, fontFace: "Yu Mincho", fontSize: 14, bold: true, color: C.blue3, align: "left", valign: "middle", margin: 0 });
    ["① 自分たちの視点・テーマを持っていること",
     "② 主体的に動けること（お客様参加はお断り）",
     "③ 月1〜2回の定期連携に参加できること",
    ].forEach((c, i) => {
      s.addText(c, { x: 0.8, y: 2.42+i*0.48, w: 4.1, h: 0.42, fontFace: "Yu Gothic", fontSize: 11, color: C.textWhite, align: "left", valign: "top", margin: 0 });
    });

    s.addShape(pres.shapes.LINE, { x: 5.05, y: 1.92, w: 0, h: 1.9, line: { color: C.blue3, width: 0.5 } });
    s.addText("問い合わせ先", { x: 5.3, y: 1.92, w: 4.0, h: 0.38, fontFace: "Yu Mincho", fontSize: 14, bold: true, color: C.blue3, align: "left", valign: "middle", margin: 0 });
    s.addText("NPO法人てんびん", { x: 5.3, y: 2.42, w: 4.0, h: 0.38, fontFace: "Yu Gothic", fontSize: 13, bold: true, color: C.textWhite, align: "left", valign: "top", margin: 0 });
    s.addText("Web: 10bin.jp", { x: 5.3, y: 2.85, w: 4.0, h: 0.35, fontFace: "Calibri", fontSize: 12, color: C.blue3, align: "left", valign: "top", margin: 0 });

    rect(s, pres, 1.0, 4.05, 8.0, 0.78, C.royalBlue);
    s.addText("「パーキンソン病を取り巻く世界を良くしたい」\nその思いがあれば、分野も規模も問いません。半年間、一緒に動きましょう。", {
      x: 1.0, y: 4.05, w: 8.0, h: 0.78,
      fontFace: "Yu Gothic", fontSize: 11, color: C.textWhite,
      align: "center", valign: "middle", margin: 0
    });
  }

  const out = "C:/Users/KATOTO/Downloads/資料作成/shiryo/10bin/20260618_10bin_WPC2029学生企画_発表資料_v2.pptx";
  await pres.writeFile({ fileName: out });
  console.log("保存完了: " + out);
}

buildDeck().catch(console.error);

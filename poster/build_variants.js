// 교원업적평가 접수 안내 포스터 시안 3종 (A3 세로, 297 x 420 mm)
//   A: 수정본(좌·우 방향 화살표) 기반 디자인 보강
//   B: 길 안내형 — 하단을 좌(연구)/우(봉사) 색면으로 나눔
//   C: 가로 안내 바형 — 진한 안내 카드 + 가로형 접수처 2줄
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa6");

const W = 11.69, H = 16.54;
const FONT = "맑은 고딕";

// 문구는 원본 그대로 사용
const TEXT = {
  head: "교원업적평가",
  l2: ["전산프로그램 출력물", "과"],
  l3: ["증빙서류 확인", " 후"],
  l4: "접수해 드리겠습니다.",
  research: { area: "연구영역", title: "접수 담당자", detail: "이중영, 이어진" },
  service: { area: "봉사영역", title: "기획처", detail: "본부동 911호" },
};

const C = {
  blue: "1E7BE0",
  navy: "1B2559",
  ink: "1A1A1A",
  cyan: "0AA5E8",
  red: "FF5A2C",
  research: "3446C6",
  service: "FF6A2B",
  white: "FFFFFF",
};

async function icon(name, color, px = 512) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(fa[name], { color: "#" + color, size: px })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

function newDeck(title) {
  const pres = new pptxgen();
  pres.defineLayout({ name: "A3_PORTRAIT", width: W, height: H });
  pres.layout = "A3_PORTRAIT";
  pres.title = title;
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  return pres;
}

const shadow = (opacity = 0.18) => ({ type: "outer", color: "000000", opacity, blur: 10, offset: 4, angle: 90 });
const runs = (parts, size, breakLast = true) => parts.map(([text, color], i) => ({
  text, options: { color, fontSize: size, breakLine: breakLast && i === parts.length - 1 },
}));

// ════════════════════════════════════════════════════════════
// 시안 A — 블루 카드 (수정본 기반)
// ════════════════════════════════════════════════════════════
async function variantA(file) {
  const pres = newDeck("교원업적평가 접수 안내 - 시안 A");
  const s = pres.addSlide();
  s.background = { color: C.blue };

  // 배경 장식 원
  s.addShape(pres.ShapeType.ellipse, { objectName: "장식 원 1", x: 7.6, y: -2.6, w: 6.5, h: 6.5, fill: { color: C.white, transparency: 88 }, line: { type: "none" } });
  s.addShape(pres.ShapeType.ellipse, { objectName: "장식 원 2", x: -3.0, y: 12.0, w: 6.5, h: 6.5, fill: { color: C.white, transparency: 90 }, line: { type: "none" } });

  const M = 0.6, cw = W - 2 * M;

  // 상단 안내 카드 + 상단 배지 아이콘
  const topY = 1.15, topH = 6.6;
  s.addShape(pres.ShapeType.roundRect, { objectName: "안내 문구 카드", x: M, y: topY, w: cw, h: topH, rectRadius: 0.35, fill: { color: C.white }, line: { type: "none" }, shadow: shadow() });
  const bd = 1.5;
  s.addShape(pres.ShapeType.ellipse, { objectName: "배지 원", x: W / 2 - bd / 2, y: topY - bd / 2, w: bd, h: bd, fill: { color: C.cyan }, line: { color: C.white, width: 6 }, shadow: shadow(0.2) });
  s.addImage({ data: await icon("FaFileCircleCheck", C.white), x: W / 2 - 0.4, y: topY - 0.4, w: 0.8, h: 0.8, altText: "서류 확인 아이콘" });

  // "교원업적평가" 제목 알약
  const pillW = 6.2, pillH = 1.2, pillY = topY + 1.0;
  s.addText(TEXT.head, {
    objectName: "제목", shape: pres.ShapeType.roundRect, rectRadius: 0.6,
    x: W / 2 - pillW / 2, y: pillY, w: pillW, h: pillH,
    fill: { color: C.navy }, color: C.white, fontFace: FONT, fontSize: 58, bold: true,
    align: "center", valign: "middle", margin: 0, charSpacing: 2,
  });
  s.addText(
    [
      ...runs([[TEXT.l2[0], C.cyan], [TEXT.l2[1], C.ink]], 60),
      ...runs([[TEXT.l3[0], C.red], [TEXT.l3[1], C.ink]], 60),
      { text: TEXT.l4, options: { color: C.ink, fontSize: 60 } },
    ],
    { objectName: "안내 문구", x: M + 0.3, y: pillY + pillH + 0.25, w: cw - 0.6, h: topY + topH - (pillY + pillH + 0.25) - 0.3, fontFace: FONT, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.2, margin: 0, isTextBox: true }
  );

  // 하단 접수처 카드
  const botY = topY + topH + 0.45, botH = H - M - botY;
  s.addShape(pres.ShapeType.roundRect, { objectName: "접수처 카드", x: M, y: botY, w: cw, h: botH, rectRadius: 0.35, fill: { color: C.white }, line: { type: "none" }, shadow: shadow() });

  const pad = 0.4, gap = 0.4, colW = (cw - 2 * pad - gap) / 2;
  const headY = botY + pad, headH = 1.8;
  const resH = 3.3, resY = botY + botH - pad - resH;
  const cd = 1.55, cY = (headY + headH + resY) / 2 - cd / 2;

  const cols = [
    { x: M + pad, color: C.research, t: TEXT.research, arrow: "FaArrowLeft", ic: "FaUserTie", icAlt: "담당자 아이콘" },
    { x: M + pad + colW + gap, color: C.service, t: TEXT.service, arrow: "FaArrowRight", ic: "FaLocationDot", icAlt: "위치 아이콘" },
  ];
  for (const c of cols) {
    s.addText(c.t.area, {
      objectName: `${c.t.area} 제목`, shape: pres.ShapeType.roundRect, rectRadius: 0.25,
      x: c.x, y: headY, w: colW, h: headH, fill: { color: c.color }, color: C.white,
      fontFace: FONT, fontSize: 64, bold: true, align: "center", valign: "middle", margin: 0, shadow: shadow(0.15),
    });
    // 방향 화살표 (수정본의 좌/우 방향 유지)
    const cx = c.x + colW / 2 - cd / 2;
    s.addShape(pres.ShapeType.ellipse, { objectName: `${c.t.area} 화살표 원`, x: cx, y: cY, w: cd, h: cd, fill: { color: c.color }, line: { color: C.white, width: 5 }, shadow: shadow(0.2) });
    s.addImage({ data: await icon(c.arrow, C.white), x: cx + cd / 2 - 0.42, y: cY + cd / 2 - 0.42, w: 0.84, h: 0.84, altText: c.arrow === "FaArrowLeft" ? "왼쪽 화살표" : "오른쪽 화살표" });

    s.addShape(pres.ShapeType.roundRect, { objectName: `${c.t.area} 접수처`, x: c.x, y: resY, w: colW, h: resH, rectRadius: 0.25, fill: { color: c.color }, line: { type: "none" }, shadow: shadow(0.15) });
    s.addImage({ data: await icon(c.ic, C.white), x: c.x + colW / 2 - 0.38, y: resY + 0.35, w: 0.76, h: 0.76, altText: c.icAlt });
    s.addText(
      [
        { text: c.t.title, options: { fontSize: 48, breakLine: true } },
        { text: c.t.detail, options: { fontSize: 38 } },
      ],
      { objectName: `${c.t.area} 접수처 문구`, x: c.x + 0.1, y: resY + 1.2, w: colW - 0.2, h: resH - 1.4, fontFace: FONT, bold: true, color: C.white, align: "center", valign: "middle", lineSpacingMultiple: 1.25, margin: 0, isTextBox: true }
    );
  }
  await pres.writeFile({ fileName: file });
}

// ════════════════════════════════════════════════════════════
// 시안 B — 길 안내형 (좌·우 색면 분할)
// ════════════════════════════════════════════════════════════
async function variantB(file) {
  const pres = newDeck("교원업적평가 접수 안내 - 시안 B");
  const s = pres.addSlide();
  s.background = { color: C.white };

  const splitY = 8.4, half = W / 2, M = 0.6;

  // 상단: 아이콘 + 제목 칩 + 안내 문구
  s.addShape(pres.ShapeType.ellipse, { objectName: "장식 원", x: -2.2, y: -2.4, w: 5.5, h: 5.5, fill: { color: "EAF2FD" }, line: { type: "none" } });
  s.addShape(pres.ShapeType.ellipse, { objectName: "장식 원 2", x: 9.4, y: 5.2, w: 3.6, h: 3.6, fill: { color: "FFF0EA" }, line: { type: "none" } });
  const id = 1.35;
  s.addShape(pres.ShapeType.ellipse, { objectName: "아이콘 원", x: W / 2 - id / 2, y: 0.7, w: id, h: id, fill: { color: C.blue }, line: { type: "none" } });
  s.addImage({ data: await icon("FaClipboardCheck", C.white), x: W / 2 - 0.36, y: 0.7 + id / 2 - 0.4, w: 0.72, h: 0.8, altText: "확인 아이콘" });

  const chipW = 6.0, chipH = 1.15;
  s.addText(TEXT.head, {
    objectName: "제목", shape: pres.ShapeType.roundRect, rectRadius: 0.575,
    x: W / 2 - chipW / 2, y: 2.35, w: chipW, h: chipH,
    fill: { color: C.navy }, color: C.white, fontFace: FONT, fontSize: 56, bold: true,
    align: "center", valign: "middle", margin: 0, charSpacing: 2,
  });
  s.addText(
    [
      ...runs([[TEXT.l2[0], C.blue], [TEXT.l2[1], C.ink]], 62),
      ...runs([[TEXT.l3[0], C.red], [TEXT.l3[1], C.ink]], 62),
      { text: TEXT.l4, options: { color: C.ink, fontSize: 62 } },
    ],
    { objectName: "안내 문구", x: M, y: 3.75, w: W - 2 * M, h: 4.3, fontFace: FONT, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.2, margin: 0, isTextBox: true }
  );

  // 하단: 좌(연구) / 우(봉사) 색면
  const sides = [
    { x: 0, color: C.research, t: TEXT.research, arrow: "FaArrowLeft", ic: "FaUserTie", align: "left" },
    { x: half, color: C.service, t: TEXT.service, arrow: "FaArrowRight", ic: "FaLocationDot", align: "right" },
  ];
  for (const c of sides) {
    s.addShape(pres.ShapeType.rect, { objectName: `${c.t.area} 색면`, x: c.x, y: splitY, w: half, h: H - splitY, fill: { color: c.color }, line: { type: "none" } });
    const innerX = c.x + M, innerW = half - 2 * M;
    const aw = 1.9, ah = 1.6;
    s.addImage({ data: await icon(c.arrow, C.white), x: c.align === "left" ? innerX : innerX + innerW - aw, y: splitY + 0.65, w: aw, h: ah, sizing: { type: "contain", w: aw, h: ah }, altText: c.align === "left" ? "왼쪽 화살표" : "오른쪽 화살표" });
    s.addText(c.t.area, { objectName: `${c.t.area} 제목`, x: innerX, y: splitY + 2.55, w: innerW, h: 1.35, fontFace: FONT, fontSize: 66, bold: true, color: C.white, align: c.align, valign: "middle", margin: 0, isTextBox: true });

    const cardY = splitY + 4.35, cardH = H - cardY - 0.75;
    s.addShape(pres.ShapeType.roundRect, { objectName: `${c.t.area} 접수처 카드`, x: innerX, y: cardY, w: innerW, h: cardH, rectRadius: 0.25, fill: { color: C.white }, line: { type: "none" }, shadow: shadow(0.2) });
    const icS = 0.7;
    s.addImage({ data: await icon(c.ic, c.color), x: c.align === "left" ? innerX + 0.4 : innerX + innerW - 0.4 - icS, y: cardY + 0.4, w: icS, h: icS, altText: c.ic === "FaUserTie" ? "담당자 아이콘" : "위치 아이콘" });
    s.addText(
      [
        { text: c.t.title, options: { fontSize: 44, color: c.color, breakLine: true } },
        { text: c.t.detail, options: { fontSize: 36, color: C.ink } },
      ],
      { objectName: `${c.t.area} 접수처`, x: innerX + 0.4, y: cardY + 1.2, w: innerW - 0.8, h: cardH - 1.45, fontFace: FONT, bold: true, align: c.align, valign: "middle", lineSpacingMultiple: 1.25, margin: 0, isTextBox: true }
    );
  }
  await pres.writeFile({ fileName: file });
}

// ════════════════════════════════════════════════════════════
// 시안 C — 가로 안내 바형
// ════════════════════════════════════════════════════════════
async function variantC(file) {
  const pres = newDeck("교원업적평가 접수 안내 - 시안 C");
  const s = pres.addSlide();
  s.background = { color: "EEF3FB" };

  const M = 0.6, cw = W - 2 * M;

  // 상단 진한 안내 카드
  const topY = M, topH = 7.4;
  s.addShape(pres.ShapeType.roundRect, { objectName: "안내 문구 카드", x: M, y: topY, w: cw, h: topH, rectRadius: 0.4, fill: { color: C.navy }, line: { type: "none" }, shadow: shadow(0.25) });
  s.addShape(pres.ShapeType.ellipse, { objectName: "장식 원 1", x: 7.9, y: -0.9, w: 4.4, h: 4.4, fill: { color: C.white, transparency: 92 }, line: { type: "none" } });
  s.addShape(pres.ShapeType.ellipse, { objectName: "장식 원 2", x: -0.2, y: 5.4, w: 3.0, h: 3.0, fill: { color: C.white, transparency: 94 }, line: { type: "none" } });

  const hY = topY + 0.75;
  s.addImage({ data: await icon("FaFileCircleCheck", "FFD54A"), x: W / 2 - 0.42, y: hY, w: 0.84, h: 0.84, altText: "서류 확인 아이콘" });
  s.addText(TEXT.head, { objectName: "제목", x: M + 0.3, y: hY + 1.0, w: cw - 0.6, h: 1.2, fontFace: FONT, fontSize: 60, bold: true, color: "FFD54A", align: "center", valign: "middle", margin: 0, charSpacing: 2, isTextBox: true });
  s.addText(
    [
      ...runs([[TEXT.l2[0], "5CCBFF"], [TEXT.l2[1], C.white]], 60),
      ...runs([[TEXT.l3[0], "FF8A65"], [TEXT.l3[1], C.white]], 60),
      { text: TEXT.l4, options: { color: C.white, fontSize: 60 } },
    ],
    { objectName: "안내 문구", x: M + 0.3, y: hY + 2.35, w: cw - 0.6, h: topY + topH - (hY + 2.35) - 0.45, fontFace: FONT, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.2, margin: 0, isTextBox: true }
  );

  // 하단 가로 안내 바 2줄 — 화살표가 가리키는 쪽에 원 배치
  const rowH = 3.5, rowGap = 0.4;
  const row1Y = topY + topH + 0.45, row2Y = row1Y + rowH + rowGap;
  const rows = [
    { y: row1Y, color: C.research, t: TEXT.research, arrow: "FaArrowLeft", side: "left", ic: "FaUserTie" },
    { y: row2Y, color: C.service, t: TEXT.service, arrow: "FaArrowRight", side: "right", ic: "FaLocationDot" },
  ];
  const cd = 1.9, inset = 0.45;
  for (const r of rows) {
    s.addShape(pres.ShapeType.roundRect, { objectName: `${r.t.area} 안내 바`, x: M, y: r.y, w: cw, h: rowH, rectRadius: 0.35, fill: { color: C.white }, line: { color: r.color, width: 4 }, shadow: shadow(0.12) });
    const circleX = r.side === "left" ? M + inset : M + cw - inset - cd;
    s.addShape(pres.ShapeType.ellipse, { objectName: `${r.t.area} 화살표 원`, x: circleX, y: r.y + (rowH - cd) / 2, w: cd, h: cd, fill: { color: r.color }, line: { type: "none" } });
    s.addImage({ data: await icon(r.arrow, C.white), x: circleX + cd / 2 - 0.55, y: r.y + rowH / 2 - 0.55, w: 1.1, h: 1.1, altText: r.side === "left" ? "왼쪽 화살표" : "오른쪽 화살표" });

    // 본문 영역: 원 반대쪽에서 시작
    const bodyX = r.side === "left" ? circleX + cd + 0.4 : M + inset + 0.1;
    const bodyW = cw - inset - cd - 0.4 - inset - 0.1;
    const labelW = 2.9, divX = bodyX + labelW + 0.2;
    s.addText(r.t.area, { objectName: `${r.t.area} 제목`, x: bodyX, y: r.y, w: labelW, h: rowH, fontFace: FONT, fontSize: 50, bold: true, color: r.color, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addShape(pres.ShapeType.line, { objectName: `${r.t.area} 구분선`, x: divX, y: r.y + 0.8, w: 0, h: rowH - 1.6, line: { color: "D5DCE8", width: 2 } });
    const detX = divX + 0.25, detW = bodyX + bodyW - detX;
    s.addImage({ data: await icon(r.ic, r.color), x: detX + detW / 2 - 0.3, y: r.y + 0.45, w: 0.6, h: 0.6, altText: r.ic === "FaUserTie" ? "담당자 아이콘" : "위치 아이콘" });
    s.addText(
      [
        { text: r.t.title, options: { fontSize: 40, color: C.ink, breakLine: true } },
        { text: r.t.detail, options: { fontSize: 34, color: "4A5468" } },
      ],
      { objectName: `${r.t.area} 접수처`, x: detX, y: r.y + 1.1, w: detW, h: rowH - 1.4, fontFace: FONT, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.2, margin: 0, isTextBox: true }
    );
  }
  await pres.writeFile({ fileName: file });
}

(async () => {
  const out = process.argv[2] || ".";
  await variantA(`${out}/교원업적평가_접수안내_A3_시안A_블루카드.pptx`);
  await variantB(`${out}/교원업적평가_접수안내_A3_시안B_길안내형.pptx`);
  await variantC(`${out}/교원업적평가_접수안내_A3_시안C_가로안내바.pptx`);
})();

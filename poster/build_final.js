// 확정 시안 다듬기
//   A3 세로 시안 A / A3 가로 시안 A (안내 문구 2줄 — 사용자 수정 반영) / A4 911호 시안 3
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa6");

const FONT = "맑은 고딕";
const A3 = { w: 11.69, h: 16.54 }, A4 = { w: 8.27, h: 11.69 };

// 문구는 확정본 그대로 사용
const TEXT = {
  head: "교원업적평가",
  l2: ["전산프로그램 출력물", "과"],
  l3: ["증빙서류 확인", " 후"],
  l4: "접수해 드리겠습니다.",
  research: { area: "연구영역", title: "산학협력단", detail: "본부동 916호" },
  service: { area: "봉사영역", title: "기획처", detail: "본부동 911호" },
};

// 인쇄 대비를 고려해 다듬은 색상
const C = {
  bg: "2172E5",
  navy: "18214F",
  ink: "1A1A1A",
  sub: "3D4556",
  cyan: "0091DA",     // 전산프로그램 출력물
  red: "E8461E",      // 증빙서류 확인
  research: "2A3AB8",
  researchTint: "EBEEFC",
  service: "F26522",
  serviceDeep: "D24E14",
  serviceTint: "FEF0E8",
  yellow: "FFC83D",
  white: "FFFFFF",
};

async function icon(name, color, px = 512) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(fa[name], { color: "#" + color, size: px })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

function newDeck(title, w, h) {
  const pres = new pptxgen();
  pres.defineLayout({ name: "CUSTOM", width: w, height: h });
  pres.layout = "CUSTOM";
  pres.title = title;
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  return pres;
}

const shadow = (opacity = 0.18) => ({ type: "outer", color: "000000", opacity, blur: 10, offset: 4, angle: 90 });
const seg = (text, color, size, breakLine = false) => ({ text, options: { color, fontSize: size, breakLine } });

const AREA_ICON = { 연구영역: "FaMicroscope", 봉사영역: "FaHandHoldingHeart" };

// 영역 제목 상자: 아이콘 + 영역 이름을 한 덩어리로 가운데 정렬
async function areaHeader(pres, s, { x, y, w, h, color, area, fontSize, iconSize, reserveL = 0, reserveR = 0 }) {
  s.addShape(pres.ShapeType.roundRect, { objectName: `${area} 제목 상자`, x, y, w, h, rectRadius: 0.22, fill: { color }, line: { type: "none" }, shadow: shadow(0.15) });
  const textW = area.length * fontSize / 72 * 1.02, gap = 0.25;
  const groupW = iconSize + gap + textW;
  const gx = x + reserveL + (w - reserveL - reserveR - groupW) / 2;
  s.addImage({ data: await icon(AREA_ICON[area], C.white), x: gx, y: y + (h - iconSize) / 2, w: iconSize, h: iconSize, altText: `${area} 아이콘` });
  s.addText(area, { objectName: `${area} 제목`, x: gx + iconSize + gap, y, w: textW + 0.2, h, fontFace: FONT, fontSize, bold: true, color: C.white, align: "left", valign: "middle", margin: 0, isTextBox: true });
}

// ════════════════════════════════════════════════════════════
// A3 세로 — 시안 A
// ════════════════════════════════════════════════════════════
async function portraitA(file) {
  const { w: W, h: H } = A3;
  const pres = newDeck("교원업적평가 접수 안내 (A3 세로)", W, H);
  const s = pres.addSlide();
  s.background = { color: C.bg };
  s.addShape(pres.ShapeType.ellipse, { objectName: "장식 원 1", x: 7.6, y: -2.6, w: 6.5, h: 6.5, fill: { color: C.white, transparency: 88 }, line: { type: "none" } });
  s.addShape(pres.ShapeType.ellipse, { objectName: "장식 원 2", x: -3.0, y: 12.0, w: 6.5, h: 6.5, fill: { color: C.white, transparency: 90 }, line: { type: "none" } });

  const M = 0.6, cw = W - 2 * M;

  const topY = 1.15, topH = 6.6;
  s.addShape(pres.ShapeType.roundRect, { objectName: "안내 문구 카드", x: M, y: topY, w: cw, h: topH, rectRadius: 0.35, fill: { color: C.white }, line: { type: "none" }, shadow: shadow() });
  const bd = 1.5;
  s.addShape(pres.ShapeType.ellipse, { objectName: "배지 원", x: W / 2 - bd / 2, y: topY - bd / 2, w: bd, h: bd, fill: { color: C.navy }, line: { color: C.white, width: 6 }, shadow: shadow(0.2) });
  s.addImage({ data: await icon("FaFileCircleCheck", C.yellow), x: W / 2 - 0.4, y: topY - 0.4, w: 0.8, h: 0.8, altText: "서류 확인 아이콘" });

  const pillW = 6.2, pillH = 1.2, pillY = topY + 1.0;
  s.addText(TEXT.head, {
    objectName: "제목", shape: pres.ShapeType.roundRect, rectRadius: 0.6,
    x: W / 2 - pillW / 2, y: pillY, w: pillW, h: pillH,
    fill: { color: C.navy }, color: C.white, fontFace: FONT, fontSize: 58, bold: true,
    align: "center", valign: "middle", margin: 0, charSpacing: 2,
  });
  const ms = 64;
  s.addText(
    [
      seg(TEXT.l2[0], C.cyan, ms), seg(TEXT.l2[1], C.ink, ms, true),
      seg(TEXT.l3[0], C.red, ms), seg(TEXT.l3[1], C.ink, ms, true),
      seg(TEXT.l4, C.ink, ms),
    ],
    { objectName: "안내 문구", x: M + 0.3, y: pillY + pillH + 0.2, w: cw - 0.6, h: topY + topH - (pillY + pillH + 0.2) - 0.25, fontFace: FONT, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.15, margin: 0, isTextBox: true }
  );

  const botY = topY + topH + 0.45, botH = H - M - botY;
  s.addShape(pres.ShapeType.roundRect, { objectName: "접수처 카드", x: M, y: botY, w: cw, h: botH, rectRadius: 0.35, fill: { color: C.white }, line: { type: "none" }, shadow: shadow() });

  const pad = 0.4, gap = 0.4, colW = (cw - 2 * pad - gap) / 2;
  const headY = botY + pad, headH = 1.8;
  const resH = 3.3, resY = botY + botH - pad - resH;
  const cd = 1.55, cY = (headY + headH + resY) / 2 - cd / 2;
  const cols = [
    { x: M + pad, color: C.research, t: TEXT.research, arrow: "FaArrowLeft", alt: "왼쪽 화살표" },
    { x: M + pad + colW + gap, color: C.service, t: TEXT.service, arrow: "FaArrowRight", alt: "오른쪽 화살표" },
  ];
  for (const c of cols) {
    await areaHeader(pres, s, { x: c.x, y: headY, w: colW, h: headH, color: c.color, area: c.t.area, fontSize: 54, iconSize: 0.66 });

    const cx = c.x + colW / 2 - cd / 2;
    s.addShape(pres.ShapeType.ellipse, { objectName: `${c.t.area} 화살표 원`, x: cx, y: cY, w: cd, h: cd, fill: { color: c.color }, line: { color: C.white, width: 5 }, shadow: shadow(0.2) });
    s.addImage({ data: await icon(c.arrow, C.white), x: cx + cd / 2 - 0.45, y: cY + cd / 2 - 0.45, w: 0.9, h: 0.9, altText: c.alt });

    s.addShape(pres.ShapeType.roundRect, { objectName: `${c.t.area} 접수처`, x: c.x, y: resY, w: colW, h: resH, rectRadius: 0.22, fill: { color: c.color }, line: { type: "none" }, shadow: shadow(0.15) });
    s.addImage({ data: await icon("FaLocationDot", C.yellow), x: c.x + colW / 2 - 0.33, y: resY + 0.32, w: 0.66, h: 0.66, altText: "위치 아이콘" });
    s.addText(
      [seg(c.t.title, C.white, 52, true), seg(c.t.detail, C.white, 42)],
      { objectName: `${c.t.area} 접수처 문구`, x: c.x + 0.1, y: resY + 1.1, w: colW - 0.2, h: resH - 1.3, fontFace: FONT, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.2, margin: 0, isTextBox: true }
    );
  }
  await pres.writeFile({ fileName: file });
}

// ════════════════════════════════════════════════════════════
// A3 가로 — 시안 A (안내 문구 2줄)
// ════════════════════════════════════════════════════════════
async function landscapeA(file) {
  const W = A3.h, H = A3.w;
  const pres = newDeck("교원업적평가 접수 안내 (A3 가로)", W, H);
  const s = pres.addSlide();
  s.background = { color: C.bg };
  s.addShape(pres.ShapeType.ellipse, { objectName: "장식 원 1", x: 12.6, y: -2.8, w: 6.0, h: 6.0, fill: { color: C.white, transparency: 88 }, line: { type: "none" } });
  s.addShape(pres.ShapeType.ellipse, { objectName: "장식 원 2", x: -2.6, y: 7.6, w: 5.5, h: 5.5, fill: { color: C.white, transparency: 90 }, line: { type: "none" } });

  const M = 0.5, cw = W - 2 * M;

  const topY = 0.95, topH = 5.05;
  s.addShape(pres.ShapeType.roundRect, { objectName: "안내 문구 카드", x: M, y: topY, w: cw, h: topH, rectRadius: 0.3, fill: { color: C.white }, line: { type: "none" }, shadow: shadow() });
  const bd = 1.1;
  s.addShape(pres.ShapeType.ellipse, { objectName: "배지 원", x: W / 2 - bd / 2, y: topY - bd / 2, w: bd, h: bd, fill: { color: C.navy }, line: { color: C.white, width: 5 }, shadow: shadow(0.2) });
  s.addImage({ data: await icon("FaFileCircleCheck", C.yellow), x: W / 2 - 0.3, y: topY - 0.3, w: 0.6, h: 0.6, altText: "서류 확인 아이콘" });

  const pillW = 5.4, pillH = 1.05, pillY = topY + 0.8;
  s.addText(TEXT.head, {
    objectName: "제목", shape: pres.ShapeType.roundRect, rectRadius: 0.525,
    x: W / 2 - pillW / 2, y: pillY, w: pillW, h: pillH,
    fill: { color: C.navy }, color: C.white, fontFace: FONT, fontSize: 50, bold: true,
    align: "center", valign: "middle", margin: 0, charSpacing: 2,
  });
  const ms = 54;
  s.addText(
    [
      seg(TEXT.l2[0], C.cyan, ms), seg(TEXT.l2[1] + " ", C.ink, ms),
      seg(TEXT.l3[0], C.red, ms), seg(TEXT.l3[1], C.ink, ms, true),
      seg(TEXT.l4, C.ink, ms),
    ],
    { objectName: "안내 문구", x: M + 0.3, y: pillY + pillH + 0.15, w: cw - 0.6, h: topY + topH - (pillY + pillH + 0.15) - 0.25, fontFace: FONT, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.25, margin: 0, isTextBox: true }
  );

  const botY = topY + topH + 0.4, botH = H - M - botY;
  s.addShape(pres.ShapeType.roundRect, { objectName: "접수처 카드", x: M, y: botY, w: cw, h: botH, rectRadius: 0.3, fill: { color: C.white }, line: { type: "none" }, shadow: shadow() });

  const pad = 0.35, gap = 0.35, colW = (cw - 2 * pad - gap) / 2;
  const headY = botY + pad, headH = 1.5;
  const resY = headY + headH + 0.25, resH = botY + botH - pad - resY;
  const cols = [
    { x: M + pad, color: C.research, tint: C.researchTint, deep: C.research, t: TEXT.research, arrow: "FaArrowLeft", side: "left" },
    { x: M + pad + colW + gap, color: C.service, tint: C.serviceTint, deep: C.serviceDeep, t: TEXT.service, arrow: "FaArrowRight", side: "right" },
  ];
  for (const c of cols) {
    const cd = 1.1, cx = c.side === "left" ? c.x + 0.22 : c.x + colW - 0.22 - cd;
    await areaHeader(pres, s, { x: c.x, y: headY, w: colW, h: headH, color: c.color, area: c.t.area, fontSize: 58, iconSize: 0.68, reserveL: c.side === "left" ? cd + 0.3 : 0, reserveR: c.side === "right" ? cd + 0.3 : 0 });
    s.addShape(pres.ShapeType.ellipse, { objectName: `${c.t.area} 화살표 원`, x: cx, y: headY + (headH - cd) / 2, w: cd, h: cd, fill: { color: C.white }, line: { type: "none" } });
    s.addImage({ data: await icon(c.arrow, c.color), x: cx + cd / 2 - 0.34, y: headY + headH / 2 - 0.34, w: 0.68, h: 0.68, altText: c.side === "left" ? "왼쪽 화살표" : "오른쪽 화살표" });

    s.addShape(pres.ShapeType.roundRect, { objectName: `${c.t.area} 접수처`, x: c.x, y: resY, w: colW, h: resH, rectRadius: 0.22, fill: { color: c.tint }, line: { color: c.color, width: 3 } });
    // 위치 아이콘 + 접수처를 가로 한 덩어리로 배치
    const icS = 0.8, tw = 4.2, gw = icS + 0.3 + tw, gx = c.x + (colW - gw) / 2;
    s.addImage({ data: await icon("FaLocationDot", c.color), x: gx, y: resY + resH / 2 - icS / 2, w: icS, h: icS, altText: "위치 아이콘" });
    s.addText(
      [seg(c.t.title, c.deep, 48, true), seg(c.t.detail, C.ink, 42)],
      { objectName: `${c.t.area} 접수처 문구`, x: gx + icS + 0.3, y: resY, w: tw, h: resH, fontFace: FONT, bold: true, align: "left", valign: "middle", lineSpacingMultiple: 1.1, margin: 0, isTextBox: true }
    );
  }
  await pres.writeFile({ fileName: file });
}

// ════════════════════════════════════════════════════════════
// A4 세로 — 911호 문 앞 시안 3
// ════════════════════════════════════════════════════════════
async function door3(file) {
  const { w: W, h: H } = A4;
  const pres = newDeck("911호 봉사영역 접수 안내 (A4)", W, H);
  const s = pres.addSlide();
  s.background = { color: C.white };

  const M = 0.45;
  s.addShape(pres.ShapeType.roundRect, { objectName: "테두리", x: M, y: M, w: W - 2 * M, h: H - 2 * M, rectRadius: 0.3, fill: { color: C.white }, line: { color: C.service, width: 7 } });

  const id = 1.5;
  s.addShape(pres.ShapeType.ellipse, { objectName: "봉사 아이콘 원", x: W / 2 - id / 2, y: 0.95, w: id, h: id, fill: { color: C.service }, line: { color: C.serviceTint, width: 6 }, shadow: shadow(0.15) });
  s.addImage({ data: await icon("FaHandHoldingHeart", C.white), x: W / 2 - 0.43, y: 0.95 + id / 2 - 0.43, w: 0.86, h: 0.86, altText: "봉사 아이콘" });
  s.addText(TEXT.head, { objectName: "제목", x: M, y: 2.6, w: W - 2 * M, h: 0.62, fontFace: FONT, fontSize: 34, bold: true, color: C.navy, align: "center", valign: "middle", margin: 0, charSpacing: 3, isTextBox: true });
  s.addText(TEXT.service.area, { objectName: "봉사영역", x: M, y: 3.22, w: W - 2 * M, h: 1.4, fontFace: FONT, fontSize: 78, bold: true, color: C.service, align: "center", valign: "middle", margin: 0, isTextBox: true });

  const ms = 40;
  s.addText(
    [
      seg(TEXT.l2[0], C.cyan, ms), seg(TEXT.l2[1], C.ink, ms, true),
      seg(TEXT.l3[0], C.red, ms), seg(TEXT.l3[1], C.ink, ms, true),
      seg(TEXT.l4, C.ink, ms),
    ],
    { objectName: "안내 문구", x: M + 0.3, y: 4.85, w: W - 2 * M - 0.6, h: 3.1, fontFace: FONT, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.3, margin: 0, isTextBox: true }
  );

  const bx = M + 0.45, bw = W - 2 * M - 0.9, by = 8.25, bh = 2.45;
  s.addShape(pres.ShapeType.roundRect, { objectName: "접수처 상자", x: bx, y: by, w: bw, h: bh, rectRadius: 0.25, fill: { color: C.navy }, line: { type: "none" }, shadow: shadow(0.2) });
  const icS = 0.9;
  s.addImage({ data: await icon("FaLocationDot", C.yellow), x: bx + 0.55, y: by + bh / 2 - icS / 2, w: icS, h: icS, altText: "위치 아이콘" });
  s.addText(
    [seg(TEXT.service.title, C.yellow, 36, true), seg(TEXT.service.detail, C.white, 46)],
    { objectName: "접수처", x: bx + 1.75, y: by, w: bw - 2.0, h: bh, fontFace: FONT, bold: true, align: "left", valign: "middle", lineSpacingMultiple: 1.15, margin: 0, isTextBox: true }
  );
  await pres.writeFile({ fileName: file });
}

(async () => {
  const out = process.argv[2] || ".";
  await portraitA(`${out}/확정_교원업적평가_접수안내_A3세로.pptx`);
  await landscapeA(`${out}/확정_교원업적평가_접수안내_A3가로.pptx`);
  await door3(`${out}/확정_911호_봉사영역_접수안내_A4.pptx`);
})();

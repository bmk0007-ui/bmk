// 911호(봉사영역 접수처) 문 앞 부착용 안내문 시안 3종 (A4 세로, 210 x 297 mm)
//   1: 오렌지 카드형  2: 호실 강조형  3: 화이트 미니멀형
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa6");

const W = 8.27, H = 11.69;
const FONT = "맑은 고딕";

// 문구는 원본 그대로 사용
const TEXT = {
  head: "교원업적평가",
  l2: ["전산프로그램 출력물", "과"],
  l3: ["증빙서류 확인", " 후"],
  l4: "접수해 드리겠습니다.",
  area: "봉사영역",
  title: "기획처",
  detail: "본부동 911호",
};

const C = {
  orange: "FF6A2B",
  orangeDeep: "E8501A",
  orangeTint: "FFF1EA",
  navy: "1B2559",
  ink: "1A1A1A",
  cyan: "0AA5E8",
  red: "FF5A2C",
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
  pres.defineLayout({ name: "A4_PORTRAIT", width: W, height: H });
  pres.layout = "A4_PORTRAIT";
  pres.title = title;
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  return pres;
}

const shadow = (opacity = 0.18) => ({ type: "outer", color: "000000", opacity, blur: 8, offset: 3, angle: 90 });
const runs = (parts, size) => parts.map(([text, color], i) => ({
  text, options: { color, fontSize: size, breakLine: i === parts.length - 1 },
}));
const message = (size, ink, c2, c3) => [
  ...runs([[TEXT.l2[0], c2], [TEXT.l2[1], ink]], size),
  ...runs([[TEXT.l3[0], c3], [TEXT.l3[1], ink]], size),
  { text: TEXT.l4, options: { color: ink, fontSize: size } },
];

// ════════════════════════════════════════════════════════════
// 시안 1 — 오렌지 카드형
// ════════════════════════════════════════════════════════════
async function door1(file) {
  const pres = newDeck("911호 봉사영역 접수 안내 - 시안 1");
  const s = pres.addSlide();
  s.background = { color: C.orange };
  s.addShape(pres.ShapeType.ellipse, { objectName: "장식 원 1", x: 5.4, y: -1.6, w: 4.4, h: 4.4, fill: { color: C.white, transparency: 88 }, line: { type: "none" } });
  s.addShape(pres.ShapeType.ellipse, { objectName: "장식 원 2", x: -1.8, y: 8.9, w: 4.0, h: 4.0, fill: { color: C.white, transparency: 90 }, line: { type: "none" } });

  const M = 0.45, cw = W - 2 * M;

  // 상단 안내 카드
  const topY = 0.95, topH = 5.0;
  s.addShape(pres.ShapeType.roundRect, { objectName: "안내 문구 카드", x: M, y: topY, w: cw, h: topH, rectRadius: 0.25, fill: { color: C.white }, line: { type: "none" }, shadow: shadow() });
  const bd = 1.05;
  s.addShape(pres.ShapeType.ellipse, { objectName: "배지 원", x: W / 2 - bd / 2, y: topY - bd / 2, w: bd, h: bd, fill: { color: C.navy }, line: { color: C.white, width: 5 }, shadow: shadow(0.2) });
  s.addImage({ data: await icon("FaFileCircleCheck", C.white), x: W / 2 - 0.28, y: topY - 0.28, w: 0.56, h: 0.56, altText: "서류 확인 아이콘" });
  const pillW = 4.3, pillH = 0.85, pillY = topY + 0.75;
  s.addText(TEXT.head, {
    objectName: "제목", shape: pres.ShapeType.roundRect, rectRadius: 0.425,
    x: W / 2 - pillW / 2, y: pillY, w: pillW, h: pillH,
    fill: { color: C.navy }, color: C.white, fontFace: FONT, fontSize: 40, bold: true,
    align: "center", valign: "middle", margin: 0, charSpacing: 2,
  });
  s.addText(message(40, C.ink, C.cyan, C.red), {
    objectName: "안내 문구", x: M + 0.2, y: pillY + pillH + 0.15, w: cw - 0.4, h: topY + topH - (pillY + pillH + 0.15) - 0.2,
    fontFace: FONT, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.2, margin: 0, isTextBox: true,
  });

  // 하단 접수처 카드
  const botY = topY + topH + 0.35, botH = H - M - botY;
  s.addShape(pres.ShapeType.roundRect, { objectName: "접수처 카드", x: M, y: botY, w: cw, h: botH, rectRadius: 0.25, fill: { color: C.white }, line: { type: "none" }, shadow: shadow() });
  const pad = 0.35;
  const hd = 1.5;
  s.addShape(pres.ShapeType.roundRect, { objectName: "봉사영역 제목 상자", x: M + pad, y: botY + pad, w: cw - 2 * pad, h: hd, rectRadius: 0.2, fill: { color: C.orange }, line: { type: "none" } });
  s.addImage({ data: await icon("FaHandHoldingHeart", C.white), x: M + pad + 0.45, y: botY + pad + hd / 2 - 0.42, w: 0.84, h: 0.84, altText: "봉사 아이콘" });
  s.addText(TEXT.area, { objectName: "봉사영역 제목", x: M + pad + 1.4, y: botY + pad, w: cw - 2 * pad - 2.8, h: hd, fontFace: FONT, fontSize: 56, bold: true, color: C.white, align: "center", valign: "middle", margin: 0, isTextBox: true });

  const ry = botY + pad + hd + 0.25, rh = botY + botH - pad - ry;
  s.addShape(pres.ShapeType.roundRect, { objectName: "접수처 상자", x: M + pad, y: ry, w: cw - 2 * pad, h: rh, rectRadius: 0.2, fill: { color: C.orangeTint }, line: { color: C.orange, width: 3 } });
  s.addImage({ data: await icon("FaLocationDot", C.orange), x: W / 2 - 0.27, y: ry + 0.3, w: 0.54, h: 0.54, altText: "위치 아이콘" });
  s.addText(
    [
      { text: TEXT.title, options: { fontSize: 40, color: C.orangeDeep, breakLine: true } },
      { text: TEXT.detail, options: { fontSize: 34, color: C.ink } },
    ],
    { objectName: "접수처", x: M + pad + 0.2, y: ry + 0.9, w: cw - 2 * pad - 0.4, h: rh - 1.05, fontFace: FONT, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.2, margin: 0, isTextBox: true }
  );
  await pres.writeFile({ fileName: file });
}

// ════════════════════════════════════════════════════════════
// 시안 2 — 호실 강조형 (911호를 크게)
// ════════════════════════════════════════════════════════════
async function door2(file) {
  const pres = newDeck("911호 봉사영역 접수 안내 - 시안 2");
  const s = pres.addSlide();
  s.background = { color: C.white };

  // 상단 오렌지 면: 봉사영역 + 기획처 + 911호
  const topH = 5.6, M = 0.55;
  s.addShape(pres.ShapeType.rect, { objectName: "상단 색면", x: 0, y: 0, w: W, h: topH, fill: { color: C.orange }, line: { type: "none" } });
  s.addShape(pres.ShapeType.ellipse, { objectName: "장식 원", x: 5.2, y: 2.4, w: 4.6, h: 4.6, fill: { color: C.white, transparency: 88 }, line: { type: "none" } });

  s.addText(TEXT.head, {
    objectName: "제목", shape: pres.ShapeType.roundRect, rectRadius: 0.325,
    x: M, y: 0.6, w: 3.2, h: 0.65, fill: { color: C.white }, color: C.orangeDeep,
    fontFace: FONT, fontSize: 28, bold: true, align: "center", valign: "middle", margin: 0, charSpacing: 1,
  });
  s.addText(TEXT.area, { objectName: "봉사영역", x: M, y: 1.5, w: W - 2 * M, h: 1.3, fontFace: FONT, fontSize: 72, bold: true, color: C.white, align: "left", valign: "middle", margin: 0, isTextBox: true });
  s.addImage({ data: await icon("FaLocationDot", C.white), x: M, y: 3.36, w: 0.5, h: 0.5, altText: "위치 아이콘" });
  s.addText(
    [
      { text: TEXT.title, options: { fontSize: 36, breakLine: true } },
      { text: TEXT.detail, options: { fontSize: 60 } },
    ],
    { objectName: "접수처", x: M + 0.75, y: 3.0, w: W - 2 * M - 0.75, h: 2.2, fontFace: FONT, bold: true, color: C.white, align: "left", valign: "middle", lineSpacingMultiple: 1.15, margin: 0, isTextBox: true }
  );

  // 하단 안내 문구
  const id = 0.95;
  s.addShape(pres.ShapeType.ellipse, { objectName: "아이콘 원", x: W / 2 - id / 2, y: topH + 0.55, w: id, h: id, fill: { color: C.orangeTint }, line: { type: "none" } });
  s.addImage({ data: await icon("FaClipboardCheck", C.orange), x: W / 2 - 0.25, y: topH + 0.55 + id / 2 - 0.28, w: 0.5, h: 0.56, altText: "확인 아이콘" });
  s.addText(message(40, C.ink, C.cyan, C.red), {
    objectName: "안내 문구", x: M, y: topH + 1.7, w: W - 2 * M, h: H - topH - 1.7 - 0.7,
    fontFace: FONT, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.35, margin: 0, isTextBox: true,
  });
  await pres.writeFile({ fileName: file });
}

// ════════════════════════════════════════════════════════════
// 시안 3 — 화이트 미니멀형
// ════════════════════════════════════════════════════════════
async function door3(file) {
  const pres = newDeck("911호 봉사영역 접수 안내 - 시안 3");
  const s = pres.addSlide();
  s.background = { color: C.white };

  const M = 0.45;
  s.addShape(pres.ShapeType.roundRect, { objectName: "테두리", x: M, y: M, w: W - 2 * M, h: H - 2 * M, rectRadius: 0.3, fill: { color: C.white }, line: { color: C.orange, width: 6 } });

  // 상단: 아이콘 + 봉사영역
  const id = 1.4;
  s.addShape(pres.ShapeType.ellipse, { objectName: "봉사 아이콘 원", x: W / 2 - id / 2, y: 1.0, w: id, h: id, fill: { color: C.orange }, line: { type: "none" }, shadow: shadow(0.15) });
  s.addImage({ data: await icon("FaHandHoldingHeart", C.white), x: W / 2 - 0.4, y: 1.0 + id / 2 - 0.4, w: 0.8, h: 0.8, altText: "봉사 아이콘" });
  s.addText(TEXT.head, { objectName: "제목", x: M, y: 2.6, w: W - 2 * M, h: 0.6, fontFace: FONT, fontSize: 30, bold: true, color: C.navy, align: "center", valign: "middle", margin: 0, charSpacing: 3, isTextBox: true });
  s.addText(TEXT.area, { objectName: "봉사영역", x: M, y: 3.2, w: W - 2 * M, h: 1.3, fontFace: FONT, fontSize: 72, bold: true, color: C.orange, align: "center", valign: "middle", margin: 0, isTextBox: true });

  // 가운데: 안내 문구
  s.addText(message(36, C.ink, C.cyan, C.red), {
    objectName: "안내 문구", x: M + 0.3, y: 4.85, w: W - 2 * M - 0.6, h: 3.0,
    fontFace: FONT, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.3, margin: 0, isTextBox: true,
  });

  // 하단: 접수처 상자
  const bx = M + 0.5, bw = W - 2 * M - 1.0, by = 8.3, bh = 2.35;
  s.addShape(pres.ShapeType.roundRect, { objectName: "접수처 상자", x: bx, y: by, w: bw, h: bh, rectRadius: 0.25, fill: { color: C.navy }, line: { type: "none" }, shadow: shadow(0.2) });
  s.addImage({ data: await icon("FaLocationDot", "FFD54A"), x: bx + 0.5, y: by + bh / 2 - 0.4, w: 0.8, h: 0.8, altText: "위치 아이콘" });
  s.addText(
    [
      { text: TEXT.title, options: { fontSize: 38, color: "FFD54A", breakLine: true } },
      { text: TEXT.detail, options: { fontSize: 36, color: C.white } },
    ],
    { objectName: "접수처", x: bx + 1.6, y: by, w: bw - 2.0, h: bh, fontFace: FONT, bold: true, align: "left", valign: "middle", lineSpacingMultiple: 1.2, margin: 0, isTextBox: true }
  );
  await pres.writeFile({ fileName: file });
}

(async () => {
  const out = process.argv[2] || ".";
  await door1(`${out}/911호_봉사영역_접수안내_A4_시안1_오렌지카드.pptx`);
  await door2(`${out}/911호_봉사영역_접수안내_A4_시안2_호실강조.pptx`);
  await door3(`${out}/911호_봉사영역_접수안내_A4_시안3_화이트미니멀.pptx`);
})();

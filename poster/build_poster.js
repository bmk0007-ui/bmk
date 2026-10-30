// 교원업적평가 접수 안내 포스터 (A3 세로, 297 x 420 mm)
const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.defineLayout({ name: "A3_PORTRAIT", width: 11.69, height: 16.54 });
pres.layout = "A3_PORTRAIT";
pres.title = "교원업적평가 접수 안내";
const FONT = "맑은 고딕";
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };

const COLOR = {
  bg: "1E7BE0",
  card: "FFFFFF",
  text: "1A1A1A",
  cyan: "0AA5E8",
  red: "FF5A2C",
  research: "3446C6",
  service: "FF6A2B",
};

const W = 11.69;
const M = 0.6; // 바깥 여백
const cardShadow = () => ({ type: "outer", color: "000000", opacity: 0.18, blur: 8, offset: 3, angle: 90 });

const slide = pres.addSlide();
slide.background = { color: COLOR.bg };

// ── 상단 안내 문구 카드 ─────────────────────────────
const topY = M, topH = 7.0;
slide.addShape(pres.ShapeType.roundRect, {
  objectName: "안내 문구 카드",
  x: M, y: topY, w: W - 2 * M, h: topH,
  fill: { color: COLOR.card }, rectRadius: 0.35, line: { color: COLOR.card }, shadow: cardShadow(),
});

const line = (runs) => runs.map(([text, color], i) => ({
  text, options: { color, breakLine: i === runs.length - 1 },
}));
slide.addText(
  [
    ...line([["교원업적평가", COLOR.text]]),
    ...line([["전산프로그램 출력물", COLOR.cyan], ["과", COLOR.text]]),
    ...line([["증빙서류 확인", COLOR.red], [" 후", COLOR.text]]),
    { text: "접수해 드리겠습니다.", options: { color: COLOR.text } },
  ],
  {
    objectName: "안내 문구",
    x: M + 0.3, y: topY + 0.3, w: W - 2 * M - 0.6, h: topH - 0.6,
    fontFace: FONT, fontSize: 64, bold: true, align: "center", valign: "middle",
    lineSpacingMultiple: 1.25, margin: 0, isTextBox: true,
  }
);

// ── 하단 접수처 카드 ────────────────────────────────
const botY = topY + topH + 0.5, botH = 16.54 - M - botY;
slide.addShape(pres.ShapeType.roundRect, {
  objectName: "접수처 카드",
  x: M, y: botY, w: W - 2 * M, h: botH,
  fill: { color: COLOR.card }, rectRadius: 0.35, line: { color: COLOR.card }, shadow: cardShadow(),
});

const gap = 0.4, pad = 0.4;
const colW = (W - 2 * M - 2 * pad - gap) / 2;
const columns = [
  { x: M + pad, color: COLOR.research, area: "연구영역", title: "산학협력단", detail: "본부동 916호" },
  { x: M + pad + colW + gap, color: COLOR.service, area: "봉사영역", title: "기획처", detail: "본부동 911호" },
];

const headY = botY + pad, headH = 1.9;
const resH = 2.9, resY = botY + botH - pad - resH;
const circleD = 1.5, circleY = (headY + headH + resY) / 2 - circleD / 2;

for (const c of columns) {
  // 영역 이름
  slide.addText(c.area, {
    objectName: `${c.area} 제목`,
    shape: pres.ShapeType.roundRect, rectRadius: 0.25,
    x: c.x, y: headY, w: colW, h: headH,
    fill: { color: c.color }, color: "FFFFFF",
    fontFace: FONT, fontSize: 66, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // 화살표 (원 + 아래 화살표)
  const cx = c.x + colW / 2 - circleD / 2;
  slide.addShape(pres.ShapeType.ellipse, {
    objectName: `${c.area} 화살표 원`,
    x: cx, y: circleY, w: circleD, h: circleD, fill: { color: c.color }, line: { color: c.color },
  });
  const aw = 0.75, ah = 0.95;
  slide.addShape(pres.ShapeType.downArrow, {
    objectName: `${c.area} 화살표`,
    x: cx + (circleD - aw) / 2, y: circleY + (circleD - ah) / 2, w: aw, h: ah,
    fill: { color: "FFFFFF" }, line: { color: "FFFFFF" },
  });

  // 접수처
  slide.addText(
    [
      { text: c.title, options: { fontSize: 50, breakLine: true } },
      { text: c.detail, options: { fontSize: 40 } },
    ],
    {
      objectName: `${c.area} 접수처`,
      shape: pres.ShapeType.roundRect, rectRadius: 0.25,
      x: c.x, y: resY, w: colW, h: resH,
      fill: { color: c.color }, color: "FFFFFF",
      fontFace: FONT, bold: true, align: "center", valign: "middle",
      lineSpacingMultiple: 1.3, margin: 0,
    }
  );
}

pres.writeFile({ fileName: process.argv[2] || "교원업적평가_접수안내_A3.pptx" });

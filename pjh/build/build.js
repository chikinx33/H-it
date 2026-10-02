// 박진희 포트폴리오: 애니메이션 / 캐릭터 IP 웹툰 / 숏폼 광고 3개 작품
// node build.js  ->  ../박진희_포트폴리오.pptx
// 코코아 브라운 바탕, 화로 앰버 강조, 작품들을 하나로 잇는 「연결 꼬리」 곡선 모티프
// 챕터 간지: 좌측 작품 색 띠 + 대형 이탤릭 번호 + 흰 꼬리 선, 우측 대표 이미지
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const IMG = path.join(__dirname, "../work/img");
const DIMS = require("../work/dims.json");
const OUT = path.join(__dirname, "../박진희_포트폴리오.pptx");

const LINK_AD = "https://drive.google.com/file/d/1nawR_erkM8cg4L6-n3Zheibxt1GWwylb/view";
const LINK_MOTION = "https://drive.google.com/file/d/1syIEN2KBbWDb8Q02cjB6IYgN444v6UqF/view";

const SANS = "Malgun Gothic";
const NUM = "Georgia";
const COCOA = "1A110D", CARD = "2A1C16", CREAM = "F7F0E6", PAPER = "FFFDF9", INKT = "2B1D16", MUTED = "8A7566";
const EMBER = "E8552B", AMBER = "F2A65A", CORAL = "EE6F7C", CORAL_D = "D2505F", KRAFT = "B8895A", KRAFT_D = "8E6236", NIGHT = "8FA3C7";

const THEME = {
  name: "Connected Tail",
  headFontFace: SANS,
  bodyFontFace: SANS,
  colors: { dk1: COCOA, lt1: CREAM, dk2: CARD, lt2: PAPER, accent1: EMBER, accent2: AMBER, accent3: CORAL, accent4: KRAFT, accent5: NIGHT, accent6: MUTED, hlink: AMBER, folHlink: AMBER },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: SANS, bodyFontFace: SANS };
pres.author = "박진희";
pres.title = "박진희 포트폴리오";
const W = 13.333, H = 7.5, MX = 0.7;

const footer = (c) => ({ text: { text: "박진희  ·  AI 콘텐츠 크리에이터 포트폴리오", options: { x: MX, y: 7.08, w: 6, h: 0.28, fontFace: SANS, fontSize: 9, color: c, margin: 0 } } });
const titlePh = (c) => ({ placeholder: { options: { name: "title", type: "title", x: MX, y: 0.74, w: 11.9, h: 0.72, fontFace: SANS, fontSize: 28, bold: true, color: c, align: "left", valign: "middle", margin: 0 }, text: "" } });
const num = (c) => ({ x: 12.13, y: 7.08, w: 0.5, h: 0.28, fontFace: NUM, fontSize: 10, italic: true, color: c, align: "right" });
pres.defineSlideMaster({ title: "DARK", background: { color: COCOA }, objects: [titlePh(CREAM), footer("8C7464")], slideNumber: num("8C7464") });
pres.defineSlideMaster({ title: "LIGHT", background: { color: CREAM }, objects: [titlePh(INKT), footer(MUTED)], slideNumber: num(MUTED) });
pres.defineSlideMaster({ title: "PLAIN", background: { color: COCOA }, objects: [] });

// ---------- helpers ----------
let sid = 0;
const nm = (s) => `${s}-${++sid}`;
const file = (key) => path.join(IMG, `${key}.${key.startsWith("tail_") ? "png" : "jpg"}`);
function fit(key, x, y, w, h) {
  const [iw, ih] = DIMS[key];
  const r = iw / ih;
  let fw = w, fh = w / r;
  if (fh > h) { fh = h; fw = h * r; }
  return { x: x + (w - fw) / 2, y: y + (h - fh) / 2, w: fw, h: fh };
}
function img(s, key, x, y, w, h, o = {}) {
  let f = fit(key, x, y, w, h);
  if (o.top) f = { ...f, y };
  if (o.left) f = { ...f, x };
  if (o.border) s.addShape(pres.shapes.RECTANGLE, { x: f.x - 0.05, y: f.y - 0.05, w: f.w + 0.1, h: f.h + 0.1, fill: { color: o.border }, line: { type: "none" }, shadow: o.shadow ? { type: "outer", color: "000000", opacity: 0.35, blur: 8, offset: 3, angle: 90 } : undefined, objectName: nm("border") });
  const link = o.url ? { hyperlink: { url: o.url, tooltip: o.tip || "영상 보기" } } : {};
  s.addImage({ path: file(key), ...f, rounding: !!o.round, altText: o.alt || key, objectName: nm(`img-${key}`), ...link });
  return f;
}
function t(s, text, o) { s.addText(text, { fontFace: SANS, color: CREAM, margin: 0, valign: "top", isTextBox: true, objectName: nm(o.name || "text"), ...o }); }
function rect(s, x, y, w, h, fill, o = {}) {
  s.addShape(o.round ? pres.shapes.ROUNDED_RECTANGLE : pres.shapes.RECTANGLE, { x, y, w, h, rectRadius: o.round || 0, fill: { color: fill, transparency: o.tr || 0 }, line: o.line ? { color: o.line, width: o.lw || 0.75 } : { type: "none" }, objectName: nm("rect") });
}
function dot(s, x, y, d, fill, o = {}) {
  s.addShape(pres.shapes.OVAL, { x: x - d / 2, y: y - d / 2, w: d, h: d, fill: { color: fill, transparency: o.tr || 0 }, line: o.line ? { color: o.line, width: o.lw || 1 } : { type: "none" }, objectName: nm("dot") });
}
function full(s, key, alt) { s.addImage({ path: file(key), x: 0, y: 0, w: W, h: H, altText: alt || "배경", objectName: nm(key) }); }
// 일반 슬라이드: 챕터 라벨(작품 색) + 제목 placeholder
function slide(master, label, labelColor, title, section, bgKey) {
  const s = pres.addSlide({ masterName: master, sectionTitle: section });
  if (bgKey) full(s, bgKey);
  dot(s, MX + 0.06, 0.55, 0.12, labelColor);
  t(s, label, { x: MX + 0.24, y: 0.4, w: 9, h: 0.3, fontSize: 11, bold: true, color: labelColor, charSpacing: 2, valign: "middle", name: "label" });
  s.addText(title, { placeholder: "title" });
  return s;
}
// 개조식 목록 (작품 색 점)
function bullets(s, items, x, y, w, o = {}) {
  const color = o.color || CREAM, mark = o.mark || EMBER, size = o.size || 15, gap = o.gap || 0.62;
  items.forEach((it, i) => {
    dot(s, x + 0.07, y + i * gap + 0.16, 0.1, mark);
    t(s, it, { x: x + 0.3, y: y + i * gap, w: w - 0.3, h: gap, fontSize: size, color, lineSpacingMultiple: 1.1, name: "bullet" });
  });
}
function chip(s, text, x, y, fill, color, o = {}) {
  const w = o.w || 0.9;
  rect(s, x, y, w, 0.3, fill, { round: 0.15 });
  t(s, text, { x, y, w, h: 0.3, fontSize: o.size || 10, bold: true, color, align: "center", valign: "middle", fontFace: o.font || SANS, name: "chip" });
}
function linkText(s, label, url, x, y, w, h, color, size = 13) {
  s.addText([{ text: label, options: { hyperlink: { url, tooltip: url }, color, bold: true } }], { x, y, w, h, fontFace: SANS, fontSize: size, valign: "middle", margin: 0, isTextBox: true, objectName: nm("link") });
}

// 챕터 간지: 좌측 작품 색 띠 + 대형 번호 + 흰 꼬리 선 / 우측은 작품별 대표 이미지(drawRight)
const BAND = 4.4;
function chapter(n, color, title, en, sub, section, drawRight) {
  const s = pres.addSlide({ masterName: "PLAIN", sectionTitle: section });
  drawRight(s);
  rect(s, 0, 0, BAND, H, color);
  t(s, "CHAPTER", { x: 0.62, y: 0.62, w: 3, h: 0.3, fontFace: NUM, fontSize: 13, italic: true, color: "FFFFFF", charSpacing: 6, name: "ch-label" });
  t(s, n, { x: 0.42, y: 0.85, w: 3.9, h: 2.9, fontFace: NUM, fontSize: 190, italic: true, bold: true, color: "FFFFFF", valign: "middle", name: "ch-num" });
  t(s, en, { x: 0.62, y: 3.62, w: 3.6, h: 0.3, fontFace: NUM, fontSize: 12, italic: true, color: "FFFFFF", transparency: 15, name: "ch-en" });
  s.addImage({ path: file("tail_ch"), x: 0, y: 0, w: W, h: H, altText: "연결 꼬리 곡선", objectName: nm("tail") });
  t(s, title, { x: 0.62, y: 4.6, w: 3.6, h: 1.3, fontSize: 32, bold: true, color: "FFFFFF", valign: "top", lineSpacingMultiple: 1.0, name: "ch-title" });
  t(s, sub, { x: 0.62, y: 6.05, w: 3.6, h: 0.8, fontSize: 13, color: "FFFFFF", lineSpacingMultiple: 1.2, name: "ch-sub" });
  return s;
}

// =====================================================================
// 01 표지
pres.addSection({ title: "소개" });
{
  const s = pres.addSlide({ masterName: "PLAIN", sectionTitle: "소개" });
  full(s, "bg_cover", "불꽃 하트를 띄운 토리 장면을 흐리게 깐 배경");
  s.addImage({ path: file("tail_cover"), x: 0, y: 0, w: W, h: H, altText: "세 작품을 잇는 꼬리 곡선", objectName: nm("tail") });
  const c = [["circ_tori", 8.4, 2.55, 2.7, "토리"], ["circ_mallang", 11.35, 3.6, 2.2, "말랑"], ["circ_cat", 9.35, 5.75, 2.0, "벅뻑 고양이"]];
  c.forEach(([k, cx, cy, d, alt]) => {
    dot(s, cx, cy, d + 0.14, EMBER);
    img(s, k, cx - d / 2, cy - d / 2, d, d, { round: true, alt });
  });
  t(s, "AI CONTENT CREATOR PORTFOLIO", { x: MX, y: 1.15, w: 6, h: 0.3, fontFace: NUM, fontSize: 13, italic: true, color: AMBER, charSpacing: 4 });
  t(s, "박진희", { x: MX, y: 1.6, w: 5.5, h: 1.2, fontSize: 60, bold: true, color: CREAM });
  t(s, "캐릭터 중심 스토리텔링의 장르 확장", { x: MX, y: 2.95, w: 6, h: 0.5, fontSize: 20, color: CREAM });
  rect(s, MX, 3.75, 0.6, 0.05, EMBER);
  [["01", "AI 애니메이션 「너와 나의 연결 꼬리」"], ["02", "캐릭터 IP 웹툰 「나는 문어, 말랑」"], ["03", "AI 숏폼 광고 「벅뻑 행복스크래쳐」"]].forEach(([n, x], i) => {
    t(s, n, { x: MX, y: 4.1 + i * 0.55, w: 0.6, h: 0.45, fontFace: NUM, fontSize: 18, italic: true, bold: true, color: AMBER, valign: "middle" });
    t(s, x, { x: MX + 0.65, y: 4.1 + i * 0.55, w: 5.4, h: 0.45, fontSize: 15, color: CREAM, valign: "middle" });
  });
  t(s, "서울시 매력일자리 AI 콘텐츠 크리에이터 양성과정  ·  2026", { x: MX, y: 6.65, w: 6.5, h: 0.3, fontSize: 10, color: "B9A595" });
}

// 02 목차
{
  const s = slide("DARK", "CONTENTS", AMBER, "목차", "소개", "bg_toc");
  s.addImage({ path: file("tail_toc"), x: 0, y: 0, w: W, h: H, altText: "목차를 잇는 꼬리 곡선", objectName: nm("tail") });
  const rows = [
    ["01", "너와 나의 연결 꼬리", "AI 애니메이션 / 홍보영상용 티저", "나에겐 있고 너에겐 없는 것", EMBER, "s1c05"],
    ["02", "나는 문어, 말랑", "캐릭터 IP / 이모티콘 / 일상 웹툰", "말랑, 물렁, 말라칸?", CORAL, "conti1"],
    ["03", "벅뻑 행복스크래쳐", "AI 숏폼 광고 / 60초", "소리로 기억되는 고양이 스크래처", KRAFT, "pan_thumb"],
  ];
  rows.forEach(([n, title, genre, line, col, key], i) => {
    const y = 1.75 + i * 1.6;
    dot(s, 1.35, y + 0.5, 0.36, col, { line: CREAM, lw: 2 });
    t(s, n, { x: 1.9, y: y - 0.05, w: 1.6, h: 1.1, fontFace: NUM, fontSize: 54, italic: true, bold: true, color: col, valign: "middle" });
    t(s, title, { x: 3.55, y: y + 0.02, w: 5.2, h: 0.5, fontSize: 24, bold: true, color: CREAM });
    t(s, genre, { x: 3.55, y: y + 0.55, w: 5.2, h: 0.3, fontSize: 13, color: col === KRAFT ? "D9B98A" : col, bold: true });
    t(s, line, { x: 3.55, y: y + 0.88, w: 5.2, h: 0.3, fontSize: 12, color: "C9B6A6" });
    img(s, key, 9.6, y - 0.08, 2.5, 1.28, { border: col });
  });
}

// 03 역량 요약
{
  const s = slide("LIGHT", "PROFILE", EMBER, "캐릭터에서 출발한 매체별 스토리텔링 역량", "소개", "bg_cream");
  const cards = [
    ["캐릭터 설계", "능력과 결핍, 평상시와 변신의 반전으로 관계와 갈등이 살아 있는 캐릭터 기획", EMBER],
    ["영상 연출", "컷 단위 샷/앵글/조명/카메라 설계로 감정선 연출 구체화", AMBER],
    ["장르 확장", "애니메이션/웹툰/광고까지 매체별 문법에 맞춘 이야기 전개", CORAL],
    ["AI 활용", "AI툴을 사용하여 창작/기획/연출 역량 극대화", KRAFT],
  ];
  cards.forEach(([h, b, col], i) => {
    const x = MX + i * 3.0, y = 2.0;
    rect(s, x, y, 2.8, 4.4, PAPER, { line: "E6D9C8" });
    rect(s, x, y, 2.8, 0.12, col);
    t(s, `0${i + 1}`, { x: x + 0.3, y: y + 0.45, w: 2, h: 0.8, fontFace: NUM, fontSize: 40, italic: true, bold: true, color: col });
    t(s, h, { x: x + 0.3, y: y + 1.5, w: 2.3, h: 0.5, fontSize: 20, bold: true, color: INKT });
    t(s, b, { x: x + 0.3, y: y + 2.2, w: 2.25, h: 1.9, fontSize: 14, color: "4A3A30", lineSpacingMultiple: 1.25 });
  });
}

// =====================================================================
// CHAPTER 01 너와 나의 연결 꼬리
const C1 = "01  너와 나의 연결 꼬리";
pres.addSection({ title: "01 너와 나의 연결 꼬리" });
chapter("01", EMBER, "너와 나의\n연결 꼬리", "AI Animation Teaser", "AI 애니메이션 홍보영상용 티저\n나에겐 있고 너에겐 없는 것", "01 너와 나의 연결 꼬리", (s) => {
  img(s, "ch1_img", BAND, 0, W - BAND, H, { alt: "토리의 꼬리 불꽃이 공중에 고리를 그리는 장면" });
});

// 작품 개요
{
  const s = slide("DARK", C1, AMBER, "다르다는 것은 틀림이 아니다", "01 너와 나의 연결 꼬리", "bg_warm");
  img(s, "s1c01", MX, 1.85, 6.0, 3.4, { border: "3A2A22", alt: "어두운 가을 숲 속 창 하나만 불 켜진 오두막" });
  t(s, "S#1 C#1  숲속 오두막, 창 하나에서 새어 나오는 불빛", { x: MX, y: 5.35, w: 6.0, h: 0.3, fontSize: 11, color: "B9A595" });
  bullets(s, [
    "코믹한 불꽃 마술 소동 이면에 찌르의 부러움을 따라가는 감정선 설계",
    "부제 '나에겐 있고 너에겐 없는 것'으로 본편에서 회수할 질문 제시",
    "전반부 코미디에서 후반부 미스터리로 톤 전환, 본편 기대감 조성",
  ], 7.15, 1.85, 5.5, { gap: 0.82 });
  const rows = [["S#1  코미디", "17컷", "2분 17초"], ["S#2  미스터리", "8컷", "39초"], ["타이틀 카드", "", "4초"], ["총 러닝타임", "", "3분 0초"]];
  rows.forEach(([a, b, c], i) => {
    const y = 4.45 + i * 0.5, last = i === rows.length - 1;
    rect(s, 7.15, y, 5.5, 0.44, last ? EMBER : CARD, { tr: last ? 0 : 15 });
    t(s, a, { x: 7.35, y, w: 2.6, h: 0.44, fontSize: 13, bold: last, color: CREAM, valign: "middle" });
    t(s, b, { x: 9.9, y, w: 1.0, h: 0.44, fontSize: 13, color: CREAM, valign: "middle", align: "center" });
    t(s, c, { x: 10.9, y, w: 1.55, h: 0.44, fontSize: 13, bold: true, color: last ? "FFFFFF" : AMBER, valign: "middle", align: "right" });
  });
}

// 캐릭터
{
  const s = slide("LIGHT", C1, EMBER, "능력과 결핍의 대비로 설계한 캐릭터 관계", "01 너와 나의 연결 꼬리", "bg_cream");
  const ch = [
    ["tori3d", "토리", "꼬리 불꽃 마술로 인정받고 싶은 무대의 주인공", EMBER],
    ["jjir3d", "찌르", "부러움을 무심한 말투로 감춘 감정선의 중심", "C9822E"],
    ["nana3d", "나나", "결계 마법으로 위기를 수습하는 조력자", "6F7FA3"],
    ["witch3d", "마녀", "한마디로 소동을 제지하는 오두막의 주인", "7A5A4A"],
  ];
  ch.forEach(([k, n, d, col], i) => {
    const x = MX + i * 3.0;
    rect(s, x, 1.8, 2.8, 4.85, "FFFFFF", { line: "E6D9C8" });
    img(s, k, x + 0.1, 1.9, 2.6, 2.75, { alt: `${n} 입체 캐릭터 시트` });
    rect(s, x + 0.3, 4.85, 0.5, 0.05, col);
    t(s, n, { x: x + 0.3, y: 5.0, w: 2.3, h: 0.45, fontSize: 20, bold: true, color: INKT });
    t(s, d, { x: x + 0.3, y: 5.5, w: 2.3, h: 0.9, fontSize: 12.5, color: "4A3A30", lineSpacingMultiple: 1.2 });
  });
}

// 캐릭터 개발
{
  const s = slide("LIGHT", C1, EMBER, "평면 초안에서 입체 캐릭터로 완성도 고도화", "01 너와 나의 연결 꼬리", "bg_cream");
  rect(s, MX, 1.8, 5.5, 3.85, "FFFFFF", { line: "E6D9C8" });
  img(s, "tori1st", MX + 0.1, 1.9, 5.3, 3.65, { alt: "토리 1차 초안 시트" });
  rect(s, 7.13, 1.8, 5.5, 3.85, "FFFFFF", { line: "E6D9C8" });
  img(s, "tori3d", 7.23, 1.9, 5.3, 3.65, { alt: "토리 입체 캐릭터 시트" });
  s.addShape(pres.shapes.RIGHT_ARROW, { x: 6.37, y: 3.45, w: 0.6, h: 0.55, fill: { color: EMBER }, line: { type: "none" }, objectName: nm("arrow") });
  chip(s, "1차 초안", MX, 5.8, INKT, CREAM, { w: 1.1 });
  chip(s, "입체 시트", 7.13, 5.8, EMBER, "FFFFFF", { w: 1.1 });
  t(s, "선화 중심 초안으로 성격과 실루엣 확정", { x: MX + 1.25, y: 5.8, w: 4.2, h: 0.3, fontSize: 13, color: "4A3A30", valign: "middle" });
  t(s, "정면/측면/후면/표정 시트로 캐릭터 일관성 확보", { x: 8.38, y: 5.8, w: 4.3, h: 0.3, fontSize: 13, color: "4A3A30", valign: "middle" });
  t(s, "붉은 스카프와 꼬리 불꽃을 시그니처로 유지한 채 입체 질감/조명 반응을 더한 영상용 캐릭터 완성", { x: MX, y: 6.3, w: 11.9, h: 0.4, fontSize: 13, bold: true, color: EMBER });
}

// 공간 설계
{
  const s = slide("DARK", C1, AMBER, "이야기가 벌어지는 오두막 공간 설계", "01 너와 나의 연결 꼬리", "bg_warm");
  img(s, "cottage", MX, 1.8, 5.0, 3.6, { left: true, border: "3A2A22", alt: "이끼 덮인 지붕의 숲속 마녀 오두막 외관" });
  img(s, "kitchen", 6.15, 1.8, 6.45, 3.6, { border: "3A2A22", alt: "출입문을 바라본 오두막 부엌 밤 전경" });
  chip(s, "외관 설정화", MX, 5.6, AMBER, COCOA, { w: 1.3 });
  chip(s, "부엌 설정화", 6.15, 5.6, AMBER, COCOA, { w: 1.3 });
  t(s, "숲속 오두막 외관을 설정화로 선제작, 세계관 분위기 통일", { x: MX, y: 6.0, w: 5.2, h: 0.6, fontSize: 13, color: CREAM });
  t(s, "화로/싱크대/창가 벤치 배치를 시나리오의 인물 동선과 구도에 반영", { x: 6.15, y: 6.0, w: 6.4, h: 0.6, fontSize: 13, color: CREAM });
}

// 감정선 연출
{
  const s = slide("DARK", C1, AMBER, "불꽃 마술 소동 속 찌르의 감정선 연출", "01 너와 나의 연결 꼬리", "bg_warm");
  const cuts = [
    ["s1c03", "C#3", "역작 시리즈를 선보이는 토리의 등장"],
    ["s1c05", "C#5", "불꽃 하트 완성, 부엌 전체가 밝아지는 정점"],
    ["s1c07", "C#7", "찌르의 \"...잘하네.\" 한마디에 꺼지는 하트"],
    ["s1c14a", "C#14", "커튼으로 옮겨붙는 불씨, 소동의 클라이맥스"],
  ];
  cuts.forEach(([k, c, d], i) => {
    const x = MX + i * 3.0;
    img(s, k, x, 1.85, 2.85, 1.61, { border: "3A2A22", alt: `S#1 ${c} 장면` });
    chip(s, c, x, 3.62, EMBER, "FFFFFF", { w: 0.8, font: NUM });
    t(s, d, { x, y: 4.0, w: 2.85, h: 0.8, fontSize: 12.5, color: CREAM, lineSpacingMultiple: 1.2 });
  });
  rect(s, MX, 5.15, 11.9, 1.5, CARD, { tr: 20 });
  rect(s, MX, 5.15, 0.08, 1.5, AMBER);
  t(s, "C#17  찌르", { x: MX + 0.35, y: 5.32, w: 3, h: 0.3, fontSize: 11, bold: true, color: AMBER });
  t(s, "\"...뭐, 오늘은 실패했지만 다음이라는 것도 있으니깐!\"", { x: MX + 0.35, y: 5.65, w: 11.2, h: 0.5, fontSize: 20, bold: true, color: CREAM });
  t(s, "소동이 가라앉은 뒤 찌르의 클로즈업으로 티저의 시작과 끝을 한 인물에게 수렴", { x: MX + 0.35, y: 6.15, w: 11.2, h: 0.35, fontSize: 12.5, color: "C9B6A6" });
}

// 조명 연출
{
  const s = slide("DARK", C1, AMBER, "빛의 온도로 인물의 감정 대비 표현", "01 너와 나의 연결 꼬리", "bg_warm");
  img(s, "s1c11a", MX, 1.85, 5.85, 3.29, { border: AMBER, alt: "화로 앰버 빛을 받은 토리 클로즈업" });
  img(s, "s1c09", 6.75, 1.85, 5.85, 3.29, { border: "6F7FA3", alt: "벽 그늘 속 쿠션에 웅크린 찌르" });
  chip(s, "화로 앰버", MX, 5.32, AMBER, COCOA, { w: 1.2 });
  chip(s, "벽 그늘 냉기", 6.75, 5.32, "6F7FA3", "FFFFFF", { w: 1.3 });
  t(s, "화로 앰버 키라이트로 토리의 자신감과 무안함을 함께 부각", { x: MX, y: 5.75, w: 5.85, h: 0.6, fontSize: 13, color: CREAM });
  t(s, "벽 그늘의 냉기 단독 조명으로 시선을 피하는 찌르의 결핍 표현", { x: 6.75, y: 5.75, w: 5.85, h: 0.6, fontSize: 13, color: CREAM });
  t(s, "투샷에서도 두 빛의 대비를 유지해 같은 공간 속 감정의 거리감 시각화", { x: MX, y: 6.42, w: 11.9, h: 0.35, fontSize: 13, bold: true, color: AMBER });
}

// S#2 미스터리
{
  const s = slide("DARK", C1, NIGHT, "코미디에서 미스터리로 장르 톤 전환", "01 너와 나의 연결 꼬리", "bg_night");
  const cuts = [
    ["s2c01", "C#1", "불 꺼진 심야 부엌 창가의 정적"],
    ["s2c02", "C#2", "꺼진 화로를 확인하다 부스럭 소리에 흔들리는 카메라"],
    ["s2c33", "C#3", "소리를 따라 출입문에서 벽장으로 이동하는 시선"],
  ];
  cuts.forEach(([k, c, d], i) => {
    const x = MX + i * 4.0;
    img(s, k, x, 1.85, 3.8, 2.14, { border: "2A3550", alt: `S#2 ${c} 심야 부엌` });
    chip(s, c, x, 4.15, NIGHT, COCOA, { w: 0.8, font: NUM });
    t(s, d, { x, y: 4.55, w: 3.8, h: 0.7, fontSize: 13, color: CREAM, lineSpacingMultiple: 1.2 });
  });
  bullets(s, [
    "달빛 한 줄기만 남긴 무광 조명과 1인칭 카메라 움직임으로 긴장감 조성",
    "봉인의 서 스파크와 어둠 속 웃음소리로 본편 복선 제시, 타이틀 카드로 담백하게 마무리",
  ], MX, 5.5, 11.9, { mark: NIGHT, size: 14, gap: 0.55 });
}

// 영상
{
  const s = slide("DARK", C1, AMBER, "티저 영상", "01 너와 나의 연결 꼬리", "bg_warm");
  const f = img(s, "poster_teaser", MX, 1.85, 7.6, 4.28, { border: "3A2A22", alt: "티저 대표 장면, 불꽃 하트" });
  chip(s, "제작 진행 중", f.x + f.w / 2 - 0.85, f.y + f.h / 2 - 0.22, EMBER, "FFFFFF", { w: 1.7, size: 14 });
  t(s, "너와 나의 연결 꼬리 / 홍보영상용 티저 / 3분", { x: MX, y: 6.3, w: 7.6, h: 0.35, fontSize: 13, color: "C9B6A6" });
  rect(s, 8.75, 1.85, 3.85, 4.28, CARD, { tr: 10, line: "3A2A22" });
  img(s, "circ_tori", 9.9, 2.15, 1.55, 1.55, { round: true, alt: "토리" });
  t(s, "캐릭터 모션 테스트", { x: 9.0, y: 3.95, w: 3.35, h: 0.4, fontSize: 17, bold: true, color: CREAM, align: "center" });
  t(s, "토리의 표정/동작을 영상으로 검증한 모션 테스트", { x: 9.0, y: 4.4, w: 3.35, h: 0.7, fontSize: 12.5, color: "C9B6A6", align: "center" });
  rect(s, 9.45, 5.25, 2.45, 0.5, EMBER, { round: 0.25 });
  linkText(s, "▶  영상 보기", LINK_MOTION, 9.45, 5.25, 2.45, 0.5, "FFFFFF", 13);
}

// =====================================================================
// CHAPTER 02 나는 문어, 말랑
const C2 = "02  나는 문어, 말랑";
pres.addSection({ title: "02 나는 문어, 말랑" });
chapter("02", CORAL, "나는 문어,\n말랑", "Character IP & Webtoon", "캐릭터 IP / 이모티콘 / 일상 웹툰\n말랑, 물렁, 말라칸?", "02 나는 문어, 말랑", (s) => {
  rect(s, BAND, 0, W - BAND, H, "FFFFFF");
  img(s, "mallang", 4.9, 0.45, 6.6, 6.6, { alt: "문어 캐릭터 말랑 기본형" });
  img(s, "malakan", 10.1, 4.35, 2.75, 2.6, { top: true, border: INKT, shadow: true, alt: "말라칸으로 변신한 말랑, 웹툰 1화 마지막 컷" });
  t(s, "변신!  말라칸", { x: 10.1, y: 3.95, w: 2.75, h: 0.32, fontSize: 12, bold: true, color: CORAL_D, align: "right" });
});

// 캐릭터 설정
{
  const s = slide("LIGHT", C2, CORAL_D, "반전 변신으로 임팩트를 더한 캐릭터 IP", "02 나는 문어, 말랑", "bg_paper");
  rect(s, MX, 1.8, 4.85, 4.85, "FFFFFF", { line: "F1D5D2" });
  img(s, "mallang", MX + 0.05, 1.85, 4.75, 4.75, { alt: "말랑 캐릭터 시트" });
  const cards = [
    ["말랑", "기본형", "밝고 긍정적이지만 비실비실하고 눈치 없는 관종 캐릭터", CORAL],
    ["말라칸", "분노폭발형", "화가 나면 산을 넘는 덩치로 각성하는 과격하고 파괴적인 변신 캐릭터", INKT],
  ];
  cards.forEach(([n, type, d, col], i) => {
    const y = 1.8 + i * 1.75;
    rect(s, 5.85, y, 6.75, 1.55, "FFFFFF", { line: "F1D5D2" });
    rect(s, 5.85, y, 0.1, 1.55, col);
    t(s, n, { x: 6.2, y: y + 0.2, w: 2, h: 0.5, fontSize: 22, bold: true, color: col });
    chip(s, type, 7.75, y + 0.28, col, "FFFFFF", { w: 1.15 });
    t(s, d, { x: 6.2, y: y + 0.82, w: 6.2, h: 0.6, fontSize: 13.5, color: "4A3A30" });
  });
  rect(s, 5.85, 5.4, 6.75, 1.25, CORAL, { tr: 88 });
  t(s, "그림 스타일", { x: 6.2, y: 5.55, w: 3, h: 0.3, fontSize: 12, bold: true, color: CORAL_D });
  t(s, "기본 선화에 포인트 컬러, 눈동자 등 사물은 실사풍으로 표현한 시그니처 스타일", { x: 6.2, y: 5.9, w: 6.2, h: 0.6, fontSize: 13.5, color: INKT });
}

// 이모티콘 기획
{
  const s = slide("LIGHT", C2, CORAL_D, "감정별 대사로 구성한 이모티콘 라인업 기획", "02 나는 문어, 말랑", "bg_paper");
  const em = [
    ["플러팅형", "오늘 나랑 오징어 먹을래?", "순한 기본형에서 클로즈업으로 확장되는 플러팅"],
    ["분노폭발형", "우에해해해~~", "해골에 깜짝 놀라 말라칸으로 변신"],
    ["내숭형", "난 질기지 않아", "말랑다운 능청스러운 내숭"],
    ["기쁨형", "가오리 사세요~", "가오리를 타고 하늘을 나는 기쁨의 춤"],
    ["우울형 (블루문)", "문리버~", "블루문 아래 눈물 한 방울, 알아봐 달라는 표정"],
    ["민망형", "난 질기지 않아!", "쑥스러움을 감추려 화난 척하는 말라칸"],
  ];
  em.forEach(([n, q, d], i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = MX + col * 4.0, y = 1.8 + row * 2.2;
    rect(s, x, y, 3.8, 2.0, "FFFFFF", { line: "F1D5D2", round: 0.12 });
    t(s, n, { x: x + 0.3, y: y + 0.22, w: 3.2, h: 0.3, fontSize: 12, bold: true, color: CORAL_D });
    t(s, `"${q}"`, { x: x + 0.3, y: y + 0.6, w: 3.3, h: 0.5, fontSize: 18, bold: true, color: INKT });
    t(s, d, { x: x + 0.3, y: y + 1.25, w: 3.3, h: 0.6, fontSize: 12, color: "6A5A50", lineSpacingMultiple: 1.15 });
  });
  t(s, "기본형 + 변형 최소 6종 기획  /  첫 인사 \"나는 문어, 말랑!\"을 시리즈 제목으로 연결  /  클립스튜디오, 포토샵, 프리미어 활용", { x: MX, y: 6.3, w: 11.9, h: 0.35, fontSize: 12, bold: true, color: CORAL_D });
}

// 웹툰 1화
{
  const s = slide("LIGHT", C2, CORAL_D, "웹툰 제1화  복권 대신, 문어 당첨", "02 나는 문어, 말랑", "bg_paper");
  img(s, "conti1", MX, 1.75, 7.4, 5.2, { left: true, border: "FFFFFF", shadow: true, alt: "나는 문어 1화 콘티, 낚시줄에 걸린 문어 말랑" });
  const x = 8.45;
  chip(s, "18컷 일상 개그툰", x, 1.8, CORAL, "FFFFFF", { w: 1.75 });
  t(s, "복권 당첨 꿈을 꾼 낚시꾼 박씨가 다 함께 끌어올린 월척은 문어 말랑 한 마리", { x, y: 2.3, w: 4.2, h: 1.0, fontSize: 15, bold: true, color: INKT, lineSpacingMultiple: 1.2 });
  bullets(s, [
    "낚시꾼들의 속마음을 말풍선 속 그림으로 표현한 위트",
    "'초고추장' 한마디에 말라칸으로 즉각 변신하는 반전 결말",
    "배경 인물은 간단한 선화, 주인공은 포인트 컬러로 시선 집중",
  ], x, 3.6, 4.2, { color: "4A3A30", mark: CORAL, size: 13.5, gap: 0.85 });
}

// 웹툰 2화
{
  const s = slide("LIGHT", C2, CORAL_D, "웹툰 제2화  AI 챗너어 대 말랑", "02 나는 문어, 말랑", "bg_paper");
  img(s, "conti2", MX, 1.65, 3.6, 5.35, { left: true, border: "FFFFFF", shadow: true, alt: "나는 문어 2화 세로 콘티, 다리 6개의 반란" });
  const x = 5.0;
  chip(s, "26컷 세로 스크롤", x, 1.8, CORAL, "FFFFFF", { w: 1.75 });
  t(s, "다리 6개의 반란", { x, y: 2.25, w: 7.5, h: 0.6, fontSize: 26, bold: true, color: INKT });
  t(s, "외로운 말랑이 AI 챗봇에게 프로필 그림을 맡기며 벌어지는 소동", { x, y: 2.95, w: 7.5, h: 0.5, fontSize: 15, color: "4A3A30" });
  bullets(s, [
    "AI 챗봇과의 대화라는 일상 소재를 캐릭터 개그로 풀어낸 시의성",
    "명화 패러디 프로필 그림으로 장면마다 새로운 볼거리 제공",
    "다리를 6개만 그리는 AI에 폭발해 말라칸으로 각성하는 시리즈 공식 유지",
    "세로 스크롤 흐름에 맞춘 컷 분할과 감정 고조",
  ], x, 3.75, 7.5, { color: "4A3A30", mark: CORAL, size: 14, gap: 0.62 });
  rect(s, x, 6.35, 7.6, 0.06, CORAL, { tr: 40 });
  t(s, "\"나는…… 자연인(산 문어)이다~!!\"", { x, y: 6.45, w: 7.6, h: 0.45, fontSize: 15, bold: true, color: CORAL_D });
}

// =====================================================================
// CHAPTER 03 벅뻑 행복스크래쳐
const C3 = "03  벅뻑 행복스크래쳐";
pres.addSection({ title: "03 벅뻑 행복스크래쳐" });
chapter("03", KRAFT, "벅뻑\n행복스크래쳐", "AI Short-form Ad", "AI 숏폼 광고 / 60초\n고양이 스크래처 브랜드 광고", "03 벅뻑 행복스크래쳐", (s) => {
  full(s, "ch3_bg", "스크래처 옆 고양이 장면을 흐리게 깐 배경");
  img(s, "pan_thumb", 5.0, 0.95, 6.9, 3.7, { border: "FFFFFF", shadow: true, alt: "엄지를 치켜든 고양이와 스크래처" });
  img(s, "pan_product", 8.35, 4.0, 4.4, 2.43, { border: "FFFFFF", shadow: true, alt: "벅뻑 행복 고양이 스크래처 제품 풀샷" });
});

// 광고 기획: 사운드 반전 구조
{
  const s = slide("LIGHT", C3, KRAFT_D, "소리로 기억되는 브랜드, 사운드 반전 구조 설계", "03 벅뻑 행복스크래쳐", "bg_kraft");
  const seasons = [["pan_season1", "여름 빗소리"], ["pan_season2", "가을 낙엽 소리"], ["pan_season3", "겨울 눈 밟는 소리"], ["pan_season4", "봄 시냇물 소리"]];
  seasons.forEach(([k, l], i) => {
    const x = MX + i * 2.25;
    img(s, k, x, 1.85, 2.1, 1.13, { border: "FFFFFF", alt: l });
    t(s, l, { x, y: 3.15, w: 2.1, h: 0.3, fontSize: 11.5, color: "5A4636", align: "center" });
  });
  t(s, "\"이 소리가 아닙니다\"", { x: MX, y: 3.5, w: 8.85, h: 0.35, fontSize: 13, italic: true, color: KRAFT_D, align: "center" });
  s.addShape(pres.shapes.RIGHT_ARROW, { x: 9.68, y: 2.15, w: 0.5, h: 0.5, fill: { color: KRAFT_D }, line: { type: "none" }, objectName: nm("arrow") });
  img(s, "pan_scratch", 10.35, 1.85, 2.25, 1.21, { border: KRAFT_D, alt: "스크래처를 긁는 고양이" });
  t(s, "\"벅벅, 벅벅벅벅\"", { x: 10.35, y: 3.15, w: 2.25, h: 0.3, fontSize: 12, bold: true, color: KRAFT_D, align: "center" });
  rect(s, MX, 4.1, 11.9, 0.03, "D8C4AA");
  bullets(s, [
    "사계절 자연의 소리를 하나씩 부정하는 내레이션으로 궁금증 유발",
    "스크래처 긁는 소리를 브랜드명 '벅뻑'과 연결해 소리만으로 브랜드 각인",
    "고양이 클로즈업과 제품 풀샷으로 '행복 스크래처' 메시지 마무리",
  ], MX, 4.35, 11.9, { color: INKT, mark: KRAFT, size: 15, gap: 0.68 });
}

// 스토리보드
{
  const s = slide("LIGHT", C3, KRAFT_D, "씬/컷/타임 단위 스토리보드 설계", "03 벅뻑 행복스크래쳐", "bg_kraft");
  img(s, "sb1", MX, 1.75, 11.9, 5.0, { border: "FFFFFF", shadow: true, alt: "벅뻑 행복스크래쳐 스토리보드 1, 씬1~2" });
}

// 영상
{
  const s = slide("LIGHT", C3, KRAFT_D, "60초 숏폼 광고 완성본", "03 벅뻑 행복스크래쳐", "bg_kraft");
  img(s, "poster_ad", MX, 1.8, 7.6, 4.19, { border: "FFFFFF", shadow: true, url: LINK_AD, alt: "벅뻑 행복스크래쳐 광고 영상 보기" });
  t(s, "이미지를 클릭하면 영상이 재생됩니다", { x: MX, y: 6.15, w: 7.6, h: 0.3, fontSize: 11, color: MUTED });
  const x = 8.75;
  chip(s, "무대사 / 60초", x, 1.85, KRAFT, "FFFFFF", { w: 1.5 });
  t(s, "벅뻑 행복 고양이 스크래처", { x, y: 2.35, w: 3.9, h: 0.5, fontSize: 20, bold: true, color: INKT });
  bullets(s, [
    "사운드와 화면만으로 메시지를 전달한 무대사 편집",
    "따뜻한 자연광 거실 톤으로 제품의 편안한 이미지 강화",
    "엄지를 치켜든 고양이로 브랜드 호감도 극대화",
  ], x, 3.1, 3.9, { color: "4A3A30", mark: KRAFT, size: 13.5, gap: 0.82 });
  rect(s, x, 5.55, 2.45, 0.5, KRAFT_D, { round: 0.25 });
  linkText(s, "▶  영상 보기", LINK_AD, x, 5.55, 2.45, 0.5, "FFFFFF", 13);
}

// =====================================================================
// 엔딩
pres.addSection({ title: "마무리" });
{
  const s = pres.addSlide({ masterName: "PLAIN", sectionTitle: "마무리" });
  full(s, "bg_end", "창 하나만 불 켜진 숲속 오두막을 흐리게 깐 배경");
  s.addImage({ path: file("tail_end"), x: 0, y: 0, w: W, h: H, altText: "꼬리 곡선", objectName: nm("tail") });
  t(s, "THANK YOU", { x: 0, y: 2.1, w: W, h: 0.4, fontFace: NUM, fontSize: 16, italic: true, color: AMBER, charSpacing: 8, align: "center" });
  t(s, "감사합니다", { x: 0, y: 2.6, w: W, h: 1.0, fontSize: 48, bold: true, color: CREAM, align: "center" });
  t(s, "박진희", { x: 0, y: 3.75, w: W, h: 0.5, fontSize: 20, bold: true, color: CREAM, align: "center" });
  t(s, "AI 애니메이션 / 캐릭터 IP 웹툰 / 숏폼 광고", { x: 0, y: 4.3, w: W, h: 0.4, fontSize: 14, color: "C9B6A6", align: "center" });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();

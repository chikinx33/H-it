// 김경수 포트폴리오: 애니메이션 「블랙홀에 기대다」 / 화투 캐릭터 IP 「光」 / 숏폼 광고 콘티 「어른폰 십팔 Pro」
// node build.js  ->  ../김경수_포트폴리오.pptx
// 모티프: 앨범(흰 인화 테두리 + 포토 코너 + 필름 날짜 각인) 콘텐츠 면, 갠지의 발광 링(블랙홀 테두리) 챕터 간지
// 색: 보이드 블랙 / 앨범 크림 / 크라프트 브라운(포토 코너) / 갠지 링 시안 / 필름 날짜 오렌지
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const IMG = path.join(__dirname, "../work/img");
const DIMS = require("../work/dims.json");
const OUT = path.join(__dirname, "../김경수_포트폴리오.pptx");

const LINK_S01 = "https://drive.google.com/file/d/1EtM8_FgEvMgrkCFFMu-GYQKSffDVupvC/view";
const LINK_S02 = "https://drive.google.com/file/d/1PhQ_nnpA_qpdteXEd8zWOHMkSZXFxL2Q/view";

const SANS = "Malgun Gothic";
const STAMPF = "Courier New"; // 필름 카메라 날짜 각인
const VOID = "0F0E13", VOID2 = "1B1A22", PAPER = "F2EADC", WHITE = "FFFFFF", INK = "2A2420", GRAY = "6E655C";
const KRAFT = "4A3A2E", LINE = "D9CDBA", RING = "46D3E6", RING_D = "1E8FA0", STAMP = "F28C28", LED = "FFB547", MIST = "C9CED6";

const THEME = {
  name: "Album Blackhole",
  headFontFace: SANS,
  bodyFontFace: SANS,
  colors: { dk1: VOID, lt1: WHITE, dk2: KRAFT, lt2: PAPER, accent1: RING_D, accent2: STAMP, accent3: KRAFT, accent4: RING, accent5: LED, accent6: GRAY, hlink: RING_D, folHlink: RING_D },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: SANS, bodyFontFace: SANS };
pres.author = "김경수";
pres.title = "김경수 포트폴리오";
const W = 13.333, H = 7.5, MX = 0.7;

const footer = (c) => ({ text: { text: "김경수  ·  AI 콘텐츠 크리에이터 포트폴리오", options: { x: MX, y: 7.08, w: 6, h: 0.28, fontFace: SANS, fontSize: 9, color: c, margin: 0 } } });
const titlePh = (c) => ({ placeholder: { options: { name: "title", type: "title", x: MX, y: 0.78, w: 11.9, h: 0.72, fontFace: SANS, fontSize: 26, bold: true, color: c, align: "left", valign: "middle", margin: 0 }, text: "" } });
const num = (c) => ({ x: 12.03, y: 7.08, w: 0.6, h: 0.28, fontFace: STAMPF, fontSize: 10, bold: true, color: c, align: "right" });
pres.defineSlideMaster({ title: "ALBUM", background: { color: PAPER }, objects: [titlePh(INK), footer("9C9081")], slideNumber: num(STAMP) });
pres.defineSlideMaster({ title: "VOID", background: { color: VOID }, objects: [titlePh(WHITE), footer("7D7F8A")], slideNumber: num(STAMP) });
pres.defineSlideMaster({ title: "PLAIN", background: { color: VOID }, objects: [] });

// ---------- helpers ----------
let sid = 0;
const nm = (s) => `${s}-${++sid}`;
const file = (key) => path.join(IMG, `${key}.${DIMS[key][2]}`);
function fit(key, x, y, w, h) {
  const [iw, ih] = DIMS[key];
  const r = iw / ih;
  let fw = w, fh = w / r;
  if (fh > h) { fh = h; fw = h * r; }
  return { x: x + (w - fw) / 2, y: y + (h - fh) / 2, w: fw, h: fh };
}
function rect(s, x, y, w, h, fill, o = {}) {
  s.addShape(o.round ? pres.shapes.ROUNDED_RECTANGLE : pres.shapes.RECTANGLE, { x, y, w, h, rectRadius: o.round || 0, rotate: o.rot || 0, fill: { color: fill, transparency: o.tr || 0 }, line: o.line ? { color: o.line, width: o.lw || 0.75 } : { type: "none" }, shadow: o.shadow ? { type: "outer", color: "000000", opacity: 0.25, blur: 7, offset: 2.5, angle: 90 } : undefined, objectName: nm("rect") });
}
function t(s, text, o) { s.addText(text, { fontFace: SANS, color: INK, margin: 0, valign: "top", isTextBox: true, objectName: nm(o.name || "text"), ...o }); }
function stamp(s, text, x, y, w, o = {}) {
  t(s, text, { x, y, w, h: o.h || 0.3, fontFace: STAMPF, fontSize: o.size || 11, bold: true, color: o.color || STAMP, align: o.align || "left", valign: "middle", charSpacing: 1, name: "stamp" });
}
function full(s, key, alt) { s.addImage({ path: file(key), x: 0, y: 0, w: W, h: H, altText: alt || "배경", objectName: nm(key) }); }
// 앨범 포토 코너: 직각삼각형을 회전해 네 귀퉁이에 끼움
function corners(s, f, c, color) {
  const d = 0.03;
  [[f.x - d, f.y - d, 90], [f.x + f.w - c + d, f.y - d, 180], [f.x + f.w - c + d, f.y + f.h - c + d, 270], [f.x - d, f.y + f.h - c + d, 0]].forEach(([x, y, rot]) => {
    s.addShape(pres.shapes.RIGHT_TRIANGLE, { x, y, w: c, h: c, rotate: rot, fill: { color }, line: { type: "none" }, objectName: nm("photo-corner") });
  });
}
// 앨범 사진: 흰 인화 테두리 + 포토 코너 + 날짜 각인
function photo(s, key, x, y, w, h, o = {}) {
  let f = fit(key, x, y, w, h);
  if (o.top) f = { ...f, y };
  if (o.left) f = { ...f, x };
  const b = o.border == null ? (f.w > 2.5 ? 0.09 : 0.06) : o.border;
  if (b) rect(s, f.x - b, f.y - b, f.w + 2 * b, f.h + 2 * b, WHITE, { shadow: o.shadow !== false });
  const link = o.url ? { hyperlink: { url: o.url, tooltip: "영상 보기" } } : {};
  s.addImage({ path: file(key), ...f, altText: o.alt || key, objectName: nm(`img-${key}`), ...link });
  if (o.corners !== false) corners(s, { x: f.x - b, y: f.y - b, w: f.w + 2 * b, h: f.h + 2 * b }, o.c || (f.w > 2.5 ? 0.3 : 0.2), o.cc || KRAFT);
  if (o.stamp) stamp(s, o.stamp, f.x + f.w - 2.2, f.y + f.h - 0.36, 2.1, { align: "right", size: o.ss || 11 });
  return f;
}
// 갠지 링 아이콘 (라벨 / 글머리)
function ringDot(s, x, y, d, color, lw) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { type: "none" }, line: { color, width: lw || 1.75 }, objectName: nm("ring-dot") });
}
function label(s, text, dark) {
  ringDot(s, MX, 0.43, 0.22, dark ? RING : RING_D, 2);
  t(s, text, { x: MX + 0.34, y: 0.36, w: 9, h: 0.36, fontSize: 11.5, bold: true, color: dark ? MIST : GRAY, valign: "middle", charSpacing: 1, name: "label" });
}
function slide(master, lab, title, section, bgKey) {
  const s = pres.addSlide({ masterName: master, sectionTitle: section });
  if (bgKey) full(s, bgKey, "작품 장면을 흐리게 깐 배경");
  label(s, lab, master === "VOID");
  s.addText(title, { placeholder: "title" });
  return s;
}
function bullets(s, items, x, y, w, o = {}) {
  const color = o.color || INK, mark = o.mark || RING_D, size = o.size || 14, gap = o.gap || 0.62;
  items.forEach((it, i) => {
    ringDot(s, x, y + i * gap + size / 72 * 0.5, 0.15, mark, 1.75);
    t(s, it, { x: x + 0.32, y: y + i * gap, w: w - 0.32, h: gap, fontSize: size, color, lineSpacingMultiple: 1.12, name: "bullet" });
  });
}
function linkBtn(s, text, url, x, y, w) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.48, rectRadius: 0.24, fill: { color: VOID2 }, line: { color: RING, width: 1.5 }, objectName: nm("link-btn") });
  s.addText([{ text, options: { hyperlink: { url, tooltip: url }, color: WHITE, bold: true } }], { x, y, w, h: 0.48, fontFace: SANS, fontSize: 12.5, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: nm("link") });
}
// 발광 링 + 원형 장면 (블랙홀 테두리)
function ringImage(s, cirKey, x, y, size, alt) {
  const d = size * 0.8, off = (size - d) / 2;
  s.addImage({ path: file(cirKey), x: x + off, y: y + off, w: d, h: d, altText: alt, objectName: nm(cirKey) });
  s.addImage({ path: file("ring"), x, y, w: size, h: size, altText: "갠지의 발광 링", objectName: nm("ring") });
}
// 챕터 간지: 큰 발광 링 속 대표 장면 + 필름 각인 번호
function chapter(n, cirKey, bgKey, title, sub, meta, section, alt) {
  const s = pres.addSlide({ masterName: "PLAIN", sectionTitle: section });
  full(s, bgKey, "챕터 대표 장면을 흐리고 어둡게 깐 배경");
  ringImage(s, cirKey, 0.75, 0.85, 5.8, alt);
  stamp(s, `No.0${n}`, 7.1, 1.55, 3, { size: 26, h: 0.6 });
  t(s, "CHAPTER", { x: 7.12, y: 2.2, w: 3, h: 0.3, fontSize: 11, bold: true, color: RING, charSpacing: 8 });
  t(s, title, { x: 7.1, y: 2.6, w: 5.8, h: 1.75, fontSize: 38, bold: true, color: WHITE, valign: "top", lineSpacingMultiple: 1.05 });
  s.addShape(pres.shapes.LINE, { x: 7.12, y: 4.5, w: 1.1, h: 0, line: { color: RING, width: 2.5 }, objectName: nm("rule") });
  t(s, sub, { x: 7.1, y: 4.75, w: 5.6, h: 1.0, fontSize: 16, color: "E3E6EB", lineSpacingMultiple: 1.25 });
  t(s, meta, { x: 7.1, y: 5.85, w: 5.6, h: 0.35, fontSize: 12, color: "9CA3AE" });
  return s;
}

// =====================================================================
// 01 표지
pres.addSection({ title: "소개" });
{
  const s = pres.addSlide({ masterName: "PLAIN", sectionTitle: "소개" });
  full(s, "bg_cover", "아빠와 간호사를 올려다보는 신생아 시점 장면을 흐리게 깐 배경");
  ringImage(s, "cir_cover", 6.75, 0.55, 6.4, "어린 경수와 갠지가 엄지를 세우는 장면");
  stamp(s, "AI CONTENT CREATOR PORTFOLIO", MX, 1.5, 5, { size: 12.5 });
  t(s, "김경수", { x: MX, y: 1.9, w: 6, h: 1.2, fontSize: 60, bold: true, color: WHITE, valign: "middle" });
  t(s, "AI 시대의 관계와 기억을 장면으로 설계하는 스토리텔러", { x: MX, y: 3.2, w: 6.2, h: 0.5, fontSize: 17, color: "E3E6EB" });
  [["No.01", "애니메이션 「블랙홀에 기대다」"], ["No.02", "화투 캐릭터 IP 「光」"], ["No.03", "숏폼 광고 콘티 「어른폰 십팔 Pro」"]].forEach(([c, x], i) => {
    const y = 4.2 + i * 0.55;
    stamp(s, c, MX, y + 0.06, 0.9, { size: 13 });
    t(s, x, { x: MX + 1.0, y, w: 5.3, h: 0.42, fontSize: 15, color: WHITE, valign: "middle" });
  });
  t(s, "서울시 매력일자리 AI 콘텐츠 크리에이터 양성과정  ·  2026", { x: MX, y: 6.7, w: 6.2, h: 0.3, fontSize: 10, color: "9CA3AE" });
}

// 02 목차: 앨범 한 면
{
  const s = slide("ALBUM", "CONTENTS", "목차", "소개", "bg_album");
  const rows = [
    ["s07_sun", "No.01", "블랙홀에 기대다", "필로소피컬 SF 드라마 애니메이션\n시나리오 / 캐릭터 시트 / 스틸 / 1차 편집본", "해변에서 노을을 바라보는 가족과 갠지"],
    ["m01", "No.02", "화투 캐릭터 IP 「光」", "화투 그림 속 동물 캐릭터\n캐릭터 시트 / 일상툰 / 이모티콘 / 굿즈", "1월 학 캐릭터 시트"],
    ["sb22", "No.03", "어른폰 십팔 Pro\n숏폼 광고 콘티", "언어유희 티저 광고\n8컷 스토리보드", "스토어 안 사과나무를 발견하는 콘티 컷"],
  ];
  rows.forEach(([k, c, title, meta, alt], i) => {
    const x = MX + 0.15 + i * 4.05;
    photo(s, k, x, 1.95, 3.55, 2.1, { alt, stamp: c, ss: 12 });
    t(s, title, { x, y: 4.45, w: 3.6, h: 0.95, fontSize: 19, bold: true, color: INK, lineSpacingMultiple: 1.05 });
    t(s, meta, { x, y: 5.45, w: 3.6, h: 0.8, fontSize: 12, color: GRAY, lineSpacingMultiple: 1.2 });
  });
}

// 03 역량: 앨범에 꽂은 메모 카드
{
  const s = slide("ALBUM", "PROFILE", "서사 기획부터 캐릭터, 광고 콘티까지 아우르는 역량", "소개", "bg_album");
  const cards = [
    ["철학적 서사 기획", "AI와 인간의 기억 비대칭을 한 사람의 일생으로 확장한 SF 드라마 기획"],
    ["무대사 영상 연출", "15씬 전체를 샷 / 구도 / 카메라 / 지문만으로 설계한 감정 중심 시나리오"],
    ["캐릭터 디자인", "AI툴을 활용한 8단계 연령 시트와 화투 동물 캐릭터 라인업으로 일관된 캐릭터 설계"],
    ["광고 콘티", "언어유희 반전 구조의 8컷 숏폼 광고 스토리보드 설계"],
  ];
  cards.forEach(([h, b], i) => {
    const x = MX + i * 3.03, y = 1.95, w = 2.83, hh = 4.4;
    rect(s, x, y, w, hh, WHITE, { shadow: true });
    corners(s, { x, y, w, h: hh }, 0.3, KRAFT);
    stamp(s, `0${i + 1}`, x + 0.3, y + 0.45, 1.2, { size: 30, h: 0.6 });
    ringDot(s, x + w - 0.62, y + 0.55, 0.36, RING_D, 2.25);
    t(s, h, { x: x + 0.3, y: y + 1.3, w: w - 0.5, h: 0.5, fontSize: 18.5, bold: true, color: INK });
    t(s, b, { x: x + 0.3, y: y + 1.95, w: w - 0.55, h: 2.2, fontSize: 13.5, color: "4E463F", lineSpacingMultiple: 1.28 });
  });
}

// =====================================================================
// CHAPTER 01 블랙홀에 기대다
const L1 = "No.01  블랙홀에 기대다", S1 = "01 블랙홀에 기대다";
pres.addSection({ title: S1 });
chapter(1, "cir_sun", "ch1_bg", "블랙홀에 기대다", "AI의 기록은 완벽하다. 기억만 없을 뿐", "필로소피컬 SF 드라마 애니메이션", S1, "노을 진 해변에서 엄마가 아빠 어깨에 기대는 뒷모습");

// 작품 개요
{
  const s = slide("ALBUM", L1, "AI와 인간의 기억 비대칭을 그린 필로소피컬 SF 드라마", S1, "bg_album");
  photo(s, "s07_sun", MX, 1.9, 5.3, 2.98, { stamp: "S#07 C#01", alt: "가족 넷이 나란히 앉아 노을을 바라보는 뒷모습" });
  t(s, "기획 단계 가제 「혼자만의 추억」", { x: MX, y: 5.1, w: 5.3, h: 0.3, fontSize: 11, color: GRAY });
  const x = 6.45;
  stamp(s, "LOGLINE", x, 1.85, 3);
  t(s, "AI의 기록은 완벽하다. 기억만 없을 뿐", { x, y: 2.15, w: 6.2, h: 0.6, fontSize: 22, bold: true, color: INK, valign: "middle" });
  const info = [["장르", "필로소피컬 SF 드라마"], ["플랫폼", "인스타그램 / 유튜브 / 광고"], ["타깃", "10대 후반~40대, AI를 일상적으로 사용하는 세대"], ["구성", "15씬, 발단 / 전개 / 위기·절정 / 결말"]];
  info.forEach(([a, b], i) => {
    const y = 2.95 + i * 0.46;
    t(s, a, { x, y, w: 1.0, h: 0.4, fontSize: 12.5, bold: true, color: RING_D, valign: "middle" });
    t(s, b, { x: x + 1.05, y, w: 5.1, h: 0.4, fontSize: 13.5, color: INK, valign: "middle" });
    s.addShape(pres.shapes.LINE, { x, y: y + 0.43, w: 6.15, h: 0, line: { color: LINE, width: 0.75 }, objectName: nm("rule") });
  });
  stamp(s, "PLANNING", x, 4.95, 3);
  bullets(s, [
    "감정의 기억은 인간에게만 쌓이고 AI는 리셋되는 비대칭 구조를 사람 사이의 관계로 확장",
    "매 순간의 진심과 그 기억이 살아가는 원동력이라는 메시지 전달",
  ], x, 5.3, 6.2, { size: 13.5, gap: 0.68 });
}

// 경수 연령 시트
{
  const s = slide("ALBUM", L1, "유아기부터 노년까지, 8단계 연령 시트로 설계한 주인공 경수", S1, "bg_album");
  const ages = [["신생아", "0세"], ["유치원생", "5~6세"], ["초등학생", "10~11세"], ["중학생", "13~15세"], ["고등학생", "16~17세"], ["대학 졸업", "졸업식"], ["30대 초반", "32~33세"], ["60대", "60~61세"]];
  ages.forEach(([a, b], i) => {
    const pw = 1.5, ph = pw * DIMS.age1[1] / DIMS.age1[0];
    const x = 0.85 + (i % 4) * 2.1, y = i < 4 ? 1.95 : 4.5;
    rect(s, x - 0.1, y - 0.1, pw + 0.2, ph + 0.62, WHITE, { shadow: true, rot: [-1.5, 1, -0.5, 1.5, 1, -1, 1.5, -1.5][i] });
    s.addImage({ path: file(`age${i + 1}`), x, y, w: pw, h: ph, altText: `${a} 경수 얼굴 클로즈업`, objectName: nm("age") });
    t(s, a, { x, y: y + ph + 0.06, w: pw, h: 0.26, fontSize: 11, bold: true, color: INK, align: "center" });
    stamp(s, b, x, y + ph + 0.3, pw, { size: 10, h: 0.2, align: "center" });
  });
  const x = 9.35;
  stamp(s, "MAIN CHARACTER", x, 1.95, 3);
  t(s, "경수", { x, y: 2.3, w: 3.3, h: 0.55, fontSize: 26, bold: true, color: INK });
  t(s, "유아기부터 노년까지 등장하는 주인공, 학생에서 회사원으로 성장", { x, y: 2.95, w: 3.3, h: 0.8, fontSize: 13, color: GRAY, lineSpacingMultiple: 1.2 });
  bullets(s, [
    "외향적인 아이에서 성숙한 내향적 성인으로 변하는 성격을 표정과 의상으로 표현",
    "연령별 전면 / 후면 / 얼굴 클로즈업 시트로 장면 간 인물 일관성 확보",
  ], x, 4.05, 3.3, { size: 13, gap: 1.2 });
}

// 갠지와 가족
{
  const s = slide("ALBUM", L1, "감정을 드러내지 않는 얼굴과 따뜻한 몸체로 설계한 AI 로봇 갠지", S1, "bg_album");
  photo(s, "ganji", MX, 1.9, 6.0, 4.5, { top: true, left: true, stamp: "GANJI", alt: "갠지 캐릭터 시트: 전면, 후면, 측면, 얼굴 클로즈업" });
  const x = 7.25;
  bullets(s, [
    "발광 링 눈과 입 없는 무표정으로 감정을 읽을 수 없는 AI 표현",
    "따뜻한 색감과 재질의 몸체로 가족 같은 친근함 부여",
  ], x, 1.85, 5.4, { size: 13.5, gap: 0.62 });
  stamp(s, "GANJI STORE LINE-UP", x, 3.15, 4);
  [["ganji3", "갠지3 아동용 소형"], ["ganji5", "갠지5 남성형"], ["ganji7", "갠지7 여성형 최상위"]].forEach(([k, c], i) => {
    const xx = x + i * 1.85;
    photo(s, k, xx, 3.5, 1.6, 1.2, { alt: c, c: 0.16 });
    t(s, c, { x: xx - 0.05, y: 4.78, w: 1.75, h: 0.26, fontSize: 10, color: GRAY, align: "center" });
  });
  stamp(s, "FAMILY", x, 5.1, 4);
  [["dad", "아버지"], ["mom", "어머니"], ["tori", "반려견 토리"]].forEach(([k, c], i) => {
    const xx = x + i * 1.85;
    photo(s, k, xx, 5.45, 1.6, 1.2, { alt: `${c} 캐릭터 시트`, c: 0.16 });
    t(s, c, { x: xx - 0.05, y: 6.73, w: 1.75, h: 0.26, fontSize: 10, color: GRAY, align: "center" });
  });
}

// 4막 구성
{
  const s = slide("ALBUM", L1, "발단 / 전개 / 위기·절정 / 결말, 4막 15씬 구성", S1, "bg_album");
  const acts = [
    ["발단", "S#01~03", "s02_thumb", "탄생과 한글 배우기, 첫 일기로 쌓는 갠지와의 유년기"],
    ["전개", "S#04~07", "s06_photo", "성장 몽타주와 제주도 가족 여행으로 쌓는 행복한 기억"],
    ["위기·절정", "S#08~12", "s09_back", "부모와의 이별 후 앨범과 일기를 거듭 확인하며 혼자만의 기억임을 자각"],
    ["결말", "S#13~15", null, "노년의 경수가 마주한 새로운 AI 광고판의 질문과 타이틀 「블랙홀에 기대다」"],
  ];
  s.addShape(pres.shapes.LINE, { x: MX + 0.12, y: 2.42, w: 11.75, h: 0, line: { color: RING_D, width: 1.5, dashType: "dash" }, objectName: nm("timeline") });
  acts.forEach(([a, sc, k, d], i) => {
    const x = MX + i * 3.05, w = 2.7;
    t(s, a, { x: x + 0.4, y: 1.85, w: 2.4, h: 0.38, fontSize: 16, bold: true, color: INK, valign: "middle" });
    rect(s, x, 2.29, 0.26, 0.26, PAPER);
    ringDot(s, x, 2.29, 0.26, RING_D, 2.25);
    if (k) photo(s, k, x, 2.8, w, 1.6, { alt: `${a} 대표 장면`, c: 0.22 });
    else {
      rect(s, x - 0.06, 2.74, w + 0.12, 1.72, WHITE, { shadow: true });
      rect(s, x, 2.8, w, 1.6, VOID);
      corners(s, { x: x - 0.06, y: 2.74, w: w + 0.12, h: 1.72 }, 0.22, KRAFT);
      t(s, "당신은 일생동안,\n혼자 살아갈 수 있습니까?", { x, y: 2.8, w, h: 1.6, fontSize: 14, bold: true, color: LED, align: "center", valign: "middle", lineSpacingMultiple: 1.25 });
    }
    stamp(s, sc, x, 4.6, 2, { size: 12 });
    t(s, d, { x, y: 4.95, w, h: 1.3, fontSize: 13, color: INK, lineSpacingMultiple: 1.22 });
  });
  ringDot(s, MX, 6.42, 0.17, STAMP, 2);
  t(s, "전 장면 대사 없이 샷 / 구도 / 카메라 / 지문만으로 설계한 시나리오", { x: MX + 0.32, y: 6.32, w: 11, h: 0.38, fontSize: 13.5, bold: true, color: KRAFT, valign: "middle" });
}

// 일기와 앨범의 복선
{
  const s = slide("ALBUM", L1, "첫 일기와 앨범 속 사진으로 심은 기억의 복선", S1, "bg_album");
  photo(s, "s03_diary", MX, 1.9, 5.15, 2.9, { stamp: "S#03 C#02", alt: "삐뚤빼뚤한 글씨로 '산타 할아버지는 없다'를 쓰는 일기 인서트" });
  photo(s, "xmas", 6.1, 1.9, 3.87, 2.9, { stamp: "ALBUM", alt: "산타 옷을 입은 아빠와 가족, 갠지의 크리스마스 앨범 사진" });
  photo(s, "s03_sleep", 10.3, 1.9, 2.33, 1.31, { stamp: "S#03 C#03", ss: 9, alt: "일기장 위에 엎드려 잠든 경수와 지켜보는 갠지" });
  photo(s, "s02_write", 10.3, 3.49, 2.33, 1.31, { stamp: "S#02 C#01", ss: 9, alt: "거실 바닥에 엎드려 한글을 쓰는 어린 경수와 갠지" });
  // 일기장 카드
  const dx = MX, dy = 5.05, dw = 6.4, dh = 1.9;
  rect(s, dx, dy, dw, dh, "FFFDF7", { shadow: true });
  s.addShape(pres.shapes.LINE, { x: dx + 0.55, y: dy, w: 0, h: dh, line: { color: "E9A3A3", width: 1 }, objectName: nm("diary-margin") });
  stamp(s, "DIARY", dx + 0.75, dy + 0.1, 2, { size: 10 });
  t(s, "\"산타 할아버지가 없다는 걸 알았다. 사실은 아빠였다. 너무 슬프고 속상했다.\"", { x: dx + 0.75, y: dy + 0.38, w: dw - 0.95, h: 0.55, fontSize: 11.5, color: INK, valign: "middle", lineSpacingMultiple: 1.1 });
  t(s, "\"갠지는 나랑 모습은 다르지만, 나한테는 형제나 다름없다. 같이 있어서 너무 행복하다.\"", { x: dx + 0.75, y: dy + 0.97, w: dw - 0.95, h: 0.55, fontSize: 11.5, color: INK, valign: "middle", lineSpacingMultiple: 1.1 });
  stamp(s, "S#12 앨범 확인 장면의 일기 인서트", dx + 0.75, dy + 1.57, 5, { size: 9.5, color: GRAY, h: 0.24 });
  bullets(s, [
    "S#03 첫 일기 '산타 할아버지는 없다'를 S#12 앨범 확인 장면에서 회수하는 복선 구조",
    "사진 속 상황과 일기 속 감정을 나란히 배치해 기록과 기억의 차이 표현",
  ], 7.45, 5.15, 5.2, { size: 13.5, gap: 0.85 });
}

// 상실의 감정선
{
  const s = slide("VOID", L1, "대사 없이 샷과 표정으로 쌓아 올린 상실의 감정선", S1, "bg_void2");
  [["s08_grad", "S#08 C#04", "갠지가 졸업 가운 차림의 경수에게 학사모를 씌워 주는 장면"], ["s09_walk", "S#09 C#02", "학사모를 쓴 경수가 아버지의 봉안함 앞에 선 장면"], ["s09_mom", "S#09 C#05", "눈물이 차오르는 엄마의 얼굴 클로즈업"], ["s09_back", "S#09 C#06", "엄마의 어깨를 토닥이는 경수와 뒤에 선 갠지의 뒷모습"]].forEach(([k, st, alt], i) => {
    photo(s, k, MX + (i % 2) * 3.97, i < 2 ? 1.95 : 4.35, 3.75, 2.11, { stamp: st, alt, cc: VOID2 });
  });
  bullets(s, [
    "졸업식의 기쁨에서 봉안당의 상실로 이어지는 감정 낙차 설계",
    "학사모 차림 그대로 아버지를 찾아가는 장면으로 부재의 의미 강조",
    "클로즈업과 뒷모습 샷의 교차로 인물의 감정 전달",
  ], 8.95, 2.0, 3.7, { size: 14, gap: 1.15, color: WHITE, mark: RING });
}

// 완급 조절
{
  const s = slide("ALBUM", L1, "무거운 서사 속 생활 유머로 완급 조절", S1, "bg_album");
  photo(s, "s11_table", MX, 1.9, 6.6, 3.71, { stamp: "S#11 C#09", alt: "카쑤 ZERO 맥주와 가솔린 ZERO가 놓인 식탁 하이앵글" });
  ringDot(s, MX, 5.95, 0.15, RING_D, 1.75);
  t(s, "엄마의 손맛을 그리워하며 벽의 가족사진으로 시선이 옮겨 가는 감정 전환", { x: MX + 0.32, y: 5.85, w: 6.3, h: 0.36, fontSize: 12.5, color: GRAY, valign: "middle" });
  photo(s, "s11_apron", 7.7, 1.9, 2.4, 1.35, { stamp: "C#06", ss: 10, alt: "앞치마를 두르고 된장찌개를 든 갠지" });
  photo(s, "s11_green", 10.25, 1.9, 2.4, 1.35, { stamp: "C#13", ss: 10, alt: "턱 아래부터 초록빛이 차오르는 경수의 얼굴" });
  bullets(s, [
    "앞치마를 두른 인물이 갠지로 드러나는 반전 컷으로 시선 집중",
    "카쑤 ZERO / 가솔린 ZERO 패러디 소품으로 세계관 위트 강화",
    "턱 아래부터 차오르는 초록빛 과장 표현으로 코믹한 리듬 형성",
  ], 7.7, 3.65, 4.95, { size: 13.5, gap: 0.85 });
}

// 수미상관과 엔딩
{
  const s = slide("VOID", L1, "수미상관 구조와 엔딩 문구 전환으로 완성한 일생의 순환", S1, "bg_void");
  [["s01_pov", "S#01 / S#13", "신생아의 첫 시점샷과 노년의 시점샷을 대구로 배치해 일생의 처음과 끝 연결", "아빠와 간호사가 내려다보는 신생아 시점샷"],
   ["s04_play", "S#04 / S#13", "갠지의 목말을 타던 놀이터에서 다른 아이의 목말 장면으로 세대의 순환 표현", "놀이터에서 갠지의 목말을 탄 어린 경수"]].forEach(([k, st, d, alt], i) => {
    const y = i ? 4.3 : 1.95;
    photo(s, k, MX, y, 3.4, 1.91, { alt, cc: VOID2 });
    stamp(s, st, 4.4, y + 0.05, 3);
    t(s, d, { x: 4.4, y: y + 0.45, w: 3.5, h: 1.3, fontSize: 13.5, color: WHITE, lineSpacingMultiple: 1.22 });
  });
  const x = 8.5, w = 4.13;
  stamp(s, "S#15 LED 전광판", x, 1.95, w);
  const led = (y, text) => {
    rect(s, x, y, w, 0.95, "07070A", { line: "3A3540", lw: 1 });
    t(s, text, { x, y, w, h: 0.95, fontSize: 14, bold: true, color: LED, align: "center", valign: "middle", lineSpacingMultiple: 1.2 });
  };
  led(2.35, "간지와 함께라면,\n당신은 혼자가 아닙니다");
  t(s, "▼", { x, y: 3.35, w, h: 0.32, fontSize: 13, color: RING, align: "center", valign: "middle" });
  led(3.72, "당신은 일생동안,\n혼자 살아갈 수 있습니까?");
  bullets(s, [
    "광고 문구 전환으로 관객에게 직접 질문을 던지는 엔딩",
    "회상(C#24-03)과 현재(C#30)를 동일 구도로 재현해 달라진 의미 강조",
  ], x, 4.95, w, { size: 12.5, gap: 0.8, color: WHITE, mark: RING });
}

// 1차 편집본
{
  const s = slide("VOID", L1, "1차 편집본 영상", S1, "bg_void");
  [["poster_s01", LINK_S01, "S#01 병원 분만실", "암전에서 페이드인되는 신생아 시점샷으로 여는 오프닝", "▶  S#01 영상 보기"],
   ["poster_s02", LINK_S02, "S#02 거실 한글 배우기", "버즈아이뷰 틸트다운과 엄지 척 투샷으로 그린 갠지와의 교감", "▶  S#02 영상 보기"]].forEach(([k, url, st, d, btn], i) => {
    const x = MX + i * 6.33;
    photo(s, k, x, 1.95, 5.6, 3.15, { url, alt: `${st} 1차 편집본 영상 보기`, cc: VOID2 });
    stamp(s, st, x, 5.3, 4, { size: 12.5 });
    t(s, d, { x, y: 5.65, w: 5.6, h: 0.4, fontSize: 13.5, color: WHITE });
    linkBtn(s, btn, url, x, 6.2, 2.5);
  });
}

// =====================================================================
// CHAPTER 02 화투 캐릭터 IP 「光」
const L2 = "No.02  화투 캐릭터 IP 「光」", S2 = "02 화투 캐릭터 IP 光";
pres.addSection({ title: S2 });
chapter(2, "cir_hak", "ch2_bg", "화투 캐릭터 IP\n「光」", "화투 열두 달 그림 속 동물로 그린 공감 일상 캐릭터", "캐릭터 시트 / 일상툰 / 이모티콘 / 굿즈", S2, "1월 학 캐릭터 시트의 정면과 측면 모습");

// IP 개요와 라인업
{
  const s = slide("ALBUM", L2, "화투 그림 속 동물을 성격 유형으로 재해석한 캐릭터 라인업", S2, "bg_ip");
  photo(s, "gwang", MX, 1.95, 3.4, 3.4, { stamp: "LOGO", alt: "붉은 원 안에 먹으로 쓴 光 브랜드 로고" });
  t(s, "브랜드 로고 「光」", { x: MX, y: 5.55, w: 3.4, h: 0.3, fontSize: 12, bold: true, color: GRAY, align: "center" });
  const x = 4.6;
  bullets(s, [
    "화투 열두 달 그림 속 동물을 현대인의 성격 유형으로 재해석",
    "흑백 손그림 라인 드로잉으로 통일한 IP 톤 앤 매너",
  ], x, 1.9, 8.0, { size: 14, gap: 0.55 });
  const chips = [["1월", "학", "지친 직장인"], ["2월", "꾀꼬리", "수다쟁이"], ["4월", "두견새", "감성 낭만파"], ["7월", "멧돼지", "돌격 행동파"],
    ["8월", "기러기 삼둥이", "사고뭉치 삼총사"], ["10월", "사슴", "칭찬 응원꾼"], ["11월", "봉황", "자뻑 과장님"], ["12월", "선비+개구리", "과묵 선비 / 장난 개구리"]];
  chips.forEach(([m, a, b], i) => {
    const cx = x + (i % 4) * 2.03, cy = i < 4 ? 3.15 : 4.85, cw = 1.88, ch = 1.5;
    rect(s, cx, cy, cw, ch, WHITE, { shadow: true });
    corners(s, { x: cx, y: cy, w: cw, h: ch }, 0.18, KRAFT);
    stamp(s, m, cx + 0.2, cy + 0.17, 1.2, { size: 15 });
    t(s, a, { x: cx + 0.2, y: cy + 0.55, w: cw - 0.3, h: 0.4, fontSize: 14.5, bold: true, color: INK, valign: "middle" });
    t(s, b, { x: cx + 0.2, y: cy + 0.97, w: cw - 0.3, h: 0.4, fontSize: 10.5, color: GRAY, valign: "middle" });
  });
}

// 대표 캐릭터 1월 학
{
  const s = slide("ALBUM", L2, "대표 캐릭터 1월 학, 배터리 8%로 출근하는 직장인", S2, "bg_ip");
  photo(s, "m01", MX, 1.9, 6.67, 5.0, { stamp: "1월 학", alt: "1월 학 캐릭터 시트: 3면도, 일상 속 학, 표정 모음, 소품" });
  const x = 7.95;
  t(s, "\"오늘도 출근하는 위대한 학…\"", { x, y: 1.95, w: 4.7, h: 0.5, fontSize: 18, bold: true, color: INK });
  const info = [["성격", "지침 / 예민 / 여림"], ["특징", "긴 목, 긴 부리, 항상 피곤함"], ["말버릇", "학~씨!"], ["좋아하는 것", "커피(필수)"]];
  info.forEach(([a, b], i) => {
    const y = 2.7 + i * 0.44;
    t(s, a, { x, y, w: 1.3, h: 0.38, fontSize: 12, bold: true, color: RING_D, valign: "middle" });
    t(s, b, { x: x + 1.35, y, w: 3.35, h: 0.38, fontSize: 13, color: INK, valign: "middle" });
    s.addShape(pres.shapes.LINE, { x, y: y + 0.41, w: 4.68, h: 0, line: { color: LINE, width: 0.75 }, objectName: nm("rule") });
  });
  bullets(s, [
    "기본 3면도 / 일상 / 표정 / 소품을 한 장에 담은 캐릭터 시트",
    "말버릇 '학~씨!'를 이모티콘과 일상툰 펀치라인으로 확장",
  ], x, 4.75, 4.7, { size: 13.5, gap: 0.85 });
}

// 캐릭터 시트 7종
{
  const s = slide("ALBUM", L2, "성격 유형별 캐릭터 시트 7종", S2, "bg_ip");
  const sheets = [["m02", "2월 꾀꼬리"], ["m04", "4월 두견새"], ["m07", "7월 멧돼지"], ["m08", "8월 기러기 삼둥이"], ["m10", "10월 사슴"], ["m11", "11월 봉황"], ["m12", "12월 선비+개구리"]];
  sheets.forEach(([k, c], i) => {
    const x = MX + (i % 4) * 3.05, y = i < 4 ? 1.9 : 4.45;
    photo(s, k, x, y, 2.85, 2.0, { alt: `${c} 캐릭터 시트`, c: 0.2 });
    stamp(s, c, x, y + 2.1, 2.85, { size: 11, align: "center", h: 0.26 });
  });
  const x = MX + 3 * 3.05, y = 4.45;
  rect(s, x, y, 2.85, 2.0, WHITE, { shadow: true });
  corners(s, { x, y, w: 2.85, h: 2.0 }, 0.2, KRAFT);
  bullets(s, [
    "캐릭터별 기본 정보 / 한마디 / 소품으로 개성 설정",
    "일상 장면과 표정 모음으로 활용 범위 확장",
  ], x + 0.22, y + 0.3, 2.5, { size: 12, gap: 0.8 });
}

// 일상툰
{
  const s = slide("ALBUM", L2, "일상툰 「네, 과장님…」으로 그린 직장인 공감 에피소드", S2, "bg_ip");
  photo(s, "toon", MX, 1.9, 3.33, 5.0, { stamp: "EP.01", alt: "일상툰 네, 과장님 1화 5컷 세로 만화" });
  const x = 4.75;
  photo(s, "toon_last", x, 1.95, 6.2, 2.38, { left: true, stamp: "LAST CUT", alt: "퇴근하지 못한 학이 옥상에서 학~씨를 외치는 마지막 컷" });
  bullets(s, [
    "자뻑 봉황 과장과 지친 학 사원의 관계로 직장 내 공감 상황 연출",
    "남 일에 관심 없는 두견새로 사무실 풍경에 현실감 부여",
    "마지막 컷 '학~씨!'로 캐릭터 말버릇을 펀치라인화",
  ], x, 4.75, 7.8, { size: 14, gap: 0.65 });
}

// 이모티콘과 굿즈
{
  const s = slide("ALBUM", L2, "움직이는 이모티콘과 굿즈로 확장한 IP", S2, "bg_ip");
  [["emo_hak", "1월 학 「학~씨!」"], ["emo_bird", "2월 꾀꼬리 「아 어딘뎅」"], ["emo_boar", "7월 멧돼지"], ["emo_deer", "10월 사슴"]].forEach(([k, c], i) => {
    const x = MX + i * 2.1;
    photo(s, k, x, 2.0, 1.85, 1.85, { alt: `${c} 움직이는 이모티콘`, c: 0.18 });
    t(s, c, { x: x - 0.1, y: 4.0, w: 2.05, h: 0.3, fontSize: 11, bold: true, color: INK, align: "center" });
  });
  stamp(s, "ANIMATED EMOTICON", MX, 1.55, 4);
  bullets(s, [
    "캐릭터 말버릇과 성격을 짧은 동작으로 압축한 움직이는 이모티콘 4종",
    "10월 사슴 캐릭터를 활용한 티셔츠 굿즈로 일상 속 IP 노출 확대",
  ], MX, 4.75, 8.2, { size: 14, gap: 0.7 });
  photo(s, "tshirt", 9.35, 1.9, 3.28, 4.1, { stamp: "GOODS", alt: "10월 사슴 캐릭터가 그려진 흰 티셔츠" });
}

// =====================================================================
// CHAPTER 03 숏폼 광고 콘티
const L3 = "No.03  어른폰 십팔 Pro 숏폼 광고 콘티", S3 = "03 어른폰 십팔 Pro";
pres.addSection({ title: S3 });
chapter(3, "cir_sb", "ch3_bg", "어른폰 십팔 Pro\n숏폼 광고 콘티", "사과를 찾아 헤매다 마주친 또 하나의 사과, 언어유희 티저 광고", "8컷 스토리보드", S3, "스토어 안 사과나무를 발견하는 콘티 컷");

{
  const s = slide("ALBUM", L3, "8컷으로 설계한 숏폼 광고 스토리보드", S3, "bg_sb");
  photo(s, "sb", MX, 1.85, 11.93, 4.9, { stamp: "STORYBOARD", alt: "숏폼 광고 스토리보드 8컷: 과일가게, 정류장, 스토어, 신제품 엔딩" });
}

{
  const s = slide("ALBUM", L3, "'사과' 언어유희로 설계한 반전 구조", S3, "bg_sb");
  const cuts = [["sb11", "S#01 C#01", "과일 사과 찾기"], ["sb22", "S#02 C#02", "사과나무 스토어 발견"], ["sb32", "S#03 C#02", "유리창 충돌 후 암전"], ["sb41", "S#04 C#01", "신제품 공개 엔딩"]];
  cuts.forEach(([k, st, c], i) => {
    const x = MX + i * 3.07;
    photo(s, k, x, 1.95, 2.7, 1.53, { alt: `${st} ${c} 콘티 컷`, c: 0.2 });
    if (i < 3) t(s, "▶", { x: x + 2.72, y: 2.5, w: 0.37, h: 0.4, fontSize: 14, color: RING_D, align: "center", valign: "middle" });
    stamp(s, st, x, 3.65, 2.7, { size: 11 });
    t(s, c, { x, y: 3.95, w: 2.7, h: 0.36, fontSize: 14, bold: true, color: INK });
  });
  bullets(s, [
    "과일 '사과'와 스마트폰 '사과'를 겹친 언어유희로 제품 공개 직전까지 궁금증 유지",
    "실사 배경 사진에 스틱 피겨를 합성한 콘티로 장소와 동선을 빠르게 시각화",
    "'커밍쑨 어른폰 십팔 Pro' 엔딩 카피로 신제품 티저 완성",
  ], MX, 4.75, 11.9, { size: 14.5, gap: 0.62 });
}

// =====================================================================
// 엔딩: 발광 링 속 인사
pres.addSection({ title: "마무리" });
{
  const s = pres.addSlide({ masterName: "PLAIN", sectionTitle: "마무리" });
  full(s, "bg_void", "노을 진 해변 장면을 흐리고 어둡게 깐 배경");
  const size = 5.6, x = (W - size) / 2, y = 0.75;
  s.addShape(pres.shapes.OVAL, { x: x + size * 0.1, y: y + size * 0.1, w: size * 0.8, h: size * 0.8, fill: { color: "07070A" }, line: { type: "none" }, objectName: nm("event-horizon") });
  s.addImage({ path: file("ring"), x, y, w: size, h: size, altText: "갠지의 발광 링", objectName: nm("ring") });
  t(s, "감사합니다", { x, y: y + 1.95, w: size, h: 0.8, fontSize: 38, bold: true, color: WHITE, align: "center", valign: "middle" });
  t(s, "김경수", { x, y: y + 2.8, w: size, h: 0.45, fontSize: 19, bold: true, color: RING, align: "center", valign: "middle" });
  t(s, "애니메이션 / 캐릭터 IP / 숏폼 광고", { x, y: y + 3.3, w: size, h: 0.35, fontSize: 12, color: "9CA3AE", align: "center", valign: "middle" });
  stamp(s, "PORTFOLIO 2026", 0, 6.7, W, { align: "center", size: 12 });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();

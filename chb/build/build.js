// 최영빈 포트폴리오: 오피스 로맨틱 코미디 애니메이션 / 캐릭터 IP 10MOO / 칸쵸 TVC 콘티
// node build.js  ->  ../최영빈_포트폴리오.pptx
// 오피스 문구 모티프: 사원증(목걸이 줄 + 사진 + 바코드) 챕터 간지, 형광펜 라벨, 테이프 붙인 자료
// 색: 차콜 / 종이 / 형광펜 옐로 / 사원증 줄 블루(10MOO 목걸이와 같은 파랑) / 칸쵸 오렌지
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const IMG = path.join(__dirname, "../work/img");
const DIMS = require("../work/dims.json");
const OUT = path.join(__dirname, "../최영빈_포트폴리오.pptx");

const LINK_EP1 = "https://drive.google.com/file/d/1cof24EAW0ZLXOta-U4wuyKgOXh34fc9D/view";
const LINK_AD = "https://drive.google.com/file/d/1Q3nT1ksGNVQFKc4LVaPQIXms32olCPhb/view";

const SANS = "Malgun Gothic";
const CHAR = "1E222C", CHAR2 = "2A303D", PAPER = "F7F5EF", WHITE = "FFFFFF", INK = "23262E", GRAY = "6B6F78";
const HL = "FFE14D", HL_D = "C9A800", BLUE = "2B59C3", BLUE_L = "8FB0F2", ORANGE = "E8892B", ORANGE_D = "B8641A", LINE = "E2DED3";

const THEME = {
  name: "Office Badge",
  headFontFace: SANS,
  bodyFontFace: SANS,
  colors: { dk1: CHAR, lt1: WHITE, dk2: CHAR2, lt2: PAPER, accent1: BLUE, accent2: HL, accent3: ORANGE, accent4: BLUE_L, accent5: HL_D, accent6: GRAY, hlink: BLUE, folHlink: BLUE },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: SANS, bodyFontFace: SANS };
pres.author = "최영빈";
pres.title = "최영빈 포트폴리오";
const W = 13.333, H = 7.5, MX = 0.7;

const footer = (c) => ({ text: { text: "최영빈  ·  AI 콘텐츠 크리에이터 포트폴리오", options: { x: MX, y: 7.08, w: 6, h: 0.28, fontFace: SANS, fontSize: 9, color: c, margin: 0 } } });
const titlePh = (c) => ({ placeholder: { options: { name: "title", type: "title", x: MX, y: 0.8, w: 11.9, h: 0.7, fontFace: SANS, fontSize: 27, bold: true, color: c, align: "left", valign: "middle", margin: 0 }, text: "" } });
const num = (c) => ({ x: 12.03, y: 7.08, w: 0.6, h: 0.28, fontFace: SANS, fontSize: 9, bold: true, color: c, align: "right" });
pres.defineSlideMaster({ title: "PAPER", background: { color: PAPER }, objects: [titlePh(INK), footer("A09B8E")], slideNumber: num("A09B8E") });
pres.defineSlideMaster({ title: "DESK", background: { color: CHAR }, objects: [titlePh(WHITE), footer("7C8291")], slideNumber: num("7C8291") });
pres.defineSlideMaster({ title: "PLAIN", background: { color: CHAR }, objects: [] });

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
  s.addShape(o.round ? pres.shapes.ROUNDED_RECTANGLE : pres.shapes.RECTANGLE, { x, y, w, h, rectRadius: o.round || 0, rotate: o.rot || 0, fill: { color: fill, transparency: o.tr || 0 }, line: o.line ? { color: o.line, width: o.lw || 0.75 } : { type: "none" }, shadow: o.shadow ? { type: "outer", color: "000000", opacity: 0.22, blur: 6, offset: 2, angle: 90 } : undefined, objectName: nm("rect") });
}
function t(s, text, o) { s.addText(text, { fontFace: SANS, color: INK, margin: 0, valign: "top", isTextBox: true, objectName: nm(o.name || "text"), ...o }); }
function full(s, key, alt) { s.addImage({ path: file(key), x: 0, y: 0, w: W, h: H, altText: alt || "배경", objectName: nm(key) }); }
// 자료 사진: 흰 종이 + 위쪽 테이프
function img(s, key, x, y, w, h, o = {}) {
  let f = fit(key, x, y, w, h);
  if (o.top) f = { ...f, y };
  if (o.left) f = { ...f, x };
  if (o.paper !== false) rect(s, f.x - 0.07, f.y - 0.07, f.w + 0.14, f.h + 0.14, WHITE, { shadow: true });
  const link = o.url ? { hyperlink: { url: o.url, tooltip: "영상 보기" } } : {};
  s.addImage({ path: file(key), ...f, altText: o.alt || key, objectName: nm(`img-${key}`), ...link });
  if (o.tape) rect(s, f.x + f.w / 2 - 0.45, f.y - 0.17, 0.9, 0.26, o.tape === true ? HL : o.tape, { tr: 25, rot: -4 });
  return f;
}
// 형광펜 라벨
function marker(s, text, x, y, color, o = {}) {
  const w = o.w || (0.3 + text.length * 0.11);
  rect(s, x - 0.05, y + 0.1, w, 0.22, color, { tr: o.tr == null ? 15 : o.tr });
  t(s, text, { x, y, w: w + 1, h: 0.36, fontSize: o.size || 11.5, bold: true, color: o.color || INK, valign: "middle", charSpacing: 1, name: "marker" });
}
function slide(master, label, color, title, section, bgKey) {
  const s = pres.addSlide({ masterName: master, sectionTitle: section });
  if (bgKey) full(s, bgKey);
  marker(s, label, MX + 0.05, 0.36, color, { color: master === "DESK" && color !== HL ? WHITE : INK });
  s.addText(title, { placeholder: "title" });
  return s;
}
function bullets(s, items, x, y, w, o = {}) {
  const color = o.color || INK, mark = o.mark || BLUE, size = o.size || 15, gap = o.gap || 0.62;
  items.forEach((it, i) => {
    rect(s, x, y + i * gap + size / 72 * 0.55, 0.14, 0.14, mark, { round: 0.03 });
    t(s, it, { x: x + 0.32, y: y + i * gap, w: w - 0.32, h: gap, fontSize: size, color, lineSpacingMultiple: 1.1, name: "bullet" });
  });
}
// 포스트잇 메모
function postit(s, x, y, w, h, fill, o = {}) {
  rect(s, x, y, w, h, fill, { shadow: true, rot: o.rot || 0 });
}
function linkBtn(s, label, url, x, y, w, fill) {
  rect(s, x, y, w, 0.5, fill, { round: 0.08 });
  s.addText([{ text: label, options: { hyperlink: { url, tooltip: url }, color: WHITE, bold: true } }], { x, y, w, h: 0.5, fontFace: SANS, fontSize: 13, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: nm("link") });
}
// 사원증: 목걸이 줄 + 카드 + 사진 + 바코드
function badge(s, x, y, color, photo, no, name, dept, o = {}) {
  const cw = o.w || 3.7, ch = cw * 1.42;
  const cx = x + cw / 2;
  rect(s, cx - 0.62, 0, 0.42, y + 0.25, color, { rot: -6 });
  rect(s, cx + 0.2, 0, 0.42, y + 0.25, color, { rot: 6 });
  rect(s, cx - 0.32, y - 0.25, 0.64, 0.45, "AEB3BD", { round: 0.08 });
  rect(s, x, y, cw, ch, WHITE, { round: 0.14, shadow: true });
  rect(s, x, y, cw, 0.62, color, { round: 0.14 });
  rect(s, x, y + 0.36, cw, 0.26, color);
  t(s, o.head || "AI CONTENT CREATOR", { x, y: y + 0.08, w: cw, h: 0.46, fontSize: 11, bold: true, color: color === HL ? INK : WHITE, align: "center", valign: "middle", charSpacing: 3 });
  const pw = cw * 0.5, ph = pw * 4 / 3;
  rect(s, cx - pw / 2 - 0.04, y + 0.85 - 0.04, pw + 0.08, ph + 0.08, LINE);
  s.addImage({ path: file(photo), x: cx - pw / 2, y: y + 0.85, w: pw, h: ph, altText: o.alt || "사원증 사진", objectName: nm("badge-photo") });
  const ty = y + 0.95 + ph;
  t(s, no, { x, y: ty, w: cw, h: 0.3, fontSize: 10.5, bold: true, color: GRAY, align: "center", charSpacing: 2 });
  t(s, name, { x: x + 0.15, y: ty + 0.3, w: cw - 0.3, h: o.nameH || 0.62, fontSize: o.nameSize || 21, bold: true, color: INK, align: "center", valign: "middle" });
  t(s, dept, { x: x + 0.15, y: ty + 0.3 + (o.nameH || 0.62), w: cw - 0.3, h: 0.3, fontSize: 10.5, color: GRAY, align: "center" });
  const bw = cw - 0.9, bh = bw * DIMS.barcode[1] / DIMS.barcode[0];
  s.addImage({ path: file("barcode"), x: x + 0.45, y: y + ch - 0.18 - bh, w: bw, h: bh, altText: "바코드", objectName: nm("barcode") });
}
// 챕터 간지: 사원증 + 큰 제목
function chapter(n, color, photo, bgKey, title, dept, sub, section, o = {}) {
  const s = pres.addSlide({ masterName: "PLAIN", sectionTitle: section });
  full(s, bgKey, "챕터 대표 장면을 흐리게 깐 배경");
  badge(s, 1.35, 1.25, color, photo, `EMPLOYEE No. CH-0${n}`, o.badgeName || title.replace("\n", " "), dept, { nameSize: o.nameSize || 17, nameH: 0.7, alt: o.alt });
  t(s, `CHAPTER 0${n}`, { x: 6.2, y: 1.7, w: 6, h: 0.4, fontSize: 15, bold: true, color: color === HL ? HL : (color === BLUE ? BLUE_L : "FFC48A"), charSpacing: 6 });
  rect(s, 6.2, 2.3, 0.9, 0.08, color === HL ? HL : (color === BLUE ? BLUE_L : ORANGE));
  t(s, title, { x: 6.2, y: 2.6, w: 6.6, h: 2.0, fontSize: 40, bold: true, color: WHITE, valign: "top", lineSpacingMultiple: 1.05 });
  t(s, sub, { x: 6.2, y: 4.85, w: 6.4, h: 1.0, fontSize: 15, color: "E4E6EC", lineSpacingMultiple: 1.25 });
  return s;
}

// =====================================================================
// 01 표지
pres.addSection({ title: "소개" });
{
  const s = pres.addSlide({ masterName: "PLAIN", sectionTitle: "소개" });
  full(s, "bg_cover", "홍보대행사 회의실을 흐리게 깐 배경");
  badge(s, 8.7, 1.3, BLUE, "id0", "PORTFOLIO 2026", "최영빈", "AI 콘텐츠 크리에이터", { alt: "10MOO 사원증 사진", nameSize: 26 });
  marker(s, "AI CONTENT CREATOR PORTFOLIO", MX + 0.05, 1.45, HL, { w: 3.8, size: 12 });
  t(s, "최영빈", { x: MX, y: 1.95, w: 6.5, h: 1.2, fontSize: 60, bold: true, color: WHITE, valign: "middle" });
  t(s, "직장인의 공감과 판타지를 장르로 엮는 스토리텔러", { x: MX, y: 3.25, w: 7.2, h: 0.5, fontSize: 18, color: "E4E6EC" });
  [["CH.01", HL, "오피스 로맨틱 코미디 「신입사원이지만 아이돌입니다」"], ["CH.02", BLUE_L, "캐릭터 IP 「10MOO」"], ["CH.03", "FFC48A", "칸쵸 × 포켓몬 TVC 콘티 「잡았다! 피카츄」"]].forEach(([c, col, x], i) => {
    const y = 4.25 + i * 0.55;
    t(s, c, { x: MX, y, w: 0.95, h: 0.42, fontSize: 13, bold: true, color: col, valign: "middle" });
    t(s, x, { x: MX + 1.0, y, w: 6.8, h: 0.42, fontSize: 14.5, color: WHITE, valign: "middle" });
  });
  t(s, "서울시 매력일자리 AI 콘텐츠 크리에이터 양성과정  ·  2026", { x: MX, y: 6.7, w: 6.5, h: 0.3, fontSize: 10, color: "A3A8B4" });
}

// 02 목차 (포스트잇)
{
  const s = slide("PAPER", "CONTENTS", HL, "목차", "소개", "bg_office");
  const rows = [
    ["c02", "CH.01", HL, "신입사원이지만\n아이돌입니다", "오피스 로맨틱 코미디 / 아이돌 성장물", -2],
    ["moo_04master", "CH.02", "CFE0FF", "10MOO", "캐릭터 IP / 캐릭터 시트 / 이모티콘", 1.5],
    ["sb_c8", "CH.03", "FFD9B3", "칸쵸 × 포켓몬\n「잡았다! 피카츄」", "TVC 콘티 / 숏폼 광고", -1],
  ];
  rows.forEach(([k, c, col, title, meta, rot], i) => {
    const x = MX + 0.1 + i * 4.05;
    postit(s, x, 1.85, 3.65, 4.75, col, { rot });
    rect(s, x + 1.45, 1.72, 0.75, 0.26, "FFFFFF", { tr: 35, rot: rot * 2 });
    t(s, c, { x: x + 0.3, y: 2.05, w: 2, h: 0.4, fontSize: 15, bold: true, color: INK });
    img(s, k.startsWith("moo") ? "10moo_04master" : k, x + 0.3, 2.5, 3.05, 1.75, { paper: false, alt: title });
    t(s, title, { x: x + 0.3, y: 4.4, w: 3.1, h: 1.0, fontSize: 19, bold: true, color: INK, lineSpacingMultiple: 1.05 });
    t(s, meta, { x: x + 0.3, y: 5.55, w: 3.1, h: 0.7, fontSize: 12, color: "4A4D55" });
  });
}

// 03 역량
{
  const s = slide("PAPER", "PROFILE", HL, "기획부터 프로덕션 스크립트까지 설계하는 역량", "소개", "bg_office");
  const cards = [
    ["IP 시리즈 기획", "100화 분량의 웹 애니메이션 세계관과 캐릭터 관계도 설계", HL],
    ["프로덕션 스크립트", "샷 / 구도 / 카메라 / 지문 / 대사 규칙으로 AI 생성용 콘티 표준화", BLUE_L],
    ["캐릭터 디자인", "5패널 턴어라운드와 의상 변형 시트로 캐릭터 일관성 확보", "CFE0FF"],
    ["광고 콘티", "제품과 IP를 결합한 TVC 콘티로 컷별 리듬과 카피 설계", "FFD9B3"],
  ];
  cards.forEach(([h, b, col], i) => {
    const x = MX + i * 3.0, y = 1.95;
    rect(s, x, y, 2.8, 4.4, WHITE, { shadow: true });
    rect(s, x, y, 2.8, 0.5, col);
    t(s, `0${i + 1}`, { x: x + 0.25, y: y + 0.07, w: 1, h: 0.36, fontSize: 15, bold: true, color: INK, valign: "middle" });
    t(s, h, { x: x + 0.25, y: y + 0.8, w: 2.35, h: 0.5, fontSize: 19, bold: true, color: INK });
    t(s, b, { x: x + 0.25, y: y + 1.45, w: 2.3, h: 2.6, fontSize: 14, color: "4A4D55", lineSpacingMultiple: 1.25 });
  });
}

// =====================================================================
// CHAPTER 01
const L1 = "CH.01  신입사원이지만 아이돌입니다", S1 = "01 신입사원이지만 아이돌입니다";
pres.addSection({ title: S1 });
chapter(1, HL, "id1", "ch1_bg", "신입사원이지만\n아이돌입니다", "마케팅팀 / 신입사원", "\"내가 해도 쟤네보단 낫겠다\"\n술 취해 뱉은 한마디가 현실이 된 오피스 로맨틱 코미디", S1, { badgeName: "나신입", nameSize: 22, alt: "나신입 사원증 사진" });

// 작품 개요
{
  const s = slide("PAPER", L1, HL, "회당 2분, 100화로 설계한 숏폼 웹 애니메이션", S1, "bg_office");
  img(s, "still_all", MX, 1.85, 3.25, 4.9, { tape: true, alt: "1화 S#2 스틸컷 모음, 9:16 세로 화면" });
  const x = 4.45;
  marker(s, "로그라인", x, 1.85, HL);
  t(s, "\"내가 해도 쟤네보단 낫겠다\" 술 취해 뱉은 한마디가 현실이 되었다! 여자 아이돌 홍보를 맡은 마케터 나신입, 까칠한 광고주 강도한 실장이 알고 보니 재벌 3세에 우리 회사 새 대표님?!", { x, y: 2.3, w: 8.15, h: 1.25, fontSize: 15, bold: true, color: INK, lineSpacingMultiple: 1.2 });
  const info = [["장르", "오피스 로맨틱 코미디 / 아이돌 성장물"], ["타깃", "2030 여성 (공감 / 대리만족 / 설렘 포인트)"], ["플랫폼", "유튜브 본편 / 틱톡 · 릴스 · 숏츠 예고편 및 홍보"], ["러닝타임", "회당 약 2분 내외, 총 100화"], ["화면", "9:16 세로 / 일본 2D 셀 애니메이션 룩"]];
  info.forEach(([a, b], i) => {
    const y = 3.8 + i * 0.58;
    rect(s, x, y, 8.15, 0.5, WHITE, { line: LINE });
    rect(s, x, y, 0.08, 0.5, i % 2 ? BLUE : HL);
    t(s, a, { x: x + 0.25, y, w: 1.3, h: 0.5, fontSize: 12.5, bold: true, color: GRAY, valign: "middle" });
    t(s, b, { x: x + 1.6, y, w: 6.4, h: 0.5, fontSize: 13.5, color: INK, valign: "middle" });
  });
}

// 기획의도: 갭
{
  const s = slide("DESK", L1, HL, "사무실과 무대 사이의 '갭(Gap)'이 만드는 재미", S1, "bg_night");
  postit(s, MX, 1.9, 6.0, 2.3, HL, { rot: -1.5 });
  t(s, "\"내가 해도 더 잘하겠다\"는 흔한 생각을\n진짜로 밀어붙인 이야기", { x: MX + 0.35, y: 2.15, w: 5.4, h: 1.2, fontSize: 21, bold: true, color: INK, lineSpacingMultiple: 1.15 });
  t(s, "작가 기획의도 중", { x: MX + 0.35, y: 3.55, w: 4, h: 0.3, fontSize: 11, color: "6B5A00" });
  bullets(s, [
    "반복되는 하루를 보내는 시청자에게 \"나도 저랬으면\" 하는 대리만족 제공",
    "까칠한 광고주의 숨은 호감과 재벌 3세 반전으로 현실 공감과 설렘을 동시에 축적",
    "눈치 보는 신입이 무대 위 아이돌로 바뀌는 갭으로 오피스물 / 아이돌물 두 팬층 공략",
  ], 7.1, 1.95, 5.5, { color: WHITE, mark: HL, size: 14.5, gap: 0.95 });
  img(s, "c11", MX, 4.65, 6.0, 1.92, { paper: false, alt: "분한 표정으로 각오를 다지는 나신입" });
  rect(s, MX, 4.65, 6.0, 1.92, "000000", { tr: 100, line: "3C4250" });
  t(s, "S#2 C#11  분한 표정의 나신입", { x: 7.1, y: 6.2, w: 5.5, h: 0.3, fontSize: 11, color: "A3A8B4" });
}

// 메인 캐릭터
{
  const s = slide("PAPER", L1, HL, "메인 캐릭터: 허당 신입과 까칠한 재벌 3세", S1, "bg_office");
  const ch = [
    ["ch_na", "나신입", "24세 / 홍보대행사 마케팅팀 신입사원", "아이돌 덕후라 이론은 빠삭하지만 몸은 뚝딱이. 데뷔 후 센터 / 입덕요정", HL],
    ["ch_kang", "강도한", "32세 / 강엔터 실장, 강그룹 재벌 3세", "차도남 + 집착남. 괴롭히면서 즐거워하는 초딩미, 고급 손목시계", BLUE],
  ];
  ch.forEach(([k, n, role, d, col], i) => {
    const x = MX + i * 6.05;
    img(s, k, x, 1.85, 5.8, 3.35, { tape: col === HL ? true : "BFD3FF", alt: `${n} 5패널 턴어라운드` });
    t(s, n, { x, y: 5.4, w: 2, h: 0.45, fontSize: 21, bold: true, color: INK });
    marker(s, role, x + 1.35, 5.45, col, { size: 11, tr: col === HL ? 15 : 60, w: 4.3 });
    t(s, d, { x, y: 5.95, w: 5.8, h: 0.7, fontSize: 12.5, color: "4A4D55", lineSpacingMultiple: 1.2 });
  });
}

// 서브 캐릭터
{
  const s = slide("PAPER", L1, HL, "회사 동료에서 아이돌 멤버로, 직급과 포지션의 이중 설정", S1, "bg_office");
  const ch = [
    ["ch_park", "박과장", "마케팅팀 팀장", "메인보컬 & 리더", "회식 노래방 고인물, 자꾸 나오는 트로트 뽕끼"],
    ["ch_lee", "이대리", "디자인팀 디자이너", "비주얼", "시크 도도 칼퇴 요정, 안경 벗으면 여신"],
    ["ch_kim", "김인턴", "마케팅팀 인턴", "메인래퍼", "MZ력 만렙 팩트 폭격기, 그 깡으로 하는 랩"],
  ];
  ch.forEach(([k, n, job, pos, d], i) => {
    const x = MX + i * 4.05;
    img(s, k, x, 1.85, 3.8, 2.53, { alt: `${n} 5패널 턴어라운드` });
    t(s, n, { x, y: 4.6, w: 1.5, h: 0.42, fontSize: 19, bold: true, color: INK });
    t(s, job, { x: x + 1.3, y: 4.66, w: 2.5, h: 0.32, fontSize: 11.5, color: GRAY });
    marker(s, `데뷔 후  ${pos}`, x, 5.1, HL, { size: 11.5, w: 2.4 });
    t(s, d, { x, y: 5.55, w: 3.8, h: 0.7, fontSize: 12.5, color: "4A4D55" });
  });
  t(s, "직장 내 캐릭터성을 무대 위 포지션으로 연결해 성장물의 재미 확보", { x: MX, y: 6.45, w: 11.9, h: 0.35, fontSize: 13, bold: true, color: BLUE });
}

// 배경 시트
{
  const s = slide("DESK", L1, HL, "두 회의실의 스케일 대비로 표현한 갑을 관계", S1, "bg_night");
  img(s, "bg_agency_room", MX, 1.85, 5.85, 3.29, { alt: "홍보대행사 회의실 배경" });
  img(s, "bg_kang_room", 6.75, 1.85, 5.85, 3.29, { alt: "강엔터 회의실 배경" });
  marker(s, "홍보대행사 회의실", MX, 5.3, HL, { w: 2.1 });
  marker(s, "강엔터 회의실", 6.75, 5.3, BLUE_L, { w: 1.8 });
  t(s, "블라인드와 낮은 천장의 좁은 실무 공간, 나신입의 현실", { x: MX, y: 5.75, w: 5.85, h: 0.6, fontSize: 13, color: WHITE });
  t(s, "도심 야경이 펼쳐진 넓은 고층 회의실, 강도한의 권력", { x: 6.75, y: 5.75, w: 5.85, h: 0.6, fontSize: 13, color: WHITE });
}

// 1화 구성
{
  const s = slide("PAPER", L1, HL, "상상과 현실을 매치컷으로 잇는 1화 구성", S1, "bg_office");
  const sc = [
    ["S#1", "상상 속 전쟁 도시", "로봇을 몰고 강도한을 조준하는 나신입", "8컷"],
    ["S#2", "홍보대행사 회의실", "\"이게 최선입니까?\" 제안서 반려", "11컷"],
    ["S#3", "탕비실", "\"쉿! 우리 회사 밥줄이야.\"", "3컷"],
    ["S#4", "사무실 야근", "노을에서 밤으로 이어지는 시간 경과", "1컷"],
    ["S#5", "강엔터 회의실", "꼬르륵 소리와 블랙카드, 열린 결말", "18컷"],
  ];
  sc.forEach(([n, place, d, c], i) => {
    const x = MX + i * 2.42;
    rect(s, x, 1.95, 2.22, 2.6, i === 0 ? CHAR : WHITE, { shadow: true });
    t(s, n, { x: x + 0.2, y: 2.1, w: 1.2, h: 0.4, fontSize: 18, bold: true, color: i === 0 ? HL : BLUE });
    t(s, c, { x: x + 1.3, y: 2.15, w: 0.75, h: 0.35, fontSize: 11, bold: true, color: i === 0 ? "C9CDD6" : GRAY, align: "right" });
    t(s, place, { x: x + 0.2, y: 2.6, w: 1.9, h: 0.45, fontSize: 13.5, bold: true, color: i === 0 ? WHITE : INK });
    t(s, d, { x: x + 0.2, y: 3.1, w: 1.9, h: 1.3, fontSize: 11.5, color: i === 0 ? "D5D8DF" : "4A4D55", lineSpacingMultiple: 1.2 });
    if (i < 4) t(s, "▶", { x: x + 2.22, y: 3.05, w: 0.2, h: 0.3, fontSize: 9, color: GRAY, align: "center" });
  });
  rect(s, MX, 4.85, 11.9, 1.8, WHITE, { line: LINE });
  rect(s, MX, 4.85, 0.1, 1.8, HL);
  t(s, "매치컷", { x: MX + 0.35, y: 5.0, w: 3, h: 0.35, fontSize: 13, bold: true, color: HL_D });
  t(s, "미사일 폭발 섬광이 그대로 제안서를 내려치는 강도한의 손으로 이어지는 전환, 메카닉 액션의 긴장감을 오피스 코미디의 웃음으로 반전", { x: MX + 0.35, y: 5.4, w: 11.3, h: 0.9, fontSize: 14, color: INK, lineSpacingMultiple: 1.2 });
}

// 스틸컷 + 컷 기술 규칙
{
  const s = slide("DESK", L1, HL, "AI 생성을 전제로 표준화한 프로덕션 스크립트", S1, "bg_night");
  const cuts = [["c01", "C#1"], ["c03", "C#3"], ["c06", "C#6"], ["c09", "C#9"]];
  cuts.forEach(([k, c], i) => {
    const x = MX + (i % 2) * 3.05, y = 1.9 + Math.floor(i / 2) * 2.15;
    img(s, k, x, y, 2.9, 1.85, { alt: `S#2 ${c} 스틸` });
    marker(s, c, x + 0.08, y + 0.08, HL, { w: 0.6, size: 10.5 });
  });
  const x = 7.05;
  t(s, "컷 기술 규칙", { x, y: 1.9, w: 5, h: 0.4, fontSize: 15, bold: true, color: HL });
  const rules = [["샷", "샷 사이즈 + 앵글"], ["구도", "피사체 배치와 시선 방향, 이미지 생성의 핵심 항목"], ["카메라", "고정 / 팬 / 틸트 / 트래킹 / 랙 포커스"], ["지문", "화면에 실제로 보이는 것만 기술"], ["대사", "화자 + 대사, 없으면 '없음' 명시"]];
  rules.forEach(([a, b], i) => {
    const y = 2.4 + i * 0.55;
    rect(s, x, y, 5.55, 0.47, CHAR2);
    t(s, a, { x: x + 0.2, y, w: 1.0, h: 0.47, fontSize: 12.5, bold: true, color: HL, valign: "middle" });
    t(s, b, { x: x + 1.2, y, w: 4.25, h: 0.47, fontSize: 12.5, color: WHITE, valign: "middle" });
  });
  t(s, "\"인물의 심리는 행동으로 번역해서 쓴다\"는 원칙으로 AI 스틸컷 / 영상의 연기 디테일 구체화", { x, y: 5.3, w: 5.55, h: 0.8, fontSize: 13, color: "E4E6EC", lineSpacingMultiple: 1.2 });
}

// 영상
{
  const s = slide("DESK", L1, HL, "1화 가편집본", S1, "bg_night");
  img(s, "poster_ep1", MX, 1.85, 7.6, 4.28, { url: LINK_EP1, alt: "신입사원이지만 아이돌입니다 가편집본 영상 보기" });
  t(s, "이미지를 클릭하면 영상이 재생됩니다", { x: MX, y: 6.3, w: 7.6, h: 0.3, fontSize: 11, color: "A3A8B4" });
  const x = 8.75;
  marker(s, "가편집 04", x, 1.9, HL, { w: 1.2 });
  t(s, "신입사원이지만\n아이돌입니다", { x, y: 2.4, w: 3.9, h: 1.0, fontSize: 21, bold: true, color: WHITE });
  t(s, "1화  /  9:16 세로 화면", { x, y: 3.5, w: 3.9, h: 0.35, fontSize: 13, color: "C9CDD6" });
  postit(s, x, 4.1, 3.85, 1.2, HL, { rot: 1 });
  t(s, "\"그럼, 저…\"\n궁금증을 남기는 열린 엔딩으로 다음 화 시청 유도", { x: x + 0.2, y: 4.2, w: 3.5, h: 1.0, fontSize: 12.5, bold: true, color: INK, lineSpacingMultiple: 1.15 });
  linkBtn(s, "▶  영상 보기", LINK_EP1, x, 5.6, 2.45, BLUE);
}

// =====================================================================
// CHAPTER 02 10MOO
const L2 = "CH.02  10MOO", S2 = "02 10MOO";
pres.addSection({ title: S2 });
chapter(2, BLUE, "id2", "ch2_bg", "10MOO", "캐릭터 IP / 이모티콘", "양털 후드를 쓴 회색 푸들 캐릭터\n파란 목걸이의 '10' 배지를 시그니처로 한 IP", S2, { nameSize: 22, alt: "10MOO 사원증 사진" });

// 캐릭터 시트
{
  const s = slide("PAPER", L2, BLUE_L, "기본형과 양털 후드형, 두 버전의 캐릭터 시트", S2, "bg_moo");
  const keys = [["10moo_01master", "정면"], ["10moo_02side", "측면"], ["10moo_03reverse", "후면"], ["10moo_06master", "정면"], ["10moo_07side", "측면"], ["10moo_08reverse", "후면"]];
  keys.forEach(([k, v], i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = MX + col * 2.45 + (row ? 0 : 0), y = 1.85 + row * 2.45;
    img(s, k, x, y, 2.2, 2.2, { alt: `10MOO ${row ? "후드형" : "기본형"} ${v}` });
    t(s, v, { x: x + 1.55, y: y + 1.85, w: 0.6, h: 0.3, fontSize: 10.5, bold: true, color: GRAY, align: "right" });
  });
  const x = 8.35;
  marker(s, "기본형", x, 1.95, "BFD3FF", { tr: 0, w: 0.9 });
  t(s, "곱슬 털의 회색 푸들, 파란 목걸이와 '10' 배지", { x, y: 2.4, w: 4.25, h: 0.7, fontSize: 13.5, color: INK });
  marker(s, "양털 후드형", x, 3.3, "BFD3FF", { tr: 0, w: 1.4 });
  t(s, "양 얼굴 후드를 더한 변형으로 귀여움과 친근함 강조", { x, y: 3.75, w: 4.25, h: 0.7, fontSize: 13.5, color: INK });
  bullets(s, [
    "정면 / 측면 / 후면 3방향으로 형태 일관성 확보",
    "의상 변형만으로 캐릭터 라인업을 넓히는 IP 확장 구조",
  ], x, 4.85, 4.25, { mark: BLUE, size: 13, gap: 0.75 });
}

// 이모티콘
{
  const s = slide("PAPER", L2, BLUE_L, "움직이는 이모티콘 4종", S2, "bg_moo");
  const em = [["emo_tilt", "넹?"], ["emo_clasp", "넵"], ["emo_ok", "오케이"], ["emo_angry", "무러?"]];
  em.forEach(([k, n], i) => {
    const x = MX + i * 3.0;
    rect(s, x, 1.85, 2.8, 3.3, WHITE, { shadow: true });
    img(s, k, x + 0.2, 1.95, 2.4, 2.4, { paper: false, alt: `10MOO 이모티콘 ${n}` });
    t(s, n, { x, y: 4.5, w: 2.8, h: 0.45, fontSize: 17, bold: true, color: BLUE, align: "center" });
  });
  bullets(s, [
    "고개 갸웃 / 손 모으기 / 윙크 / 팔짱 등 동작 중심의 감정 표현",
    "투명 배경 360px 규격으로 메신저 이모티콘 활용 대응",
  ], MX, 5.5, 11.9, { mark: BLUE, size: 14, gap: 0.55 });
}

// =====================================================================
// CHAPTER 03 칸쵸 TVC
const L3 = "CH.03  칸쵸 × 포켓몬 TVC", S3 = "03 칸쵸 TVC 콘티";
pres.addSection({ title: S3 });
chapter(3, ORANGE, "id3", "ch3_bg", "칸쵸 × 포켓몬\n「잡았다! 피카츄」", "TVC 콘티 / 숏폼 광고", "칸쵸를 포켓볼처럼 던져\n피카츄 얼굴이 새겨진 과자를 잡는 콜라보 광고", S3, { badgeName: "잡았다! 피카츄", nameSize: 18, alt: "콘티 속 칸쵸를 든 주인공" });

// 스토리보드
{
  const s = slide("PAPER", L3, "FFC48A", "8컷, 약 15초로 설계한 TVC 콘티", S3, "bg_sb");
  img(s, "sb", MX, 1.85, 11.9, 5.04, { tape: "FFC48A", alt: "칸쵸 포켓몬 콜라보 TVC 콘티 8컷" });
}

// 광고 포인트 + 영상
{
  const s = slide("PAPER", L3, "FFC48A", "제품과 IP를 하나의 놀이로 엮은 광고 연출", S3, "bg_sb");
  const cuts = [["sb_c5", "C#5", "날아간 칸쵸가 닿는 순간 소용돌이와 함께 빨려 들어가는 피카츄"], ["sb_c7", "C#7", "\"잡았다 피카츄!\" 얼굴이 새겨진 칸쵸를 내미는 주인공"], ["sb_c8", "C#8", "다양한 포켓몬 얼굴 칸쵸가 쏟아지는 제품 엔딩"]];
  cuts.forEach(([k, c, d], i) => {
    const y = 1.85 + i * 1.62;
    img(s, k, MX, y, 2.55, 1.42, { alt: `콘티 ${c}` });
    marker(s, c, MX + 2.8, y + 0.05, "FFC48A", { tr: 0, w: 0.6 });
    t(s, d, { x: MX + 2.8, y: y + 0.45, w: 3.9, h: 0.9, fontSize: 12.5, color: INK, lineSpacingMultiple: 1.15 });
  });
  const x = 7.65;
  img(s, "poster_ad", x, 1.9, 4.95, 2.78, { url: LINK_AD, alt: "칸쵸 광고 영상 보기" });
  bullets(s, [
    "칸쵸를 포켓볼처럼 던지는 놀이로 제품과 IP를 자연스럽게 결합",
    "Tap! / Click! / Ta-da! 효과음 카피로 컷 전환 리듬 강화",
    "\"포켓몬 X 칸쵸 지금 모으러 가자!\" 내레이션으로 수집 욕구 자극",
  ], x, 4.9, 4.95, { mark: ORANGE, size: 12.5, gap: 0.6 });
  linkBtn(s, "▶  광고 영상 보기", LINK_AD, x, 6.45, 2.4, ORANGE);
}

// =====================================================================
// 엔딩
pres.addSection({ title: "마무리" });
{
  const s = pres.addSlide({ masterName: "PLAIN", sectionTitle: "마무리" });
  full(s, "bg_night", "강엔터 회의실을 흐리게 깐 배경");
  postit(s, 4.15, 1.6, 5.0, 4.3, HL, { rot: -2 });
  rect(s, 6.2, 1.45, 0.9, 0.3, WHITE, { tr: 35, rot: -6 });
  t(s, "감사합니다", { x: 4.35, y: 2.35, w: 4.6, h: 0.9, fontSize: 40, bold: true, color: INK, align: "center", valign: "middle" });
  t(s, "최영빈", { x: 4.35, y: 3.4, w: 4.6, h: 0.5, fontSize: 22, bold: true, color: INK, align: "center" });
  t(s, "오피스 로맨틱 코미디 / 캐릭터 IP / TVC 콘티", { x: 4.35, y: 4.05, w: 4.6, h: 0.4, fontSize: 13, color: "5A4E00", align: "center" });
  s.addImage({ path: file("barcode"), x: 5.4, y: 4.85, w: 2.5, h: 2.5 * DIMS.barcode[1] / DIMS.barcode[0], altText: "바코드", objectName: nm("barcode") });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();

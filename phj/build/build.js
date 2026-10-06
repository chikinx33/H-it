// 박현진 포트폴리오: 아이돌, 출발 / 너무 좋아 1000% / 일렁이는 뻄씨 / 숏폼 광고 2편
// node build.js  ->  ../박현진_포트폴리오.pptx
// 모티프: 콘서트 티켓(절취선 + 반원 노치 + 스텁) 챕터 간지와 목차, 별 참 글머리, 팬레터 봉투 엔딩
// 색: 콘서트 바이올렛 / 응원봉 라일락 / 체리 레드(팬의 체리 키링) / 편지지 크림 / 은빛
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const IMG = path.join(__dirname, "../work/img");
const DIMS = require("../work/dims.json");
const OUT = path.join(__dirname, "../박현진_포트폴리오.pptx");

const V = (id) => `https://drive.google.com/file/d/${id}/view`;
const LINK_IDOL = V("17Z0WyQogL-lNrDR5xIX4TQRBBudd0MGz");
const LINK_BBEM = V("1cN0moZzlcjhu3kQcTy5YQjEjh3OqM4XU");
const LINK_HAT = V("1FefrZd72wzog5w2L43n1-h5u9knYitG-");
const LINK_SPK = V("1s-JXzPu1EC2nhFelQepjlmBFu1Ptx7IT");
const LINK_TUM = V("1J_nYql8gBW3Kp7d5poH8K72Mv7EYwD8t");
const LINK_MAGIC = V("1KKHFpEc7DSNOmQ0AThXyHrO0jktwjwlx");
const LINK_SLIM = V("17-MGQA1YUiZOZXjLN2MosLBWYy1VwaZy");

const SANS = "Malgun Gothic";
const TK = "Trebuchet MS"; // 티켓 인쇄 글자
const VIOLET = "241640", VIOLET2 = "36235C", LILAC = "B69CFF", LILAC_D = "7A5CD6", LILAC_L = "ECE4FF";
const CHERRY = "D7263D", CREAM = "FFF6F1", PAPER = "FFFBF8", WHITE = "FFFFFF", INK = "2B2238", GRAY = "6F6680", SILVER = "C9CCD6", LINE = "E6DCEF";

const THEME = {
  name: "Concert Ticket",
  headFontFace: SANS,
  bodyFontFace: SANS,
  colors: { dk1: VIOLET, lt1: WHITE, dk2: VIOLET2, lt2: CREAM, accent1: LILAC_D, accent2: CHERRY, accent3: LILAC, accent4: SILVER, accent5: VIOLET2, accent6: GRAY, hlink: LILAC_D, folHlink: LILAC_D },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: SANS, bodyFontFace: SANS };
pres.author = "박현진";
pres.title = "박현진 포트폴리오";
const W = 13.333, H = 7.5, MX = 0.7;

const footer = (c) => ({ text: { text: "박현진  ·  AI 콘텐츠 크리에이터 포트폴리오", options: { x: MX, y: 7.08, w: 6, h: 0.28, fontFace: SANS, fontSize: 9, color: c, margin: 0 } } });
const titlePh = (c) => ({ placeholder: { options: { name: "title", type: "title", x: MX, y: 0.78, w: 11.9, h: 0.72, fontFace: SANS, fontSize: 26, bold: true, color: c, align: "left", valign: "middle", margin: 0 }, text: "" } });
const num = (c) => ({ x: 12.03, y: 7.08, w: 0.6, h: 0.28, fontFace: TK, fontSize: 10, bold: true, color: c, align: "right" });
pres.defineSlideMaster({ title: "LETTER", background: { color: CREAM }, objects: [titlePh(INK), footer("A094A8")], slideNumber: num(CHERRY) });
pres.defineSlideMaster({ title: "STAGE", background: { color: VIOLET }, objects: [titlePh(WHITE), footer("9C90B8")], slideNumber: num(LILAC) });
pres.defineSlideMaster({ title: "PLAIN", background: { color: VIOLET }, objects: [] });

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
  s.addShape(o.round ? pres.shapes.ROUNDED_RECTANGLE : pres.shapes.RECTANGLE, { x, y, w, h, rectRadius: o.round || 0, rotate: o.rot || 0, fill: { color: fill, transparency: o.tr || 0 }, line: o.line ? { color: o.line, width: o.lw || 0.75, dashType: o.dash || "solid" } : { type: "none" }, shadow: o.shadow ? { type: "outer", color: "1E1233", opacity: 0.25, blur: 8, offset: 2.5, angle: 90 } : undefined, objectName: nm("rect") });
}
function oval(s, x, y, d, fill, o = {}) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: fill ? { color: fill } : { type: "none" }, line: o.line ? { color: o.line, width: o.lw || 1 } : { type: "none" }, objectName: nm("oval") });
}
function star(s, x, y, d, color) {
  s.addShape(pres.shapes.STAR_5_POINT, { x, y, w: d, h: d, fill: { color }, line: { type: "none" }, objectName: nm("star") });
}
function t(s, text, o) { s.addText(text, { fontFace: SANS, color: INK, margin: 0, valign: "top", isTextBox: true, objectName: nm(o.name || "text"), ...o }); }
function full(s, key, alt) { s.addImage({ path: file(key), x: 0, y: 0, w: W, h: H, altText: alt || "배경", objectName: nm(key) }); }
// 응원봉 빛 테두리 사진
function pic(s, key, x, y, w, h, o = {}) {
  let f = fit(key, x, y, w, h);
  if (o.top) f = { ...f, y };
  if (o.left) f = { ...f, x };
  const b = o.border == null ? 0.06 : o.border;
  if (b) rect(s, f.x - b, f.y - b, f.w + 2 * b, f.h + 2 * b, o.dark ? VIOLET2 : WHITE, { round: 0.05, line: o.dark ? LILAC_D : LINE, lw: 1, shadow: o.shadow !== false });
  const link = o.url ? { hyperlink: { url: o.url, tooltip: "영상 보기" } } : {};
  s.addImage({ path: file(key), ...f, altText: o.alt || key, objectName: nm(`img-${key}`), ...link });
  if (o.cap) t(s, o.cap, { x: f.x - 0.1, y: f.y + f.h + b + 0.05, w: f.w + 0.2, h: 0.28, fontSize: o.capSize || 10.5, bold: true, color: o.dark ? "D9D2EA" : GRAY, align: "center", valign: "middle", name: "caption" });
  return f;
}
// 티켓 스텁 모양 라벨
function label(s, text, dark) {
  const w = 0.75 + text.length * 0.105;
  rect(s, MX, 0.36, w, 0.34, dark ? VIOLET2 : LILAC_L, { round: 0.06 });
  star(s, MX + 0.1, 0.43, 0.2, CHERRY);
  s.addShape(pres.shapes.LINE, { x: MX + 0.4, y: 0.4, w: 0, h: 0.26, line: { color: dark ? LILAC_D : LILAC, width: 1, dashType: "dash" }, objectName: nm("perf") });
  t(s, text, { x: MX + 0.5, y: 0.36, w: w - 0.5, h: 0.34, fontSize: 11, bold: true, color: dark ? "E3DCF5" : LILAC_D, valign: "middle", charSpacing: 1, name: "label" });
}
function slide(master, lab, title, section, bgKey) {
  const s = pres.addSlide({ masterName: master, sectionTitle: section });
  if (bgKey) full(s, bgKey, "작품 장면을 흐리게 깐 배경");
  label(s, lab, master === "STAGE");
  s.addText(title, { placeholder: "title" });
  return s;
}
function bullets(s, items, x, y, w, o = {}) {
  const color = o.color || INK, size = o.size || 14, gap = o.gap || 0.62;
  items.forEach((it, i) => {
    star(s, x, y + i * gap + size / 72 * 0.42, 0.18, o.mark || CHERRY);
    t(s, it, { x: x + 0.32, y: y + i * gap, w: w - 0.32, h: gap, fontSize: size, color, lineSpacingMultiple: 1.12, name: "bullet" });
  });
}
function linkBtn(s, text, url, x, y, w) {
  rect(s, x, y, w, 0.46, CHERRY, { round: 0.23 });
  s.addText([{ text, options: { hyperlink: { url, tooltip: url }, color: WHITE, bold: true } }], { x, y, w, h: 0.46, fontFace: SANS, fontSize: 12, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: nm("link") });
}
function info(s, rows, x, y, w, o = {}) {
  rows.forEach(([a, b], i) => {
    const yy = y + i * (o.gap || 0.44);
    t(s, a, { x, y: yy, w: 1.05, h: 0.38, fontSize: 12, bold: true, color: o.dark ? LILAC : LILAC_D, valign: "middle" });
    t(s, b, { x: x + 1.1, y: yy, w: w - 1.1, h: 0.38, fontSize: 13, color: o.dark ? WHITE : INK, valign: "middle" });
    s.addShape(pres.shapes.LINE, { x, y: yy + 0.41, w, h: 0, line: { color: o.dark ? "4A3A70" : LINE, width: 0.75 }, objectName: nm("rule") });
  });
}
function tag(s, text, x, y, dark) {
  t(s, text, { x, y, w: 4, h: 0.3, fontFace: TK, fontSize: 11, bold: true, color: dark ? LILAC : CHERRY, charSpacing: 3, name: "tag" });
}
// 콘서트 티켓: 본권(사진) + 절취선 + 스텁, 위아래 반원 노치
function ticket(s, x, y, w, h, stubW, imgKey, notch, alt) {
  rect(s, x, y, w, h, PAPER, { round: 0.06, shadow: true });
  const px = x + w - stubW, nd = Math.min(0.42, h * 0.16);
  if (imgKey) {
    const ix = x + 0.2, iy = y + 0.2, iw = w - stubW - 0.4, ih = h - 0.4;
    rect(s, ix, iy, iw, ih, VIOLET2);
    s.addImage({ path: file(imgKey), ...fit(imgKey, ix, iy, iw, ih), altText: alt || "티켓 대표 이미지", objectName: nm(`ticket-${imgKey}`) });
  }
  s.addShape(pres.shapes.LINE, { x: px, y: y + nd / 2 + 0.06, w: 0, h: h - nd - 0.12, line: { color: LILAC, width: 1.5, dashType: "dash" }, objectName: nm("perforation") });
  oval(s, px - nd / 2, y - nd / 2, nd, notch);
  oval(s, px - nd / 2, y + h - nd / 2, nd, notch);
  return px;
}
function chapter(n, tkKey, bgKey, title, sub, meta, section, alt) {
  const s = pres.addSlide({ masterName: "PLAIN", sectionTitle: section });
  full(s, bgKey, "챕터 대표 장면을 흐리고 어둡게 깐 배경");
  t(s, "PARK HYUNJIN  PORTFOLIO TOUR 2026", { x: 0.85, y: 0.55, w: 8, h: 0.35, fontFace: TK, fontSize: 12, bold: true, color: LILAC, charSpacing: 4 });
  const px = ticket(s, 0.85, 1.2, 11.63, 5.1, 3.75, tkKey, "2A1A48", alt);
  const sx = px + 0.3, sw = 3.2;
  t(s, "ADMIT ONE", { x: sx, y: 1.55, w: sw, h: 0.3, fontFace: TK, fontSize: 11, bold: true, color: LILAC_D, charSpacing: 5 });
  t(s, [{ text: "CH.", options: { fontSize: 22 } }, { text: `0${n}`, options: { fontSize: 60 } }], { x: sx, y: 1.85, w: sw, h: 1.05, fontFace: TK, bold: true, color: CHERRY, valign: "bottom" });
  t(s, title, { x: sx, y: 3.05, w: sw, h: 1.15, fontSize: 24, bold: true, color: INK, lineSpacingMultiple: 1.05 });
  t(s, sub, { x: sx, y: 4.3, w: sw, h: 1.0, fontSize: 13, color: GRAY, lineSpacingMultiple: 1.2 });
  [0, 1, 2].forEach((i) => star(s, sx + i * 0.32, 5.45, 0.2, i === 0 ? CHERRY : LILAC));
  t(s, meta, { x: sx, y: 5.72, w: sw, h: 0.3, fontSize: 10.5, bold: true, color: LILAC_D });
  return s;
}

// =====================================================================
// 01 표지
pres.addSection({ title: "소개" });
{
  const s = pres.addSlide({ masterName: "PLAIN", sectionTitle: "소개" });
  full(s, "bg_cover", "무대 위에서 노래하는 박수아 장면을 흐리게 깐 배경");
  const f = fit("cover_card", 8.55, 0.85, 4.1, 5.8);
  rect(s, f.x - 0.14, f.y - 0.14, f.w + 0.28, f.h + 0.28, LILAC, { round: 0.06, tr: 55 });
  rect(s, f.x - 0.06, f.y - 0.06, f.w + 0.12, f.h + 0.12, WHITE, { round: 0.04 });
  s.addImage({ path: file("cover_card"), ...f, altText: "마이크를 들고 무대에서 노래하는 박수아", objectName: nm("cover-card") });
  oval(s, f.x + f.w - 0.55, f.y + f.h - 0.55, 0.9, CHERRY);
  star(s, f.x + f.w - 0.3, f.y + f.h - 0.3, 0.4, WHITE);
  tag(s, "AI CONTENT CREATOR PORTFOLIO", MX, 1.35, true);
  t(s, "박현진", { x: MX, y: 1.75, w: 6.5, h: 1.2, fontSize: 60, bold: true, color: WHITE, valign: "middle" });
  t(s, "설렘과 응원의 순간을 장면으로 설계하는 크리에이터", { x: MX, y: 3.0, w: 7.4, h: 0.5, fontSize: 18, color: "E8E2F5" });
  [["No.01", "아이돌 애니메이션 「아이돌, 출발」"], ["No.02", "순정 숏폼 애니메이션 「너무 좋아 1000%」"], ["No.03", "캐릭터 IP 「일렁이는 뻄씨」"], ["No.04", "숏폼 광고 「3초매직」 / 「확빠져」"]].forEach(([c, x], i) => {
    const y = 3.85 + i * 0.52;
    star(s, MX, y + 0.11, 0.2, i % 2 ? LILAC : CHERRY);
    t(s, c, { x: MX + 0.35, y, w: 0.9, h: 0.42, fontFace: TK, fontSize: 13, bold: true, color: LILAC, valign: "middle" });
    t(s, x, { x: MX + 1.25, y, w: 6.3, h: 0.42, fontSize: 15, color: WHITE, valign: "middle" });
  });
  t(s, "서울시 매력일자리 AI 콘텐츠 크리에이터 양성과정  ·  2026", { x: MX, y: 6.7, w: 7, h: 0.3, fontSize: 10, color: "A89DC4" });
}

// 02 목차: 미니 티켓 4장
{
  const s = slide("LETTER", "CONTENTS", "목차", "소개", "bg_paper");
  const rows = [
    ["toc1", "아이돌, 출발", "아이돌 / 청춘 / 쌍방 성장물\n장편 애니메이션 3씬", "무대에서 노래하는 박수아"],
    ["toc2", "너무 좋아 1000%", "순정 학원물 로맨스\n1분 9:16 숏폼 애니메이션", "등굣길에서 마주친 유하나와 강민"],
    ["toc3", "일렁이는 뻄씨", "캐릭터 IP\n일상툰 / 이모티콘 / 굿즈", "꽃밭 속 뻄씨와 똘랑뻄"],
    ["toc4", "3초매직 / 확빠져", "숏폼 광고 2편\n컨셉 / 스토리보드 / 완성 영상", "트리트먼트와 다이어트 보조제 제품 컨셉"],
  ];
  rows.forEach(([k, title, meta, alt], i) => {
    const x = MX + (i % 2) * 6.1, y = i < 2 ? 1.85 : 4.4, w = 5.85, h = 2.3;
    const px = ticket(s, x, y, w, h, 3.0, k, CREAM, alt);
    t(s, `No.0${i + 1}`, { x: px + 0.28, y: y + 0.22, w: 2.5, h: 0.45, fontFace: TK, fontSize: 20, bold: true, color: CHERRY });
    t(s, title, { x: px + 0.28, y: y + 0.72, w: 2.6, h: 0.5, fontSize: 17, bold: true, color: INK, valign: "middle" });
    t(s, meta, { x: px + 0.28, y: y + 1.28, w: 2.62, h: 0.8, fontSize: 11, color: GRAY, lineSpacingMultiple: 1.2 });
  });
}

// 03 역량
{
  const s = slide("LETTER", "PROFILE", "성장 서사 기획부터 캐릭터 IP, 광고까지 아우르는 역량", "소개", "bg_paper");
  const cards = [
    ["청춘 성장 서사 기획", "아이돌과 팬의 쌍방 성장을 기승전결 3씬으로 압축한 애니메이션 기획"],
    ["캐릭터 설계", "AI툴을 활용한 인물별 프로필 / 의상 / 포인트 컬러 설계로 캐릭터 일관성 확보"],
    ["장면 연출 / 사운드", "색보정 / 소품 / 시선 규칙과 컷별 보이스 / BGM / 현장음으로 몰입감 극대화"],
    ["IP / 광고 확장", "캐릭터 굿즈 상세페이지와 9:16 숏폼 광고로 상품성 있는 콘텐츠 기획"],
  ];
  cards.forEach(([h, b], i) => {
    const x = MX + i * 3.03, y = 1.95, w = 2.83, hh = 4.4;
    rect(s, x, y, w, hh, WHITE, { round: 0.06, line: LINE, shadow: true });
    rect(s, x, y, w, 0.95, i % 2 ? LILAC_L : "FBE3E6", { round: 0.06 });
    rect(s, x, y + 0.6, w, 0.35, i % 2 ? LILAC_L : "FBE3E6");
    t(s, `0${i + 1}`, { x: x + 0.28, y: y + 0.18, w: 1.2, h: 0.6, fontFace: TK, fontSize: 28, bold: true, color: i % 2 ? LILAC_D : CHERRY, valign: "middle" });
    star(s, x + w - 0.6, y + 0.32, 0.32, i % 2 ? LILAC : CHERRY);
    s.addShape(pres.shapes.LINE, { x: x + 0.15, y: y + 0.95, w: w - 0.3, h: 0, line: { color: LILAC, width: 1.25, dashType: "dash" }, objectName: nm("perf") });
    t(s, h, { x: x + 0.28, y: y + 1.25, w: w - 0.45, h: 0.5, fontSize: 18, bold: true, color: INK });
    t(s, b, { x: x + 0.28, y: y + 1.9, w: w - 0.5, h: 2.2, fontSize: 13.5, color: "4E465C", lineSpacingMultiple: 1.28 });
  });
}

// =====================================================================
// CHAPTER 01 아이돌, 출발
const L1 = "CH.01  아이돌, 출발", S1 = "01 아이돌, 출발";
pres.addSection({ title: S1 });
chapter(1, "tk_ch1", "ch1_bg", "아이돌, 출발", "평범한 소녀가 팬의 응원으로 정상의 무대에 서기까지", "아이돌 / 청춘 / 쌍방 성장물", S1, "박수아와 백설하가 함께 선 데뷔 무대");

// 작품 개요
{
  const s = slide("STAGE", L1, "평범한 소녀가 팬의 응원으로 무대에 서는 쌍방 성장물", S1, "bg_stage");
  pic(s, "i42", MX, 1.9, 5.4, 3.04, { dark: true, alt: "응원봉 물결 앞에서 노래하는 박수아" });
  t(s, "16:9 FHD / 일본 아이돌 애니메이션 화풍 / 시네마틱 연출", { x: MX, y: 5.15, w: 5.4, h: 0.3, fontSize: 11, color: "B9AFD3" });
  const x = 6.55;
  tag(s, "LOGLINE", x, 1.85, true);
  t(s, "평범한 여중생 A가 우연히 본 아이돌 무대에 반해 아이돌을 꿈꾸고, 갖은 노력 끝에 정상에 오르는 이야기", { x, y: 2.15, w: 6.05, h: 0.9, fontSize: 15, bold: true, color: WHITE, lineSpacingMultiple: 1.2 });
  info(s, [["장르", "아이돌 / 청춘 / 쌍방(팬, 아이돌) 성장물"], ["플랫폼", "카카오 / 네이버"], ["타깃", "10~30대 여성 / 남성"]], x, 3.2, 6.05, { dark: true });
  tag(s, "PLANNING", x, 4.65, true);
  bullets(s, [
    "단순한 동경을 넘어 확고한 꿈을 향해 나아가는 청춘의 성장통 표현",
    "팬과 아이돌이 서로를 비추는 등대처럼 함께 역경을 넘는 과정 연결",
  ], x, 5.0, 6.05, { size: 13.5, gap: 0.66, color: WHITE });
}

// 기승전결
{
  const s = slide("LETTER", L1, "기 / 승 / 전 / 결, 4단계로 설계한 성장 서사", S1, "bg_paper");
  const acts = [
    ["기", "운명적인 끌림", "i18", "하굣길 야외 무대에서 아이돌의 꿈을 품고 소속사의 문을 두드림"],
    ["승", "녹록지 않은 현실, 뜻밖의 위로", "i29", "천재들 사이의 비교와 무명의 설움 속 팬레터와 선물로 다시 일어섬"],
    ["전", "편지가 노래가 되던 날", "i39a", "팬의 편지 문장을 가사로 옮겨 데뷔곡 완성"],
    ["결", "함께 비추는 등대", "i411", "데뷔 무대 객석에서 같은 액세서리를 한 팬을 발견하고 전속 작곡가로 제안"],
  ];
  acts.forEach(([a, h, k, d], i) => {
    const x = MX + i * 3.05, w = 2.85;
    oval(s, x, 1.85, 0.5, i === 3 ? CHERRY : LILAC_D);
    t(s, a, { x, y: 1.85, w: 0.5, h: 0.5, fontSize: 16, bold: true, color: WHITE, align: "center", valign: "middle" });
    t(s, h, { x: x + 0.6, y: 1.85, w: w - 0.55, h: 0.5, fontSize: 13, bold: true, color: INK, valign: "middle" });
    pic(s, k, x, 2.6, w, 1.6, { alt: `${h} 대표 장면` });
    t(s, d, { x, y: 4.45, w, h: 1.3, fontSize: 13, color: INK, lineSpacingMultiple: 1.25 });
    if (i < 3) t(s, "▶", { x: x + w - 0.02, y: 3.2, w: 0.22, h: 0.4, fontSize: 11, color: LILAC_D, align: "center", valign: "middle" });
  });
  rect(s, MX, 5.95, 11.93, 0.62, LILAC_L, { round: 0.1 });
  t(s, "\"모든 걸 포기하고 싶어질 무렵, A를 다시 일으켜 세운 것은 묵묵히 곁을 지켜준 팬들의 믿음 어린 응원\"", { x: MX + 0.3, y: 5.95, w: 11.4, h: 0.62, fontSize: 13, bold: true, color: LILAC_D, valign: "middle" });
}

// 캐릭터 시트
{
  const s = slide("LETTER", L1, "포인트 컬러와 의상으로 역할을 구분한 캐릭터 설계", S1, "bg_paper");
  const sheets = [
    ["sheet_sua", "박수아  ·  16세 주인공, 검정 / 은색 / 흰색"],
    ["sheet_sua2", "박수아 연습복  ·  연습생 시절 의상 변형"],
    ["sheet_yujin", "이유진  ·  수아가 동경하게 된 아이돌, 연보라"],
    ["sheet_jinwoo", "김진우  ·  팬 출신 프로듀서, 청자색"],
  ];
  sheets.forEach(([k, c], i) => {
    const x = MX + (i % 2) * 4.1, y = i < 2 ? 1.8 : 4.35;
    pic(s, k, x, y, 3.9, 2.15, { alt: `${c} 캐릭터 시트`, cap: c });
  });
  const x = 9.15;
  tag(s, "CHARACTER", x, 1.85);
  t(s, "박수아", { x, y: 2.2, w: 3.5, h: 0.5, fontSize: 24, bold: true, color: INK });
  t(s, "밝고 당찬 노력파, 은빛 눈동자와 삼지창 앞머리 포니테일", { x, y: 2.8, w: 3.5, h: 0.8, fontSize: 12.5, color: GRAY, lineSpacingMultiple: 1.2 });
  bullets(s, [
    "인물마다 포인트 컬러를 지정해 화면 속 역할 구분 강화",
    "턴어라운드 / 표정 / 의상 변형 시트로 컷 간 인물 일관성 확보",
  ], x, 3.85, 3.5, { size: 13, gap: 1.05 });
}

// 라이벌과 팬
{
  const s = slide("STAGE", L1, "라이벌 백설하와 무명팬 송아연, 주인공을 비추는 두 인물", S1, "bg_stage");
  pic(s, "i210", MX, 1.9, 3.7, 2.08, { dark: true, alt: "백설하 클로즈업" });
  pic(s, "i45f", MX + 3.9, 1.9, 3.7, 2.08, { dark: true, alt: "제복풍 무대 의상의 백설하" });
  pic(s, "i411", MX, 4.35, 7.6, 2.2, { dark: true, alt: "객석에서 휠체어에 앉아 무대를 바라보는 송아연" });
  const x = 8.75;
  tag(s, "RIVAL", x, 1.9, true);
  t(s, "백설하  ·  17세 톱 아이돌", { x, y: 2.2, w: 3.9, h: 0.4, fontSize: 16, bold: true, color: WHITE });
  t(s, "천재성과 자존심의 라이벌, 수아의 진심과 노력을 보고 동반자로 인정", { x, y: 2.65, w: 3.9, h: 0.9, fontSize: 12.5, color: "D9D2EA", lineSpacingMultiple: 1.2 });
  tag(s, "FAN", x, 4.35, true);
  t(s, "송아연  ·  16세 무명팬", { x, y: 4.65, w: 3.9, h: 0.4, fontSize: 16, bold: true, color: WHITE });
  t(s, "사고로 아이돌의 꿈을 접은 소녀, 무명 시절부터 편지와 선물로 응원하다 수아의 전속 작곡가로", { x, y: 5.1, w: 3.9, h: 1.2, fontSize: 12.5, color: "D9D2EA", lineSpacingMultiple: 1.2 });
}

// 씬1
{
  const s = slide("LETTER", L1, "씬1 하굣길에서 되찾은 꿈, 색과 소품 규칙으로 설계한 연속성", S1, "bg_paper");
  pic(s, "i18", MX, 1.85, 6.2, 3.49, { alt: "하굣길에 드러나는 대형 야외 무대" });
  [["i12", "인파 속 하굣길"], ["i16", "응원봉 빛에 물든 얼굴"], ["i112", "눈동자 위 겹치는 잔상"], ["i114", "나도 저렇게 되고 싶어"]].forEach(([k, alt], i) => {
    pic(s, k, 7.15 + (i % 2) * 2.79, i < 2 ? 1.85 : 3.6, 2.69, 1.51, { alt });
  });
  bullets(s, [
    "회상은 황금빛 / 세피아, 현재는 청록빛으로 색보정을 규칙화해 시간대 구분",
    "이어폰 착용 / 제거와 은빛 눈동자 고정 규칙으로 컷 간 연속성 확보",
    "익스트림 클로즈업과 이중노출로 꿈을 되찾는 순간의 감정 극대화",
  ], MX, 5.6, 11.9, { size: 13.5, gap: 0.46 });
}

// 씬2
{
  const s = slide("STAGE", L1, "씬2 연습생의 벽, 청록 톤으로 그린 무명의 시간", S1, "bg_stage");
  [["i22b", "월말 평가 심사석"], ["i23d", "프로듀서 김진우"], ["i24", "트레이닝 평가 결과 공지"], ["i28", "우편함 앞의 수아"], ["i29", "편지와 체리 손뜨개 키링"], ["i211", "편지를 가슴에 안은 수아"]].forEach(([k, alt], i) => {
    pic(s, k, MX + (i % 3) * 4.02, i < 3 ? 1.85 : 4.12, 3.85, 2.07, { dark: true, alt, cap: alt });
  });
  t(s, "평가표 / 순위 명단 / 우편함 인서트로 경쟁과 좌절, 위로의 흐름을 대사 없이 전달", { x: MX, y: 6.62, w: 11.9, h: 0.35, fontSize: 12.5, bold: true, color: LILAC, valign: "middle" });
}

// 씬3
{
  const s = slide("LETTER", L1, "씬3 편지가 노래가 되던 날, 청록에서 따뜻한 톤으로", S1, "bg_paper");
  pic(s, "i314", MX, 1.85, 6.2, 3.49, { alt: "수아와 백설하의 데뷔 무대" });
  [["i34a", "팔로워 상승 SNS 화면"], ["i34d", "연습생 월간 평가 순위"], ["i34f", "데뷔조가 된 세 사람"], ["i39a", "녹음 부스의 수아와 설하"]].forEach(([k, alt], i) => {
    pic(s, k, 7.15 + (i % 2) * 2.79, i < 2 ? 1.85 : 3.6, 2.69, 1.51, { alt });
  });
  bullets(s, [
    "후반부로 갈수록 따뜻해지는 색보정으로 감정의 해소 시각화",
    "SNS 팔로워 / 순위표 인서트로 인지도 상승을 한눈에 전달",
    "씬1의 객석 시점과 대구를 이루는 데뷔 무대 연출",
  ], MX, 5.6, 11.9, { size: 13.5, gap: 0.46 });
}

// 반복 모티프
{
  const s = slide("STAGE", L1, "체리 키링과 팬레터, 팬과 아이돌을 잇는 반복 모티프", S1, "bg_stage");
  [["i29", "편지와 체리 키링"], ["i32", "키링을 단 첫 거리 무대"], ["i42", "데뷔 무대 위 체리 참"], ["i44", "응원봉 물결 속 팬들"]].forEach(([k, c], i) => {
    pic(s, k, MX + i * 3.03, 1.9, 2.85, 1.6, { dark: true, alt: c, cap: c });
  });
  bullets(s, [
    "팬이 보낸 체리 손뜨개 키링을 무대 의상 액세서리로 이어 응원의 흔적 시각화",
    "편지 문장이 데뷔곡 가사가 되는 구조로 팬과 아이돌의 쌍방 성장 완성",
    "같은 체리 굿즈를 든 팬을 객석에서 발견하는 엔딩으로 감동 극대화",
  ], MX, 4.25, 11.9, { size: 15, gap: 0.7, color: WHITE, mark: CHERRY });
}

// 편집본과 사운드
{
  const s = slide("STAGE", L1, "통합 편집본과 사운드 설계", S1, "bg_stage");
  pic(s, "poster_idol", MX, 1.9, 6.7, 3.77, { dark: true, url: LINK_IDOL, alt: "아이돌, 출발 편집본 영상 보기" });
  linkBtn(s, "▶  편집본 보기", LINK_IDOL, MX, 5.95, 2.4);
  const x = 7.9;
  tag(s, "SOUND DESIGN", x, 1.9, true);
  const rows = [["보이스", "컷별 캐릭터 대사 / 속생각 녹음"], ["BGM", "노을빛 하굣길 / 노을 끝의 첫 무대 / 체리 키링의 약속"], ["현장음", "콘서트 환호 / 웅성 / 거리 일상"]];
  rows.forEach(([a, b], i) => {
    const y = 2.3 + i * 0.95;
    rect(s, x, y, 4.73, 0.8, VIOLET2, { round: 0.08, line: "4A3A70" });
    t(s, a, { x: x + 0.2, y, w: 1.0, h: 0.8, fontSize: 13, bold: true, color: LILAC, valign: "middle" });
    t(s, b, { x: x + 1.2, y, w: 3.4, h: 0.8, fontSize: 12, color: WHITE, valign: "middle", lineSpacingMultiple: 1.15 });
  });
  bullets(s, ["장면별 BGM과 현장음 레이어로 무대의 몰입감 강화"], x, 5.3, 4.73, { size: 13, gap: 0.7, color: WHITE });
}

// =====================================================================
// CHAPTER 02 너무 좋아 1000%
const L2 = "CH.02  너무 좋아 1000%", S2 = "02 너무 좋아 1000%";
pres.addSection({ title: S2 });
chapter(2, "tk_ch2", "ch2_bg", "너무 좋아\n1000%", "등굣길의 우연한 부딪힘으로 시작되는 첫사랑의 떨림", "순정 학원물 / 1분 숏폼", S2, "벚꽃 골목에서 부딪혀 안기듯 넘어지는 유하나와 강민");

{
  const s = slide("LETTER", L2, "1분 숏폼에 압축한 첫사랑의 설렘, 순정 학원물", S2, "bg_spring");
  pic(s, "game1", MX, 1.85, 2.81, 5.0, { alt: "호감도 게이지와 대사창이 있는 게임 UI 버전 컷" });
  const x = 3.85;
  tag(s, "LOGLINE", x, 1.85);
  t(s, "등굣길에서 부딪혀 첫눈에 반한 두 동급생, 교내 곳곳에서 시선이 마주칠수록 커지는 설렘 끝에 처음 만난 그 길에서 동시에 고백", { x, y: 2.15, w: 4.85, h: 1.2, fontSize: 13.5, bold: true, color: INK, lineSpacingMultiple: 1.2 });
  info(s, [["장르", "순정만화 / 학원물 / 로맨스"], ["플랫폼", "카카오 / 네이버"], ["타깃", "10~20대 여성"], ["포맷", "1분 내외 / 9:16 세로"]], x, 3.5, 4.85);
  bullets(s, [
    "대사보다 표정 / 시선 / 색감 변화로 감정 전달",
    "파스텔톤과 벚꽃, 봄 햇살로 산뜻한 분위기 연출",
  ], x, 5.4, 4.85, { size: 13, gap: 0.55 });
  pic(s, "sheet_hana", 9.05, 1.85, 3.58, 2.02, { alt: "유하나 캐릭터 시트", cap: "유하나  ·  17세, 감정이 얼굴에 드러나는 덜렁이" });
  pic(s, "sheet_min", 9.05, 4.4, 3.58, 2.02, { alt: "강민 캐릭터 시트", cap: "강민  ·  17세, 무심한 듯 다정한 관찰자" });
}

{
  const s = slide("LETTER", L2, "등굣길의 충돌, 슬로우와 클로즈업으로 만든 첫 두근거림", S2, "bg_spring");
  [["l11", "늦었다, 늦었어!"], ["l14", "모퉁이 정면충돌"], ["l15", "안기듯 넘어지는 찰나"], ["l18", "동시 사과, 이마 콩!"], ["l19", "반대 방향으로 도망"]].forEach(([k, c], i) => {
    pic(s, k, MX + i * 2.4, 1.85, 2.25, 4.0, { alt: c, cap: c });
  });
  bullets(s, ["식빵 등교 / 정면충돌 / 동시 사과 등 순정만화 클리셰를 경쾌한 컷 리듬으로 재해석"], MX, 6.4, 11.9, { size: 14, gap: 0.5 });
}

{
  const s = slide("LETTER", L2, "교내 몽타주와 게임 UI, 수미상관으로 완성한 감정의 원", S2, "bg_spring");
  [["l21", "운동장"], ["l22", "눈 마주침"], ["l23", "급식실"], ["game2", "게임 UI 버전"]].forEach(([k, c], i) => {
    pic(s, k, MX + i * 2.12, 1.85, 1.98, 3.52, { alt: c, cap: c });
  });
  const x = 9.35;
  tag(s, "STRUCTURE", x, 1.85);
  const steps = ["등굣길  ·  만남", "운동장 / 급식실", "동아리실 / 음악실", "등굣길  ·  동시 고백"];
  steps.forEach((st, i) => {
    const y = 2.25 + i * 0.85;
    rect(s, x, y, 3.28, 0.58, i === 0 || i === 3 ? "FBE3E6" : LILAC_L, { round: 0.12 });
    t(s, st, { x, y, w: 3.28, h: 0.58, fontSize: 13, bold: true, color: i === 0 || i === 3 ? CHERRY : LILAC_D, align: "center", valign: "middle" });
    if (i < 3) t(s, "▼", { x, y: y + 0.58, w: 3.28, h: 0.27, fontSize: 10, color: LILAC_D, align: "center", valign: "middle" });
  });
  bullets(s, [
    "장소마다 햇살 / 소란 / 차분 / 따뜻함으로 톤을 바꿔 호감 상승 표현",
    "호감도 게이지 UI 버전으로 게임형 콘텐츠 확장 가능성 제시",
  ], MX, 5.85, 11.9, { size: 13.5, gap: 0.48 });
}

// =====================================================================
// CHAPTER 03 일렁이는 뻄씨
const L3 = "CH.03  일렁이는 뻄씨", S3 = "03 일렁이는 뻄씨";
pres.addSection({ title: S3 });
chapter(3, "tk_ch3", "ch3_bg", "일렁이는 뻄씨", "걱정 없이 사는 듯하지만 일상이 고단한 말랑 캐릭터 듀오", "캐릭터 IP / 일상툰 / 굿즈", S3, "꽃밭 속에서 웃는 뻄씨와 똘랑뻄");

{
  const s = slide("LETTER", L3, "T 성향 뻄씨와 F 성향 똘랑뻄, 성격 대비로 만든 캐릭터 듀오", S3, "bg_bbem");
  pic(s, "sheet_bbem", MX, 1.9, 5.8, 4.35, { alt: "뻄씨 캐릭터 시트", cap: "뻄씨  ·  감정 표현이 확실한 T 성향" });
  pic(s, "sheet_ddol", 6.83, 1.9, 5.8, 4.35, { alt: "똘랑뻄 캐릭터 시트", cap: "똘랑뻄  ·  상냥하고 다정한 F 성향, 꼬리 방울 포인트" });
}

{
  const s = slide("LETTER", L3, "일상툰과 움직이는 이모티콘으로 넓힌 공감 포인트", S3, "bg_bbem");
  pic(s, "toon", MX, 1.85, 4.0, 5.0, { alt: "출근길 일상툰 4컷" });
  pic(s, "logo2", 5.2, 1.85, 3.3, 1.99, { alt: "일렁이는 뻄씨 브랜드 로고, 똘랑뻄 버전" });
  pic(s, "logo1", 8.95, 1.85, 3.3, 1.99, { alt: "일렁이는 뻄씨 브랜드 로고" });
  [["emo_nyam", "냠냠"], ["emo_hi", "안녕"], ["emo_bye", "바이"], ["emo_bus", "버스"], ["emo_q", "물음표"]].forEach(([k, c], i) => {
    pic(s, k, 5.2 + i * 1.5, 4.3, 1.32, 1.32, { alt: `${c} 움직이는 이모티콘`, cap: c });
  });
  bullets(s, [
    "비몽사몽 기상 / 지옥철 출근 / 커피 충전 등 직장인 일상 공감 에피소드",
    "말랑한 라인 드로잉 움직이는 이모티콘으로 메신저 활용도 강화",
  ], 5.2, 6.12, 7.45, { size: 12.5, gap: 0.42 });
}

{
  const s = slide("LETTER", L3, "볼캡 / 블루투스 스피커 / 텀블러, 굿즈 3종 상세페이지 기획", S3, "bg_bbem");
  [["hat_detail", "hat_page", "포근 퍼 볼캡", LINK_HAT], ["spk_detail", "spk_page", "블루투스 스피커", LINK_SPK], ["tum_detail", "tum_page", "뻄 텀블러", LINK_TUM]].forEach(([d, p, c, url], i) => {
    const x = MX + i * 4.05;
    pic(s, d, x, 1.85, 2.55, 3.4, { alt: `${c} 상세 설명 이미지` });
    pic(s, p, x + 2.75, 1.85, 1.05, 3.4, { alt: `${c} 광고 페이지 전체` });
    t(s, c, { x, y: 5.45, w: 3.8, h: 0.4, fontSize: 15, bold: true, color: INK });
    linkBtn(s, "▶  실사 광고 영상", url, x, 5.95, 2.3);
  });
  t(s, "캐릭터 포인트를 살린 제품 디자인 / 상세 설명 / 광고 페이지 / 실사 광고까지 판매 동선 설계", { x: MX, y: 6.55, w: 11.9, h: 0.35, fontSize: 12.5, bold: true, color: LILAC_D, valign: "middle" });
}

{
  const s = slide("STAGE", L3, "홍보영상과 실사 합성 광고 이미지로 완성한 IP 브랜딩", S3, "bg_stage");
  pic(s, "poster_bbem", MX, 1.9, 6.4, 3.6, { dark: true, url: LINK_BBEM, alt: "일렁이는 뻄씨 홍보영상 보기" });
  linkBtn(s, "▶  홍보영상 보기", LINK_BBEM, MX, 5.8, 2.4);
  pic(s, "spk_main", 7.5, 1.9, 2.6, 3.25, { dark: true, alt: "스피커 굿즈 실사 합성 광고 이미지", cap: "실사 합성 광고 이미지" });
  pic(s, "promo1", 10.35, 1.9, 2.28, 1.28, { dark: true, alt: "단순 캐릭터 홍보 영상 1 장면" });
  pic(s, "promo2", 10.35, 3.55, 2.28, 1.28, { dark: true, alt: "단순 캐릭터 홍보 영상 2 장면" });
  t(s, "SNS용 짧은 캐릭터 모션", { x: 10.25, y: 4.95, w: 2.5, h: 0.28, fontSize: 10.5, bold: true, color: "D9D2EA", align: "center" });
  bullets(s, ["캐릭터 홍보영상과 짧은 모션 클립으로 SNS 노출용 콘텐츠 확보"], 7.5, 5.8, 5.13, { size: 13, gap: 0.6, color: WHITE });
}

// =====================================================================
// CHAPTER 04 숏폼 광고
const L4 = "CH.04  숏폼 광고 2편", S4 = "04 숏폼 광고";
pres.addSection({ title: S4 });
chapter(4, "tk_ch4", "ch4_bg", "3초매직 / 확빠져\n숏폼 광고", "좌절에서 기쁨까지, 15초 감성 스토리텔링 광고 2편", "세로 9:16 / 15초", S4, "3초매직 트리트먼트와 확빠져 다이어트 보조제 제품 컨셉");

{
  const s = slide("LETTER", L4, "3초매직 트리트먼트, 즉각 변화를 보여 주는 15초 스토리텔링", S4, "bg_ad");
  pic(s, "poster_magic", MX, 1.85, 3.6, 4.5, { url: LINK_MAGIC, alt: "3초매직 광고 영상 보기" });
  linkBtn(s, "▶  광고 영상 보기", LINK_MAGIC, MX, 6.5, 2.4);
  const x = 4.75;
  pic(s, "magic_concept", x, 1.85, 4.4, 2.93, { alt: "3초매직 컨셉 보드: 박곱슬 캐릭터와 보라색 튜브 제품" });
  info(s, [["포맷", "세로 9:16 / 15초"], ["타깃", "머릿결 스트레스 20~30대"], ["USP", "쓰자마자 보이는 매직 효과"]], 9.45, 1.9, 3.18, { gap: 0.6 });
  const beats = [["훅", "00:00"], ["의심", "00:04.5"], ["사용", "00:06"], ["변화", "00:09"], ["CTA", "00:12"]];
  beats.forEach(([b, tc], i) => {
    const bx = x + i * 1.6;
    rect(s, bx, 5.15, 1.4, 0.9, i === 3 ? "FBE3E6" : LILAC_L, { round: 0.1 });
    t(s, tc, { x: bx, y: 5.2, w: 1.4, h: 0.3, fontFace: TK, fontSize: 10.5, bold: true, color: GRAY, align: "center" });
    t(s, b, { x: bx, y: 5.5, w: 1.4, h: 0.45, fontSize: 15, bold: true, color: i === 3 ? CHERRY : LILAC_D, align: "center", valign: "middle" });
    if (i < 4) t(s, "▶", { x: bx + 1.4, y: 5.4, w: 0.2, h: 0.4, fontSize: 9, color: LILAC_D, align: "center", valign: "middle" });
  });
  t(s, "10컷 타임코드 설계 / 대사 없이 표정 연기로 감정 전달 / 반짝임 이펙트로 변화 강조", { x, y: 6.25, w: 7.9, h: 0.35, fontSize: 12.5, bold: true, color: INK, valign: "middle" });
}

{
  const s = slide("LETTER", L4, "확빠져 다이어트 보조제, 계절의 흐름으로 그린 꾸준함", S4, "bg_ad");
  pic(s, "sb", MX, 1.85, 6.9, 4.42, { alt: "확빠져 10컷 스토리보드" });
  const x = 7.95;
  pic(s, "poster_slim", x, 1.85, 2.4, 3.0, { url: LINK_SLIM, alt: "확빠져 광고 영상 보기" });
  pic(s, "char_before", 10.6, 1.85, 1.0, 1.35, { alt: "박복슬 캐릭터 레퍼런스" });
  pic(s, "char_after", 11.68, 1.85, 1.0, 1.35, { alt: "박곱슬 캐릭터 레퍼런스" });
  t(s, "캐릭터 레퍼런스", { x: 10.55, y: 3.35, w: 2.1, h: 0.28, fontSize: 10, bold: true, color: GRAY, align: "center" });
  linkBtn(s, "▶  광고 영상 보기", LINK_SLIM, 10.6, 4.35, 2.05);
  bullets(s, [
    "여름 / 가을 / 겨울 / 봄 창밖 풍경 변화로 꾸준한 섭취의 시간 경과 표현",
    "과장된 감량 수치 / 의료적 효과 표현을 배제한 신뢰감 있는 연출",
  ], x, 5.15, 4.68, { size: 12.5, gap: 0.68 });
}

// =====================================================================
// 엔딩: 팬레터 봉투
pres.addSection({ title: "마무리" });
{
  const s = pres.addSlide({ masterName: "PLAIN", sectionTitle: "마무리" });
  full(s, "bg_stage", "응원봉 물결을 흐리고 어둡게 깐 배경");
  t(s, "THANK YOU FOR WATCHING", { x: 0, y: 0.8, w: W, h: 0.4, fontFace: TK, fontSize: 14, bold: true, color: LILAC, align: "center", charSpacing: 6 });
  const ew = 6.2, eh = 3.9, ex = (W - ew) / 2, ey = 1.55;
  rect(s, ex, ey, ew, eh, PAPER, { shadow: true });
  s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: ex, y: ey, w: ew, h: 1.75, rotate: 180, fill: { color: "F6E9E4" }, line: { color: "E2CFC8", width: 1 }, objectName: nm("envelope-flap") });
  oval(s, W / 2 - 0.45, ey + 1.75 - 0.45, 0.9, CHERRY);
  t(s, "♥", { x: W / 2 - 0.45, y: ey + 1.75 - 0.45, w: 0.9, h: 0.9, fontSize: 24, color: WHITE, align: "center", valign: "middle" });
  t(s, "감사합니다", { x: ex, y: ey + 2.35, w: ew, h: 0.7, fontSize: 32, bold: true, color: INK, align: "center", valign: "middle" });
  t(s, "박현진", { x: ex, y: ey + 3.0, w: ew, h: 0.4, fontSize: 17, bold: true, color: CHERRY, align: "center", valign: "middle" });
  t(s, "아이돌 애니메이션 / 순정 숏폼 / 캐릭터 IP / 숏폼 광고", { x: ex, y: ey + 3.4, w: ew, h: 0.32, fontSize: 11, color: GRAY, align: "center", valign: "middle" });
  [-1, 0, 1].forEach((i) => star(s, W / 2 + i * 0.45 - 0.11, 5.9, 0.22, i ? LILAC : CHERRY));
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();

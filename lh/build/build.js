// 임현 포트폴리오: AI 애니메이션 / 캐릭터 IP TENSHI / 학원물 캐릭터 기획
// node build.js  ->  ../임현_포트폴리오.pptx
// 버튜버 방송 화면 모티프: 플레이어 창 + LIVE 배지 + 채팅창, 챕터가 넘어갈수록 차오르는 재생 바
// 색: 철수 방의 밤 인디고, 방송 핑크, TENSHI 하늘색(#BDE8FF / #1F8FE6), 비 오는 교실 네이비
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const IMG = path.join(__dirname, "../work/img");
const DIMS = require("../work/dims.json");
const OUT = path.join(__dirname, "../임현_포트폴리오.pptx");

const LINK_MV = "https://drive.google.com/file/d/1cZqWLTYq6Tai5tQZoIr-jXhBxo9EdzGs/view";
const LINK_TENSHI = "https://drive.google.com/file/d/1YcNxksqSytBIN9XAH8KY3LJ4UgyTcMVg/view";

const SANS = "Malgun Gothic";
const UI = "Segoe UI";
const NIGHT = "15132B", PANEL = "221F3F", PANEL2 = "2C2850", CREAM = "FFF8F0", INK = "221D33", MUTED = "8F89AE";
const PINK = "FF5C9E", PINK_L = "FF9CC6", LAV = "A99BFF", SKY = "1F8FE6", SKY_L = "BDE8FF", SKY_BG = "EEF8FF", RAIN = "3D5A99", RAIN_L = "9DB4E8", SUN = "F2A65A";

const THEME = {
  name: "Live Stream",
  headFontFace: SANS,
  bodyFontFace: SANS,
  colors: { dk1: NIGHT, lt1: "FFFFFF", dk2: PANEL, lt2: CREAM, accent1: PINK, accent2: SKY, accent3: LAV, accent4: RAIN, accent5: SUN, accent6: MUTED, hlink: PINK, folHlink: PINK },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: SANS, bodyFontFace: SANS };
pres.author = "임현";
pres.title = "임현 포트폴리오";
const W = 13.333, H = 7.5, MX = 0.7;

const footer = (c) => ({ text: { text: "임현  ·  AI 콘텐츠 크리에이터 포트폴리오", options: { x: MX, y: 7.08, w: 6, h: 0.28, fontFace: SANS, fontSize: 9, color: c, margin: 0 } } });
const titlePh = (c) => ({ placeholder: { options: { name: "title", type: "title", x: MX, y: 0.78, w: 11.9, h: 0.7, fontFace: SANS, fontSize: 27, bold: true, color: c, align: "left", valign: "middle", margin: 0 }, text: "" } });
const num = (c) => ({ x: 12.03, y: 7.08, w: 0.6, h: 0.28, fontFace: UI, fontSize: 9, bold: true, color: c, align: "right" });
pres.defineSlideMaster({ title: "NIGHT", background: { color: NIGHT }, objects: [titlePh("FFFFFF"), footer("6E6890")], slideNumber: num("6E6890") });
pres.defineSlideMaster({ title: "CREAM", background: { color: CREAM }, objects: [titlePh(INK), footer("A39C93")], slideNumber: num("A39C93") });
pres.defineSlideMaster({ title: "SKYBG", background: { color: SKY_BG }, objects: [titlePh("143A66"), footer("7FA6CC")], slideNumber: num("7FA6CC") });
pres.defineSlideMaster({ title: "PLAIN", background: { color: NIGHT }, objects: [] });

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
  s.addShape(o.round ? pres.shapes.ROUNDED_RECTANGLE : pres.shapes.RECTANGLE, { x, y, w, h, rectRadius: o.round || 0, fill: { color: fill, transparency: o.tr || 0 }, line: o.line ? { color: o.line, width: o.lw || 0.75 } : { type: "none" }, objectName: nm("rect") });
}
function img(s, key, x, y, w, h, o = {}) {
  let f = fit(key, x, y, w, h);
  if (o.top) f = { ...f, y };
  if (o.left) f = { ...f, x };
  if (o.frame) rect(s, f.x - 0.06, f.y - 0.06, f.w + 0.12, f.h + 0.12, o.frame, { round: 0.08 });
  const link = o.url ? { hyperlink: { url: o.url, tooltip: o.tip || "영상 보기" } } : {};
  s.addImage({ path: file(key), ...f, altText: o.alt || key, objectName: nm(`img-${key}`), ...link });
  return f;
}
function t(s, text, o) { s.addText(text, { fontFace: SANS, color: "FFFFFF", margin: 0, valign: "top", isTextBox: true, objectName: nm(o.name || "text"), ...o }); }
function dot(s, x, y, d, fill) { s.addShape(pres.shapes.OVAL, { x: x - d / 2, y: y - d / 2, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: nm("dot") }); }
function full(s, key, alt) { s.addImage({ path: file(key), x: 0, y: 0, w: W, h: H, altText: alt || "배경", objectName: nm(key) }); }
// 방송 라벨 알약: ● LIVE  CH.01 ...
function pill(s, text, x, y, fill, color, o = {}) {
  const w = o.w || (0.5 + text.length * 0.095);
  rect(s, x, y, w, 0.32, fill, { round: 0.16 });
  dot(s, x + 0.2, y + 0.16, 0.1, color);
  t(s, text, { x: x + 0.32, y, w: w - 0.38, h: 0.32, fontFace: o.font || SANS, fontSize: o.size || 10.5, bold: true, color, valign: "middle", name: "pill" });
  return w;
}
// 일반 슬라이드: 챕터 알약 라벨 + 제목 placeholder
function slide(master, label, color, title, section, bgKey) {
  const s = pres.addSlide({ masterName: master, sectionTitle: section });
  if (bgKey) full(s, bgKey);
  pill(s, label, MX, 0.36, color, "FFFFFF");
  s.addText(title, { placeholder: "title" });
  return s;
}
function bullets(s, items, x, y, w, o = {}) {
  const color = o.color || "FFFFFF", mark = o.mark || PINK, size = o.size || 15, gap = o.gap || 0.62;
  items.forEach((it, i) => {
    dot(s, x + 0.08, y + i * gap + size / 72 * 0.62, 0.11, mark);
    t(s, it, { x: x + 0.3, y: y + i * gap, w: w - 0.3, h: gap, fontSize: size, color, lineSpacingMultiple: 1.1, name: "bullet" });
  });
}
// 플레이어 창: 16:9 화면 + 하단 재생 바(prog 0~1)
function player(s, key, x, y, w, color, o = {}) {
  const h = w * 9 / 16, bar = 0.42;
  rect(s, x - 0.08, y - 0.08, w + 0.16, h + bar + 0.16, o.frameColor || "0B0A18", { round: 0.12 });
  s.addImage({ path: file(key), x, y, w, h, altText: o.alt || key, objectName: nm(`player-${key}`), ...(o.url ? { hyperlink: { url: o.url, tooltip: "영상 보기" } } : {}) });
  const by = y + h + 0.08;
  t(s, "▶", { x: x + 0.05, y: by, w: 0.3, h: bar - 0.12, fontSize: 12, color: "FFFFFF", valign: "middle", name: "play" });
  const px = x + 0.42, pw = w - (o.time ? 1.55 : 0.6);
  rect(s, px, by + 0.12, pw, 0.06, "55507A", { round: 0.03 });
  if (o.prog) rect(s, px, by + 0.12, pw * o.prog, 0.06, color, { round: 0.03 });
  if (o.prog) dot(s, px + pw * o.prog, by + 0.15, 0.16, color);
  if (o.time) t(s, o.time, { x: x + w - 1.05, y: by, w: 1.05, h: bar - 0.12, fontFace: UI, fontSize: 10, bold: true, color: "D8D4F0", align: "right", valign: "middle", name: "time" });
  if (o.live) {
    rect(s, x + 0.2, y + 0.2, 0.95, 0.34, PINK, { round: 0.06 });
    t(s, "● LIVE", { x: x + 0.2, y: y + 0.2, w: 0.95, h: 0.34, fontFace: UI, fontSize: 12, bold: true, color: "FFFFFF", align: "center", valign: "middle", name: "live" });
  }
  return { x, y, w, h };
}
// 채팅 말풍선
function chat(s, user, text, x, y, w, h, o = {}) {
  rect(s, x, y, w, h, o.fill || PANEL2, { round: 0.1, tr: o.tr || 0 });
  t(s, user, { x: x + 0.2, y: y + 0.1, w: w - 0.4, h: 0.28, fontSize: 10, bold: true, color: o.userColor || PINK_L, name: "chat-user" });
  t(s, text, { x: x + 0.2, y: y + 0.38, w: w - 0.4, h: h - 0.45, fontSize: o.size || 13, bold: !!o.bold, color: o.color || "FFFFFF", lineSpacingMultiple: 1.15, name: "chat-text", ...(o.font ? { fontFace: o.font } : {}) });
}
function linkBtn(s, label, url, x, y, w, fill) {
  rect(s, x, y, w, 0.5, fill, { round: 0.25 });
  s.addText([{ text: label, options: { hyperlink: { url, tooltip: url }, color: "FFFFFF", bold: true } }], { x, y, w, h: 0.5, fontFace: SANS, fontSize: 13, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: nm("link") });
}

// 챕터 간지: 방송 시작 화면 (플레이어 + 채팅창), 재생 바가 챕터마다 차오름
function chapter(n, color, playKey, bgKey, title, sub, prog, section) {
  const s = pres.addSlide({ masterName: "PLAIN", sectionTitle: section });
  full(s, bgKey, "챕터 대표 장면을 흐리게 깐 배경");
  player(s, playKey, 0.75, 1.05, 7.9, color, { live: true, prog, time: `0${n} / 03`, alt: "챕터 대표 화면" });
  const x = 9.0, w = 3.6;
  rect(s, x, 1.0, w, 5.35, "0B0A18", { round: 0.12, tr: 15 });
  t(s, "LIVE CHAT", { x: x + 0.25, y: 1.15, w: 2.5, h: 0.3, fontFace: UI, fontSize: 11, bold: true, color: "B9B3DA", charSpacing: 3 });
  rect(s, x + 0.25, 1.52, w - 0.5, 0.02, "3A3560");
  t(s, `CH.0${n}`, { x: x + 0.25, y: 1.7, w: w - 0.5, h: 1.0, fontFace: UI, fontSize: 54, bold: true, color, valign: "middle", name: "ch-num" });
  chat(s, "임현", title, x + 0.25, 2.85, w - 0.5, 1.45, { size: 22, bold: true, userColor: color });
  chat(s, "임현", sub, x + 0.25, 4.45, w - 0.5, 1.6, { size: 13, userColor: color, color: "E6E2FA" });
  return s;
}

// =====================================================================
// 01 표지
pres.addSection({ title: "소개" });
{
  const s = pres.addSlide({ masterName: "PLAIN", sectionTitle: "소개" });
  full(s, "bg_cover", "영희의 무대 장면을 흐리게 깐 배경");
  player(s, "play0", 6.55, 1.25, 6.05, PINK, { live: true, prog: 0.0, time: "00:00", alt: "영희의 무대 장면" });
  pill(s, "ON AIR", MX, 1.3, PINK, "FFFFFF", { font: UI, size: 11, w: 1.15 });
  t(s, "AI CONTENT CREATOR PORTFOLIO", { x: MX, y: 1.85, w: 5.6, h: 0.35, fontFace: UI, fontSize: 13, bold: true, color: PINK_L, charSpacing: 3 });
  t(s, "임현", { x: MX, y: 2.25, w: 5.5, h: 1.2, fontSize: 60, bold: true, color: "FFFFFF", valign: "middle" });
  t(s, "화면 너머의 작은 응원을 이야기로 연결하는 크리에이터", { x: MX, y: 3.55, w: 5.6, h: 0.8, fontSize: 17, color: "E6E2FA", lineSpacingMultiple: 1.2 });
  [["CH.01", PINK, "AI 애니메이션 「남들과 다른, 그래서 평범」"], ["CH.02", SKY_L, "캐릭터 IP 「TENSHI」"], ["CH.03", RAIN_L, "학원물 캐릭터 / 숏폼 기획"]].forEach(([c, col, x], i) => {
    const y = 4.6 + i * 0.5;
    t(s, c, { x: MX, y, w: 0.9, h: 0.4, fontFace: UI, fontSize: 13, bold: true, color: col, valign: "middle" });
    t(s, x, { x: MX + 0.95, y, w: 4.8, h: 0.4, fontSize: 14, color: "FFFFFF", valign: "middle" });
  });
  t(s, "서울시 매력일자리 AI 콘텐츠 크리에이터 양성과정  ·  2026", { x: MX, y: 6.7, w: 6.5, h: 0.3, fontSize: 10, color: "9E98C2" });
}

// 02 목차 (재생목록)
{
  const s = slide("NIGHT", "PLAYLIST", PINK, "목차", "소개", "bg_room");
  const rows = [
    ["play1", "CH.01", PINK, "남들과 다른, 그래서 평범", "AI 애니메이션 / 유튜브 숏폼 / 일상, 성장", 1 / 3],
    ["play2", "CH.02", SKY, "TENSHI", "캐릭터 IP / 이모티콘 / 굿즈 / 일상툰", 2 / 3],
    ["play3", "CH.03", RAIN, "학원물 캐릭터 / 숏폼 기획", "캐릭터 삼면도 / 표정 시트 / 가사 스토리 시트", 1],
  ];
  rows.forEach(([k, c, col, title, meta, p], i) => {
    const x = MX + i * 4.05;
    player(s, k, x + 0.08, 1.95, 3.6, col, { prog: p, alt: title });
    t(s, c, { x, y: 4.6, w: 1.2, h: 0.35, fontFace: UI, fontSize: 13, bold: true, color: col === RAIN ? RAIN_L : col });
    t(s, title, { x, y: 4.95, w: 3.8, h: 0.5, fontSize: 19, bold: true, color: "FFFFFF" });
    t(s, meta, { x, y: 5.5, w: 3.8, h: 0.6, fontSize: 12, color: "C9C4E6" });
  });
}

// 03 역량
{
  const s = slide("CREAM", "PROFILE", PINK, "경험에서 출발한 공감형 스토리와 캐릭터", "소개", "bg_morning");
  const cards = [
    ["경험 기반 기획", "실직 후 무기력했던 시절을 솔직하게 돌아본 진정성 있는 서사 기획", PINK],
    ["조명/색 연출", "시간대와 빛의 회복으로 인물의 감정 변화를 시각화한 컷 설계", SUN],
    ["캐릭터 IP", "컬러 팔레트와 디테일을 정의한 시트로 이모티콘/굿즈까지 확장", SKY],
    ["2D 작화 + AI", "직접 그린 작화와 AI툴을 결합해 창작/기획/연출 역량 극대화", LAV],
  ];
  cards.forEach(([h, b, col], i) => {
    const x = MX + i * 3.0, y = 1.95;
    rect(s, x, y, 2.8, 4.4, "FFFFFF", { round: 0.12, line: "EADFD2" });
    pill(s, `0${i + 1}`, x + 0.25, y + 0.3, col, "FFFFFF", { font: UI, size: 11, w: 0.75 });
    t(s, h, { x: x + 0.25, y: y + 1.0, w: 2.35, h: 0.5, fontSize: 19, bold: true, color: INK });
    t(s, b, { x: x + 0.25, y: y + 1.65, w: 2.3, h: 2.4, fontSize: 14, color: "4A4458", lineSpacingMultiple: 1.25 });
  });
}

// =====================================================================
// CHAPTER 01
const L1 = "CH.01  남들과 다른, 그래서 평범", S1 = "01 남들과 다른, 그래서 평범";
pres.addSection({ title: S1 });
chapter(1, PINK, "play1", "ch1_bg", "남들과 다른,\n그래서 평범", "AI 애니메이션 / 유튜브 숏폼\n무기력한 철수가 버튜버 영희의 방송에서 다시 문을 나설 용기를 얻는 이야기", 1 / 3, S1);

// 작품 개요
{
  const s = slide("NIGHT", L1, PINK, "\"이런 식의 구원도 있을 수 있다.\"", S1, "bg_room");
  player(s, "a1_2a", MX, 1.85, 6.0, PINK, { prog: 0.1, alt: "커튼 사이로 빛이 새는 어두운 방, 컴퓨터 앞으로 가는 철수" });
  chat(s, "로그라인", "나아갈 수 있는 '용기'를 주는 그런 존재가 누구에게는 필요하다.", 7.15, 1.85, 5.45, 1.2, { size: 15, bold: true });
  bullets(s, [
    "실직 후 무기력했던 작가 자신의 경험을 바탕으로 한 진정성 있는 서사",
    "비슷한 좌절을 겪는 이들에게 \"혼자가 아니다\"라는 메시지 전달",
  ], 7.15, 3.25, 5.45, { size: 14, gap: 0.78 });
  const info = [["장르", "일상 / 성장"], ["플랫폼", "유튜브 숏폼"], ["타깃", "서브컬처에 관심 있는 20~30대 사회초년생 / 실직자"]];
  info.forEach(([a, b], i) => {
    const y = 5.0 + i * 0.45;
    t(s, a, { x: 7.15, y, w: 1.1, h: 0.4, fontSize: 12, bold: true, color: PINK_L, valign: "middle" });
    t(s, b, { x: 8.25, y, w: 4.35, h: 0.4, fontSize: 12.5, color: "FFFFFF", valign: "middle" });
  });
}

// 캐릭터
{
  const s = slide("CREAM", L1, PINK, "두 모습으로 설계한 두 인물", S1, "bg_morning");
  const rows = [
    ["철수 (30세)", "정규직 전환에 탈락한 뒤 무기력한 일상을 보내는 주인공", [["cs_out", "실직자"], ["cs_job", "구직자"]], "6A5ACD"],
    ["영희 (자칭 1000세)", "어색한 첫 방송에서 팬들의 응원으로 용기를 얻는 신인 버튜버", [["yh_vt", "버튜버"], ["yh_idol", "아이돌"]], PINK],
  ];
  // 철수: 세로 시트 두 장 / 영희: 가로 시트 두 장
  rect(s, MX, 1.8, 4.3, 4.85, "FFFFFF", { round: 0.1, line: "EADFD2" });
  img(s, "cs_out", MX + 0.1, 1.9, 2.05, 3.6, { alt: "철수 실직자 시트" });
  img(s, "cs_job", MX + 2.15, 1.9, 2.05, 3.6, { alt: "철수 구직자 시트" });
  t(s, rows[0][0], { x: MX + 0.25, y: 5.55, w: 3.9, h: 0.4, fontSize: 17, bold: true, color: INK });
  t(s, "실직자 / 구직자 두 버전으로 변화 전후를 의상으로 표현", { x: MX + 0.25, y: 5.95, w: 3.9, h: 0.6, fontSize: 12, color: "5A5468" });
  rect(s, 5.2, 1.8, 7.4, 4.85, "FFFFFF", { round: 0.1, line: "EADFD2" });
  img(s, "yh_vt", 5.3, 1.85, 7.2, 1.85, { alt: "영희 버튜버 시트" });
  img(s, "yh_idol", 5.3, 3.7, 7.2, 1.85, { alt: "영희 아이돌 시트" });
  t(s, rows[1][0], { x: 5.45, y: 5.55, w: 3, h: 0.4, fontSize: 17, bold: true, color: PINK });
  t(s, "방송용 버튜버 의상과 무대용 아이돌 의상으로 감정의 정점을 시각적으로 분리", { x: 5.45, y: 5.95, w: 7.0, h: 0.6, fontSize: 12, color: "5A5468" });
}

// 빛의 회복 설계
{
  const s = slide("NIGHT", L1, PINK, "빛이 돌아오는 방, 감정 변화를 조명으로 설계", S1, "bg_room");
  const steps = [
    ["a1_3a", "아침", "커튼에 막힌 햇빛과 모니터 반사광"],
    ["a2_8a", "심야", "핸드폰 불빛만 남은 완전한 어둠"],
    ["a3_5", "저녁", "화면 속 무대 조명이 얼굴에 반사"],
    ["a4_1", "아침", "커튼을 걷고 처음 방을 채우는 햇살"],
  ];
  steps.forEach(([k, time, d], i) => {
    const x = MX + i * 3.0;
    img(s, k, x, 1.95, 2.8, 1.58, { frame: "0B0A18", alt: d });
    const col = [MUTED, "3F3A78", LAV, SUN][i];
    rect(s, x, 3.75, 2.8, 0.08, col);
    pill(s, time, x, 3.98, PANEL2, "FFFFFF", { w: 0.95 });
    t(s, d, { x, y: 4.42, w: 2.8, h: 0.7, fontSize: 12.5, color: "E6E2FA" });
    if (i < 3) t(s, "▶", { x: x + 2.8, y: 2.55, w: 0.2, h: 0.35, fontSize: 10, color: PINK, align: "center" });
  });
  rect(s, MX, 5.35, 11.9, 1.3, PANEL, { round: 0.1 });
  t(s, "작가의 공통 비주얼 사양", { x: MX + 0.3, y: 5.48, w: 4, h: 0.3, fontSize: 11, bold: true, color: PINK_L });
  t(s, "\"커튼과 짐으로 자연광이 차단된 상태를 기본으로 하며, 이야기가 진행될수록 인공조명과 자연광이 점차 되돌아오는 방향으로 설계함.\"", { x: MX + 0.3, y: 5.82, w: 11.3, h: 0.7, fontSize: 13.5, color: "FFFFFF", lineSpacingMultiple: 1.15 });
}

// 희로애락 교차 편집
{
  const s = slide("NIGHT", L1, PINK, "희로애락 방송과 철수의 하루를 교차 편집", S1, "bg_stream");
  const top = [["a2_1", "희", "\"끝까지 잘 부탁할게!\""], ["a2_3", "노", "\"나 지금 완전 진지하거든?!\""], ["a2_5", "애", "\"…노래도 너무 슬퍼\""], ["a2_7a", "락", "\"드디어 깼다?!\""]];
  top.forEach(([k, e, d], i) => {
    const x = MX + i * 3.0;
    img(s, k, x, 1.9, 2.8, 1.58, { frame: PINK, alt: `영희 방송 ${e}` });
    pill(s, `LIVE · ${e}`, x + 0.1, 1.98, PINK, "FFFFFF", { w: 1.05, size: 9.5 });
    t(s, d, { x, y: 3.6, w: 2.8, h: 0.35, fontSize: 12, bold: true, color: PINK_L });
  });
  const bottom = [["a2_2", "점심, 컵라면을 먹으며"], ["a2_8a", "심야, 잠들기 직전까지"]];
  bottom.forEach(([k, d], i) => {
    const x = MX + i * 3.0;
    img(s, k, x, 4.2, 2.8, 1.58, { frame: "0B0A18", alt: `철수 ${d}` });
    t(s, d, { x, y: 5.9, w: 2.8, h: 0.35, fontSize: 12, color: "C9C4E6" });
  });
  bullets(s, [
    "영희의 방송은 시간과 무관한 파스텔 핑크 스튜디오로 고정, 철수의 방과 다른 가상 공간임을 표현",
    "방송이 철수의 하루 루틴 속으로 스며드는 과정을 교차 편집으로 전개",
  ], 6.85, 4.25, 5.75, { size: 13.5, gap: 0.95, mark: PINK });
}

// 무대
{
  const s = slide("NIGHT", L1, PINK, "버튜버에서 아이돌로, 화면 너머의 무대", S1, "bg_stream");
  img(s, "a3_1", MX, 1.9, 3.7, 2.08, { frame: "0B0A18", alt: "영희의 SNS 예고 게시글" });
  t(s, "SNS 예고", { x: MX, y: 4.05, w: 3.7, h: 0.3, fontSize: 11.5, bold: true, color: PINK_L });
  player(s, "a3_3", 4.65, 1.9, 7.95, PINK, { live: true, prog: 0.62, time: "S#3", alt: "핑크와 골드 조명 속 무대 위 영희" });
  img(s, "a3_5", MX, 4.55, 3.7, 2.08, { frame: "0B0A18", alt: "모니터 빛이 비친 철수의 놀란 얼굴" });
  t(s, "화면 속 빛이 반사된 철수의 얼굴", { x: MX, y: 6.68, w: 3.7, h: 0.3, fontSize: 11, color: "C9C4E6" });
}

// 다시 시작할 준비
{
  const s = slide("CREAM", L1, PINK, "다시 시작할 준비, 일상을 되찾는 네 장면", S1, "bg_morning");
  const cuts = [["a4_1", "방을 정리하다 발견한 정장"], ["a4_2", "미용실에서 다듬는 머리"], ["a4_3", "스탠드 불빛 아래 쓰는 이력서"], ["a4_4", "거울 앞에서 다시 매는 넥타이"]];
  cuts.forEach(([k, d], i) => {
    const x = MX + (i % 2) * 6.0, y = 1.8 + Math.floor(i / 2) * 2.5;
    img(s, k, x, y, 3.6, 2.03, { frame: "FFFFFF", alt: d });
    pill(s, `S#4 C#${i + 1}`, x + 3.8, y + 0.1, PINK, "FFFFFF", { font: UI, w: 1.15 });
    t(s, d, { x: x + 3.8, y: y + 0.6, w: 2.0, h: 1.2, fontSize: 14, bold: true, color: INK, lineSpacingMultiple: 1.2 });
  });
  t(s, "대사 없이 행동과 조명 변화만으로 마음을 다잡는 과정을 전달하는 절제된 연출", { x: MX, y: 6.62, w: 11.9, h: 0.35, fontSize: 13, bold: true, color: PINK });
}

// 수미상관 엔딩
{
  const s = slide("CREAM", L1, PINK, "과거의 나와 작별하는 수미상관 엔딩", S1, "bg_morning");
  const cuts = [["a1_1b", "첫 컷, 닫히는 문"], ["a5_3r", "컵라면을 먹는 과거의 철수"], ["a5_4", "다녀오라는 손짓"], ["a5_6_4", "빛이 쏟아진 뒤 닫힌 문"]];
  cuts.forEach(([k, d], i) => {
    const x = MX + i * 3.0;
    img(s, k, x, 1.95, 2.8, 1.58, { frame: i === 3 ? SUN : "FFFFFF", alt: d });
    t(s, d, { x, y: 3.65, w: 2.8, h: 0.35, fontSize: 12.5, bold: true, color: INK });
    if (i === 0) pill(s, "S#1", x + 0.1, 2.03, MUTED, "FFFFFF", { font: UI, w: 0.7 });
    else pill(s, "S#5", x + 0.1, 2.03, PINK, "FFFFFF", { font: UI, w: 0.7 });
  });
  bullets(s, [
    "과거의 자신과 현재의 자신을 한 공간에 마주 세운 연출로 내적 변화를 시각화",
    "첫 컷의 문과 마지막 컷의 문을 대응시킨 수미상관 구조로 여운 극대화",
    "가장 밝은 햇살이 문틈으로 쏟아지는 순간으로 새로운 시작 표현",
  ], MX, 4.45, 11.9, { color: "3A3448", size: 14.5, gap: 0.62, mark: PINK });
}

// AI 워크플로우
{
  const s = slide("NIGHT", L1, PINK, "컷 단위 비주얼 사양으로 일관성을 지킨 AI 제작 공정", S1, "bg_room");
  img(s, "wf_claude", MX, 1.9, 5.85, 3.1, { frame: "0B0A18", alt: "Claude 프롬프트 작성 화면" });
  img(s, "wf_premiere", 6.75, 1.9, 5.85, 3.1, { frame: "0B0A18", alt: "프리미어 편집 화면" });
  pill(s, "Claude", MX, 5.15, PINK, "FFFFFF", { font: UI, w: 1.0 });
  pill(s, "Premiere Pro", 6.75, 5.15, LAV, "FFFFFF", { font: UI, w: 1.45 });
  t(s, "카메라 / 시간대 / 색감·조명 / 동작 / 효과음을 컷마다 명시한 시나리오로 프롬프트 체계화", { x: MX, y: 5.6, w: 5.85, h: 0.7, fontSize: 12.5, color: "E6E2FA" });
  t(s, "생성한 컷과 효과음, BGM을 편집해 감정 흐름에 맞춘 리듬 완성", { x: 6.75, y: 5.6, w: 5.85, h: 0.7, fontSize: 12.5, color: "E6E2FA" });
  t(s, "얼굴 각도가 바뀌어도 외형과 신체 비율이 변하지 않도록 고정한 캐릭터 일관성 관리", { x: MX, y: 6.4, w: 11.9, h: 0.35, fontSize: 13, bold: true, color: PINK_L });
}

// 영상
{
  const s = slide("NIGHT", L1, PINK, "최종 편집본", S1, "bg_stream");
  player(s, "poster_mv", MX, 1.9, 7.4, PINK, { prog: 1, time: "FINAL", url: LINK_MV, alt: "남들과 다른, 그래서 평범 영상 보기" });
  t(s, "화면을 클릭하면 영상이 재생됩니다", { x: MX, y: 6.55, w: 7.4, h: 0.3, fontSize: 11, color: "9E98C2" });
  const x = 8.65;
  chat(s, "영희", "오늘도 다들 파이팅이에요!\n저도 항상 응원할게요!", x, 1.9, 3.95, 1.35, { size: 15, bold: true });
  chat(s, "영희", "헤헤, 다음에도 잘 부탁해.", x, 3.4, 3.95, 0.95, { size: 14 });
  t(s, "남들과 다른, 그래서 평범", { x, y: 4.6, w: 3.95, h: 0.45, fontSize: 18, bold: true, color: "FFFFFF" });
  t(s, "AI 애니메이션 / 유튜브 숏폼", { x, y: 5.05, w: 3.95, h: 0.35, fontSize: 12.5, color: "C9C4E6" });
  linkBtn(s, "▶  영상 보기", LINK_MV, x, 5.6, 2.45, PINK);
}

// =====================================================================
// CHAPTER 02 TENSHI
const L2 = "CH.02  TENSHI", S2 = "02 TENSHI";
pres.addSection({ title: S2 });
chapter(2, SKY_L, "play2", "ch2_bg", "TENSHI", "캐릭터 IP / 이모티콘 / 굿즈 / 일상툰\n천사 고리와 날개를 시그니처로 한 캐릭터 브랜드", 2 / 3, S2);

// 캐릭터 시트 + 로고
{
  const s = slide("SKYBG", L2, SKY, "컬러 팔레트와 디테일까지 정의한 캐릭터 시트", S2, "bg_sky");
  rect(s, MX, 1.8, 7.6, 5.05, "FFFFFF", { round: 0.1, line: "CFE6F7" });
  img(s, "t_sheet", MX + 0.1, 1.85, 7.4, 4.95, { alt: "TENSHI 캐릭터 시트 최종본" });
  rect(s, 8.6, 1.8, 4.0, 2.4, "FFFFFF", { round: 0.1, line: "CFE6F7" });
  img(s, "t_logo", 8.7, 1.85, 3.8, 2.3, { alt: "TENSHI 캐릭터 로고" });
  const pal = [["BDE8FF", "머리색"], ["1F8FE6", "눈 색상"], ["FFFFFF", "원피스"]];
  pal.forEach(([c, n], i) => {
    const x = 8.6 + i * 1.35;
    rect(s, x, 4.4, 1.2, 0.55, c, { round: 0.08, line: "B5D6EE" });
    t(s, `${n}\n#${c}`, { x, y: 5.0, w: 1.3, h: 0.5, fontSize: 9.5, color: "35577A" });
  });
  bullets(s, [
    "천사 고리 / 날개 / 하트 장식을 반복 요소로 브랜드 각인",
    "정면 / 45도 / 후면과 감정별 표정 예시로 활용 기준 확립",
  ], 8.6, 5.65, 4.0, { color: "143A66", mark: SKY, size: 12, gap: 0.55 });
}

// 비율 변형
{
  const s = slide("SKYBG", L2, SKY, "SD에서 6등신까지, 비율 변형으로 넓힌 활용 범위", S2, "bg_sky");
  rect(s, MX, 1.8, 7.9, 3.75, "FFFFFF", { round: 0.1, line: "CFE6F7" });
  img(s, "t_sd", MX + 0.1, 1.85, 7.7, 3.65, { alt: "TENSHI SD 삼면도" });
  rect(s, 8.9, 1.8, 3.7, 3.75, "FFFFFF", { round: 0.1, line: "CFE6F7" });
  img(s, "t_six", 9.0, 1.85, 3.5, 3.65, { alt: "TENSHI 6등신 캐릭터" });
  pill(s, "SD 삼면도", MX, 5.7, SKY, "FFFFFF", { w: 1.25 });
  pill(s, "6등신", 8.9, 5.7, SKY, "FFFFFF", { w: 0.9 });
  t(s, "이모티콘 / 굿즈용 SD와 일러스트용 6등신 비율을 함께 설계해 매체별 활용도 확보", { x: MX, y: 6.2, w: 11.9, h: 0.4, fontSize: 14, bold: true, color: "143A66" });
}

// 이모티콘 (움직이는 GIF)
{
  const s = slide("SKYBG", L2, SKY, "감정별 움직이는 이모티콘 14종", S2, "bg_sky");
  const names = ["이 몸 등장", "최고", "사랑해", "축하해~", "파이팅!", "건배", "미안해~", "뿌에엥~", "싫은데", "날 속인거니?", "귀찮아", "졸려…", "힘들어", "쉬면서해"];
  names.forEach((n, i) => {
    const col = i % 7, row = Math.floor(i / 7);
    const x = MX + col * 1.72, y = 1.8 + row * 2.45;
    rect(s, x, y, 1.6, 1.6, "FFFFFF", { round: 0.12, line: "CFE6F7" });
    img(s, `emo${String(i + 1).padStart(2, "0")}`, x + 0.08, y + 0.08, 1.44, 1.44, { alt: `TENSHI 이모티콘 ${n}` });
    t(s, n, { x, y: y + 1.65, w: 1.6, h: 0.3, fontSize: 11, bold: true, color: "143A66", align: "center" });
  });
  t(s, "긍정 / 위로 / 장난 / 휴식까지 일상 대화에 바로 쓰는 감정 라인업 구성, 슬라이드 쇼에서 움직임 확인 가능", { x: MX, y: 6.65, w: 11.9, h: 0.35, fontSize: 12, color: "35577A" });
}

// 굿즈
{
  const s = slide("SKYBG", L2, SKY, "피규어와 인형으로 확장한 굿즈 시안", S2, "bg_sky");
  const g = [["g_nendo", "넨도로이드", "표정 파츠와 소품을 갖춘 피규어 시안"], ["g_nui", "누이구루미", "부드러운 봉제 질감의 인형 시안"], ["g_puchi", "푸치 사이즈", "손바닥 크기의 미니 인형 시안"]];
  g.forEach(([k, n, d], i) => {
    const x = MX + i * 4.0;
    rect(s, x, 1.8, 3.8, 3.4, "FFFFFF", { round: 0.1, line: "CFE6F7" });
    img(s, k, x + 0.1, 1.9, 3.6, 3.2, { alt: `TENSHI 굿즈 ${n}` });
    t(s, n, { x, y: 5.35, w: 3.8, h: 0.4, fontSize: 18, bold: true, color: "143A66" });
    t(s, d, { x, y: 5.8, w: 3.8, h: 0.4, fontSize: 12.5, color: "35577A" });
  });
  t(s, "캐릭터 시트의 컬러와 디테일을 입체 상품으로 일관되게 구현", { x: MX, y: 6.4, w: 11.9, h: 0.35, fontSize: 13, bold: true, color: SKY });
}

// 일상툰 + 홍보영상
{
  const s = slide("SKYBG", L2, SKY, "일상툰과 30초 홍보영상", S2, "bg_sky");
  img(s, "toon1", MX, 1.8, 4.9, 4.9, { frame: "FFFFFF", alt: "TENSHI 일상툰 1" });
  const x = 5.95;
  player(s, "poster_tenshi", x, 1.85, 6.6, SKY, { prog: 0.4, time: "0:30", url: LINK_TENSHI, frameColor: "0E3560", alt: "TENSHI 30초 홍보영상 보기" });
  t(s, "\"…너라도 있어서 다행이야\"", { x, y: 5.85, w: 4.2, h: 0.4, fontSize: 15, bold: true, color: "143A66" });
  t(s, "어두운 방의 고독한 순간에 다가가는 위로형 일상툰", { x, y: 6.3, w: 4.2, h: 0.4, fontSize: 12, color: "35577A" });
  linkBtn(s, "▶  홍보영상 보기", LINK_TENSHI, 10.35, 5.9, 2.2, SKY);
}

// =====================================================================
// CHAPTER 03 학원물
const L3 = "CH.03  학원물 캐릭터 / 숏폼 기획", S3 = "03 학원물 캐릭터";
pres.addSection({ title: S3 });
chapter(3, RAIN_L, "play3", "ch3_bg", "학원물 캐릭터\n/ 숏폼 기획", "캐릭터 삼면도 / 표정 시트 / 가사 스토리 시트\n비 오는 교실에서 시작하는 학원물", 1, S3);

// 캐릭터
{
  const s = slide("NIGHT", L3, RAIN, "교복과 소품으로 성격을 드러낸 남녀 캐릭터", S3, "bg_rain");
  rect(s, MX, 1.8, 7.9, 2.35, "FFFFFF", { round: 0.08 });
  img(s, "s_m3", MX + 0.05, 1.82, 7.8, 2.31, { alt: "남자 캐릭터 삼면도" });
  rect(s, MX, 4.3, 7.9, 2.35, "FFFFFF", { round: 0.08 });
  img(s, "s_f3", MX + 0.05, 4.32, 7.8, 2.31, { alt: "여자 캐릭터 삼면도" });
  rect(s, 8.85, 1.8, 3.75, 4.85, "FFFFFF", { round: 0.08 });
  img(s, "s_mfront", 8.9, 1.85, 1.2, 4.75, { alt: "남자 캐릭터 정면" });
  img(s, "s_mside", 10.1, 1.85, 1.2, 4.75, { alt: "남자 캐릭터 좌측면" });
  img(s, "s_ffront", 11.35, 1.85, 1.2, 4.75, { alt: "여자 캐릭터 정면" });
}

// 표정 시트
{
  const s = slide("NIGHT", L3, RAIN, "18종 이상의 표정 시트로 감정 연기 범위 확보", S3, "bg_rain");
  rect(s, MX, 1.8, 5.85, 3.95, "FFFFFF", { round: 0.08 });
  img(s, "s_mface", MX + 0.05, 1.85, 5.75, 3.85, { alt: "남자 캐릭터 표정 시트" });
  rect(s, 6.75, 1.8, 5.85, 3.95, "FFFFFF", { round: 0.08 });
  img(s, "s_fface", 6.8, 1.85, 5.75, 3.85, { alt: "여자 캐릭터 표정 시트" });
  bullets(s, [
    "기쁨 / 미소 / 당황 / 부끄러움 / 충격 / 체념 등 미세한 감정까지 세분화",
    "표정 단위로 정리한 시트로 컷마다 캐릭터 일관성과 연기 디테일 유지",
  ], MX, 5.95, 11.9, { mark: RAIN_L, size: 13.5, gap: 0.45 });
}

// 가사 스토리 시트
{
  const s = slide("NIGHT", L3, RAIN, "가사와 장면을 1:1로 매칭한 스토리 시트 「아도레나」", S3, "bg_rain");
  img(s, "s_class", MX, 1.85, 4.6, 2.59, { frame: "0B0A18", alt: "비 오는 교실 배경" });
  t(s, "비 오는 교실 배경", { x: MX, y: 4.55, w: 4.6, h: 0.3, fontSize: 11, color: "C9C4E6" });
  bullets(s, [
    "가사 한 구절마다 인물 동선과 장면을 지정한 뮤직비디오형 구성",
    "학생회와 부활동 캐릭터 군상으로 세계관 확장",
  ], MX, 5.05, 4.6, { mark: RAIN_L, size: 13, gap: 0.8 });
  const rows = [
    ["그래, 그날 이후 모든 게 완전히 달라졌어", "교문 앞, 길을 몰라 헤매는 후야"],
    ["번뜩이는 마음을 따라 뛰어든 나의", "학생회 회의, 회장 / 회계 / 서기 구도"],
    ["자, 이제 움직여!", "학생회장과 부활동 시찰 진행"],
    ["큰맘 먹고 짧게 자른 머리는 나 자신에게 건 마법이야", "머리를 자른 학생회장"],
    ["널 꼭 내 쪽으로 돌아보게 만들 테니까", "후야의 고백, 얼굴 클로즈업"],
  ];
  const x = 5.65, w1 = 3.65, w2 = 3.3;
  rect(s, x, 1.85, w1 + w2, 0.42, RAIN, { round: 0.06 });
  t(s, "가사", { x: x + 0.2, y: 1.85, w: w1, h: 0.42, fontSize: 12, bold: true, color: "FFFFFF", valign: "middle" });
  t(s, "장면", { x: x + w1 + 0.2, y: 1.85, w: w2, h: 0.42, fontSize: 12, bold: true, color: "FFFFFF", valign: "middle" });
  rows.forEach(([a, b], i) => {
    const y = 2.35 + i * 0.86;
    rect(s, x, y, w1 + w2, 0.78, PANEL, { round: 0.06, tr: 10 });
    t(s, a, { x: x + 0.2, y, w: w1 - 0.3, h: 0.78, fontSize: 12, color: "FFFFFF", valign: "middle", italic: true });
    t(s, b, { x: x + w1 + 0.2, y, w: w2 - 0.3, h: 0.78, fontSize: 12, bold: true, color: RAIN_L, valign: "middle" });
  });
}

// 2D 작화 역량
pres.addSection({ title: "2D 작화" });
{
  const s = slide("NIGHT", "BONUS  2D ANIMATION", LAV, "직접 그리는 2D 작화와 프레임 애니메이션", "2D 작화", "bg_room");
  img(s, "gif_turn", MX, 1.85, 5.85, 3.29, { frame: "FFFFFF", alt: "캐릭터 턴어라운드 애니메이션" });
  img(s, "gif_practice", 6.75, 1.85, 5.85, 3.29, { frame: "FFFFFF", alt: "애니메이션 활용 실습" });
  pill(s, "턴어라운드", MX, 5.3, LAV, "FFFFFF", { w: 1.25 });
  pill(s, "애니메이션 실습", 6.75, 5.3, LAV, "FFFFFF", { w: 1.6 });
  bullets(s, [
    "클립스튜디오에서 러프 / 선화 / 프레임 타이밍까지 직접 작업한 손그림 애니메이션",
    "직접 그리는 작화 기초 위에 AI툴을 사용하여 창작/기획/연출 역량 극대화",
  ], MX, 5.85, 11.9, { mark: LAV, size: 13.5, gap: 0.48 });
}

// =====================================================================
// 엔딩
pres.addSection({ title: "마무리" });
{
  const s = pres.addSlide({ masterName: "PLAIN", sectionTitle: "마무리" });
  full(s, "bg_end", "햇살이 비치는 현관을 흐리게 깐 배경");
  rect(s, 3.4, 1.7, 6.53, 4.1, "FFFFFF", { round: 0.16, tr: 8 });
  pill(s, "STREAM ENDED", 5.55, 2.05, PINK, "FFFFFF", { font: UI, size: 11, w: 2.2 });
  t(s, "시청해 주셔서 감사합니다", { x: 3.4, y: 2.65, w: 6.53, h: 0.8, fontSize: 32, bold: true, color: INK, align: "center", valign: "middle" });
  t(s, "임현", { x: 3.4, y: 3.55, w: 6.53, h: 0.5, fontSize: 20, bold: true, color: PINK, align: "center" });
  t(s, "AI 애니메이션 / 캐릭터 IP / 학원물 캐릭터 기획", { x: 3.4, y: 4.15, w: 6.53, h: 0.4, fontSize: 13, color: "5A5468", align: "center" });
  rect(s, 4.2, 4.95, 4.93, 0.06, "E2DCEF", { round: 0.03 });
  rect(s, 4.2, 4.95, 4.93, 0.06, PINK, { round: 0.03 });
  dot(s, 9.13, 4.98, 0.18, PINK);
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();

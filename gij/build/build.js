// 고일준 포트폴리오: 라푼젤 SF 재해석 애니메이션 기획 (가제)
// node build.js  ->  ../고일준_포트폴리오.pptx
// 모티프: 유럽 그래픽 노블(BD) 지면. 잉크 테두리 패널 + 흰 거터, 내레이션 캡션 박스 라벨, 케이블 구분선, 황금 큐브 엔딩과 FIN
// 색: 양피지 크림 / 잉크 / 새벽 청록 / 큐브 황금 (작가가 설정한 청록과 황금의 보색 대비)
const path = require("path");
const pptxgen = require("pptxgenjs");
// 테마 색 적용 스크립트(pptx 스킬)는 있을 때만 사용. 없으면 건너뛰며 슬라이드 모양은 같음
let applyTheme = null;
try { if (process.env.PPTX_SKILL) ({ applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js")); } catch (e) { applyTheme = null; }

const IMG = path.join(__dirname, "../work/img");
const DIMS = require("../work/dims.json");
const OUT = path.join(__dirname, "../고일준_포트폴리오.pptx");

const V = (id) => `https://drive.google.com/file/d/${id}/view`;
const LINK_ANIM = V("1qgZ3MNQNXUd2Jvw82OL8JyUNG1nOVp1b");
const LINK_ROT = V("1tJBJx1PWxtNXg1bALg6v3TdEt8SJ0HH2");

const SANS = "Malgun Gothic";
const SERIF = "Palatino Linotype"; // 라틴 제목 (BD 표지 느낌)
const PARCH = "EFE6D6", PARCH2 = "E4D7C2", INK = "1F1B16", INK2 = "4A4236", GRAY = "7A6E5E";
const TEAL = "0E2225", TEAL2 = "1B3B3F", TEAL_L = "5E9A9C", GOLD = "C9952F", GOLD_L = "E6C27A", WHITE = "FFFFFF";
const SWATCH = ["D7C5B1", "B99E7D", "C3A47E", "978B78", "898576", "8A898B", "A29384", "857664"]; // R 캐릭터 시트 컬러 팔레트

const THEME = {
  name: "Ligne Claire",
  headFontFace: SANS,
  bodyFontFace: SANS,
  colors: { dk1: INK, lt1: WHITE, dk2: TEAL, lt2: PARCH, accent1: TEAL2, accent2: GOLD, accent3: TEAL_L, accent4: GOLD_L, accent5: INK2, accent6: GRAY, hlink: TEAL2, folHlink: TEAL2 },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: SANS, bodyFontFace: SANS };
pres.author = "고일준";
pres.title = "고일준 포트폴리오";
const W = 13.333, H = 7.5, MX = 0.7;

const footer = (c) => ({ text: { text: "고일준  ·  AI 콘텐츠 크리에이터 포트폴리오", options: { x: MX, y: 7.08, w: 6, h: 0.28, fontFace: SANS, fontSize: 9, color: c, margin: 0 } } });
const titlePh = (c) => ({ placeholder: { options: { name: "title", type: "title", x: MX, y: 0.78, w: 11.9, h: 0.72, fontFace: SANS, fontSize: 26, bold: true, color: c, align: "left", valign: "middle", margin: 0 }, text: "" } });
const num = (c) => ({ x: 12.03, y: 7.08, w: 0.6, h: 0.28, fontFace: SERIF, fontSize: 11, bold: true, color: c, align: "right" });
pres.defineSlideMaster({ title: "PARCH", background: { color: PARCH }, objects: [titlePh(INK), footer("9A8E7C")], slideNumber: num(GOLD) });
pres.defineSlideMaster({ title: "NIGHT", background: { color: TEAL }, objects: [titlePh(PARCH), footer("7F9A98")], slideNumber: num(GOLD_L) });
pres.defineSlideMaster({ title: "PLAIN", background: { color: TEAL }, objects: [] });

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
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: fill ? { color: fill, transparency: o.tr || 0 } : { type: "none" }, line: o.line ? { color: o.line, width: o.lw || 1 } : { type: "none" }, shadow: o.shadow ? { type: "outer", color: "000000", opacity: 0.2, blur: 5, offset: 2, angle: 90 } : undefined, objectName: nm(o.name || "rect") });
}
function diamond(s, x, y, d, color) {
  s.addShape(pres.shapes.DIAMOND, { x, y, w: d, h: d, fill: { color }, line: { type: "none" }, objectName: nm("diamond") });
}
function t(s, text, o) { s.addText(text, { fontFace: SANS, color: INK, margin: 0, valign: "top", isTextBox: true, objectName: nm(o.name || "text"), ...o }); }
function full(s, key, alt) { s.addImage({ path: file(key), x: 0, y: 0, w: W, h: H, altText: alt || "배경", objectName: nm(key) }); }
// BD 패널: 흰 거터 + 잉크 테두리 안에 원본 비율 그림
function panel(s, key, x, y, w, h, o = {}) {
  let f = fit(key, x, y, w, h);
  if (o.top) f = { ...f, y };
  if (o.left) f = { ...f, x };
  const g = 0.07;
  rect(s, f.x - g, f.y - g, f.w + 2 * g, f.h + 2 * g, o.dark ? TEAL2 : WHITE, { line: o.dark ? GOLD : INK, lw: 2.25, shadow: o.shadow });
  const link = o.url ? { hyperlink: { url: o.url, tooltip: "영상 보기" } } : {};
  s.addImage({ path: file(key), ...f, altText: o.alt || key, objectName: nm(`img-${key}`), ...link });
  if (o.cap) caption(s, o.cap, f.x + 0.12, f.y + f.h - 0.24, { dark: o.dark, size: o.capSize });
  return f;
}
// 내레이션 캡션 박스 (BD récitatif)
function caption(s, text, x, y, o = {}) {
  const w = o.w || (0.35 + text.length * (o.size || 10.5) * 0.0105);
  rect(s, x, y, w, 0.34, o.dark ? TEAL2 : "F7E9B8", { line: o.dark ? GOLD : INK, lw: 1.25 });
  t(s, text, { x: x + 0.12, y, w: w - 0.2, h: 0.34, fontSize: o.size || 10.5, bold: true, color: o.dark ? GOLD_L : INK, valign: "middle", name: "caption" });
}
// 케이블 구분선: 선 + 연결 포트
function cable(s, x, y, w, color) {
  s.addShape(pres.shapes.LINE, { x, y, w, h: 0, line: { color, width: 1.75 }, objectName: nm("cable") });
  [0.18, 0.5, 0.82].forEach((p) => rect(s, x + w * p - 0.11, y - 0.05, 0.22, 0.1, color));
}
function label(s, text, dark) {
  const w = 0.62 + text.length * 0.1;
  rect(s, MX, 0.36, w, 0.34, dark ? TEAL2 : "F7E9B8", { line: dark ? GOLD : INK, lw: 1.25 });
  diamond(s, MX + 0.12, 0.46, 0.14, GOLD);
  t(s, text, { x: MX + 0.36, y: 0.36, w: w - 0.4, h: 0.34, fontSize: 11, bold: true, color: dark ? GOLD_L : INK, valign: "middle", charSpacing: 1, name: "label" });
}
function slide(master, lab, title, bgKey) {
  const s = pres.addSlide({ masterName: master });
  if (bgKey) full(s, bgKey, "작품 이미지를 흐리게 깐 배경");
  label(s, lab, master === "NIGHT");
  s.addText(title, { placeholder: "title" });
  return s;
}
function bullets(s, items, x, y, w, o = {}) {
  const color = o.color || INK, size = o.size || 14, gap = o.gap || 0.62;
  items.forEach((it, i) => {
    diamond(s, x, y + i * gap + size / 72 * 0.45, 0.15, o.mark || GOLD);
    t(s, it, { x: x + 0.3, y: y + i * gap, w: w - 0.3, h: gap, fontSize: size, color, lineSpacingMultiple: 1.12, name: "bullet" });
  });
}
function info(s, rows, x, y, w, o = {}) {
  rows.forEach(([a, b], i) => {
    const yy = y + i * (o.gap || 0.5);
    t(s, a, { x, y: yy, w: 1.15, h: 0.44, fontSize: 12, bold: true, color: o.dark ? GOLD_L : TEAL2, valign: "middle" });
    t(s, b, { x: x + 1.2, y: yy, w: w - 1.2, h: 0.44, fontSize: 12.5, color: o.dark ? PARCH : INK, valign: "middle", lineSpacingMultiple: 1.1 });
    s.addShape(pres.shapes.LINE, { x, y: yy + 0.47, w, h: 0, line: { color: o.dark ? "2F5357" : "CDBFA8", width: 0.75 }, objectName: nm("rule") });
  });
}
// 글만 담는 패널 (줄거리 / 리서치 칸)
function textPanel(s, x, y, w, h, head, body, o = {}) {
  rect(s, x, y, w, h, o.dark ? TEAL2 : "FBF6EC", { line: o.dark ? GOLD : INK, lw: 2.25 });
  if (o.no) t(s, o.no, { x: x + 0.2, y: y + 0.15, w: 1, h: 0.55, fontFace: SERIF, fontSize: 28, bold: true, color: GOLD });
  t(s, head, { x: x + 0.2, y: y + (o.no ? 0.75 : 0.2), w: w - 0.4, h: 0.45, fontSize: 15, bold: true, color: o.dark ? GOLD_L : INK });
  t(s, body, { x: x + 0.2, y: y + (o.no ? 1.25 : 0.7), w: w - 0.4, h: h - (o.no ? 1.4 : 0.85), fontSize: o.size || 12.5, color: o.dark ? "E8DECB" : INK2, lineSpacingMultiple: 1.25 });
}

// =====================================================================
// 01 표지
{
  const s = pres.addSlide({ masterName: "PLAIN" });
  full(s, "bg_cover", "성층권 연구시설의 빛나는 큐브 이미지보드를 흐리게 깐 배경");
  panel(s, "board", 6.85, 1.0, 5.8, 4.64, { dark: true, alt: "어두운 연구시설 한가운데 황금빛으로 빛나는 큐브와 그 앞에 선 연구원" });
  t(s, "AI CONTENT CREATOR PORTFOLIO", { x: MX, y: 1.3, w: 6, h: 0.32, fontFace: SERIF, fontSize: 12, bold: true, color: GOLD_L, charSpacing: 4 });
  t(s, "고일준", { x: MX, y: 1.7, w: 5.8, h: 1.1, fontSize: 58, bold: true, color: PARCH, valign: "middle" });
  cable(s, MX, 3.05, 5.6, GOLD);
  t(s, "RAPUNZEL", { x: MX, y: 3.3, w: 5.8, h: 0.8, fontFace: SERIF, fontSize: 40, bold: true, color: GOLD_L, charSpacing: 8, valign: "middle" });
  t(s, "라푼젤 SF 재해석 애니메이션 기획 (가제)", { x: MX, y: 4.15, w: 5.8, h: 0.42, fontSize: 16, color: PARCH });
  t(s, "고전 동화의 원형을 SF 서사와 그래픽 노블 미학으로 다시 쓰는 기획", { x: MX, y: 4.65, w: 5.9, h: 0.7, fontSize: 13, color: "B9C9C5", lineSpacingMultiple: 1.2 });
  t(s, "서울시 매력일자리 AI 콘텐츠 크리에이터 양성과정  ·  2026", { x: MX, y: 6.7, w: 6.5, h: 0.3, fontSize: 10, color: "8FA6A3" });
}

// 02 역량
{
  const s = slide("PARCH", "PROFILE", "원작 재해석 기획부터 스타일 바이블까지 설계하는 역량", "bg_parch");
  const cards = [
    ["원작 재해석 기획", "고전 동화 '라푼젤'의 원형 요소를 계승하고 SF 설정으로 새롭게 해석한 서사 기획"],
    ["세계관 / 캐릭터 설계", "AI툴을 활용한 캐릭터 시트 / 표정 시트 / 이미지보드로 세계관 시각화"],
    ["비주얼 스타일 바이블", "리뉴 클레르 손그림 스타일과 청록 / 황금 보색 대비로 일관된 톤 설계"],
    ["레퍼런스 리서치", "유사 작품 검토와 원작 비평 해석을 인물 구축에 반영"],
  ];
  cards.forEach(([h, b], i) => {
    const x = MX + i * 3.03, y = 1.95, w = 2.83, hh = 3.85;
    rect(s, x, y, w, hh, "FBF6EC", { line: INK, lw: 2.25 });
    rect(s, x, y, w, 1.0, i % 2 ? TEAL2 : "F7E9B8", { line: INK, lw: 2.25 });
    t(s, `0${i + 1}`, { x: x + 0.25, y: y + 0.15, w: 1.2, h: 0.7, fontFace: SERIF, fontSize: 30, bold: true, color: i % 2 ? GOLD_L : GOLD, valign: "middle" });
    diamond(s, x + w - 0.55, y + 0.38, 0.26, i % 2 ? GOLD_L : INK);
    t(s, h, { x: x + 0.25, y: y + 1.3, w: w - 0.4, h: 0.5, fontSize: 18, bold: true, color: INK });
    t(s, b, { x: x + 0.25, y: y + 1.95, w: w - 0.45, h: 1.75, fontSize: 13.5, color: INK2, lineSpacingMultiple: 1.28 });
  });
  rect(s, MX, 6.1, 11.93, 0.6, "F7E9B8", { line: INK, lw: 1.5 });
  t(s, "\"작은 규모의 이야기 안에 웅장하고 시네마틱한 분위기를 담아내는 것\"  작품개요 기획의도 중", { x: MX + 0.25, y: 6.1, w: 11.5, h: 0.6, fontSize: 13, bold: true, color: INK, valign: "middle" });
}

// 03 작품 개요
{
  const s = slide("NIGHT", "OVERVIEW", "완전무결한 유리 감옥과 구원자의 반전, SF 판타지 기획", "bg_night");
  panel(s, "board", MX, 1.9, 5.5, 4.4, { dark: true, alt: "빛나는 큐브 이미지보드", cap: "이미지보드  ·  큐브" });
  const x = 6.65;
  t(s, "LOGLINE", { x, y: 1.85, w: 3, h: 0.3, fontFace: SERIF, fontSize: 12, bold: true, color: GOLD_L, charSpacing: 4 });
  rect(s, x, 2.2, 5.98, 1.75, "F7E9B8", { line: GOLD, lw: 1.5 });
  t(s, "성층권 위 격리된 연구시설, 완전무결한 유리 감옥에 갇힌 소녀를 사랑하게 된 순진한 연구원은 그녀를 구하기 위해 자신의 몸마저 내어주지만, 그것이 죽은 딸과 다시 '닿기' 위해 준비해 온 한 어머니와 그녀가 만들어낸 존재의 계획이었다는 사실을 알지 못한다.", { x: x + 0.2, y: 2.3, w: 5.6, h: 1.55, fontSize: 12.5, bold: true, color: INK, lineSpacingMultiple: 1.2, valign: "middle" });
  info(s, [["장르", "SF / 판타지, 절제된 스릴러 서스펜스"], ["러닝타임", "3분 내외 프로모션 비디오, 본편은 장 / 단편으로 발전"], ["타깃", "고등학생~성인, 독창적 / 예술 애니메이션 관객층"], ["플랫폼", "추후 결정"]], x, 4.2, 5.98, { dark: true });
}

// 04 원작 계승과 반전
{
  const s = slide("PARCH", "CONCEPT", "원작의 원형은 계승하고 구원자의 역할은 뒤집은 재해석", "bg_parch");
  const rows = [
    ["탑", "성층권 연구시설의 완전 밀폐 공간 '큐브'"],
    ["머리카락", "데이터 케이블 / 에너지 파이프, 빛과 생명을 잇는 생명선"],
    ["라푼젤", "R, 죽은 딸의 데이터로 재구성된 존재"],
    ["왕자", "연구원 루카스, 구원자가 되려다 이용당하는 존재"],
    ["마녀", "K. 멜루진 박사, 딸을 잃은 신경공학자"],
    ["주문", "'라푼젤, 머리카락을 내려다오'를 변형해 닫힌 창을 여는 방법"],
  ];
  t(s, "원작 라푼젤", { x: MX, y: 1.8, w: 2.2, h: 0.35, fontSize: 12, bold: true, color: GRAY, align: "center" });
  t(s, "이 작품", { x: MX + 2.75, y: 1.8, w: 5.0, h: 0.35, fontSize: 12, bold: true, color: GRAY });
  rows.forEach(([a, b], i) => {
    const y = 2.2 + i * 0.72;
    rect(s, MX, y, 2.2, 0.56, "F7E9B8", { line: INK, lw: 1.5 });
    t(s, a, { x: MX, y, w: 2.2, h: 0.56, fontSize: 14, bold: true, color: INK, align: "center", valign: "middle" });
    t(s, "▶", { x: MX + 2.22, y, w: 0.5, h: 0.56, fontSize: 12, color: GOLD, align: "center", valign: "middle" });
    rect(s, MX + 2.75, y, 5.2, 0.56, "FBF6EC", { line: INK, lw: 1.5 });
    t(s, b, { x: MX + 2.9, y, w: 5.0, h: 0.56, fontSize: 12.5, color: INK, valign: "middle" });
  });
  const x = 9.05, y = 2.2, w = 3.58;
  rect(s, x, y, w, 4.16, TEAL2, { line: INK, lw: 2.25 });
  t(s, "THEME", { x: x + 0.25, y: y + 0.2, w: 2, h: 0.3, fontFace: SERIF, fontSize: 12, bold: true, color: GOLD_L, charSpacing: 4 });
  t(s, "닿음(실재)", { x: x + 0.25, y: y + 0.55, w: w - 0.4, h: 0.6, fontSize: 26, bold: true, color: PARCH });
  cable(s, x + 0.25, y + 1.35, w - 0.5, GOLD);
  t(s, "만질 수 없는 것을 향한 갈망, 닿기 위해 치러야 하는 대가, 그리고 닿고 나서도 완전히 확신할 수 없는 것들에 관한 이야기", { x: x + 0.25, y: y + 1.6, w: w - 0.45, h: 2.3, fontSize: 13, color: "E8DECB", lineSpacingMultiple: 1.3 });
}

// 05 줄거리: 네 칸 만화
{
  const s = slide("NIGHT", "STORY", "네 칸으로 읽는 줄거리, 순수한 로맨스 뒤에 숨은 계획", "bg_night");
  const beats = [
    ["큐브 속 소녀", "성층권 연구시설에 새로 배치된 연구원 루카스가 큐브 속 최고중증 환자 R에게 마음을 빼앗김"],
    ["디기탈리스와 창", "꽃을 만져보고 싶다는 R의 소원, 동화 속 주문을 변형한 방법으로 닫힌 창을 열어 꽃을 건넴"],
    ["뒤바뀐 존재", "회복되지 않은 척하는 R, 의식을 케이블로 전송한 루카스, 두 사람의 자아가 서로 뒤바뀜"],
    ["재회", "루카스의 몸으로 어머니 멜루진 박사에게 돌아간 R, 큐브에 홀로 남은 루카스가 지켜보는 재회"],
  ];
  beats.forEach(([h, b], i) => {
    const x = MX + (i % 2) * 6.03, y = i < 2 ? 1.85 : 4.2;
    textPanel(s, x, y, 5.88, 2.18, h, b, { dark: true, no: `0${i + 1}`, size: 13 });
  });
  t(s, "관객은 사건의 전모를 목격하되, 도덕적 판단은 스스로에게 맡기는 열린 결말", { x: MX, y: 6.5, w: 11.9, h: 0.35, fontSize: 12.5, bold: true, color: GOLD_L, valign: "middle" });
}

// 06 캐릭터 R
{
  const s = slide("PARCH", "CHARACTER", "R, 케이블 머리카락으로 재해석한 라푼젤", "bg_parch");
  panel(s, "sheet_r", MX, 1.85, 7.2, 4.8, { alt: "R 캐릭터 시트: 정면, 측면, 후면, 45도, 머리카락 구조", cap: "R 캐릭터 시트" });
  const x = 8.2;
  panel(s, "face_r", x, 1.85, 4.43, 2.95, { alt: "R 표정 시트", cap: "표정 시트" });
  bullets(s, [
    "머리카락을 데이터 케이블 / 에너지 파이프로 재해석해 빛 / 정보 / 생명의 통로로 시각화",
    "순수해 보이지만 계획적으로 목적을 이루는 이중적 인물",
  ], x, 5.05, 4.43, { size: 12.5, gap: 0.82 });
}

// 07 캐릭터 루카스 / 멜루진 박사
{
  const s = slide("PARCH", "CHARACTER", "루카스와 멜루진 박사, 구원자와 설계자", "bg_parch");
  panel(s, "sheet_lukas", MX, 1.85, 6.0, 4.76, { alt: "남자 연구원 루카스 캐릭터 시트", cap: "루카스 캐릭터 시트" });
  panel(s, "lukas_close", 7.0, 1.85, 2.2, 2.2, { alt: "루카스 얼굴과 안경 디테일 클로즈업", cap: "디테일 보정" });
  t(s, "루카스  ·  23세 전후 연구원", { x: 9.45, y: 1.85, w: 3.2, h: 0.4, fontSize: 14, bold: true, color: INK });
  t(s, "순진하고 호기심 많은 데이터 분석 담당, R을 구하려다 몸과 정체성을 빼앗기는 비극적 인물", { x: 9.45, y: 2.3, w: 3.2, h: 1.6, fontSize: 12, color: INK2, lineSpacingMultiple: 1.22 });
  bullets(s, ["둥근 금속 안경 / ID 카드 겸 데이터 저장 장치 / 데이터 장비 가방으로 직업 특성 표현"], 7.0, 4.35, 5.63, { size: 12.5, gap: 0.7 });
  rect(s, 7.0, 5.2, 5.63, 1.45, TEAL2, { line: INK, lw: 2.25 });
  t(s, "K. 멜루진 박사  ·  마녀 포지션", { x: 7.2, y: 5.3, w: 5.3, h: 0.36, fontSize: 13.5, bold: true, color: GOLD_L });
  t(s, "신경공학(뇌-기계 인터페이스) 분야의 천재 과학자, 딸을 잃은 슬픔을 차갑고 이성적인 태도 뒤에 감춘 인물", { x: 7.2, y: 5.7, w: 5.3, h: 0.9, fontSize: 12, color: "E8DECB", lineSpacingMultiple: 1.22 });
}

// 08 비주얼 스타일 바이블
{
  const s = slide("NIGHT", "STYLE BIBLE", "리뉴 클레르 손그림과 청록 / 황금 대비의 스타일 바이블", "bg_night");
  panel(s, "sheet_lukas1", MX, 1.9, 5.6, 4.48, { dark: true, alt: "유럽 만화 스타일의 루카스 캐릭터 시트 초기 버전" });
  const x = 6.8, w = 5.83;
  const rows = [["그래픽", "유럽 그래픽 노블의 리뉴 클레르 2D 손그림, 케이블 / 파이프 / 심장 모티브 장식"], ["미학", "사이버펑크 / 스팀펑크 / 레트로가 혼재된 독자적 디자인"], ["색감", "비비드하지 않은 차분한 단일 톤, 청록과 황금의 보색 대비"], ["모션", "초당 12~15프레임 리미티드, 감정적 순간에는 풀프레임 전환"]];
  rows.forEach(([a, b], i) => {
    const y = 1.9 + i * 0.86;
    rect(s, x, y, w, 0.74, TEAL2, { line: "2F5357", lw: 1 });
    t(s, a, { x: x + 0.2, y, w: 1.0, h: 0.74, fontSize: 13, bold: true, color: GOLD_L, valign: "middle" });
    t(s, b, { x: x + 1.2, y, w: w - 1.35, h: 0.74, fontSize: 12, color: PARCH, valign: "middle", lineSpacingMultiple: 1.15 });
  });
  t(s, "PALETTE", { x, y: 5.4, w: 3, h: 0.3, fontFace: SERIF, fontSize: 11, bold: true, color: GOLD_L, charSpacing: 4 });
  [TEAL_L, GOLD, ...SWATCH].forEach((c, i) => rect(s, x + i * 0.58, 5.78, 0.46, 0.46, c, { line: INK, lw: 1 }));
  t(s, "청록 / 황금 대비 색과 캐릭터 시트 팔레트", { x, y: 6.32, w: w, h: 0.3, fontSize: 10.5, color: "B9C9C5" });
}

// 09 애니메이션 테스트
{
  const s = slide("PARCH", "MOTION TEST", "키프레임 설계와 영상 테스트로 검증한 움직임", "bg_parch");
  panel(s, "key_a", MX, 1.85, 7.0, 3.94, { alt: "라푼젤 애니메이션 테스트 키프레임 5장, 16프레임 24fps" });
  const x = 8.0, w = 4.63;
  panel(s, "poster_anim", x, 1.85, w, 1.96, { url: LINK_ANIM, alt: "머리카락 모션 테스트 영상 보기", cap: "▶ 머리카락 모션 테스트" });
  panel(s, "poster_rot", x, 4.05, w, 1.96, { url: LINK_ROT, alt: "루카스 턴어라운드 테스트 영상 보기", cap: "▶ 루카스 턴어라운드 테스트" });
  bullets(s, [
    "16프레임 / 24fps 키프레임 설계로 머리카락의 흐름과 고개 움직임 테스트",
    "턴어라운드 영상으로 각도별 캐릭터 일관성 검증",
  ], MX, 6.15, 11.9, { size: 13, gap: 0.42 });
}

// 10 레퍼런스 리서치
{
  const s = slide("PARCH", "RESEARCH", "유사 작품 검토와 원작 비평으로 다진 기획의 깊이", "bg_parch");
  const cols = [
    ["유사성 검토", "엑스 마키나(2014)", "데이터 생명체 설정과 후반 전개의 장르적 유사성을 검토해 차별점 확인"],
    ["SF 원형 계보", "죽은 자식을 기술로 재현한다", "우주소년 아톰 / 초비츠 / 마조리 프라임 / 블랙 미러 'Be Right Back' / 다큐멘터리 '너를 만났다' 등"],
    ["원작 해석", "마녀와 라푼젤의 관계", "나르시시즘적 모성 / 앤 섹스턴의 시 / 페미니즘적 해석 / 융의 '삼키는 어머니' 원형, 탑의 상징성(소유욕 / 심리적 조종 / 수직성)"],
  ];
  cols.forEach(([tag, h, b], i) => {
    const x = MX + i * 4.03, y = 1.9, w = 3.83, hh = 3.65;
    rect(s, x, y, w, hh, "FBF6EC", { line: INK, lw: 2.25 });
    caption(s, tag, x + 0.2, y + 0.25, { w: 1.6, size: 11 });
    t(s, h, { x: x + 0.2, y: y + 0.8, w: w - 0.4, h: 0.8, fontSize: 16, bold: true, color: INK, lineSpacingMultiple: 1.1 });
    cable(s, x + 0.2, y + 1.7, w - 0.4, GOLD);
    t(s, b, { x: x + 0.2, y: y + 1.95, w: w - 0.4, h: 2.0, fontSize: 12.5, color: INK2, lineSpacingMultiple: 1.28 });
  });
  bullets(s, ["리서치 결과를 R / 루카스 / 멜루진 박사의 인물 구축에 반영"], MX, 5.95, 11.9, { size: 13.5, gap: 0.5 });
}

// =====================================================================
// 엔딩: 황금 큐브와 FIN
{
  const s = pres.addSlide({ masterName: "PLAIN" });
  full(s, "bg_night", "빛나는 큐브 이미지보드를 흐리고 어둡게 깐 배경");
  const size = 5.6, x = (W - size) / 2, y = 0.75;
  s.addImage({ path: file("cube"), x, y, w: size, h: size, altText: "황금빛으로 빛나는 큐브", objectName: nm("cube") });
  t(s, "감사합니다", { x, y: y + 2.0, w: size, h: 0.8, fontSize: 34, bold: true, color: PARCH, align: "center", valign: "middle" });
  t(s, "고일준", { x, y: y + 2.8, w: size, h: 0.45, fontSize: 18, bold: true, color: GOLD_L, align: "center", valign: "middle" });
  t(s, "FIN", { x: 0, y: 6.45, w: W, h: 0.5, fontFace: SERIF, fontSize: 22, bold: true, italic: true, color: GOLD_L, align: "center", charSpacing: 10 });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  if (applyTheme) await applyTheme(OUT, THEME);
  else console.log("PPTX_SKILL 없음: 테마 색 적용 단계 건너뜀");
  console.log("written", OUT);
})();

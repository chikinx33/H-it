// 김은상 포트폴리오: 「삼국무신전」 공식 티저
// node build.js  ->  ../김은상_포트폴리오.pptx
// 먹색 · 남색 바탕, 천야검 은회색 강조, 습격의 밤 적색, 궁서 제목
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const IMG = path.join(__dirname, "../work/img");
const DIMS = require("../work/dims.json");
const OUT = path.join(__dirname, "../김은상_포트폴리오.pptx");

const LINK_FILM = "https://drive.google.com/file/d/1wwMN4ngbUAvgDN58iO2qLhRP2ic8gfBX/view";
const LINK_AD_REAL = "https://drive.google.com/file/d/1G3ppCFwc6yH5SNcySnm3HA5aJxHTdTLz/view";
const LINK_AD_2D = "https://drive.google.com/file/d/1iZv3gmaNhSpLXgV4h9pFB1rm82GRha4J/view";

const BRUSH = "Gungsuh";
const SANS = "Malgun Gothic";
const INK = "0B0F15", CARD = "1A2230", LINE = "2E3A4C";
const SILVER = "D5DAE1", STEEL = "8E96A3", CRIMSON = "C0453A", GOLD = "C9A45C", WHITE = "F2F4F7";
const BEIGE = "F6EEE3", BEAN = "D9B98A", BROWN = "6B4A2E", BTEXT = "4A3A2A";

const THEME = {
  name: "Samguk Musinjeon",
  headFontFace: BRUSH,
  bodyFontFace: SANS,
  colors: { dk1: INK, lt1: WHITE, dk2: CARD, lt2: SILVER, accent1: SILVER, accent2: CRIMSON, accent3: GOLD, accent4: STEEL, accent5: BEAN, accent6: BROWN, hlink: SILVER, folHlink: SILVER },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: BRUSH, bodyFontFace: SANS };
pres.author = "김은상";
pres.title = "삼국무신전: 김은상 포트폴리오";
const W = 13.333, H = 7.5, MX = 0.7;

const footer = (c) => ({ text: { text: "김은상  ·  삼국무신전", options: { x: MX, y: 7.08, w: 6, h: 0.28, fontFace: SANS, fontSize: 9, color: c, margin: 0 } } });
const titlePh = (c) => ({ placeholder: { options: { name: "title", type: "title", x: MX, y: 0.72, w: 11.9, h: 0.75, fontFace: BRUSH, fontSize: 30, bold: true, color: c, align: "left", valign: "middle", margin: 0 }, text: "" } });
pres.defineSlideMaster({ title: "INK", background: { color: INK }, objects: [titlePh(WHITE), footer(STEEL)], slideNumber: { x: 12.13, y: 7.08, w: 0.5, h: 0.28, fontFace: SANS, fontSize: 9, color: STEEL, align: "right" } });
pres.defineSlideMaster({ title: "KONGI", background: { color: BEIGE }, objects: [titlePh(BROWN), footer("9A8A78")], slideNumber: { x: 12.13, y: 7.08, w: 0.5, h: 0.28, fontFace: SANS, fontSize: 9, color: "9A8A78", align: "right" } });
pres.defineSlideMaster({ title: "PLAIN", background: { color: INK }, objects: [] });

// ---------- helpers ----------
let sid = 0;
const nm = (s) => `${s}-${++sid}`;
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
  const link = o.url ? { hyperlink: { url: o.url, tooltip: "영상 보기" } } : {};
  if (o.border) s.addShape(pres.shapes.RECTANGLE, { x: f.x - 0.03, y: f.y - 0.03, w: f.w + 0.06, h: f.h + 0.06, fill: { color: o.border }, line: { type: "none" }, objectName: nm("border") });
  s.addImage({ path: path.join(IMG, `${key}.jpg`), ...f, altText: o.alt || key, objectName: nm(`img-${key}`), ...link });
  return f;
}
function t(s, text, o) { s.addText(text, { fontFace: SANS, color: WHITE, margin: 0, valign: "top", isTextBox: true, objectName: nm(o.name || "text"), ...o }); }
function rect(s, x, y, w, h, fill, o = {}) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: fill, transparency: o.tr || 0 }, line: o.line ? { color: o.line, width: 0.75 } : { type: "none" }, objectName: nm("rect") });
}
// 흐린 배경 이미지 (16:9로 미리 잘라 둔 사본)
function bg(s, key) { s.addImage({ path: path.join(IMG, `${key}.jpg`), x: 0, y: 0, w: W, h: H, altText: "배경", objectName: nm("bg") }); }
function slide(master, label, title, section, bgKey) {
  const s = pres.addSlide({ masterName: master, sectionTitle: section });
  if (bgKey) bg(s, bgKey);
  t(s, label, { x: MX, y: 0.4, w: 8, h: 0.3, fontSize: 11, bold: true, color: master === "KONGI" ? "B08A5A" : CRIMSON, charSpacing: 3, valign: "middle", name: "label" });
  s.addText(title, { placeholder: "title" });
  return s;
}
function linkText(s, label, url, x, y, w, h, color) {
  s.addText([{ text: label, options: { hyperlink: { url, tooltip: url }, color, bold: true } }], { x, y, w, h, fontFace: SANS, fontSize: 13, valign: "middle", margin: 0, isTextBox: true, objectName: nm("link") });
}

// =====================================================================
// 01 표지
pres.addSection({ title: "작품" });
{
  const s = pres.addSlide({ masterName: "PLAIN", sectionTitle: "작품" });
  s.addImage({ path: path.join(IMG, "cover.jpg"), x: 0, y: 0, w: W, h: H, altText: "새벽 한강 위의 삼국무신전 타이틀", objectName: nm("cover") });
  t(s, "601년", { x: MX, y: 0.6, w: 3, h: 0.4, fontFace: BRUSH, fontSize: 18, color: SILVER });
  t(s, "「삼국무신전」 공식 티저 · AI 애니메이션 포트폴리오 2026", { x: MX, y: 5.55, w: 11.9, h: 0.4, fontSize: 14, bold: true, color: SILVER, align: "center" });
  t(s, "김은상", { x: MX, y: 6.0, w: 11.9, h: 0.55, fontFace: BRUSH, fontSize: 24, bold: true, color: WHITE, align: "center" });
}

// 02 작품 소개
{
  const s = slide("INK", "LOGLINE", "작품 소개", "작품", "bg_river");
  t(s, "전설의 검을 품은 소년 김운.\n그의 선택이 삼국의 운명을 바꾼다.", { x: MX, y: 1.8, w: 6.4, h: 1.5, fontFace: BRUSH, fontSize: 26, bold: true, color: WHITE, lineSpacingMultiple: 1.3 });
  const info = [["장르", "역사 무협"], ["플랫폼", "유튜브"], ["티저 러닝타임", "약 90초 · 16:9 FHD"], ["타겟", "전연령, 30~50대 남성 메인"]];
  info.forEach(([k, v], i) => {
    const y = 3.65 + i * 0.72;
    rect(s, MX, y, 6.4, 0.6, CARD, { tr: 15 });
    t(s, k, { x: MX + 0.25, y, w: 1.8, h: 0.6, fontSize: 12, bold: true, color: SILVER, valign: "middle" });
    t(s, v, { x: MX + 2.1, y, w: 4.1, h: 0.6, fontSize: 13, color: WHITE, valign: "middle" });
  });
  const f = img(s, "f_raise", 7.45, 1.8, 5.18, 3.0, { border: STEEL, alt: "천야검을 들어 올린 김운" });
  t(s, "티저 S#10  천야검을 들어 올린 김운", { x: f.x, y: f.y + f.h + 0.12, w: 5.2, h: 0.3, fontSize: 10, color: STEEL });
}

// 03 기획의도
{
  const s = slide("INK", "CONCEPT", "기획의도: 한국형 역사 무협", "작품", "bg_slash");
  const pts = [
    ["대체역사 설정", "삼국시대 실제 역사에 무협적 상상력 결합\n한강 패권을 전쟁 대신 삼국무술대회로 결정"],
    ["인물의 선택", "복수와 화해, 나라와 백성, 과거와 미래 사이에서\n진정한 영웅의 선택을 그리는 서사"],
    ["확장 세계관", "개성 있는 인물 · 무공 · 삼국의 대립과 연대 기반\n웹소설 · 웹툰 · 애니메이션 확장 가능한 세계관 구축"],
  ];
  pts.forEach(([k, v], i) => {
    const x = MX + i * 4.06;
    rect(s, x, 1.85, 3.88, 3.7, CARD, { tr: 20, line: LINE });
    t(s, String(i + 1).padStart(2, "0"), { x: x + 0.3, y: 2.05, w: 1, h: 0.6, fontFace: BRUSH, fontSize: 28, bold: true, color: CRIMSON });
    t(s, k, { x: x + 0.3, y: 2.8, w: 3.3, h: 0.5, fontFace: BRUSH, fontSize: 20, bold: true, color: WHITE });
    t(s, v, { x: x + 0.3, y: 3.45, w: 3.3, h: 1.8, fontSize: 12.5, color: SILVER, lineSpacingMultiple: 1.3 });
  });
  t(s, "레퍼런스 장르: 무협 웹소설 · 웹툰 · 극장판 2D 애니메이션", { x: MX, y: 5.95, w: 12, h: 0.4, fontSize: 12, color: STEEL });
}

// 04 세계관
pres.addSection({ title: "세계관 · 인물" });
{
  const s = slide("INK", "WORLD  601", "세계관: 한강을 둘러싼 삼국", "세계관 · 인물", "bg_river");
  const f = img(s, "f_flag", MX, 1.8, 5.4, 3.0, { border: STEEL, alt: "고구려 군기와 한강" });
  t(s, "티저 S#2  한강을 내려다보는 군기", { x: f.x, y: f.y + f.h + 0.12, w: 5.4, h: 0.3, fontSize: 10, color: STEEL });
  t(s, "「한강을 차지하는 자가… 천하를 얻는다.」", { x: MX, y: 5.45, w: 5.6, h: 0.5, fontFace: BRUSH, fontSize: 17, bold: true, color: WHITE });
  t(s, "티저 오프닝 내레이션", { x: MX, y: 5.95, w: 5, h: 0.3, fontSize: 10, color: STEEL });
  const rx = 6.45, rw = W - MX - rx;
  const nations = [["고구려", "삼족오 군기", "최강의 권법가 연호개"], ["백제", "용 군기", "최고의 여성 검객 부여설화"], ["신라", "뿔 달린 천마 군기", "화랑 김운 · 김유신"]];
  nations.forEach(([n, flag, who], i) => {
    const y = 1.8 + i * 1.05;
    rect(s, rx, y, rw, 0.92, CARD, { tr: 15, line: LINE });
    t(s, n, { x: rx + 0.3, y, w: 1.4, h: 0.92, fontFace: BRUSH, fontSize: 20, bold: true, color: WHITE, valign: "middle" });
    t(s, flag, { x: rx + 1.75, y: y + 0.14, w: rw - 2.0, h: 0.32, fontSize: 12, bold: true, color: SILVER });
    t(s, who, { x: rx + 1.75, y: y + 0.48, w: rw - 2.0, h: 0.3, fontSize: 11, color: STEEL });
  });
  rect(s, rx, 5.0, rw, 1.5, "2A1614", { tr: 10, line: "5A2A24" });
  t(s, "삼국무신전", { x: rx + 0.3, y: 5.12, w: 3, h: 0.45, fontFace: BRUSH, fontSize: 20, bold: true, color: CRIMSON });
  t(s, "10년마다 삼국 대표 무인이 겨루는 대회\n우승국이 한강 지배권 획득", { x: rx + 0.3, y: 5.62, w: rw - 0.6, h: 0.8, fontSize: 12, color: SILVER, lineSpacingMultiple: 1.25 });
}

// 05 등장인물 (티저)
{
  const s = slide("INK", "CHARACTERS", "티저 등장인물", "세계관 · 인물", "bg_night");
  const chars = [["c_김운", "김운", "주인공", "가야 무인 가문의 아들, 신라 화랑\n의리와 정의감, 냉철한 판단력"], ["c_김진백", "김진백", "아버지", "문무를 겸비한 인재\n천야검의 비밀을 아는 마지막 인물"], ["c_설연화", "설연화", "어머니", "신라 귀족 출신\n끝까지 김운을 지키는 강인한 어머니"]];
  const cw = 2.62, gap = 0.16;
  chars.forEach(([key, n, role, d], i) => {
    const x = MX + i * (cw + gap);
    rect(s, x, 1.8, cw, 5.0, CARD, { tr: 15, line: LINE });
    img(s, key, x + 0.1, 1.9, cw - 0.2, cw - 0.2);
    t(s, n, { x: x + 0.2, y: 4.45, w: 1.4, h: 0.45, fontFace: BRUSH, fontSize: 20, bold: true, color: WHITE, valign: "middle" });
    t(s, role, { x: x + 1.4, y: 4.45, w: cw - 1.6, h: 0.45, fontSize: 11, bold: true, color: CRIMSON, align: "right", valign: "middle" });
    t(s, d, { x: x + 0.2, y: 5.05, w: cw - 0.4, h: 1.5, fontSize: 11, color: SILVER, lineSpacingMultiple: 1.25 });
  });
  const rx = MX + 3 * (cw + gap) + 0.1, rw = W - MX - rx;
  rect(s, rx, 1.8, rw, 5.0, CARD, { tr: 15, line: LINE });
  t(s, "티저 속 운명의 인물들", { x: rx + 0.3, y: 1.98, w: rw - 0.6, h: 0.35, fontSize: 13, bold: true, color: SILVER });
  const others = [["무진", "은둔 고수, 천야검의 주인"], ["김유신", "의형제, 문무 겸비의 전략가"], ["부여설화", "백제 최고의 검객"], ["연호개", "고구려 최강의 권법가"], ["소명공주", "신라 왕실, 김운을 사랑하는 여인"], ["살수", "정체를 감춘 암살 조직"]];
  others.forEach(([n, d], i) => {
    const y = 2.5 + i * 0.68;
    t(s, n, { x: rx + 0.3, y, w: 1.3, h: 0.55, fontFace: BRUSH, fontSize: 15, bold: true, color: WHITE, valign: "middle" });
    t(s, d, { x: rx + 1.65, y, w: rw - 1.9, h: 0.55, fontSize: 11, color: STEEL, valign: "middle" });
  });
}

// 06 인물 설계 15인
{
  const s = slide("INK", "CAST  15", "세계관 인물 설계: 캐릭터 시트 15인", "세계관 · 인물", "bg_slash");
  const groups = [
    ["가야 · 가족", ["김운", "김진백", "설연화", "무진"], SILVER],
    ["신라 · 화랑", ["김유신", "비형랑", "석도윤", "모가대", "소명공주", "천관녀", "진평왕"], GOLD],
    ["삼국 라이벌", ["부여설화 (백제)", "연호개 (고구려)"], STEEL],
    ["흑막 · 암살", ["사마령", "살수"], CRIMSON],
  ];
  const cw = 2.92, gap = 0.11;
  groups.forEach(([g, names, col], i) => {
    const x = MX + i * (cw + gap);
    rect(s, x, 1.85, cw, 4.6, CARD, { tr: 15, line: LINE });
    t(s, g, { x: x + 0.3, y: 2.05, w: cw - 0.6, h: 0.45, fontFace: BRUSH, fontSize: 18, bold: true, color: col });
    s.addText(names.map((n, j) => ({ text: n, options: { breakLine: j < names.length - 1 } })), { x: x + 0.3, y: 2.7, w: cw - 0.6, h: 3.5, fontFace: SANS, fontSize: 14, color: WHITE, valign: "top", margin: 0, paraSpaceAfter: 8, isTextBox: true, objectName: nm("names") });
  });
  t(s, "웹소설 · 웹툰 확장을 전제로 한 인물 관계와 소속 설계", { x: MX, y: 6.6, w: 12, h: 0.35, fontSize: 13, bold: true, color: SILVER });
}

// 07 티저 구성 90초
pres.addSection({ title: "연출" });
{
  const s = slide("INK", "STRUCTURE  00:00 - 01:30", "티저 구성: 11씬 90초", "연출", "bg_river");
  const scenes = [["오프닝", 4, "601년"], ["S#1", 7, "한강"], ["S#2", 6, "삼국 대치"], ["S#3", 9, "삼국무신전"], ["S#4", 12, "어린 시절"], ["S#5", 10, "운명의 밤"], ["S#6", 4, "무진"], ["S#7", 7, "천야검"], ["S#8", 4, "전쟁"], ["S#9", 14, "운명의 인물들"], ["S#10", 10, "결의"], ["S#11", 4, "타이틀"]];
  const total = scenes.reduce((a, b) => a + b[1], 0), u = 12.0 / total;
  const warm = { "S#4": "6A4A2A" }, red = { "S#5": "6E2420", "S#8": "4A2A26" };
  let x = MX;
  scenes.forEach(([no, d, name]) => {
    const w = d * u, col = warm[no] || red[no] || "243042";
    rect(s, x + 0.015, 2.0, w - 0.03, 1.6, col, { line: LINE });
    t(s, no, { x: x + 0.05, y: 2.1, w: w - 0.1, h: 0.3, fontSize: 10, bold: true, color: SILVER, align: "center" });
    t(s, name, { x: x + 0.05, y: 2.45, w: w - 0.1, h: 0.7, fontSize: 9.5, color: WHITE, align: "center" });
    t(s, `${d}초`, { x: x + 0.05, y: 3.2, w: w - 0.1, h: 0.3, fontSize: 9.5, color: STEEL, align: "center" });
    x += w;
  });
  t(s, "시나리오 기준 컷 길이 합산 (교차편집 구간은 일부 겹침)", { x: MX, y: 3.75, w: 9, h: 0.3, fontSize: 10, color: STEEL });
  const notes = [["웅장함", "한강 · 삼국 대치 · 무신전\n버드아이 뷰와 로우 앵글의 스케일", "243042"], ["따뜻함과 상실", "가을빛 김해의 가족에서\n검붉은 밤의 습격으로 하드컷", "6E2420"], ["결의", "질문과 답의 구조 위에\n천야검 검광으로 타이틀 연결", "8E96A3"]];
  notes.forEach(([k, v, col], i) => {
    const nx = MX + i * 4.06;
    rect(s, nx, 4.35, 3.88, 2.2, CARD, { tr: 15, line: LINE });
    rect(s, nx + 0.3, 4.6, 0.3, 0.3, col);
    t(s, k, { x: nx + 0.75, y: 4.5, w: 2.9, h: 0.5, fontFace: BRUSH, fontSize: 18, bold: true, color: WHITE, valign: "middle" });
    t(s, v, { x: nx + 0.3, y: 5.2, w: 3.3, h: 1.1, fontSize: 12, color: SILVER, lineSpacingMultiple: 1.25 });
  });
}

// 08 공통 비주얼 사양
{
  const s = slide("INK", "VISUAL SPEC", "공통 비주얼 사양", "연출", "bg_slash");
  const spec = [["스타일", "한국 2D 애니메이션 · 웹툰 영향의 시네마틱 2D 애니메이션"], ["작화", "깨끗한 선화, 셀 셰이딩, 영화적인 빛과 대기 원근감"], ["톤", "웅장함 / 신비 / 비장함 / 역사 판타지"], ["카메라", "극장판 2D 애니메이션의 안정적이고 역동적인 카메라 워크"], ["영상화", "안개 · 강물 · 불꽃 · 깃발 · 머리카락 · 옷자락의 움직임 공간 확보"], ["캐릭터", "프로젝트 등록 확정 이미지를 최우선 레퍼런스로 고정"]];
  spec.forEach(([k, v], i) => {
    const y = 1.8 + i * 0.75;
    rect(s, MX, y, 8.0, 0.63, CARD, { tr: 15, line: LINE });
    t(s, k, { x: MX + 0.25, y, w: 1.3, h: 0.63, fontSize: 12.5, bold: true, color: SILVER, valign: "middle" });
    t(s, v, { x: MX + 1.6, y, w: 6.25, h: 0.63, fontSize: 12, color: WHITE, valign: "middle" });
  });
  const rx = 9.0, rw = W - MX - rx;
  rect(s, rx, 1.8, rw, 2.25, CARD, { tr: 15, line: LINE });
  t(s, "천야검 검신 색 (확정)", { x: rx + 0.3, y: 1.95, w: rw - 0.6, h: 0.35, fontSize: 12, bold: true, color: SILVER });
  s.addShape(pres.shapes.RECTANGLE, { x: rx + 0.3, y: 2.45, w: rw - 0.6, h: 0.75, fill: { color: "B9BEC6" }, line: { color: WHITE, width: 0.75 }, objectName: nm("swatch") });
  t(s, "은회색, 모든 컷 통일", { x: rx + 0.3, y: 3.35, w: rw - 0.6, h: 0.35, fontSize: 12, color: WHITE });
  rect(s, rx, 4.25, rw, 2.25, CARD, { tr: 15, line: LINE });
  t(s, "자막 서체", { x: rx + 0.3, y: 4.4, w: rw - 0.6, h: 0.35, fontSize: 12, bold: true, color: SILVER });
  t(s, "오프닝 「601년」과 엔딩 「삼국무신전」을 같은 서체와 은회색 톤으로 통일", { x: rx + 0.3, y: 4.85, w: rw - 0.6, h: 1.4, fontSize: 12, color: WHITE, lineSpacingMultiple: 1.25 });
}

// 09 내레이션 설계
{
  const s = slide("INK", "NARRATION", "내레이션 설계: 4인 보이스와 질문 · 답 구조", "연출", "bg_night");
  const voices = [["내레이터", "40대 남성, 중저음의 중후한 톤", "오프닝 내레이션"], ["김운 V.O.", "10대 후반, 슬프지만 차분한 톤", "회상과 독백"], ["무진", "60대, 차분한 톤", "질문"], ["김진백", "40대, 적과 대치한 급한 톤", "「빨리 가시오.」"]];
  voices.forEach(([n, tone, use], i) => {
    const y = 1.8 + i * 0.95;
    rect(s, MX, y, 5.4, 0.82, CARD, { tr: 15, line: LINE });
    t(s, n, { x: MX + 0.25, y: y + 0.08, w: 2.2, h: 0.35, fontFace: BRUSH, fontSize: 15, bold: true, color: WHITE });
    t(s, use, { x: MX + 2.5, y: y + 0.1, w: 2.7, h: 0.3, fontSize: 10.5, color: CRIMSON, align: "right" });
    t(s, tone, { x: MX + 0.25, y: y + 0.45, w: 5.0, h: 0.3, fontSize: 11, color: SILVER });
  });
  const rx = 6.45, rw = W - MX - rx;
  const lines = [["S#4", "그 시절의 나는 아무것도 몰랐다.", SILVER], ["S#5", "그날 밤 모든 것이 끝났다.", CRIMSON], ["S#6", "무진  「너는 무엇을 위해 검을 드느냐?」", GOLD], ["S#7", "그때 나는 답하지 못했다.", STEEL], ["S#10", "이제는 답할 수 있다.", SILVER], ["S#10", "…이 검은 사람을 지키기 위해 있다.", WHITE]];
  lines.forEach(([sc, l, col], i) => {
    const y = 1.8 + i * 0.66;
    t(s, sc, { x: rx, y, w: 0.8, h: 0.55, fontSize: 11, bold: true, color: STEEL, valign: "middle" });
    t(s, l.startsWith("무진") ? l : `「${l}」`, { x: rx + 0.85, y, w: rw - 0.85, h: 0.55, fontFace: BRUSH, fontSize: i === 5 ? 17 : 14, bold: i === 5, color: col, valign: "middle" });
  });
  t(s, "질문에서 답으로 이어지는 독백 구조를 통한 주인공 성장 압축", { x: MX, y: 6.55, w: 12, h: 0.35, fontSize: 13, bold: true, color: SILVER });
}

// 10 연출 포인트
{
  const s = slide("INK", "DIRECTING", "연출 포인트: 하드컷과 교차편집", "연출", "bg_family");
  const iw = 5.85;
  const a = img(s, "f_family", MX, 1.8, iw, 3.29, { border: GOLD, alt: "가족의 마지막 평화로운 식사" });
  const b = img(s, "s_assassin", W - MX - iw, 1.8, iw, 3.29, { border: CRIMSON, alt: "검은 후드의 살수" });
  t(s, "S#4  가족의 마지막 평화로운 식사", { x: a.x, y: a.y + a.h + 0.12, w: iw, h: 0.3, fontSize: 11, bold: true, color: GOLD });
  t(s, "S#9  어둠 속 검은 후드의 살수", { x: b.x, y: b.y + b.h + 0.12, w: iw, h: 0.3, fontSize: 11, bold: true, color: CRIMSON });
  const pts = [["하드컷", "평화로운 가을 낮에서 검붉은 밤으로 급전환"], ["교차편집", "삼국 대치 · 인물 소개 · 결의 구간의 장면 교차"], ["음악 설계", "00:07 배경음악 시작, 어린 시절에서 따뜻한 톤 전환, 타이틀에서 페이드아웃"]];
  pts.forEach(([k, v], i) => {
    const y = 5.65 + i * 0.45;
    t(s, k, { x: MX, y, w: 1.6, h: 0.4, fontSize: 12.5, bold: true, color: SILVER, valign: "middle" });
    t(s, v, { x: MX + 1.7, y, w: 10.3, h: 0.4, fontSize: 12, color: WHITE, valign: "middle" });
  });
}

// 11 주요 장면
{
  const s = slide("INK", "KEY SCENES", "주요 장면", "연출", "bg_river");
  const shots = [["s_river", "S#1", "새벽 한강, 버드아이 뷰"], ["f_flag", "S#2", "삼국 대치의 군기"], ["f_arena", "S#3", "거대한 야외 삼국무신전"], ["f_family", "S#4", "가족의 마지막 식사"], ["s_assassin", "S#9", "검은 후드의 살수"], ["f_raise", "S#10", "천야검을 들어 올린 김운"], ["f_slash", "S#10", "천야검, 마지막 일격"], ["s_title", "S#11", "은회색 타이틀"]];
  const gap = 0.14, iw = (12.0 - 3 * gap) / 4, ih = iw / 1.777;
  shots.forEach(([key, sc, cap], i) => {
    const col = i % 4, row = Math.floor(i / 4);
    const x = MX + col * (iw + gap), y = 1.75 + row * (ih + 0.72);
    const f = img(s, key, x, y, iw, ih, { top: true, border: LINE });
    t(s, [{ text: sc + "   ", options: { bold: true, color: CRIMSON } }, { text: cap, options: { color: SILVER } }], { x, y: f.y + f.h + 0.1, w: iw, h: 0.32, fontSize: 11 });
  });
  t(s, "S#2 · S#3 · S#4 · S#10 장면은 제작 과정 캡처에서 발췌", { x: MX, y: 6.65, w: 12, h: 0.3, fontSize: 10, color: STEEL });
}

// 12~13 AI 제작 워크플로우
pres.addSection({ title: "제작 과정" });
function workflow(label, title, steps) {
  const s = slide("INK", label, title, "제작 과정", "bg_slash");
  const gap = 0.2, iw = (12.0 - 2 * gap) / 3, ih = iw * 1021 / 1912;
  steps.forEach(([key, no, k, v], i) => {
    const x = MX + i * (iw + gap);
    const f = img(s, key, x, 1.8, iw, ih, { top: true, border: LINE, alt: `${k} 작업 화면` });
    t(s, no, { x, y: f.y + f.h + 0.2, w: 0.7, h: 0.5, fontFace: BRUSH, fontSize: 24, bold: true, color: CRIMSON, valign: "middle" });
    t(s, k, { x: x + 0.75, y: f.y + f.h + 0.2, w: iw - 0.75, h: 0.5, fontFace: BRUSH, fontSize: 17, bold: true, color: WHITE, valign: "middle" });
    t(s, v, { x, y: f.y + f.h + 0.8, w: iw, h: 1.2, fontSize: 11.5, color: SILVER, lineSpacingMultiple: 1.25 });
  });
  return s;
}
workflow("WORKFLOW 1", "AI 제작 워크플로우 ①", [
  ["w_시나리오작업컷", "01", "시나리오 설계", "AI 대화를 활용한 시나리오 초안과 공통 비주얼 사양 정리, 컷별 초 단위 확정"],
  ["w_캐릭터시트컷2", "02", "캐릭터 시트", "정면 · 측면 · 후면 · 얼굴의 4분할 턴어라운드, 15인 캐릭터 고정"],
  ["w_티저영상컷1", "03", "장면 영상 생성", "확정 캐릭터 이미지를 레퍼런스로 한 컷별 영상 생성"],
]);
workflow("WORKFLOW 2", "AI 제작 워크플로우 ②", [
  ["w_나레이션컷2", "04", "내레이션", "4인 보이스 구성, 장면별 대사 톤 지정"],
  ["w_티저배경음악컷", "05", "배경음악", "무협 OST 분위기의 배경음악, 구간별 길이 조정"],
  ["w_프리미어편집컷", "06", "편집", "프리미어 프로 편집, 교차편집과 조정 레이어 적용"],
]);

// 14 영상 결과물
pres.addSection({ title: "결과물 · IP" });
{
  const s = slide("INK", "TEASER", "영상 결과물: 삼국무신전 공식 티저 (1차)", "결과물 · IP", "bg_river");
  const f = img(s, "poster_film", MX, 1.75, 7.9, 4.44, { border: STEEL, url: LINK_FILM, alt: "티저 대표 이미지 (클릭 시 영상 재생)" });
  t(s, "이미지 클릭 시 영상 재생 (Google Drive)", { x: MX, y: f.y + f.h + 0.15, w: 7.9, h: 0.3, fontSize: 10, color: STEEL });
  const rx = 8.9, rw = W - MX - rx;
  rect(s, rx, 1.75, rw, 4.44, CARD, { tr: 15, line: LINE });
  const info = [["버전", "1차 편집본"], ["러닝타임", "약 90초 (시나리오 기준)"], ["화면비", "16:9 FHD"], ["구성", "11씬, 4인 내레이션"]];
  info.forEach(([k, v], i) => {
    const y = 2.0 + i * 0.98;
    t(s, k, { x: rx + 0.3, y, w: rw - 0.6, h: 0.3, fontSize: 11, bold: true, color: CRIMSON });
    t(s, v, { x: rx + 0.3, y: y + 0.33, w: rw - 0.6, h: 0.45, fontSize: 14, color: WHITE });
  });
  linkText(s, "▶  티저 영상 보기", LINK_FILM, rx, 6.35, rw, 0.4, SILVER);
}

// 15 AI 숏폼 광고
{
  const s = slide("INK", "SHORT-FORM AD", "AI 숏폼 광고: 향이나", "결과물 · IP", "bg_slash");
  const iw = 5.85;
  const a = img(s, "poster_ad_실사", MX, 1.8, iw, 3.29, { border: LINE, url: LINK_AD_REAL, alt: "향이나 실사 버전 (클릭 시 재생)" });
  const b = img(s, "poster_ad_2D", W - MX - iw, 1.8, iw, 3.29, { border: LINE, url: LINK_AD_2D, alt: "향이나 2D 버전 (클릭 시 재생)" });
  t(s, "실사 버전", { x: a.x, y: a.y + a.h + 0.15, w: 3, h: 0.4, fontFace: BRUSH, fontSize: 17, bold: true, color: WHITE });
  t(s, "2D 버전", { x: b.x, y: b.y + b.h + 0.15, w: 3, h: 0.4, fontFace: BRUSH, fontSize: 17, bold: true, color: WHITE });
  linkText(s, "▶  영상 보기", LINK_AD_REAL, a.x + a.w - 2.0, a.y + a.h + 0.15, 2.0, 0.4, SILVER);
  linkText(s, "▶  영상 보기", LINK_AD_2D, b.x + b.w - 2.0, b.y + b.h + 0.15, 2.0, 0.4, SILVER);
  t(s, "하나의 광고를 실사와 2D 두 가지 스타일로 제작한 비교 결과물", { x: MX, y: 5.85, w: 12, h: 0.4, fontSize: 13, bold: true, color: SILVER });
  t(s, "광고 내용 · 길이: [영상 확인 후 기입]", { x: MX, y: 6.3, w: 12, h: 0.35, fontSize: 11, color: STEEL });
}

// 16 캐릭터 IP: 콩이 (캐릭터 고유의 베이지 톤)
{
  const s = slide("KONGI", "CHARACTER IP", "캐릭터 IP: 콩이", "결과물 · IP", "bg_kongi");
  const a = img(s, "k_캐릭터시트", MX, 1.75, 6.0, 4.0, { top: true, border: "E6D8C3", alt: "콩이 캐릭터 시트" });
  const rx = MX + 6.3, rw = W - MX - rx;
  img(s, "k_브랜드로고", rx, 1.75, 2.6, 2.6, { top: true, left: true, alt: "콩이 브랜드 로고" });
  t(s, "작은 콩, 큰 행복", { x: rx + 2.85, y: 1.95, w: rw - 2.85, h: 0.6, fontFace: BRUSH, fontSize: 22, bold: true, color: BROWN });
  t(s, "SMALL BEAN, BIG HAPPINESS", { x: rx + 2.85, y: 2.6, w: rw - 2.85, h: 0.35, fontSize: 9, bold: true, color: "B08A5A" });
  s.addShape(pres.shapes.RECTANGLE, { x: rx + 2.85, y: 3.2, w: 0.55, h: 0.55, fill: { color: BEAN }, line: { color: "B08A5A", width: 0.75 }, objectName: nm("swatch") });
  t(s, "공식 컬러\n볶은 콩 베이지 #D9B98A", { x: rx + 3.55, y: 3.15, w: rw - 3.55, h: 0.7, fontSize: 11, color: BTEXT });
  img(s, "k_서브로고", rx, 4.55, rw, 2.2, { top: true, left: true, border: "E6D8C3", alt: "콩이 서브 로고 시스템" });
  t(s, "브랜드 로고 · 서브 로고 4종 · 캐릭터 심볼 체계", { x: MX, y: a.y + a.h + 0.15, w: 6, h: 0.35, fontSize: 12, bold: true, color: BTEXT });
}

// 17 콩이 확장
{
  const s = slide("KONGI", "IP EXPANSION", "콩이 IP 확장: 이모티콘 · 일상툰 · 굿즈", "결과물 · IP", "bg_kongi");
  const sq = 2.35, g = 0.15;
  const tiles = [["k_24가지_표정", "이모티콘 24종"], ["k_이불", "일상툰 「이불」"], ["k_알람", "일상툰 「알람」"], ["k_과자", "일상툰 「과자」"]];
  tiles.forEach(([k, cap], i) => {
    const x = MX + i * (sq + g);
    const f = img(s, k, x, 1.7, sq, sq, { top: true, left: true, border: "E6D8C3" });
    t(s, cap, { x, y: f.y + f.h + 0.06, w: sq, h: 0.3, fontSize: 11, bold: true, color: BTEXT });
  });
  const tx = MX + 4 * (sq + g) + 0.1, tw = W - MX - tx;
  t(s, "이모티콘 24종\n일상툰 4컷 3편\n굿즈 5종", { x: tx, y: 1.75, w: tw, h: 1.3, fontFace: BRUSH, fontSize: 16, bold: true, color: BROWN, lineSpacingMultiple: 1.3 });
  t(s, "상품별 컬러 옵션과 패키지 예시까지 정리한 굿즈 기획", { x: tx, y: 3.15, w: tw, h: 0.9, fontSize: 11, color: BTEXT, lineSpacingMultiple: 1.25 });
  const goods = [["k_인형", "인형"], ["k_펜", "볼펜"], ["k_스마트폰그립", "스마트폰 그립"], ["k_알람시계", "알람시계"], ["k_노트", "노트"]];
  const gw = (12.0 - 4 * 0.14) / 5, gh = gw / 1.5;
  goods.forEach(([k, n], i) => {
    const x = MX + i * (gw + 0.14);
    const f = img(s, k, x, 4.65, gw, gh, { top: true, border: "E6D8C3" });
    t(s, n, { x, y: 4.65 + gh + 0.08, w: gw, h: 0.3, fontSize: 11, bold: true, color: BTEXT });
  });
}

// 18 엔딩
{
  const s = pres.addSlide({ masterName: "PLAIN", sectionTitle: "결과물 · IP" });
  bg(s, "bg_slash");
  t(s, "「…이 검은 사람을 지키기 위해 있다.」", { x: MX, y: 1.3, w: 11.9, h: 1.1, fontFace: BRUSH, fontSize: 34, bold: true, color: WHITE, align: "center", valign: "middle" });
  t(s, "「삼국무신전」 티저 S#10 김운의 결의", { x: MX, y: 2.45, w: 11.9, h: 0.35, fontSize: 11, color: STEEL, align: "center" });
  const skills = [["세계관 기획", "대체역사 무협 세계관과 15인 인물 설계"], ["AI 제작 파이프라인", "시나리오 · 캐릭터 · 영상 · 내레이션 · 음악 · 편집 일괄 제작"], ["캐릭터 IP", "콩이 브랜드 · 이모티콘 · 일상툰 · 굿즈 확장"]];
  skills.forEach(([k, v], i) => {
    const x = MX + i * 4.06;
    rect(s, x, 3.3, 3.88, 1.45, CARD, { tr: 20, line: LINE });
    t(s, k, { x: x + 0.3, y: 3.45, w: 3.3, h: 0.45, fontFace: BRUSH, fontSize: 17, bold: true, color: SILVER });
    t(s, v, { x: x + 0.3, y: 3.95, w: 3.3, h: 0.7, fontSize: 11.5, color: WHITE, lineSpacingMultiple: 1.2 });
  });
  rect(s, MX, 5.4, 11.9, 1.25, CARD, { tr: 20, line: LINE });
  t(s, "김은상", { x: MX + 0.35, y: 5.55, w: 2, h: 0.5, fontFace: BRUSH, fontSize: 22, bold: true, color: WHITE, valign: "middle" });
  [["이메일", 2.7], ["연락처", 6.0], ["포트폴리오 링크", 9.0]].forEach(([k, x]) => {
    t(s, k, { x: MX + x, y: 5.55, w: 2.6, h: 0.28, fontSize: 10, color: STEEL });
    s.addShape(pres.shapes.LINE, { x: MX + x, y: 6.3, w: 2.6, h: 0, line: { color: STEEL, width: 1 }, objectName: nm("blank-line") });
  });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();

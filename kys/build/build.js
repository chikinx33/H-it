// 김윤성 포트폴리오: 「지금이라는 문」
// node build.js  ->  ../김윤성_포트폴리오.pptx
// 작품의 색 설계(차갑고 탁한 청회색 -> 노을빛)를 덱 진행에도 그대로 적용
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const IMG = path.join(__dirname, "../work/img");
const DIMS = require("../work/dims.json");
const OUT = path.join(__dirname, "../김윤성_포트폴리오.pptx");

const LINK_FILM = "https://drive.google.com/file/d/1MyGeTKwiceS10SmJRmEL3EWLXy3p9L3t/view";
const LINK_SP = "https://drive.google.com/file/d/1IdLq_u1jngfyKZQjePuPLi2wdwnOzN5D/view";

const SANS = "Malgun Gothic";
const TC = "Consolas"; // 편집 타임코드 모티프

// 단계별 톤
const T = {
  COLD: { bg: "1E252D", card: "2A333D", text: "E3E8EC", muted: "8C99A6", accent: "7FA3B8", line: "3A4652" },
  HARD: { bg: "14181D", card: "232A31", text: "E3E8EC", muted: "8C99A6", accent: "B8C4CE", line: "343E48" },
  TRANS: { bg: "2A2826", card: "3A3633", text: "F2ECE6", muted: "B0A69C", accent: "D9B48A", line: "4A443F" },
  WARM: { bg: "2A1C16", card: "3B2920", text: "FFF1E4", muted: "D1B29C", accent: "F28C4B", line: "5A4033" },
  IP: { bg: "FCF6EE", card: "FFFFFF", text: "4A3F38", muted: "8A7D73", accent: "E97A8E", line: "EADFD2" },
};
const GOLD = "F6C45E", GREEN = "8DBF6A", PINK = "F2899B", SKYB = "7DB4E3";

const THEME = {
  name: "Now Door",
  headFontFace: SANS,
  bodyFontFace: SANS,
  colors: {
    dk1: "1E252D", lt1: "FFF1E4", dk2: "2A1C16", lt2: "E3E8EC",
    accent1: "7FA3B8", accent2: "F28C4B", accent3: "F6C45E", accent4: "8DBF6A", accent5: "F2899B", accent6: "7DB4E3",
    hlink: "F28C4B", folHlink: "F28C4B",
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: SANS, bodyFontFace: SANS };
pres.author = "김윤성";
pres.title = "지금이라는 문: 김윤성 포트폴리오";
const W = 13.333, H = 7.5, MX = 0.7;

// 단계별 레이아웃 (배경 · 제목색 · 푸터)
for (const [name, c] of Object.entries(T)) {
  pres.defineSlideMaster({
    title: name,
    background: { color: c.bg },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: MX, y: 0.78, w: 11.9, h: 0.68, fontFace: SANS, fontSize: 26, bold: true, color: c.text, align: "left", valign: "middle", margin: 0 }, text: "" } },
      { text: { text: "김윤성  ·  지금이라는 문", options: { x: MX, y: 7.08, w: 6, h: 0.28, fontFace: SANS, fontSize: 9, color: c.muted, margin: 0 } } },
    ],
    slideNumber: { x: 12.13, y: 7.08, w: 0.5, h: 0.28, fontFace: SANS, fontSize: 9, color: c.muted, align: "right" },
  });
}

pres.defineSlideMaster({ title: "PLAIN", background: { color: "111418" }, objects: [] });

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
  const ext = key === "title" ? "png" : "jpg";
  const f = o.top ? (() => { const g = fit(key, x, y, w, h); return { ...g, y }; })() : fit(key, x, y, w, h);
  const link = o.url ? { hyperlink: { url: o.url, tooltip: "영상 보기" } } : {};
  if (o.border) s.addShape(pres.shapes.RECTANGLE, { x: f.x - 0.03, y: f.y - 0.03, w: f.w + 0.06, h: f.h + 0.06, fill: { color: o.border }, line: { type: "none" }, objectName: nm("border") });
  s.addImage({ path: path.join(IMG, `${key}.${ext}`), ...f, altText: o.alt || key, objectName: nm(`img-${key}`), ...link });
  return f;
}
function t(s, text, o) {
  s.addText(text, { fontFace: SANS, margin: 0, valign: "top", isTextBox: true, objectName: nm(o.name || "text"), ...o });
}
function rect(s, x, y, w, h, fill, o = {}) {
  s.addShape(o.round ? pres.shapes.ROUNDED_RECTANGLE : pres.shapes.RECTANGLE, { x, y, w, h, rectRadius: o.round || 0, fill: { color: fill }, line: o.line ? { color: o.line, width: 0.75 } : { type: "none" }, objectName: nm("rect") });
}
function slide(stage, tc, title, section) {
  const c = T[stage];
  const s = pres.addSlide({ masterName: stage, sectionTitle: section });
  t(s, tc, { x: MX, y: 0.4, w: 8, h: 0.3, fontFace: TC, fontSize: 11, bold: true, color: c.accent, valign: "middle", name: "timecode" });
  s.addText(title, { placeholder: "title" });
  return { s, c };
}
function linkText(s, label, url, x, y, w, h, color) {
  s.addText([{ text: label, options: { hyperlink: { url, tooltip: url }, color, bold: true } }], { x, y, w, h, fontFace: SANS, fontSize: 13, valign: "middle", margin: 0, isTextBox: true, objectName: nm("link") });
}

// =====================================================================
// 01 표지: 차가운 출근길(S#1)과 노을의 퇴근길(S#6)을 좌우로 나란히
pres.addSection({ title: "작품" });
{
  const s = pres.addSlide({ masterName: "PLAIN", sectionTitle: "작품" });
  s.background = { color: "111418" };
  s.addImage({ path: path.join(IMG, "cover_left.jpg"), x: 0, y: 0, w: W / 2, h: H, altText: "출근길 지하철의 수현", objectName: nm("cover-cold") });
  s.addImage({ path: path.join(IMG, "cover_right.jpg"), x: W / 2, y: 0, w: W / 2, h: H, altText: "노을 속 수현의 미소", objectName: nm("cover-warm") });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 2.55, w: W, h: 2.55, fill: { color: "000000", transparency: 40 }, line: { type: "none" }, objectName: nm("cover-band") });
  img(s, "title", 3.9, 2.75, 5.5, 1.45, { alt: "지금이라는 문 제목 캘리그라피" });
  t(s, "AI 애니메이션 포트폴리오 2026", { x: MX, y: 4.35, w: W - 2 * MX, h: 0.35, fontSize: 13, bold: true, color: "E3E8EC", align: "center" });
  t(s, "김윤성", { x: MX, y: 4.7, w: W - 2 * MX, h: 0.45, fontSize: 20, bold: true, color: "FFFFFF", align: "center" });
  t(s, "00:00:00", { x: 0.35, y: 0.3, w: 2, h: 0.3, fontFace: TC, fontSize: 11, bold: true, color: "B8C4CE" });
  t(s, "00:02:07", { x: W - 2.35, y: 0.3, w: 2, h: 0.3, fontFace: TC, fontSize: 11, bold: true, color: "FFD9B8", align: "right" });
}

// 02 작품 소개
{
  const { s, c } = slide("COLD", "INTRO", "작품 소개", "작품");
  t(s, "불안한 생각에 갇힌 중년 직장인 수현이\n상담사 도윤을 만나 생각과 자신을 분리하고,\n퇴근길에서 '지금 이 순간'의 감각을 되찾는 이야기", { x: MX, y: 1.75, w: 6.6, h: 1.7, fontSize: 19, bold: true, color: c.text, lineSpacingMultiple: 1.3 });
  const stats = [["127초", "러닝타임"], ["6씬", "구성"], ["24컷", "시나리오 기준"]];
  stats.forEach(([v, k], i) => {
    const x = MX + i * 2.25;
    rect(s, x, 3.85, 2.05, 1.35, c.card);
    t(s, v, { x: x + 0.25, y: 3.98, w: 1.7, h: 0.65, fontSize: 28, bold: true, color: c.accent, valign: "middle" });
    t(s, k, { x: x + 0.25, y: 4.68, w: 1.7, h: 0.35, fontSize: 12, color: c.muted });
  });
  rect(s, MX, 5.5, 6.55, 1.2, c.card);
  t(s, "엔딩 대사", { x: MX + 0.3, y: 5.65, w: 2, h: 0.3, fontSize: 11, bold: true, color: GOLD });
  t(s, "수현 (V.O.)  「지금... 이 순간.」", { x: MX + 0.3, y: 6.0, w: 6, h: 0.5, fontSize: 17, bold: true, color: c.text, valign: "middle" });
  img(s, "s3-1", 7.65, 1.75, 4.98, 2.8, { border: c.line, alt: "지금이라는 문 심리치료실 외관" });
  img(s, "s6-1", 7.65, 4.75, 4.98, 1.95, { border: c.line, alt: "노을 진 퇴근길" });
}

// 03 등장인물
pres.addSection({ title: "인물 · 공간" });
{
  const { s, c } = slide("COLD", "CHARACTERS", "등장인물", "인물 · 공간");
  const chars = [
    ["char_suhyun", "김수현", "주인공 · 부장", "실적 압박과 퇴직 후 걱정에 갇힌 중년 직장인\n생각과 자신의 분리를 통한 감각의 회복"],
    ["char_doyun", "이도윤", "상담사", "'지금이라는 문' 심리치료실의 상담사\n담담한 질문과 숙제로 수현을 이끄는 인물"],
    ["char_taejun", "박태준", "상사", "프로젝트 책임을 매섭게 묻는 상사\n수현의 패닉을 촉발하는 인물"],
  ];
  const cw = 3.85, gap = 0.18;
  chars.forEach(([key, name, role, desc], i) => {
    const x = MX + i * (cw + gap);
    rect(s, x, 1.7, cw, 5.1, c.card);
    rect(s, x + 0.15, 1.85, cw - 0.3, 2.65, "F4F4F2");
    img(s, key, x + 0.2, 1.9, cw - 0.4, 2.55, { alt: `${name} 캐릭터 시트` });
    t(s, name, { x: x + 0.3, y: 4.72, w: 1.6, h: 0.45, fontSize: 20, bold: true, color: c.text, valign: "middle" });
    t(s, role, { x: x + 1.85, y: 4.72, w: cw - 2.15, h: 0.45, fontSize: 12, bold: true, color: c.accent, align: "right", valign: "middle" });
    t(s, desc, { x: x + 0.3, y: 5.35, w: cw - 0.6, h: 1.2, fontSize: 12, color: c.muted, lineSpacingMultiple: 1.2 });
  });
}

// 04 공간: 심리치료실 '지금이라는 문'
{
  const { s, c } = slide("COLD", "LOCATION", "공간: 심리치료실 '지금이라는 문'", "인물 · 공간");
  const iw = 5.9;
  const a = img(s, "bg_day", MX, 1.75, iw, 3.36, { border: c.line, alt: "배경 설정: 낮" });
  const b = img(s, "bg_night", W - MX - iw, 1.75, iw, 3.36, { border: c.line, alt: "배경 설정: 밤" });
  t(s, "배경 설정 · 낮", { x: a.x, y: a.y + a.h + 0.15, w: 3, h: 0.3, fontSize: 12, bold: true, color: c.accent });
  t(s, "배경 설정 · 밤", { x: b.x, y: b.y + b.h + 0.15, w: 3, h: 0.3, fontSize: 12, bold: true, color: GOLD });
  t(s, "작품 제목과 같은 이름의 상담실 간판", { x: MX, y: 5.75, w: 12, h: 0.4, fontSize: 15, bold: true, color: c.text });
  t(s, "화려하지 않은 빌딩 3층의 평범한 상담실, 초록 외벽과 따뜻한 실내등으로 차가운 도시와 구분", { x: MX, y: 6.2, w: 12, h: 0.4, fontSize: 12, color: c.muted });
}

// 05 색의 설계
pres.addSection({ title: "연출" });
{
  const { s, c } = slide("COLD", "COLOR SCRIPT", "색의 설계: 청회색에서 노을빛으로", "연출");
  const scenes = [
    ["s1-1", "S#1", "차갑고 탁한 청회색", "31404B"],
    ["s2-1", "S#2", "무채색, 낮은 채도", "3C494F"],
    ["s3-3", "S#3", "차분하고 탁한 실내광", "4A4F4F"],
    ["s4-2", "S#4", "가장 차갑고 대비 강한 하드라이트", "262D33"],
    ["s5-3", "S#5", "채도가 서서히 오르는 전환", "493E30"],
    ["s6-1", "S#6", "노을빛 · 신록 · 햇살, 높은 채도", "4B3429"],
  ];
  const cw = 1.92, gap = 0.12;
  scenes.forEach(([key, no, tone, hex], i) => {
    const x = MX + i * (cw + gap);
    img(s, key, x, 1.75, cw, 1.08, { top: true });
    rect(s, x, 3.0, cw, 0.75, "#" === hex ? c.card : hex, { line: c.line });
    t(s, `#${hex}`, { x, y: 3.0, w: cw, h: 0.75, fontFace: TC, fontSize: 11, color: "E3E8EC", align: "center", valign: "middle" });
    t(s, no, { x, y: 3.95, w: cw, h: 0.3, fontFace: TC, fontSize: 12, bold: true, color: i < 4 ? c.accent : "F2A66E" });
    t(s, tone, { x, y: 4.3, w: cw, h: 0.75, fontSize: 11, color: c.text, lineSpacingMultiple: 1.15 });
  });
  t(s, "장면별 평균 색 (스틸컷 실측)", { x: MX, y: 5.2, w: 6, h: 0.3, fontSize: 10, color: c.muted });
  rect(s, MX, 5.7, 12.0, 1.0, c.card);
  t(s, "S#5 상담실 2차 방문을 색조 전환점으로 설정, 인물의 심리 변화를 화면의 채도와 색온도로 표현", { x: MX + 0.3, y: 5.7, w: 11.4, h: 1.0, fontSize: 14, bold: true, color: c.text, valign: "middle" });
}

// 06 전반부와 후반부의 연출 대비
{
  const { s, c } = slide("TRANS", "BEFORE / AFTER", "전반부와 후반부의 연출 대비", "연출");
  const iw = 4.0;
  const a = img(s, "s2-1", MX, 1.75, iw, 2.25, { border: "8C99A6", alt: "전반부: 파티션 사이의 수현" });
  const b = img(s, "s6-2", W - MX - iw, 1.75, iw, 2.25, { border: "F28C4B", alt: "후반부: 노을 속 수현" });
  t(s, "전반부", { x: a.x, y: a.y + a.h + 0.12, w: 2, h: 0.32, fontSize: 13, bold: true, color: "B8C4CE" });
  t(s, "후반부", { x: b.x, y: b.y + b.h + 0.12, w: 2, h: 0.32, fontSize: 13, bold: true, color: "F2A66E" });
  const rows = [
    ["색", "청회색 · 무채색 · 낮은 채도", "노을빛 · 신록 · 햇살, 높은 채도"],
    ["조명", "형광등 하드라이트, 차가운 인공광", "역광 산란광, 소프트라이트"],
    ["카메라", "핸드헬드의 미세한 흔들림", "안정적인 카메라 무브"],
    ["구도", "손잡이 · 창틀 · 파티션에 갇힌 구도", "오픈된 거리, 열린 구도"],
  ];
  const cx = MX + iw + 0.25, cwid = W - 2 * MX - 2 * iw - 0.5;
  rows.forEach(([k, before, after], i) => {
    const y = 1.75 + i * 1.22;
    rect(s, cx, y, cwid, 1.1, c.card);
    t(s, k, { x: cx, y: y + 0.08, w: cwid, h: 0.3, fontSize: 12, bold: true, color: c.accent, align: "center" });
    t(s, before, { x: cx + 0.15, y: y + 0.42, w: cwid - 0.3, h: 0.28, fontSize: 10.5, color: "B8C4CE", align: "center" });
    t(s, after, { x: cx + 0.15, y: y + 0.72, w: cwid - 0.3, h: 0.28, fontSize: 10.5, color: "F2A66E", align: "center" });
  });
  t(s, "색 · 조명 · 카메라 · 구도의 동시 전환을 통한 해방감 표현", { x: MX, y: 6.55, w: 12, h: 0.35, fontSize: 14, bold: true, color: c.text });
}

// 07 생각의 시각화
{
  const { s, c } = slide("TRANS", "VISUALIZING THOUGHT", "생각의 시각화", "연출");
  const items = [
    ["s1-3", "S#1 컷03", "창에 비친 얼굴과 도시 풍경의 오버랩", "멈추지 않는 불안한 생각"],
    ["s5-4", "S#5 컷04", "먹구름으로 표현한 머릿속 소음", "에고가 만든 소음"],
    ["s5-2-1", "S#5 심상 인서트", "구름이 걷힌 맑은 도시 하늘", "생각과 자신의 분리"],
  ];
  const iw = 3.9, gap = 0.15;
  items.forEach(([key, cut, cap, meaning], i) => {
    const x = MX + i * (iw + gap);
    const f = img(s, key, x, 1.75, iw, 2.2, { top: true });
    t(s, cut, { x, y: f.y + f.h + 0.15, w: iw, h: 0.3, fontFace: TC, fontSize: 11, bold: true, color: c.accent });
    t(s, cap, { x, y: f.y + f.h + 0.5, w: iw, h: 0.35, fontSize: 13, bold: true, color: c.text });
    t(s, meaning, { x, y: f.y + f.h + 0.88, w: iw, h: 0.3, fontSize: 11, color: c.muted });
  });
  rect(s, MX, 5.6, 12.0, 1.05, c.card);
  t(s, "도윤 「생각을 바꾸려 싸우지 말고, 그냥 가만히 바라보세요. 생각과 자신을 분리하는 겁니다.」", { x: MX + 0.3, y: 5.6, w: 11.4, h: 1.05, fontSize: 14, color: c.text, valign: "middle" });
}

// 08 대사 설계: 내면의 목소리 변화
{
  const { s, c } = slide("TRANS", "VOICE", "대사 설계: 내면의 목소리 변화", "연출");
  const lines = [
    ["S#1", "수현 (V.O.)", "「그때 이직을 했어야 했나...」", "B8C4CE"],
    ["S#1", "수현 (V.O.)", "「요즘 젊은 애들한테 밀리면... 퇴직 후엔 어쩌지?」", "B8C4CE"],
    ["S#3", "도윤", "「부장님, 방금 마신 차의 온도와 향은 어땠나요?」", "D9B48A"],
    ["S#4", "수현 (V.O.)", "「나 이제... 끝났어.」", "8C99A6"],
    ["S#5", "도윤", "「'난 끝장이야'라는 그 생각은 진짜 상황이 아니라 에고가 만든 소음일 뿐이에요.」", "D9B48A"],
    ["S#6", "수현 (V.O.)", "「지금... 이 순간.」", "F28C4B"],
  ];
  lines.forEach(([sc, who, line, col], i) => {
    const y = 1.7 + i * 0.83;
    rect(s, MX, y, 12.0, 0.7, c.card);
    t(s, sc, { x: MX + 0.25, y, w: 0.8, h: 0.7, fontFace: TC, fontSize: 12, bold: true, color: col, valign: "middle" });
    t(s, who, { x: MX + 1.1, y, w: 1.5, h: 0.7, fontSize: 12, bold: true, color: c.muted, valign: "middle" });
    t(s, line, { x: MX + 2.6, y, w: 9.2, h: 0.7, fontSize: i === 5 ? 17 : 13.5, bold: i === 5, color: i === 5 ? "F28C4B" : c.text, valign: "middle" });
  });
  t(s, "과거와 미래에 대한 독백에서 현재 감각의 한마디로 수렴하는 대사 구조", { x: MX, y: 6.7, w: 12, h: 0.32, fontSize: 12, bold: true, color: c.accent });
}

// 09 127초 타임라인
{
  const { s, c } = slide("TRANS", "TIMELINE  00:00 - 02:07", "127초 타임라인", "연출");
  const scenes = [
    ["S#1", "지하철 출근길", 17, "31404B"], ["S#2", "일상 몽타주", 16, "3C494F"], ["S#3", "상담실 1차", 19, "4A4F4F"],
    ["S#4", "질책과 패닉", 16, "262D33"], ["S#5", "상담실 2차 · 전환", 29, "8A6A45"], ["S#6", "퇴근길 · 감각의 회복", 30, "C8673A"],
  ];
  const total = 127, u = 12.0 / total;
  let x = MX, acc = 0;
  scenes.forEach(([no, name, d, col]) => {
    const w = d * u;
    rect(s, x + 0.02, 2.2, w - 0.04, 1.4, col, { line: c.line });
    t(s, no, { x: x + 0.15, y: 2.32, w: w - 0.3, h: 0.3, fontFace: TC, fontSize: 12, bold: true, color: "FFFFFF" });
    t(s, name, { x: x + 0.15, y: 2.68, w: w - 0.3, h: 0.55, fontSize: 11, color: "FFFFFF" });
    t(s, `${d}초`, { x: x + 0.15, y: 3.18, w: w - 0.3, h: 0.3, fontSize: 11, bold: true, color: "FFFFFF" });
    const mm = (n) => `00:${String(Math.floor(n / 60)).padStart(2, "0")}:${String(n % 60).padStart(2, "0")}`;
    t(s, mm(acc), { x: x, y: 1.85, w: 1.2, h: 0.28, fontFace: TC, fontSize: 9.5, color: c.muted });
    acc += d; x += w;
  });
  t(s, "00:02:07", { x: MX + 12.0 - 1.2, y: 1.85, w: 1.2, h: 0.28, fontFace: TC, fontSize: 9.5, color: c.muted, align: "right" });
  const notes = [["16~19초", "S#1 ~ S#4\n불안과 압박의 전반부"], ["29초", "S#5 가장 긴 전환 구간\n색이 풀리는 시간 확보"], ["30초", "S#6 엔딩\n감각 인서트 중심의 여유"]];
  notes.forEach(([big, small], i) => {
    const nx = MX + i * 4.06;
    rect(s, nx, 4.15, 3.88, 2.3, c.card);
    t(s, big, { x: nx + 0.3, y: 4.35, w: 3.3, h: 0.7, fontSize: 28, bold: true, color: i === 0 ? "B8C4CE" : i === 1 ? c.accent : "F28C4B", valign: "middle" });
    t(s, small, { x: nx + 0.3, y: 5.15, w: 3.3, h: 0.9, fontSize: 13, color: c.text, lineSpacingMultiple: 1.2 });
  });
}

// 10 컷 설계 방식
{
  const { s, c } = slide("TRANS", "CUT SHEET", "컷 설계 방식: 6항목 기술", "연출");
  const fields = [["카메라", "핸드헬드, 고정, 로우앵글 등 카메라 운용"], ["구도", "인물을 가두는 프레임, 오픈 구도 등 화면 배치"], ["조명", "형광등 하드라이트, 노을 역광 등 빛의 성격"], ["재생시간", "컷 단위 초 단위 지정, 씬 소계와 합계 관리"], ["지문", "인물 행동과 심리의 시각적 묘사"], ["대사", "화자와 V.O. 구분"]];
  fields.forEach(([k, v], i) => {
    const y = 1.75 + i * 0.8;
    rect(s, MX, y, 5.6, 0.68, c.card);
    t(s, k, { x: MX + 0.25, y, w: 1.3, h: 0.68, fontSize: 13, bold: true, color: c.accent, valign: "middle" });
    t(s, v, { x: MX + 1.6, y, w: 3.9, h: 0.68, fontSize: 11.5, color: c.text, valign: "middle" });
  });
  const rx = 6.65, rw = W - MX - rx;
  img(s, "s5-4", rx, 1.75, rw, 3.0, { top: true, alt: "S#5 컷04 심상 컷" });
  rect(s, rx, 5.25, rw, 1.45, "F4EDE4");
  t(s, [
    { text: "S#5 컷04  ", options: { fontFace: TC, bold: true, color: "8A6A45" } },
    { text: "추상적 인서트 몽타주 · 8초", options: { bold: true, color: "2A1C16", breakLine: true } },
    { text: "조명: 무채색에서 옅은 색이 스며드는 그라데이션", options: { color: "4A3F38", breakLine: true } },
    { text: "지문: 불안한 생각들이 구름처럼 지나가는 것을 관조", options: { color: "4A3F38" } },
  ], { x: rx + 0.3, y: 5.35, w: rw - 0.6, h: 1.25, fontSize: 12, lineSpacingMultiple: 1.25 });
}

// 11~15 스토리보드 (씬이 진행될수록 배경 톤 전환)
pres.addSection({ title: "스토리보드" });
const SB = {
  "s1-1": ["컷01 · 6초", "핸드헬드", "손잡이와 창틀에 갇힌 수현"],
  "s1-2": ["컷02 · 5초", "클로즈업", "초점 없는 눈동자, 「그때 이직을 했어야 했나...」"],
  "s1-3": ["컷03 · 6초", "핸드헬드", "창에 비친 얼굴과 도시 풍경의 오버랩"],
  "s2-1": ["컷01 · 5초", "고정", "서류를 보는 굳은 표정의 수현"],
  "s2-2": ["컷02 · 5초", "핸드헬드", "식당 창가, 맛을 느끼지 못하는 점심"],
  "s2-3": ["컷03 · 6초", "롱샷", "퇴근길 인파 속 무표정한 수현"],
  "s3-1": ["컷01 · 4초", "고정 와이드샷", "빌딩 3층 상담실 '지금이라는 문'"],
  "s3-2": ["컷01 · 4초", "고정 와이드샷", "차를 우려내는 도윤, 문을 열고 들어선 수현"],
  "s3-3": ["컷02 · 5초", "투샷", "「방금 마신 차의 온도와 향은 어땠나요?」"],
  "s3-4": ["컷03 · 4초", "클로즈업", "말문이 막혀 시선을 피하는 수현"],
  "s3-5": ["컷04 · 6초", "클로즈업", "담담하게 숙제를 내어주는 도윤"],
  "s4-1": ["컷01 · 5초", "로우앵글", "「이 프로젝트, 대체 어떻게 책임질 겁니까?」"],
  "s4-1-1": ["추가 컷", "미러샷", "거울 앞에서 고개를 떨군 수현"],
  "s4-2": ["컷02 · 5초", "클로즈업", "하얗게 질린 얼굴, 「나 이제... 끝났어.」"],
  "s4-3": ["컷03 · 6초", "핸드헬드", "좁은 복도에 갇힌 패닉"],
  "s4-3-1": ["추가 컷", "미디엄샷", "건물 밖, 가슴을 움켜쥔 수현"],
  "s5-1": ["컷01 · 5초", "투샷", "따뜻한 물 한 잔을 건네는 도윤"],
  "s5-2": ["컷02 · 5초", "클로즈업", "「생각과 자신을 분리하는 겁니다.」"],
  "s5-3": ["컷03 · 6초", "클로즈업", "눈을 감고 숨소리에 귀 기울이는 수현"],
  "s5-4": ["컷04 · 8초", "심상 몽타주", "먹구름처럼 지나가는 생각"],
  "s5-2-1": ["심상 인서트", "와이드", "구름이 걷힌 도시의 하늘"],
  "s5-5": ["컷05 · 5초", "클로즈업", "부드러운 자연광 속 편안해진 표정"],
  "s6-1": ["컷01 · 6초", "롱샷", "노을 역광의 퇴근길을 천천히 걷는 수현"],
  "s6-2": ["컷02 · 4초", "클로즈업 인서트", "뺨을 스치는 저녁 바람"],
  "s6-3": ["인서트", "클로즈업", "발밑 보도블록의 감촉"],
  "s6-4": ["컷05 · 6초", "미디엄샷", "깊은 숨과 미소, 「지금... 이 순간.」"],
  "s6-5": ["컷06 · 6초", "롱샷, 페이드아웃", "노을 진 거리, 황금빛 엔딩"],
};
function board(stage, tc, title, keys, tone, cols = 3) {
  const { s, c } = slide(stage, tc, title, "스토리보드");
  const gap = 0.15, gw = 10.2, iw = (gw - (cols - 1) * gap) / cols, ih = iw / 1.777;
  rect(s, MX + gw + 0.2, 1.65, 12.0 - gw - 0.2, 5.2, c.card);
  t(s, "톤", { x: MX + gw + 0.38, y: 1.82, w: 1.4, h: 0.3, fontSize: 11, bold: true, color: c.accent });
  t(s, tone, { x: MX + gw + 0.38, y: 2.18, w: 12.0 - gw - 0.55, h: 4.5, fontSize: 11, color: c.text, lineSpacingMultiple: 1.3 });
  keys.forEach((key, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const x = MX + col * (iw + gap), y = 1.65 + row * (ih + 0.78);
    const f = img(s, key, x, y, iw, ih, { top: true });
    const [a, b, d] = SB[key];
    s.addText([
      { text: a + "   ", options: { fontFace: TC, bold: true, color: c.accent, fontSize: 10.5 } },
      { text: b, options: { color: c.text, fontSize: 10.5, bold: true, breakLine: true } },
      { text: d, options: { color: c.muted, fontSize: 10.5 } },
    ], { x, y: f.y + f.h + 0.08, w: iw, h: 0.62, fontFace: SANS, valign: "top", margin: 0, isTextBox: true, objectName: nm("sb-caption") });
  });
  return { s, c };
}
board("COLD", "S#1 00:00-00:17  ·  S#2 00:17-00:33", "S#1 지하철 출근길 · S#2 일상 몽타주", ["s1-1", "s1-2", "s1-3", "s2-1", "s2-2", "s2-3"], "차갑고 탁한 청회색\n\n형광등, 차가운 인공광\n\n무채색, 낮은 채도");
board("COLD", "S#3  00:33 - 00:52", "S#3 상담실 1차 방문", ["s3-1", "s3-2", "s3-3", "s3-4", "s3-5"], "은은하나 여전히 차분하고 탁한 실내광\n\n창가광, 차분한 톤");
board("HARD", "S#4  00:52 - 01:08", "S#4 회사, 질책과 패닉", ["s4-1", "s4-1-1", "s4-2", "s4-3", "s4-3-1"], "가장 차갑고 대비 강한 하드라이트\n\n형광등 깜빡임, 강한 그림자");
board("TRANS", "S#5  01:08 - 01:37", "S#5 상담실 2차 방문, 색조 전환점", ["s5-1", "s5-2", "s5-3", "s5-4", "s5-2-1", "s5-5"], "무채색에서 서서히 채도가 올라가는 그라데이션\n\n부드러운 자연광으로 전환");
board("WARM", "S#6  01:37 - 02:07", "S#6 퇴근길, 감각의 회복", ["s6-1", "s6-2", "s6-3", "s6-4", "s6-5"], "붉은 노을빛, 신록의 초록, 햇살의 노란빛\n\n높은 채도, 소프트라이트");

// 16 제작 공정
pres.addSection({ title: "결과물 · IP" });
{
  const { s, c } = slide("WARM", "PIPELINE", "제작 공정", "결과물 · IP");
  const steps = [["시나리오", "컷별 카메라 · 구도 · 조명 · 재생시간 설계"], ["캐릭터 · 배경 시트", "인물 3인 턴어라운드, 상담실 낮 · 밤 설정"], ["스틸컷 생성", "씬별 색조 지시를 반영한 장면 이미지"], ["영상화", "스틸컷 기반 움직임 부여"], ["편집", "127초 타임라인, 색조 전환 편집"]];
  const cw = 2.3, gap = 0.125;
  steps.forEach(([k, v], i) => {
    const x = MX + i * (cw + gap);
    rect(s, x, 1.75, cw, 2.9, c.card);
    t(s, String(i + 1).padStart(2, "0"), { x: x + 0.25, y: 1.9, w: 1, h: 0.5, fontFace: TC, fontSize: 22, bold: true, color: c.accent, valign: "middle" });
    t(s, k, { x: x + 0.25, y: 2.5, w: cw - 0.4, h: 0.65, fontSize: 14, bold: true, color: c.text });
    t(s, v, { x: x + 0.25, y: 3.15, w: cw - 0.4, h: 0.9, fontSize: 11, color: c.muted, lineSpacingMultiple: 1.15 });
    t(s, "[툴 기입]", { x: x + 0.25, y: 4.2, w: cw - 0.4, h: 0.3, fontSize: 10.5, color: GOLD });
  });
  const refs = [["char_suhyun", "캐릭터 시트"], ["bg_night", "배경 설정"], ["s5-3", "스틸컷"]];
  refs.forEach(([key, cap], i) => {
    const x = MX + i * 4.06;
    rect(s, x, 4.95, 3.88, 1.6, key === "char_suhyun" ? "F4F4F2" : c.card);
    img(s, key, x + 0.05, 5.0, 3.78, 1.5);
    t(s, cap, { x, y: 6.62, w: 3, h: 0.28, fontSize: 10, color: c.muted });
  });
}

// 17 영상 결과물
{
  const { s, c } = slide("WARM", "FILM  ver 1", "영상 결과물: 지금이라는 문 ver 1", "결과물 · IP");
  const f = img(s, "poster_film", MX, 1.7, 7.9, 4.44, { border: c.line, url: LINK_FILM, alt: "영상 대표 이미지 (클릭 시 영상 재생)" });
  t(s, "이미지 클릭 시 영상 재생 (Google Drive) / 대표 이미지: S#6 컷05", { x: MX, y: f.y + f.h + 0.15, w: 7.9, h: 0.3, fontSize: 10, color: c.muted });
  const rx = 8.9, rw = W - MX - rx;
  rect(s, rx, 1.7, rw, 4.44, c.card);
  const info = [["버전", "ver 1"], ["시나리오 러닝타임", "127초 (약 2분 7초)"], ["영상 길이", "[영상 확인 후 기입]"], ["해상도", "[영상 확인 후 기입]"]];
  info.forEach(([k, v], i) => {
    const y = 1.95 + i * 0.98;
    t(s, k, { x: rx + 0.3, y, w: rw - 0.6, h: 0.3, fontSize: 11, bold: true, color: c.accent });
    t(s, v, { x: rx + 0.3, y: y + 0.33, w: rw - 0.6, h: 0.45, fontSize: 14, color: c.text });
  });
  linkText(s, "▶  영상 보기", LINK_FILM, rx, 6.3, rw, 0.4, GOLD);
}

// 18 캐릭터 IP: 스피키 & 포니 (IP 섹션은 캐릭터 고유의 크림 · 핑크 · 하늘 톤)
{
  const { s, c } = slide("IP", "CHARACTER IP", "스피키 & 포니", "결과물 · IP");
  const a = img(s, "sp_캐릭터_시트", MX, 1.7, 6.6, 5.0, { top: true, alt: "스피키 & 포니 캐릭터 시트" });
  const rx = MX + 6.9, rw = W - MX - rx;
  img(s, "sp_브랜드_로고", rx, 1.7, rw, 2.75, { top: true, alt: "스피키 & 포니 브랜드 로고" });
  const ch = [["스피키 (Spiky)", "분위기 메이커 스피커", "작지만 빵빵한 사운드", PINK], ["포니 (Pony)", "언제나 내 편 거치대", "든든한 나의 거치대", SKYB]];
  ch.forEach(([n, role, line, col], i) => {
    const y = 4.75 + i * 0.95;
    rect(s, rx, y, rw, 0.82, c.card, { round: 0.12, line: c.line });
    s.addShape(pres.shapes.OVAL, { x: rx + 0.22, y: y + 0.29, w: 0.24, h: 0.24, fill: { color: col }, line: { type: "none" }, objectName: nm("dot") });
    t(s, n, { x: rx + 0.6, y: y + 0.08, w: 2.3, h: 0.35, fontSize: 14, bold: true, color: c.text, valign: "middle" });
    t(s, role, { x: rx + 0.6, y: y + 0.43, w: rw - 0.8, h: 0.3, fontSize: 11, color: c.muted, valign: "middle" });
  });
}

// 19 IP 확장: 이모티콘 · 웹툰 · 굿즈
{
  const { s, c } = slide("IP", "IP EXPANSION", "IP 확장: 이모티콘 · 웹툰 · 굿즈", "결과물 · IP");
  const e = img(s, "sp_이모티콘", MX, 1.7, 5.0, 3.35, { top: true, alt: "스피키 & 포니 이모티콘 32종" });
  t(s, "이모티콘 32종", { x: MX, y: e.y + e.h + 0.12, w: 3, h: 0.32, fontSize: 13, bold: true, color: c.text });
  t(s, "말풍선 · 아이콘 · 미니 스티커 포함", { x: MX, y: e.y + e.h + 0.45, w: 5, h: 0.3, fontSize: 11, color: c.muted });
  const wt = img(s, "sp_웹툰", MX + 5.25, 1.7, 2.45, 4.35, { top: true, alt: "스피키 & 포니 일상툰 10컷" });
  t(s, "일상툰 10컷", { x: MX + 5.25, y: wt.y + wt.h + 0.12, w: 2.45, h: 0.32, fontSize: 13, bold: true, color: c.text });
  const gx = MX + 7.95, gw = W - MX - gx;
  const g = img(s, "sp_굿즈_상품", gx, 1.7, gw, 2.8, { top: true, alt: "스피키 & 포니 굿즈 가이드" });
  t(s, "굿즈 가이드 12종", { x: gx, y: g.y + g.h + 0.12, w: gw, h: 0.32, fontSize: 13, bold: true, color: c.text });
  t(s, "봉제 인형, 키링, 머그컵, 폰케이스 등\n컬러 팔레트와 브랜드 패턴 정리", { x: gx, y: g.y + g.h + 0.45, w: gw, h: 0.6, fontSize: 11, color: c.muted });
  img(s, "poster_sp", gx, 5.25, 1.9, 1.27, { url: LINK_SP, alt: "스피키 & 포니 영상 (클릭 시 재생)" });
  linkText(s, "▶  스피키 & 포니 영상", LINK_SP, gx + 2.05, 5.5, gw - 2.05, 0.4, "D9586E");
  t(s, "캐릭터의 기능 (음악 · 거치) 을 성격과 관계로 확장한 IP 설계", { x: MX, y: 6.6, w: 12, h: 0.32, fontSize: 12.5, bold: true, color: "D9586E" });
}

// 20 엔딩: 노을의 퇴근길 위에 엔딩 대사
{
  const s = pres.addSlide({ masterName: "PLAIN", sectionTitle: "결과물 · IP" });
  s.addImage({ path: path.join(IMG, "s6-5.jpg"), x: 0, y: 0, w: W, h: H, altText: "노을 진 거리", objectName: nm("ending-bg") });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: H, fill: { color: "1A110C", transparency: 35 }, line: { type: "none" }, objectName: nm("ending-shade") });
  t(s, "00:02:07", { x: MX, y: 0.45, w: 3, h: 0.3, fontFace: TC, fontSize: 11, bold: true, color: "FFD9B8" });
  t(s, "「지금... 이 순간.」", { x: MX, y: 1.4, w: 12, h: 1.1, fontSize: 40, bold: true, color: "FFFFFF", valign: "middle" });
  const skills = [["색채 설계", "색온도 · 채도 변화를 통한 감정 전환"], ["컷 설계", "카메라 · 구도 · 조명 · 재생시간 명시"], ["캐릭터 IP", "브랜드 · 이모티콘 · 웹툰 · 굿즈 확장"]];
  skills.forEach(([k, v], i) => {
    const y = 3.0 + i * 0.62;
    t(s, k, { x: MX, y, w: 2.0, h: 0.45, fontSize: 14, bold: true, color: "F6C45E", valign: "middle" });
    t(s, v, { x: MX + 2.1, y, w: 7, h: 0.45, fontSize: 14, color: "FFF1E4", valign: "middle" });
  });
  s.addShape(pres.shapes.RECTANGLE, { x: MX, y: 5.45, w: 12.0, h: 1.3, fill: { color: "1A110C", transparency: 20 }, line: { type: "none" }, objectName: nm("contact") });
  t(s, "김윤성", { x: MX + 0.35, y: 5.6, w: 2, h: 0.5, fontSize: 20, bold: true, color: "FFFFFF", valign: "middle" });
  [["이메일", 2.7], ["연락처", 6.0], ["포트폴리오 링크", 9.1]].forEach(([k, x]) => {
    t(s, k, { x: MX + x, y: 5.62, w: 2.6, h: 0.28, fontSize: 10, color: "D1B29C" });
    s.addShape(pres.shapes.LINE, { x: MX + x, y: 6.4, w: 2.6, h: 0, line: { color: "D1B29C", width: 1 }, objectName: nm("blank-line") });
  });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();

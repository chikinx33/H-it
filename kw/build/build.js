// 강원 포트폴리오: 「상상 속의 하늘을 날다」
// node build.js  ->  ../강원_포트폴리오.pptx
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const IMG = path.join(__dirname, "../work/img");
const DIMS = require("../work/dims.json");
const OUT = path.join(__dirname, "../강원_포트폴리오.pptx");

const LINK_FILM = "https://drive.google.com/file/d/156VyLgx8t4YyaAT5u1FqRg_KaATZ55XH/view";
const LINK_EMO = "https://drive.google.com/file/d/1Jtj3e9SkLrNs45NBc4cImHn3D2m0-vm_/view";

// 작품 톤: 석양의 과거(주황) / 한낮의 현재(하늘색) / 밤하늘 배경(남색) / 손 그림 종이(크림)
const SERIF = "Batang";
const SANS = "Malgun Gothic";
const THEME = {
  name: "Sangsang Sky",
  headFontFace: SERIF,
  bodyFontFace: SANS,
  colors: {
    dk1: "1C1B2E", // 밤하늘 남색
    lt1: "F7EFE3", // 종이 크림
    dk2: "2A2940", // 어스름 카드
    lt2: "A8A3B8", // 보조 텍스트(어두운 배경용)
    accent1: "E8743B", // 석양 주황
    accent2: "F2B84B", // 노을 금색
    accent3: "6FA8DC", // 한낮 하늘
    accent4: "C9677A", // 저녁 장밋빛
    accent5: "3A3858", // 선
    accent6: "6B5E55", // 종이 위 보조 텍스트
    hlink: "F2B84B",
    folHlink: "F2B84B",
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: SERIF, bodyFontFace: SANS };
pres.author = "강원";
pres.title = "상상 속의 하늘을 날다: 강원 포트폴리오";

const C = pres.SchemeColor;
const NIGHT = C.text1, PAPER = C.background1, DUSK = C.text2, HAZE = C.background2;
const SUNSET = C.accent1, GOLD = C.accent2, DAY = C.accent3, ROSE = C.accent4, LINE = C.accent5, BROWN = C.accent6;
const W = 13.333, H = 7.5, MX = 0.65;

// ---------- layouts ----------
const titlePh = (color) => ({ placeholder: { options: { name: "title", type: "title", x: MX, y: 0.72, w: 12.0, h: 0.7, fontFace: SERIF, fontSize: 28, bold: true, color, align: "left", valign: "middle", margin: 0 }, text: "" } });
pres.defineSlideMaster({
  title: "FILM",
  background: { color: NIGHT },
  objects: [titlePh(PAPER),
    { text: { text: "강원  |  상상 속의 하늘을 날다", options: { x: MX, y: 7.08, w: 6, h: 0.28, fontFace: SANS, fontSize: 9, color: HAZE, margin: 0 } } }],
  slideNumber: { x: 12.18, y: 7.08, w: 0.5, h: 0.28, fontFace: SANS, fontSize: 9, color: HAZE, align: "right" },
});
pres.defineSlideMaster({
  title: "PAPER",
  background: { color: PAPER },
  objects: [titlePh(NIGHT),
    { text: { text: "강원  |  상상 속의 하늘을 날다", options: { x: MX, y: 7.08, w: 6, h: 0.28, fontFace: SANS, fontSize: 9, color: BROWN, margin: 0 } } }],
  slideNumber: { x: 12.18, y: 7.08, w: 0.5, h: 0.28, fontFace: SANS, fontSize: 9, color: BROWN, align: "right" },
});
pres.defineSlideMaster({ title: "BLACK", background: { color: "000000" }, objects: [] });

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
  const ext = key.startsWith("emo_") ? "png" : "jpg";
  const f = fit(key, x, y, w, h);
  const link = o.url ? { hyperlink: { url: o.url, tooltip: "영상 보기" } } : {};
  if (o.frame) {
    const p = o.pad ?? 0.05;
    s.addShape(pres.shapes.RECTANGLE, { x: f.x - p, y: f.y - p, w: f.w + 2 * p, h: f.h + 2 * p, fill: { color: o.frame }, line: { type: "none" }, objectName: nm("frame") });
  }
  s.addImage({ path: path.join(IMG, `${key}.${ext}`), ...f, altText: o.alt || key, objectName: nm(`img-${key}`), ...link });
  return f;
}
function t(s, text, o) {
  s.addText(text, { fontFace: SANS, color: PAPER, margin: 0, valign: "top", isTextBox: true, objectName: nm(o.name || "text"), ...o });
}
function box(s, x, y, w, h, fill, line) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: fill }, line: line ? { color: line, width: 0.75 } : { type: "none" }, objectName: nm("box") });
}
function slide(master, label, title, section) {
  const s = pres.addSlide({ masterName: master, sectionTitle: section });
  t(s, label, { x: MX, y: 0.38, w: 8, h: 0.32, fontSize: 11, bold: true, color: master === "PAPER" ? SUNSET : GOLD, charSpacing: 3, valign: "middle", name: "label" });
  s.addText(title, { placeholder: "title" });
  return s;
}
function linkText(s, label, url, x, y, w, h, color) {
  s.addText([{ text: label, options: { hyperlink: { url, tooltip: url }, color, bold: true } }], {
    x, y, w, h, fontFace: SANS, fontSize: 13, align: "left", valign: "middle", margin: 0, isTextBox: true, objectName: nm("link"),
  });
}

// =====================================================================
// 01 표지: 소년이 노을을 향해 두 팔을 벌린 컷을 전면에, 위아래 시네마 바
pres.addSection({ title: "작품" });
{
  const s = pres.addSlide({ masterName: "BLACK", sectionTitle: "작품" });
  s.addImage({ path: path.join(IMG, "s1-5.jpg"), x: 0, y: 0, w: W, h: H, altText: "노을을 향해 두 팔을 벌린 소년", objectName: nm("cover-img") });
  box(s, 0, 0, W, 0.55, "000000");
  box(s, 0, H - 0.55, W, 0.55, "000000");
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 4.35, w: W, h: 2.6, fill: { color: "000000", transparency: 45 }, line: { type: "none" }, objectName: nm("cover-shade") });
  t(s, "AI 2D 애니메이션 포트폴리오 2026", { x: MX, y: 4.55, w: 8, h: 0.35, fontSize: 13, bold: true, color: GOLD, charSpacing: 2 });
  t(s, "상상 속의 하늘을 날다", { x: MX, y: 4.95, w: 9, h: 0.95, fontFace: SERIF, fontSize: 44, bold: true, color: PAPER, valign: "middle" });
  t(s, "강원", { x: MX, y: 6.05, w: 4, h: 0.5, fontFace: SERIF, fontSize: 22, bold: true, color: PAPER, valign: "middle" });
  t(s, "S1-C5  「나도.. 하늘을.. 날고 싶다」", { x: 8.2, y: 6.18, w: 4.5, h: 0.3, fontSize: 10, color: PAPER, align: "right" });
}

// 02 작품 소개
{
  const s = slide("FILM", "INTRO", "작품 소개", "작품");
  img(s, "s1-4", 6.55, 1.75, 6.13, 3.45, { frame: PAPER, pad: 0.05, alt: "소년의 눈동자에 비친 송골매" });
  t(s, "S1-C4  눈동자 속 송골매", { x: 6.55, y: 5.3, w: 6, h: 0.3, fontSize: 10, color: HAZE });
  t(s, "하늘을 날고 싶던 8살 소년이\n중년의 화가가 되어 다시 송골매를 마주하고,\n상상 속 비행으로 잊혀진 꿈을 깨우는 이야기", {
    x: MX, y: 1.75, w: 5.6, h: 1.8, fontFace: SERIF, fontSize: 20, bold: true, color: PAPER, lineSpacingMultiple: 1.3,
  });
  const tags = [["2D 셀 애니메이션", SUNSET], ["약 3~5분 단편", GOLD], ["과거와 현재의 교차", DAY]];
  tags.forEach(([k, col], i) => {
    s.addText(k, { x: MX + i * 1.9, y: 3.85, w: 1.78, h: 0.4, fontFace: SANS, fontSize: 11, bold: true, color: NIGHT, fill: { color: col }, align: "center", valign: "middle", margin: 0, objectName: nm("tag") });
  });
  box(s, MX, 4.6, 5.6, 1.5, DUSK);
  t(s, "작가의 메시지", { x: MX + 0.3, y: 4.75, w: 3, h: 0.3, fontSize: 11, bold: true, color: GOLD });
  t(s, "「날개가 없어도, 마음만 있다면,\n우리는 언제든 하늘을 날 수 있다.」", { x: MX + 0.3, y: 5.1, w: 5.1, h: 0.85, fontFace: SERIF, fontSize: 15, color: PAPER, lineSpacingMultiple: 1.2 });
}

// 03 제작 기본 사양
{
  const s = slide("PAPER", "SPEC", "제작 기본 사양", "작품");
  const specs = [
    ["러닝타임", "약 3~5분", "단편 애니메이션"],
    ["최종 영상", "16:9 FHD", "1920 × 1080"],
    ["프레임 레이트", "24fps", "영화 같은 움직임"],
    ["배경 원본", "2K~4K", "고해상도 제작 후 FHD 출력"],
    ["컷 전환", "디졸브 · 오버랩", "장면 간 시간의 연결감 강조"],
    ["구성", "6씬 · 30컷", "+ 엔딩 자막"],
  ];
  const cw = 3.9, ch = 2.45, gx = 0.2, gy = 0.25;
  specs.forEach(([k, v, sub], i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = MX + col * (cw + gx), y = 1.75 + row * (ch + gy);
    box(s, x, y, cw, ch, "FFFFFF", "E3D6C3");
    t(s, k, { x: x + 0.35, y: y + 0.3, w: cw - 0.7, h: 0.3, fontSize: 12, bold: true, color: SUNSET });
    t(s, v, { x: x + 0.35, y: y + 0.75, w: cw - 0.7, h: 0.8, fontFace: SERIF, fontSize: 30, bold: true, color: NIGHT, valign: "middle" });
    t(s, sub, { x: x + 0.35, y: y + 1.7, w: cw - 0.7, h: 0.4, fontSize: 12, color: BROWN });
  });
}

// 04 비주얼 방향
{
  const s = slide("FILM", "LOOK", "비주얼 방향: 손 그림 셀 애니메이션", "작품");
  const items = [
    ["영상 스타일", "전통적인 2D 셀 애니메이션\n90~2000년대 극장판 애니메이션의 평면적이고 단순한 표현\n사실적인 디지털 페인팅 배제"],
    ["연출 방향", "따뜻한 손 그림과 자연스러운 카메라 움직임\n서정적인 빛과 힐링 감성, 따뜻한 색조"],
    ["배경 원칙", "단색 또는 옅은 그라데이션의 단순한 배경\n인물과 하늘에 시선 집중"],
  ];
  items.forEach(([k, v], i) => {
    const y = 1.75 + i * 1.62;
    t(s, k, { x: MX, y, w: 4.8, h: 0.35, fontSize: 13, bold: true, color: GOLD });
    t(s, v, { x: MX, y: y + 0.42, w: 5.2, h: 1.1, fontSize: 13, color: PAPER, lineSpacingMultiple: 1.2 });
  });
  img(s, "s1-1", 6.35, 1.75, 6.33, 3.56, { frame: PAPER, pad: 0.05, alt: "1960년대 서울 옥상의 노을" });
  img(s, "s2-1", 6.35, 5.45, 3.07, 1.5, { frame: PAPER, pad: 0.04, alt: "화실의 강 작가" });
  img(s, "s3-2", 9.61, 5.45, 3.07, 1.5, { frame: PAPER, pad: 0.04, alt: "역광 노을 속 화실 문" });
}

// 05 캐릭터
pres.addSection({ title: "인물 · 공간" });
{
  const s = slide("PAPER", "CHARACTERS", "등장인물", "인물 · 공간");
  const a = img(s, "char_kang", MX, 1.7, 5.05, 4.0, { frame: "FFFFFF", pad: 0.08, alt: "강 작가 캐릭터 시트" });
  t(s, "강 작가", { x: a.x, y: 5.95, w: 2.2, h: 0.45, fontFace: SERIF, fontSize: 20, bold: true, color: NIGHT, valign: "middle" });
  t(s, "옥탑 화실의 중년 화가, 어린 시절의 꿈을 잊은 채 지내는 일상", { x: a.x, y: 6.45, w: 5.1, h: 0.3, fontSize: 11, color: BROWN });
  const b = img(s, "char_boy", 6.0, 1.7, 2.65, 4.0, { frame: "FFFFFF", pad: 0.08, alt: "어린 시절 캐릭터 시트" });
  t(s, "어린 시절", { x: b.x, y: 5.95, w: 2.8, h: 0.45, fontFace: SERIF, fontSize: 20, bold: true, color: NIGHT, valign: "middle" });
  t(s, "1960년대 서울, 8살 소년", { x: b.x, y: 6.45, w: 3, h: 0.3, fontSize: 11, color: BROWN });
  const cf = fit("s2-4", 9.0, 1.7, 3.68, 4.0);
  img(s, "s2-4", 9.0, 1.7, 3.68, cf.h, { frame: "FFFFFF", pad: 0.08, alt: "하늘을 나는 송골매" });
  t(s, "송골매", { x: 9.0, y: 1.7 + cf.h + 0.25, w: 3, h: 0.45, fontFace: SERIF, fontSize: 20, bold: true, color: NIGHT, valign: "middle" });
  t(s, "과거와 현재를 잇는 매개\n소년과 강 작가가 같은 하늘에서 마주하는 존재\n클라이맥스에서 두 인물과 하나로 융합", { x: 9.0, y: 1.7 + cf.h + 0.75, w: 3.68, h: 1.0, fontSize: 11, color: BROWN, lineSpacingMultiple: 1.2 });
}

// 06 공간과 시간
{
  const s = slide("FILM", "TIME & PLACE", "공간과 시간: 석양의 과거, 한낮의 현재", "인물 · 공간");
  const iw = 5.9;
  const L = img(s, "s1-1", MX, 1.75, iw, 3.32, { frame: SUNSET, pad: 0.06, alt: "과거: 1960년대 서울 옥상" });
  const R = img(s, "bg_studio", W - MX - iw, 1.75, iw, 3.32, { frame: DAY, pad: 0.06, alt: "현재: 옥탑 화실" });
  const col = (x, head, colr, place, feel) => {
    s.addText(head, { x, y: 5.3, w: 1.2, h: 0.4, fontFace: SANS, fontSize: 12, bold: true, color: NIGHT, fill: { color: colr }, align: "center", valign: "middle", margin: 0, objectName: nm("tag") });
    t(s, place, { x: x + 1.4, y: 5.3, w: 4.4, h: 0.4, fontSize: 14, bold: true, color: PAPER, valign: "middle" });
    t(s, feel, { x, y: 5.85, w: 5.9, h: 0.7, fontSize: 12, color: HAZE, lineSpacingMultiple: 1.2 });
  };
  col(L.x - 0.06, "과거", SUNSET, "1960년대 서울 골목, 연립주택 옥상 · 석양", "붉은 노을과 흔들리는 빨래\n동경과 순수함을 담은 따뜻한 색온도");
  col(R.x - 0.06, "현재", DAY, "옥탑의 작은 화실 · 어느 오후", "창가로 드는 햇살과 맑은 하늘\n잊혀진 꿈과 일상을 담은 밝은 색온도");
}

// 07 6씬 구성과 핵심 감정
pres.addSection({ title: "연출" });
{
  const s = slide("FILM", "STRUCTURE", "6씬 구성과 핵심 감정", "연출");
  const scenes = [
    ["SCENE 1", "어린 시절", "옥상 · 석양", "동경\n순수함\n첫 번째 꿈", SUNSET, NIGHT],
    ["SCENE 2", "중년", "화실 · 오후", "세월\n기억\n잊혀진 꿈", DAY, NIGHT],
    ["SCENE 3", "충동", "옥상 · 석양", "억눌렸던 꿈의\n각성", GOLD, NIGHT],
    ["SCENE 4", "교감", "환상", "교감\n환상\n해방", "F6D58E", NIGHT],
    ["SCENE 5", "각성", "지는 해", "꿈\n현실\n여운", ROSE, NIGHT],
    ["SCENE 6", "엔딩", "검은 화면 · 자막", "메시지와\n여운", "000000", PAPER],
  ];
  const cw = 1.9, gap = 0.12;
  scenes.forEach(([no, name, where, emo, fill, tc], i) => {
    const x = MX + i * (cw + gap);
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.75, w: cw, h: 4.55, fill: { color: fill }, line: fill === "000000" ? { color: LINE, width: 0.75 } : { type: "none" }, objectName: nm("scene-card") });
    t(s, no, { x: x + 0.2, y: 1.95, w: cw - 0.4, h: 0.3, fontSize: 10, bold: true, color: tc, charSpacing: 2 });
    t(s, name, { x: x + 0.2, y: 2.3, w: cw - 0.4, h: 0.55, fontFace: SERIF, fontSize: 22, bold: true, color: tc, valign: "middle" });
    t(s, where, { x: x + 0.2, y: 2.9, w: cw - 0.4, h: 0.35, fontSize: 11, color: tc });
    t(s, "핵심 감정", { x: x + 0.2, y: 4.1, w: cw - 0.4, h: 0.3, fontSize: 10, bold: true, color: tc });
    t(s, emo, { x: x + 0.2, y: 4.45, w: cw - 0.4, h: 1.6, fontSize: 14, bold: true, color: tc, lineSpacingMultiple: 1.15 });
  });
  t(s, "장면의 빛과 시간대 변화를 통한 감정 흐름 설계", { x: MX, y: 6.5, w: 12, h: 0.4, fontSize: 14, bold: true, color: GOLD });
}

// 08 반복 모티프
{
  const s = slide("FILM", "MOTIF", "반복 모티프: 송골매와 눈동자", "연출");
  const items = [
    ["s1-4", "S1-C4", "소년의 눈동자 속 송골매"],
    ["s2-4", "S2-C4", "중년의 강 작가가 다시 마주한 송골매"],
    ["s1-5", "S1-C5", "송골매를 향해 두 팔을 벌린 소년"],
  ];
  const iw = 3.88, gap = 0.18;
  items.forEach(([key, cut, cap], i) => {
    const x = MX + i * (iw + gap);
    const f = img(s, key, x, 1.75, iw, 2.18, { frame: PAPER, pad: 0.04 });
    t(s, cut, { x, y: f.y + f.h + 0.15, w: 1.2, h: 0.3, fontSize: 11, bold: true, color: GOLD });
    t(s, cap, { x, y: f.y + f.h + 0.45, w: iw, h: 0.35, fontSize: 12, color: PAPER });
  });
  const pts = [
    ["같은 피사체의 반복", "S1 소년의 시선, S2 · S3 강 작가의 시선으로 이어지는 송골매"],
    ["눈동자 클로즈업의 반복", "S1-C4, S2-C5, S3-C5 눈동자 속 송골매와 플래시백"],
    ["같은 동작의 반복", "S1-C5 소년과 S3-C6 강 작가의 두 팔 벌린 동작"],
  ];
  pts.forEach(([k, v], i) => {
    const y = 4.95 + i * 0.62;
    s.addShape(pres.shapes.OVAL, { x: MX, y: y + 0.11, w: 0.2, h: 0.2, fill: { color: [SUNSET, GOLD, DAY][i] }, line: { type: "none" }, objectName: nm("dot") });
    t(s, k, { x: MX + 0.4, y, w: 2.8, h: 0.42, fontSize: 13, bold: true, color: PAPER, valign: "middle" });
    t(s, v, { x: MX + 3.3, y, w: 8.7, h: 0.42, fontSize: 13, color: HAZE, valign: "middle" });
  });
}

// 09 빛의 연출
{
  const s = slide("FILM", "LIGHT", "빛의 연출: 석양, 한낮, 역광 노을", "연출");
  const items = [
    ["s1-6", "석양", "S1-C6", "노을 속 소년의 실루엣, 사라지는 송골매", SUNSET],
    ["s2-3", "한낮", "S2-C3", "창밖 하늘을 바라보는 강 작가의 일상", DAY],
    ["s3-2", "역광 노을", "S3-C2~C3", "문 너머 쏟아지는 노을빛, 꿈의 각성", GOLD],
  ];
  const iw = 3.88, gap = 0.18;
  items.forEach(([key, light, cut, cap, col], i) => {
    const x = MX + i * (iw + gap);
    const f = img(s, key, x, 1.75, iw, 2.18, { frame: col, pad: 0.06 });
    s.addText(light, { x, y: f.y + f.h + 0.25, w: 1.4, h: 0.4, fontFace: SANS, fontSize: 12, bold: true, color: NIGHT, fill: { color: col }, align: "center", valign: "middle", margin: 0, objectName: nm("tag") });
    t(s, cut, { x: x + 1.55, y: f.y + f.h + 0.25, w: 2.2, h: 0.4, fontSize: 11, bold: true, color: HAZE, valign: "middle" });
    t(s, cap, { x, y: f.y + f.h + 0.8, w: iw, h: 0.6, fontSize: 12, color: PAPER });
  });
  t(s, "현재의 하늘과 과거의 하늘을 오버랩하여 시간의 연결감 강조", { x: MX, y: 5.2, w: 12, h: 0.4, fontSize: 14, bold: true, color: GOLD });
  t(s, "씬별 빛의 흐름", { x: MX, y: 5.8, w: 3, h: 0.3, fontSize: 11, bold: true, color: HAZE });
  const flow = [["S1 석양", SUNSET, NIGHT], ["S2 한낮", DAY, NIGHT], ["S3 석양", GOLD, NIGHT], ["S4 환상", "F6D58E", NIGHT], ["S5 지는 해", ROSE, NIGHT], ["S6 검은 화면", "000000", PAPER]];
  const fw = 12.03 / 6;
  flow.forEach(([k, col, tc], i) => {
    s.addText(k, { x: MX + i * fw, y: 6.12, w: fw - 0.04, h: 0.5, fontFace: SANS, fontSize: 12, bold: true, color: tc, fill: { color: col }, line: col === "000000" ? { color: LINE, width: 0.75 } : undefined, align: "center", valign: "middle", margin: 0, objectName: nm("light-flow") });
  });
}

// 10~12 스토리보드
pres.addSection({ title: "스토리보드" });
const SB = {
  "s1-1": ["S1-C1", "익스트림 롱샷 (부감)", "1960년대 서울 골목, 연립주택 옥상의 노을과 빨래"],
  "s1-2": ["S1-C2", "롱샷 (아이레벨)", "옥상 난간에서 하늘을 바라보는 8살 소년"],
  "s1-3": ["S1-C3", "소년 시점 (POV)", "하늘 높이 원을 그리며 나는 송골매"],
  "s1-4": ["S1-C4", "클로즈업", "눈동자에 선명하게 비친 송골매"],
  "s1-5": ["S1-C5", "미디엄, 로우앵글", "두 팔을 벌린 소년 「나도.. 하늘을.. 날고 싶다」"],
  "s1-6": ["S1-C6", "익스트림 롱샷", "노을 속 소년과 사라지는 송골매, 디졸브"],
  "s2-1": ["S2-C1", "풀샷", "햇살 드는 작은 화실에서 그림을 그리는 강 작가"],
  "s2-2": ["S2-C2", "클로즈업", "얼음이 든 아이스커피를 드는 손"],
  "s2-3": ["S2-C3", "미디엄샷", "커피를 마시며 무심코 창밖 하늘을 바라봄"],
  "s2-4": ["S2-C4", "시점 (POV)", "창밖 하늘을 원을 그리며 나는 송골매"],
  "s2-6": ["S2-C6", "미디엄샷", "창밖에 시선을 둔 채 천천히 일어서는 강 작가"],
  "s3-1": ["S3-C1", "미디엄샷", "무언가에 홀린 듯 화실 문으로 향하는 강 작가"],
  "s3-2": ["S3-C2~C3", "도어 클로즈업, 역광 롱샷", "문을 열자 쏟아지는 노을빛과 길게 드리운 그림자"],
};
function caption(s, key, x, y, w) {
  const [a, b, c] = SB[key];
  s.addText([
    { text: a + "   ", options: { bold: true, color: GOLD, fontSize: 11 } },
    { text: b, options: { color: PAPER, fontSize: 11, breakLine: true } },
    { text: c, options: { color: HAZE, fontSize: 10.5 } },
  ], { x, y, w, h: 0.62, fontFace: SANS, valign: "top", margin: 0, isTextBox: true, objectName: nm("sb-caption") });
}
function textCut(s, x, y, w, h, cut, shot, desc) {
  box(s, x, y, w, h, DUSK, LINE);
  t(s, cut, { x: x + 0.2, y: y + 0.2, w: w - 0.4, h: 0.3, fontSize: 11, bold: true, color: GOLD });
  t(s, shot, { x: x + 0.2, y: y + 0.55, w: w - 0.4, h: 0.3, fontSize: 12, bold: true, color: PAPER });
  t(s, desc, { x: x + 0.2, y: y + 0.9, w: w - 0.4, h: h - 1.05, fontSize: 11, color: HAZE, lineSpacingMultiple: 1.15 });
}
function board(num, title, cells) {
  const s = slide("FILM", num, title, "스토리보드");
  const iw = 3.88, ih = 2.18, gx = 0.18;
  cells.forEach((cell, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = MX + col * (iw + gx), y = 1.6 + row * 2.75;
    if (typeof cell === "string") {
      const f = img(s, cell, x, y, iw, ih, { frame: PAPER, pad: 0.03 });
      caption(s, cell, x, f.y + f.h + 0.1, iw);
    } else {
      textCut(s, x, y, iw, ih, ...cell);
    }
  });
  return s;
}
board("SCENE 1  어린 시절", "옥상 / 석양", ["s1-1", "s1-2", "s1-3", "s1-4", "s1-5", "s1-6"]);
board("SCENE 2  중년", "화실 / 어느 오후", ["s2-1", "s2-2", "s2-3", "s2-4", ["S2-C5", "눈 클로즈업", "강 작가의 눈동자 속 어린 시절 옥상 장면, 인서트 플래시백"], "s2-6"]);
{
  const s = slide("FILM", "SCENE 3  충동", "하늘을 보다 / 석양", "스토리보드");
  const iw = 5.95;
  ["s3-1", "s3-2"].forEach((key, i) => {
    const x = MX + i * (iw + 0.13);
    const f = img(s, key, x, 1.6, iw, 3.3, { frame: PAPER, pad: 0.03 });
    caption(s, key, x, f.y + f.h + 0.12, iw);
  });
  const rest = [
    ["S3-C4", "롱샷", "원을 그리며 나는 송골매를 우두커니 서서 바라보는 강 작가"],
    ["S3-C5", "익스트림 클로즈업", "눈동자 속으로 점점 다가오는 송골매"],
    ["S3-C6", "미디엄샷", "송골매가 나는 하늘로 두 팔을 뻗는 강 작가 「하늘을… 날고.. 싶다」"],
  ];
  rest.forEach(([a, b, c], i) => {
    const x = MX + i * (3.97 + 0.11);
    box(s, x, 5.7, 3.97, 1.15, DUSK, LINE);
    t(s, `${a}   ${b}`, { x: x + 0.2, y: 5.8, w: 3.6, h: 0.3, fontSize: 11, bold: true, color: GOLD });
    t(s, c, { x: x + 0.2, y: 6.12, w: 3.6, h: 0.65, fontSize: 10.5, color: HAZE });
  });
}

// 13 클라이맥스 설계
pres.addSection({ title: "클라이맥스 · 엔딩" });
{
  const s = slide("FILM", "SCENE 4  교감", "클라이맥스 설계: 하나가 되는 순간", "클라이맥스 · 엔딩");
  const cx = 3.55, cy = 4.05, r = 1.35;
  [["어린 소년", cx - 0.95, cy - 0.55, SUNSET], ["강 작가", cx + 0.95, cy - 0.55, DAY], ["송골매", cx, cy + 1.0, GOLD]].forEach(([k, x, y, col]) => {
    s.addShape(pres.shapes.OVAL, { x: x - r, y: y - r, w: 2 * r, h: 2 * r, fill: { color: col, transparency: 35 }, line: { color: col, width: 1 }, objectName: nm("venn") });
  });
  t(s, "어린 소년", { x: cx - 2.45, y: cy - 1.45, w: 1.6, h: 0.4, fontFace: SERIF, fontSize: 15, bold: true, color: PAPER, align: "center" });
  t(s, "강 작가", { x: cx + 0.85, y: cy - 1.45, w: 1.6, h: 0.4, fontFace: SERIF, fontSize: 15, bold: true, color: PAPER, align: "center" });
  t(s, "송골매", { x: cx - 0.8, y: cy + 1.6, w: 1.6, h: 0.4, fontFace: SERIF, fontSize: 15, bold: true, color: PAPER, align: "center" });
  t(s, "하나의\n존재", { x: cx - 0.6, y: cy - 0.3, w: 1.2, h: 0.75, fontFace: SERIF, fontSize: 14, bold: true, color: NIGHT, align: "center", valign: "middle" });
  const rx = 7.0, rw = W - MX - rx;
  const cuts = [
    ["S4-C1", "클로즈업", "강 작가를 향해 다가오는 송골매"],
    ["S4-C2", "교차 클로즈업", "인간의 눈과 송골매의 눈이 마주치는 찰나"],
    ["S4-C3", "판타지 오버랩", "소년 · 강 작가 · 송골매가 하나로 겹쳐짐"],
    ["S4-C4", "로우앵글, 부감", "날개를 펼친 존재가 옥상에서 떠오름"],
    ["S4-C5", "롱샷 트래킹", "구름 사이 비행, 아래로 펼쳐지는 도시"],
    ["S4-C6", "시점 샷 (POV)", "석양빛이 날개에 비치는 해방감"],
  ];
  cuts.forEach(([a, b, c], i) => {
    const y = 1.7 + i * 0.66;
    box(s, rx, y, rw, 0.56, DUSK);
    t(s, a, { x: rx + 0.2, y, w: 0.85, h: 0.56, fontSize: 11, bold: true, color: GOLD, valign: "middle" });
    t(s, b, { x: rx + 1.05, y, w: 1.6, h: 0.56, fontSize: 11, bold: true, color: PAPER, valign: "middle" });
    t(s, c, { x: rx + 2.65, y, w: rw - 2.8, h: 0.56, fontSize: 11, color: HAZE, valign: "middle" });
  });
  t(s, "AI 이미지와 영상 보간 작업을 집중 투입하는 작품의 클라이맥스", { x: rx, y: 5.8, w: rw, h: 0.6, fontSize: 13, bold: true, color: GOLD, valign: "middle" });
}

// 14 각성과 엔딩
{
  const s = slide("FILM", "SCENE 5 · 6  각성과 엔딩", "호접지몽, 그리고 검은 화면의 자막", "클라이맥스 · 엔딩");
  const lw = 5.6;
  box(s, MX, 1.7, lw, 5.1, DUSK);
  t(s, "SCENE 5  각성 / 지는 해", { x: MX + 0.3, y: 1.9, w: lw - 0.6, h: 0.35, fontSize: 12, bold: true, color: ROSE });
  const s5 = ["빛이 옅어지며 흩어지는 비행의 존재", "다시 소년 · 강 작가 · 송골매 셋으로 분리", "옥상에 홀로 남은 강 작가", "석양 속으로 멀어지는 송골매의 실루엣", "아쉬움과 미소가 함께 담긴 클로즈업", "바람에 옷자락이 흔들리는 역광 뒷모습"];
  s.addText(s5.map((x, j) => ({ text: x, options: { bullet: true, breakLine: j < s5.length - 1 } })), {
    x: MX + 0.3, y: 2.4, w: lw - 0.6, h: 3.0, fontFace: SANS, fontSize: 12, color: PAPER, valign: "top", margin: 0, paraSpaceAfter: 6, isTextBox: true, objectName: nm("s5"),
  });
  t(s, "긴 여운의 화면으로 마무리", { x: MX + 0.3, y: 6.15, w: lw - 0.6, h: 0.4, fontSize: 13, bold: true, color: GOLD });
  const rx = MX + lw + 0.3, rw = W - MX - rx;
  box(s, rx, 1.7, rw, 5.1, "000000", LINE);
  t(s, "SCENE 6  엔딩 / 검은 화면에 자막", { x: rx + 0.35, y: 1.9, w: rw - 0.7, h: 0.35, fontSize: 12, bold: true, color: HAZE });
  t(s, "하늘을 날고 싶던 동심은\n상상 속에 추억의 한 조각으로 남고,\n간절함은\n꿈을 이루어지게 한다는 것을,\n이 시대를 살아가는 나에게\n희망의 메시지로 전해준다.", {
    x: rx + 0.35, y: 2.45, w: rw - 0.7, h: 3.0, fontFace: SERIF, fontSize: 15, color: PAPER, lineSpacingMultiple: 1.35,
  });
  t(s, "타자기 소리와 함께 한 자씩 쳐지는 흰 자막, 페이드아웃", { x: rx + 0.35, y: 6.15, w: rw - 0.7, h: 0.4, fontSize: 12, color: HAZE });
}

// 15 제작 공정
pres.addSection({ title: "결과물 · IP" });
{
  const s = slide("PAPER", "PIPELINE", "제작 공정", "결과물 · IP");
  const steps = [
    ["AI 스틸컷 생성", "시나리오 컷별 구도와 지문 기반의 장면 이미지 생성"],
    ["컷별 이미지 통일", "캐릭터 시트와 배경 시트를 기준으로 인물 · 공간 일관성 유지"],
    ["AI 영상화", "스틸컷에 카메라 움직임과 동작 부여"],
    ["편집 · 사운드", "디졸브 · 오버랩 전환, 효과음과 BGM 설계"],
  ];
  const cw = 2.86, gap = 0.2;
  steps.forEach(([k, v], i) => {
    const x = MX + i * (cw + gap);
    box(s, x, 1.75, cw, 3.1, "FFFFFF", "E3D6C3");
    t(s, `0${i + 1}`, { x: x + 0.3, y: 1.95, w: 1, h: 0.6, fontFace: SERIF, fontSize: 30, bold: true, color: SUNSET, valign: "middle" });
    t(s, k, { x: x + 0.3, y: 2.7, w: cw - 0.6, h: 0.45, fontSize: 15, bold: true, color: NIGHT });
    t(s, v, { x: x + 0.3, y: 3.2, w: cw - 0.6, h: 1.0, fontSize: 12, color: BROWN, lineSpacingMultiple: 1.15 });
    t(s, "[툴 기입]", { x: x + 0.3, y: 4.35, w: cw - 0.6, h: 0.3, fontSize: 11, color: SUNSET });
  });
  const refs = [["char_kang", "캐릭터 시트"], ["bg_studio", "배경 시트"], ["s2-6", "스틸컷"]];
  refs.forEach(([key, cap], i) => {
    const x = MX + i * (3.97 + 0.11);
    const f = img(s, key, x + (i === 0 ? 0 : 0), 5.1, 3.97, 1.45, { frame: "FFFFFF", pad: 0.05 });
    t(s, cap, { x: f.x, y: 6.65, w: 2, h: 0.28, fontSize: 10, color: BROWN });
  });
}

// 16 영상 결과물
{
  const s = slide("FILM", "FILM", "영상 결과물: 상상 속의 하늘을 날다 #1", "결과물 · IP");
  const f = img(s, "poster_film", MX, 1.7, 7.9, 4.44, { frame: PAPER, pad: 0.05, url: LINK_FILM, alt: "영상 대표 이미지 (클릭 시 영상 재생)" });
  t(s, "이미지 클릭 시 영상 재생 (Google Drive) / 대표 이미지: S1-C4", { x: MX, y: f.y + f.h + 0.15, w: 7.9, h: 0.3, fontSize: 10, color: HAZE });
  const rx = 8.85, rw = W - MX - rx;
  box(s, rx, 1.7, rw, 4.44, DUSK);
  const info = [["파일", "상상 속의 하늘을 날다-#1"], ["영상 규격", "16:9 FHD · 24fps (시나리오 기준)"], ["영상 길이", "[영상 확인 후 기입]"], ["포함 범위", "[영상 확인 후 기입]"]];
  info.forEach(([k, v], i) => {
    const y = 1.95 + i * 0.98;
    t(s, k, { x: rx + 0.3, y, w: rw - 0.6, h: 0.3, fontSize: 11, bold: true, color: GOLD });
    t(s, v, { x: rx + 0.3, y: y + 0.33, w: rw - 0.6, h: 0.5, fontSize: 13, color: PAPER });
  });
  linkText(s, "▶  영상 보기", LINK_FILM, rx, 6.35, rw, 0.4, GOLD);
}

// 17 캐릭터 IP: 야구 이모티콘
{
  const s = slide("PAPER", "CHARACTER IP 1", "캐릭터 IP: 야구 소년 이모티콘", "결과물 · IP");
  const pos = [["emo_타자", "타자"], ["emo_투수", "투수"], ["emo_포수", "포수"], ["emo_유격수", "유격수"]];
  const cw = 1.68, gap = 0.15;
  pos.forEach(([key, k], i) => {
    const x = MX + i * (cw + gap);
    box(s, x, 1.75, cw, 2.75, "FFFFFF", "E3D6C3");
    img(s, key, x + 0.12, 1.9, cw - 0.24, 2.1);
    t(s, k, { x, y: 4.08, w: cw, h: 0.35, fontSize: 13, bold: true, color: NIGHT, align: "center" });
  });
  t(s, "포지션별 동작을 통한 캐릭터 성격 표현", { x: MX, y: 4.75, w: 7.2, h: 0.4, fontSize: 14, bold: true, color: NIGHT });
  t(s, "정지 이모티콘 4종 (타자 / 투수 / 포수 / 유격수)\n움직이는 이모티콘 1종 「아웃!」 (9:16, 5초)", { x: MX, y: 5.25, w: 7.2, h: 0.8, fontSize: 12, color: BROWN, lineSpacingMultiple: 1.2 });
  const rx = 8.25, rw = W - MX - rx;
  box(s, rx, 1.75, rw, 5.0, NIGHT);
  t(s, "움직이는 이모티콘 「아웃!」", { x: rx + 0.3, y: 1.95, w: rw - 0.6, h: 0.35, fontSize: 12, bold: true, color: GOLD });
  const pw = 1.4;
  img(s, "poster_emo", rx + 0.3, 2.45, pw, 2.49, { url: LINK_EMO, alt: "이모티콘 아웃 (클릭 시 재생)" });
  [1, 3, 5].forEach((n, i) => img(s, `emoanim_${n}`, rx + 0.3 + pw + 0.2 + i * 0.75, 2.95, 0.65, 1.16));
  t(s, "프레임 흐름", { x: rx + pw + 0.5, y: 2.55, w: 2.4, h: 0.3, fontSize: 10, color: HAZE });
  t(s, "이미지 클릭 시 재생", { x: rx + 0.3, y: 5.35, w: rw - 0.6, h: 0.3, fontSize: 10, color: HAZE });
  linkText(s, "▶  이모티콘 보기", LINK_EMO, rx + 0.3, 5.85, rw - 0.6, 0.4, GOLD);
}

// 18 캐릭터 IP: 굿즈
{
  const s = slide("PAPER", "CHARACTER IP 2", "캐릭터 IP: 스포츠 말 캐릭터 굿즈", "결과물 · IP");
  const goods = [["goods_tumbler", "텀블러", "골프 스윙"], ["goods_bottle", "보틀", "러닝"], ["goods_case", "휴대폰 케이스", "농구 드리블"], ["goods_cap", "모자", "점프"]];
  const cw = 2.86, gap = 0.2;
  goods.forEach(([key, k, pose], i) => {
    const x = MX + i * (cw + gap);
    box(s, x, 1.75, cw, 4.1, "FFFFFF", "E3D6C3");
    img(s, key, x + 0.15, 1.9, cw - 0.3, 3.4);
    t(s, k, { x: x + 0.25, y: 5.4, w: 1.6, h: 0.35, fontSize: 13, bold: true, color: NIGHT, valign: "middle" });
    t(s, pose, { x: x + 1.5, y: 5.4, w: cw - 1.75, h: 0.35, fontSize: 11, color: SUNSET, align: "right", valign: "middle" });
  });
  t(s, "종목별 동작 포즈를 활용한 캐릭터 상품화", { x: MX, y: 6.15, w: 12, h: 0.45, fontSize: 14, bold: true, color: NIGHT });
}

// 19 엔딩: 작품의 마지막처럼 검은 화면에 흰 자막
{
  const s = pres.addSlide({ masterName: "BLACK", sectionTitle: "결과물 · IP" });
  t(s, "날개가 없어도, 마음만 있다면,\n우리는 언제든 하늘을 날 수 있다.", { x: 1.2, y: 1.6, w: 10.9, h: 1.9, fontFace: SERIF, fontSize: 30, color: PAPER, align: "center", valign: "middle", lineSpacingMultiple: 1.4 });
  t(s, "「상상 속의 하늘을 날다」 SCENE 6 엔딩 자막", { x: 1.2, y: 3.55, w: 10.9, h: 0.3, fontSize: 11, color: HAZE, align: "center" });
  const skills = [["서정 연출", "빛과 시간대를 활용한 감정 흐름 설계"], ["AI 파이프라인", "스틸컷 생성부터 영상화 · 편집까지 일괄 제작"], ["캐릭터 IP", "이모티콘 · 굿즈로의 캐릭터 확장"]];
  skills.forEach(([k, v], i) => {
    const x = 1.2 + i * 3.7;
    t(s, k, { x, y: 4.35, w: 3.4, h: 0.35, fontSize: 13, bold: true, color: GOLD, align: "center" });
    t(s, v, { x, y: 4.72, w: 3.4, h: 0.55, fontSize: 11, color: PAPER, align: "center" });
  });
  t(s, "강원", { x: 1.2, y: 5.75, w: 2, h: 0.5, fontFace: SERIF, fontSize: 20, bold: true, color: PAPER, valign: "middle" });
  [["이메일", 3.4], ["연락처", 6.3], ["포트폴리오 링크", 9.2]].forEach(([k, x]) => {
    t(s, k, { x, y: 5.68, w: 2.6, h: 0.28, fontSize: 10, color: HAZE });
    s.addShape(pres.shapes.LINE, { x, y: 6.35, w: 2.6, h: 0, line: { color: LINE, width: 1 }, objectName: nm("blank-line") });
  });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();

// 이민상 포트폴리오 — 곰나나 「점심메뉴 대소동」
// node build.js  →  ../이민상_포트폴리오.pptx
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const IMG = path.join(__dirname, "../work/img");
const DIMS = require("../work/dims.json");
const OUT = path.join(__dirname, "../이민상_포트폴리오.pptx");

const LINK_LUNCH = "https://drive.google.com/file/d/1kwQ9-Hc1DLFkLV5Mx-d7NjC70e1IsnNo/view";
const LINK_AD = "https://drive.google.com/file/d/1QdZ01a6mN7L7mlONQdThAWSbANCpgrNe/view";

const FONT = "Malgun Gothic";
const THEME = {
  name: "Gomnana",
  headFontFace: FONT,
  bodyFontFace: FONT,
  colors: {
    dk1: "1F2A37", // 잉크
    lt1: "FFFFFF",
    dk2: "5B6B7C", // 보조 텍스트
    lt2: "F5FAFD", // 카드
    accent1: "4FA3D9", // 진하늘
    accent2: "7CC36B", // 초원
    accent3: "F49AB8", // 분홍
    accent4: "E8B931", // 금색
    accent5: "BFE3F7", // 하늘
    accent6: "2E7FB8", // 진하늘(글자용)
    hlink: "2E7FB8",
    folHlink: "2E7FB8",
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
pres.author = "이민상";
pres.title = "곰나나 「점심메뉴 대소동」 이민상 포트폴리오";

const C = pres.SchemeColor;
const INK = C.text1, MUTED = C.text2, WHITE = C.background1, CARD = C.background2;
const BLUE = C.accent1, GREEN = C.accent2, PINK = C.accent3, GOLD = C.accent4, SKY = C.accent5, BLUE_T = C.accent6;
const W = 13.333, H = 7.5, MX = 0.6;

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: WHITE },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 1.45, y: 0.4, w: 11.3, h: 0.66, fontFace: FONT, fontSize: 28, bold: true, color: INK, align: "left", valign: "middle", margin: 0 }, text: "" } },
    { text: { text: "이민상  ·  곰나나 「점심메뉴 대소동」", options: { x: MX, y: 7.05, w: 6, h: 0.3, fontFace: FONT, fontSize: 9, color: MUTED, margin: 0 } } },
  ],
  slideNumber: { x: 12.23, y: 7.05, w: 0.5, h: 0.3, fontFace: FONT, fontSize: 9, color: MUTED, align: "right" },
});
pres.defineSlideMaster({
  title: "DARK",
  background: { color: INK },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 1.45, y: 0.4, w: 11.3, h: 0.66, fontFace: FONT, fontSize: 28, bold: true, color: WHITE, align: "left", valign: "middle", margin: 0 }, text: "" } },
  ],
});

// ---------- helpers ----------
let shapeId = 0;
const nm = (s) => `${s}-${++shapeId}`;

function fit(key, x, y, w, h) {
  const [iw, ih] = DIMS[key];
  const r = iw / ih;
  let fw = w, fh = w / r;
  if (fh > h) { fh = h; fw = h * r; }
  return { x: x + (w - fw) / 2, y: y + (h - fh) / 2, w: fw, h: fh };
}
function img(slide, key, x, y, w, h, alt, url) {
  const ext = key === "logo" ? "png" : "jpg";
  const f = fit(key, x, y, w, h);
  const link = url ? { hyperlink: { url, tooltip: "영상 보기" } } : {};
  slide.addImage({ path: path.join(IMG, `${key}.${ext}`), ...f, altText: alt || key, objectName: nm(`img-${key}`), ...link });
  return f;
}
function card(slide, x, y, w, h, fill, opts = {}) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: opts.r ?? 0.14, fill: { color: fill },
    line: opts.line ? { color: opts.line, width: opts.lineW ?? 1 } : { type: "none" },
    objectName: nm(opts.name || "card"),
  });
}
// 흰 배경 이미지용 액자: 연한 카드 + 패딩 안에 이미지
function framedImg(slide, key, x, y, w, h, opts = {}) {
  const pad = opts.pad ?? 0.1;
  const f = fit(key, x + pad, y + pad, w - 2 * pad, h - 2 * pad);
  card(slide, f.x - pad, f.y - pad, f.w + 2 * pad, f.h + 2 * pad, opts.fill || SKY, { r: opts.r ?? 0.12, name: "frame" });
  img(slide, key, f.x, f.y, f.w, f.h, opts.alt, opts.url);
  return { x: f.x - pad, y: f.y - pad, w: f.w + 2 * pad, h: f.h + 2 * pad };
}
function text(slide, t, o) {
  slide.addText(t, { fontFace: FONT, color: INK, margin: 0, valign: "top", isTextBox: true, objectName: nm(o.name || "text"), ...o });
}
function chip(slide, t, x, y, w, h, fill, color, size = 11, bold = true) {
  slide.addText(t, {
    shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: h / 2, x, y, w, h,
    fill: { color: fill }, color, fontFace: FONT, fontSize: size, bold, align: "center", valign: "middle", margin: 0,
    objectName: nm("chip"),
  });
}
function addContent(num, title, section, master = "CONTENT") {
  const s = pres.addSlide({ masterName: master, sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  chip(s, String(num).padStart(2, "0"), MX, 0.5, 0.7, 0.46, master === "DARK" ? BLUE : BLUE, WHITE, 15);
  return s;
}
function linkButton(slide, label, url, x, y, w, h) {
  slide.addText([{ text: label, options: { hyperlink: { url, tooltip: url }, color: WHITE, bold: true } }], {
    shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: h / 2, x, y, w, h, fill: { color: INK },
    fontFace: FONT, fontSize: 14, align: "center", valign: "middle", margin: 0, objectName: nm("link-button"),
  });
}
function placeholderFrame(slide, x, y, w, h, label) {
  slide.addText(label, {
    shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, x, y, w, h,
    fill: { color: CARD }, line: { color: BLUE, width: 1.25, dashType: "dash" },
    fontFace: FONT, fontSize: 12, color: MUTED, align: "center", valign: "middle", margin: 0.1, objectName: nm("video-placeholder"),
  });
}

// =====================================================================
// 01 표지
pres.addSection({ title: "소개" });
{
  const s = pres.addSlide({ masterName: "DARK", sectionTitle: "소개" });
  img(s, "logo", MX, 0.75, 3.3, 2.2, "곰나나 로고");
  text(s, "AI 애니메이션 · 캐릭터 IP 포트폴리오 2026", { x: MX, y: 3.15, w: 5.2, h: 0.35, fontSize: 13, color: SKY, bold: true });
  text(s, "곰나나", { x: MX, y: 3.6, w: 5.6, h: 0.45, fontSize: 18, bold: true, color: WHITE });
  text(s, "「점심메뉴 대소동」", { x: MX, y: 4.05, w: 5.6, h: 0.8, fontSize: 36, bold: true, color: WHITE, valign: "middle" });
  text(s, "이민상", { x: MX, y: 5.1, w: 5, h: 0.55, fontSize: 24, bold: true, color: WHITE });
  text(s, "AI를 도구로 쓰는 1인 풀파이프라인 디렉터", { x: MX, y: 5.7, w: 5.4, h: 0.4, fontSize: 15, color: SKY });
  const f = framedImg(s, "TEST", 6.35, 1.35, 6.4, 3.75, { fill: SKY, pad: 0.12, alt: "키 비주얼: 콩가리니쿠를 들고 달려오는 지니와 뒤따르는 미샤" });
  text(s, "키 비주얼: 콩가리니쿠를 든 지니와 뒤따르는 미샤", { x: f.x, y: f.y + f.h + 0.15, w: f.w, h: 0.3, fontSize: 10, color: SKY });
}

// 02 한 줄 소개
{
  const s = addContent(2, "한 줄 소개", "소개");
  const k = framedImg(s, "0202", MX, 1.4, 7.0, 4.4, { alt: "컷2 국밥 선언: 콩가리니쿠를 들어 올린 지니" });
  text(s, "컷2 · 씬2-B  국밥 선언", { x: MX, y: k.y + k.h + 0.1, w: 7, h: 0.3, fontSize: 10, color: MUTED });
  card(s, 8.0, 1.4, 4.73, k.h, CARD);
  text(s, "LOGLINE", { x: 8.35, y: 1.7, w: 4, h: 0.3, fontSize: 11, bold: true, color: BLUE_T, charSpacing: 2 });
  text(s, "우유부단한 곰 지니가 점심 메뉴를 정하려 하지만, 짝사랑 상대 고양이 오대오의 등장으로 정신이 팔려 결국 아무것도 정하지 못한 채 끝나는 에피소드", {
    x: 8.35, y: 2.1, w: 4.05, h: 2.6, fontSize: 17, bold: true, color: INK, lineSpacingMultiple: 1.25,
  });
  const tags = [["일상 코미디", SKY, INK], ["59초 숏폼", PINK, INK], ["캐릭터 IP", GREEN, INK]];
  tags.forEach(([t, f, c], i) => chip(s, t, 8.35 + i * 1.38, 1.4 + k.h - 0.7, 1.28, 0.38, f, c, 11));
  text(s, "기획부터 IP 확장까지 전 과정 1인 제작", { x: MX, y: 6.35, w: 12.1, h: 0.4, fontSize: 15, bold: true, color: INK });
}

// 03 작품 기본 정보
{
  const s = addContent(3, "작품 기본 정보", "소개");
  const stats = [
    ["장르", "일상 + 코믹", "Slice of Life / Comedy"],
    ["플랫폼", "OTT", "유튜브 외"],
    ["러닝타임", "59초", "8컷 / 13씬"],
    ["화면비", "16:9", "가로형"],
  ];
  const cw = 2.86, gap = 0.22;
  stats.forEach(([k, v, sub], i) => {
    const x = MX + i * (cw + gap);
    card(s, x, 1.4, cw, 1.95, CARD);
    text(s, k, { x: x + 0.3, y: 1.6, w: cw - 0.6, h: 0.3, fontSize: 12, bold: true, color: BLUE_T });
    text(s, v, { x: x + 0.3, y: 1.95, w: cw - 0.6, h: 0.75, fontSize: 30, bold: true, color: INK, valign: "middle" });
    text(s, sub, { x: x + 0.3, y: 2.75, w: cw - 0.6, h: 0.35, fontSize: 12, color: MUTED });
  });
  const ty = 3.65, th = 3.15, tw = 3.75;
  [["주 타겟", "10대 후반~30대 초반 MZ세대", "릴스 · 쇼츠 주 소비층", BLUE],
   ["부가 타겟", "귀여운 캐릭터 콘텐츠를 좋아하는 전 연령대", "캐릭터 굿즈 · 이모티콘 구매층과 겹침", PINK]].forEach(([k, v, sub, col], i) => {
    const x = MX + i * (tw + gap);
    card(s, x, ty, tw, th, CARD);
    chip(s, k, x + 0.3, ty + 0.3, 1.2, 0.38, col, i === 0 ? WHITE : INK, 11);
    text(s, v, { x: x + 0.3, y: ty + 0.9, w: tw - 0.6, h: 1.1, fontSize: 16, bold: true, color: INK });
    text(s, sub, { x: x + 0.3, y: ty + 2.2, w: tw - 0.6, h: 0.7, fontSize: 13, color: MUTED });
  });
  const ix = MX + 2 * (tw + gap);
  framedImg(s, "0102", ix, ty, W - MX - ix, th, { alt: "언덕에 나란히 앉은 지니와 미샤" });
}

// 04 기획의도
{
  const s = addContent(4, "기획의도", "소개");
  text(s, "산동네 곰 캐릭터들의 일상 코미디를 통해\n시청자 공감과 캐릭터 몰입을 유도하는\n숏폼 시리즈", { x: MX, y: 1.45, w: 5.9, h: 1.9, fontSize: 24, bold: true, color: INK, lineSpacingMultiple: 1.15 });
  const pts = [
    ["공감", "점심 메뉴 고민이라는 일상 소재 활용", SKY],
    ["캐릭터 몰입", "행동과 소품을 통한 캐릭터 성격 표현", PINK],
    ["시리즈성", "열린 결말을 통한 다음 에피소드 연결", GREEN],
  ];
  pts.forEach(([k, v, col], i) => {
    const y = 3.65 + i * 1.02;
    card(s, MX, y, 5.9, 0.85, CARD);
    chip(s, k, MX + 0.25, y + 0.22, 1.45, 0.42, col, INK, 12);
    text(s, v, { x: MX + 1.95, y: y + 0.12, w: 3.8, h: 0.62, fontSize: 14, color: INK, valign: "middle" });
  });
  const f = framedImg(s, "0802", 6.9, 1.45, 5.83, 4.9, { alt: "뒷모습 와이드: 초원을 바라보는 지니와 미샤" });
  text(s, "컷8 · 씬8-B  오픈 엔딩", { x: f.x, y: f.y + f.h + 0.1, w: 5, h: 0.3, fontSize: 10, color: MUTED });
}

// 05 세계관 · 무대
pres.addSection({ title: "세계관 · 캐릭터" });
{
  const s = addContent(5, "세계관 · 무대: 고고산 언덕", "세계관 · 캐릭터");
  const f = framedImg(s, "Background", MX, 1.4, 7.6, 4.6, { alt: "배경 시트: 파란 하늘과 뭉게구름, 초원" });
  text(s, "배경 시트", { x: f.x, y: f.y + f.h + 0.1, w: 3, h: 0.3, fontSize: 10, color: MUTED });
  card(s, 8.5, 1.4, 4.23, 4.6, CARD);
  const rows = [["장소", "고고산 언덕 · 산동네 공터"], ["시간", "전 컷 정오"], ["하늘", "선명한 파란 하늘 · 뭉게구름"], ["땅", "자연광이 도는 초록 초원"], ["캐릭터", "순백(#FFFFFF), 하늘/초원과의 대비로 가독성 확보"]];
  rows.forEach(([k, v], i) => {
    const y = 1.75 + i * 0.8;
    text(s, k, { x: 8.8, y, w: 1.0, h: 0.6, fontSize: 12, bold: true, color: BLUE_T });
    text(s, v, { x: 9.8, y, w: 2.7, h: 0.66, fontSize: 13, color: INK });
  });
  const sw = [["FFFFFF", "캐릭터"], ["4FA3D9", "하늘"], ["FFFFFF", "구름"], ["7CC36B", "초원"]];
  sw.forEach(([hex, lab], i) => {
    const x = MX + i * 1.55;
    s.addShape(pres.shapes.OVAL, { x, y: 6.45, w: 0.36, h: 0.36, fill: { color: hex }, line: { color: "BFE3F7", width: 1 }, objectName: nm("swatch") });
    text(s, lab, { x: x + 0.45, y: 6.48, w: 1.0, h: 0.3, fontSize: 11, color: INK, valign: "middle" });
  });
}

// 06 캐릭터 소개
{
  const s = addContent(6, "캐릭터 소개", "세계관 · 캐릭터");
  const chars = [
    ["JIni", "지니", "주동인물 · 흰 곰", BLUE, WHITE, ["까칠·힙한 텐션, 결국 모든 상황에 참여", "한식파 (국밥, 순대국)", "결정 앞에선 흔들림, 실수엔 당당", "오른손엔 늘 콩가리니쿠 (먹지 않음)", "오대오를 일방적으로 좋아함"]],
    ["Misha", "미샤", "조력자 · 흰 곰", PINK, INK, ["상냥하고 호기심 많음", "지니 의견을 대체로 따름", "자기 취향 앞에선 우물쭈물", "최애 음식: 햄버거", "소심해 보여도 계획적 · 적극적"]],
    ["Odeo", "오대오", "카메오 · 고양이", GOLD, INK, ["시크함의 끝판왕", "건드려도 목만 돌려 쳐다보는 무반응", "분홍 목걸이 + 금색 방울", "맥락상 필요할 때만 등장"]],
  ];
  const cw = 3.9, gap = 0.215;
  chars.forEach(([key, name, role, col, tcol, lines], i) => {
    const x = MX + i * (cw + gap);
    card(s, x, 1.35, cw, 5.5, CARD);
    framedImg(s, key, x + 0.15, 1.5, cw - 0.3, 2.15, { fill: SKY, pad: 0.06, alt: `${name} 캐릭터 시트` });
    text(s, name, { x: x + 0.3, y: 3.92, w: 1.5, h: 0.5, fontSize: 22, bold: true, color: INK, valign: "middle" });
    chip(s, role, x + cw - 1.95, 3.98, 1.65, 0.38, col, tcol, 10);
    s.addText(lines.map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < lines.length - 1 } })), {
      x: x + 0.3, y: 4.6, w: cw - 0.55, h: 2.1, fontFace: FONT, fontSize: 13, color: INK, valign: "top", margin: 0, paraSpaceAfter: 5, isTextBox: true, objectName: nm("traits"),
    });
  });
}

// 07 시그니처 소품 콩가리니쿠
{
  const s = addContent(7, "시그니처 소품: 콩가리니쿠", "세계관 · 캐릭터");
  text(s, "소품의 위치와 각도 변화를 통한 캐릭터 감정 표현", { x: MX, y: 1.3, w: 12, h: 0.4, fontSize: 16, color: INK });
  const steps = [
    ["0201", "늘어뜨림", "컷1~2-A · 배고픔, 결의 직전", SKY],
    ["0202", "들어 올림", "컷2-B · 국밥 선언의 확신", GREEN],
    ["0401", "힘없이 기울어짐", "컷4 · 확신이 흔들림", PINK],
  ];
  const iw = 3.75, gap = 0.44;
  steps.forEach(([key, k, v, col], i) => {
    const x = MX + i * (iw + gap);
    const f = framedImg(s, key, x, 1.95, iw, 2.2, { alt: `콩가리니쿠 ${k}` });
    chip(s, k, x, f.y + f.h + 0.2, 1.9, 0.42, col, INK, 13);
    text(s, v, { x, y: f.y + f.h + 0.72, w: iw, h: 0.35, fontSize: 12, color: MUTED });
    if (i < 2) text(s, "→", { x: x + iw + 0.02, y: 2.75, w: 0.4, h: 0.5, fontSize: 24, bold: true, color: BLUE, align: "center", valign: "middle" });
  });
  card(s, MX, 5.6, 12.13, 1.2, CARD);
  text(s, "연속성 규칙", { x: MX + 0.3, y: 5.75, w: 2, h: 0.3, fontSize: 12, bold: true, color: BLUE_T });
  const rules = ["항상 오른손", "찌르는 건 왼손 손가락 (컷7-A)", "뒷모습에서도 오른손 (컷8-B)"];
  rules.forEach((t, i) => chip(s, t, MX + 0.3 + i * 3.7, 6.13, 3.45, 0.42, WHITE, INK, 12, false));
}

// 08 줄거리 3막
{
  const s = addContent(8, "줄거리: 3막 구성", "세계관 · 캐릭터");
  const acts = [
    ["0202", "S#1  배고픔", BLUE, WHITE, "지니의 국밥 선언과 미샤의 햄버거 선호 사이에서 흔들리는 메뉴 결정", "「국밥… 아니, 순대국인가…」"],
    ["0602", "S#2  오대오", PINK, INK, "오대오 등장과 지니의 광속 질주, 손가락으로 찔러도 목만 돌리는 오대오의 무반응", "「오대오다~」"],
    ["0802", "S#3  엔딩", GREEN, INK, "끝내 메뉴를 정하지 못한 채 열린 결말로 마무리", "「근데 저녁은 뭐 먹지?」"],
  ];
  const cw = 3.9, gap = 0.215;
  acts.forEach(([key, act, col, tcol, body, line], i) => {
    const x = MX + i * (cw + gap);
    card(s, x, 1.35, cw, 5.5, CARD);
    framedImg(s, key, x + 0.15, 1.5, cw - 0.3, 2.15, { fill: SKY, pad: 0.06 });
    chip(s, act, x + 0.25, 3.9, 1.85, 0.4, col, tcol, 12);
    text(s, body, { x: x + 0.3, y: 4.5, w: cw - 0.6, h: 1.4, fontSize: 14, color: INK, lineSpacingMultiple: 1.15 });
    text(s, line, { x: x + 0.3, y: 5.95, w: cw - 0.6, h: 0.65, fontSize: 15, bold: true, color: BLUE_T, valign: "middle" });
  });
}

// 09 연출 설계 ① 감정 ↔ 샷
pres.addSection({ title: "연출 · AI 설계" });
{
  const s = addContent(9, "연출 설계 ①: 감정에 맞춘 샷/앵글", "연출 · AI 설계");
  const rows = [
    ["0102", "배고픈 공허함", "하이앵글 3/4 부감\n원경으로 작게"],
    ["0202", "국밥 선언의 확신", "로우앵글\n미디엄샷"],
    ["0401", "확신이 무너지는 순간", "더치틸트\n+ 두 캐릭터 사이 여백"],
    ["0602", "오대오를 향한 폭주", "다이나믹 더치틸트\n+ 스피드라인 · 흙먼지"],
    ["0802", "허무한 결말", "뒷모습 와이드\n화면 하단에 작게"],
  ];
  const cw = 2.29, gap = 0.17;
  text(s, "감정", { x: MX, y: 3.08, w: 1, h: 0.3, fontSize: 11, bold: true, color: MUTED });
  text(s, "샷 · 앵글", { x: MX, y: 4.55, w: 1.2, h: 0.3, fontSize: 11, bold: true, color: MUTED });
  rows.forEach(([key, emo, shot], i) => {
    const x = MX + i * (cw + gap);
    framedImg(s, key, x, 1.4, cw, 1.4, { pad: 0.06 });
    chip(s, emo, x, 3.42, cw, 0.46, PINK, INK, 12);
    text(s, "↓", { x, y: 3.92, w: cw, h: 0.5, fontSize: 20, bold: true, color: BLUE, align: "center", valign: "middle" });
    card(s, x, 4.9, cw, 1.2, SKY);
    text(s, shot, { x: x + 0.1, y: 4.95, w: cw - 0.2, h: 1.1, fontSize: 13, bold: true, color: INK, align: "center", valign: "middle" });
  });
  text(s, "장면별 감정에 맞춘 샷/앵글을 프롬프트에 반영", { x: MX, y: 6.4, w: 12, h: 0.4, fontSize: 15, color: INK });
}

// 10 연출 설계 ② 반전 대비
{
  const s = addContent(10, "연출 설계 ②: 반전 대비", "연출 · AI 설계");
  const iw = 5.6;
  const L = framedImg(s, "0602", MX, 1.45, iw, 3.3, { fill: PINK, pad: 0.1, alt: "지니의 광속 질주" });
  const R = framedImg(s, "0702", W - MX - iw, 1.45, iw, 3.3, { fill: SKY, pad: 0.1, alt: "목만 돌린 오대오의 무반응" });
  s.addText("VS", { shape: pres.shapes.OVAL, x: W / 2 - 0.45, y: 2.65, w: 0.9, h: 0.9, fill: { color: INK }, color: WHITE, fontFace: FONT, fontSize: 18, bold: true, align: "center", valign: "middle", margin: 0, objectName: nm("vs") });
  const colTxt = (x, w, head, col, tcol, lines) => {
    chip(s, head, x, 5.0, 2.3, 0.46, col, tcol, 14);
    text(s, lines, { x, y: 5.6, w, h: 0.75, fontSize: 14, color: INK, lineSpacingMultiple: 1.2 });
  };
  colTxt(L.x, L.w, "과장된 행동", PINK, INK, "흙먼지와 스피드라인으로 표현한 지니의 광속 질주\n응답을 기대하지 않는 「오대오다~」");
  text(s, "과장된 행동과 무반응의 대비를 통한 웃음 유발", { x: MX, y: 6.5, w: 12.1, h: 0.4, fontSize: 15, bold: true, color: BLUE_T });
  colTxt(R.x, R.w, "무반응", SKY, INK, "몸은 그대로 두고 목만 돌리는 오대오의 반응\n무표정 클로즈업으로 마무리");
}

// 11 AI 프롬프트 설계
{
  const s = addContent(11, "AI 프롬프트 설계: 공통 사양과 컷 기술 규칙", "연출 · AI 설계");
  const lx = MX, lw = 6.6;
  card(s, lx, 1.35, lw, 5.45, CARD);
  text(s, "공통 비주얼 사양", { x: lx + 0.3, y: 1.55, w: 4, h: 0.4, fontSize: 16, bold: true, color: INK });
  text(s, "모든 컷 프롬프트에 공통 적용", { x: lx + 3.6, y: 1.6, w: 2.75, h: 0.3, fontSize: 11, color: MUTED, align: "right" });
  const spec = [
    ["화면비", "16:9"],
    ["아트 스타일", "소프트 애니메 셀셰이딩 (실사풍/과한 반사광 금지)"],
    ["색감", "순백 캐릭터 #FFFFFF · 파란 하늘 · 뭉게구름 · 초록 초원"],
    ["시간대", "전 컷 정오"],
    ["장소", "고고산 언덕 (공터)"],
    ["BGM", "경쾌하고 발랄한 CM송 느낌"],
    ["내레이션", "없음"],
  ];
  spec.forEach(([k, v], i) => {
    const y = 2.1 + i * 0.64;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: lx + 0.25, y, w: lw - 0.5, h: 0.54, rectRadius: 0.1, fill: { color: WHITE }, line: { type: "none" }, objectName: nm("spec-row") });
    text(s, k, { x: lx + 0.45, y, w: 1.4, h: 0.54, fontSize: 12, bold: true, color: BLUE_T, valign: "middle" });
    text(s, v, { x: lx + 1.9, y, w: lw - 2.3, h: 0.54, fontSize: 12, color: INK, valign: "middle" });
  });
  const rx = lx + lw + 0.3, rw = W - MX - rx;
  card(s, rx, 1.35, rw, 3.55, SKY);
  text(s, "컷 기술 5항목 규칙", { x: rx + 0.3, y: 1.55, w: rw - 0.6, h: 0.4, fontSize: 16, bold: true, color: INK });
  const five = [["샷", "사이즈 + 앵글"], ["등장인물", "누가 · 어디에 · 어떤 자세"], ["표정·감정", "시선 · 감정 상태"], ["소품·이펙트", "콩가리니쿠 쥔 손 · 흙먼지 · 스피드라인"], ["대사", "화자 + 대사 + 톤, 없으면 '없음'"]];
  five.forEach(([k, v], i) => {
    const y = 2.1 + i * 0.54;
    s.addText(String(i + 1), { shape: pres.shapes.OVAL, x: rx + 0.3, y: y + 0.04, w: 0.36, h: 0.36, fill: { color: BLUE }, color: WHITE, fontFace: FONT, fontSize: 11, bold: true, align: "center", valign: "middle", margin: 0, objectName: nm("num") });
    text(s, k, { x: rx + 0.8, y, w: 1.35, h: 0.44, fontSize: 12, bold: true, color: INK, valign: "middle" });
    text(s, v, { x: rx + 2.15, y, w: rw - 2.35, h: 0.44, fontSize: 11, color: INK, valign: "middle" });
  });
  const rules = [["정지 이미지 단계의 카메라 무브 배제", "움직임은 영상화 단계에서 별도 지정", GOLD], ["심리 묘사의 시각적 구체화", "예: 눈 흔들림, 힘없이 기울어진 고기", PINK]];
  rules.forEach(([h1, h2, col], i) => {
    const y = 5.05 + i * 0.9;
    card(s, rx, y, rw, 0.78, CARD);
    s.addShape(pres.shapes.OVAL, { x: rx + 0.25, y: y + 0.27, w: 0.24, h: 0.24, fill: { color: col }, line: { type: "none" }, objectName: nm("dot") });
    text(s, h1, { x: rx + 0.65, y: y + 0.08, w: rw - 0.85, h: 0.34, fontSize: 13, bold: true, color: INK, valign: "middle" });
    text(s, h2, { x: rx + 0.65, y: y + 0.42, w: rw - 0.85, h: 0.3, fontSize: 11, color: MUTED, valign: "middle" });
  });
}

// 12~14 스토리보드
const SB = {
  "0102": ["컷1 · 씬1-A · 3초", "하이앵글 3/4 부감, 미디엄 와이드", "대사 없음"],
  "0201": ["컷2 · 씬2-A · 5초", "로우앵글 미디엄샷", "대사 없음"],
  "0202": ["컷2 · 씬2-B · 5초", "로우앵글 미디엄샷", "지니 「오늘은 내가 딱 정할게. 국밥 어때? 이만한 게 없지.」"],
  "0301": ["컷3 · 씬3-A · 5초", "아이레벨 클로즈업", "미샤 「어… 나는 국물보단 햄버거 쪽이 좋은데…」"],
  "0401": ["컷4 · 씬4-A · 5초", "더치틸트 미디엄샷", "지니 「국밥… 아니, 순대국인가…」"],
  "0501": ["컷5 · 씬5-A · 5초", "원경 와이드", "대사 없음"],
  "0601": ["컷6 · 씬6-A · 5초", "더치틸트 미디엄샷", "대사 없음"],
  "0602": ["컷6 · 씬6-B · 5초", "다이나믹 더치틸트", "지니 「오대오다~」"],
  "0701": ["컷7 · 씬7-A · 3초", "사이드 프로필", "대사 없음 · 왼손 손가락으로 콕"],
  "0702": ["컷7 · 씬7-B · 4초", "오대오 클로즈업", "대사 없음 · 목만 스윽"],
  "0801": ["컷8 · 씬8-A · 5초", "아이레벨 미디엄샷", "미샤 「…그래서 뭐 먹어?」"],
  "0802": ["컷8 · 씬8-B · 5초", "와이드 · 뒷모습", "지니 「근데 저녁은 뭐 먹지?」"],
};
function sbCaption(s, key, x, y, w, big = false) {
  const [a, b, c] = SB[key];
  s.addText([
    { text: a, options: { bold: true, color: BLUE_T, fontSize: big ? 13 : 11, breakLine: true } },
    { text: b, options: { color: INK, fontSize: big ? 13 : 11, breakLine: true } },
    { text: c, options: { color: MUTED, fontSize: big ? 13 : 11 } },
  ], { x, y, w, h: big ? 1.0 : 0.82, fontFace: FONT, valign: "top", margin: 0, paraSpaceAfter: 1, isTextBox: true, objectName: nm("sb-caption") });
}
function storyboardFive(num, title, keys, info) {
  const s = addContent(num, title, "연출 · AI 설계");
  const iw = 3.6, ih = 1.85, gap = 0.4, x0 = (W - (3 * iw + 2 * gap)) / 2;
  keys.forEach((key, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = x0 + col * (iw + gap), y = 1.3 + row * 2.85;
    const f = framedImg(s, key, x, y, iw, ih, { pad: 0.06 });
    sbCaption(s, key, f.x, f.y + f.h + 0.1, iw);
  });
  const x = x0 + 2 * (iw + gap), y = 1.3 + 2.85;
  card(s, x, y, iw, 2.6, CARD);
  chip(s, info[0], x + 0.25, y + 0.25, 1.9, 0.4, info[2], info[3], 12);
  text(s, info[1], { x: x + 0.25, y: y + 0.85, w: iw - 0.5, h: 1.6, fontSize: 13, color: INK, lineSpacingMultiple: 1.2 });
  return s;
}
storyboardFive(12, "스토리보드 ①: S#1 배고픔", ["0102", "0201", "0202", "0301", "0401"], ["S#1 · 4컷 · 27초", "대사 중심의 갈등 전개", BLUE, WHITE]);
storyboardFive(13, "스토리보드 ②: S#2 오대오", ["0501", "0601", "0602", "0701", "0702"], ["S#2 · 3컷 · 22초", "대사 최소화, 동작 중심의 전개", PINK, INK]);
{
  const s = addContent(14, "스토리보드 ③: S#3 엔딩", "연출 · AI 설계");
  const iw = 5.9, gap = 0.33;
  ["0801", "0802"].forEach((key, i) => {
    const x = MX + i * (iw + gap);
    const f = framedImg(s, key, x, 1.35, iw, 3.4, { pad: 0.08 });
    sbCaption(s, key, f.x, f.y + f.h + 0.15, iw, true);
  });
  card(s, MX, 5.95, 12.13, 0.85, CARD);
  chip(s, "S#3 · 1컷 · 10초", MX + 0.25, 6.17, 1.9, 0.4, GREEN, INK, 12);
  text(s, "열린 결말을 통한 시리즈 연속성 확보", { x: MX + 2.45, y: 6.1, w: 9.4, h: 0.55, fontSize: 14, color: INK, valign: "middle" });
}

// 15 리듬 설계
{
  const s = addContent(15, "리듬 설계: 13씬 59초 타임라인", "연출 · AI 설계");
  const scenes = [["1-A", 3, 1], ["1-B", 4, 1], ["2-A", 5, 1], ["2-B", 5, 1], ["3-A", 5, 1], ["4-A", 5, 1], ["5-A", 5, 2], ["6-A", 5, 2], ["6-B", 5, 2], ["7-A", 3, 2], ["7-B", 4, 2], ["8-A", 5, 3], ["8-B", 5, 3]];
  const cuts = [["컷1", 7], ["컷2", 10], ["컷3", 5], ["컷4", 5], ["컷5", 5], ["컷6", 10], ["컷7", 7], ["컷8", 10]];
  const acts = [["S#1 배고픔 · 27초", 27, BLUE, WHITE], ["S#2 오대오 · 22초", 22, PINK, INK], ["S#3 엔딩 · 10초", 10, GREEN, INK]];
  const total = 59, x0 = MX, tw = 12.13, u = tw / total;
  let x = x0;
  acts.forEach(([t, d, col, tc]) => { chip(s, t, x + 0.03, 1.45, d * u - 0.06, 0.44, col, tc, 12); x += d * u; });
  x = x0;
  const actFill = { 1: SKY, 2: "FBD3E1", 3: "CDE9C4" };
  scenes.forEach(([id, d, a]) => {
    const w = d * u;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.03, y: 2.1, w: w - 0.06, h: 1.25, rectRadius: 0.08, fill: { color: actFill[a] }, line: { type: "none" }, objectName: nm("scene-bar") });
    text(s, [{ text: id, options: { bold: true, fontSize: 12, breakLine: true } }, { text: `${d}초`, options: { fontSize: 11 } }], { x: x + 0.03, y: 2.1, w: w - 0.06, h: 1.25, align: "center", valign: "middle", color: INK });
    x += w;
  });
  x = x0;
  cuts.forEach(([t, d]) => {
    const w = d * u;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.03, y: 3.5, w: w - 0.06, h: 0.42, rectRadius: 0.08, fill: { color: CARD }, line: { color: SKY, width: 1 }, objectName: nm("cut-bar") });
    text(s, `${t} · ${d}초`, { x: x + 0.03, y: 3.5, w: w - 0.06, h: 0.42, fontSize: 11, align: "center", valign: "middle", color: INK });
    x += w;
  });
  text(s, "0초", { x: x0, y: 4.0, w: 1, h: 0.3, fontSize: 10, color: MUTED });
  text(s, "59초", { x: x0 + tw - 1, y: 4.0, w: 1, h: 0.3, fontSize: 10, color: MUTED, align: "right" });
  const notes = [["3~4초", "오프닝 (1-A · 1-B)\n무반응 (7-A · 7-B)"], ["5초", "대사 · 액션 씬의\n기본 단위"], ["59초", "13씬 합계\n1분 이내 숏폼"]];
  const cw = 3.9, gap = 0.215;
  notes.forEach(([big, small], i) => {
    const nx = MX + i * (cw + gap);
    card(s, nx, 4.65, cw, 2.1, CARD);
    text(s, big, { x: nx + 0.3, y: 4.85, w: cw - 0.6, h: 0.75, fontSize: 30, bold: true, color: BLUE_T, valign: "middle" });
    text(s, small, { x: nx + 0.3, y: 5.7, w: cw - 0.6, h: 0.85, fontSize: 13, color: INK });
  });
}

// 16 제작 공정 6단계
{
  const s = addContent(16, "제작 공정 6단계", "연출 · AI 설계");
  const steps = [
    ["기획 · 시나리오", "작품개요 · 8컷 13씬\n컷 기술 5항목"],
    ["캐릭터 · 배경 시트", "턴어라운드 · 표정\n고고산 배경"],
    ["스틸컷 생성", "씬별 정지 이미지\n카메라 무브 없음"],
    ["영상화", "스틸컷에\n움직임 지정"],
    ["사운드", "경쾌한\nCM송 느낌 BGM"],
    ["편집", "59초\n타임라인"],
  ];
  const cw = 1.86, gap = 0.194;
  steps.forEach(([t, d], i) => {
    const x = MX + i * (cw + gap);
    card(s, x, 1.4, cw, 2.75, CARD);
    s.addText(String(i + 1), { shape: pres.shapes.OVAL, x: x + 0.2, y: 1.6, w: 0.5, h: 0.5, fill: { color: BLUE }, color: WHITE, fontFace: FONT, fontSize: 15, bold: true, align: "center", valign: "middle", margin: 0, objectName: nm("step-num") });
    text(s, t, { x: x + 0.2, y: 2.25, w: cw - 0.35, h: 0.65, fontSize: 13, bold: true, color: INK });
    text(s, d, { x: x + 0.2, y: 2.9, w: cw - 0.35, h: 0.7, fontSize: 11, color: MUTED });
    text(s, "[툴 기입]", { x: x + 0.2, y: 3.7, w: cw - 0.35, h: 0.3, fontSize: 10, color: BLUE_T });
  });
  const iw = 3.75, g2 = 0.44, y = 4.5;
  const a = framedImg(s, "JIni", MX, y, iw, 2.1, { pad: 0.06, alt: "지니 캐릭터 시트" });
  text(s, "2  캐릭터 시트", { x: a.x, y: a.y + a.h + 0.05, w: iw, h: 0.28, fontSize: 10, color: MUTED });
  text(s, "→", { x: MX + iw, y: y + 0.8, w: g2, h: 0.5, fontSize: 22, bold: true, color: BLUE, align: "center", valign: "middle" });
  const b = framedImg(s, "0202", MX + iw + g2, y, iw, 2.1, { pad: 0.06, alt: "스틸컷 0202" });
  text(s, "3  스틸컷", { x: b.x, y: b.y + b.h + 0.05, w: iw, h: 0.28, fontSize: 10, color: MUTED });
  text(s, "→", { x: MX + 2 * iw + g2, y: y + 0.8, w: g2, h: 0.5, fontSize: 22, bold: true, color: BLUE, align: "center", valign: "middle" });
  const c = framedImg(s, "poster_lunch", MX + 2 * (iw + g2), y, iw, 2.1, { pad: 0.06, alt: "점심메뉴 영상 대표 이미지 (클릭 시 영상 재생)", url: LINK_LUNCH });
  text(s, "4  영상화 (제작 진행 중, 클릭 시 영상 재생)", { x: c.x, y: c.y + c.h + 0.05, w: iw, h: 0.28, fontSize: 10, color: MUTED });
}

// 17 디렉터의 결정 vs AI의 생성
{
  const s = addContent(17, "직접 기획/연출과 AI툴 활용", "연출 · AI 설계");
  const lw = 6.5;
  card(s, MX, 1.35, lw, 5.45, INK);
  text(s, "직접 기획/연출", { x: MX + 0.35, y: 1.6, w: 4, h: 0.5, fontSize: 20, bold: true, color: WHITE });
  const mine = [["캐릭터 설정", "성격 · 취향 · 소품 · 관계"], ["컷 구성", "8컷 / 13씬"], ["샷 · 앵글", "감정별 샷/앵글 설계"], ["대사", "화자 · 톤까지 지정"], ["씬 길이", "3~5초 단위, 합계 59초"], ["비주얼 사양", "스타일 · 색감 · 시간대 고정"]];
  mine.forEach(([k, v], i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = MX + 0.35 + col * 3.0, y = 2.35 + row * 1.42;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 2.8, h: 1.2, rectRadius: 0.12, fill: { color: "2C3A4B" }, line: { type: "none" }, objectName: nm("decision") });
    text(s, k, { x: x + 0.2, y: y + 0.15, w: 2.4, h: 0.4, fontSize: 15, bold: true, color: WHITE });
    text(s, v, { x: x + 0.2, y: y + 0.6, w: 2.45, h: 0.45, fontSize: 11, color: SKY });
  });
  const rx = MX + lw + 0.3, rw = W - MX - rx;
  card(s, rx, 1.35, rw, 5.45, CARD);
  text(s, "AI툴 활용", { x: rx + 0.35, y: 1.6, w: 4, h: 0.5, fontSize: 20, bold: true, color: INK });
  [["이미지", "캐릭터 시트 · 배경 · 스틸컷"], ["영상 소스", "스틸컷 기반 영상화"]].forEach(([k, v], i) => {
    const y = 2.35 + i * 0.75;
    chip(s, k, rx + 0.35, y, 1.3, 0.42, SKY, INK, 12);
    text(s, v, { x: rx + 1.85, y, w: rw - 2.1, h: 0.42, fontSize: 13, color: INK, valign: "middle" });
  });
  const iw = (rw - 0.7 - 0.2) / 2;
  framedImg(s, "Misha", rx + 0.35, 4.0, iw, 1.5, { pad: 0.05, alt: "미샤 캐릭터 시트" });
  framedImg(s, "0602", rx + 0.35 + iw + 0.2, 4.0, iw, 1.5, { pad: 0.05, alt: "스틸컷 0602" });
  text(s, "AI툴을 사용하여 창작/기획/연출 역량 극대화", { x: rx + 0.35, y: 5.95, w: rw - 0.7, h: 0.5, fontSize: 13, bold: true, color: BLUE_T, valign: "middle" });
}

// 18 영상 결과물 (제작 진행 중)
pres.addSection({ title: "결과물 · IP 확장" });
{
  const s = addContent(18, "영상 결과물: 점심메뉴 대소동", "결과물 · IP 확장");
  const p = framedImg(s, "poster_lunch", MX, 1.4, 7.6, 4.275, { pad: 0.08, alt: "점심메뉴 영상 대표 이미지 (클릭 시 영상 재생)", url: LINK_LUNCH });
  const rx = 8.55, rw = W - MX - rx;
  card(s, rx, 1.4, rw, 4.275, CARD);
  chip(s, "제작 진행 중", rx + 0.3, 1.7, 1.75, 0.42, GOLD, INK, 12);
  const info = [["목표 러닝타임", "59초 · 16:9"], ["구성", "8컷 / 13씬"], ["현재 영상 길이", "[영상 확인 후 기입]"], ["해상도", "[영상 확인 후 기입]"]];
  info.forEach(([k, v], i) => {
    const y = 2.4 + i * 0.75;
    text(s, k, { x: rx + 0.3, y, w: rw - 0.6, h: 0.3, fontSize: 11, bold: true, color: BLUE_T });
    text(s, v, { x: rx + 0.3, y: y + 0.3, w: rw - 0.6, h: 0.36, fontSize: 14, color: INK });
  });
  text(s, "이미지 클릭 시 영상 재생 (Google Drive) / 대표 이미지: 스틸컷 0601", { x: p.x, y: p.y + p.h + 0.12, w: 7.6, h: 0.3, fontSize: 10, color: MUTED });
  linkButton(s, "▶  영상 보기", LINK_LUNCH, rx, 6.0, rw, 0.6);
}

// 19 IP 확장 ① 브랜드 로고
{
  const s = addContent(19, "IP 확장 ①: 브랜드 로고", "결과물 · IP 확장", "DARK");
  card(s, MX, 1.4, 7.6, 5.35, BLUE, { r: 0.2 });
  img(s, "logo", MX + 0.4, 1.7, 6.8, 4.75, "곰나나 로고: 곰내 나는 나라");
  const rx = 8.6;
  text(s, "곰나나", { x: rx, y: 1.6, w: 4.1, h: 0.75, fontSize: 32, bold: true, color: WHITE });
  text(s, "곰내 나는 나라", { x: rx, y: 2.35, w: 4.1, h: 0.45, fontSize: 18, color: SKY });
  const pts = [["흰 글자 + 초록 외곽선", "순백 캐릭터, 초원 색감과의 통일감"], ["꽃 · 나뭇잎 장식", "고고산 언덕의 자연 이미지 반영"], ["나무 현판 서브 타이틀", "산동네의 소박한 분위기 표현"]];
  pts.forEach(([a, b], i) => {
    const y = 3.25 + i * 1.12;
    s.addShape(pres.shapes.OVAL, { x: rx, y: y + 0.1, w: 0.22, h: 0.22, fill: { color: [GREEN, PINK, GOLD][i] }, line: { type: "none" }, objectName: nm("dot") });
    text(s, a, { x: rx + 0.4, y, w: 3.7, h: 0.42, fontSize: 15, bold: true, color: WHITE, valign: "middle" });
    text(s, b, { x: rx + 0.4, y: y + 0.45, w: 3.7, h: 0.4, fontSize: 12, color: SKY, valign: "middle" });
  });
}

// 20 IP 확장 ② 이모티콘 + 굿즈
{
  const s = addContent(20, "IP 확장 ②: 이모티콘 · 굿즈", "결과물 · IP 확장");
  const cw = 5.95, gap = 0.23;
  const blocks = [
    ["emoticon", "이모티콘", "15종 (지니/미샤/오대오)", ["표정과 동작을 통한 캐릭터 성격 표현", "콩가리니쿠, 금방울 등 시그니처 소품 활용"]],
    ["Goods", "상품 굿즈", "작품 속 소품을 활용한 상품화", ["콩가리니쿠 봉제 인형 · 대형 쿠션", "지니 스카프 브로치 뱃지 · 오대오 금방울 키링"]],
  ];
  blocks.forEach(([key, t, sub, lines], i) => {
    const x = MX + i * (cw + gap);
    card(s, x, 1.35, cw, 5.1, CARD);
    framedImg(s, key, x + 0.25, 1.6, cw - 0.5, 3.15, { pad: 0.06, alt: t });
    text(s, t, { x: x + 0.3, y: 4.95, w: 2.0, h: 0.45, fontSize: 18, bold: true, color: INK, valign: "middle" });
    text(s, sub, { x: x + 2.1, y: 4.95, w: cw - 2.4, h: 0.45, fontSize: 12, color: BLUE_T, valign: "middle" });
    s.addText(lines.map((l, j) => ({ text: l, options: { bullet: true, breakLine: j < lines.length - 1 } })), {
      x: x + 0.3, y: 5.55, w: cw - 0.6, h: 1.0, fontFace: FONT, fontSize: 12, color: INK, valign: "top", margin: 0, paraSpaceAfter: 4, isTextBox: true, objectName: nm("ip-lines"),
    });
  });
}

// 21 AI 숏폼 광고
{
  const s = addContent(21, "AI 숏폼 광고: 당근 15초 광고", "결과물 · IP 확장");
  const p = framedImg(s, "poster_ad", MX, 1.4, 7.6, 4.275, { pad: 0.08, alt: "당근 15초 광고 대표 이미지 (클릭 시 영상 재생)", url: LINK_AD });
  text(s, "이미지 클릭 시 영상 재생 (Google Drive)", { x: p.x, y: p.y + p.h + 0.12, w: 7.6, h: 0.3, fontSize: 10, color: MUTED });
  const fw = 1.75, fg = (p.w - 4 * fw) / 3;
  for (let i = 0; i < 4; i++) placeholderFrame(s, p.x + i * (fw + fg), 6.05, fw, 0.8, `프레임 ${i + 1}`);
  const rx = 8.55, rw = W - MX - rx;
  card(s, rx, 1.4, rw, 4.275, CARD);
  const info = [["길이", "15초"], ["광고 대상", "[영상 확인 후 기입]"], ["캐릭터 등장", "[영상 확인 후 기입]"], ["광고 흐름", "[영상 확인 후 기입]"]];
  info.forEach(([k, v], i) => {
    const y = 1.7 + i * 0.95;
    text(s, k, { x: rx + 0.3, y, w: rw - 0.6, h: 0.3, fontSize: 11, bold: true, color: BLUE_T });
    text(s, v, { x: rx + 0.3, y: y + 0.32, w: rw - 0.6, h: 0.4, fontSize: 14, color: INK });
  });
  linkButton(s, "▶  광고 영상 보기", LINK_AD, rx, 6.0, rw, 0.6);
}

// 22 엔딩
{
  const s = pres.addSlide({ masterName: "DARK", sectionTitle: "결과물 · IP 확장" });
  text(s, "「근데 저녁은 뭐 먹지?」", { x: MX, y: 0.75, w: 7.2, h: 1.0, fontSize: 36, bold: true, color: WHITE, valign: "middle" });
  text(s, "다음 에피소드로 자연스럽게 연결", { x: MX, y: 1.8, w: 7, h: 0.4, fontSize: 14, color: SKY });
  const f = framedImg(s, "0802", 8.2, 0.75, 4.53, 2.6, { fill: SKY, pad: 0.08, alt: "뒷모습 와이드 엔딩" });
  const skills = [["연출 설계", "감정에 맞춘 샷/앵글 설계", BLUE, WHITE], ["AI 프롬프트 설계", "공통 비주얼 사양과 컷 기술 규칙 수립", PINK, INK], ["캐릭터 IP 확장", "애니메이션에서 로고/이모티콘/굿즈/광고로 확장", GREEN, INK]];
  skills.forEach(([k, v, col, tc], i) => {
    const y = 2.75 + i * 0.85;
    chip(s, k, MX, y, 2.4, 0.5, col, tc, 13);
    text(s, v, { x: MX + 2.7, y, w: 4.9, h: 0.5, fontSize: 15, color: WHITE, valign: "middle" });
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MX, y: 5.55, w: 12.13, h: 1.25, rectRadius: 0.14, fill: { color: "2C3A4B" }, line: { type: "none" }, objectName: nm("contact") });
  text(s, "이민상", { x: MX + 0.35, y: 5.75, w: 2, h: 0.5, fontSize: 20, bold: true, color: WHITE, valign: "middle" });
  [["이메일", 2.6], ["연락처", 6.0], ["포트폴리오 링크", 9.1]].forEach(([k, x]) => {
    text(s, k, { x: MX + x, y: 5.72, w: 2.8, h: 0.3, fontSize: 11, color: SKY });
    s.addShape(pres.shapes.LINE, { x: MX + x, y: 6.45, w: 2.7, h: 0, line: { color: "5B6B7C", width: 1 }, objectName: nm("blank-line") });
  });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();

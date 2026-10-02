// 조성기(CHOSK) 포트폴리오: 뮤직비디오 / 캐릭터 IP 웹툰 / 숏폼 광고 3개 작품
// node build.js  ->  ../조성기_포트폴리오.pptx
// 「I'm Not Joker」에서 가져온 트럼프 카드 모티프: 누아르 블랙, 조커 립 레드, 카드 아이보리, Impact 디스플레이
// 챕터 간지: 작품별 카드(A♥ / 2♦ / 3♠) 앞면과 뒷면을 흩뿌린 테이블 위 연출
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const IMG = path.join(__dirname, "../work/img");
const DIMS = require("../work/dims.json");
const OUT = path.join(__dirname, "../조성기_포트폴리오.pptx");

const LINK_MV = "https://drive.google.com/file/d/1FxAWrfKejtH5CH7youn3iG33vNt5Ns2T/view";

const SANS = "Malgun Gothic";
const DISP = "Impact";
const NOIR = "0E0C0B", CARD = "1C1817", LINE = "3A302C", IVORY = "F3EBDD", PAPER = "FBF7F0", INK = "1A1514", MUTED = "9A8E84";
const RED = "B3122E", RED_L = "E0475E", AMBER = "E08A2E", AMBER_D = "B8661A", NAVY = "3E5A9E", NAVY_L = "8EA6DA", BRASS = "B8955A";

const THEME = {
  name: "Wild Card",
  headFontFace: SANS,
  bodyFontFace: SANS,
  colors: { dk1: NOIR, lt1: IVORY, dk2: CARD, lt2: PAPER, accent1: RED, accent2: AMBER, accent3: NAVY, accent4: BRASS, accent5: RED_L, accent6: MUTED, hlink: RED_L, folHlink: RED_L },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: SANS, bodyFontFace: SANS };
pres.author = "조성기";
pres.title = "조성기 포트폴리오";
const W = 13.333, H = 7.5, MX = 0.7;

const footer = (c) => ({ text: { text: "CHOSK  ·  조성기 포트폴리오", options: { x: MX, y: 7.08, w: 6, h: 0.28, fontFace: SANS, fontSize: 9, color: c, margin: 0 } } });
const titlePh = (c) => ({ placeholder: { options: { name: "title", type: "title", x: MX, y: 0.74, w: 11.9, h: 0.72, fontFace: SANS, fontSize: 28, bold: true, color: c, align: "left", valign: "middle", margin: 0 }, text: "" } });
const num = (c) => ({ x: 12.13, y: 7.06, w: 0.5, h: 0.3, fontFace: DISP, fontSize: 12, color: c, align: "right" });
pres.defineSlideMaster({ title: "NOIR", background: { color: NOIR }, objects: [titlePh(IVORY), footer("7A6E66")], slideNumber: num("7A6E66") });
pres.defineSlideMaster({ title: "IVORY", background: { color: IVORY }, objects: [titlePh(INK), footer(MUTED)], slideNumber: num(MUTED) });
pres.defineSlideMaster({ title: "PLAIN", background: { color: NOIR }, objects: [] });

// ---------- helpers ----------
let sid = 0;
const nm = (s) => `${s}-${++sid}`;
const PNG = new Set(["cardback", "card1_face", "card2_face", "card3_face"]);
const file = (key) => path.join(IMG, `${key}.${PNG.has(key) ? "png" : "jpg"}`);
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
  if (o.border) s.addShape(pres.shapes.RECTANGLE, { x: f.x - 0.04, y: f.y - 0.04, w: f.w + 0.08, h: f.h + 0.08, fill: { color: o.border }, line: { type: "none" }, objectName: nm("border") });
  const link = o.url ? { hyperlink: { url: o.url, tooltip: o.tip || "영상 보기" } } : {};
  s.addImage({ path: file(key), ...f, rotate: o.rot || 0, altText: o.alt || key, objectName: nm(`img-${key}`), ...link });
  return f;
}
// 카드 한 장 (앞면 합성 PNG 또는 뒷면), 높이 기준 크기
function cardImg(s, key, x, y, h, rot, alt) {
  const [iw, ih] = DIMS[key];
  const w = h * iw / ih;
  s.addImage({ path: file(key), x, y, w, h, rotate: rot || 0, altText: alt || key, objectName: nm(`card-${key}`) });
  return { x, y, w, h };
}
function t(s, text, o) { s.addText(text, { fontFace: SANS, color: IVORY, margin: 0, valign: "top", isTextBox: true, objectName: nm(o.name || "text"), ...o }); }
function rect(s, x, y, w, h, fill, o = {}) {
  s.addShape(o.round ? pres.shapes.ROUNDED_RECTANGLE : pres.shapes.RECTANGLE, { x, y, w, h, rectRadius: o.round || 0, fill: { color: fill, transparency: o.tr || 0 }, line: o.line ? { color: o.line, width: o.lw || 0.75 } : { type: "none" }, objectName: nm("rect") });
}
function full(s, key, alt) { s.addImage({ path: file(key), x: 0, y: 0, w: W, h: H, altText: alt || "배경", objectName: nm(key) }); }
// 일반 슬라이드: 카드 무늬 라벨 + 제목 placeholder
function slide(master, suit, label, color, title, section, bgKey) {
  const s = pres.addSlide({ masterName: master, sectionTitle: section });
  if (bgKey) full(s, bgKey);
  t(s, suit, { x: MX, y: 0.36, w: 0.3, h: 0.36, fontSize: 15, bold: true, color, valign: "middle", name: "suit" });
  t(s, label, { x: MX + 0.32, y: 0.38, w: 9, h: 0.32, fontFace: DISP, fontSize: 14, color, charSpacing: 2, valign: "middle", name: "label" });
  s.addText(title, { placeholder: "title" });
  return s;
}
function bullets(s, items, x, y, w, o = {}) {
  const color = o.color || IVORY, mark = o.mark || RED, size = o.size || 15, gap = o.gap || 0.62;
  items.forEach((it, i) => {
    t(s, o.suit || "♦", { x, y: y + i * gap + 0.01, w: 0.25, h: 0.3, fontSize: size - 5, color: mark, name: "mark" });
    t(s, it, { x: x + 0.3, y: y + i * gap, w: w - 0.3, h: gap, fontSize: size, color, lineSpacingMultiple: 1.1, name: "bullet" });
  });
}
function tag(s, text, x, y, fill, color, o = {}) {
  const w = o.w || 1.0;
  rect(s, x, y, w, 0.3, fill);
  t(s, text, { x, y, w, h: 0.3, fontSize: o.size || 10, bold: true, color, align: "center", valign: "middle", fontFace: o.font || SANS, name: "tag" });
}
function linkText(s, label, url, x, y, w, h, color, size = 13) {
  s.addText([{ text: label, options: { hyperlink: { url, tooltip: url }, color, bold: true } }], { x, y, w, h, fontFace: SANS, fontSize: size, valign: "middle", margin: 0, isTextBox: true, objectName: nm("link") });
}

// 챕터 간지: 흐린 대표 장면 위, 뒷면 카드 두 장과 앞면 카드 한 장을 흩뿌린 연출
function chapter(n, rank, suit, color, cardKey, bgKey, title, en, sub, section) {
  const s = pres.addSlide({ masterName: "PLAIN", sectionTitle: section });
  full(s, bgKey, "챕터 대표 장면을 흐리게 깐 배경");
  rect(s, 0, 0, W, H, NOIR, { tr: 55 });
  cardImg(s, "cardback", 0.55, 1.35, 4.6, -16, "카드 뒷면");
  cardImg(s, "cardback", 2.9, 1.55, 4.6, 13, "카드 뒷면");
  cardImg(s, `${cardKey}_face`, 1.45, 0.75, 5.9, -3, `${rank}${suit} 카드 앞면 대표 이미지`);
  t(s, `CHAPTER ${n}`, { x: 7.2, y: 1.25, w: 5.4, h: 0.4, fontFace: DISP, fontSize: 18, color, charSpacing: 6 });
  t(s, [
    { text: rank, options: { fontFace: DISP, fontSize: 150, color: IVORY } },
    { text: suit, options: { fontFace: SANS, fontSize: 110, color } },
  ], { x: 7.1, y: 1.6, w: 5.6, h: 2.4, valign: "middle", name: "ch-rank" });
  rect(s, 7.2, 4.15, 1.2, 0.06, color);
  const two = title.includes("\n");
  t(s, title, { x: 7.2, y: 4.35, w: 5.6, h: two ? 1.3 : 0.7, fontSize: two ? 30 : 34, bold: true, color: IVORY, valign: "top", lineSpacingMultiple: 1.0 });
  const y2 = two ? 5.75 : 5.2;
  t(s, en, { x: 7.2, y: y2, w: 5.6, h: 0.35, fontFace: DISP, fontSize: 15, color: "C9BFB4", charSpacing: 2 });
  t(s, sub, { x: 7.2, y: y2 + 0.45, w: 5.6, h: 0.8, fontSize: 14, color: IVORY, lineSpacingMultiple: 1.2 });
  return s;
}

// =====================================================================
// 01 표지
pres.addSection({ title: "소개" });
{
  const s = pres.addSlide({ masterName: "PLAIN", sectionTitle: "소개" });
  full(s, "bg_cover", "트럼프 카드가 흩날리는 뮤지컬 무대를 흐리게 깐 배경");
  cardImg(s, "card3_face", 9.75, 1.55, 4.4, 15, "3♠ 숏폼 광고 카드");
  cardImg(s, "card2_face", 8.35, 1.15, 4.4, 2, "2♦ 캐릭터 IP 카드");
  cardImg(s, "card1_face", 6.9, 1.45, 4.4, -12, "A♥ 뮤직비디오 카드");
  t(s, "AI CONTENT CREATOR PORTFOLIO", { x: MX, y: 1.2, w: 6, h: 0.35, fontFace: DISP, fontSize: 15, color: RED_L, charSpacing: 4 });
  t(s, "CHOSK", { x: MX, y: 1.6, w: 6, h: 1.3, fontFace: DISP, fontSize: 88, color: IVORY, valign: "middle" });
  t(s, "조성기", { x: MX, y: 2.95, w: 5, h: 0.6, fontSize: 28, bold: true, color: IVORY });
  t(s, "AI와 인간 콜라보레이션으로 완성한 스토리텔링", { x: MX, y: 3.6, w: 6, h: 0.45, fontSize: 17, color: "D8CFC4" });
  [["A", "♥", RED_L, "뮤직비디오 「I'm Not Joker」"], ["2", "♦", AMBER, "캐릭터 IP 웹툰 「돈깨비 방망이」 / 「오해소지」"], ["3", "♠", NAVY_L, "AI 숏폼 광고 세로형 스토리보드"]].forEach(([r, su, c, x], i) => {
    const y = 4.45 + i * 0.55;
    t(s, [{ text: r, options: { fontFace: DISP, color: IVORY } }, { text: su, options: { color: c } }], { x: MX, y, w: 0.7, h: 0.45, fontSize: 20, valign: "middle" });
    t(s, x, { x: MX + 0.75, y, w: 5.6, h: 0.45, fontSize: 14.5, color: IVORY, valign: "middle" });
  });
  t(s, "서울시 매력일자리 AI 콘텐츠 크리에이터 양성과정  ·  2026", { x: MX, y: 6.7, w: 6.5, h: 0.3, fontSize: 10, color: "A89C92" });
}

// 02 목차
{
  const s = slide("NOIR", "♣", "CONTENTS", BRASS, "목차", "소개", "bg_noir");
  const rows = [
    ["card1_face", "A", "♥", RED_L, "I'm Not Joker", "뮤직비디오 / EINSX 「No Smile」"],
    ["card2_face", "2", "♦", AMBER, "돈깨비 방망이 / 오해소지", "캐릭터 IP / 스크롤 웹툰 / 이모티콘 / 일상툰"],
    ["card3_face", "3", "♠", NAVY_L, "숏폼 광고 스토리보드", "AI 숏폼 광고 / 9:16 세로형"],
  ];
  rows.forEach(([k, r, su, c, title, genre], i) => {
    const x = 1.05 + i * 4.0;
    cardImg(s, k, x + 0.55, 1.75, 3.4, 0, title);
    t(s, [{ text: r, options: { fontFace: DISP, color: IVORY } }, { text: su, options: { color: c } }], { x, y: 5.3, w: 1.0, h: 0.6, fontSize: 30, valign: "middle" });
    t(s, title, { x: x + 0.95, y: 5.3, w: 2.9, h: 0.6, fontSize: 18, bold: true, color: IVORY, valign: "middle" });
    t(s, genre, { x, y: 5.95, w: 3.8, h: 0.6, fontSize: 12, color: c, bold: true });
  });
}

// 03 역량
{
  const s = slide("NOIR", "♣", "PROFILE", BRASS, "AI와 인간 콜라보레이션으로 넓힌 창작 영역", "소개", "bg_noir");
  const cards = [
    ["A", "♥", RED_L, "스토리 기획", "청년의 방황과 재기를 담은 메시지를 뮤직비디오 서사로 구조화"],
    ["2", "♦", AMBER, "캐릭터 IP", "작가 본인 캐릭터부터 부산 사투리 소시지까지 개성 있는 IP 기획"],
    ["3", "♠", NAVY_L, "영상 연출", "컷 단위 샷/앵글 설계와 흑백에서 컬러로 이어지는 시간 전환 연출"],
    ["J", "♣", BRASS, "AI 제작 공정", "AI툴을 사용하여 글/그림/영상/음악 제작 역량 극대화"],
  ];
  cards.forEach(([r, su, c, h, b], i) => {
    const x = MX + i * 3.0, y = 1.95;
    rect(s, x, y, 2.8, 4.5, IVORY, { round: 0.12 });
    t(s, [{ text: r, options: { fontFace: DISP } }, { text: "\n" + su }], { x: x + 0.18, y: y + 0.15, w: 0.6, h: 1.0, fontSize: 22, color: c === BRASS ? "8C6A34" : (c === NAVY_L ? NAVY : (c === AMBER ? AMBER_D : RED)), align: "center", lineSpacingMultiple: 0.9 });
    t(s, h, { x: x + 0.3, y: y + 1.45, w: 2.3, h: 0.5, fontSize: 20, bold: true, color: INK });
    t(s, b, { x: x + 0.3, y: y + 2.1, w: 2.25, h: 2.0, fontSize: 14, color: "4A403A", lineSpacingMultiple: 1.25 });
  });
}

// =====================================================================
// CHAPTER A♥ I'm Not Joker
const L1 = "I'M NOT JOKER", S1 = "A♥ I'm Not Joker";
pres.addSection({ title: S1 });
chapter("A", "A", "♥", RED_L, "card1", "ch1_bg", "I'm Not Joker", "EINSX  \"NO SMILE\"  MUSIC VIDEO", "아무도 봐주지 않던 소극장 희극인 배우가\n은인을 만나 다시 태어나는 이야기", S1);

// 작품 개요
{
  const s = slide("NOIR", "♥", L1, RED_L, "노력은 배신하지 않는다는 청년의 이야기", S1, "bg_noir");
  img(s, "a1_1", MX, 1.85, 6.2, 3.49, { border: LINE, alt: "비 온 뒤 홍대 골목 벤치에 앉은 현준" });
  t(s, "S#1  홍대 거리, 소극장 앞 벤치", { x: MX, y: 5.45, w: 6.2, h: 0.3, fontSize: 11, color: MUTED });
  t(s, "로그라인", { x: 7.25, y: 1.85, w: 3, h: 0.3, fontSize: 12, bold: true, color: RED_L });
  t(s, "아무도 봐주지 않던 소극장 희극인 배우가, 은인을 만나 새로운 꿈을 찾게 되어 다시 태어나는 이야기", { x: 7.25, y: 2.2, w: 5.4, h: 1.0, fontSize: 16, bold: true, color: IVORY, lineSpacingMultiple: 1.2 });
  bullets(s, [
    "방황하는 청년에게 노력한 사람은 꼭 빛을 본다는 메시지 전달",
    "아이돌 그룹 EINSX 타이틀곡 「No Smile」을 스토리텔링 뮤직비디오로 기획",
  ], 7.25, 3.4, 5.4, { mark: RED_L, size: 14, gap: 0.8, suit: "♥" });
  const info = [["장르", "뮤직비디오 (심리 드라마)"], ["플랫폼", "유튜브"], ["타깃", "20~30대 / 아이돌 팬덤"], ["러닝타임", "4분"]];
  info.forEach(([a, b], i) => {
    const y = 5.1 + i * 0.4;
    t(s, a, { x: 7.25, y, w: 1.3, h: 0.36, fontSize: 12, bold: true, color: MUTED, valign: "middle" });
    t(s, b, { x: 8.55, y, w: 4.1, h: 0.36, fontSize: 13, color: IVORY, valign: "middle" });
  });
}

// 캐릭터
{
  const s = slide("NOIR", "♥", L1, RED_L, "퇴폐미 있는 외모의 노잼 희극인과 그를 알아본 은인", S1, "bg_noir");
  const ch = [
    ["hyunjun", "박현준", "주인공 / EINSX 멤버", "ISTJ, 퇴폐미 있게 잘생긴 외모. 주위 권유로 희극인이 됐지만 웃음의 재능이 없어 방황하는 인물"],
    ["kanghoon", "강훈", "소극장 연출자 / 40대", "현준의 마스크를 알아보고 뮤지컬에 캐스팅, 가창력까지 발견해 새 길을 열어 주는 은인"],
  ];
  ch.forEach(([k, n, role, d], i) => {
    const x = MX + i * 6.1;
    img(s, k, x, 1.85, 5.8, 3.25, { border: IVORY, alt: `${n} 캐릭터 시트` });
    t(s, n, { x, y: 5.3, w: 2, h: 0.45, fontSize: 21, bold: true, color: IVORY });
    t(s, role, { x: x + 1.55, y: 5.36, w: 4, h: 0.35, fontSize: 12, bold: true, color: RED_L });
    t(s, d, { x, y: 5.85, w: 5.8, h: 0.8, fontSize: 13, color: "D8CFC4", lineSpacingMultiple: 1.2 });
  });
}

// 3막 서사 구조
{
  const s = slide("NOIR", "♥", L1, RED_L, "몰락에서 비상까지, 3막 서사 구조", S1, "bg_alley");
  const acts = [
    ["몰락", "소극장 무대의 싸늘한 반응과 계란 세례, 골목에서의 주먹다짐", ["a5_3", "a7_9"]],
    ["전환", "쓰러진 현준에게 던져진 뮤지컬 대본, 녹음실에서 드러난 가창력", ["a7_14", "a9_1"]],
    ["비상", "트럼프 카드가 흩날리는 대극장 무대와 처음으로 활짝 웃는 현준", ["a12_6", "a16_5"]],
  ];
  acts.forEach(([h, d, ks], i) => {
    const x = MX + i * 4.1;
    t(s, [{ text: `ACT ${i + 1}  `, options: { fontFace: DISP, color: RED_L } }, { text: h, options: { bold: true, color: IVORY } }], { x, y: 1.8, w: 3.6, h: 0.45, fontSize: 18, valign: "middle" });
    img(s, ks[0], x, 2.35, 3.6, 2.03, { border: LINE, alt: `${h} 장면 1` });
    img(s, ks[1], x, 4.5, 3.6, 2.03, { border: LINE, alt: `${h} 장면 2` });
    if (i < 2) t(s, "▶", { x: x + 3.68, y: 3.25, w: 0.35, h: 0.4, fontSize: 16, color: RED_L, align: "center" });
  });
  acts.forEach(([h, d], i) => t(s, d, { x: MX + i * 4.1, y: 6.6, w: 3.6, h: 0.45, fontSize: 11, color: "D8CFC4" }));
}

// 조커 메이크업 모티프
{
  const s = slide("NOIR", "♥", L1, RED_L, "붉게 번진 입꼬리, 작품을 관통하는 조커 모티프", S1, "bg_noir");
  const cuts = [
    ["a3_3", "S#3", "분장실, 입꼬리까지 삐져나오게 그린 붉은 립"],
    ["a7_15", "S#7", "분장인지 피인지 모르게 번진 입술과 WILD CARD 대본"],
    ["a16_5", "S#17", "같은 얼굴에 처음으로 피어난 진짜 웃음"],
  ];
  cuts.forEach(([k, c, d], i) => {
    const x = MX + i * 4.0;
    img(s, k, x, 1.85, 3.8, 2.14, { border: i === 2 ? RED : LINE, alt: `${c} 장면` });
    tag(s, c, x, 4.15, RED, IVORY, { w: 0.8, font: DISP, size: 12 });
    t(s, d, { x, y: 4.55, w: 3.8, h: 0.7, fontSize: 13, color: IVORY, lineSpacingMultiple: 1.2 });
  });
  rect(s, MX, 5.45, 11.9, 1.2, CARD, { tr: 15 });
  rect(s, MX, 5.45, 0.08, 1.2, RED);
  bullets(s, [
    "웃음을 강요받는 분장을 몰락과 비상을 잇는 시각 모티프로 반복 배치",
    "'WILD CARD' 대본과 트럼프 카드 연출로 인생 역전의 계기를 상징화",
  ], MX + 0.35, 5.6, 11.4, { mark: RED_L, size: 14, gap: 0.48, suit: "♥" });
}

// 뮤지컬 무대 연출
{
  const s = slide("NOIR", "♥", L1, RED_L, "카드와 샹들리에로 완성한 무대 스펙터클", S1, "bg_noir");
  img(s, "a12_1", MX, 1.85, 7.3, 4.11, { border: LINE, alt: "샹들리에와 붉은 커튼 앞에서 노래하는 현준" });
  img(s, "a11_1", 8.25, 1.85, 4.35, 2.37, { border: BRASS, alt: "트럼프 카드 위로 떠오른 몇 달 후 인서트" });
  img(s, "a12_7", 8.25, 4.36, 4.35, 1.6, { top: true, border: LINE, alt: "거대한 공연장 한가운데 선 현준" });
  bullets(s, [
    "'몇 달 후' 카드 인서트로 시간 경과를 장면 전환에 녹인 연출",
    "붉은 커튼 / 샹들리에 / 흩날리는 카드로 소극장과 대비되는 무대 스케일 표현",
    "OST 「내 심장이 하는 일」을 SUNO AI로 제작, 감정 고조 구간과 컷 호흡 일치",
  ], MX, 6.12, 11.9, { mark: RED_L, size: 12.5, gap: 0.3, suit: "♥" });
}

// AI 워크플로우 1: 기획
{
  const s = slide("NOIR", "♥", L1, RED_L, "Claude 프로젝트로 체계화한 기획 / 프롬프트 설계", S1, "bg_wf");
  img(s, "wf_claude", MX, 1.85, 5.85, 3.29, { border: LINE, alt: "Claude 프로젝트 I'm Not Joker 화면" });
  img(s, "wf_prompt", 6.75, 1.85, 5.85, 3.29, { border: LINE, alt: "립싱크 영상 프롬프트 작성 화면" });
  tag(s, "Claude", MX, 5.3, RED, IVORY, { w: 1.0 });
  tag(s, "Claude", 6.75, 5.3, RED, IVORY, { w: 1.0 });
  t(s, "캐릭터 설정 / 씬별 이미지 프롬프트 / 촬영 계획을 한 프로젝트에서 관리", { x: MX, y: 5.72, w: 5.85, h: 0.7, fontSize: 13, color: IVORY });
  t(s, "카메라 고정 / 표정 / 대사 타이밍까지 지정한 립싱크 영상 프롬프트로 연출 의도 정밀 반영", { x: 6.75, y: 5.72, w: 5.85, h: 0.7, fontSize: 13, color: IVORY });
}

// AI 워크플로우 2: 생성과 후반
{
  const s = slide("NOIR", "♥", L1, RED_L, "생성부터 후반 편집까지 이어진 AI 제작 공정", S1, "bg_wf");
  const wf = [
    ["wf_gen1", "PixVerse", "흑백 시네마틱 영상 생성"],
    ["wf_lipsync", "LipSync", "노래 립싱크 영상 생성"],
    ["wf_clip", "클립스튜디오", "얼굴형 / 입꼬리 디지털 성형 보정"],
    ["wf_premiere", "프리미어", "컷 편집 / 타이틀 / 사운드 완성"],
  ];
  wf.forEach(([k, tool, d], i) => {
    const x = MX + i * 3.0;
    img(s, k, x, 1.9, 2.85, 1.6, { border: LINE, alt: `${tool} 작업 화면` });
    tag(s, tool, x, 3.65, i === 3 ? BRASS : RED, IVORY, { w: 1.35 });
    t(s, d, { x, y: 4.05, w: 2.85, h: 0.6, fontSize: 12.5, color: IVORY });
    if (i < 3) t(s, "▶", { x: x + 2.85, y: 2.5, w: 0.15, h: 0.4, fontSize: 10, color: RED_L, align: "center" });
  });
  rect(s, MX, 4.95, 11.9, 1.7, CARD, { tr: 15 });
  t(s, "Project: Digital Face Remodeling", { x: MX + 0.3, y: 5.1, w: 6, h: 0.35, fontFace: DISP, fontSize: 15, color: RED_L });
  bullets(s, [
    "얼굴 크기 축소 / 턱선 V라인 정리 / 조커 콘셉트의 입꼬리 리프팅 / 황금 비율 이목구비 재배치",
    "AI 생성 결과를 그대로 쓰지 않고 캐릭터 설정에 맞춰 직접 보정한 완성도 관리",
  ], MX + 0.3, 5.5, 11.3, { mark: RED_L, size: 13, gap: 0.5, suit: "♥" });
}

// 영상
{
  const s = slide("NOIR", "♥", L1, RED_L, "뮤직비디오 1차 편집본", S1, "bg_noir");
  img(s, "poster_mv", MX, 1.85, 7.6, 4.28, { border: LINE, url: LINK_MV, alt: "I'm Not Joker 1차 편집본 영상 보기" });
  t(s, "이미지를 클릭하면 영상이 재생됩니다", { x: MX, y: 6.3, w: 7.6, h: 0.3, fontSize: 11, color: MUTED });
  const x = 8.75;
  tag(s, "1차 편집 / 미완성", x, 1.9, RED, IVORY, { w: 1.7 });
  t(s, "I'm Not Joker", { x, y: 2.35, w: 3.9, h: 0.6, fontFace: DISP, fontSize: 32, color: IVORY });
  t(s, "EINSX  「No Smile (웃을 수 없는 남자)」", { x, y: 3.0, w: 3.9, h: 0.4, fontSize: 13, color: RED_L, bold: true });
  bullets(s, [
    "2022년 흑백 화면에서 2026년 컬러 데뷔 무대로 이어지는 시간 전환",
    "분장실과 대기실을 대구로 배치한 데뷔 엔딩",
  ], x, 3.65, 3.9, { mark: RED_L, size: 13, gap: 0.85, suit: "♥" });
  rect(s, x, 5.55, 2.45, 0.5, RED);
  linkText(s, "▶  영상 보기", LINK_MV, x, 5.55, 2.45, 0.5, IVORY, 13);
}

// =====================================================================
// CHAPTER 2♦ 캐릭터 IP
const L2 = "CHARACTER IP", S2 = "2♦ 캐릭터 IP";
pres.addSection({ title: S2 });
chapter("2", "2", "♦", AMBER, "card2", "ch2_bg", "돈깨비 방망이\n오해소지", "CHARACTER IP  &  WEBTOON", "스크롤 웹툰 / 이모티콘 / 일상툰\n작가 본인 캐릭터에서 출발한 IP 확장", S2);

// 캐릭터 시트
{
  const s = slide("IVORY", "♦", L2, AMBER_D, "작가 본인 캐릭터와 두 조력자", S2, "bg_ivory");
  const ch = [
    ["eun", "조은탁", "작가 본인 캐릭터", "통장 잔고 100원에 한숨 쉬는 현실 공감형 주인공"],
    ["isin", "이신", "도깨비", "돈깨비 방망이를 든 냉정한 조력자, 노력 없는 소원은 단호히 거절"],
    ["reaper_toon", "킹여", "저승사자", "이신과 함께 나타나는 검은 코트의 동행자"],
  ];
  ch.forEach(([k, n, role, d], i) => {
    const x = MX + i * 4.0;
    rect(s, x, 1.8, 3.8, 3.4, "FFFFFF", { line: "E3D7C4" });
    img(s, k, x + 0.1, 1.9, 3.6, 3.2, { alt: `${n} 캐릭터 시트` });
    t(s, n, { x, y: 5.35, w: 1.6, h: 0.45, fontSize: 20, bold: true, color: INK });
    tag(s, role, x + 1.5, 5.42, AMBER, "FFFFFF", { w: 1.6 });
    t(s, d, { x, y: 5.9, w: 3.8, h: 0.75, fontSize: 12.5, color: "4A403A", lineSpacingMultiple: 1.2 });
  });
}

// 스크롤 웹툰
{
  const s = slide("IVORY", "♦", L2, AMBER_D, "스크롤 웹툰 「돈깨비 방망이」", S2, "bg_ivory");
  ["web1", "web2", "web3", "web4"].forEach((k, i) => {
    img(s, k, MX + i * 1.95, 1.8, 1.8, 4.95, { top: true, border: "FFFFFF", alt: `돈깨비 방망이 장면 ${i + 1}` });
  });
  const x = 8.65;
  img(s, "web_title", x, 1.8, 3.95, 1.9, { left: true, alt: "돈깨비 방망이 타이틀 카드" });
  t(s, "돈이 하늘에서 떨어지길 바라는 조은탁 앞에 나타난 도깨비와 저승사자, 노력하지 않는 인간은 소원을 빌 자격도 없다는 반전", { x, y: 3.85, w: 3.95, h: 1.2, fontSize: 13, bold: true, color: INK, lineSpacingMultiple: 1.2 });
  bullets(s, [
    "놀라거나 설렐 때 그림체가 바뀌는 연출로 감정 변화를 코믹하게 표현",
    "'지옥 같은 회사'를 이겨내라는 결말로 직장인 공감 극대화",
  ], x, 5.1, 3.95, { color: "4A403A", mark: AMBER, size: 12, gap: 0.62 });
  t(s, "글: CHOSK, claude_AI  /  그림: gemini_AI, pixverse.AI, googleAI", { x, y: 6.4, w: 3.95, h: 0.3, fontSize: 9.5, color: AMBER_D, bold: true });
}

// 오해소지
{
  const s = slide("IVORY", "♦", L2, AMBER_D, "부산 사투리 소시지 캐릭터 「오해소지」", S2, "bg_ivory");
  img(s, "sausage_prof", MX, 1.75, 2.0, 5.0, { left: true, border: "FFFFFF", alt: "오해소지 캐릭터 소개 프로필" });
  const x = 3.2;
  t(s, "\"마, 오해하지 마라카이. 내는 햄이다! 알긋나?\"", { x, y: 1.85, w: 6.0, h: 0.5, fontSize: 17, bold: true, color: INK });
  const info = [["모티브", "가성비 좋은 분홍 소시지"], ["설정", "부산 토박이 / 180g / 5살 / 외동아들 / INFP"], ["성격", "부끄러움을 감추려 괜히 센 척, 듬직한 '햄(형님)'이 되고 싶은 막내"]];
  info.forEach(([a, b], i) => {
    const y = 2.55 + i * 0.5;
    t(s, a, { x, y, w: 1.0, h: 0.4, fontSize: 12, bold: true, color: AMBER_D, valign: "middle" });
    t(s, b, { x: x + 1.0, y, w: 5.1, h: 0.4, fontSize: 13, color: "4A403A", valign: "middle" });
  });
  rect(s, x, 4.15, 6.0, 2.55, "FFFFFF", { line: "E3D7C4", round: 0.1 });
  img(s, "sausage_wanna", x + 0.2, 4.3, 2.25, 2.25, { alt: "왔나?! 이모티콘" });
  t(s, "부산 사투리 이모티콘", { x: x + 2.65, y: 4.4, w: 3.2, h: 0.35, fontSize: 13, bold: true, color: AMBER_D });
  t(s, "왔나?!  /  모라꼬?  /  밥도!\n마! 카톡 봤나 안 봤나? 으이?", { x: x + 2.65, y: 4.85, w: 3.2, h: 0.9, fontSize: 13.5, bold: true, color: INK, lineSpacingMultiple: 1.3 });
  t(s, "채팅 화면 형식의 프로필로 캐릭터 세계관을 대화하듯 소개", { x: x + 2.65, y: 5.85, w: 3.2, h: 0.7, fontSize: 11.5, color: "6A5E56" });
  img(s, "sausage_emo", 9.45, 1.85, 3.15, 3.15, { border: "FFFFFF", alt: "오해소지 이모티콘 4종" });
  t(s, "기본형 + 감정별 변형 이모티콘", { x: 9.45, y: 5.15, w: 3.15, h: 0.3, fontSize: 11.5, bold: true, color: AMBER_D, align: "center" });
}

// 일상툰
{
  const s = slide("IVORY", "♦", L2, AMBER_D, "일상툰 「햄이 되고싶어」", S2, "bg_ivory");
  img(s, "hamtoon", MX, 1.7, 3.25, 5.1, { left: true, border: "FFFFFF", alt: "일상툰 햄이 되고싶어" });
  const x = 4.5;
  tag(s, "6컷 일상툰", x, 1.85, AMBER, "FFFFFF", { w: 1.3 });
  t(s, "부산에서 형님을 '햄'이라 부르는 사투리를 소재로 한 언어유희", { x, y: 2.35, w: 8.1, h: 0.5, fontSize: 18, bold: true, color: INK });
  bullets(s, [
    "냉장고에서 떨어지는 아기햄을 구한 오해소지가 '햄' 소리에 빠져드는 짧은 에피소드",
    "수채화풍 냉장고 배경과 둥근 캐릭터로 친근한 브랜드 톤 확보",
    "프로필 / 이모티콘 / 일상툰으로 이어지는 캐릭터 IP 확장 구조",
  ], x, 3.1, 8.1, { color: "4A403A", mark: AMBER, size: 14.5, gap: 0.72 });
  rect(s, x, 5.4, 8.1, 1.3, "FFFFFF", { line: "E3D7C4" });
  t(s, "\"엣헴! 내를 이제 햄이라 부르레이~\"", { x: x + 0.3, y: 5.55, w: 7.5, h: 0.45, fontSize: 17, bold: true, color: AMBER_D });
  t(s, "\"영원히 햄으로 모시겠습니다!!\"", { x: x + 0.3, y: 6.05, w: 7.5, h: 0.45, fontSize: 15, color: INK });
}

// =====================================================================
// CHAPTER 3♠ 숏폼 광고
const L3 = "SHORT-FORM AD", S3 = "3♠ 숏폼 광고";
pres.addSection({ title: S3 });
chapter("3", "3", "♠", NAVY_L, "card3", "ch3_bg", "숏폼 광고\n스토리보드", "AI SHORT-FORM AD  9:16", "어두운 골목의 추격에서\n출근길 지하철로 이어지는 세로형 시퀀스", S3);

// 캐릭터
{
  const s = slide("NOIR", "♠", L3, NAVY_L, "세로형 화면에 맞춘 두 인물 설계", S3, "bg_ad");
  rect(s, MX, 1.8, 8.0, 2.85, "E6E6E6");
  img(s, "bays", MX + 0.1, 1.85, 7.8, 2.75, { alt: "배이스 캐릭터 시트" });
  rect(s, 9.0, 1.8, 3.6, 4.85, "E6E6E6");
  img(s, "reaper_ad", 9.05, 1.85, 3.5, 4.75, { alt: "저승사자 캐릭터 시트" });
  t(s, "배이스", { x: MX, y: 4.85, w: 2, h: 0.45, fontSize: 20, bold: true, color: IVORY });
  t(s, "블라우스 / 데님 롱스커트 / 사원증 / 붉은 마스크를 시그니처로 한 직장인 캐릭터, 7방향 시트로 컷마다 의상 일관성 유지", { x: MX, y: 5.35, w: 8.0, h: 0.75, fontSize: 13, color: "D8CFC4", lineSpacingMultiple: 1.2 });
  t(s, "저승사자", { x: 9.0, y: 6.7, w: 1.5, h: 0.3, fontSize: 12, bold: true, color: NAVY_L });
  t(s, "검은 로브 / 정장 두 가지 버전", { x: 10.25, y: 6.7, w: 2.4, h: 0.3, fontSize: 11, color: "D8CFC4" });
}

// 시퀀스 1
{
  const s = slide("NOIR", "♠", L3, NAVY_L, "S#1  골목 추격, 긴장감을 쌓는 세로 구도", S3, "bg_ad");
  const ks = ["d1_1", "d1_2", "d1_3", "d1_4"];
  ks.forEach((k, i) => {
    const f = img(s, k, MX + 0.4 + i * 2.85, 1.8, 2.7, 4.0, { top: true, border: LINE, alt: `S#1 C${i + 1}` });
    tag(s, `S#1 C${i + 1}`, f.x, 5.9, NAVY, IVORY, { w: 0.95, font: DISP, size: 12 });
  });
  bullets(s, [
    "좁은 골목의 소실점 구도와 세로 화면으로 도망치는 인물의 고립감 강조",
    "점점 다가오는 저승사자의 실루엣과 붙잡히는 순간까지 컷 단위로 긴장감 고조",
  ], MX, 6.35, 11.9, { mark: NAVY_L, size: 12.5, gap: 0.33, suit: "♠" });
}

// 시퀀스 2
{
  const s = slide("NOIR", "♠", L3, NAVY_L, "S#2~S#3  반전의 미소에서 출근길로", S3, "bg_ad");
  const ks = [["d2_1", "S#2 C1"], ["d2_2", "S#2 C2"], ["d2_3", "S#2 C3"], ["d3_1", "S#3 C1"], ["d3_2", "S#3 C2"]];
  ks.forEach(([k, c], i) => {
    const f = img(s, k, MX + i * 2.4, 1.8, 2.25, 4.0, { border: LINE, alt: c });
    tag(s, c, f.x, 5.9, NAVY, IVORY, { w: 0.95, font: DISP, size: 12 });
  });
  bullets(s, [
    "붙잡힌 순간의 섬뜩한 미소와 붉은 마스크로 장르를 뒤집는 반전 연출",
    "심야 골목에서 붐비는 아침 지하철로 공간을 전환하며 붉은 마스크를 반복 노출",
  ], MX, 6.35, 11.9, { mark: NAVY_L, size: 12.5, gap: 0.33, suit: "♠" });
}

// =====================================================================
// 엔딩
pres.addSection({ title: "마무리" });
{
  const s = pres.addSlide({ masterName: "PLAIN", sectionTitle: "마무리" });
  full(s, "bg_end", "조명이 쏟아지는 대극장 무대를 흐리게 깐 배경");
  cardImg(s, "cardback", 5.55, 0.75, 3.1, -8, "카드 뒷면");
  cardImg(s, "card1_face", 5.95, 0.95, 3.1, 7, "A♥ 카드");
  t(s, "THANK YOU", { x: 0, y: 4.45, w: W, h: 0.4, fontFace: DISP, fontSize: 18, color: RED_L, charSpacing: 8, align: "center" });
  t(s, "감사합니다", { x: 0, y: 4.85, w: W, h: 0.9, fontSize: 44, bold: true, color: IVORY, align: "center" });
  t(s, "조성기  ·  CHOSK", { x: 0, y: 5.8, w: W, h: 0.45, fontSize: 18, bold: true, color: IVORY, align: "center" });
  t(s, "뮤직비디오 / 캐릭터 IP 웹툰 / 숏폼 광고", { x: 0, y: 6.3, w: W, h: 0.4, fontSize: 13, color: "C9BFB4", align: "center" });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();

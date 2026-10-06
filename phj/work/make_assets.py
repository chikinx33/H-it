# 티켓 간지용 크롭(가로:세로 = 7.2:4.4), 목차 티켓 사진(4:3), 표지 세로 카드, 영상 포스터(체리 재생 버튼)
from PIL import Image, ImageDraw, ImageFilter
import os
os.chdir(os.path.join(os.path.dirname(__file__), ".."))
A, L, C, D, O = "source/애니메이션_아이돌출발", "source/애니메이션_너무좋아", "source/캐릭터", "source/숏폼광고", "work/img"

def crop_ratio(src, name, ratio, fx=0.5, fy=0.5, maxw=1800, img=None):
    im = img if img is not None else Image.open(src).convert("RGB")
    w, h = im.size
    if w / h > ratio:
        nw = round(h * ratio); l = round((w - nw) * fx); box = (l, 0, l + nw, h)
    else:
        nh = round(w / ratio); t = round((h - nh) * fy); box = (0, t, w, t + nh)
    im = im.crop(box)
    if im.width > maxw: im = im.resize((maxw, round(maxw / ratio)), Image.LANCZOS)
    im.save(f"{O}/{name}.jpg", quality=88)

T = 7.2 / 4.4
crop_ratio(f"{A}/컷 4-1B.png", "tk_ch1", T)
crop_ratio(f"{L}/컷 1-5.png", "tk_ch2", T, fy=0.42)
crop_ratio(f"{O}/promo.jpg", "tk_ch3", T)
# 숏폼 광고 간지: 두 제품 컨셉을 나란히 합성
cv = Image.new("RGB", (1636, 1000), (246, 238, 250))
a = Image.open(f"{D}/컨셉.png").convert("RGB").crop((820, 110, 1536, 1010)); a.thumbnail((800, 960))
b = Image.open(f"{D}/확빠져_컨셉.png").convert("RGB").crop((962, 12, 1660, 462)); b.thumbnail((800, 960))
cv.paste(a, (18 + (800 - a.width) // 2, (1000 - a.height) // 2)); cv.paste(b, (818 + (800 - b.width) // 2, (1000 - b.height) // 2))
cv.save(f"{O}/tk_ch4.jpg", quality=88)
# 목차 티켓 사진 4:3
crop_ratio(f"{A}/컷 4-2.png", "toc1", 4 / 3)
crop_ratio(f"{L}/컷 1-6.png", "toc2", 4 / 3, fy=0.35)
crop_ratio(f"{O}/promo.jpg", "toc3", 4 / 3)
crop_ratio(f"{O}/tk_ch4.jpg", "toc4", 4 / 3)
# 표지 세로 카드 3:4
crop_ratio(f"{A}/컷 4-2.png", "cover_card", 3 / 4, fx=0.38)
# 영상 포스터: 16:9, 체리 레드 원 + 흰 삼각형
def poster(src, name, ratio=16 / 9, fx=0.5, fy=0.5, box=None):
    img = Image.open(src).convert("RGB").crop(box) if box else None
    crop_ratio(src, name, ratio, fx, fy, maxw=1600, img=img)
    im = Image.open(f"{O}/{name}.jpg").convert("RGBA")
    w, h = im.size
    dark = Image.new("RGBA", (w, h), (30, 18, 51, 70)); im = Image.alpha_composite(im, dark)
    lay = Image.new("RGBA", (w, h), (0, 0, 0, 0)); d = ImageDraw.Draw(lay)
    cx, cy, r = w // 2, h // 2, int(h * 0.11)
    d.ellipse([cx - r - 10, cy - r - 10, cx + r + 10, cy + r + 10], fill=(255, 255, 255, 110))
    d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=(215, 38, 61, 255))
    s = r * 0.5; d.polygon([(cx - s * 0.7, cy - s), (cx - s * 0.7, cy + s), (cx + s * 1.05, cy)], fill=(255, 255, 255, 255))
    Image.alpha_composite(im, lay).convert("RGB").save(f"{O}/{name}.jpg", quality=88)
poster(f"{A}/컷 4-1B.png", "poster_idol")
poster(f"{O}/promo.jpg", "poster_bbem")
poster(f"{D}/컨셉.png", "poster_magic", ratio=4 / 5, box=(815, 0, 1536, 1024))
poster(f"{D}/확빠져_컨셉.png", "poster_slim", ratio=4 / 5, fx=1.0, fy=0.0)

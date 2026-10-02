# 갠지의 발광 링(블랙홀 테두리) PNG, 원형으로 잘라 낸 장면 사본, 링 재생 버튼 포스터
from PIL import Image, ImageDraw, ImageFilter
import os
os.chdir(os.path.join(os.path.dirname(__file__), ".."))
A, C, D, O = "source/애니메이션", "source/캐릭터", "source/숏폼광고", "work/img"
RING = (70, 211, 230)

def ring(name, size=1200, r=480, width=10):
    glow = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    g = ImageDraw.Draw(glow)
    c = size // 2
    g.ellipse([c - r, c - r, c + r, c + r], outline=RING + (255,), width=width * 5)
    glow = glow.filter(ImageFilter.GaussianBlur(28))
    sharp = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(sharp)
    d.ellipse([c - r, c - r, c + r, c + r], outline=(200, 248, 255, 255), width=width)
    out = Image.alpha_composite(glow, sharp)
    out.save(f"{O}/{name}.png")

def circle(src, name, box, size=960):
    im = Image.open(src).convert("RGB").crop(box).resize((size, size), Image.LANCZOS)
    mask = Image.new("L", (size * 4, size * 4), 0)
    ImageDraw.Draw(mask).ellipse([0, 0, size * 4 - 1, size * 4 - 1], fill=255)
    mask = mask.resize((size, size), Image.LANCZOS)
    im.putalpha(mask)
    im.save(f"{O}/{name}.png")

def poster(src, name, w=1600, h=900):
    im = Image.open(src).convert("RGB")
    sw, sh = im.size
    s = max(w / sw, h / sh)
    im = im.resize((round(sw * s), round(sh * s)), Image.LANCZOS)
    l, t = (im.width - w) // 2, (im.height - h) // 2
    im = im.crop((l, t, l + w, t + h))
    dark = Image.new("RGB", (w, h), (15, 14, 19))
    im = Image.blend(im, dark, 0.3).convert("RGBA")
    lay = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    g = ImageDraw.Draw(lay)
    cx, cy, r = w // 2, h // 2, 92
    g.ellipse([cx - r, cy - r, cx + r, cy + r], outline=RING + (255,), width=34)
    lay = lay.filter(ImageFilter.GaussianBlur(14))
    d = ImageDraw.Draw(lay)
    d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=(15, 14, 19, 150), outline=(200, 248, 255, 255), width=8)
    d.polygon([(cx - 26, cy - 44), (cx - 26, cy + 44), (cx + 48, cy)], fill=(255, 255, 255, 255))
    Image.alpha_composite(im, lay).convert("RGB").save(f"{O}/{name}.jpg", quality=88)

ring("ring")
# 원형 사본: 정사각 영역을 잘라 원으로 (비율 1:1 유지)
circle(f"{A}/씬07_컷02.png", "cir_sun", (466, 0, 1407, 941))
circle(f"{A}/씬02컷05.png", "cir_cover", (400, 0, 1341, 941))
circle(f"{C}/1월 캐릭터시트.png", "cir_hak", (170, 70, 730, 630))
circle(f"{D}/쇼츠광고-스토리보드8컷.png", "cir_sb", (56 + 120, 598, 56 + 120 + 326, 598 + 326))
circle(f"{A}/08.60대 경수.png", "cir_old", (880, 190, 1440, 750))
poster(f"{A}/씬01 컷02.png", "poster_s01")
poster(f"{A}/씬02컷05.png", "poster_s02")

# 황금 큐브 글로우(엔딩), 영상 포스터(잉크 테두리 원 + 황금 삼각형)
from PIL import Image, ImageDraw, ImageFilter
import os
os.chdir(os.path.join(os.path.dirname(__file__), ".."))
O = "work/img"
GOLD = (222, 178, 84)

# 큐브: 정사각 프레임 + 안쪽 은은한 빛 + 위에서 내려오는 케이블 다발
S = 1200
glow = Image.new("RGBA", (S, S), (0, 0, 0, 0)); g = ImageDraw.Draw(glow)
a, b = 270, 930
g.rectangle([a, a, b, b], fill=GOLD + (60,), outline=GOLD + (255,), width=40)
glow = glow.filter(ImageFilter.GaussianBlur(34))
top = Image.new("RGBA", (S, S), (0, 0, 0, 0)); d = ImageDraw.Draw(top)
d.rectangle([a, a, b, b], outline=(255, 226, 150, 255), width=7)
for i, x in enumerate(range(560, 660, 14)):
    d.line([(x, 0), (x + (i % 3 - 1) * 10, a)], fill=(205, 170, 95, 200), width=4)
Image.alpha_composite(glow, top).save(f"{O}/cube.png")

def poster(src, name):
    im = Image.open(f"{O}/{src}.jpg").convert("RGBA")
    w, h = im.size
    im = Image.alpha_composite(im, Image.new("RGBA", (w, h), (14, 34, 37, 80)))
    d = ImageDraw.Draw(im)
    cx, cy, r = w // 2, h // 2, int(h * 0.16)
    d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=(239, 230, 214, 235), outline=(31, 27, 22, 255), width=8)
    s = r * 0.48
    d.polygon([(cx - s * 0.65, cy - s), (cx - s * 0.65, cy + s), (cx + s * 1.05, cy)], fill=(196, 146, 40, 255))
    im.convert("RGB").save(f"{O}/{name}.jpg", quality=88)
poster("frame_anim", "poster_anim")
poster("frame_rot", "poster_rot")

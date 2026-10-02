# 「연결 꼬리」 모티프: 작품들을 하나로 잇는 꼬리 선 PNG 생성 (슬라이드 좌표 인치 x 120px)
from PIL import Image, ImageDraw, ImageFilter
import os
OUT = os.path.join(os.path.dirname(__file__), "img")
PX = 120
W, H = 1600, 900

def bez(p0, p1, p2, p3, n=120):
    pts = []
    for i in range(n + 1):
        t = i / n
        a = (1 - t) ** 3; b = 3 * (1 - t) ** 2 * t; c = 3 * (1 - t) * t ** 2; d = t ** 3
        pts.append((a * p0[0] + b * p1[0] + c * p2[0] + d * p3[0], a * p0[1] + b * p1[1] + c * p2[1] + d * p3[1]))
    return pts

def chain(start, segs):
    pts, cur = [], start
    for c1, c2, end in segs:
        pts += bez(cur, c1, c2, end)
        cur = end
    return [(x * PX, y * PX) for x, y in pts]

def stroke(draw, pts, width, color, taper=False):
    n = len(pts)
    for i, (x, y) in enumerate(pts):
        w = width * (0.35 + 0.65 * (1 - i / n)) if taper else width
        r = w / 2
        draw.ellipse([x - r, y - r, x + r, y + r], fill=color)

def make(name, start, segs, color, width=7, glow=True, tip=True, taper=False):
    pts = chain(start, segs)
    img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    if glow:
        g = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        stroke(ImageDraw.Draw(g), pts, width * 4, color[:3] + (110,))
        img = Image.alpha_composite(img, g.filter(ImageFilter.GaussianBlur(14)))
    line = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(line)
    stroke(d, pts, width, color, taper)
    if tip:  # 꼬리 끝 불씨
        x, y = pts[-1]
        tg = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        ImageDraw.Draw(tg).ellipse([x - 34, y - 34, x + 34, y + 34], fill=(255, 170, 80, 170))
        line = Image.alpha_composite(tg.filter(ImageFilter.GaussianBlur(12)), line)
        d = ImageDraw.Draw(line)
        d.ellipse([x - 13, y - 13, x + 13, y + 13], fill=(255, 236, 200, 255))
    img = Image.alpha_composite(img, line)
    img.save(os.path.join(OUT, name + ".png"))

EMBER = (240, 110, 50, 255)
# 표지: 세 작품 원형 이미지를 잇는 꼬리
make("tail_cover", (5.2, 7.7), [
    ((6.2, 5.2), (6.6, 2.6), (8.4, 2.55)),
    ((10.0, 2.5), (10.0, 3.6), (11.35, 3.6)),
    ((12.7, 3.6), (11.2, 5.75), (9.35, 5.75)),
    ((7.9, 5.75), (7.5, 6.9), (8.6, 7.0)),
], EMBER, width=8)
# 목차: 세로로 이어지는 꼬리
make("tail_toc", (1.35, 1.25), [
    ((0.85, 1.6), (1.85, 1.85), (1.35, 2.25)),
    ((0.8, 2.8), (1.9, 3.3), (1.35, 3.85)),
    ((0.8, 4.4), (1.9, 4.9), (1.35, 5.45)),
    ((0.9, 5.95), (1.8, 6.3), (1.45, 6.75)),
], EMBER, width=7)
# 챕터: 번호와 제목 사이를 가르는 흰 꼬리, 이미지 경계에서 불씨로 연결
make("tail_ch", (-0.2, 4.35), [
    ((1.1, 3.85), (2.3, 4.75), (3.4, 4.25)),
    ((3.8, 4.05), (4.15, 4.1), (4.4, 4.25)),
], (255, 255, 255, 255), width=5, glow=False)
# 엔딩
make("tail_end", (-0.2, 5.9), [
    ((2.5, 4.9), (4.5, 6.9), (6.67, 6.0)),
    ((7.6, 5.6), (8.2, 5.3), (8.9, 5.7)),
], EMBER, width=7)
print("ok")

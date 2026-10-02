# 트럼프 카드 뒷면 (I'm Not Joker 카드 모티프)
from PIL import Image, ImageDraw
import os
W, H = 600, 840
RED, DEEP, IVORY, GOLD = (179, 18, 46), (110, 8, 26), (243, 235, 221), (201, 164, 92)
im = Image.new("RGB", (W, H), IVORY)
d = ImageDraw.Draw(im)
d.rounded_rectangle([24, 24, W - 25, H - 25], radius=26, fill=DEEP)
d.rounded_rectangle([44, 44, W - 45, H - 45], radius=18, outline=GOLD, width=3)
step = 36
for k in range(-H, W + H, step):  # 마름모 격자
    d.line([(k, 50), (k + H, 50 + H)], fill=RED, width=3)
    d.line([(k, H - 50), (k + H, H - 50 - H)], fill=RED, width=3)
mask = Image.new("L", (W, H), 0)
ImageDraw.Draw(mask).rounded_rectangle([50, 50, W - 51, H - 51], radius=14, fill=255)
base = Image.new("RGB", (W, H), IVORY)
bd = ImageDraw.Draw(base)
bd.rounded_rectangle([24, 24, W - 25, H - 25], radius=26, fill=DEEP)
bd.rounded_rectangle([44, 44, W - 45, H - 45], radius=18, outline=GOLD, width=3)
base.paste(im, (0, 0), mask)
cx, cy = W // 2, H // 2
bd.ellipse([cx - 70, cy - 70, cx + 70, cy + 70], fill=DEEP, outline=GOLD, width=4)
bd.polygon([(cx, cy - 40), (cx + 28, cy), (cx, cy + 40), (cx - 28, cy)], fill=GOLD)
# 모서리를 배경색(투명 대신 검정 계열)으로 둥글게
out = Image.new("RGBA", (W, H), (0, 0, 0, 0))
cm = Image.new("L", (W, H), 0)
ImageDraw.Draw(cm).rounded_rectangle([0, 0, W - 1, H - 1], radius=34, fill=255)
out.paste(base, (0, 0), cm)
out.save(os.path.join(os.path.dirname(__file__), "img", "cardback.png"))
print("ok")

# 챕터 카드 앞면: 모서리 랭크/무늬 + 대표 이미지 (한 장의 PNG로 합성해 회전해도 어긋나지 않게)
from PIL import Image, ImageDraw, ImageFont
import os
D = os.path.join(os.path.dirname(__file__), "img")
W, H = 750, 1050
IVORY = (243, 235, 221)
SERIF = "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf"
SYM = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
cards = [("card1", "A", "♥", (179, 18, 46)), ("card2", "2", "♦", (224, 138, 46)), ("card3", "3", "♠", (62, 90, 158))]
fr, fs = ImageFont.truetype(SERIF, 92), ImageFont.truetype(SYM, 80)
for key, rank, suit, col in cards:
    face = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(face)
    d.rounded_rectangle([0, 0, W - 1, H - 1], radius=42, fill=IVORY, outline=(210, 198, 178), width=3)
    art = Image.open(os.path.join(D, key + ".jpg")).convert("RGB")
    bx, by, bw, bh = 120, 150, 510, 750  # 안쪽 그림 칸
    r = min(bw / art.width, bh / art.height)
    art = art.resize((round(art.width * r), round(art.height * r)), Image.LANCZOS)
    ax, ay = bx + (bw - art.width) // 2, by + (bh - art.height) // 2
    d.rectangle([ax - 8, ay - 8, ax + art.width + 7, ay + art.height + 7], fill=col)
    face.paste(art, (ax, ay))
    def corner(img):
        dd = ImageDraw.Draw(img)
        dd.text((60, 30), rank, font=fr, fill=col, anchor="ma")
        dd.text((60, 135), suit, font=fs, fill=col, anchor="ma")
    corner(face)
    face = face.rotate(180)
    corner(face)
    face = face.rotate(180)
    face.save(os.path.join(D, key + "_face.png"))
print("ok")

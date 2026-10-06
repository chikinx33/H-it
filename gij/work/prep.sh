#!/bin/bash
# 고일준(gij) 이미지 가공: 원본 비율 사본, 영상 포스터, 큐브 글로우, 흐린 배경, 크기 표
set -e
cd "$(dirname "$0")/.."
S=source/애니메이션; O=work/img
mkdir -p $O
cv(){ convert "$1" -resize "${3:-1600x1600>}" -background white -flatten -quality 88 "$O/$2.jpg"; }
cv "$S/라푼젤 이미지보드 test 3.png" board
cv "$S/라푼젤 캐릭터 디자인 테스트_캐릭터 시트01.png" sheet_r 2000x2000
cv "$S/라푼젤 캐릭터 디자인 테스트_표정01.png" face_r
cv "$S/루카스.png" sheet_lukas 2000x2000
cv "$S/루카스(Lukas).png" sheet_lukas1
cv "$S/ChatGPT Image 2026년 8월 11일 오후 01_14_58.png" lukas_close 1200x1200
cv "$S/35b451f6-89c3-442d-9357-531551eb7378.png" key_a
cv "$S/e78f3d52-c97c-4d36-a365-25146b91add9.png" key_b
ffmpeg -v error -y -ss 2.6 -i "$S/Animation Test.mp4" -frames:v 1 -q:v 3 $O/frame_anim.jpg
ffmpeg -v error -y -ss 0.3 -i "$S/Rotation Test.mp4" -frames:v 1 -q:v 3 $O/frame_rot.jpg
# 흐린 배경 1600x900
B(){ convert "$1" -resize 1600x900^ -gravity center -extent 1600x900 -blur 0x$2 -fill "#$3" -colorize $4 -quality 85 "$O/$5.jpg"; }
B "$S/라푼젤 이미지보드 test 3.png" 16 0E2225 62 bg_night
B "$S/라푼젤 이미지보드 test 3.png" 10 0E2225 40 bg_cover
B "$S/라푼젤 캐릭터 디자인 테스트_캐릭터 시트01.png" 14 EFE6D6 78 bg_parch
python3 work/make_assets.py
python3 - <<'PY'
import json, os, glob
from PIL import Image
d = {}
for f in sorted(glob.glob("work/img/*")):
    k, e = os.path.splitext(os.path.basename(f)); im = Image.open(f); d[k] = [im.width, im.height, e[1:]]
json.dump(d, open("work/dims.json", "w"), ensure_ascii=False, indent=0)
PY

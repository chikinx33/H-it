#!/bin/bash
# 박현진(phj) 이미지 가공: 원본 비율 사본, 티켓 간지용 크롭, 영상 포스터, 흐린 배경, 크기 표
set -e
cd "$(dirname "$0")/.."
S=source; O=work/img; A=$S/애니메이션_아이돌출발; L=$S/애니메이션_너무좋아; C=$S/캐릭터; D=$S/숏폼광고
mkdir -p $O
cv(){ convert "$1" -resize "${3:-1600x1600>}" -background white -flatten -quality 88 "$O/$2.jpg"; }
# 아이돌, 출발
cv "$A/박수아.png" sheet_sua 2000x2000; cv "$A/박수아트레이닝복.png" sheet_sua2 2000x2000
cv "$A/이유진.png" sheet_yujin; cv "$A/김진우.png" sheet_jinwoo 2000x2000
for c in 1-2 1-5 1-6 1-8 1-12 1-13 1-14 1-15 1-16 2-2B 2-3D 2-4 2-7 2-8 2-9 2-10 2-11 2-12 3-2 3-4A 3-4D 3-4E3 3-4F 3-9A 3-14 4-1B 4-2 4-4 4-5F 4-8 4-11; do
  k=i$(echo "$c" | tr -d '-' | tr 'A-Z' 'a-z'); cv "$A/컷 $c.png" "$k"; done
# 너무 좋아 1000% (9:16 세로)
for c in 1-1 1-2 1-3 1-4 1-5 1-6 1-7 1-8 1-9 2-1 2-2 2-3; do k=l$(echo "$c" | tr -d '-'); cv "$L/컷 $c.png" "$k" 1400x1400; done
cv "$L/게임01.png" game1 1400x1400; cv "$L/게임02.png" game2 1400x1400
cv "$L/유하나.png" sheet_hana; cv "$L/강민.png" sheet_min
# 일렁이는 뻄씨
cv "$C/일렁이는 뻄씨_브랜드_로고.png" logo1; cv "$C/일렁이는 뻄씨_브랜드_로고2(뾸랑뻄_캐릭터_버전).png" logo2
cv "$C/뻄.png" sheet_bbem; cv "$C/뚈랑뻄.png" sheet_ddol; cv "$C/일렁이는 뻄씨.png" emo_grid 1800x1800
cv "$C/일상툰260831.png" toon; cv "$C/모자_상세설명.png" hat_detail; cv "$C/스피커_상세설명.png" spk_detail; cv "$C/텀블러_상세설명.png" tum_detail
cv "$C/스피커_메인.png" spk_main
cv "$C/모자_광고_페이지.png" hat_page 900x2700; cv "$C/스피커_광고_페이지.png" spk_page 900x2700; cv "$C/텀블러_광고_페이지.png" tum_page 900x2700
for g in 냠냠:nyam 안녕:hi 바이:bye 버스:bus 물음표:q; do cp "$C/${g%%:*}.gif" "$O/emo_${g##*:}.gif"; done
for v in 영상:promo 단순_캐릭터_홍보_영상_1:promo1 단순_캐릭터_홍보_영상_2:promo2; do
  ffmpeg -v error -y -ss 2 -i "$C/${v%%:*}.mp4" -frames:v 1 -q:v 3 "$O/${v##*:}.jpg"; done
# 숏폼 광고
cv "$D/컨셉.png" magic_concept 1800x1800; cv "$D/확빠져_컨셉.png" slim_concept 1800x1800
cv "$D/확빠져_스토리보드.png" sb 2560x2560; cv "$D/살찐박복슬.png" char_before; cv "$D/살빠진박곱슬.png" char_after
# 흐린 배경 1600x900
B(){ convert "$1" -resize 1600x900^ -gravity center -extent 1600x900 -blur 0x$2 -fill "#$3" -colorize $4 -quality 85 "$O/$5.jpg"; }
B "$A/컷 4-4.png" 20 FFF6F1 84 bg_paper
B "$A/컷 4-4.png" 18 1E1233 66 bg_stage
B "$A/컷 4-2.png" 22 1E1233 70 bg_cover
B "$A/컷 4-1B.png" 16 1E1233 62 ch1_bg
B "$L/컷 1-9.png" 18 FFF4F6 80 bg_spring
B "$L/컷 1-5.png" 16 1E1233 72 ch2_bg
B "$C/일렁이는 뻄씨.png" 8 FFF8F0 70 bg_bbem
B "$O/promo.jpg" 14 1E1233 74 ch3_bg
B "$D/확빠져_컨셉.png" 18 FFF5F8 84 bg_ad
B "$D/컨셉.png" 14 1E1233 76 ch4_bg
python3 work/make_assets.py
# 이미지 크기 표 (build.js 비율 계산용)
python3 - <<'PY'
import json, os, glob
from PIL import Image
d = {}
for f in sorted(glob.glob("work/img/*")):
    k, e = os.path.splitext(os.path.basename(f)); im = Image.open(f); d[k] = [im.width, im.height, e[1:]]
json.dump(d, open("work/dims.json", "w"), ensure_ascii=False, indent=0)
PY

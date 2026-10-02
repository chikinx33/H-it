#!/bin/bash
# 김경수(kks) 이미지 가공: 원본 비율 사본, 연령 시트 얼굴 크롭, 콘티 컷 발췌, 원형 사본, 발광 링, 흐린 배경
set -e
cd "$(dirname "$0")/.."
S=source; O=work/img; A=$S/애니메이션; C=$S/캐릭터; D=$S/숏폼광고
mkdir -p $O
cv(){ convert "$1" -resize "${3:-1600x1600>}" -background white -flatten -quality 88 "$O/$2.jpg"; }
# 애니메이션 스틸 (1672x941)
cv "$A/씬01 컷01-5.png" s01_pov; cv "$A/씬01 컷02.png" s01_ward; cv "$A/씬01 컷04.png" s01_dad
cv "$A/씬02컷01.png" s02_write; cv "$A/씬02컷05.png" s02_thumb
cv "$A/씬03컷02.png" s03_diary; cv "$A/씬03컷03.png" s03_sleep
cv "$A/씬04 컷01.png" s04_play; cv "$A/씬04 컷03.png" s04_tv
cv "$A/씬05_컷02.png" s05_jeju; cv "$A/씬06_컷03.png" s06_photo
cv "$A/씬07_컷01.png" s07_sun; cv "$A/씬07_컷02.png" s07_lean
cv "$A/씬08_컷04.png" s08_grad; cv "$A/씬09_컷02.png" s09_walk; cv "$A/씬09_컷05.png" s09_mom; cv "$A/씬09_컷06.png" s09_back
cv "$A/씬10_컷02.png" s10_street
cv "$A/씬11_컷06.png" s11_apron; cv "$A/씬11_컷09.png" s11_table; cv "$A/씬11_컷13.png" s11_green; cv "$A/씬11_컷14.png" s11_green2
cv "$A/크리스마스 사진.png" xmas; cv "$A/봉안당 어머니.png" urn_mom
# 캐릭터 시트
cv "$A/갠지 캐릭터시트.png" ganji; cv "$A/갠지3.png" ganji3; cv "$A/갠지5.png" ganji5; cv "$A/갠지7.png" ganji7
cv "$A/중년 아버지.png" dad; cv "$A/중년 어머니.png" mom; cv "$A/토리 캐릭터시트.png" tori
# 경수 연령별 얼굴 클로즈업 (시트 오른쪽 영역)
i=1; for f in "$A"/0[1-8].*경수.png; do convert "$f" -crop 560x660+880+190 +repage -quality 90 "$O/age$i.jpg"; i=$((i+1)); done
# 화투 캐릭터 IP
cv "$C/1월 캐릭터시트.png" m01 2000x2000; cv "$C/2월 캐릭터 시트.png" m02; cv "$C/4월 캐릭터 시트.png" m04
cv "$C/7월 캐릭터 시트.png" m07; cv "$C/8월 캐릭터 시트.png" m08; cv "$C/10월 캐릭터 시트.png" m10
cv "$C/11월 캐릭터 시트.png" m11; cv "$C/12월 캐릭터 시트.png" m12
cv "$C/광.png" gwang 1000x1000; cv "$C/티셔츠.png" tshirt 1200x1200; cv "$C/일상툰 1화.png" toon 1024x1536
convert "$C/일상툰 1화.png" -crop 990x380+17+1150 +repage -quality 90 $O/toon_last.jpg
cp "$C/학~씨!.gif" $O/emo_hak.gif; cp "$C/아 어딘뎅.gif" $O/emo_bird.gif; cp "$C/멧돼지.gif" $O/emo_boar.gif; cp "$C/사슴.gif" $O/emo_deer.gif
# 숏폼 광고 스토리보드 (2550x1080) 컷 발췌
SB="$D/쇼츠광고-스토리보드8컷.png"
cv "$SB" sb 2550x2550
cr(){ convert "$SB" -crop "$1" +repage -quality 92 "$O/$2.jpg"; }
cr 576x326+56+110 sb11; cr 576x326+672+110 sb12; cr 576x326+1288+110 sb13; cr 576x326+1904+110 sb21
cr 576x326+56+598 sb22; cr 576x326+672+598 sb31; cr 576x326+1288+598 sb32; cr 576x326+1904+598 sb41
# 흐린 배경 1600x900
B(){ convert "$1" -resize 1600x900^ -gravity center -extent 1600x900 -blur 0x$2 -fill "#$3" -colorize $4 -quality 85 "$O/$5.jpg"; }
B "$A/씬07_컷01.png" 22 0F0E13 70 bg_void
B "$A/씬07_컷02.png" 14 0F0E13 58 ch1_bg
B "$A/씬11_컷09.png" 22 F2EADC 82 bg_album
B "$A/씬09_컷06.png" 22 0F0E13 72 bg_void2
B "$C/1월 캐릭터시트.png" 10 0F0E13 88 ch2_bg
B "$C/8월 캐릭터 시트.png" 6 F2EADC 70 bg_ip
B "$SB" 14 0F0E13 72 ch3_bg
B "$SB" 16 F2EADC 86 bg_sb
B "$A/씬01 컷01-5.png" 22 0F0E13 74 bg_cover
# 원형 사본 + 발광 링 + 영상 포스터
python3 work/make_rings.py
# 이미지 크기 표 (build.js의 비율 계산용)
python3 - <<'PY'
import json, os, glob
from PIL import Image
d = {}
for f in sorted(glob.glob("work/img/*")):
    k, e = os.path.splitext(os.path.basename(f)); im = Image.open(f); d[k] = [im.width, im.height, e[1:]]
json.dump(d, open("work/dims.json", "w"), ensure_ascii=False, indent=0)
PY

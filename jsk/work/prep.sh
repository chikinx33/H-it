#!/bin/bash
# 조성기(jsk) 이미지 가공: 원본 비율 유지 사본, 카드용 크롭, 흐린 배경
set -e
cd "$(dirname "$0")/.."
S=source; O=work/img; A=$S/애니메이션; C=$S/캐릭터; D=$S/숏폼광고; F=$S/워크플로우
cv(){ convert "$1" -resize "${3:-1600x1600>}" -background white -flatten -quality 88 "$O/$2.jpg"; }
# 1장 애니메이션 스틸
for f in "$A"/S#*; do b=$(basename "$f"); k=$(echo "${b%.*}" | sed 's/S#\([0-9]*\) C\([0-9]*\)/a\1_\2/'); cv "$f" "$k"; done
cv "$A/현준.png" hyunjun; cv "$A/강훈.png" kanghoon
# 워크플로우 캡처
cv "$F/클로드 프로젝트.png" wf_claude; cv "$F/클로드 립싱크영상 프롬프트.png" wf_prompt
cv "$F/영상생성.png" wf_gen1; cv "$F/영상생성2.png" wf_gen2; cv "$F/립싱크 영상생성 .png" wf_lipsync
cv "$F/디지털성형(클립스튜디오).png" wf_clip; cv "$F/프리미어편집.png" wf_premiere
# 2장 캐릭터 IP
cv "$C/조은탁.png" eun; cv "$C/이신.png" isin; cv "$C/저승사자.png" reaper_toon
cv "$C/오해소지프로필.png" sausage_prof 1600x2800; cv "$C/왔나.gif[0]" sausage_wanna
cv "$C/일상툰6컷 햄이되고싶어.png" hamtoon 1400x2200
cr(){ convert "$1" -crop "$2" +repage -quality 90 "$O/$3.jpg"; }
cr "$C/1페이지(스크롤웹툰).png" 690x1500+0+0 web1
cr "$C/1페이지(스크롤웹툰).png" 690x1400+0+2100 web2
cr "$C/3페이지(스크롤웹툰).png" 690x1420+0+2080 web3
cr "$C/끝페이지(스크롤웹툰).png" 690x1400+0+1550 web4
cr "$C/2페이지(스크롤웹툰).png" 690x660+0+3210 web_title
cr "$C/오해소지프로필.png" 675x665+110+1752 sausage_emo
# 3장 숏폼 광고
for f in "$D"/S#*; do b=$(basename "$f"); k=$(echo "${b%.*}" | sed 's/S#\([0-9]*\) C\([0-9]*\)/d\1_\2/'); cv "$f" "$k"; done
cv "$D/배이스 캐릭터시트.png" bays 2400x2400; cv "$D/저승사자캐릭터시트찐.png" reaper_ad
# 챕터 카드 그림 (카드 안쪽 비율 0.79)
cr "$A/S#12 C3.png" 853x1080+330+0 card1
cr "$C/3페이지(스크롤웹툰).png" 690x873+0+0 card2
cr "$D/S#2 C1.png" 768x972+0+150 card3
# 흐린 배경 1600x900
B(){ convert "$1" -resize 1600x900^ -gravity center -extent 1600x900 -blur 0x$2 -fill "#$3" -colorize $4 -quality 85 "$O/$5.jpg"; }
B "$A/S#12 C6.png" 10 0B0909 58 bg_cover
B "$A/S#12 C1.png" 20 0E0C0B 74 bg_noir
B "$A/S#7 C9.png" 20 0B0B0C 70 bg_alley
B "$A/S#12 C7.png" 6 0B0909 45 bg_end
B "$A/S#12 C6.png" 9 1A0508 52 ch1_bg
B "$C/1페이지(스크롤웹툰).png" 14 1E1408 62 ch2_bg
B "$D/S#1 C1.png" 12 070B1C 55 ch3_bg
B "$C/3페이지(스크롤웹툰).png" 18 F3EBDD 84 bg_ivory
B "$D/S#1 C2.png" 18 080C1A 70 bg_ad
B "$F/클로드 프로젝트.png" 16 0E0C0B 80 bg_wf
# 재생 버튼 포스터
convert "$A/S#12 C3.png" -resize 1600x900 -fill black -colorize 30 -fill "rgba(243,235,221,0.92)" -draw "circle 800,450 800,350" -fill "#B3122E" -draw "polygon 768,395 768,505 858,450" -quality 88 $O/poster_mv.jpg
echo prep done

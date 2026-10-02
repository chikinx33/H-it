#!/bin/bash
# 최영빈(chb) 이미지 가공: 원본 비율 사본, 스틸 패널 발췌, 사원증 사진 크롭, 흐린 배경
set -e
cd "$(dirname "$0")/.."
S=source; O=work/img; A=$S/애니메이션; C=$S/캐릭터; D=$S/숏폼광고
cv(){ convert "$1" -resize "${3:-1600x1600>}" -background white -flatten -quality 88 "$O/$2.jpg"; }
cv "$A/나신입_1화_5패널_턴어라운드.png" ch_na 2000x2000; cv "$A/강도한_1화_5패널_턴어라운드.png" ch_kang 2000x2000
cv "$A/박과장_1화_5패널_턴어라운드.png" ch_park 2000x2000; cv "$A/이대리_1화_5패널_턴어라운드.png" ch_lee 2000x2000; cv "$A/김인턴_1화_5패널_턴어라운드.png" ch_kim 2000x2000
cv "$A/홍보대행사_회의실.png" bg_agency_room; cv "$A/강엔터_회의실.png" bg_kang_room
cv "$A/최영빈_애니메이션_스틸컷.png" still_all 1024x1536
# 스틸컷 패널 (S#2 C1~C11)
P="$A/최영빈_애니메이션_스틸컷.png"
cr(){ convert "$1" -crop "$2" +repage -quality 92 "$O/$3.jpg"; }
cr "$P" 429x276+3+3 c01; cr "$P" 578x276+443+3 c02; cr "$P" 332x258+3+291 c03; cr "$P" 328x258+347+291 c04
cr "$P" 334x258+687+291 c05; cr "$P" 560x181+3+561 c06; cr "$P" 446x181+575+561 c07; cr "$P" 584x212+3+755 c08
cr "$P" 421x212+600+755 c09; cr "$P" 1018x216+3+979 c10; cr "$P" 1018x326+3+1207 c11
# 10MOO
for f in "$C"/10moo_*.jpg; do b=$(basename "$f" .jpg); cv "$f" "${b,,}" 1200x1200; done
cp "$C/10MOO_clasping_transparent_360x360.gif" $O/emo_clasp.gif; cp "$C/10MOO_sheep_hood_tilt_transparent_360x360.gif" $O/emo_tilt.gif
cp "$C/10MOO_wink_OK_transparent_360x360.gif" $O/emo_ok.gif; cp "$C/10MOO_angry_snap_transparent_360x360.gif" $O/emo_angry.gif
# 광고 스토리보드
cv "$D/최영빈_광고_스토리보드.jpg" sb 2550x2550
SB="$D/최영빈_광고_스토리보드.jpg"
cr "$SB" 560x312+1914+620 sb_c8; cr "$SB" 560x312+56+620 sb_c5; cr "$SB" 560x312+1331+620 sb_c7
# 사원증 사진 (3:4)
cr "$A/나신입_1화_5패널_턴어라운드.png" 600x800+1835+100 id1
convert "$C/10moo_04master.jpg" -crop 1200x1600+424+200 +repage -resize 600x800 -quality 90 $O/id2.jpg
cr "$SB" 240x320+1405+606 id3
convert "$C/10moo_01master.jpg" -crop 1200x1600+424+200 +repage -resize 600x800 -quality 90 $O/id0.jpg
# 영상 포스터
convert $O/c11.jpg -resize 1600x -gravity center -extent 1600x900 -background "#1F2430" -fill "#1F2430" -colorize 25 -fill "rgba(255,255,255,0.92)" -draw "circle 800,450 800,355" -fill "#2B59C3" -draw "polygon 770,398 770,502 858,450" -quality 88 $O/poster_ep1.jpg
convert $O/sb_c8.jpg -resize 1600x900^ -gravity center -extent 1600x900 -fill black -colorize 20 -fill "rgba(255,255,255,0.92)" -draw "circle 800,450 800,355" -fill "#E8892B" -draw "polygon 770,398 770,502 858,450" -quality 88 $O/poster_ad.jpg
# 흐린 배경 1600x900
B(){ convert "$1" -resize 1600x900^ -gravity center -extent 1600x900 -blur 0x$2 -fill "#$3" -colorize $4 -quality 85 "$O/$5.jpg"; }
B "$A/홍보대행사_회의실.png" 18 F7F5EF 80 bg_office
B "$A/강엔터_회의실.png" 16 1E222C 62 bg_night
B "$A/강엔터_회의실.png" 8 1E222C 40 ch1_bg
B "$C/10moo_06master.jpg" 22 EEF3FB 84 bg_moo
B "$C/10moo_06master.jpg" 14 1D3E8A 55 ch2_bg
B "$SB" 16 FFF6EC 86 bg_sb
B "$SB" 10 6B3A12 55 ch3_bg
B "$A/홍보대행사_회의실.png" 8 1E222C 45 bg_cover
echo prep done

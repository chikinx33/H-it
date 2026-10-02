#!/bin/bash
# 임현(lh) 이미지 가공: 원본 비율 유지 사본, 플레이어용 16:9 크롭, 흐린 배경
set -e
cd "$(dirname "$0")/.."
S=source; O=work/img; A=$S/애니메이션; C=$S/캐릭터; D=$S/숏폼광고; F=$S/워크플로우
cv(){ convert "$1" -resize "${3:-1600x1600>}" -background white -flatten -quality 88 "$O/$2.jpg"; }
# 1장 애니메이션 스틸: S#1_C#2a -> a1_2a
for f in "$A"/S#*; do b=$(basename "$f"); k=$(echo "${b%.*}" | sed 's/S#\([0-9]*\)_C#\(.*\)/a\1_\2/' | tr 'A-Z' 'a-z'); cv "$f" "$k"; done
cv "$A/철수(실직자).png" cs_out; cv "$A/철수(구직자).png" cs_job
cv "$A/영희(버튜버).png" yh_vt 2400x2400; cv "$A/영희(아이돌).png" yh_idol 2400x2400
# 워크플로우
cv "$F/클로드.png" wf_claude; cv "$F/프리미어.png" wf_premiere
cv "$F/스크린샷 2026-10-02 123044.png" wf_clip1; cv "$F/스크린샷 2026-10-02 123246.png" wf_clip2; cv "$F/스크린샷 2026-10-02 123510.png" wf_clip3
cp "$F/애니메이션 턴어라운드.gif" $O/gif_turn.gif; cp "$F/애니메이션 활용실습.gif" $O/gif_practice.gif
# 2장 TENSHI
cv "$C/캐릭터 시트 최종본.png" t_sheet; cv "$C/캐릭터 수정 삼면도.png" t_sd; cv "$C/6등신비 캐릭터.png" t_six; cv "$C/캐릭터 로고.png" t_logo
cv "$C/넨도로이드.png" g_nendo; cv "$C/누이구루미.png" g_nui; cv "$C/푸치슈.png" g_puchi
cv "$C/일상툰예시1-1.png" toon1; cv "$C/일상툰예시1-2.png" toon2
i=0; for n in 이몸등장:등장 최고:최고 사랑해:사랑해-최종본 축하해:"축하해 최종본" 파이팅:화이팅최종본 건배:건배 미안해:사과 뿌에엥:뿌엥 싫은데:거절-최종본 날속인거니:기만 귀찮아:나태-최종본 졸려:피곤-최종본 힘들어:후회최종본 쉬면서해:휴식최종본; do
  i=$((i+1)); src="${n#*:}"; cp "$C/$src.gif" "$O/emo$(printf %02d $i).gif"; done
# 3장 학원물 캐릭터
cv "$D/S#1_C#1.png" s_class; cv "$D/남자 캐릭터 삼면도.png" s_m3; cv "$D/여자 캐릭터 삼면도.png" s_f3
cv "$D/남자 캐릭터 표정.png" s_mface; cv "$D/여자 캐릭터 표정.png" s_fface
cv "$D/남주 시트 정면도.png" s_mfront; cv "$D/여주 시트 정면도.png" s_ffront; cv "$D/좌측면도.png" s_mside
# 플레이어 화면용 16:9 크롭
convert "$A/S#3_C#4.png" -resize 1600x900^ -gravity center -extent 1600x900 -quality 90 $O/play1.jpg
convert "$C/넨도로이드.png" -gravity center -crop 1536x864+0+0 +repage -resize 1600x900 -quality 90 $O/play2.jpg
convert "$D/S#1_C#1.png" -resize 1600x900^ -gravity center -extent 1600x900 -quality 90 $O/play3.jpg
convert "$A/S#3_C#3.png" -resize 1600x900^ -gravity center -extent 1600x900 -quality 90 $O/play0.jpg
# 영상 포스터
convert $O/play1.jpg -fill "#15132B" -colorize 30 -fill "rgba(255,255,255,0.92)" -draw "circle 800,450 800,355" -fill "#FF5C9E" -draw "polygon 770,398 770,502 858,450" -quality 88 $O/poster_mv.jpg
convert $O/play2.jpg -fill "#0F3B66" -colorize 25 -fill "rgba(255,255,255,0.92)" -draw "circle 800,450 800,355" -fill "#1F8FE6" -draw "polygon 770,398 770,502 858,450" -quality 88 $O/poster_tenshi.jpg
# 흐린 배경 1600x900
B(){ convert "$1" -resize 1600x900^ -gravity center -extent 1600x900 -blur 0x$2 -fill "#$3" -colorize $4 -quality 85 "$O/$5.jpg"; }
B "$A/S#3_C#3.png" 12 15132B 55 bg_cover
B "$A/S#2_C#8a.png" 18 15132B 62 bg_room
B "$A/S#2_C#1.png" 20 15132B 72 bg_stream
B "$A/S#4_C#1.png" 18 FFF7EE 82 bg_morning
B "$A/S#3_C#4.png" 10 15132B 50 ch1_bg
B "$C/캐릭터 시트 최종본.png" 16 EEF8FF 86 bg_sky
B "$C/넨도로이드.png" 12 0E3560 50 ch2_bg
B "$D/S#1_C#1.png" 16 141B2E 55 bg_rain
B "$D/S#1_C#1.png" 10 141B2E 45 ch3_bg
B "$A/S#5_C#6_4.png" 8 FFF7EE 55 bg_end
echo prep done

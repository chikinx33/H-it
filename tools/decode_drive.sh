#!/bin/bash
# 사용: tools/decode_drive.sh <출력폴더> <최소 타임스탬프(ms)>
# Drive 커넥터 다운로드 결과(JSON base64)를 원래 파일명으로 복원
OUT="$1"; MIN="$2"; mkdir -p "$OUT"
D=/root/.claude/projects/-home-user-H-it/22044d0b-fd83-5b89-8123-76ac982323f9/tool-results
for f in $D/mcp-Google_Drive-download_file_content-*.txt; do
  ts=$(basename "$f" .txt | sed 's/.*-//'); [ "$ts" -lt "$MIN" ] && continue
  m=$(jq -r .mimeType "$f"); case "$m" in image/png) e=png;; image/jpeg) e=jpg;; video/mp4) e=mp4;; *) continue;; esac
  t=$(jq -r .title "$f" | sed 's/ *$//'); base="${t%.*}"; dest="$OUT/$base.$e"
  [ -f "$dest" ] && continue
  jq -r .content "$f" | base64 -d > "$dest" && echo "$base.$e $(identify -format '%wx%h' "$dest" 2>/dev/null)"
done

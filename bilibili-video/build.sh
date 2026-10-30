#!/usr/bin/env bash
# Full pipeline: fonts -> frames -> music -> mux. Usage: ./build.sh [out_dir] [workers]
set -euo pipefail
cd "$(dirname "$0")"
OUT="${1:-out}"; WORKERS="${2:-4}"
mkdir -p "$OUT" assets/fonts
fetch() { [ -s "assets/fonts/$1" ] || curl -sSfL -o "assets/fonts/$1" "$2"; }
fetch NotoSerifSC-Regular.ttf https://fonts.gstatic.com/s/notoserifsc/v36/H4cyBXePl9DZ0Xe7gG9cyOj7uK2-n-D2rd4FY7SCqyWv.ttf
fetch NotoSerifSC-Bold.ttf    https://fonts.gstatic.com/s/notoserifsc/v36/H4cyBXePl9DZ0Xe7gG9cyOj7uK2-n-D2rd4FY7RlrCWv.ttf
fetch NotoSerifSC-Black.ttf   https://fonts.gstatic.com/s/notoserifsc/v36/H4cyBXePl9DZ0Xe7gG9cyOj7uK2-n-D2rd4FY7QrrCWv.ttf
fetch NotoSansSC-Light.ttf    https://fonts.gstatic.com/s/notosanssc/v41/k3kCo84MPvpLmixcA63oeAL7Iqp5IZJF9bmaG4HFnYw.ttf
[ -d node_modules/mathjax ] || npm install --no-audit --no-fund
node render.js "$OUT" "$WORKERS"
python3 music.py "$OUT/timeline.json" "$OUT/music.wav"
python3 make_meta.py "$OUT/timeline.json" "$OUT"
(cd "$OUT" && ffmpeg -y -loglevel error -f concat -safe 0 -i chunks.txt -c copy video_only.mp4)
ffmpeg -y -loglevel error -i "$OUT/video_only.mp4" -i "$OUT/music.wav" \
  -map 0:v -map 1:a -c:v copy -af "loudnorm=I=-16:TP=-1.5:LRA=11" -c:a aac -b:a 320k -ar 48000 \
  -shortest -movflags +faststart "$OUT/人生的分岔.mp4"
ffmpeg -y -loglevel error -i "$OUT/cover.png" -q:v 2 "$OUT/cover.jpg"
echo "done: $OUT/人生的分岔.mp4"

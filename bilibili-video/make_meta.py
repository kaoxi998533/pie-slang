#!/usr/bin/env python3
"""timeline.json -> subtitles.srt (CC字幕) + chapter list for the B站 description."""
import json
import sys

tl = json.load(open(sys.argv[1]))
out_dir = sys.argv[2]


def ts(t, sep=","):
    h, m, s = int(t // 3600), int(t % 3600 // 60), t % 60
    return f"{h:02d}:{m:02d}:{int(s):02d}{sep}{int(round((s % 1) * 1000)) % 1000:03d}"


with open(f"{out_dir}/subtitles.srt", "w", encoding="utf-8") as f:
    for i, b in enumerate(tl["beats"], 1):
        text = b["text"].replace("*", "")
        if b.get("sub"):
            text += "\n" + b["sub"]
        f.write(f"{i}\n{ts(b['start'])} --> {ts(b['end'])}\n{text}\n\n")

chapters = [("00:00", "开场：小径分岔的花园")]
for s in tl["scenes"]:
    if s["card"]:
        m, sec = int(s["start"] // 60), int(s["start"] % 60)
        chapters.append((f"{m:02d}:{sec:02d}", s["chapter"]))
with open(f"{out_dir}/chapters.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(f"{t} {name}" for t, name in chapters) + "\n")
print(open(f"{out_dir}/chapters.txt").read())
print(f"duration {tl['duration']:.1f}s, {len(tl['beats'])} subtitle lines")

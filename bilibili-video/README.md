# 人生的分岔，可以数得清吗？

这是一部约 12 分半钟的文字动画短片，可以直接上传 B站（1080p30，H.264 + AAC）。片子分十章，从集合论里的无穷讲起，经过 Bekenstein 界和紧致性，最后落到"绝症晚期的人，未来的选择会不会瞬间坍缩"这个问题。

- 画面：Canvas 2D，每一帧都只由时间 t 决定。用 Playwright（Chromium）逐帧截图，再交给 ffmpeg 编码。
- 公式：MathJax 先把 TeX 渲染成 SVG，再画到 Canvas 上。
- 字体：Noto Serif SC 和 Noto Sans SC（SIL OFL），构建时从 Google Fonts 下载。
- 配乐：`music.py` 用 numpy 现场合成，包括钢琴、铺底音色、钟声和混响。它会按时间线上每一章的情绪切换和声，不使用任何外部素材。

## 文件

| 文件 | 作用 |
|---|---|
| `scenes.js` | 文案（分镜）和每个场景的画面 |
| `engine.js` | 时间轴、逐字出现的字幕、分岔树、辉光和合成 |
| `render.js` | 并行渲染全部帧，并分段编码 |
| `preview.js` | 渲染单帧，用于检查画面，例如 `node preview.js out s:s5b:b1+3 cover` |
| `music.py` | 根据 `timeline.json` 生成配乐 |
| `make_meta.py` | 生成 `subtitles.srt` 和章节时间戳 |
| `build.sh` | 一键构建 |
| `投稿信息.md` | 标题、简介、标签和分区建议 |

## 构建

```bash
./build.sh out 4     # 输出 out/人生的分岔.mp4、cover.jpg、subtitles.srt、chapters.txt
```

需要 Node 18 及以上、Playwright 自带的 Chromium、ffmpeg，以及装有 numpy 和 scipy 的 Python 3。

如果要改文案，编辑 `scenes.js` 里各个 `beats` 的 `text` 就行。每一句的停留时间按字数自动计算，参数是 `engine.js` 里的 `READ_BASE` 和 `READ_CPS`。

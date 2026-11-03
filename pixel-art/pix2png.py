#!/usr/bin/env python3
"""把 .pix 文本图片编码解码成 PNG。

只用 Python 标准库（zlib + struct），不依赖网络或第三方包。

用法:
    python3 pix2png.py apple.pix apple.png [放大倍数，默认 16]
"""
import struct
import sys
import zlib


def parse_pix(text):
    width = height = None
    palette = {}
    rows = []
    section = None
    for raw in text.splitlines():
        line = raw.rstrip("\n")
        if section != "data":
            stripped = line.strip()
            if not stripped or stripped.startswith("#"):
                continue
            if stripped.startswith("size"):
                _, w, h = stripped.split()[:3]
                width, height = int(w), int(h)
            elif stripped == "palette":
                section = "palette"
            elif stripped == "data":
                section = "data"
            elif section == "palette":
                parts = stripped.split()
                ch, hexcolor = parts[0], parts[1]
                if len(ch) != 1:
                    raise ValueError(f"调色板字符必须是单个字符: {ch!r}")
                palette[ch] = tuple(int(hexcolor[i:i + 2], 16) for i in (0, 2, 4))
        else:
            if line.strip():
                rows.append(line.strip())

    if width is None:
        raise ValueError("缺少 size 行")
    if len(rows) != height:
        raise ValueError(f"data 行数为 {len(rows)}，应为 {height}")
    for y, row in enumerate(rows):
        if len(row) != width:
            raise ValueError(f"第 {y} 行长度为 {len(row)}，应为 {width}: {row!r}")
        for x, ch in enumerate(row):
            if ch not in palette:
                raise ValueError(f"({x},{y}) 处的字符 {ch!r} 不在调色板中")
    return width, height, palette, rows


def write_png(path, width, height, pixel_rows):
    """pixel_rows: 每行是 bytes（RGB，每像素 3 字节）。"""
    def chunk(tag, data):
        body = tag + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)

    raw = b"".join(b"\x00" + row for row in pixel_rows)  # 每行前加过滤类型 0
    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)  # 8 位 RGB
    with open(path, "wb") as f:
        f.write(b"\x89PNG\r\n\x1a\n")
        f.write(chunk(b"IHDR", ihdr))
        f.write(chunk(b"IDAT", zlib.compress(raw, 9)))
        f.write(chunk(b"IEND", b""))


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    src, dst = sys.argv[1], sys.argv[2]
    scale = int(sys.argv[3]) if len(sys.argv) > 3 else 16

    with open(src, encoding="utf-8") as f:
        width, height, palette, rows = parse_pix(f.read())

    out_rows = []
    for row in rows:
        line = b"".join(bytes(palette[ch]) * scale for ch in row)  # 最近邻放大
        out_rows.extend([line] * scale)

    write_png(dst, width * scale, height * scale, out_rows)
    print(f"{src}: {width}x{height} 像素, {len(palette)} 种颜色 -> {dst} ({width * scale}x{height * scale})")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""room.c 渲染结果的后期处理：降噪 -> 曝光 -> 泛光 -> ACES 色调映射 -> 暗角 -> sRGB PNG。

用法: python3 post.py render.bin out.png [--exposure E] [--denoise N] [--no-bloom]
只依赖 numpy 和 Pillow（本地已安装，不需要网络）。
"""
import argparse

import numpy as np
from PIL import Image


def load(path):
    with open(path, "rb") as f:
        w, h = np.frombuffer(f.read(8), dtype=np.int32)
        n = w * h
        data = np.frombuffer(f.read(), dtype=np.float32)
    col = data[: 3 * n].reshape(h, w, 3).astype(np.float64)
    alb = data[3 * n: 6 * n].reshape(h, w, 3).astype(np.float64)
    nrm = data[6 * n: 9 * n].reshape(h, w, 3).astype(np.float64)
    dep = data[9 * n: 10 * n].reshape(h, w).astype(np.float64)
    return col, alb, nrm, dep


def shift(a, dy, dx):
    """带边缘夹取的平移（out[y,x] = a[y+dy, x+dx]）。"""
    h, w = a.shape[:2]
    ys = np.clip(np.arange(h) + dy, 0, h - 1)
    xs = np.clip(np.arange(w) + dx, 0, w - 1)
    return a[ys][:, xs]


def lum(c):
    return 0.2126 * c[..., 0] + 0.7152 * c[..., 1] + 0.0722 * c[..., 2]


def atrous_denoise(col, alb, nrm, dep, iterations=4, sigma_l=4.0):
    """边缘保持的 à-trous 小波滤波（只滤光照，纹理通过反照率解调保留）。"""
    a = np.maximum(alb, 0.03)
    illum = col / a
    nl = np.linalg.norm(nrm, axis=2, keepdims=True)
    n = nrm / np.maximum(nl, 1e-6)
    kernel = [1 / 16, 1 / 4, 3 / 8, 1 / 4, 1 / 16]

    # 局部方差估计（5x5）用作亮度边缘阈值
    L = lum(illum)
    m1 = np.zeros_like(L)
    m2 = np.zeros_like(L)
    cnt = 0
    for dy in range(-2, 3):
        for dx in range(-2, 3):
            s = shift(L, dy, dx)
            m1 += s
            m2 += s * s
            cnt += 1
    var = np.maximum(m2 / cnt - (m1 / cnt) ** 2, 0)

    cur = illum
    for it in range(iterations):
        step = 1 << it
        Lc = lum(cur)
        sd = np.sqrt(var) + 1e-4
        acc = np.zeros_like(cur)
        wsum = np.zeros(cur.shape[:2])
        for j, ky in enumerate(kernel):
            for i, kx in enumerate(kernel):
                dy, dx = (j - 2) * step, (i - 2) * step
                cq = shift(cur, dy, dx)
                nq = shift(n, dy, dx)
                zq = shift(dep, dy, dx)
                aq = shift(a, dy, dx)
                w_n = np.maximum(np.sum(n * nq, axis=2), 0) ** 64
                w_z = np.exp(-np.abs(dep - zq) / (0.012 * dep * (abs(dy) + abs(dx) + 1) + 1e-3))
                w_a = np.exp(-np.sum((a - aq) ** 2, axis=2) / 0.004)
                w_l = np.exp(-np.abs(Lc - lum(cq)) / (sigma_l * sd))
                w = ky * kx * w_n * w_z * w_a * w_l
                acc += cq * w[..., None]
                wsum += w
        cur = acc / wsum[..., None]
        var = var / 4.0
    return cur * a


def box_blur(img, r):
    """可分离盒式模糊（积分图），重复 3 次近似高斯。"""
    out = img
    for _ in range(3):
        for axis in (0, 1):
            pad = [(0, 0)] * out.ndim
            pad[axis] = (r + 1, r)
            p = np.pad(out, pad, mode="edge")
            c = np.cumsum(p, axis=axis)
            hi = np.take(c, np.arange(2 * r + 1, c.shape[axis]), axis=axis)
            lo = np.take(c, np.arange(0, c.shape[axis] - 2 * r - 1), axis=axis)
            out = (hi - lo) / (2 * r + 1)
    return out


def aces(x):
    a, b, c, d, e = 2.51, 0.03, 2.43, 0.59, 0.14
    return np.clip((x * (a * x + b)) / (x * (c * x + d) + e), 0, 1)


def srgb(x):
    return np.where(x <= 0.0031308, 12.92 * x, 1.055 * np.power(np.maximum(x, 0), 1 / 2.4) - 0.055)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("dst")
    ap.add_argument("--exposure", type=float, default=None)
    ap.add_argument("--denoise", type=int, default=4, help="à-trous 迭代次数，0 关闭")
    ap.add_argument("--no-bloom", action="store_true")
    args = ap.parse_args()

    col, alb, nrm, dep = load(args.src)
    h, w = col.shape[:2]
    if args.denoise > 0:
        col = atrous_denoise(col, alb, nrm, dep, iterations=args.denoise)

    interior = dep < 49
    L = lum(col)
    exposure = args.exposure if args.exposure else 0.20 / np.percentile(L[interior], 50)
    print(f"{w}x{h}  median interior luminance={np.percentile(L[interior], 50):.4f}  exposure={exposure:.3f}")
    img = col * exposure

    # 轻微暖色白平衡
    img *= np.array([1.03, 1.0, 0.95])

    if not args.no_bloom:
        bright = np.maximum(img - 1.2, 0)
        bloom = 0.5 * box_blur(bright, max(1, w // 400)) + 0.35 * box_blur(bright, max(2, w // 120)) \
            + 0.25 * box_blur(bright, max(4, w // 40))
        img = img + 0.08 * bloom

    yy, xx = np.mgrid[0:h, 0:w]
    r2 = ((xx - w / 2) / (w / 2)) ** 2 + ((yy - h / 2) / (h / 2)) ** 2
    img *= (1 - 0.18 * r2)[..., None]

    out = srgb(aces(img))
    # 极轻的胶片颗粒，避免大面积色带
    rng = np.random.default_rng(7)
    out = out + rng.normal(0, 0.004, out.shape)
    Image.fromarray((np.clip(out, 0, 1) * 255 + 0.5).astype(np.uint8)).save(args.dst)
    print("saved", args.dst)


if __name__ == "__main__":
    main()

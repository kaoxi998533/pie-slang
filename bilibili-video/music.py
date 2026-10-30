#!/usr/bin/env python3
"""Procedural ambient score that follows the video's timeline.

usage: music.py timeline.json out.wav

Moods come from the scenes (wonder/think/heavy/warm/end); chapter cards
get a soft bell, the "collapse" moments get a low swell. Everything is
synthesized with numpy (additive piano, detuned sine pads, convolution
reverb), seeded so the result is reproducible.
"""
import json
import sys

import numpy as np
from scipy.signal import fftconvolve, lfilter

SR = 44100
tl = json.load(open(sys.argv[1]))
OUT = sys.argv[2]
DUR = tl["duration"] + 0.5
N = int(DUR * SR)
dry = np.zeros((2, N))
send = np.zeros((2, N))
rng = np.random.default_rng(7)


def midi(m):
    return 440.0 * 2 ** ((m - 69) / 12)


def place(buf, start, sig, pan=0.0, gain=1.0):
    i = int(round(start * SR))
    if i >= N or gain == 0:
        return
    if i < 0:
        sig, i = sig[-i:], 0
    n = min(len(sig), N - i)
    l, r = np.cos((pan + 1) * np.pi / 4), np.sin((pan + 1) * np.pi / 4)
    buf[0, i:i + n] += sig[:n] * gain * l
    buf[1, i:i + n] += sig[:n] * gain * r


def piano(m, vel):
    f0 = midi(m)
    T = float(np.clip(9.0 * (220.0 / f0) ** 0.45, 2.5, 9.0))
    t = np.arange(int(T * SR)) / SR
    sig = np.zeros_like(t)
    for k in range(1, 10):
        fk = k * f0 * np.sqrt(1 + 0.0003 * k * k)
        if fk > 9000:
            break
        amp = (1.0 / k ** 1.6) * (0.5 + vel) ** (0.35 * (k - 1))
        tau = (T / (1.0 + 0.7 * (k - 1))) / 6.9
        ph = rng.uniform(0, 2 * np.pi)
        for det in (-0.7, 0.7):
            sig += 0.5 * amp * np.sin(2 * np.pi * fk * 2 ** (det / 1200) * t + ph) * np.exp(-t / tau)
    sig *= np.minimum(1.0, t / 0.005) * (0.55 * np.exp(-t / 0.5) + 0.45)
    return sig * vel * 0.32


def pad(m, dur, att=3.0, rel=4.0, bright=0.22):
    f0 = midi(m)
    t = np.arange(int((dur + rel) * SR)) / SR
    sig = np.zeros_like(t)
    for d in (-0.0035, 0.0, 0.004):
        sig += np.sin(2 * np.pi * f0 * (1 + d) * t + rng.uniform(0, 6.28))
    sig = sig / 3 + bright * np.sin(2 * np.pi * 2 * f0 * t) + 0.06 * np.sin(2 * np.pi * 3 * f0 * t)
    env = np.minimum(1.0, t / att) ** 2
    env *= np.where(t > dur, np.exp(-(t - dur) / (rel / 3.5)), 1.0)
    env *= 1 + 0.18 * np.sin(2 * np.pi * rng.uniform(0.05, 0.11) * t + rng.uniform(0, 6.28))
    return sig * env


def bell(f0=698.46):
    t = np.arange(int(7 * SR)) / SR
    parts = [(1.0, 1.0, 3.2), (2.0, 0.45, 2.2), (2.76, 0.35, 1.6), (4.07, 0.18, 1.1), (5.4, 0.12, 0.8)]
    sig = sum(a * np.sin(2 * np.pi * f0 * r * t) * np.exp(-t / d) for r, a, d in parts)
    return sig * np.minimum(1.0, t / 0.003) * 0.16


def low_swell():
    t = np.arange(int(6 * SR)) / SR
    f = 44 + 20 * np.exp(-t / 0.6)
    tone = np.sin(2 * np.pi * np.cumsum(f) / SR) * np.exp(-t / 2.2) * np.minimum(1.0, t / 0.05)
    noise = lfilter([0.02], [1, -0.98], rng.standard_normal(len(t)))
    noise *= np.sin(np.pi * np.clip(t / 5.5, 0, 1)) ** 2
    return 0.45 * tone + 1.6 * noise


# ---------- harmony ----------
CH = {  # bass, chord tones (for pads and arpeggios)
    "Dm": (38, [57, 62, 64, 65, 69]),
    "Bb": (46, [57, 62, 65, 69, 70]),
    "F": (41, [57, 60, 64, 65, 69]),
    "C": (36, [55, 60, 62, 64, 67]),
    "Gm": (43, [55, 58, 62, 65, 67]),
    "A": (45, [57, 61, 64, 69, 64]),
}
MOOD = {  # progression, seconds per chord, note interval, play prob, octave, velocity, pad gain
    "wonder": (["Dm", "Bb", "F", "C"], 8.0, 1.5, 0.75, 12, (0.30, 0.50), 0.050),
    "think": (["Dm", "Bb", "F", "C"], 8.0, 1.0, 0.70, 12, (0.30, 0.48), 0.045),
    "heavy": (["Dm", "Gm", "Bb", "A"], 10.0, 2.0, 0.55, 0, (0.28, 0.42), 0.055),
    "warm": (["F", "C", "Dm", "Bb"], 8.0, 0.8, 0.70, 12, (0.30, 0.50), 0.050),
    "end": (["Bb", "C", "Dm", "F"], 8.0, 1.1, 0.65, 12, (0.30, 0.46), 0.050),
}
PATTERN = [0, 2, 1, 3, 2, 4, 3, 1]

sections = []
for s in tl["scenes"]:
    if sections and sections[-1][0] == s["mood"]:
        sections[-1][2] = s["start"] + s["dur"]
    else:
        sections.append([s["mood"], s["start"], s["start"] + s["dur"]])

accents = tl["accents"]
final_t = max(a["t"] for a in accents if a["kind"] == "bell")

events = []  # (start, dur, chord, mood)
for mood, a, b in sections:
    prog, cd = MOOD[mood][0], MOOD[mood][1]
    if mood == "end":
        # land on the tonic F exactly at the last line, then let it ring to the end
        k = int(np.ceil((final_t - a) / cd))
        t0 = final_t - k * cd
        for i in range(k):
            events.append((max(a, t0 + i * cd), t0 + (i + 1) * cd - max(a, t0 + i * cd), prog[(i - k) % 4], mood))
        events.append((final_t, b - final_t, "F", "final"))
    else:
        t, i = a, 0
        while t < b - 0.5:
            d = min(cd, b - t)
            events.append((t, d, prog[i % len(prog)], mood)); t += cd; i += 1

for start, dur, name, mood in events:
    bass, tones = CH[name]
    m = MOOD["end" if mood == "final" else mood]
    pg = m[6]
    rel = 6.0 if mood == "final" else 4.0
    place(dry, start, pad(bass, dur + 0.6, rel=rel, bright=0.3), 0.0, pg * 1.1)
    for j, n in enumerate(tones[:3]):
        place(dry, start, pad(n, dur + 0.6, rel=rel), [-0.45, 0.1, 0.5][j], pg * 0.55)
    if mood == "heavy":
        place(dry, start, pad(bass - 12, dur + 0.6, att=4.0, bright=0.4), 0.0, pg * 0.9)
    # arpeggiated piano
    if mood == "final":
        for j, (dt, n) in enumerate([(0.0, 53), (0.9, 57), (1.8, 60), (2.7, 64), (4.4, 65), (6.5, 72)]):
            place(dry, start + dt, piano(n + 12, 0.42 - j * 0.03), (j % 3 - 1) * 0.35)
            place(send, start + dt, piano(n + 12, 0.42 - j * 0.03), (j % 3 - 1) * 0.35)
        continue
    _, _, step, prob, octv, (v0, v1), _ = m
    idx = int(rng.integers(0, 4))
    t = start + rng.uniform(0.0, 0.3)
    while t < start + dur - 0.2:
        if rng.random() < prob:
            n = tones[PATTERN[idx % len(PATTERN)] % len(tones)] + octv
            vel = rng.uniform(v0, v1)
            sig = piano(n, vel)
            pan = float(np.clip((n - 72) / 14, -0.7, 0.7))
            place(dry, t, sig, pan); place(send, t, sig, pan, 1.0)
        idx += 1
        t += step * rng.choice([1.0, 1.0, 1.0, 0.5, 1.5])

for a in accents:
    if a["kind"] == "bell":
        b = bell(698.46 if a["t"] < final_t else 523.25)
        place(dry, a["t"], b, 0.15, 0.22); place(send, a["t"], b, 0.15, 0.5)
    elif a["kind"] == "low":
        s = low_swell(); place(dry, a["t"], s, 0.0, 0.45); place(send, a["t"], s, 0.0, 0.3)

# ---------- reverb ----------
def make_ir(rt, seed):
    r = np.random.default_rng(seed)
    t = np.arange(int(rt * SR)) / SR
    ir = lfilter([0.18], [1, -0.82], r.standard_normal(len(t))) * np.exp(-6.9 * t / rt)
    ir = np.concatenate([np.zeros(int(0.025 * SR)), ir])
    return ir / np.sqrt(np.sum(ir ** 2))

wet = np.stack([fftconvolve(send[c] + 0.35 * dry[c], make_ir(3.6, 11 + c))[:N] for c in range(2)])
mix = dry + 0.42 * wet
mix = lfilter([1, -1], [1, -0.995], mix, axis=1)  # DC / rumble blocker
fade_in = np.minimum(1.0, np.arange(N) / (3.0 * SR))
fade_out = np.clip((N - np.arange(N)) / (4.0 * SR), 0, 1) ** 1.5
mix *= fade_in * fade_out
mix = np.tanh(mix / np.max(np.abs(mix)) * 1.2) / np.tanh(1.2) * 0.89
pcm = (mix.T * 32767).astype(np.int16)

import wave
with wave.open(OUT, "wb") as w:
    w.setnchannels(2); w.setsampwidth(2); w.setframerate(SR); w.writeframes(pcm.tobytes())
print(f"wrote {OUT}: {DUR:.1f}s, {len(events)} chords, {len(sections)} sections")

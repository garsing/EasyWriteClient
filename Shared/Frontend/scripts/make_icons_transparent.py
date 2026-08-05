"""Make near-white / light-gray icon backgrounds transparent (flood from edges)."""
from __future__ import annotations

from collections import deque
from pathlib import Path

from PIL import Image

ROOT = Path(__file__).resolve().parents[1] / "src" / "assets" / "images"

# Icons that are line/glyph style with solid light backgrounds
TARGETS = [
    "add.png",
    "user.png",
    "knowledge_base.png",
    "setting.png",
    "close.png",
    "stop.png",
    "submit.png",
    "unfold.png",
    "chat-history-line.png",
    "manipulate_tool.png",
    "read_tool.png",
    "a.png",
    "r.png",
]


def is_bg(r: int, g: int, b: int, a: int, threshold: int = 235) -> bool:
    if a == 0:
        return True
    # near-white / very light gray (icon plates)
    return r >= threshold and g >= threshold and b >= threshold and abs(r - g) < 18 and abs(g - b) < 18


def remove_bg_flood(im: Image.Image, threshold: int = 235) -> Image.Image:
    im = im.convert("RGBA")
    w, h = im.size
    px = im.load()
    visited = [[False] * w for _ in range(h)]
    q: deque[tuple[int, int]] = deque()

    def try_push(x: int, y: int) -> None:
        if x < 0 or y < 0 or x >= w or y >= h or visited[y][x]:
            return
        r, g, b, a = px[x, y]
        if is_bg(r, g, b, a, threshold):
            visited[y][x] = True
            q.append((x, y))

    for x in range(w):
        try_push(x, 0)
        try_push(x, h - 1)
    for y in range(h):
        try_push(0, y)
        try_push(w - 1, y)

    while q:
        x, y = q.popleft()
        px[x, y] = (0, 0, 0, 0)
        for nx, ny in ((x - 1, y), (x + 1, y), (x, y - 1), (x, y + 1)):
            try_push(nx, ny)

    # Soften fringe: near-white pixels adjacent to transparent become more transparent
    for y in range(h):
        for x in range(w):
            r, g, b, a = px[x, y]
            if a == 0:
                continue
            if not (r >= 220 and g >= 220 and b >= 220):
                continue
            edge = False
            for nx, ny in ((x - 1, y), (x + 1, y), (x, y - 1), (x, y + 1)):
                if 0 <= nx < w and 0 <= ny < h and px[nx, ny][3] == 0:
                    edge = True
                    break
            if edge:
                # keep glyph tint but drop plate
                luma = (r + g + b) / 3
                if luma >= 245:
                    px[x, y] = (0, 0, 0, 0)
                else:
                    # anti-alias fringe
                    na = max(0, min(255, int(255 * (1 - (luma - 220) / 35))))
                    px[x, y] = (r, g, b, na)
    return im


def analyze(path: Path) -> str:
    im = Image.open(path).convert("RGBA")
    w, h = im.size
    px = im.load()
    whiteish = 0
    for y in range(h):
        for x in range(w):
            r, g, b, a = px[x, y]
            if is_bg(r, g, b, a, 245):
                whiteish += 1
    c = [px[0, 0], px[w - 1, 0], px[0, h - 1], px[w - 1, h - 1]]
    return f"{path.name:22} {w}x{h} whiteish={whiteish / (w * h):.1%} corners={c}"


def main() -> None:
    print("--- before ---")
    for name in TARGETS:
        p = ROOT / name
        if p.exists():
            print(analyze(p))

    for name in TARGETS:
        p = ROOT / name
        if not p.exists():
            print(f"skip missing {name}")
            continue
        out = remove_bg_flood(Image.open(p))
        out.save(p, optimize=True)
        print(f"wrote {name}")

    print("--- after ---")
    for name in TARGETS:
        p = ROOT / name
        if p.exists():
            print(analyze(p))


if __name__ == "__main__":
    main()

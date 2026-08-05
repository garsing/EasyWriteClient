from pathlib import Path
from PIL import Image

ROOT = Path(__file__).resolve().parents[1] / "src" / "assets" / "images"
for name in ["add.png", "user.png", "knowledge_base.png", "submit.png", "stop.png"]:
    p = ROOT / name
    im = Image.open(p).convert("RGBA")
    w, h = im.size
    px = im.load()
    opaque = trans = black_opaque = white_opaque = 0
    for y in range(h):
        for x in range(w):
            r, g, b, a = px[x, y]
            if a == 0:
                trans += 1
            else:
                opaque += 1
                if r < 40 and g < 40 and b < 40:
                    black_opaque += 1
                if r > 240 and g > 240 and b > 240:
                    white_opaque += 1
    n = w * h
    print(
        f"{name:22} trans={trans/n:.1%} opaque={opaque/n:.1%} "
        f"black_opaque={black_opaque/n:.1%} white_opaque={white_opaque/n:.1%} "
        f"corner={px[0,0]} center={px[w//2,h//2]}"
    )

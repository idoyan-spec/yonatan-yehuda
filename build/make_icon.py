# -*- coding: utf-8 -*-
"""מייצר favicon.ico ו-assets/icon.ico מתוך icon.svg (רינדור בכרומיום)."""
import os
from playwright.sync_api import sync_playwright
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SVG = os.path.join(HERE, "assets", "icon.svg")
PNG = os.path.join(HERE, "assets", "icon-512.png")

html = f"""<!doctype html><meta charset=utf-8>
<style>html,body{{margin:0;background:transparent}}svg{{width:512px;height:512px;display:block}}</style>
{open(SVG, encoding='utf-8').read()}"""

tmp = os.path.join(HERE, "_icon.html")
open(tmp, "w", encoding="utf-8").write(html)

with sync_playwright() as p:
    b = p.chromium.launch()
    pg = b.new_page(viewport={"width": 512, "height": 512})
    pg.goto("file:///" + tmp.replace("\\", "/"))
    pg.wait_for_timeout(250)
    pg.screenshot(path=PNG, omit_background=True)
    b.close()
os.remove(tmp)

im = Image.open(PNG).convert("RGBA")
sizes = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
im.save(os.path.join(HERE, "assets", "favicon.ico"), sizes=sizes)
os.makedirs(os.path.join(ROOT, "assets"), exist_ok=True)
im.save(os.path.join(ROOT, "assets", "icon.ico"), sizes=sizes)
im.resize((180, 180), Image.LANCZOS).save(os.path.join(HERE, "assets", "apple-touch-icon.png"))
print("אייקונים נוצרו:", os.path.join(ROOT, "assets", "icon.ico"))

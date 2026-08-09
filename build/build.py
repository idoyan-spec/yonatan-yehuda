# -*- coding: utf-8 -*-
r"""
בונה את אתר כתבי יונתן יהודה ז"ל.

הרצה:   python build\build.py
פלט:    docs\   (אתר סטטי מוכן להעלאה)

הוספת כתבים חדשים: פשוט שמים קובץ .docx בתיקיית "כתבים_גיבוי" ומריצים שוב.
הוספת שירים:       תיקיית "שירים"  (.docx / .txt / .md / .pdf / תמונות)
הוספת הקלטות:      תיקיית "מוסיקה" או "הקלטות"  (.mp3/.wav/.m4a/.amr/.ogg)
"""
import html
import json
import os
import re
import shutil
import subprocess
import sys
import zipfile
from datetime import date, datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from pyluach import dates as pl

from config import BUILD, SITE, DOCS, COLLECTIONS, SKIP, SONGS
from docx_read import read_docx

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BUILD_DIR = os.path.join(ROOT, "build")
OUT = os.path.join(ROOT, "docs")
WRITINGS_DIRS = [os.path.join(ROOT, "כתבים_גיבוי"), os.path.join(ROOT, "כתבים")]
SONGS_DIR = os.path.join(ROOT, "שירים")
AUDIO_DIRS = [os.path.join(ROOT, "מוסיקה"), os.path.join(ROOT, "הקלטות")]
PHOTO = os.path.join(ROOT, "IMG_9166.JPG")

AUDIO_EXT = {".mp3", ".wav", ".m4a", ".amr", ".ogg", ".aac", ".wma", ".flac"}
IMAGE_EXT = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".heic"}

BIRTH = date(*SITE["birth_greg"])
DEATH = pl.HebrewDate(*SITE["death_heb"]).to_greg()
DEATH = date(DEATH.year, DEATH.month, DEATH.day)

MONTHS_HE = ["בינואר", "בפברואר", "במרץ", "באפריל", "במאי", "ביוני",
             "ביולי", "באוגוסט", "בספטמבר", "באוקטובר", "בנובמבר", "בדצמבר"]


# ─────────────────────────────  תאריכים וגיל  ─────────────────────────────

def heb_date(d):
    """כ״ג תמוז תשע״ג"""
    return pl.GregorianDate(d.year, d.month, d.day).to_heb().hebrew_date_string()


def greg_date(d):
    return f"{d.day} {MONTHS_HE[d.month - 1]} {d.year}"


def age_at(d):
    """מחזיר (שנים, חודשים) בתאריך נתון."""
    y = d.year - BIRTH.year
    m = d.month - BIRTH.month
    if d.day < BIRTH.day:
        m -= 1
    if m < 0:
        y -= 1
        m += 12
    return y, m


def age_text(d):
    y, m = age_at(d)
    if m == 0:
        return f"בגיל {y}"
    if m == 1:
        return f"בגיל {y} וחודש"
    if m == 2:
        return f"בגיל {y} וחודשיים"
    return f"בגיל {y} ו־{m} חודשים"


def year_span(dates_):
    ys = sorted({d.year for d in dates_})
    return str(ys[0]) if len(ys) == 1 else f"{ys[0]}–{ys[-1]}"


def age_short(d):
    return f"גיל {age_at(d)[0]}"


# ─────────────────────────────  עזרי HTML  ─────────────────────────────

def e(s):
    return html.escape(s, quote=True)


def slugify(s):
    s = re.sub(r"[^\w֐-׿-]+", "-", s, flags=re.UNICODE).strip("-")
    return s or "item"


def translit_slug(s):
    """סלאג לטיני פשוט לשמות קבצים (למניעת בעיות קידוד בכתובות)."""
    table = {"א": "a", "ב": "b", "ג": "g", "ד": "d", "ה": "h", "ו": "v", "ז": "z",
             "ח": "ch", "ט": "t", "י": "y", "כ": "k", "ך": "k", "ל": "l", "מ": "m",
             "ם": "m", "נ": "n", "ן": "n", "ס": "s", "ע": "a", "פ": "p", "ף": "f",
             "צ": "tz", "ץ": "tz", "ק": "k", "ר": "r", "ש": "sh", "ת": "t",
             " ": "-", "_": "-"}
    out = "".join(table.get(c, c if re.match(r"[A-Za-z0-9-]", c) else "") for c in s)
    return re.sub(r"-+", "-", out).strip("-").lower() or "item"


NAV = [("index.html", "בית"), ("ketavim.html", "הכתבים"),
       ("shirim.html", "שירים"), ("haklatot.html", "הקלטות")]


def shell(title, body, current, depth=0, desc="", extra_head=""):
    up = "../" * depth
    nav = "".join(
        f'<li><a class="lnk" href="{up}{href}"'
        f'{" aria-current=\"page\"" if href == current else ""}>{e(label)}</a></li>'
        for href, label in NAV)
    return f"""<!doctype html>
<html lang="he" dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{e(title)}</title>
<meta name="description" content="{e(desc or SITE['tagline'])}">
<meta property="og:title" content="{e(title)}">
<meta property="og:description" content="{e(desc or SITE['tagline'])}">
<meta property="og:type" content="website">
<link rel="icon" href="{up}assets/icon.svg" type="image/svg+xml">
<link rel="alternate icon" href="{up}assets/favicon.ico">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Frank+Ruhl+Libre:wght@400;500;700&family=Assistant:wght@400;600;700&display=swap" rel="stylesheet">
<link rel="stylesheet" href="{up}assets/style.css">
{extra_head}
</head>
<body>
<nav class="nav no-print">
  <div class="nav-in">
    <a class="brand" href="{up}index.html">יונתן יהודה <span>ז״ל</span></a>
    <ul>{nav}</ul>
  </div>
</nav>
{body}
<footer class="no-print">
  <div class="foot-in">
    <div>לזכרו של {e(SITE['name'])}<br>
      <span class="faint">{e(heb_date(BIRTH))} – {e(heb_date(DEATH))}</span></div>
    <div class="build">גרסת אתר {e(BUILD)}</div>
  </div>
</footer>
</body>
</html>
"""


ORNAMENT = (
    '<svg class="ornament" viewBox="0 0 160 20" fill="none" aria-hidden="true">'
    '<path d="M2 10h56M102 10h56" stroke="#C9AE74" stroke-width="1"/>'
    '<path d="M80 2.5c3.2 3.6 7.5 5.6 7.5 8.4 0 3-3.4 5.1-7.5 5.1s-7.5-2.1-7.5-5.1'
    'c0-2.8 4.3-4.8 7.5-8.4z" fill="#C9AE74" opacity=".85"/>'
    '<circle cx="66" cy="10" r="1.6" fill="#C9AE74"/>'
    '<circle cx="94" cy="10" r="1.6" fill="#C9AE74"/></svg>'
)

I_PRINT = ('<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" '
           'stroke-linecap="round" stroke-linejoin="round"><path d="M6 9V2h12v7"/>'
           '<path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/>'
           '<path d="M6 14h12v8H6z"/></svg>')
I_DL = ('<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" '
        'stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v12"/>'
        '<path d="m7 12 5 5 5-5"/><path d="M4 21h16"/></svg>')


# ─────────────────────────────  איסוף הכתבים  ─────────────────────────────

def collect_writings():
    items = []
    for base in WRITINGS_DIRS:
        if not os.path.isdir(base):
            continue
        for dp, _dn, fn in os.walk(base):
            for f in sorted(fn):
                if not f.lower().endswith(".docx") or f.startswith("~$"):
                    continue
                path = os.path.join(dp, f)
                rel = os.path.relpath(path, base)
                if rel in SKIP or os.path.getsize(path) == 0:
                    continue
                cfg = DOCS.get(rel, {})
                try:
                    blocks = read_docx(path)
                except Exception as exc:
                    print(f"  ! דילוג על {rel}: {exc}")
                    continue
                if not blocks:
                    continue

                created = docx_created(path)
                folder = os.path.dirname(rel).split(os.sep)[0] if os.sep in rel else ""
                items.append({
                    "rel": rel,
                    "path": path,
                    "blocks": blocks,
                    "created": created,
                    "title": cfg.get("title") or os.path.splitext(os.path.basename(rel))[0],
                    "subtitle": cfg.get("subtitle", ""),
                    "slug": cfg.get("slug") or translit_slug(os.path.splitext(os.path.basename(rel))[0]),
                    "collection": cfg.get("collection", folder or None),
                    "order": cfg.get("order", 99),
                    "sections": cfg.get("sections"),
                    "order_blocks": cfg.get("order_blocks"),
                    "join": cfg.get("join"),
                    "list": cfg.get("list", False),
                })
    items.sort(key=lambda x: (x["collection"] or "", x["order"], x["created"]))
    return items


def docx_created(path):
    with zipfile.ZipFile(path) as z:
        core = z.read("docProps/core.xml").decode("utf-8", "ignore")
    m = re.search(r"<dcterms:created[^>]*>(\d{4})-(\d{2})-(\d{2})", core)
    if m:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    ts = os.path.getmtime(path)
    d = datetime.fromtimestamp(ts)
    return date(d.year, d.month, d.day)


def runs_html(runs):
    out = []
    for r in runs:
        t = e(r["t"]).replace("\n", "<br>")
        if r.get("b"):
            t = f"<strong>{t}</strong>"
        if r.get("i"):
            t = f"<em>{t}</em>"
        out.append(t)
    return "".join(out)


def blocks_html(item):
    blocks = item["blocks"]

    # סדר קריאה מתוקן (למסמכים שנכתבו בתיבות טקסט פזורות על הדף)
    if item.get("order_blocks"):
        idx = item["order_blocks"]
        joins = dict(item.get("join") or [])
        parts, i = [], 0
        while i < len(idx):
            cur = idx[i]
            merged = list(blocks[cur]["runs"])
            while i + 1 < len(idx) and joins.get(cur) == idx[i + 1]:
                nxt = idx[i + 1]
                merged += blocks[nxt]["runs"]
                cur = nxt
                i += 1
            parts.append({"type": "p", "runs": merged})
            i += 1
        blocks = parts

    # מבנה של שאלה/תשובה
    if item.get("sections"):
        out = []
        for sec in item["sections"]:
            out.append(f'<p class="sec-label">{e(sec["label"])}</p>')
            inner = "".join(f"<p>{runs_html(blocks[i]['runs'])}</p>"
                            for i in sec["blocks"] if i < len(blocks))
            out.append(f"<blockquote>{inner}</blockquote>" if sec.get("quote") else inner)
        return "\n".join(out)

    # מצב רשימה (מסמך שכולו רשימת פריטים)
    if item.get("list"):
        head = f"<p>{runs_html(blocks[0]['runs'])}</p>" if blocks else ""
        lis = "".join(f"<li>{runs_html(b['runs'])}</li>" for b in blocks[1:])
        return f"{head}<ul>{lis}</ul>"

    out = []
    n = len(blocks)
    in_list = False
    for i, b in enumerate(blocks):
        inner = runs_html(b["runs"])
        text = "".join(r["t"] for r in b["runs"]).strip()

        # פסקה ראשונה שחוזרת על כותרת המסמך — מיותרת באתר
        if i == 0 and _norm_t(text) == _norm_t(item.get("title", "")):
            continue

        # פריטי רשימה ממוספרים מ-Word
        if b.get("li"):
            if not in_list:
                out.append("<ul>")
                in_list = True
            out.append(f"<li>{inner}</li>")
            continue
        if in_list:
            out.append("</ul>")
            in_list = False

        kind = "h" if b.get("type") == "h" else classify(text, i, n, item)
        if kind == "bsd":
            out.append('<p class="bsd">בס"ד</p>')
        elif kind == "byline":
            out.append(f'<p class="byline">{inner}</p>')
        elif kind == "sign":
            out.append(f'<p class="sign">{inner}</p>')
        elif kind == "h":
            out.append(f"<h2>{inner}</h2>")
        else:
            out.append(f"<p>{inner}</p>")
    if in_list:
        out.append("</ul>")
    return "\n".join(out)


def _norm_t(s):
    return re.sub(r"[\s'\"״׳.,!?]+", "", s or "")


BSD = {"בס''ד", 'בס"ד', "בס״ד"}
END_PUNCT = (".", "!", "?", ",", ";", "״", '"', "…")


def classify(text, i, n, item):
    """מסווג פסקה: בס"ד / שורת מחבר / חתימה / כותרת / פסקה."""
    t = text.strip()
    flat = t.replace("״", '"').replace("''", '"')
    if flat in BSD or (len(t) <= 5 and flat.startswith('בס"')):
        return "bsd"
    # שורת מחבר כמו "תפילה/יונתן ינקלוביץ"
    if i <= 2 and "/" in t and "יונתן" in t and len(t) <= 45:
        return "byline"
    # חתימה בסוף
    if i >= n - 3 and len(t) <= 24 and "יונתן" in t:
        return "sign"
    if len(t) <= 34 and i > 0 and not t.endswith(END_PUNCT) and len(t.split()) <= 5:
        return "h"
    return "p"


def excerpt(item, n=125):
    for b in item["blocks"]:
        t = "".join(r["t"] for r in b["runs"]).strip()
        if len(t) > 70:
            return (t[:n].rstrip() + "…") if len(t) > n else t
    t = "".join(r["t"] for r in item["blocks"][0]["runs"]).strip()
    return t[:n]


# ─────────────────────────────  עמודי הכתבים  ─────────────────────────────

def stamp_html(d, extra=""):
    return (f'<div class="stamp">נכתב ב<b>{e(heb_date(d))}</b> · {e(greg_date(d))}'
            f'<br>{e(age_text(d))}{extra}</div>')


def write_doc_pages(items):
    os.makedirs(os.path.join(OUT, "k"), exist_ok=True)
    for i, it in enumerate(items):
        prev = items[i - 1] if i > 0 else None
        nxt = items[i + 1] if i < len(items) - 1 else None
        d = it["created"]
        docx_name = f'{it["slug"]}.docx'
        shutil.copy2(it["path"], os.path.join(OUT, "k", docx_name))

        tools = f"""<div class="tools no-print">
  <button class="btn primary" onclick="window.print()">{I_PRINT} הדפסה</button>
  <a class="btn" href="{it['slug']}.pdf" download>{I_DL} PDF</a>
  <a class="btn" href="{e(docx_name)}" download>{I_DL} קובץ Word מקורי</a>
</div>"""

        nav_prev = (f'<a href="{prev["slug"]}.html">→ הקודם<span class="t">{e(prev["title"])}</span></a>'
                    if prev else '<a href="../ketavim.html">→ כל הכתבים</a>')
        nav_next = (f'<a href="{nxt["slug"]}.html">הבא ←<span class="t">{e(nxt["title"])}</span></a>'
                    if nxt else '<a href="../ketavim.html">כל הכתבים ←</a>')

        coll = ""
        if it["collection"]:
            coll = f'<p class="eyebrow">מתוך · {e(it["collection"])}</p>'

        body = f"""<main class="wrap">
  <header class="doc-head">
    {coll}
    <h1>{e(it['title'])}</h1>
    {f'<p class="sub">{e(it["subtitle"])}</p>' if it['subtitle'] else ''}
    {stamp_html(d)}
  </header>
  {tools}
  <article class="doc">
{blocks_html(it)}
    <div class="print-foot">
      {e(it['title'])} · {e(SITE['name'])} · נכתב ב{e(heb_date(d))} ({e(greg_date(d))}), {e(age_text(d))}
    </div>
  </article>
  <nav class="docnav no-print">{nav_prev}{nav_next}</nav>
</main>"""
        page = shell(f"{it['title']} · {SITE['short']}", body, "ketavim.html",
                     depth=1, desc=excerpt(it))
        with open(os.path.join(OUT, "k", f"{it['slug']}.html"), "w", encoding="utf-8") as f:
            f.write(page)


def write_index_page(items):
    groups = {}
    for it in items:
        groups.setdefault(it["collection"], []).append(it)

    parts = []
    singles = groups.pop(None, [])
    if singles:
        parts.append('<div class="grid">' + "".join(card(it) for it in singles) + "</div>")
    for name, lst in groups.items():
        info = COLLECTIONS.get(name, {})
        parts.append(f"""<section>
  <div class="collection">
    <h2>{e(name)}</h2>
    {f'<p>{e(info["blurb"])}</p>' if info.get('blurb') else ''}
  </div>
  <div class="grid">{''.join(card(it) for it in lst)}</div>
</section>""")

    span = year_span([i["created"] for i in items])
    body = f"""<main class="wrap">
  <section style="padding-top:2.6rem">
    <p class="eyebrow">{len(items)} כתבים · {e(span)}</p>
    <h1>הכתבים</h1>
    <p class="lead narrow" style="margin-inline:0">
      כל מה שיונתן כתב, לפי סדר. ליד כל כתב מופיע התאריך שבו נכתב — עברי ולועזי —
      והגיל שלו באותו יום. אפשר לקרוא כאן, להדפיס בלחיצה אחת, או להוריד כקובץ.
    </p>
    <div class="tools" style="margin-inline:0">
      <a class="btn" href="ketavim-all.pdf" download>{I_DL} כל הכתבים בקובץ PDF אחד</a>
      <a class="btn" href="ketavim.zip" download>{I_DL} הורדת הכל (ZIP)</a>
    </div>
  </section>
  {''.join(parts)}
</main>"""
    write(os.path.join(OUT, "ketavim.html"),
          shell(f"הכתבים · {SITE['short']}", body, "ketavim.html",
                desc="כל כתביו של יונתן יהודה ז\"ל — לקריאה, להדפסה ולהורדה."))


def card(it):
    d = it["created"]
    return f"""<a class="card" href="k/{it['slug']}.html">
  <h3>{e(it['title'])}</h3>
  <p class="sub">{e(it['subtitle'] or excerpt(it, 95))}</p>
  <div class="meta"><span>{e(heb_date(d))}</span><span>·</span>
    <span>{d.year}</span><span class="age">{e(age_short(d))}</span></div>
</a>"""


# ─────────────────────────────  עמוד הבית  ─────────────────────────────

def write_home(items, songs, recs):
    span = year_span([i["created"] for i in items])
    ages = sorted({age_at(i["created"])[0] for i in items})
    age_span = f"{min(ages)}" if len(ages) == 1 else f"{min(ages)}–{max(ages)}"
    body = f"""<main class="wrap">
  <section class="hero">
    <div class="portrait"><img src="assets/yonatan.jpg" alt="יונתן יהודה ז״ל" width="640" height="640"></div>
    <h1>יונתן יהודה ינקלוביץ ז״ל</h1>
    <p class="dates">
      <b>{e(heb_date(BIRTH))}</b> · {e(greg_date(BIRTH))}<br>
      <b>{e(heb_date(DEATH))}</b> · {e(greg_date(DEATH))}<br>
      <span class="faint">נהרג בגיל 18</span>
    </p>
    {ORNAMENT}
    <div class="intro">
      <p>כשיונתן נהרג, מצאנו את הדברים שכתב. הוא היה אז ילד — בן {e(age_span)} —
         והדפים מלאים בשאלות גדולות ובתשובות שחיפש לעצמו: על תפילה, על שמחה,
         על איך לזכור את מה שחשוב בכל רגע ורגע.</p>
      <p>האתר הזה נועד כדי שאפשר יהיה לקרוא אותם, ולשמוע אותו. ליד כל כתב,
         שיר והקלטה רשום התאריך — עברי ולועזי — והגיל שלו באותו יום.
         הכול פתוח לקריאה, להאזנה, להדפסה ולהורדה.</p>
    </div>
    <div class="tools" style="justify-content:center; margin-top:2rem">
      <a class="btn primary" href="ketavim.html">אל הכתבים</a>
      <a class="btn" href="haklatot.html">הקלטות</a>
      <a class="btn" href="shirim.html">שירים</a>
    </div>
  </section>

  <hr class="soft">

  <section>
    <h2 style="text-align:center">מה יש כאן</h2>
    <div class="grid" style="margin-top:1.4rem">
      <a class="card" href="ketavim.html"><h3>הכתבים</h3>
        <p class="sub">{len(items)} חיבורים משנת {e(span)}.</p>
        <div class="meta"><span>קריאה · הדפסה · הורדה</span></div></a>
      <a class="card" href="haklatot.html"><h3>הקלטות</h3>
        <p class="sub">{"%d הקלטות של קולו — לימוד, שירה ורגעים מהבית" % len(recs)
                        if recs else "טרם נוספו הקלטות"}.</p>
        <div class="meta"><span>האזנה · הורדה</span></div></a>
      <a class="card" href="shirim.html"><h3>שירים</h3>
        <p class="sub">{"%d שירים ששר והקליט, בערך בגיל %d" % (len(songs), SONGS["approx_age"])
                        if songs else "מקום לשירים שיתווספו"}.</p>
        <div class="meta"><span>האזנה · הורדה</span></div></a>
    </div>
  </section>
</main>"""
    write(os.path.join(OUT, "index.html"),
          shell(SITE["name"], body, "index.html",
                desc="כתביו, שיריו והקלטותיו של יונתן יהודה ינקלוביץ ז״ל."))


# ─────────────────────────────  שירים  ─────────────────────────────

def collect_songs():
    items = []
    if not os.path.isdir(SONGS_DIR):
        return items
    for dp, _dn, fn in os.walk(SONGS_DIR):
        for f in sorted(fn):
            if f.startswith("~$") or f.lower() == "readme.txt":
                continue
            path = os.path.join(dp, f)
            ext = os.path.splitext(f)[1].lower()
            name = os.path.splitext(f)[0]
            item = {"path": path, "title": name, "ext": ext, "kind": "file",
                    "slug": translit_slug(name), "created": None, "blocks": None,
                    "note": ""}
            if ext in AUDIO_EXT:
                item["kind"] = "audio"
                title, note, d = read_sidecar(path, name, "")
                item["title"] = title or name
                item["note"] = note
                item["created"] = d or audio_tag_date(path)
            elif ext == ".docx":
                item["kind"] = "text"
                try:
                    item["blocks"] = read_docx(path)
                    item["created"] = docx_created(path)
                except Exception as exc:
                    print(f"  ! שיר {f}: {exc}")
                    continue
            elif ext in (".txt", ".md"):
                if os.path.exists(os.path.splitext(path)[0]) or _is_sidecar(path):
                    continue          # קובץ תיאור של שיר אחר, לא שיר בפני עצמו
                item["kind"] = "text"
                txt = open(path, encoding="utf-8", errors="replace").read()
                item["blocks"] = [{"type": "p", "text": p, "runs": [{"t": p}]}
                                  for p in txt.split("\n") if p.strip()]
                item["created"] = mtime_date(path)
            elif ext in IMAGE_EXT or ext == ".pdf":
                item["created"] = mtime_date(path)
            else:
                continue
            items.append(item)
    items.sort(key=lambda x: (x["kind"] != "audio", x["title"]))
    return items


def _is_sidecar(path):
    """קובץ כמו 'שיר.mp3.txt' הוא תיאור לשיר, לא פריט בפני עצמו."""
    stem = os.path.splitext(path)[0]
    return os.path.splitext(stem)[1].lower() in AUDIO_EXT


def read_sidecar(path, name, default_title):
    """קורא קובץ .txt צמוד: שורה 1 = כותרת, 'תאריך: YYYY-MM-DD' = תאריך, השאר = הסבר."""
    title, note, d = default_title or name, "", None
    sidecar = path + ".txt"
    if not os.path.exists(sidecar):
        return title, note, d
    lines = [l.strip() for l in
             open(sidecar, encoding="utf-8", errors="replace").read().splitlines()]
    lines = [l for l in lines if l]
    body = []
    for k, line in enumerate(lines):
        m = re.match(r"^(?:תאריך|date)\s*[:=]\s*(\d{4})-(\d{1,2})-(\d{1,2})\s*$", line)
        if m:
            d = date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        elif k == 0:
            title = line
        else:
            body.append(line)
    return title, " ".join(body), d


def audio_tag_date(path):
    """
    תאריך מתוך התגים המוטבעים בקובץ הקול (מכשירי הקלטה רבים כותבים אותו).
    מוחזר רק אם הוא נופל בתוך שנות חייו של יונתן - אחרת הוא חסר משמעות.
    """
    try:
        out = subprocess.run(
            ["ffprobe", "-v", "quiet", "-show_entries",
             "format_tags=date,creation_time,TDRC,year", "-of", "default=nw=1",
             path], capture_output=True, text=True, timeout=60).stdout
    except Exception:
        return None
    for m in re.finditer(r"(\d{4})-(\d{1,2})-(\d{1,2})", out):
        try:
            d = date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            continue
        if BIRTH <= d <= DEATH:
            return d
    return None


def mtime_date(path):
    d = datetime.fromtimestamp(os.path.getmtime(path))
    return date(d.year, d.month, d.day)


def write_songs(songs):
    if not songs:
        body = """<main class="wrap">
  <section style="padding-top:2.6rem">
    <p class="eyebrow">שירים</p>
    <h1>שירים</h1>
    <p class="lead narrow" style="margin-inline:0">העמוד הזה מחכה לשירים של יונתן.</p>
  </section>
  <section>
    <div class="empty">
      <h2>עוד לא נוספו שירים</h2>
      <p>כשיימצאו שירים — מספיק לשים אותם בתיקיית <b>שירים</b> שבתיקיית הפרויקט
         (קובץ Word, טקסט, PDF או צילום של הדף), ולהריץ מחדש את הבנייה.
         הם יופיעו כאן מיד, עם התאריך העברי והלועזי והגיל, בדיוק כמו הכתבים.</p>
    </div>
  </section>
</main>"""
        write(os.path.join(OUT, "shirim.html"),
              shell(f"שירים · {SITE['short']}", body, "shirim.html",
                    desc="שיריו של יונתן יהודה ז\"ל."))
        return

    os.makedirs(os.path.join(OUT, "s"), exist_ok=True)
    media = os.path.join(OUT, "media")
    os.makedirs(media, exist_ok=True)

    rows, cards = [], []
    for it in songs:
        d = it["created"]
        meta = (f'<div class="meta"><span>{e(heb_date(d))}</span><span>·</span>'
                f'<span>{d.year}</span><span class="age">{e(age_short(d))}</span></div>') if d else ""

        if it["kind"] == "audio":
            rows.append(song_row(it, media))
            continue

        if it["blocks"] is not None:
            shutil.copy2(it["path"], os.path.join(OUT, "s", os.path.basename(it["path"])))
            body = f"""<main class="wrap">
  <header class="doc-head">
    <p class="eyebrow">שיר</p>
    <h1>{e(it['title'])}</h1>
    {stamp_html(d) if d else ''}
  </header>
  <div class="tools no-print">
    <button class="btn primary" onclick="window.print()">{I_PRINT} הדפסה</button>
    <a class="btn" href="{e(os.path.basename(it['path']))}" download>{I_DL} הורדת המקור</a>
  </div>
  <article class="doc">
{blocks_html(it)}
    <div class="print-foot">{e(it['title'])} · {e(SITE['name'])}</div>
  </article>
  <nav class="docnav no-print"><a href="../shirim.html">→ כל השירים</a></nav>
</main>"""
            write(os.path.join(OUT, "s", f"{it['slug']}.html"),
                  shell(f"{it['title']} · {SITE['short']}", body, "shirim.html", depth=1))
            href = f"s/{it['slug']}.html"
        else:
            dst = os.path.join(OUT, "s", os.path.basename(it["path"]))
            shutil.copy2(it["path"], dst)
            href = f"s/{os.path.basename(it['path'])}"
        cards.append(f'<a class="card" href="{href}"><h3>{e(it["title"])}</h3>'
                     f'<p class="sub">{e(it["ext"].lstrip("."))}</p>{meta}</a>')

    approx = SONGS.get("approx_age")
    approx_line = ""
    if approx and rows:
        y0 = date(BIRTH.year + approx, BIRTH.month, BIRTH.day)
        y1 = date(BIRTH.year + approx + 1, BIRTH.month, BIRTH.day)
        hy0 = pl.GregorianDate(y0.year, y0.month, y0.day).to_heb().hebrew_date_string()
        hy1 = pl.GregorianDate(y1.year, y1.month, y1.day).to_heb().hebrew_date_string()
        approx_line = (f'<p class="muted narrow" style="margin-inline:0; font-size:.94rem">'
                       f'לקובצי הקול אין תאריך מדויק שנשמר, ולכן הגיל כאן משוער — '
                       f'גיל {approx} נופל בין {e(hy0.split()[-1])} ל{e(hy1.split()[-1])} '
                       f'({y0.year}–{y1.year}).</p>')

    body = f"""<main class="wrap">
  <section style="padding-top:2.6rem">
    <p class="eyebrow">{len(songs)} שירים</p>
    <h1>שירים</h1>
    <p class="lead narrow" style="margin-inline:0">{e(SONGS.get('blurb', ''))}</p>
    {approx_line}
  </section>
  {f'<section>{"".join(rows)}</section>' if rows else ''}
  {f'<section><div class="grid">{"".join(cards)}</div></section>' if cards else ''}
</main>"""
    write(os.path.join(OUT, "shirim.html"),
          shell(f"שירים · {SITE['short']}", body, "shirim.html",
                desc="שירים שיונתן יהודה ז\"ל שר והקליט בילדותו."))


def song_row(it, media):
    """כרטיס שיר עם נגן — כמו בעמוד ההקלטות."""
    dst = os.path.join(media, it["slug"] + ".mp3")
    if not os.path.exists(dst):
        print(f"  ♪ ממיר {os.path.basename(it['path'])} …")
        res = subprocess.run(
            ["ffmpeg", "-y", "-loglevel", "error", "-i", it["path"], "-vn",
             "-ar", "44100", "-b:a", "128k", dst], capture_output=True, text=True)
        if res.returncode != 0:
            print(f"  ! המרה נכשלה: {res.stderr[:200]}")
            return ""
    dur = fmt_dur(ffprobe_duration(dst))
    size = os.path.getsize(dst) / 1048576
    d = it["created"]
    meta = []
    if d:
        meta.append(f"{heb_date(d)} · {greg_date(d)}")
        meta.append(age_text(d))
    elif SONGS.get("approx_age"):
        meta.append(f"בערך בגיל {SONGS['approx_age']}")
    if dur:
        meta.append(dur)
    meta.append(f"{size:.1f} MB")
    note = (f'<p class="muted" style="margin:.2rem 0 .7rem">{e(it["note"])}</p>'
            if it.get("note") else "")
    return f"""<div class="rec">
  <h3>{e(it['title'])}</h3>
  <div class="meta">{e(' · '.join(meta))}</div>
  {note}
  <audio controls preload="none" src="media/{e(it['slug'])}.mp3"></audio>
  <div class="dl"><a class="btn" href="media/{e(it['slug'])}.mp3" download>{I_DL} הורדה</a></div>
</div>"""


# ─────────────────────────────  הקלטות  ─────────────────────────────

def collect_recordings():
    items = []
    for base in AUDIO_DIRS:
        if not os.path.isdir(base):
            continue
        for dp, _dn, fn in os.walk(base):
            for f in sorted(fn):
                ext = os.path.splitext(f)[1].lower()
                if ext not in AUDIO_EXT:
                    continue
                path = os.path.join(dp, f)
                name = os.path.splitext(f)[0]
                title, note = nice_recording_name(name)
                override_date = None
                sidecar = path + ".txt"
                if os.path.exists(sidecar):
                    lines = [l.strip() for l in
                             open(sidecar, encoding="utf-8", errors="replace").read().splitlines()]
                    lines = [l for l in lines if l]
                    body = []
                    for k, line in enumerate(lines):
                        m = re.match(r"^(?:תאריך|date)\s*[:=]\s*"
                                     r"(\d{4})-(\d{1,2})-(\d{1,2})\s*$", line)
                        if m:
                            override_date = date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
                        elif k == 0:
                            title = line
                        else:
                            body.append(line)
                    note = " ".join(body)
                items.append({"path": path, "title": title, "note": note,
                              "override_date": override_date,
                              "named": bool(title != "הקלטה"),
                              "slug": translit_slug(name) or slugify(name),
                              "date": override_date or recording_date(path, name)})
    # לפי סדר כרונולוגי; מה שאין לו תאריך אמין — בסוף
    items.sort(key=lambda x: (x["date"] is None, x["date"] or date(2100, 1, 1)))
    # הקלטות שאין להן שם — ממוספרות לפי הסדר הכרונולוגי
    k = 0
    for it in items:
        if not it["named"]:
            k += 1
            it["title"] = f"הקלטה {k}"
    return items


def recording_date(path, name):
    """
    תאריך ההקלטה. שמות כמו 160731-004 מקודדים את התאריך בעצמם
    ולכן הם מהימנים. אחרת נופלים על חותמת הקובץ, שיכולה להיות תאריך
    העתקה ולא ההקלטה - על כן תאריך מחוץ לשנות חייו נחשב לא-ידוע.
    """
    m = re.match(r"^(\d{2})(\d{2})(\d{2})-\d+$", name)   # 160731-004
    if m:
        try:
            return date(2000 + int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            pass
    tagged = audio_tag_date(path)          # תג תאריך מוטבע בקובץ - מהימן
    if tagged:
        return tagged
    d = mtime_date(path)
    return d if BIRTH <= d <= DEATH else None


def nice_recording_name(name):
    m = re.match(r"^(\d{2})(\d{2})(\d{2})-(\d+)$", name)
    if m:
        return "הקלטה", ""
    if re.match(r"^[A-Z]?\d{5,}$", name):
        return "הקלטה", ""
    return name.replace("_", " ").strip(" ."), ""


def ffprobe_duration(path):
    try:
        out = subprocess.run(
            ["ffprobe", "-v", "quiet", "-show_entries", "format=duration",
             "-of", "csv=p=0", path], capture_output=True, text=True, timeout=90)
        return float(out.stdout.strip())
    except Exception:
        return None


def fmt_dur(sec):
    if not sec:
        return ""
    s = int(round(sec))
    return f"{s // 60}:{s % 60:02d}"


def write_recordings(recs):
    media = os.path.join(OUT, "media")
    os.makedirs(media, exist_ok=True)

    if not recs:
        body = """<main class="wrap">
  <section style="padding-top:2.6rem"><p class="eyebrow">הקלטות</p><h1>הקלטות</h1></section>
  <section><div class="empty"><h2>עוד לא נוספו הקלטות</h2>
    <p>אפשר לשים קבצי קול בתיקיית <b>מוסיקה</b> או <b>הקלטות</b> ולהריץ מחדש את הבנייה.</p>
  </div></section></main>"""
        write(os.path.join(OUT, "haklatot.html"),
              shell(f"הקלטות · {SITE['short']}", body, "haklatot.html"))
        return

    rows = []
    for r in recs:
        dst = os.path.join(media, r["slug"] + ".mp3")
        if not os.path.exists(dst):
            print(f"  ♪ ממיר {os.path.basename(r['path'])} …")
            cmd = ["ffmpeg", "-y", "-loglevel", "error", "-i", r["path"],
                   "-vn", "-ac", "1", "-ar", "44100", "-b:a", "112k", dst]
            res = subprocess.run(cmd, capture_output=True, text=True)
            if res.returncode != 0:
                print(f"  ! המרה נכשלה: {res.stderr[:200]}")
                continue
        dur = fmt_dur(ffprobe_duration(dst))
        size = os.path.getsize(dst) / 1048576
        d = r["date"]
        meta = []
        if d:
            meta.append(f"{heb_date(d)} · {greg_date(d)}")
            meta.append(age_text(d))
        else:
            meta.append("תאריך לא ידוע")
        if dur:
            meta.append(dur)
        meta.append(f"{size:.1f} MB")
        rows.append(f"""<div class="rec">
  <h3>{e(r['title'])}</h3>
  <div class="meta">{e(' · '.join(meta))}</div>
  {f'<p class="muted" style="margin:.2rem 0 .7rem">{e(r["note"])}</p>' if r['note'] else ''}
  <audio controls preload="none" src="media/{e(r['slug'])}.mp3"></audio>
  <div class="dl"><a class="btn" href="media/{e(r['slug'])}.mp3" download>{I_DL} הורדה</a></div>
</div>""")

    body = f"""<main class="wrap">
  <section style="padding-top:2.6rem">
    <p class="eyebrow">{len(rows)} הקלטות</p>
    <h1>הקלטות</h1>
    <p class="lead narrow" style="margin-inline:0">
      קולו של יונתן. ליד כל הקלטה רשום התאריך העברי והלועזי והגיל שלו באותו יום.
    </p>
    <p class="muted narrow" style="margin-inline:0; font-size:.94rem">
      התאריכים נלקחים משם הקובץ או מחותמת הזמן שלו. היכן שהחותמת אינה אמינה
      רשום "תאריך לא ידוע" — ואפשר להשלים אותו ידנית.
    </p>
  </section>
  <section>{''.join(rows)}</section>
</main>"""
    write(os.path.join(OUT, "haklatot.html"),
          shell(f"הקלטות · {SITE['short']}", body, "haklatot.html",
                desc="הקלטות קול של יונתן יהודה ז\"ל."))


# ─────────────────────────────  PDF ו-ZIP  ─────────────────────────────

def make_pdfs(items):
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("  ! playwright לא מותקן — מדלג על יצירת PDF")
        return False
    ok = True
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch()
            page = browser.new_page()
            for it in items:
                src = os.path.join(OUT, "k", f"{it['slug']}.html")
                page.goto("file:///" + src.replace("\\", "/"))
                page.emulate_media(media="print")
                page.wait_for_timeout(350)
                page.pdf(path=os.path.join(OUT, "k", f"{it['slug']}.pdf"),
                         format="A4", print_background=True,
                         margin={"top": "18mm", "bottom": "20mm",
                                 "left": "20mm", "right": "20mm"})
            # קובץ אחד עם הכל
            src = os.path.join(OUT, "_all.html")
            page.goto("file:///" + src.replace("\\", "/"))
            page.emulate_media(media="print")
            page.wait_for_timeout(400)
            page.pdf(path=os.path.join(OUT, "ketavim-all.pdf"), format="A4",
                     print_background=True,
                     margin={"top": "18mm", "bottom": "20mm",
                             "left": "20mm", "right": "20mm"})
            browser.close()
    except Exception as exc:
        print(f"  ! יצירת PDF נכשלה: {exc}")
        ok = False
    return ok


def write_all_in_one(items):
    """עמוד מאוחד לכל הכתבים — משמש ליצירת PDF אחד גדול."""
    parts = [f"""<div style="text-align:center; padding:22vh 0 0">
  <h1 style="font-size:2.4rem; margin-bottom:.4rem">כתבי יונתן יהודה ז״ל</h1>
  <p class="muted">{e(heb_date(BIRTH))} – {e(heb_date(DEATH))}</p>
  <p class="faint" style="font-size:.9rem">{len(items)} כתבים · נכתבו בגילאי
     {min(age_at(i['created'])[0] for i in items)}–{max(age_at(i['created'])[0] for i in items)}</p>
</div>"""]
    for it in items:
        d = it["created"]
        parts.append(f"""<article class="doc" style="break-before:page; max-width:none">
  <h1 style="font-size:1.9rem">{e(it['title'])}</h1>
  {f'<p class="muted">{e(it["subtitle"])}</p>' if it['subtitle'] else ''}
  <p class="faint" style="font-size:.86rem; border-top:1px solid #ddd; border-bottom:1px solid #ddd; padding:.4rem 0">
    נכתב ב{e(heb_date(d))} · {e(greg_date(d))} · {e(age_text(d))}</p>
{blocks_html(it)}
</article>""")
    write(os.path.join(OUT, "_all.html"),
          shell(f"כל הכתבים · {SITE['short']}", f'<main class="wrap">{"".join(parts)}</main>',
                "ketavim.html"))


def make_zip(items):
    zp = os.path.join(OUT, "ketavim.zip")
    with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as z:
        for it in items:
            pdf = os.path.join(OUT, "k", f"{it['slug']}.pdf")
            if os.path.exists(pdf):
                z.write(pdf, f"כתבי יונתן/PDF/{it['title']}.pdf")
            z.write(it["path"], f"כתבי יונתן/מקור Word/{it['title']}.docx")
        allpdf = os.path.join(OUT, "ketavim-all.pdf")
        if os.path.exists(allpdf):
            z.write(allpdf, "כתבי יונתן/כל הכתבים.pdf")


# ─────────────────────────────  נכסים  ─────────────────────────────

def write(path, text):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)


def copy_assets():
    dst = os.path.join(OUT, "assets")
    os.makedirs(dst, exist_ok=True)
    for f in os.listdir(os.path.join(BUILD_DIR, "assets")):
        shutil.copy2(os.path.join(BUILD_DIR, "assets", f), os.path.join(dst, f))
    if os.path.exists(PHOTO):
        try:
            from PIL import Image
            im = Image.open(PHOTO)
            w, h = im.size
            side = min(w, h)
            im = im.crop(((w - side) // 2, 0, (w - side) // 2 + side, side))
            im.resize((900, 900), Image.LANCZOS).convert("RGB").save(
                os.path.join(dst, "yonatan.jpg"), quality=88, optimize=True)
            im2 = Image.open(PHOTO)
            im2.thumbnail((1600, 1600), Image.LANCZOS)
            im2.convert("RGB").save(os.path.join(dst, "yonatan-full.jpg"),
                                    quality=86, optimize=True)
        except Exception as exc:
            print(f"  ! עיבוד התמונה נכשל ({exc}) — מעתיק כמות שהיא")
            shutil.copy2(PHOTO, os.path.join(dst, "yonatan.jpg"))
    write(os.path.join(OUT, ".nojekyll"), "")
    write(os.path.join(OUT, "robots.txt"), "User-agent: *\nAllow: /\n")


# ─────────────────────────────  ראשי  ─────────────────────────────

def main():
    print(f"בונה את האתר · {BUILD}")
    if os.path.isdir(OUT):
        for name in os.listdir(OUT):
            if name == "media":       # שומרים MP3 שכבר הומרו
                continue
            p = os.path.join(OUT, name)
            shutil.rmtree(p) if os.path.isdir(p) else os.remove(p)
    os.makedirs(OUT, exist_ok=True)

    copy_assets()
    items = collect_writings()
    print(f"  · {len(items)} כתבים")
    songs = collect_songs()
    print(f"  · {len(songs)} שירים")
    recs = collect_recordings()
    print(f"  · {len(recs)} הקלטות")

    write_doc_pages(items)
    write_index_page(items)
    write_all_in_one(items)
    write_home(items, songs, recs)
    write_songs(songs)
    write_recordings(recs)

    print("  · מייצר PDF …")
    make_pdfs(items)
    if os.path.exists(os.path.join(OUT, "_all.html")):
        os.remove(os.path.join(OUT, "_all.html"))
    make_zip(items)

    for it in items:
        d = it["created"]
        print(f"    {it['title']:<28} {heb_date(d):<18} {greg_date(d):<18} {age_text(d)}")
    print(f"\nהאתר מוכן: {OUT}")


if __name__ == "__main__":
    main()

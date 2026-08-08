# -*- coding: utf-8 -*-
"""
קורא מסמכי Word (.docx) והופך אותם למבנה בלוקים נקי לאתר.
מטפל נכון בתיבות טקסט (textbox) שגורמות לכפילויות ב-XML של Word:
Word שומר כל תיבה פעמיים - פעם ב-mc:Choice ופעם ב-mc:Fallback.
"""
import re
import zipfile
import xml.etree.ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
MC = '{http://schemas.openxmlformats.org/markup-compatibility/2006}'


def _strip_fallbacks(root):
    """מסיר את ענפי mc:Fallback - הם עותק כפול של mc:Choice."""
    doomed = [(parent, child)
              for parent in root.iter()
              for child in parent
              if child.tag == MC + 'Fallback']
    for parent, child in doomed:
        parent.remove(child)


def read_docx(path):
    """מחזיר רשימת בלוקים: {'type': 'h'|'p', 'runs': [{'t':str,'b':bool,'i':bool}]}"""
    with zipfile.ZipFile(path) as z:
        xml = z.read('word/document.xml')
    root = ET.fromstring(xml)
    _strip_fallbacks(root)

    # מיפוי הורים, כדי לדעת לאיזו פסקה שייך כל קטע טקסט
    parent = {}
    for el in root.iter():
        for ch in el:
            parent[id(ch)] = el

    def nearest_p(el):
        cur = el
        while id(cur) in parent:
            cur = parent[id(cur)]
            if cur.tag == W + 'p':
                return cur
        return None

    paras = [p for p in root.iter(W + 'p')]
    order = {id(p): i for i, p in enumerate(paras)}
    buckets = {id(p): [] for p in paras}

    for el in root.iter():
        if el.tag not in (W + 't', W + 'tab', W + 'br'):
            continue
        p = nearest_p(el)
        if p is None:
            continue
        # סגנון ההרצה (bold / italic) נלקח מ-w:rPr של ה-w:r העוטף
        b = i = False
        r = parent.get(id(el))
        if r is not None and r.tag == W + 'r':
            rpr = r.find(W + 'rPr')
            if rpr is not None:
                b = rpr.find(W + 'b') is not None
                i = rpr.find(W + 'i') is not None
        if el.tag == W + 't':
            txt = el.text or ''
        elif el.tag == W + 'tab':
            txt = ' '
        else:
            txt = '\n'
        if txt:
            buckets[id(p)].append({'t': txt, 'b': b, 'i': i})

    blocks = []
    for p in sorted(paras, key=lambda x: order[id(x)]):
        runs = buckets[id(p)]
        text = ''.join(r['t'] for r in runs)
        if not text.strip():
            continue
        style = ''
        is_list = False
        ppr = p.find(W + 'pPr')
        if ppr is not None:
            st = ppr.find(W + 'pStyle')
            if st is not None:
                style = st.get(W + 'val') or ''
            is_list = ppr.find(W + 'numPr') is not None or 'listparagraph' in style.lower()
        kind = 'h' if style.lower().startswith('heading') else 'p'
        blocks.append({'type': kind, 'text': text, 'runs': runs, 'li': is_list})

    return _dedupe(blocks)


def _norm(s):
    return re.sub(r'\s+', ' ', s).strip()


def _dedupe(blocks):
    """
    מסיר כפילויות: (א) פסקאות זהות עוקבות, (ב) פסקה שכבר הופיעה קודם במסמך.
    בקבצים של יונתן חלק מהתוכן נכתב בתוך תיבות טקסט, ו-Word שכפל אותו.
    """
    seen = set()
    out = []
    for b in blocks:
        key = _norm(b['text'])
        if len(key) < 3:
            out.append(b)
            continue
        if key in seen:
            continue
        seen.add(key)
        out.append(b)
    # הסרת פסקאות שהן שרשור של פסקאות אחרות (תוצר של תיבת טקסט מקוננת)
    longs = [b for b in out if len(_norm(b['text'])) > 400]
    keep = []
    for b in out:
        n = _norm(b['text'])
        swallowed = any(
            other is not b and len(_norm(other['text'])) > len(n) * 1.5 and n in _norm(other['text'])
            for other in longs
        )
        if swallowed:
            continue
        keep.append(b)
    return keep

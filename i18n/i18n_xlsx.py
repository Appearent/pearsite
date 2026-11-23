#!/usr/bin/env python3
"""Export every text string of the four pages to translations.xlsx, or write an edited sheet back.

  python3 i18n/i18n_xlsx.py export   -> i18n/translations.xlsx
  python3 i18n/i18n_xlsx.py import   -> rewrites the four index.html files from the sheet
                                        (add --dry-run to only report what would change)

Rows are matched to the pages BY POSITION, so do not add, delete or reorder rows;
only edit the language columns. Structure changes belong in the HTML itself.
"""
import sys, re, urllib.parse
from html.parser import HTMLParser
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PAGES = [("EN", "index.html"), ("繁體中文", "zh-hant/index.html"),
         ("简体中文", "zh-hans/index.html"), ("日本語", "ja/index.html")]
BLOCK = {"title", "h1", "h2", "h3", "p", "li", "summary", "dt", "dd", "span", "a", "strong"}
SKIP_CLASSES = {"stages"}      # step-number pills: not text
META = {("name", "description"), ("property", "og:title"), ("property", "og:description")}

class Scan(HTMLParser):
    """Finds editable strings as (kind, label, start, end) character ranges in the source."""
    def __init__(s, src):
        super().__init__(convert_charrefs=False)
        s.src = src
        s.nl = [0]
        for m in re.finditer("\n", src): s.nl.append(m.end())
        s.items = []; s.sec = "head"; s.open = None; s.skip = 0; s.stack = []
    def off(s): l, c = s.getpos(); return s.nl[l-1] + c
    def handle_starttag(s, t, a):
        d = dict(a); start = s.off(); txt = s.get_starttag_text(); end = start + len(txt)
        cls = set((d.get("class") or "").split())
        if t == "section": s.sec = d.get("id") or (d.get("class") or "section")
        if t == "header": s.sec = "header"
        if t == "footer": s.sec = "footer"
        if t == "meta":
            for k, v in META:
                if d.get(k) == v:
                    i = s.src.index('content="', start) + 9
                    s.items.append(("meta", f"{s.sec} · {v}", i, s.src.index('"', i)))
        if t == "img" and "alt" in d and d["alt"]:
            i = s.src.index('alt="', start) + 5
            s.items.append(("attr", "logo alt text", i, s.src.index('"', i)))
        if t == "nav" and "aria-label" in d:
            i = s.src.index('aria-label="', start) + 12
            s.items.append(("attr", "nav aria-label", i, s.src.index('"', i)))
        if t == "a" and d.get("href", "").startswith("mailto:") and "subject=" in d["href"]:
            i = s.src.index("subject=", start) + 8
            s.items.append(("mailto", f"{s.sec} · email subject", i, s.src.index('"', i)))
        s.stack.append((t, cls))
        if "stages" in cls: s.skip += 1
        if t == "nav" and "lang-switch" in cls: s.skip += 1; s.langnav = True
        if s.skip and not (t == "summary") and not (s.open): pass
        if t in BLOCK and s.open is None and not (s.skip and t != "summary") :
            s.open = [t, end, 1, s.sec]
        elif s.open and t == s.open[0]: s.open[2] += 1
    def handle_endtag(s, t):
        if s.open and t == s.open[0]:
            s.open[2] -= 1
            if s.open[2] == 0:
                inner = s.src[s.open[1]:s.off()]
                if re.search(r"[^\W\d_]", inner):   # contains letters/CJK
                    s.items.append(("html", f"{s.open[3]} · <{s.open[0]}>", s.open[1], s.off()))
                s.open = None
        if s.stack:
            tt, cls = s.stack.pop()
            if "stages" in cls: s.skip -= 1
            if tt == "nav" and "lang-switch" in cls: s.skip -= 1
    # nested blocks while a block is open must not close it early: handled by same-tag depth only

def scan(path):
    src = path.read_text(encoding="utf-8"); p = Scan(src); p.feed(src)
    # drop the outer <title>-less duplicates and sort
    items = p.items
    mail = [i for i in items if i[0] == "mailto"]
    # a block that only wraps an email link has nothing to translate (the subject has its own row)
    items = [i for i in items if not (i[0] == "html" and any(i[2] <= m[2] < i[3] for m in mail))]
    items.sort(key=lambda x: x[2])
    for x, y in zip(items, items[1:]): assert x[3] <= y[2], f"overlapping strings in {path}: {x} {y}"
    return src, items

def value(kind, src, a, b):
    v = src[a:b]
    return urllib.parse.unquote(v) if kind == "mailto" else v

def export():
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    data = [scan(ROOT / f) for _, f in PAGES]
    n = {len(it) for _, it in data}
    assert len(n) == 1, f"pages have different numbers of strings: {[len(it) for _, it in data]}"
    wb = Workbook(); ws = wb.active; ws.title = "Strings"
    head = ["#", "Section · element"] + [l for l, _ in PAGES]
    ws.append(head)
    for i in range(n.pop()):
        row = [i + 1, data[0][1][i][1]]
        for src, it in data:
            k, _, a, b = it[i]; row.append(value(k, src, a, b))
        ws.append(row)
    ft = Font(name="Arial", size=10)
    thin = Side(style="thin", color="D0D0D0")
    for r in ws.iter_rows():
        for c in r:
            c.font = ft; c.alignment = Alignment(wrap_text=True, vertical="top")
            c.border = Border(bottom=thin)
    for c in ws[1]:
        c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="637159")
    for col, w in zip("ABCDEF", (6, 26, 48, 48, 48, 48)): ws.column_dimensions[col].width = w
    ws.freeze_panes = "C2"; ws.auto_filter.ref = ws.dimensions
    g = wb.create_sheet("Read me")
    for line in [
        "How to use this file",
        "1. Edit only the language columns (C to F) on the Strings sheet. Do not add, delete, sort or filter-and-move rows: rows are matched to the pages by position (column A).",
        "2. Some cells contain small HTML tags such as <b>…</b> or <span class=\"han\">…</span>. Keep the tags, translate the words between them.",
        "3. Sheet-row meaning: 'email subject' cells hold the plain subject line (the import encodes it); 'meta' rows are the search/social descriptions.",
        "4. Hand the file back and run:  python3 i18n/i18n_xlsx.py import   (add --dry-run first to preview). Then review with git diff.",
        "5. Not in this sheet (edit in the HTML): links, image names, the language dropdown, step-number pills, and the site's email address.",
    ]: g.append([line])
    g.column_dimensions["A"].width = 130; g["A1"].font = Font(name="Arial", bold=True, size=12)
    for r in g.iter_rows(min_row=2):
        r[0].font = ft; r[0].alignment = Alignment(wrap_text=True)
    out = ROOT / "i18n" / "translations.xlsx"; wb.save(out); print("wrote", out)

def do_import(dry):
    from openpyxl import load_workbook
    ws = load_workbook(ROOT / "i18n" / "translations.xlsx")["Strings"]
    rows = [r for r in ws.iter_rows(min_row=2, values_only=True) if r[0] is not None]
    for col, (label, f) in enumerate(PAGES):
        src, items = scan(ROOT / f)
        assert len(items) == len(rows), f"{f}: {len(items)} strings in page vs {len(rows)} rows in sheet"
        out, last, changed = [], 0, 0
        for (k, _, a, b), r in zip(items, rows):
            new = r[2 + col]; new = "" if new is None else str(new)
            if k == "mailto": new = urllib.parse.quote(new, safe="")
            elif k in ("meta", "attr"): new = new.replace('"', "&quot;")
            if new != src[a:b]:
                changed += 1; print(f"  {f} #{r[0]}: {src[a:b][:40]!r} -> {new[:40]!r}")
            out.append(src[last:a]); out.append(new); last = b
        out.append(src[last:])
        if not dry and changed: (ROOT / f).write_text("".join(out), encoding="utf-8")
        print(f"{f}: {changed} change(s){' (dry run)' if dry else ''}")

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "export": export()
    elif len(sys.argv) > 1 and sys.argv[1] == "import": do_import("--dry-run" in sys.argv)
    else: print(__doc__)

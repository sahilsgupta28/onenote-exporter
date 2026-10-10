#!/usr/bin/env python3
"""
Export one OneNote section (personal account) via Microsoft Graph.

Usage:
  pip3 install requests
  export ONENOTE_TOKEN="<access token from Graph Explorer>"
  python3 export_section.py '<section_id>'            # use SINGLE quotes (ids contain "!")
  python3 export_section.py '<section_id>' -o ~/Backup
  python3 export_section.py '<section_id>' --utc      # show dates in UTC (default: local time)

Output layout:
  <Section name>/
      metadata/
          pages.csv                   one row per page
          <Section name>.json         section + page metadata from the API (sub-pages nested)
      html/
          index.html                  ordered list of all pages, linked to the pages below
          1_Page name.html            one rendered, OneNote-style page per page
          1.2_Sub page.html
      pages/
          1_Page name/
              1_Page name.md          page content as markdown
              images/                 images on the page (only if there are any)
              tables/                 each table as CSV (only if there are any)
              attachments/            attached files (only if there are any)
              pages/                  sub-pages, same structure, nested further
                  1.2_Sub page name/
                      1.2_Sub page name.md ...

Sequence numbers follow the page order and indentation: top-level pages are
1, 2, 3...; sub-pages of page 2 are 2.1, 2.2...; their sub-pages 2.2.1, and so on.
"""
import argparse
import csv
import html
import json
import mimetypes
import os
import re
import sys
import time
import warnings
from datetime import datetime, timezone
from html.parser import HTMLParser
from string import Template
from urllib.parse import quote, urlparse

warnings.filterwarnings("ignore", message=".*OpenSSL.*")
import requests  # noqa: E402

GRAPH = "https://graph.microsoft.com/v1.0"
BETA = "https://graph.microsoft.com/beta"
HEADERS = {}


# --------------------------------------------------------------------------
# Graph helpers
# --------------------------------------------------------------------------
class GraphError(Exception):
    def __init__(self, status, body, url):
        super().__init__(f"HTTP {status} for {url}\n{body[:600]}")
        self.status = status


def get(url, stream=False, auth=True, tries=4):
    """GET with retries on throttling (429) and server errors (5xx)."""
    headers = HEADERS if auth else {}
    r = None
    for attempt in range(tries):
        r = requests.get(url, headers=headers, stream=stream, timeout=120)
        if r.status_code == 429 or r.status_code >= 500:
            try:
                wait = int(r.headers.get("Retry-After", 2 ** attempt * 2))
            except ValueError:
                wait = 2 ** attempt * 2
            print(f"    (HTTP {r.status_code}, retrying in {wait}s)")
            time.sleep(wait)
            continue
        if not r.ok:
            raise GraphError(r.status_code, r.text, url)
        return r
    raise GraphError(r.status_code, r.text, url)


def paged(url):
    """Yield every item of a paged Graph collection."""
    while url:
        data = get(url).json()
        yield from data.get("value", [])
        url = data.get("@odata.nextLink")


# --------------------------------------------------------------------------
# File / name helpers
# --------------------------------------------------------------------------
def safe_name(name, fallback="Untitled", maxlen=100):
    name = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", name or "")
    name = re.sub(r"\s+", " ", name).strip(" .")
    return (name or fallback)[:maxlen].rstrip(" .") or fallback


def unique_name(name, used):
    """Return name, or 'name (2)', 'name (3)'... if already used (case-insensitive)."""
    candidate, n = name, 1
    while candidate.lower() in used:
        n += 1
        candidate = f"{name} ({n})"
    used.add(candidate.lower())
    return candidate


def unique_file(directory, filename):
    base, ext = os.path.splitext(filename)
    candidate, n = filename, 1
    while os.path.exists(os.path.join(directory, candidate)):
        n += 1
        candidate = f"{base} ({n}){ext}"
    return candidate


def write_text(path, text):
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)


def write_json(path, obj):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, indent=2, ensure_ascii=False)
        f.write("\n")


def write_csv(path, rows):
    # utf-8-sig so Excel on Mac/Windows detects UTF-8 correctly
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        csv.writer(f).writerows(rows)


def save_stream(resp, path):
    with open(path, "wb") as f:
        for chunk in resp.iter_content(65536):
            f.write(chunk)


def ext_for(content_type, fallback=""):
    ct = (content_type or "").split(";")[0].strip().lower()
    ext = mimetypes.guess_extension(ct) if ct else None
    ext = {".jpe": ".jpg", ".jpeg": ".jpg"}.get(ext, ext)
    return ext or fallback


def sniff_image_ext(data):
    """Detect the image type from its first bytes (Graph often sends octet-stream)."""
    if data.startswith(b"\x89PNG"):
        return ".png"
    if data.startswith(b"\xff\xd8"):
        return ".jpg"
    if data[:6] in (b"GIF87a", b"GIF89a"):
        return ".gif"
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return ".webp"
    if data[:2] == b"BM":
        return ".bmp"
    if data[:4] in (b"II*\x00", b"MM\x00*"):
        return ".tif"
    if data.lstrip()[:4] in (b"<svg", b"<?xm"):
        return ".svg"
    return None


def image_ext(data, content_type, hint):
    """File extension for an image: sniffed bytes, then declared type, else .png."""
    ext = sniff_image_ext(data)
    if ext:
        return ext
    for declared in (hint, content_type):
        ct = (declared or "").split(";")[0].strip().lower()
        if ct.startswith("image/"):
            ext = ext_for(ct)
            if ext:
                return ext
    return ".png"


def is_graph(url):
    return urlparse(url).hostname == "graph.microsoft.com"


# --------------------------------------------------------------------------
# Images and attachments
# --------------------------------------------------------------------------
def attr(tag, name):
    """Raw (still HTML-escaped) value of an attribute inside a tag string."""
    m = re.search(
        rf'(?<![\w-]){re.escape(name)}\s*=\s*(?:"([^"]*)"|\'([^\']*)\')', tag, re.I
    )
    if not m:
        return None
    return m.group(1) if m.group(1) is not None else m.group(2)


def localize_resources(content, page_dir):
    """Download images and attachments; point the HTML at the local copies."""
    images, attachments, cache = [], [], {}

    def repl_img(m):
        tag = m.group(0)
        full, small = attr(tag, "data-fullres-src"), attr(tag, "src")
        pick = full or small
        if not pick:
            return tag
        url = html.unescape(pick)
        if not url.startswith("http"):
            return tag
        if url not in cache:
            try:
                r = get(url, stream=True, auth=is_graph(url))
            except GraphError as e:
                print(f"    ! image download failed: {e}")
                return tag
            hint = attr(tag, "data-fullres-src-type") or attr(tag, "data-src-type")
            data = r.content
            ext = image_ext(data, r.headers.get("Content-Type"), hint)
            name = f"image-{len(images) + 1:02d}{ext}"
            os.makedirs(os.path.join(page_dir, "images"), exist_ok=True)
            with open(os.path.join(page_dir, "images", name), "wb") as f:
                f.write(data)
            images.append(name)
            cache[url] = f"images/{name}"
        # replace the longer URL first: the small src can be a prefix of the full-res one
        for raw in sorted({full, small} - {None, ""}, key=len, reverse=True):
            tag = tag.replace(raw, quote(cache[url]))
        return tag

    def repl_obj(m):
        tag = m.group(0)
        fname, raw = attr(tag, "data-attachment"), attr(tag, "data")
        if not fname or not raw:
            return tag
        url = html.unescape(raw)
        if not url.startswith("http"):
            return tag
        if url not in cache:
            att_dir = os.path.join(page_dir, "attachments")
            try:
                r = get(url, stream=True, auth=is_graph(url))
            except GraphError as e:
                print(f"    ! attachment download failed: {e}")
                return tag
            os.makedirs(att_dir, exist_ok=True)
            local = unique_file(att_dir, safe_name(html.unescape(fname), "attachment", 150))
            save_stream(r, os.path.join(att_dir, local))
            attachments.append(local)
            cache[url] = f"attachments/{local}"
        return tag.replace(raw, quote(cache[url]))

    content = re.sub(r"<img\b[^>]*>", repl_img, content, flags=re.I | re.S)
    content = re.sub(r"<object\b[^>]*>", repl_obj, content, flags=re.I | re.S)
    return content, images, attachments


# --------------------------------------------------------------------------
# HTML -> text / markdown / tables
# --------------------------------------------------------------------------
SKIP_TAGS = {"script", "style", "head", "title"}
HEADINGS = {"h1", "h2", "h3", "h4", "h5", "h6"}


class TextExtractor(HTMLParser):
    BLOCKS = HEADINGS | {"p", "div", "br", "li", "tr", "table", "ul", "ol"}

    def __init__(self):
        super().__init__()
        self.out, self.skip, self.cells = [], 0, 0

    def handle_starttag(self, tag, attrs):
        if tag in SKIP_TAGS:
            self.skip += 1
        elif self.skip:
            return
        elif tag in ("td", "th"):
            self.cells += 1
        elif tag in self.BLOCKS:
            self.out.append(" " if self.cells else "\n")

    def handle_endtag(self, tag):
        if tag in SKIP_TAGS:
            self.skip = max(0, self.skip - 1)
        elif self.skip:
            return
        elif tag in ("td", "th"):
            self.cells = max(0, self.cells - 1)
            if not self.cells:
                self.out.append("\t")
        elif tag in self.BLOCKS:
            self.out.append(" " if self.cells else "\n")

    def handle_data(self, data):
        if not self.skip:
            self.out.append(re.sub(r"\s+", " ", data))

    def result(self):
        lines = []
        for line in "".join(self.out).split("\n"):
            cells = [c.strip() for c in line.split("\t")]
            line = "\t".join(cells).rstrip("\t")
            if line.strip():
                lines.append(line)
        return "\n".join(lines) + "\n"


def clean_cell(s):
    return re.sub(r"\s+", " ", s).strip()


def md_table(rows):
    n = max(len(r) for r in rows)
    rows = [r + [""] * (n - len(r)) for r in rows]

    def esc(c):
        return c.replace("|", "\\|")

    lines = ["| " + " | ".join(esc(c) for c in rows[0]) + " |",
             "| " + " | ".join(["---"] * n) + " |"]
    lines += ["| " + " | ".join(esc(c) for c in r) + " |" for r in rows[1:]]
    return "\n".join(lines)


class MarkdownConverter(HTMLParser):
    """Small HTML -> Markdown converter that also collects top-level tables."""

    def __init__(self):
        super().__init__()
        self.stack = [[]]       # output buffers; a new one is pushed per table cell
        self.skip = 0
        self.lists = []         # [kind, counter]
        self.tables = []        # finished top-level tables (list of rows)
        self.table_stack = []
        self.fmt = []           # open bold/italic/strike markers
        self.links = []

    # -- helpers
    def w(self, s):
        self.stack[-1].append(s)

    def in_cell(self):
        return len(self.stack) > 1

    def last_char(self):
        for chunk in reversed(self.stack[-1]):
            if chunk:
                return chunk[-1]
        return "\n"

    def para_break(self):
        self.w(" " if (self.in_cell() or self.lists) else "\n\n")

    def open_fmt(self, tag, marker):
        self.fmt.append((tag, marker))
        if marker:
            self.w(marker)

    def close_fmt(self, tag):
        for i in range(len(self.fmt) - 1, -1, -1):
            if self.fmt[i][0] == tag:
                _, marker = self.fmt.pop(i)
                if not marker:
                    return
                closer, buf = marker[::-1], self.stack[-1]
                if buf and buf[-1] == marker:      # empty emphasis: drop it
                    buf.pop()
                elif buf and buf[-1].strip() and buf[-1] != buf[-1].rstrip(" "):
                    buf[-1] = buf[-1].rstrip(" ")  # keep trailing space outside marker
                    self.w(closer + " ")
                else:
                    self.w(closer)
                return

    # -- parser callbacks
    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag in SKIP_TAGS:
            self.skip += 1
            return
        if self.skip:
            return
        if tag in HEADINGS:
            self.para_break()
            if not self.in_cell() and not self.lists:
                self.w("#" * int(tag[1]) + " ")
        elif tag in ("p", "div"):
            self.para_break()
            todo = a.get("data-tag") or ""
            if "to-do" in todo:
                box = "[x] " if "completed" in todo else "[ ] "
                self.w(box if (self.lists or self.in_cell()) else "- " + box)
        elif tag == "br":
            self.w(" " if (self.in_cell() or self.lists) else "  \n")
        elif tag in ("b", "strong"):
            self.open_fmt(tag, "**")
        elif tag in ("i", "em"):
            self.open_fmt(tag, "*")
        elif tag in ("s", "strike", "del"):
            self.open_fmt(tag, "~~")
        elif tag == "span":
            style = (a.get("style") or "").replace(" ", "").lower()
            marks = []
            if "font-weight:bold" in style:
                marks.append("**")
            if "font-style:italic" in style:
                marks.append("*")
            if "line-through" in style:
                marks.append("~~")
            self.open_fmt("span", "".join(marks))
        elif tag == "a":
            self.links.append((a.get("href"), len(self.stack[-1])))
        elif tag == "img":
            src = a.get("src") or a.get("data-fullres-src") or ""
            alt = (a.get("alt") or "").replace("\n", " ")
            img = f"![{alt}]({src})"
            self.w(img if (self.in_cell() or self.lists) else f"\n\n{img}\n\n")
        elif tag == "object":
            name, data = a.get("data-attachment"), a.get("data")
            if name:
                ref = f"[Attachment: {name}]({data})" if data else f"Attachment: {name}"
                self.w(ref if (self.in_cell() or self.lists) else f"\n\n{ref}\n\n")
        elif tag in ("ul", "ol"):
            self.w(" " if self.in_cell() else ("" if self.lists else "\n"))
            self.lists.append([tag, 0])
        elif tag == "li":
            if self.in_cell():
                self.w(" \u2022 ")
            else:
                depth = max(len(self.lists) - 1, 0)
                marker = "- "
                if self.lists:
                    kind, n = self.lists[-1]
                    self.lists[-1][1] += 1
                    if kind == "ol":
                        marker = f"{n + 1}. "
                self.w("\n" + "  " * depth + marker)
        elif tag == "table":
            self.table_stack.append({"rows": [], "row": None})
        elif tag == "tr":
            if self.table_stack:
                self.table_stack[-1]["row"] = []
        elif tag in ("td", "th"):
            if self.table_stack:
                self.stack.append([])

    def handle_endtag(self, tag):
        if tag in SKIP_TAGS:
            self.skip = max(0, self.skip - 1)
            return
        if self.skip:
            return
        if tag in HEADINGS or tag in ("p", "div"):
            self.para_break()
        elif tag in ("b", "strong", "i", "em", "s", "strike", "del", "span"):
            self.close_fmt(tag)
        elif tag == "a" and self.links:
            href, idx = self.links.pop()
            buf = self.stack[-1]
            text = "".join(buf[idx:]).strip()
            del buf[idx:]
            if text:
                self.w(f"[{text}]({href})" if href else text)
        elif tag in ("ul", "ol"):
            if self.lists:
                self.lists.pop()
            self.w(" " if self.in_cell() else ("" if self.lists else "\n"))
        elif tag in ("td", "th"):
            if self.table_stack and self.in_cell():
                cell = clean_cell("".join(self.stack.pop()))
                t = self.table_stack[-1]
                if t["row"] is None:
                    t["row"] = []
                t["row"].append(cell)
        elif tag == "tr":
            if self.table_stack:
                t = self.table_stack[-1]
                if t["row"]:
                    t["rows"].append(t["row"])
                t["row"] = None
        elif tag == "table" and self.table_stack:
            t = self.table_stack.pop()
            rows = t["rows"]
            if not rows:
                return
            if self.table_stack:                   # nested table: flatten into parent cell
                self.w(" / ".join(" ; ".join(r) for r in rows))
            else:
                self.tables.append(rows)
                self.w("\n\n" + md_table(rows) + "\n\n")

    def handle_data(self, data):
        if self.skip:
            return
        text = re.sub(r"\s+", " ", data)
        if self.last_char() in " \n":
            text = text.lstrip(" ")
        if text:
            self.w(text)

    def result(self):
        lines = []
        for line in "".join(self.stack[0]).split("\n"):
            hard_break = line.endswith("  ") and line.strip()
            line = line.rstrip()
            lines.append(line + "  " if hard_break else line)
        return re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip() + "\n"


# --------------------------------------------------------------------------
# Dates and HTML rendering
# --------------------------------------------------------------------------
def fmt_dt(iso, utc=False):
    """'2020-05-16T09:57:00.000Z' -> 'Saturday, May 16, 2020 2:57 AM PDT' (local time)."""
    if not iso:
        return ""
    try:
        dt = datetime.fromisoformat(iso.replace("Z", "+00:00"))
    except ValueError:
        return iso
    dt = dt.astimezone(timezone.utc) if utc else dt.astimezone()
    hour = dt.hour % 12 or 12
    ampm = "AM" if dt.hour < 12 else "PM"
    return (f"{dt.strftime('%A, %B')} {dt.day}, {dt.year} {hour}:{dt.minute:02d} {ampm} "
            f"{dt.tzname() or ''}").strip()


def body_of(page_html):
    m = re.search(r"<body[^>]*>(.*)</body>", page_html, re.S | re.I)
    body = m.group(1) if m else page_html
    return re.sub(r"<script\b.*?</script>", "", body, flags=re.S | re.I)


PAGE_CSS = r"""
*{box-sizing:border-box}
body{margin:0;background:#f3f2f1;color:#201f1e;font-family:"Segoe UI","Segoe UI Web (West European)",-apple-system,BlinkMacSystemFont,Roboto,"Helvetica Neue",Arial,sans-serif}
.topbar{background:#7719aa;color:#fff;padding:10px 24px;font-size:14px;display:flex;gap:16px;flex-wrap:wrap;align-items:center}
.topbar a{color:#fff;text-decoration:none;font-weight:600}
.topbar a:hover{text-decoration:underline}
.crumbs{opacity:.92}
.crumbs a{font-weight:400}
.page{max-width:920px;margin:24px auto;background:#fff;padding:36px 56px 72px;box-shadow:0 1px 4px rgba(0,0,0,.18);min-height:60vh}
.title{font-family:"Segoe UI Light","Segoe UI",-apple-system,"Helvetica Neue",Arial,sans-serif;font-weight:300;font-size:2.25rem;line-height:1.2;margin:0;padding-bottom:6px;border-bottom:1px solid #c8c6c4;display:inline-block;min-width:50%;max-width:100%;overflow-wrap:anywhere}
.dates{color:#767676;font-size:.9rem;margin:10px 0 32px;display:flex;gap:24px;flex-wrap:wrap}
.dates b{font-weight:600;margin-right:6px}
.content{font-size:1rem;line-height:1.55;overflow-wrap:anywhere}
.content div[style*="absolute"]{position:static!important;left:auto!important;top:auto!important;width:auto!important;margin:0 0 16px!important}
.content p{margin:0 0 .6em}
.content img{max-width:100%;height:auto}
.content table{border-collapse:collapse;margin:8px 0;max-width:100%}
.content td,.content th{border:1px solid #c8c6c4;padding:4px 8px;vertical-align:top}
.content a{color:#0f6cbd}
.content ul,.content ol{margin:.3em 0 .6em;padding-left:1.8em}
.content [data-tag*="to-do"]::before{content:"\2610\00a0"}
.content [data-tag*="to-do:completed"]::before{content:"\2611\00a0"}
.pager{max-width:920px;margin:0 auto 32px;padding:0 4px;display:flex;justify-content:space-between;gap:16px;font-size:.9rem}
.pager a{color:#7719aa;text-decoration:none}
@media print{.topbar,.pager{display:none}body{background:#fff}.page{box-shadow:none;margin:0}}
"""

INDEX_CSS = r"""
*{box-sizing:border-box}
body{margin:0;background:#f3f2f1;color:#201f1e;font-family:"Segoe UI","Segoe UI Web (West European)",-apple-system,BlinkMacSystemFont,Roboto,"Helvetica Neue",Arial,sans-serif}
.topbar{background:#7719aa;color:#fff;padding:10px 24px;font-size:14px;font-weight:600}
.topbar a{color:#fff;text-decoration:none}
.topbar a:hover{text-decoration:underline}
.count{color:#767676;font-size:.85em;margin-left:.5em}
.page{max-width:920px;margin:24px auto;background:#fff;padding:36px 56px 56px;box-shadow:0 1px 4px rgba(0,0,0,.18)}
h1{font-weight:300;font-size:2.25rem;margin:0 0 4px}
.meta{color:#767676;font-size:.9rem;margin:0 0 28px}
ol.toc,ol.toc ol{list-style:none;margin:0;padding-left:1.6em}
ol.toc{padding-left:0}
ol.toc li{margin:.25em 0;line-height:1.45}
.seq{display:inline-block;min-width:3.2em;color:#767676;font-variant-numeric:tabular-nums}
a{color:#0f6cbd;text-decoration:none}
a:hover{text-decoration:underline}
em{color:#a4262c}
"""

PAGE_TEMPLATE = Template("""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>$title_text</title>
<style>$css</style>
</head>
<body>
<header class="topbar">
  <a href="index.html">&#9776; $section</a>
  <nav class="crumbs">$crumbs</nav>
</header>
<main class="page">
  <h1 class="title">$title</h1>
  <div class="dates"><span><b>Created</b>$created</span><span><b>Modified</b>$modified</span></div>
  <div class="content">
$body
  </div>
</main>
<nav class="pager"><span>$prev</span><span>$next</span></nav>
</body>
</html>
""")

INDEX_TEMPLATE = Template("""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>$title_text</title>
<style>$css</style>
</head>
<body>
<header class="topbar">$topbar</header>
<main class="page">
  <h1>$title</h1>
  <p class="meta">$meta</p>
$toc
</main>
</body>
</html>
""")


def html_href(entry, prefix=""):
    return prefix + quote(entry["html_name"])


def render_page_html(e, section_name, body, created, modified, prev_e, next_e):
    links, anc = [], e["parent"]
    while anc:
        links.append(f'<a href="{html_href(anc)}">{html.escape(anc["title"])}</a>')
        anc = anc["parent"]
    crumbs = "".join(l + " &rsaquo; " for l in reversed(links)) + html.escape(e["title"])

    def pager(other, arrow_left):
        if not other or other["status"] == "failed":
            return ""
        label = f'{other["seq"]} {html.escape(other["title"])}'
        return (f'<a href="{html_href(other)}">&larr; {label}</a>' if arrow_left
                else f'<a href="{html_href(other)}">{label} &rarr;</a>')

    return PAGE_TEMPLATE.substitute(
        title_text=html.escape(e["title"]), css=PAGE_CSS, section=html.escape(section_name),
        crumbs=crumbs, title=html.escape(e["title"]),
        created=html.escape(created), modified=html.escape(modified), body=body,
        prev=pager(prev_e, True), next=pager(next_e, False))


def render_index(section, entries, utc, notebook_link=None):
    roots = [e for e in entries if not e["parent"]]

    def render(nodes):
        out = ["<ol class=\"toc\">" if nodes and not nodes[0]["parent"] else "<ol>"]
        for e in nodes:
            label = html.escape(e["title"])
            if e["status"] == "ok":
                link = f'<a href="{quote(e["html_name"])}">{label}</a>'
            else:
                link = f"{label} <em>(not exported)</em>"
            out.append(f'<li><span class="seq">{e["seq"]}</span>{link}')
            if e["children"]:
                out.append(render(e["children"]))
            out.append("</li>")
        out.append("</ol>")
        return "\n".join(out)

    notebook = (section.get("parentNotebook") or {}).get("displayName")
    name = section.get("displayName") or "Section"
    bits = ([f"Notebook: {html.escape(notebook)}"] if notebook else [])
    bits.append(f"{len(entries)} pages")
    bits.append("Exported " + html.escape(fmt_dt(datetime.now(timezone.utc).isoformat(), utc)))
    topbar = "OneNote export"
    if notebook_link:     # (notebook name, relative href to the notebook's index.html)
        topbar = f'<a href="{html.escape(notebook_link[1])}">&larr; {html.escape(notebook_link[0])}</a>'
    return INDEX_TEMPLATE.substitute(
        title_text=html.escape(name), css=INDEX_CSS, title=html.escape(name), topbar=topbar,
        meta=" &middot; ".join(bits), toc=render(roots) if roots else "<p>No pages.</p>")


# --------------------------------------------------------------------------
# Export logic
# --------------------------------------------------------------------------
def build_entries(pages, pages_dir):
    """Number the pages (1, 1.1, 1.2.1 ...) and decide each page's folder."""
    entries, stack, counters = [], [], {}
    for page in pages:
        lvl = page.get("level") or 0
        title = page.get("title") or "Untitled"
        while stack and stack[-1]["level"] >= lvl:
            stack.pop()
        parent = stack[-1] if stack else None
        key = parent["seq"] if parent else ""
        counters[key] = counters.get(key, 0) + 1
        seq = f"{key}.{counters[key]}" if parent else str(counters[key])
        base = f"{seq}_{safe_name(title)}"      # unique because the sequence is unique
        folder = base
        rel_parts = (parent["rel_parts"] + ["pages", folder]) if parent else [folder]
        e = {
            "page": page, "title": title, "level": lvl, "order": page.get("order", ""),
            "parent": parent, "children": [], "seq": seq,
            "rel_parts": rel_parts, "dir": os.path.join(pages_dir, *rel_parts),
            "md_name": base + ".md", "html_name": base + ".html",
            "path": (parent["path"] + " / " + title) if parent else title,
            "status": "pending", "stats": None,
        }
        if parent:
            parent["children"].append(e)
        stack.append(e)
        entries.append(e)
    return entries


def page_size(page_dir):
    """Bytes in the page's own files (md, images, tables, attachments), not its sub-pages."""
    total = 0
    for dirpath, dirnames, filenames in os.walk(page_dir):
        if dirpath == page_dir and "pages" in dirnames:
            dirnames.remove("pages")
        for f in filenames:
            total += os.path.getsize(os.path.join(dirpath, f))
    return total


def export_page(e, html_dir, section_name, prev_e, next_e, utc):
    page, title, page_dir = e["page"], e["title"], e["dir"]
    os.makedirs(page_dir, exist_ok=True)
    created = fmt_dt(page.get("createdDateTime"), utc)
    modified = fmt_dt(page.get("lastModifiedDateTime"), utc)

    r = get(f"{GRAPH}/me/onenote/pages/{page['id']}/content")
    raw = r.content.decode("utf-8", "replace")
    content, images, attachments = localize_resources(raw, page_dir)

    # word count from plain text
    te = TextExtractor()
    te.feed(content)
    words = len(te.result().split())

    # markdown (+ tables as CSV)
    mc = MarkdownConverter()
    mc.feed(content)
    write_text(os.path.join(page_dir, e["md_name"]),
               f"# {title}\n\nCreated: {created}  \nModified: {modified}\n\n---\n\n{mc.result()}")
    if mc.tables:
        os.makedirs(os.path.join(page_dir, "tables"), exist_ok=True)
        for i, rows in enumerate(mc.tables, 1):
            write_csv(os.path.join(page_dir, "tables", f"table-{i:02d}.csv"), rows)

    # html: re-point image/attachment links so they work from the html/ folder
    prefix = "../pages/" + "/".join(quote(p) for p in e["rel_parts"]) + "/"
    linked = re.sub(r'((?<![\w-])(?:src|data-fullres-src|data)\s*=\s*")((?:images|attachments)/)',
                    lambda m: m.group(1) + prefix + m.group(2), content, flags=re.I)
    write_text(os.path.join(html_dir, e["html_name"]),
               render_page_html(e, section_name, body_of(linked), created, modified,
                                prev_e, next_e))

    return {"words": words, "size": page_size(page_dir), "images": len(images),
            "tables": len(mc.tables), "attachments": len(attachments)}


PAGES_CSV_HEADER = ["Number", "Sequence", "Page name", "Created", "Modified", "WordCount",
                    "Size (bytes)", "Level", "Order", "Images", "Tables", "Attachments",
                    "Path"]


def pages_csv_rows(entries):
    rows = [PAGES_CSV_HEADER]
    for n, e in enumerate(entries, 1):
        st = e["stats"] or {}
        rows.append([
            n, e["seq"], e["title"], e["page"].get("createdDateTime", ""),
            e["page"].get("lastModifiedDateTime", ""), st.get("words", ""), st.get("size", ""),
            e["level"], e["order"], st.get("images", ""), st.get("tables", ""),
            st.get("attachments", ""),
            "pages/" + "/".join(e["rel_parts"]),   # page folder, relative to the root folder
        ])
    return rows


def section_json(section, entries):
    """Section metadata plus every page's API metadata, sub-pages nested under their parent."""
    def node(e):
        d = dict(e["page"])                      # as returned by the API (no page content)
        d["pages"] = [node(c) for c in e["children"]]
        return d

    doc = dict(section)
    doc["pages"] = [node(e) for e in entries if not e["parent"]]
    return doc


class ExportError(Exception):
    """A section could not be exported; the caller decides whether to continue."""


class TokenExpired(ExportError):
    """Graph answered 401: the token is expired or invalid."""


def fail(err, what):
    cls = TokenExpired if getattr(err, "status", None) == 401 else ExportError
    raise cls(f"{what}:\n{err}")


def set_token(raw):
    token = re.sub(r"\s+", "", raw or "").strip("\"'")
    if token.lower().startswith("bearer"):
        token = token[6:]
    HEADERS["Authorization"] = f"Bearer {token}"


def export_section(sid, outdir=".", utc=False, folder_name=None, notebook_link=None):
    """Export one section. Returns a dict describing the result.

    folder_name   -- override the root folder name (default: the section name)
    notebook_link -- (notebook name, relative href) to show a back-link on index.html
    """
    # Section details
    base = f"{GRAPH}/me/onenote/sections/{sid}"
    try:
        try:
            section = get(base + "?$expand=parentNotebook($select=id,displayName)").json()
        except GraphError as e:
            if e.status == 401:
                raise
            section = get(base).json()
    except GraphError as e:
        fail(e, "Could not fetch section")

    # Root folder and its subfolders
    sec_name = section.get("displayName") or sid
    root = os.path.join(outdir, folder_name or safe_name(sec_name, "Section"))
    meta_dir, pages_dir, html_dir = (os.path.join(root, d) for d in ("metadata", "pages", "html"))
    for d in (meta_dir, pages_dir, html_dir):
        os.makedirs(d, exist_ok=True)
    print(f"Section: {sec_name}")

    # Pages list. "pagelevel=true" is REQUIRED to get indentation level and order;
    # without it Graph omits both. Try several request forms until levels show up.
    # (No $top: it caused HTTP 500 earlier.)
    fields = ("id,title,createdDateTime,lastModifiedDateTime,level,order,createdByAppId,"
              "contentUrl,self,links,userTags")
    variants = [
        f"{base}/pages?$select={fields}&pagelevel=true",
        f"{base}/pages?pagelevel=true",
        f"{BETA}/me/onenote/sections/{sid}/pages?pagelevel=true",
    ]
    pages = None
    try:
        for i, url in enumerate(variants, 1):
            try:
                result = list(paged(url))
            except GraphError as e:
                if e.status == 401:
                    raise
                print(f"  (page list request {i} failed: HTTP {e.status})")
                continue
            if pages is None:
                pages = result
            if any("level" in p for p in result):
                pages = result
                break
    except GraphError as e:
        fail(e, "Could not list pages")
    if pages is None:
        raise ExportError("Could not list pages (all request forms failed).")
    if pages and not any("level" in p for p in pages):
        print("  ! Graph returned no level/order values; pages will be numbered flat "
              "(no sub-page nesting).")
    pages.sort(key=lambda p: p.get("order", 0))

    entries = build_entries(pages, pages_dir)
    write_json(os.path.join(meta_dir, safe_name(sec_name, "section") + ".json"),
               section_json(section, entries))
    print(f"Found {len(entries)} pages -> {os.path.abspath(root)}")

    # Export each page
    aborted = False
    for n, e in enumerate(entries, 1):
        prev_e = entries[n - 2] if n > 1 else None
        next_e = entries[n] if n < len(entries) else None
        print(f"  [{n}/{len(entries)}] {e['seq']}  {'  ' * e['level']}{e['title']}")
        try:
            e["stats"] = export_page(e, html_dir, sec_name, prev_e, next_e, utc)
            e["status"] = "ok"
        except GraphError as err:
            e["status"] = "failed"
            print(f"    ! {err}")
            if err.status == 401:
                print("\nToken expired. Get a fresh one from Graph Explorer and re-run.")
                aborted = True
        except Exception as err:  # keep going on unexpected per-page problems
            e["status"] = "failed"
            print(f"    ! {err}")
        if aborted:
            break

    # metadata/pages.csv and html/index.html
    write_csv(os.path.join(meta_dir, "pages.csv"), pages_csv_rows(entries))
    write_text(os.path.join(html_dir, "index.html"),
               render_index(section, entries, utc, notebook_link))

    done = [e["stats"] for e in entries if e["stats"]]
    totals = {k: sum(st[k] for st in done)
              for k in ("words", "size", "images", "tables", "attachments")}
    totals.update(pages=len(entries), ok=sum(1 for e in entries if e["status"] == "ok"))
    return {"section": section, "root": root, "index": os.path.join(html_dir, "index.html"),
            "entries": entries, "totals": totals, "aborted": aborted}


def main():
    ap = argparse.ArgumentParser(description="Export a OneNote section to local files.")
    ap.add_argument("section_id", help="Section ID (use single quotes in zsh/bash)")
    ap.add_argument("-o", "--outdir", default=".", help="parent folder (default: current)")
    ap.add_argument("--utc", action="store_true",
                    help="show page created/modified dates in UTC instead of local time")
    args = ap.parse_args()

    raw = os.environ.get("ONENOTE_TOKEN") or sys.exit("Set ONENOTE_TOKEN first")
    set_token(raw)
    try:
        res = export_section(args.section_id, args.outdir, args.utc)
    except ExportError as e:
        sys.exit(str(e))

    t = res["totals"]
    print(f"Done: {t['ok']} of {t['pages']} pages exported.")
    for e in res["entries"]:
        if e["status"] != "ok":
            print(f"  not exported: {e['seq']} {e['title']}")
    print(f"Open: {os.path.abspath(res['index'])}")


if __name__ == "__main__":
    main()

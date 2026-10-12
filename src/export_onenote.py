#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""
Export OneNote content (personal account) via Microsoft Graph.

Usage:
  pip3 install requests
  export ONENOTE_TOKEN='<access token from Graph Explorer>'      # use SINGLE quotes
  python3 export_onenote.py N '<notebook id>'      # a whole notebook
  python3 export_onenote.py S '<section id>'       # one section
  python3 export_onenote.py P '<page id>'          # one page

Options:
  -o DIR      parent folder for the export (default: current folder)
  --utc       show dates in UTC instead of local time
  --no-ink    do not ask Graph for handwriting (InkML)
  --ocr       run local OCR (needs the `tesseract` program) on downloaded images
  --render M  save pages as PDF/PNG: auto = pages with handwriting, OCR text or only
              images (default), all, or none (needs Chrome/Chromium/Edge/Brave;
              or set ONENOTE_BROWSER=/path/to/browser)
  --render-as pdf|png|both   format of the saved page (default pdf)

Output (notebook mode; section and page mode use the same layout starting at the
section / page folder):
  <Notebook name>/
      index.html                numbered list of sections
      errors.csv                only if something could not be downloaded or handled
      metadata/
          sections.csv, <Notebook name>.json
          logs.txt              everything printed to the console, all runs appended
          progress.json         only while a run is unfinished (used to resume)
      sections/<seq>_<Section>/ (section groups are folders holding their sections)
          metadata/ (pages.csv, <Section name>.json)   html/ (index.html + one file per page)
          pages/<seq>_<Page>/
              <seq>_<Page>.md      (+ <seq>_<Page>.pdf / .png for handwriting / OCR pages)
              images/  tables/  attachments/  ink/      (only when there is something)
              pages/    sub-pages, same structure

Safe to stop and re-run
  * If Graph keeps answering HTTP 429 (too many requests), a request is retried up to
    10 times (waits 2, 4, 8, 16, then 20 seconds). After that the run stops cleanly and
    asks you to run the same command later.
  * Run the same command again to continue: finished sections and pages are skipped,
    the page that was in progress is deleted and downloaded again.
  * Nothing is overwritten until its replacement has been fully downloaded.
  * errors.csv lists every page / section / item that could not be downloaded (these
    are retried on the next run) and every data format the script cannot handle
    (kept for your review).
"""
import argparse
import base64
import csv
import hashlib
import html
import io
import json
import mimetypes
import os
import re
import shutil
import subprocess
import sys
import time
import traceback
import warnings
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from email import message_from_bytes
from html.parser import HTMLParser
from pathlib import Path
from string import Template
from urllib.parse import quote, unquote_to_bytes, urlparse

warnings.filterwarnings("ignore", message=".*OpenSSL.*")
import requests  # noqa: E402

GRAPH = "https://graph.microsoft.com/v1.0"
BETA = "https://graph.microsoft.com/beta"
HEADERS = {}
MS_HOSTS = {"graph.microsoft.com", "www.onenote.com", "onenote.com"}

MAX_429_RETRIES = 10     # HTTP 429: give up (and ask the user to continue later) after this many
MAX_BACKOFF = 20         # seconds; waits are 2, 4, 8, 16, 20, 20 ...
MAX_SERVER_TRIES = 4     # HTTP 5xx / network errors


# --------------------------------------------------------------------------
# Exceptions and run state
# --------------------------------------------------------------------------
class GraphError(Exception):
    def __init__(self, status, body, url):
        super().__init__(f"HTTP {status} for {url}\n{(body or '')[:600]}")
        self.status, self.url, self.body = status, url, (body or "")[:600]


class ExportError(Exception):
    """One section (or the notebook) could not be exported; others can continue."""


class StopRun(Exception):
    """The whole run must stop now. Progress is saved so a later run can continue."""
    result = None


class RateLimitExceeded(StopRun):
    pass


class TokenExpired(StopRun):
    pass


OPTS = {"utc": False, "ink": True, "ocr": False, "render": "auto", "render_as": "pdf"}
RUN = {"kind": "", "id": ""}
CTX = {"section": None, "page": None}      # (name, id) / page entry, used when logging errors
INK_STATE = {"unsupported": False}
OCR_STATE = {"warned": False}
ERRORS = None
PROGRESS = None


class Tee:
    """Copies everything printed to the console into metadata/logs.txt."""

    def __init__(self, stream):
        self.stream, self.file, self.buffer = stream, None, []

    def attach(self, path):
        if self.file:
            return
        os.makedirs(os.path.dirname(path), exist_ok=True)
        self.file = open(path, "a", encoding="utf-8", buffering=1)
        stamp = datetime.now().astimezone().isoformat(timespec="seconds")
        self.file.write(f"\n===== Run started {stamp} =====\n")
        self.file.write("".join(self.buffer))
        self.buffer = []

    def write(self, s):
        self.stream.write(s)
        if self.file:
            self.file.write(s)
        else:
            self.buffer.append(s)
        return len(s)

    def flush(self):
        self.stream.flush()
        if self.file:
            self.file.flush()

    def isatty(self):
        return False

    def close(self):
        if self.file:
            self.file.close()
            self.file = None


ERROR_HEADER = ["Time", "Category", "Section", "Section ID", "Page", "Page ID", "Sequence",
                "Item", "URL", "Error"]
# download / page / section problems are retried on the next run;
# format / ocr problems are kept for review only
RETRY_CATEGORIES = {"download", "page", "section"}


class ErrorLog:
    """Collects problems and writes errors.csv in the root folder."""

    def __init__(self):
        self.path, self.rows, self.prior, self.reprocessed = None, [], [], set()

    def attach(self, root):
        if self.path:
            return
        self.path = os.path.join(root, "errors.csv")
        if os.path.exists(self.path):
            try:
                with open(self.path, newline="", encoding="utf-8-sig") as f:
                    rows = list(csv.reader(f))
                self.prior = [r for r in rows[1:] if len(r) == len(ERROR_HEADER)]
            except OSError:
                self.prior = []

    @staticmethod
    def key(row):
        return row[5] or "section:" + row[3]

    def retry_pages(self):
        return {r[5] for r in self.prior if r[5] and r[1] in RETRY_CATEGORIES}

    def add(self, category, item, error, url=""):
        sec, page = CTX["section"], CTX["page"]
        text = " ".join(str(error).split())[:500]
        self.rows.append([
            datetime.now().astimezone().isoformat(timespec="seconds"), category,
            sec[0] if sec else "", sec[1] if sec else "",
            page["title"] if page else "", page["page"].get("id", "") if page else "",
            page["seq"] if page else "", item, url, text])
        print(f"    ! [{category}] {item}: {text[:160]}")
        self.flush()
        if PROGRESS:
            PROGRESS.flush()          # write progress as soon as an error is encountered

    def flush(self):
        if not self.path:
            return
        keep = [r for r in self.prior if self.key(r) not in self.reprocessed]
        rows = keep + self.rows
        if rows:
            write_csv(self.path, [ERROR_HEADER] + rows)
        elif os.path.exists(self.path):
            os.remove(self.path)      # everything that failed before has now been fixed

    def count(self):
        keep = [r for r in self.prior if self.key(r) not in self.reprocessed]
        return len(keep) + len(self.rows)

    def retry_count(self):
        keep = [r for r in self.prior if self.key(r) not in self.reprocessed]
        return len([r for r in keep + self.rows if r[1] in RETRY_CATEGORIES])


class Progress:
    """In-memory index of finished sections / pages, saved to metadata/progress.json."""

    def __init__(self):
        self.path, self.data = None, None

    def attach(self, meta_dir, kind, target_id):
        if self.path:
            return
        self.path = os.path.join(meta_dir, "progress.json")
        self.data = {"version": 1, "kind": kind, "id": target_id, "sections": {}}
        if os.path.exists(self.path):
            try:
                with open(self.path, encoding="utf-8") as f:
                    old = json.load(f)
                if (old.get("version"), old.get("kind"), old.get("id")) == (1, kind, target_id):
                    self.data = old
                    print("Found progress from an earlier run; continuing where it stopped.")
            except (OSError, ValueError):
                print("Progress file unreadable; falling back to checking the files on disk.")

    def section(self, sid):
        return self.data["sections"].setdefault(
            sid, {"status": "pending", "page_count": None, "done": {}, "in_progress": None,
                  "totals": None})

    def flush(self):
        if self.path and self.data is not None:
            self.data["updated"] = datetime.now().astimezone().isoformat(timespec="seconds")
            write_json(self.path, self.data)

    def clear(self):
        if self.path and os.path.exists(self.path):
            os.remove(self.path)


def init_run(kind, target):
    global ERRORS, PROGRESS
    ERRORS, PROGRESS = ErrorLog(), Progress()
    RUN.update(kind=kind, id=target)
    CTX.update(section=None, page=None)
    INK_STATE["unsupported"] = False
    OCR_STATE["warned"] = False
    RENDER_STATE["warned"] = False


def attach_run_files(root):
    """Start logs.txt / errors.csv / progress.json in this root (first caller wins)."""
    meta = os.path.join(root, "metadata")
    if isinstance(sys.stdout, Tee):
        sys.stdout.attach(os.path.join(meta, "logs.txt"))
    ERRORS.attach(root)
    PROGRESS.attach(meta, RUN["kind"], RUN["id"])


# --------------------------------------------------------------------------
# Graph helpers: retry policy
# --------------------------------------------------------------------------
def is_graph(url):
    return (urlparse(url).hostname or "").lower() == "graph.microsoft.com"


def is_ms_host(url):
    return (urlparse(url).hostname or "").lower() in MS_HOSTS


def backoff(n, retry_after=None):
    """Wait before retry n (1, 2, 3 ...): 2, 4, 8, 16, then 20 seconds at most."""
    wait = min(MAX_BACKOFF, 2 ** n)
    try:
        wait = min(MAX_BACKOFF, max(wait, float(retry_after)))
    except (TypeError, ValueError):
        pass
    return wait


def get(url, stream=False, auth=True):
    headers = HEADERS if auth else {}
    retries_429, server_tries = 0, 0
    while True:
        try:
            r = requests.get(url, headers=headers, stream=stream, timeout=120)
        except requests.RequestException as exc:
            server_tries += 1
            if server_tries >= MAX_SERVER_TRIES:
                raise GraphError(0, f"{type(exc).__name__}: {exc}", url)
            wait = backoff(server_tries)
            print(f"    (network error, retry {server_tries}/{MAX_SERVER_TRIES - 1} in {wait:g}s)")
            time.sleep(wait)
            continue
        if r.status_code == 429:
            if retries_429 >= MAX_429_RETRIES:
                raise RateLimitExceeded(
                    f"HTTP 429 (too many requests) persisted after {MAX_429_RETRIES} retries")
            retries_429 += 1
            wait = backoff(retries_429, r.headers.get("Retry-After"))
            print(f"    (HTTP 429, retry {retries_429}/{MAX_429_RETRIES} in {wait:g}s)")
            time.sleep(wait)
            continue
        if r.status_code >= 500:
            server_tries += 1
            if server_tries >= MAX_SERVER_TRIES:
                raise GraphError(r.status_code, r.text, url)
            wait = backoff(server_tries)
            print(f"    (HTTP {r.status_code}, retry {server_tries}/{MAX_SERVER_TRIES - 1} in {wait:g}s)")
            time.sleep(wait)
            continue
        if r.status_code == 401 and is_graph(url):
            raise TokenExpired("HTTP 401: the access token is expired or invalid")
        if not r.ok:
            raise GraphError(r.status_code, r.text, url)
        return r


def paged(url):
    """Yield every item of a paged Graph collection."""
    while url:
        data = get(url).json()
        yield from data.get("value", [])
        url = data.get("@odata.nextLink")


def set_token(raw):
    token = re.sub(r"\s+", "", raw or "").strip("\"'")
    if token.lower().startswith("bearer"):
        token = token[6:]
    HEADERS["Authorization"] = f"Bearer {token}"


# --------------------------------------------------------------------------
# File helpers (every write goes to a .part file first, then is renamed)
# --------------------------------------------------------------------------
def atomic_write(path, data, binary=False):
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    tmp = path + ".part"
    if binary:
        with open(tmp, "wb") as f:
            f.write(data)
    else:
        with open(tmp, "w", encoding="utf-8", newline="") as f:
            f.write(data)
    os.replace(tmp, path)


def write_text(path, text):
    atomic_write(path, text)


def write_json(path, obj):
    atomic_write(path, json.dumps(obj, indent=2, ensure_ascii=False) + "\n")


def write_csv(path, rows):
    # utf-8 BOM so Excel on Mac/Windows detects UTF-8 correctly
    buf = io.StringIO(newline="")
    csv.writer(buf).writerows(rows)
    atomic_write(path, "\ufeff" + buf.getvalue())


def save_stream(resp, path):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".part"
    with open(tmp, "wb") as f:
        for chunk in resp.iter_content(65536):
            f.write(chunk)
    os.replace(tmp, path)




# --------------------------------------------------------------------------
# Name / image helpers
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


def attr(tag, name):
    """Raw (still HTML-escaped) value of an attribute inside a tag string."""
    m = re.search(
        rf'(?<![\w-]){re.escape(name)}\s*=\s*(?:"([^"]*)"|\'([^\']*)\')', tag, re.I
    )
    if not m:
        return None
    return m.group(1) if m.group(1) is not None else m.group(2)


# --------------------------------------------------------------------------
# Images, attachments and other embedded resources
# --------------------------------------------------------------------------
RES_ID_RE = re.compile(r"/resources/([^/?#]+)/\$value", re.I)
DATA_URI_RE = re.compile(r"^data:([\w/+.-]*)(;base64)?,(.*)$", re.I | re.S)

# Elements the script cannot turn into files / markdown. They are logged to errors.csv
# (category "format") so you can review the affected pages.
UNHANDLED_TAGS = {
    "iframe": "embedded video / web content cannot be downloaded (link kept in markdown and html)",
    "embed": "embedded plug-in content is not supported",
    "video": "video element is not supported",
    "audio": "audio element is not supported",
    "canvas": "canvas drawing is not supported",
    "svg": "inline SVG is kept in html but dropped from markdown",
}


def resource_candidates(url):
    """URLs to try for one resource: as given, then the documented Graph form
    https://graph.microsoft.com/v1.0/me/onenote/resources/{id}/$value"""
    cands = [url]
    m = RES_ID_RE.search(url)
    if m:
        canon = f"{GRAPH}/me/onenote/resources/{m.group(1)}/$value"
        if canon not in cands:
            cands.append(canon)
    return cands


def fetch_resource(url):
    last = None
    for cand in resource_candidates(url):
        try:
            return get(cand, stream=True, auth=is_ms_host(cand))
        except GraphError as e:
            last = e
    raise last


def reorder_split_images(content):
    """OneNote splits large images (web captures, PDF printouts) into several <img>
    elements that share data-id; data-index is the top-to-bottom order and the tags may
    arrive out of order. Put each run of such tags back in order."""
    def index_of(tag):
        try:
            return int(attr(tag, "data-index"))
        except (TypeError, ValueError):
            return 0

    def fix(m):
        tags = re.findall(r"<img\b[^>]*>", m.group(0), flags=re.I | re.S)
        groups = {}
        for t in tags:
            groups.setdefault(attr(t, "data-id") or "", []).append(t)
        ordered = [t for g in groups.values() for t in sorted(g, key=index_of)]
        return "\n".join(ordered) + "\n"

    return re.sub(r"(?:<img\b[^>]*\bdata-index\s*=[^>]*>\s*){2,}", fix, content,
                  flags=re.I | re.S)


def ocr_image(path):
    """Text found in an image by the local tesseract program ('' if none / unavailable)."""
    exe = shutil.which("tesseract")
    if not exe:
        if not OCR_STATE["warned"]:
            OCR_STATE["warned"] = True
            ERRORS.add("ocr", "tesseract", "tesseract is not installed, so OCR was skipped "
                       "for all images (e.g. brew install tesseract, then re-run with --ocr)")
        return ""
    name = os.path.basename(path)
    try:
        r = subprocess.run([exe, path, "stdout"], capture_output=True, text=True, timeout=180)
    except (OSError, subprocess.SubprocessError) as e:
        ERRORS.add("ocr", "image", f"OCR failed for {name}: {e}")
        return ""
    if r.returncode != 0:
        ERRORS.add("ocr", "image", f"tesseract exit {r.returncode} for {name}: {r.stderr.strip()[:150]}")
        return ""
    return r.stdout.strip()


def localize_resources(content, page_dir):
    """Download the page's images and attachments into page_dir and re-point the HTML
    at the local copies. Returns (html, image names, attachment names). Anything that
    cannot be downloaded stays linked to its original URL and is logged."""
    images, attachments, cache = [], [], {}

    def store_image(data, content_type, hint):
        ext = sniff_image_ext(data)
        if not ext:
            ext = image_ext(data, content_type, hint)
            ERRORS.add("format", "image", f"unrecognized image data (content type "
                       f"{content_type or 'unknown'}); saved as {ext}")
        name = f"image-{len(images) + 1:02d}{ext}"
        path = os.path.join(page_dir, "images", name)
        atomic_write(path, data, binary=True)
        images.append(name)
        return name, path

    def repl_img(m):
        tag = m.group(0)
        full, small = attr(tag, "data-fullres-src"), attr(tag, "src")
        raws = [r for r in dict.fromkeys([full, small]) if r]   # full-res first, then optimized
        if not raws:
            return tag
        key = html.unescape(raws[0])
        ocr_text = ""
        if key in cache:
            rel, ocr_text = cache[key]
        else:
            rel, last_err, path = None, None, None
            for raw in raws:
                url = html.unescape(raw)
                if url.lower().startswith("data:"):
                    mm = DATA_URI_RE.match(url)
                    try:
                        data = (base64.b64decode(mm.group(3)) if mm.group(2)
                                else unquote_to_bytes(mm.group(3)))
                    except Exception as e:     # bad data URI: try the next candidate
                        last_err = e
                        continue
                    name, path = store_image(data, mm.group(1), None)
                    rel = f"images/{name}"
                    break
                if url.lower().startswith(("http://", "https://")):
                    try:
                        r = fetch_resource(url)
                    except GraphError as e:
                        last_err = e
                        continue
                    hint = attr(tag, "data-fullres-src-type") or attr(tag, "data-src-type")
                    name, path = store_image(r.content, r.headers.get("Content-Type"), hint)
                    rel = f"images/{name}"
                    break
                ERRORS.add("format", "image", f"unsupported image source: {url[:80]}", url[:200])
                return tag
            if rel is None:
                ERRORS.add("download", "image", last_err, getattr(last_err, "url", raws[0])[:300])
                return tag
            if OPTS["ocr"]:
                ocr_text = ocr_image(path)
                if ocr_text:
                    atomic_write(os.path.splitext(path)[0] + ".txt", ocr_text + "\n")
            cache[key] = (rel, ocr_text)
        for raw in sorted(set(raws), key=len, reverse=True):   # longer URL first: src can be a prefix
            tag = tag.replace(raw, quote(rel))
        if ocr_text:
            esc = html.escape(ocr_text, quote=True).replace("\n", "&#10;")
            tag = tag.replace("<img", f'<img data-ocr-text="{esc}"', 1)
        return tag

    def repl_obj(m):
        tag = m.group(0)
        raw = attr(tag, "data")
        if not raw:
            return tag
        url = html.unescape(raw)
        if not url.lower().startswith(("http://", "https://")):
            return tag
        if url not in cache:
            try:
                r = fetch_resource(url)
            except GraphError as e:
                ERRORS.add("download", "attachment", e, e.url[:300])
                return tag
            fname = attr(tag, "data-attachment")
            if fname:
                name = html.unescape(fname)
            else:
                name = f"object-{len(attachments) + 1:02d}" + ext_for(
                    attr(tag, "type") or r.headers.get("Content-Type"), ".bin")
                ERRORS.add("format", "object", f"embedded object has no file name; saved as {name}")
            att_dir = os.path.join(page_dir, "attachments")
            os.makedirs(att_dir, exist_ok=True)
            local = unique_file(att_dir, safe_name(name, "attachment", 150))
            save_stream(r, os.path.join(att_dir, local))
            attachments.append(local)
            cache[url] = (f"attachments/{local}", "")
        return tag.replace(raw, quote(cache[url][0]))

    content = reorder_split_images(content)
    content = re.sub(r"<img\b[^>]*>", repl_img, content, flags=re.I | re.S)
    content = re.sub(r"<object\b[^>]*>", repl_obj, content, flags=re.I | re.S)

    for tagname, why in UNHANDLED_TAGS.items():
        found = re.findall(rf"<{tagname}\b[^>]*>", content, flags=re.I | re.S)
        if found:
            src = attr(found[0], "data-original-src") or attr(found[0], "src") or ""
            ERRORS.add("format", tagname, f"{len(found)} <{tagname}> element(s): {why}",
                       html.unescape(src)[:300])
    return content, images, attachments


# --------------------------------------------------------------------------
# Handwriting (InkML)
# --------------------------------------------------------------------------
def split_multipart(resp):
    """(html bytes or None, [InkML bytes]) from a page-content response."""
    ctype = resp.headers.get("Content-Type", "")
    if "multipart" not in ctype.lower():
        return resp.content, []
    msg = message_from_bytes(b"MIME-Version: 1.0\r\nContent-Type: "
                             + ctype.encode("latin-1", "replace") + b"\r\n\r\n" + resp.content)
    page_html, inks = None, []
    for part in msg.walk():
        if part.is_multipart():
            continue
        ctype_part, payload = part.get_content_type(), part.get_payload(decode=True) or b""
        if ctype_part == "text/html":
            page_html = payload
        elif "inkml" in ctype_part:
            inks.append(payload)
        else:
            ERRORS.add("format", "multipart part",
                       f"unhandled part type {ctype_part} ({len(payload)} bytes), not saved")
    return page_html, inks


def fetch_page_content(pid):
    """(html bytes, [InkML bytes]). Asks for handwriting with includeInkML=true; if the
    API refuses that option, plain content is used for the rest of the run."""
    base = f"{GRAPH}/me/onenote/pages/{pid}/content"
    if OPTS["ink"] and not INK_STATE["unsupported"]:
        try:
            page_html, inks = split_multipart(get(base + "?includeInkML=true"))
            if page_html is not None:
                return page_html, inks
        except GraphError as e:
            if 400 <= e.status < 500:
                INK_STATE["unsupported"] = True
                ERRORS.add("format", "ink", "the API rejected includeInkML (HTTP "
                           f"{e.status}); handwriting cannot be exported and was not requested "
                           f"for the rest of this run: {e.body[:100]}", e.url[:300])
            else:
                print(f"    (ink request failed with HTTP {e.status}; using plain content)")
    return get(base).content, []


def _local(tag):
    return tag.rsplit("}", 1)[-1]


def decode_trace(text, xi, yi):
    """Points of one InkML trace (handles the ! ' \" explicit / difference prefixes)."""
    pts, state = [], None
    for chunk in text.split(","):
        vals = chunk.split()
        if not vals:
            continue
        if state is None:
            state = [{"v": 0.0, "d": 0.0, "m": "!"} for _ in range(max(len(vals), xi + 1, yi + 1))]
        for i, tok in enumerate(vals[:len(state)]):
            s = state[i]
            if tok and tok[0] in "!'\"":
                s["m"], tok = tok[0], tok[1:]
            if tok in ("", "*", "?"):
                continue
            try:
                num = float(tok)
            except ValueError:
                continue
            if s["m"] == "!":
                s["v"], s["d"] = num, 0.0
            elif s["m"] == "'":
                s["d"] = num
                s["v"] += num
            else:
                s["d"] += num
                s["v"] += s["d"]
        pts.append((state[xi]["v"], state[yi]["v"]))
    return pts


def inkml_to_svg(data):
    """Approximate SVG drawing of InkML strokes (None if there are no strokes)."""
    root = ET.fromstring(data)
    xml_id = "{http://www.w3.org/XML/1998/namespace}id"
    brushes, channels = {}, []
    for el in root.iter():
        name = _local(el.tag)
        if name == "brush":
            props = {c.get("name"): c.get("value") for c in el if _local(c.tag) == "brushProperty"}
            brushes[el.get(xml_id) or el.get("id") or ""] = props
        elif name == "traceFormat" and not channels:
            channels = [c.get("name") for c in el if _local(c.tag) == "channel"]
    xi = channels.index("X") if "X" in channels else 0
    yi = channels.index("Y") if "Y" in channels else 1
    strokes = []
    for el in root.iter():
        if _local(el.tag) == "trace":
            pts = decode_trace(el.text or "", xi, yi)
            if pts:
                strokes.append((pts, brushes.get((el.get("brushRef") or "").lstrip("#"), {})))
    if not strokes:
        return None
    xs = [p[0] for pts, _ in strokes for p in pts]
    ys = [p[1] for pts, _ in strokes for p in pts]
    minx, miny = min(xs), min(ys)
    w, h = max(max(xs) - minx, 1.0), max(max(ys) - miny, 1.0)
    pad, sw = max(w, h) * 0.02, max(w, h) / 600
    shapes = []
    for pts, props in strokes:
        color = props.get("color") or "#000000"
        if not re.match(r"^#[0-9a-fA-F]{3,8}$", color):
            color = "#000000"
        if len(pts) == 1:
            shapes.append(f'<circle cx="{pts[0][0]:g}" cy="{pts[0][1]:g}" r="{sw:g}" fill="{color}"/>')
        else:
            coords = " ".join(f"{x:g},{y:g}" for x, y in pts)
            shapes.append(f'<polyline points="{coords}" fill="none" stroke="{color}" '
                          f'stroke-width="{sw:g}" stroke-linecap="round" stroke-linejoin="round"/>')
    vw, vh = w + 2 * pad, h + 2 * pad
    return (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="{minx - pad:g} {miny - pad:g} '
            f'{vw:g} {vh:g}" width="800" height="{800 * vh / vw:.0f}">\n'
            + "\n".join(shapes) + "\n</svg>\n")


def save_ink(inks, page_dir):
    """Keep the raw InkML and an approximate SVG. Returns relative SVG paths."""
    svgs = []
    for i, data in enumerate(inks, 1):
        suffix = "" if len(inks) == 1 else f"-{i:02d}"
        atomic_write(os.path.join(page_dir, "ink", f"ink{suffix}.inkml"), data, binary=True)
        try:
            svg = inkml_to_svg(data)
        except Exception as e:
            ERRORS.add("format", "ink", f"InkML could not be converted to SVG ({e}); raw InkML kept")
            continue
        if svg:
            atomic_write(os.path.join(page_dir, "ink", f"ink{suffix}.svg"), svg)
            svgs.append(f"ink/ink{suffix}.svg")
        else:
            ERRORS.add("format", "ink", "InkML has no drawable strokes; raw InkML kept")
    if inks:
        ERRORS.add("format", "ink", f"{len(inks)} handwriting part(s) saved as InkML + approximate "
                   "SVG and shown at the end of the page; their original position is not reproduced")
    return svgs


# --------------------------------------------------------------------------
# Saving a page as PDF / PNG (pages with handwriting, OCR text or only images)
# --------------------------------------------------------------------------
BROWSER_NAMES = ["google-chrome", "google-chrome-stable", "chromium", "chromium-browser",
                 "chrome", "microsoft-edge", "msedge", "brave-browser"]
BROWSER_APPS = [
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Chromium.app/Contents/MacOS/Chromium",
    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    "/Applications/Brave Browser.app/Contents/MacOS/Brave Browser",
]
RENDER_STATE = {"warned": False}


def find_browser():
    """A Chromium-based browser that can print a page headlessly (or None)."""
    env = os.environ.get("ONENOTE_BROWSER")
    if env and (shutil.which(env) or os.path.exists(env)):
        return shutil.which(env) or env
    for name in BROWSER_NAMES:
        path = shutil.which(name)
        if path:
            return path
    return next((p for p in BROWSER_APPS if os.path.exists(p)), None)


def render_page(html_path, out_dir, base_name):
    """Render the saved page (html/<seq>_<name>.html) to PDF and/or PNG inside out_dir.
    Returns the file names created. Nothing is raised: problems go to errors.csv."""
    exe = find_browser()
    if not exe:
        if not RENDER_STATE["warned"]:
            RENDER_STATE["warned"] = True
            ERRORS.add("format", "render", "no Chrome / Chromium / Edge / Brave found, so pages "
                       "with handwriting or OCR text could not be saved as PDF/PNG (install one, "
                       "or set ONENOTE_BROWSER=/path/to/browser, and re-run)")
        return []
    fmt = OPTS["render_as"]
    targets = []
    if fmt in ("pdf", "both"):
        targets.append(("pdf", ["--print-to-pdf={out}", "--no-pdf-header-footer"]))
    if fmt in ("png", "both"):
        targets.append(("png", ["--screenshot={out}", "--window-size=1100,3200",
                                "--hide-scrollbars"]))
    url = Path(html_path).resolve().as_uri()
    sandbox = ["--no-sandbox"] if (hasattr(os, "geteuid") and os.geteuid() == 0) else []
    created = []
    for ext, flags in targets:
        final = os.path.join(out_dir, f"{base_name}.{ext}")
        tmp, err = final + ".part", "no output produced"
        for headless in ("--headless=new", "--headless"):
            cmd = ([exe, headless, "--disable-gpu", "--virtual-time-budget=8000"]
                   + [f.format(out=tmp) for f in flags] + sandbox + [url])
            try:
                r = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
            except (OSError, subprocess.SubprocessError) as e:
                err = str(e)
                continue
            if os.path.exists(tmp) and os.path.getsize(tmp) > 0:
                os.replace(tmp, final)
                created.append(os.path.basename(final))
                break
            err = (r.stderr or "").strip()[-150:] or err
        else:
            if os.path.exists(tmp):
                os.remove(tmp)
            ERRORS.add("format", "render", f"could not save the page as {ext.upper()}: {err}")
    return created


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

    @staticmethod
    def tag_prefix(value):
        """Label for note tags other than to-do: data-tag="important" -> '**[Important]** '."""
        labels = []
        for t in (value or "").split(","):
            t = t.strip()
            if t and not t.startswith("to-do"):
                labels.append(t.replace(":", " ").replace("-", " ").strip().title())
        return "".join(f"**[{label}]** " for label in labels)

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
            self.w(self.tag_prefix(a.get("data-tag")))
        elif tag in ("p", "div"):
            self.para_break()
            todo = a.get("data-tag") or ""
            if "to-do" in todo:
                box = "[x] " if "completed" in todo else "[ ] "
                self.w(box if (self.lists or self.in_cell()) else "- " + box)
            else:
                self.w(self.tag_prefix(todo))
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
            dt = a.get("data-tag") or ""
            if "to-do" in dt:
                self.w("[x] " if "completed" in dt else "[ ] ")
            else:
                self.w(self.tag_prefix(dt))
            self.open_fmt("span", "".join(marks))
        elif tag == "a":
            self.links.append((a.get("href"), len(self.stack[-1])))
        elif tag == "img":
            src = a.get("src") or a.get("data-fullres-src") or ""
            alt = (a.get("alt") or "").replace("\n", " ")
            source = (a.get("data-render-original-src") or "").replace('"', "%22")
            img = f'![{alt}]({src} "Source: {source}")' if source else f"![{alt}]({src})"
            ocr = (a.get("data-ocr-text") or "").strip()
            if ocr and not (self.in_cell() or self.lists):
                img += "\n\n" + "\n".join("> " + ln for ln in ["**OCR text**"] + ocr.splitlines())
            self.w(img if (self.in_cell() or self.lists) else f"\n\n{img}\n\n")
        elif tag == "object":
            name, data = a.get("data-attachment"), a.get("data")
            if name:
                ref = f"[Attachment: {name}]({data})" if data else f"Attachment: {name}"
                self.w(ref if (self.in_cell() or self.lists) else f"\n\n{ref}\n\n")
        elif tag == "iframe":
            url = a.get("data-original-src") or a.get("src") or ""
            if url:
                ref = f"[Video / embedded content: {url}]({url})"
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
.content img[data-index]{display:block;margin:0}
.content iframe{max-width:100%}
.content p.attachment{margin:.4em 0}
.content p.attachment a::before{content:"\1F4CE\00a0"}
.content details.ocr{margin:.3em 0 1em;color:#605e5c;font-size:.9rem}
.content details.ocr pre{white-space:pre-wrap;margin:.4em 0}
.ink{margin-top:2em;border-top:1px solid #c8c6c4;padding-top:1em}
.ink h2{font-weight:400;font-size:1.1rem;color:#605e5c}
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
$extra
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


def render_page_html(e, section_name, body, created, modified, prev_e, next_e, extra=""):
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
        created=html.escape(created), modified=html.escape(modified), body=body, extra=extra,
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
# Section numbering and tables of contents
# --------------------------------------------------------------------------
def build_entries(pages, pages_dir, flat=False):
    """Number the pages (1, 1.1, 1.2.1 ...) and decide each page's folder."""
    entries, stack, counters = [], [], {}
    for page in pages:
        lvl = 0 if flat else (page.get("level") or 0)
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


PAGES_CSV_HEADER = ["Number", "Sequence", "Page name", "Created", "Modified", "WordCount",
                    "Size (bytes)", "Level", "Order", "Images", "Tables", "Attachments",
                    "Page ID", "Path"]


def pages_csv_rows(entries):
    rows = [PAGES_CSV_HEADER]
    for n, e in enumerate(entries, 1):
        st = e["stats"] or {}
        rows.append([
            n, e["seq"], e["title"], e["page"].get("createdDateTime", ""),
            e["page"].get("lastModifiedDateTime", ""), st.get("words", ""), st.get("size", ""),
            e["level"], e["order"], st.get("images", ""), st.get("tables", ""),
            st.get("attachments", ""), e["page"].get("id", ""),
            "pages/" + "/".join(e["rel_parts"]),   # page folder, relative to the root folder
        ])
    return rows


def pages_csv_current(meta_dir):
    """True if metadata/pages.csv exists and already has the current columns (e.g. Page ID).
    A finished section whose pages.csv is out of date is not skipped, so the file gets
    rewritten (no page is downloaded again)."""
    try:
        with open(os.path.join(meta_dir, "pages.csv"), encoding="utf-8-sig", newline="") as f:
            return next(csv.reader(f)) == PAGES_CSV_HEADER
    except (OSError, StopIteration):
        return False


def section_json(section, entries):
    """Section metadata plus every page's API metadata, sub-pages nested under their parent."""
    def node(e):
        d = dict(e["page"])                      # as returned by the API (no page content)
        d["pages"] = [node(c) for c in e["children"]]
        return d

    doc = dict(section)
    doc["pages"] = [node(e) for e in entries if not e["parent"]]
    return doc


# --------------------------------------------------------------------------
# Building and saving one page
# --------------------------------------------------------------------------
def prepare_html_body(content, prefix):
    """Page HTML as shown in html/<seq>_<name>.html (the part inside <body>)."""
    # <object/> and <iframe/> are not void elements in HTML5: written self-closed they would
    # swallow the rest of the page, so expand them.
    content = re.sub(r"<(object|iframe)\b([^>]*?)\s*/>", r"<\1\2></\1>", content,
                     flags=re.I | re.S)
    # image / attachment links are relative to the page folder; html/ lives next to pages/
    content = re.sub(r'((?<![\w-])(?:src|data-fullres-src|data)\s*=\s*")((?:images|attachments)/)',
                     lambda m: m.group(1) + prefix + m.group(2), content, flags=re.I)

    def attachment_link(m):
        tag = m.group(0)
        href = attr(tag, "data")
        if not href:
            return tag
        name = html.unescape(attr(tag, "data-attachment") or "") or "attachment"
        return f'<p class="attachment"><a href="{href}">{html.escape(name)}</a></p>'

    content = re.sub(r"<object\b[^>]*>\s*(?:</object>)?", attachment_link, content,
                     flags=re.I | re.S)

    def with_ocr(m):
        tag = m.group(0)
        raw = attr(tag, "data-ocr-text")
        if not raw:
            return tag
        return (tag + '<details class="ocr"><summary>OCR text</summary><pre>'
                + html.escape(html.unescape(raw)) + "</pre></details>")

    content = re.sub(r"<img\b[^>]*>", with_ocr, content, flags=re.I | re.S)
    return body_of(content)


def build_page(e, tmp_dir):
    """Download one page into tmp_dir (nothing outside tmp_dir is touched)."""
    page, title = e["page"], e["title"]
    created = fmt_dt(page.get("createdDateTime"), OPTS["utc"])
    modified = fmt_dt(page.get("lastModifiedDateTime"), OPTS["utc"])

    raw_bytes, inks = fetch_page_content(page["id"])
    os.makedirs(tmp_dir, exist_ok=True)
    content, images, attachments = localize_resources(raw_bytes.decode("utf-8", "replace"), tmp_dir)

    te = TextExtractor()
    te.feed(content)
    words = len(te.result().split())

    mc = MarkdownConverter()
    mc.feed(content)
    ink_svgs = save_ink(inks, tmp_dir) if inks else []
    md = f"# {title}\n\nCreated: {created}  \nModified: {modified}\n\n---\n\n{mc.result()}"
    if ink_svgs:
        md += "\n## Handwriting\n\n" + "\n\n".join(
            f"![Handwriting]({quote(p)})\n\n[Open {os.path.basename(p)} in a browser]({quote(p)})"
            for p in ink_svgs) + "\n"
    write_text(os.path.join(tmp_dir, e["md_name"]), md)

    for i, rows in enumerate(mc.tables, 1):
        write_csv(os.path.join(tmp_dir, "tables", f"table-{i:02d}.csv"), rows)

    stats = {"words": words, "size": page_size(tmp_dir), "images": len(images),
             "tables": len(mc.tables), "attachments": len(attachments)}
    return {"content": content, "ink": ink_svgs, "created": created, "modified": modified,
            "stats": stats,
            "needs_render": bool(inks) or "data-ocr-text=" in content or (words == 0 and bool(images))}


def commit_page_dir(tmp, final):
    """Replace the page's files by the freshly downloaded ones. Sub-pages (final/pages)
    are never touched."""
    os.makedirs(final, exist_ok=True)
    for name in os.listdir(final):
        if name == "pages":
            continue
        path = os.path.join(final, name)
        shutil.rmtree(path) if os.path.isdir(path) else os.remove(path)
    for name in os.listdir(tmp):
        shutil.move(os.path.join(tmp, name), os.path.join(final, name))
    os.rmdir(tmp)


def write_page_html(e, html_dir, section_name, prev_e, next_e, built):
    prefix = "../pages/" + "/".join(quote(p) for p in e["rel_parts"]) + "/"
    body = prepare_html_body(built["content"], prefix)
    extra = ""
    if built["ink"]:
        items = []
        for i, p in enumerate(built["ink"], 1):
            href = f"{prefix}{quote(p)}"
            label = f" {i}" if len(built["ink"]) > 1 else ""
            items.append(
                f'<p><a href="{href}" target="_blank" rel="noopener">'
                f'Open handwriting{label} in a browser (SVG)</a></p>'
                f'<a href="{href}" target="_blank" rel="noopener">'
                f'<img src="{href}" alt="Handwriting{label}"></a>')
        extra = f'<div class="ink"><h2>Handwriting</h2>{"".join(items)}</div>'
    # written last: its existence marks the page as completely exported
    write_text(os.path.join(html_dir, e["html_name"]),
               render_page_html(e, section_name, body, built["created"], built["modified"],
                                prev_e, next_e, extra))


def stats_from_disk(e):
    """Rebuild a finished page's numbers from its files (used when no progress is saved)."""
    d = e["dir"]

    def count(sub, keep=lambda f: True):
        p = os.path.join(d, sub)
        return len([f for f in os.listdir(p) if keep(f)]) if os.path.isdir(p) else 0

    try:
        with open(os.path.join(d, e["md_name"]), encoding="utf-8") as f:
            words = len(f.read().split("\n---\n", 1)[-1].split())
    except OSError:
        words = 0
    return {"words": words, "size": page_size(d),
            "images": count("images", lambda f: not f.endswith(".txt")),
            "tables": count("tables"), "attachments": count("attachments")}


# --------------------------------------------------------------------------
# Export one section (resumable)
# --------------------------------------------------------------------------
def export_section(sid, outdir=".", folder_name=None, notebook_link=None, only_page=None,
                   section_hint=None):
    """Export a section, or only one of its pages (only_page = page id).
    Returns a result dict. Raises ExportError (this section failed) or StopRun (stop the
    whole run; progress is saved and exc.result holds the partial result)."""
    utc = OPTS["utc"]
    CTX["section"], CTX["page"] = (sid, sid), None
    if section_hint is not None and folder_name and not only_page and PROGRESS.data is not None:
        # Efficient resume: a section that the progress file says is finished needs no API call
        st = PROGRESS.section(sid)
        root0 = os.path.join(outdir, folder_name)
        index0 = os.path.join(root0, "html", "index.html")
        if (st["status"] == "done" and os.path.exists(index0)
                and pages_csv_current(os.path.join(root0, "metadata"))):
            print(f"Section: {section_hint.get('displayName') or sid} - already complete, skipped")
            return {"section": section_hint, "root": root0, "index": index0, "entries": [],
                    "totals": st["totals"], "aborted": False, "complete": True, "skipped": True}
    base = f"{GRAPH}/me/onenote/sections/{sid}"
    try:
        try:
            section = get(base + "?$expand=parentNotebook($select=id,displayName)").json()
        except GraphError:
            section = get(base).json()
    except GraphError as e:
        ERRORS.add("section", "section details", e, e.url[:300])
        raise ExportError(f"Could not fetch section:\n{e}")

    sec_name = section.get("displayName") or sid
    CTX["section"] = (sec_name, sid)
    root = os.path.join(outdir, folder_name or safe_name(sec_name, "Section"))
    meta_dir, pages_dir, html_dir = (os.path.join(root, d) for d in ("metadata", "pages", "html"))
    for d in (meta_dir, pages_dir, html_dir):
        os.makedirs(d, exist_ok=True)
    attach_run_files(root)

    state = PROGRESS.section(sid)
    index_path = os.path.join(html_dir, "index.html")
    if (state["status"] == "done" and not only_page and os.path.exists(index_path)
            and pages_csv_current(meta_dir)):
        print(f"Section: {sec_name} - already complete, skipped")
        return {"section": section, "root": root, "index": index_path, "entries": [],
                "totals": state["totals"], "aborted": False, "complete": True, "skipped": True}
    ERRORS.reprocessed.add("section:" + sid)
    print(f"Section: {sec_name}")

    # Pages list. "pagelevel=true" is REQUIRED to get indentation level and order; without
    # it Graph omits both. Try several request forms until levels show up.
    fields = ("id,title,createdDateTime,lastModifiedDateTime,level,order,createdByAppId,"
              "contentUrl,self,links,userTags")
    variants = [
        f"{base}/pages?$select={fields}&pagelevel=true",
        f"{base}/pages?pagelevel=true",
        f"{BETA}/me/onenote/sections/{sid}/pages?pagelevel=true",
    ]
    pages, last_err = None, None
    for i, url in enumerate(variants, 1):
        try:
            result = list(paged(url))
        except GraphError as e:
            last_err = e
            print(f"  (page list request {i} failed: HTTP {e.status})")
            continue
        if pages is None:
            pages = result
        if any("level" in p for p in result):
            pages = result
            break
    if pages is None:
        ERRORS.add("section", "page list", f"all request forms failed: {last_err}")
        raise ExportError(f"Could not list pages (all request forms failed): {last_err}")
    if only_page:
        pages = [p for p in pages if p.get("id") == only_page]
        if not pages:
            ERRORS.add("page", "page lookup", f"page {only_page} was not found in its section")
            raise ExportError(f"Page {only_page} was not found in section {sec_name}")
    elif pages and not any("level" in p for p in pages):
        print("  ! Graph returned no level/order values; pages will be numbered flat "
              "(no sub-page nesting).")
    pages.sort(key=lambda p: p.get("order", 0))

    entries = build_entries(pages, pages_dir, flat=bool(only_page))
    write_json(os.path.join(meta_dir, safe_name(sec_name, "section") + ".json"),
               section_json(section, entries))
    print(f"Found {len(entries)} pages -> {os.path.abspath(root)}")

    # Resume check: compare the page count with what is already on disk
    def on_disk(e):
        return (os.path.exists(os.path.join(os.path.join(html_dir, e["html_name"])))
                and os.path.exists(os.path.join(e["dir"], e["md_name"])))

    if state.get("page_count") not in (None, len(entries)):
        print(f"  Page count changed since the last run ({state['page_count']} -> "
              f"{len(entries)}); every page is checked.")
    have = sum(1 for e in entries if on_disk(e))
    if have:
        print(f"  {have} of {len(entries)} pages are already on disk"
              + (" - section looks complete." if have == len(entries) else " - resuming."))

    retry = ERRORS.retry_pages()
    stop, failed, incomplete = None, 0, 0
    try:
        for n, e in enumerate(entries, 1):
            page, pid = e["page"], e["page"]["id"]
            CTX["page"] = e
            prev_e = entries[n - 2] if n > 1 else None
            next_e = entries[n] if n < len(entries) else None
            rel_folder = "/".join(["pages"] + e["rel_parts"])
            tmp_dir = e["dir"] + ".tmp"
            label = f"  [{n}/{len(entries)}] {e['seq']}  {'  ' * e['level']}{e['title']}"

            rec = state["done"].get(pid)
            if (on_disk(e) and pid not in retry
                    and (rec is None or rec.get("folder") == rel_folder)):
                if os.path.isdir(tmp_dir):
                    shutil.rmtree(tmp_dir)
                e["stats"] = (rec or {}).get("stats") or stats_from_disk(e)
                e["status"] = "ok"
                state["done"][pid] = {"folder": rel_folder, "stats": e["stats"]}
                print(label + "  (already exported)")
                continue

            print(label)
            ERRORS.reprocessed.add(pid)
            state["in_progress"] = pid
            before = len(ERRORS.rows)
            try:
                if os.path.isdir(tmp_dir):          # left behind by an interrupted run
                    shutil.rmtree(tmp_dir)
                built = build_page(e, tmp_dir)
                commit_page_dir(tmp_dir, e["dir"])
                write_page_html(e, html_dir, sec_name, prev_e, next_e, built)
                if OPTS["render"] == "all" or (OPTS["render"] == "auto" and built["needs_render"]):
                    if render_page(os.path.join(html_dir, e["html_name"]), e["dir"],
                                   e["md_name"][:-3]):
                        built["stats"]["size"] = page_size(e["dir"])
                e["stats"], e["status"] = built["stats"], "ok"
                if any(r[1] in RETRY_CATEGORIES for r in ERRORS.rows[before:]):
                    e["incomplete"] = True          # saved, but something is missing: retried next run
                    incomplete += 1
                else:
                    state["done"][pid] = {"folder": rel_folder, "stats": e["stats"]}
                state["in_progress"] = None
            except StopRun:
                raise                                # keep in_progress so the next run redoes it
            except GraphError as err:
                ERRORS.add("page", "page content", err, err.url[:300])
                e["status"], state["in_progress"] = "failed", None
                failed += 1
                shutil.rmtree(tmp_dir, ignore_errors=True)
            except Exception as err:                 # never let one page stop the export
                ERRORS.add("page", "page export", f"{type(err).__name__}: {err}")
                print(traceback.format_exc())
                e["status"], state["in_progress"] = "failed", None
                failed += 1
                shutil.rmtree(tmp_dir, ignore_errors=True)
            if n % 20 == 0:
                PROGRESS.flush()
    except StopRun as exc:
        stop = exc
    CTX["page"] = None

    # Always write the section's tables of contents, even after a stop
    write_csv(os.path.join(meta_dir, "pages.csv"), pages_csv_rows(entries))
    write_text(index_path, render_index(section, entries, utc, notebook_link))

    done_stats = [e["stats"] for e in entries if e["stats"]]
    totals = {k: sum(s.get(k, 0) for s in done_stats)
              for k in ("words", "size", "images", "tables", "attachments")}
    totals.update(pages=len(entries), ok=sum(1 for e in entries if e["status"] == "ok"),
                  incomplete=incomplete, failed=failed)
    complete = stop is None and failed == 0 and incomplete == 0
    state.update(page_count=len(entries), totals=totals,
                 status="done" if (complete and not only_page) else "partial")
    PROGRESS.flush()
    result = {"section": section, "root": root, "index": index_path, "entries": entries,
              "totals": totals, "aborted": stop is not None, "complete": complete,
              "skipped": False}
    if stop:
        stop.result = result
        raise stop
    return result


# --------------------------------------------------------------------------
# Notebook structure and files
# --------------------------------------------------------------------------
SECTIONS_CSV_HEADER = [
    "Number", "Sequence", "Name", "Created", "Modified", "Size (bytes)", "Page count",
    "WordCount", "Images", "Tables", "Attachments", "Type", "ID", "Parent section group", "Path",
]


# --------------------------------------------------------------------------
# Reading the notebook structure
# --------------------------------------------------------------------------
def new_node(kind, meta, parent):
    return {"kind": kind, "meta": meta, "parent": parent, "children": [],
            "status": "pending", "result": None, "seq": "", "folder": ""}


def fetch_nodes(base_url, parent):
    """Sections first, then section groups (recursively)."""
    nodes = [new_node("section", s, parent) for s in paged(base_url + "/sections")]
    for g in paged(base_url + "/sectionGroups"):
        node = new_node("group", g, parent)
        node["children"] = fetch_nodes(f"{GRAPH}/me/onenote/sectionGroups/{g['id']}", node)
        nodes.append(node)
    return nodes


def number(nodes, prefix=""):
    for i, n in enumerate(nodes, 1):
        n["seq"] = f"{prefix}{i}"
        number(n["children"], n["seq"] + ".")


def assign_paths(nodes, parent_parts):
    """Folder for every section/group: <sequence>_<name>, nested like the hierarchy."""
    for n in nodes:
        n["folder"] = f"{n['seq']}_{safe_name(name_of(n), 'Section')}"
        n["parts"] = parent_parts + [n["folder"]]          # path below the notebook root
        assign_paths(n["children"], n["parts"])


def flatten(nodes):
    for n in nodes:
        yield n
        yield from flatten(n["children"])


def name_of(n):
    return n["meta"].get("displayName") or "Untitled"


def node_json(n):
    d = dict(n["meta"])                       # as returned by the API
    if n["kind"] == "group":
        d["sections"] = [node_json(c) for c in n["children"]]
    return d


# --------------------------------------------------------------------------
# Totals per section / group
# --------------------------------------------------------------------------
TOTAL_KEYS = ("pages", "ok", "size", "words", "images", "tables", "attachments")


def totals_of(n):
    """Totals for a section, or the sum over everything inside a section group."""
    if n["kind"] == "section":
        return n["result"]["totals"] if n["result"] else None
    sums, found = dict.fromkeys(TOTAL_KEYS, 0), False
    for c in n["children"]:
        t = totals_of(c)
        if t:
            found = True
            for k in TOTAL_KEYS:
                sums[k] += t[k]
    return sums if found else None


def status_of(n):
    if n["kind"] == "section":
        return "not exported" if n["status"] == "pending" else n["status"]
    states = {status_of(c) for c in n["children"]}
    if not states or states == {"ok"}:
        return "ok"
    return "not exported" if states <= {"failed", "not exported"} else "partial"


# --------------------------------------------------------------------------
# Output files
# --------------------------------------------------------------------------
def sections_csv_rows(flat):
    rows = [SECTIONS_CSV_HEADER]
    for i, n in enumerate(flat, 1):
        t = totals_of(n) or {}
        m = n["meta"]
        has_folder = n["kind"] == "group" or bool(n["result"])
        rows.append([
            i, n["seq"], name_of(n), m.get("createdDateTime", ""), m.get("lastModifiedDateTime", ""),
            t.get("size", ""), t.get("pages", ""), t.get("words", ""), t.get("images", ""),
            t.get("tables", ""), t.get("attachments", ""),
            "section" if n["kind"] == "section" else "section group", m.get("id", ""),
            name_of(n["parent"]) if n["parent"] else "",
            "/".join(n["parts"]) if has_folder else "",
        ])
    return rows


def render_notebook_index(nb, nodes, utc):
    def pages_label(count):
        return f"{count} page" + ("" if count == 1 else "s")

    def render(items, top):
        out = ['<ol class="toc">' if top else "<ol>"]
        for n in items:
            label = html.escape(name_of(n))
            if n["kind"] == "group":
                link = f"<strong>{label}</strong>"
            elif n["status"] in ("ok", "partial"):
                href = "/".join(quote(p) for p in n["parts"]) + "/html/index.html"
                count = n["result"]["totals"]["pages"]
                link = f'<a href="{href}">{label}</a><span class="count">{pages_label(count)}</span>'
            else:
                link = f"{label} <em>(not exported)</em>"
            out.append(f'<li><span class="seq">{n["seq"]}</span>{link}')
            if n["children"]:
                out.append(render(n["children"], False))
            out.append("</li>")
        out.append("</ol>")
        return "\n".join(out)

    flat = list(flatten(nodes))
    sections = [n for n in flat if n["kind"] == "section"]
    pages = sum(n["result"]["totals"]["pages"] for n in sections if n["result"])
    name = nb.get("displayName") or "Notebook"
    meta = [f"{len(sections)} sections", pages_label(pages),
            "Exported " + html.escape(fmt_dt(datetime.now(timezone.utc).isoformat(), utc))]
    return INDEX_TEMPLATE.substitute(
        title_text=html.escape(name), css=INDEX_CSS, title=html.escape(name),
        topbar="OneNote export", meta=" &middot; ".join(meta),
        toc=render(nodes, True) if nodes else "<p>No sections.</p>")




# --------------------------------------------------------------------------
# Running a notebook / section / page
# --------------------------------------------------------------------------
def run_notebook(nid, outdir):
    utc = OPTS["utc"]
    CTX["section"] = None
    try:
        nb = get(f"{GRAPH}/me/onenote/notebooks/{nid}").json()
        nodes = fetch_nodes(f"{GRAPH}/me/onenote/notebooks/{nid}", None)
    except GraphError as e:
        ERRORS.add("section", "notebook", e, e.url[:300])
        raise ExportError(f"Could not read notebook:\n{e}")

    nb_name = nb.get("displayName") or nid
    root = os.path.join(outdir, safe_name(nb_name, "Notebook"))
    meta_dir = os.path.join(root, "metadata")
    os.makedirs(meta_dir, exist_ok=True)
    os.makedirs(os.path.join(root, "sections"), exist_ok=True)
    attach_run_files(root)

    number(nodes)
    flat = list(flatten(nodes))
    sections = [n for n in flat if n["kind"] == "section"]
    assign_paths(nodes, ["sections"])
    for n in flat:
        if n["kind"] == "group":
            os.makedirs(os.path.join(root, *n["parts"]), exist_ok=True)

    doc = dict(nb)
    doc["sections"] = [node_json(n) for n in nodes]
    write_json(os.path.join(meta_dir, safe_name(nb_name, "notebook") + ".json"), doc)
    print(f"Notebook: {nb_name} ({len(sections)} sections) -> {os.path.abspath(root)}")

    stop, current = None, None
    try:
        for i, n in enumerate(sections, 1):
            current = n
            print(f"\n=== Section {i}/{len(sections)} ({n['seq']}): {name_of(n)} ===")
            # from sections/.../<folder>/html/ back up to the notebook's index.html
            backlink = "../" * (len(n["parts"]) + 1) + "index.html"
            try:
                res = export_section(n["meta"]["id"], os.path.join(root, *n["parts"][:-1]),
                                     folder_name=n["folder"], notebook_link=(nb_name, backlink),
                                     section_hint=n["meta"])
                n["result"] = res
                n["status"] = "ok" if res["complete"] else "partial"
            except ExportError as e:
                n["status"] = "failed"
                print(f"  ! section failed: {e}")
    except StopRun as exc:
        stop = exc
        if current is not None and exc.result:
            current["result"], current["status"] = exc.result, "partial"

    write_csv(os.path.join(meta_dir, "sections.csv"), sections_csv_rows(flat))
    write_text(os.path.join(root, "index.html"), render_notebook_index(nb, nodes, utc))

    ok = sum(1 for n in sections if n["status"] == "ok")
    print(f"\n{ok} of {len(sections)} sections fully exported.")
    for n in sections:
        if n["status"] != "ok":
            print(f"  {status_of(n)}: {n['seq']} {name_of(n)}")
    print(f"Open: {os.path.abspath(os.path.join(root, 'index.html'))}")
    if stop:
        raise stop
    return {"complete": ok == len(sections)}


def report_section(res):
    t = res["totals"] or {}
    print(f"\nDone: {t.get('ok', 0)} of {t.get('pages', 0)} pages exported.")
    for e in res["entries"]:
        if e["status"] != "ok" or e.get("incomplete"):
            print(f"  needs attention: {e['seq']} {e['title']}")
    print(f"Open: {os.path.abspath(res['index'])}")


def run_section(sid, outdir):
    res = export_section(sid, outdir)
    report_section(res)
    return {"complete": res["complete"]}


def run_page(pid, outdir):
    CTX["section"] = None
    try:
        page = get(f"{GRAPH}/me/onenote/pages/{pid}").json()
    except GraphError as e:
        ERRORS.add("page", "page details", e, e.url[:300])
        raise ExportError(f"Could not read page:\n{e}")
    section_id = (page.get("parentSection") or {}).get("id")
    if not section_id:
        ERRORS.add("page", "page details", "Graph did not say which section the page is in")
        raise ExportError("Could not find the section of this page.")
    res = export_section(section_id, outdir, folder_name=safe_name(page.get("title"), "Page"),
                         only_page=pid)
    report_section(res)
    return {"complete": res["complete"]}


def finish(outdir, complete):
    """Write errors.csv, then delete progress.json if everything is done."""
    if not ERRORS.path:       # failed before a root folder existed: keep the notes in outdir
        ERRORS.attach(outdir)
        if isinstance(sys.stdout, Tee):
            sys.stdout.attach(os.path.join(outdir, "metadata", "logs.txt"))
    ERRORS.flush()
    if complete and ERRORS.retry_count() == 0:
        PROGRESS.clear()
    else:
        PROGRESS.flush()
    n = ERRORS.count()
    if n:
        print(f"\n{n} item(s) recorded in {ERRORS.path} for review.")


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Export OneNote content to local files (safe to stop and re-run).")
    ap.add_argument("kind", type=str.upper, choices=["N", "S", "P"], metavar="N|S|P",
                    help="what the ID is: N = notebook, S = section, P = page")
    ap.add_argument("target_id", metavar="ID", help="notebook / section / page ID (use single quotes)")
    ap.add_argument("-o", "--outdir", default=".", help="parent folder (default: current)")
    ap.add_argument("--utc", action="store_true", help="show dates in UTC instead of local time")
    ap.add_argument("--no-ink", action="store_true", help="do not request handwriting (InkML)")
    ap.add_argument("--ocr", action="store_true", help="OCR downloaded images with tesseract")
    ap.add_argument("--render", choices=["auto", "all", "none"], default="auto",
                    help="save pages as PDF/PNG: auto = pages with handwriting, OCR text or "
                         "only images (default), all, none")
    ap.add_argument("--render-as", choices=["pdf", "png", "both"], default="pdf",
                    help="format of the saved page (default pdf)")
    args = ap.parse_args(argv)

    raw = os.environ.get("ONENOTE_TOKEN")
    if not raw:
        print("Set ONENOTE_TOKEN first")
        return 1
    set_token(raw)
    OPTS.update(utc=args.utc, ink=not args.no_ink, ocr=args.ocr, render=args.render,
                render_as=args.render_as)

    real_stdout = sys.stdout
    sys.stdout = Tee(real_stdout)
    init_run(args.kind, args.target_id)
    code, complete = 0, False
    try:
        runner = {"N": run_notebook, "S": run_section, "P": run_page}[args.kind]
        complete = runner(args.target_id, args.outdir)["complete"]
        code = 0 if complete else 3
    except RateLimitExceeded as e:
        print(f"\nStopped: {e}.\nYour progress is saved. Please run the same command again "
              "later to continue where it stopped.")
        code = 2
    except TokenExpired as e:
        print(f"\nStopped: {e}.\nYour progress is saved. Copy a fresh token from Graph Explorer, "
              "set ONENOTE_TOKEN again and run the same command to continue.")
        code = 2
    except ExportError as e:
        print(f"\n{e}")
        code = 1
    except KeyboardInterrupt:
        print("\nInterrupted. Your progress is saved; run the same command again to continue.")
        code = 130
    except Exception as e:     # unexpected: keep the evidence, stay resumable
        print(traceback.format_exc())
        ERRORS.add("section", "unexpected error", f"{type(e).__name__}: {e}")
        code = 1
    finally:
        finish(args.outdir, complete)
        sys.stdout.close()
        sys.stdout = real_stdout
    return code


if __name__ == "__main__":
    sys.exit(main())

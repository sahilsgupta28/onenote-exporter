# OneNote Graph Exporter

A single-file Python command-line tool that exports OneNote notebooks, sections or single pages
from a **personal Microsoft account** to open, offline formats: Markdown, browsable HTML, CSV
tables, original images and attachments, and (optionally) PDF/PNG copies of pages that contain
handwriting. It talks to the **Microsoft Graph OneNote API** over HTTPS, so it needs no Windows,
no OneNote desktop app and no Office installation.

```
python3 export_onenote.py N '<notebook id>'     # a whole notebook
python3 export_onenote.py S '<section id>'      # one section
python3 export_onenote.py P '<page id>'         # one page
```

It is built for large notebooks and flaky connections: it retries when Microsoft throttles
requests, **stops cleanly** if throttling persists, and **resumes where it stopped** the next
time you run the same command. Nothing is overwritten until its replacement is complete, and
anything that could not be downloaded or understood is written to `errors.csv` for review.

> **Status.** The script was developed against documented Graph behaviour and tested with
> mocked API responses (rate limits, failed downloads, expired tokens, nested sections and
> pages, ink, OCR, rendering). Anything described below as approximate or assumed (for example the
> placement of handwriting) could not be confirmed against every kind of live page. Always keep your original notebooks and review a sample of the export.

## Contents

- [Highlights](#highlights)
- [How it compares](#how-it-compares)
- [Requirements](#requirements)
- [Quick start](#quick-start)
- [Getting a token and IDs](#getting-a-token-and-ids)
- [Manual page](#manual-page)
- [Output layout](#output-layout)
- [Numbering rules](#numbering-rules)
- [File formats](#file-formats)
- [Content support](#content-support)
- [Reliability: retries, resume and errors](#reliability-retries-resume-and-errors)
- [Architecture](#architecture)
- [Known limitations](#known-limitations)
- [Troubleshooting](#troubleshooting)
- [Testing](#testing)
- [AI assistance](#ai-assistance)
- [License and disclaimer](#license-and-disclaimer)

## Highlights

- Export a notebook, a section or a single page, selected with one command-line letter (`N`, `S`, `P`).
- Keeps the **hierarchy**: section groups become folders, sub-pages nest inside their parent page.
- Every section and page gets a **sequence number** (`1`, `1.2`, `3.4.1`) used in folder names,
  file names, tables of contents and CSV files.
- Three views of the same content: **Markdown** (per page), **HTML** (OneNote-style pages with an
  index, breadcrumbs and previous/next links) and **CSV/JSON metadata**.
  Each page shows its Created date, and its Modified date only when that differs from Created.
- Downloads **images** (full resolution first, detected file type), **attachments**, inline
  data-URI images, and puts OneNote's split images back in order.
- Every table is also saved as a **CSV** file.
- Handwriting (InkML to SVG) shown in the pages; optional **OCR** of images (local `tesseract`);
  optional **PDF/PNG copies** of pages (`--render`, off by default; local Chromium-based browser).
- **Rate-limit aware**: up to 10 retries on HTTP 429 with exponential back-off (2, 4, 8, 16, then
  20 seconds), then a clean stop with a "run again later" message.
- **Resumable**: progress is tracked in memory, written to `metadata/progress.json` when something
  goes wrong, and removed when everything finished.
- **Loss-averse**: atomic file writes, page replacement only after a complete download, a failed
  image never discards the rest of the page, and `errors.csv` records everything unusual.
- Console output is copied to `metadata/logs.txt`.

## How it compares

Tools such as [OneNote Backup Exporter](https://github.com/jli-software/OneNoteBackupExporter)
(exports `.onepkg`, HTML, PDF and XPS) and
[OneNote Md Exporter](https://github.com/alxnbl/onenote-md-exporter) (Markdown via Word and
Pandoc) drive the **OneNote desktop application through its Windows COM API**. This tool takes
the other route:

| | COM-based desktop exporters | This tool |
|---|---|---|
| Needs | Windows, OneNote desktop (and Word for some) | Python 3 and an internet connection |
| Data source | Local, synced notebooks | Microsoft Graph (the cloud copy) |
| Account types | Whatever OneNote desktop has open | Personal Microsoft accounts (the target of this tool) |
| Works on macOS / Linux | No | Yes |
| Page order and indentation | Via COM XML | Via Graph `order` and `level` |
| Resumes after throttling | Not applicable | Yes |

Choose a COM-based exporter if you want a re-importable `.onepkg` backup; this tool does not
produce one.

## Requirements

- **Python 3.8 or newer** and the `requests` package: `pip3 install requests`
- A Microsoft Graph **access token** with the `Notes.Read` permission (see below)
- Optional:
  - [`tesseract`](https://github.com/tesseract-ocr/tesseract) on your `PATH` for `--ocr`
  - Google Chrome, Chromium, Microsoft Edge or Brave, only if you use `--render` to save pages as PDF/PNG

The script runs wherever Python runs. It uses no other third-party Python packages.

## Quick start

```bash
pip3 install requests

# 1. Get a token (see "Getting a token and IDs"), then set it. Single quotes matter.
read -s ONENOTE_TOKEN && export ONENOTE_TOKEN     # paste the token, press Enter

# 2. Export a notebook to ~/OneNoteExport
python3 export_onenote.py N '0-1234567890ABCDEF!101' -o ~/OneNoteExport

# 3. Open the result
open ~/OneNoteExport/<Notebook name>/index.html
```

If it stops because of rate limiting or an expired token, set a fresh token if needed and run
**the same command again**; it continues where it stopped.

## Getting a token and IDs

The tool uses the token you copy from Graph Explorer, so no app registration (and no Azure
tenant) is needed.

1. Open <https://developer.microsoft.com/graph/graph-explorer> and sign in with your personal
   Microsoft account.
2. Run `GET https://graph.microsoft.com/v1.0/me/onenote/notebooks` and accept the `Notes.Read`
   permission when asked (Modify permissions tab).
3. Open the **Access token** tab and copy the token. It lasts about an hour. Treat it like a
   password.
4. Export it: `export ONENOTE_TOKEN='<token>'` (or `read -s ONENOTE_TOKEN && export ONENOTE_TOKEN`
   to keep it out of your shell history).

Find the IDs in Graph Explorer:

| You want | Query | Take the `id` of the item |
|---|---|---|
| Notebook ID | `GET /v1.0/me/onenote/notebooks` | the notebook |
| Section ID | `GET /v1.0/me/onenote/sections` (all) or `.../notebooks/{id}/sections` | the section |
| Page ID | `GET /v1.0/me/onenote/sections/{id}/pages` | the page |

IDs contain `!`, which zsh/bash treat specially, so **always wrap IDs in single quotes**.

## Manual page

```
NAME
    export_onenote.py - export OneNote notebooks, sections or pages through Microsoft Graph

SYNOPSIS
    python3 export_onenote.py N|S|P ID [-o DIR] [--utc] [--no-ink] [--ink-layout canvas|end] [--html-images embed|link] [--ocr]
                                       [--render auto|all|none] [--render-as pdf|png|both]
    python3 export_onenote.py -h

DESCRIPTION
    Downloads the content named by ID and writes Markdown, HTML, CSV, JSON, images,
    attachments and (optionally) ink, OCR text and PDF/PNG copies below DIR.
    The first argument selects what ID refers to:

        N   a notebook   (all sections and section groups inside it)
        S   a section    (all pages, with sub-pages nested)
        P   a page       (only that page, numbered 1)

    The console shows one line per page; when a page brings extra files, a second line
    says what was exported, for example "Exported 2 images, 1 table, 2 attachments,
    94 handwriting strokes" (nothing is printed for pages without extra files).

    The command is safe to stop and run again: see RESUMING.

OPTIONS
    -o, --outdir DIR
            Parent folder for the export. Default: the current folder.
    --utc
            Show dates in the Markdown and HTML pages in UTC. Default: local time with
            its time-zone abbreviation. CSV and JSON files always keep the ISO UTC
            timestamps returned by Graph.
    --no-ink
            Do not ask Graph for handwriting (InkML). Saves one probe per run when you
            know your pages have no handwriting.
    --html-images embed|link
            How the HTML pages show images.
              embed  (default) the images are placed inside each page file, so they show
                     wherever the page is opened. Browsers such as Safari often refuse to
                     load pictures from outside the folder of a local HTML file.
              link   the page refers to the files in pages/ (smaller HTML files, but the
                     images may not show in every browser). Pictures over 20 MB stay linked.
            The image files are saved in pages/.../images/ either way.
    --ink-layout canvas|end
            How the HTML pages show handwriting.
              canvas  (default) the page keeps OneNote's layout: text boxes stay where
                      OneNote put them and the handwriting is drawn over them as an inline
                      SVG, so the page looks like it does in OneNote.
              end     the page text is shown in reading order and the handwriting is
                      shown below it.
            The standalone SVG (ink/ink.svg, openable in any browser) is saved either way and
            is linked from the page's Markdown file.
    --ocr
            Run the local `tesseract` program on every downloaded image. The text is saved
            next to the image (image-01.txt), as a quote in the Markdown and in a
            collapsible block in the HTML. Graph itself does not return OCR text.
    --render auto|all|none
            Save the finished page as PDF/PNG with a headless Chromium-based browser.
              none  (default) no PDF/PNG export
              auto  pages with handwriting, OCR text, or images but no text
              all   every page
    --render-as pdf|png|both
            Format of the saved page. Default: pdf. PNG is one screenshot of a fixed-size
            window, so very long pages can be cut off; use pdf for those.
    -h, --help
            Show the built-in help.

ENVIRONMENT
    ONENOTE_TOKEN     Required. Graph access token (Notes.Read). Whitespace, quotes and a
                      leading "Bearer" are ignored.
    ONENOTE_BROWSER   Optional. Path or name of the Chromium-based browser to use for
                      --render. Default: search PATH and the usual macOS locations.

EXIT STATUS
    0    Everything was exported.
    1    Fatal error (for example the notebook or section could not be read).
    2    Stopped early: HTTP 429 persisted after 10 retries, or the token expired.
         Progress is saved; run the same command later.
    3    Finished, but some pages need a retry (see errors.csv); run the same command again.
    130  Interrupted with Ctrl-C. Progress is saved.

FILES
    <root>/errors.csv              problems and review items (only when there are any)
    <root>/metadata/logs.txt       console output of every run (appended)
    <root>/metadata/progress.json  resume state (exists only while a run is unfinished)

RESUMING
    Run the identical command again. Finished sections are skipped without any API call,
    finished pages are skipped, and the page that was in progress is downloaded again.
    Pages listed in errors.csv with a download problem are retried.

EXAMPLES
    python3 export_onenote.py N '0-1234567890ABCDEF!101' -o ~/OneNoteExport
    python3 export_onenote.py S '0-1234567890ABCDEF!202' --ocr
    python3 export_onenote.py P '0-0123456789abcdef0123456789abcdef!1-1234567890ABCDEF!202'
    python3 export_onenote.py N '<id>' --render all --render-as both --utc
```

## Output layout

Notebook mode (`N`):

```
<Notebook name>/
├── index.html                       numbered list of sections -> each section's index
├── errors.csv                       only if something failed or needs review
├── metadata/
│   ├── sections.csv                 one row per section / section group
│   ├── <Notebook name>.json         notebook + sections (in order, groups nested)
│   ├── logs.txt                     console output, all runs
│   └── progress.json                only while a run is unfinished
└── sections/
    ├── 1_Mathematics/               a section (layout below)
    └── 2_Science/                   a section group = a folder
        ├── 2.1_Physics/
        └── 2.2_Chemistry/
```

Section mode (`S`) writes the section folder directly under `-o`; page mode (`P`) writes a
folder named after the page, using the same layout and numbering the page as `1`.

A section folder:

```
<Section>/
├── metadata/
│   ├── pages.csv                    one row per page
│   └── <Section name>.json          section + page metadata (sub-pages nested)
├── html/
│   ├── index.html                   numbered, nested list of pages (links to the files below)
│   ├── 1_Introduction.html          one OneNote-style page per page
│   └── 3.1_Exercises.html
└── pages/
    ├── 1_Introduction/
    │   └── 1_Introduction.md
    └── 3_Chapter Three/
        ├── 3_Chapter Three.md
        └── pages/                   sub-pages, same structure
            └── 3.1_Exercises/
                ├── 3.1_Exercises.md
                ├── 3.1_Exercises.pdf      (only with --render)
                ├── images/image-01.png            (only if the page has images)
                ├── images/image-01.txt            (only with --ocr)
                ├── tables/table-01.csv            (only if the page has tables)
                ├── attachments/report.pdf         (only if the page has attachments)
                └── ink/ink.inkml, ink/ink.svg     (only if the page really has handwriting)
```

Folder and file names replace the characters `\ / : * ? " < > |` with `_` and are limited to
100 characters; the sequence prefix keeps names unique.

## Numbering rules

Items are numbered by their position and indentation level:

- Top-level pages (or sections) are `1`, `2`, `3` ...
- Sub-pages extend their parent's number: `3.1`, `3.2`; their sub-pages `3.2.1` and so on.
- **Pages** follow Graph's `order` value and are nested using Graph's `level` (0 = top level,
  1 = sub-page, ...). A page belongs to the nearest earlier page with a lower level.
- **Sections** are numbered in the order Graph returns them (Graph has no section order field).
  Sections come first, then section groups, because the API lists them separately; items inside
  a group are numbered under it, for example `3`, `3.1`, `3.4.1`.

## File formats

### `metadata/pages.csv`

| Column | Meaning |
|---|---|
| Number | Row counter (1, 2, 3 ...) |
| Sequence | Outline number such as `3.1` |
| Page name | Page title |
| Created / Modified | ISO UTC timestamps from Graph |
| WordCount | Words in the page's text (OCR text and ink not counted) |
| Size (bytes) | Bytes of the page's own files (Markdown, images, tables, attachments, ink, PDF/PNG), not its sub-pages or its HTML file |
| Level | Graph indentation level (0 = top) |
| Order | Raw `order` value from Graph (starts at 0) |
| Images / Tables / Attachments | Counts |
| Page ID | The Graph ID of the page (usable with `P`) |
| Path | Page folder relative to the section folder |

### `metadata/sections.csv` (notebook mode)

Number, Sequence, Name, Created, Modified, Size (bytes), Page count, WordCount, Images, Tables,
Attachments, **Type** (`section` or `section group`), **ID**, **Parent section group**, **Path**
(relative to the notebook folder; empty for sections that were not exported). Section group rows
carry the totals of everything inside them.

### `errors.csv`

Columns: Time, Category, Section, Section ID, Page, Page ID, Sequence, Item, URL, Error.

| Category | Meaning | Retried on the next run? |
|---|---|---|
| `download` | An image or attachment could not be downloaded | Yes (the page is downloaded again) |
| `page` | A page's content could not be fetched or processed | Yes |
| `section` | A section / notebook could not be read | Yes |
| `format` | A data format the script cannot fully handle: embedded video (`iframe`), `embed`, `audio`, `video`, `canvas`, inline `svg`, objects without a file name, unrecognised image data, unreadable InkML, unsupported API options, missing browser | No; kept for your review |
| `ocr` | OCR could not run (for example `tesseract` is not installed) | No |

The file is rewritten on every run: rows for pages that were downloaded again are replaced, rows
for untouched pages are kept, and the file is deleted when nothing is left.

### JSON metadata

`<Section name>.json` is the section object returned by Graph plus a `pages` array in notebook
order; each page is its API metadata (never its content) with its own `pages` array for
sub-pages. `<Notebook name>.json` is the notebook plus a `sections` array in which section
groups carry their own `sections` array.

### `progress.json`

Internal resume state: for each section its status (`done`/`partial`), the page count, the page
IDs already finished (with their folder and statistics) and the page that was in progress. You
normally never need to open it.

## Content support

| Content | Markdown | HTML | Notes |
|---|---|---|---|
| Text, headings, bold, italic, strikethrough | Yes | Yes | |
| Links | Yes | Yes | `onenote:` links are kept as they are |
| Lists, nested lists | Yes | Yes | |
| To-do checkboxes | Yes (`[ ]` / `[x]`) | Yes | |
| Other note tags (Important, Question ...) | Yes, as a bold label | Attribute kept | |
| Tables | Yes, plus a CSV per table | Yes | First row becomes the Markdown header; nested tables are flattened into their cell |
| Images | Yes (linked file) | Yes (embedded in the page by default) | Full resolution first; type detected from the file bytes (default `.png`); split images reordered |
| Inline (data-URI) images | Yes | Yes | |
| File attachments | Yes (link) | Yes (link) | Includes vCards from OneNote data extraction |
| Videos / embedded web content | Link only | Embedded frame | Logged for review |
| Handwriting (ink) | SVG image with an "open in a browser" link | Inline SVG drawn over the text boxes at OneNote's positions (or below the text with `--ink-layout end`) | Graph returns it for `includeInkML` requests (pages without handwriting get an empty part, which is ignored without any log entry). Placement assumes OneNote's ink units (himetric, 0.01 mm) share the page origin with the text boxes; check a sample page |
| OCR text of images | With `--ocr` | With `--ocr` | Local `tesseract`; weak on real handwriting |
| Font colours, highlighting, fonts | No | Yes (inline styles kept) | |
| Page layout (text boxes at fixed positions) | n/a | Stacked in reading order | OneNote's absolute positions are not reproduced |
| Whole page as picture | PDF/PNG file | | Off by default; `--render` turns it on; needs a Chromium-based browser |

OneNote's own data-extraction output (business cards, recipes, products) arrives from Graph as
ordinary HTML plus a vCard attachment, so it needs no special handling.

## Reliability: retries, resume and errors

### Retry policy

| Situation | Behaviour |
|---|---|
| HTTP **429** | Up to **10** retries. Wait before retry *n* = `min(20, 2^n)` seconds (2, 4, 8, 16, 20, 20, ...); a larger `Retry-After` is honoured up to the 20 s cap. After the 10th retry the run **stops** and tells you to continue later (exit status 2). |
| HTTP **5xx**, network errors | 3 retries (waits 2, 4, 8 s), then the item fails and is logged. |
| HTTP **401** from Graph | The run stops: the token expired. Paste a fresh token and run again. |
| Other HTTP errors on an image / attachment | The alternate URL forms are tried; if all fail it is logged as `download` and the page is retried next run. |

The token is only ever sent to `graph.microsoft.com`, `www.onenote.com` and `onenote.com`.

### Resume algorithm

1. **In memory:** after each page the script records it as finished (folder and statistics).
2. **On any problem** (an error row, a stop, Ctrl-C, an exception) and every 20 pages and after
   every section, the index is written to `metadata/progress.json`.
3. **Next run, per section:**
   - *Finished* sections (per `progress.json`) are skipped with no API call.
   - Otherwise the page list is fetched (cheap) and the **page count is compared with what is on
     disk**: "7 of 12 pages are already on disk - resuming".
   - A page counts as finished only if its Markdown and its HTML file exist (the HTML file is
     written last) and it is not waiting for a retry in `errors.csv`.
   - The page that was in progress is thrown away and downloaded again.
4. **When everything finished with no retryable errors,** `progress.json` is deleted.

If `progress.json` is missing (for example after a power cut), the same disk check is used, so
re-running is always safe and cheap.

Exports made by older versions are upgraded in place: pages that have images but whose HTML
does not carry them are downloaded again, pages that have handwriting but whose
HTML does not draw it are downloaded again, `ink/` folders that hold only empty InkML are
removed, and earlier "no drawable strokes" / "handwriting saved" rows in `errors.csv` are
dropped.

A finished section whose `pages.csv` is out of date (for example an export made before the
`Page ID` column existed) is not skipped: its `pages.csv` and indexes are rewritten with the
current columns, and no page is downloaded again.

### Protection against data loss

- All files are written to a `.part` file and renamed, so a half-written file never looks
  finished.
- A page is downloaded into `<page folder>.tmp` and swapped in only after it is complete; its
  sub-pages are never touched. Leftover `.tmp` folders are removed on the next run.
- A page whose image failed to download is still written (with the original link kept), so you
  keep the text while the page is queued for a retry.
- Unexpected exceptions are logged to `errors.csv` and the run continues with the next page.
- `index.html`, `pages.csv` and `sections.csv` are written even when a run stops early.

## Architecture

Everything lives in one file, `export_onenote.py` (about 2,000 lines), organised in layers.
Nothing outside Python's standard library is used except `requests`.

```
                         main()  -- argument parsing, run state, exit codes
                           |
          +----------------+-----------------+
          |                |                 |
     run_notebook()   run_section()      run_page()
          |                |                 |
          |                +-----> export_section() <----+
          |                              |               (P: only_page)
          +-- for each section --------->|
                                         v
            list pages (level/order) -> build_entries() -> sequence numbers, folders
                                         |
                       for each page:  resume check (progress + disk + errors.csv)
                                         |
                                  build_page() --> fetch_page_content()  [Graph]
                                         |            localize_resources()  [images, files]
                                         |            TextExtractor / MarkdownConverter
                                         |            save_ink()  ocr_image()
                                         v
                                  commit_page_dir()  (swap .tmp into place)
                                         |
                                  write_page_html()  (completion marker)
                                         |
                                  render_page()  (optional PDF/PNG)
                                         v
                       pages.csv, <Section>.json, html/index.html
```

| Layer | Main pieces | Responsibility |
|---|---|---|
| Run state | `OPTS`, `CTX`, `init_run`, `attach_run_files` | Options, "current section/page" context for error rows, per-run singletons |
| Logging | `Tee` | Copies console output into `metadata/logs.txt` (buffers until the folder exists) |
| Problems | `ErrorLog` | Collects rows, rewrites `errors.csv` (merging rows of untouched pages), decides what is retried |
| Progress | `Progress` | In-memory resume index, `progress.json` read/write/clear |
| Network | `get`, `backoff`, `paged`, `fetch_resource` | One place for the retry policy, auth, paging and alternate resource URLs |
| Files | `atomic_write`, `write_csv`, `write_json`, `save_stream` | `.part` + rename writes; CSVs with a UTF-8 BOM so Excel reads them correctly |
| Resources | `localize_resources`, `reorder_split_images`, `image_ext` | Download images/attachments, data URIs, split-image order, re-point links, flag unsupported tags |
| Ink / OCR / render | `split_multipart`, `parse_inkml`, `save_ink`, `canvas_size`, `ocr_image`, `render_page` | Optional extras, each failing softly into `errors.csv` |
| Conversion | `TextExtractor`, `MarkdownConverter` | HTML to word count / Markdown; collects tables for CSV export |
| HTML output | `PAGE_TEMPLATE`, `prepare_html_body`, `render_page_html`, `render_index` | OneNote-style pages, breadcrumbs, previous/next, section index |
| Section export | `build_entries`, `export_section`, `build_page`, `commit_page_dir`, `stats_from_disk` | Numbering, resume logic, atomic page replacement |
| Notebook export | `fetch_nodes`, `number`, `assign_paths`, `sections_csv_rows`, `render_notebook_index`, `run_notebook` | Section groups, folders, totals, notebook index and back-links |

### Key design decisions

- **Graph, not COM.** It removes the Windows/OneNote dependency, but Graph returns less than the
  desktop app: no section order, no OCR text, and ink only through the `includeInkML` option.
- **Page levels need `pagelevel=true`.** Graph omits `level` and `order` unless asked; the script
  tries three request forms until levels appear.
- **The HTML file is the completion marker.** It is written last, atomically, so "HTML exists"
  reliably means "page complete" without extra marker files.
- **`.tmp` swap instead of in-place edits.** Downloads can fail at any moment; replacing only
  after success means a failed re-download never destroys a good copy.
- **Sequence numbers instead of titles for uniqueness.** Titles repeat (two sections named
  `Notes`); sequences never do, and they preserve the original order on disk.
- **Soft failures.** A problem with one image, one page or one optional feature is logged and the
  export carries on; only throttling and token expiry stop the whole run.

## Known limitations

- Built for **personal Microsoft accounts**; work or school accounts were not tested.
- Graph provides no section order, so sections follow the API's order and may differ from your
  tab order in OneNote. Sections are listed before section groups.
- **Handwriting** is drawn from the InkML Graph returns. Strokes are approximate (pen pressure
  and exact brush shapes are not reproduced) and their placement relies on the assumption that
  ink coordinates (himetric) and text-box positions (pixels) share one origin; if a page looks
  misaligned, use `--ink-layout end` and report it. Pages without handwriting need no handling.
- OCR uses `tesseract` on your machine, not Microsoft's OCR, and is unreliable for handwriting.
- With `--html-images embed` the HTML files are larger, because they contain their pictures.
- Page layout is flattened to reading order in HTML; Markdown loses colours and fonts.
- Embedded videos and web content are kept as links; unsupported elements are logged, not
  converted.
- `Sequence` values such as `1.10` may be shown as `1.1` if you open `pages.csv` in Excel; import
  the column as text.
- Graph tokens expire after about an hour; long exports may need a fresh token mid-way (resume
  handles this).
- No `.onepkg` export and no Markdown front-matter or wikilink conversion (unlike COM-based
  Markdown exporters).

## Troubleshooting

| Symptom | Cause and fix |
|---|---|
| `Set ONENOTE_TOKEN first` | The variable is not set in this terminal window. Export it again. |
| `HTTP 401` / "token is expired" | Copy a fresh token from Graph Explorer (and check the `Notes.Read` permission), then re-run. |
| `zsh: no such event: 101` | The ID contains `!`; wrap it in single quotes. |
| "HTTP 429 persisted after 10 retries" | Microsoft is throttling you. Wait a while and run the same command again. |
| Pages not nested / no sub-page folders | Graph returned no `level` values; the script prints a warning. Re-run, and check the section in Graph Explorer with `?$select=title,level,order&pagelevel=true`. |
| Images missing in a page | See `errors.csv` (`download` rows). Re-run; those pages are retried automatically. |
| Pictures show a broken-image icon in the HTML pages | Fixed by the default `--html-images embed` (re-run to upgrade older pages). With `--html-images link`, some browsers block images outside the HTML file's folder. If a picture is also missing from the Markdown file, it failed to download: check `errors.csv`. |
| `no Chrome / Chromium ... found` in `errors.csv` | Install a Chromium-based browser or set `ONENOTE_BROWSER`. Only affects the PDF/PNG copies. |
| `tesseract is not installed` | `brew install tesseract` (macOS) or your package manager, then re-run with `--ocr` on the pages you want. |
| Want to redo a finished section | Delete its folder (and `progress.json`) and run again, or export it with `S` into a new folder. |

## Testing

Development tests ran the script against a mocked Graph API and checked, among other things: the
exact back-off sequence and the stop after 10 rate-limit retries; resume after a stop (finished
pages skipped, the in-progress page redone, `progress.json` removed on completion); recovery with
no progress file and with leftover `.tmp` folders; failed image downloads and their retry;
expired tokens; nested section groups and sub-pages; image ordering, data-URI images, alternate
resource URLs; ink conversion; OCR and rendering with stand-in programs; and that every local link
in the generated HTML resolves. The mock tests are not part of this file, and live-account
behaviour beyond what Microsoft documents has not been verified.

## AI assistance

This project was developed with the assistance of an AI model (Claude, by Anthropic), directed
and reviewed by the maintainer. An earlier version of the section and notebook export was run by
the maintainer on their own notebooks. The current version, including resuming, handwriting, OCR
and PDF/PNG copies, has so far been tested mainly with mocked Microsoft Graph responses (see
[Testing](#testing)), so please review your exports and report problems as issues.

## License and disclaimer

This project is released under the [MIT License](LICENSE): anyone may use, copy, modify, merge,
publish, distribute, sublicense and sell it, provided the copyright notice and licence text are
kept.

**No warranty, no liability.** The software is provided "as is", without warranty of any kind,
and the authors are not liable for any claim, damages or other liability arising from its use
(the full wording is in [LICENSE](LICENSE)). In particular:

- Exports may be incomplete or imperfect: Microsoft Graph does not expose everything OneNote
  stores. **Keep your original notebooks** and check a sample of the export before relying on it.
- The tool only reads from your OneNote account (it makes read-only `GET` requests with the
  `Notes.Read` permission) and never changes or deletes anything there. It writes only inside
  the output folder you choose, where it may delete its own temporary files and replace a page
  it is downloading again.
- Your access token is sent only to Microsoft hosts (`graph.microsoft.com`, `www.onenote.com`,
  `onenote.com`). Treat the token like a password and do not share it, paste it into issues, or
  commit it to a repository. Exported notes may contain private information; handle and store
  them accordingly.
- Use of the Microsoft Graph API is subject to Microsoft's own terms and rate limits.

This is an independent, unofficial project. It is not affiliated with, endorsed by or sponsored
by Microsoft. "Microsoft", "OneNote" and "Microsoft Graph" are trademarks of Microsoft
Corporation.

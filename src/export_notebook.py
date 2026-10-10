#!/usr/bin/env python3
"""
Export a whole OneNote notebook (personal account) via Microsoft Graph.

Every section is exported with export_section.py, which must sit in the same folder
as this script.

Usage:
  pip3 install requests
  export ONENOTE_TOKEN="<access token from Graph Explorer>"
  python3 export_notebook.py '<notebook_id>'            # use SINGLE quotes (ids contain "!")
  python3 export_notebook.py '<notebook_id>' -o ~/Backup
  python3 export_notebook.py '<notebook_id>' --utc      # show dates in UTC (default: local time)

Output layout:
  <Notebook name>/
      index.html                      numbered list of sections -> each section's html/index.html
      metadata/
          sections.csv                one row per section (and section group)
          <Notebook name>.json        notebook + section metadata from the API (in order)
      sections/
          1_Section name/             created by export_section.py (metadata/, pages/, html/)
          2_Section group/            section groups are folders holding their sections
              2.1_Section name/
              2.2_Nested group/
                  2.2.1_Section name/

Numbering: sections come first, then section groups (the API lists them separately and
gives no mixed order). Items inside a group are numbered under it: 3, 3.1, 3.2, 3.2.1.
Every folder is named <sequence>_<name>.
The order of sections is the order Graph returns them in.
"""
import argparse
import html
import os
import sys
from datetime import datetime, timezone
from urllib.parse import quote

import export_section as es

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
    nodes = [new_node("section", s, parent) for s in es.paged(base_url + "/sections")]
    for g in es.paged(base_url + "/sectionGroups"):
        node = new_node("group", g, parent)
        node["children"] = fetch_nodes(f"{es.GRAPH}/me/onenote/sectionGroups/{g['id']}", node)
        nodes.append(node)
    return nodes


def number(nodes, prefix=""):
    for i, n in enumerate(nodes, 1):
        n["seq"] = f"{prefix}{i}"
        number(n["children"], n["seq"] + ".")


def assign_paths(nodes, parent_parts):
    """Folder for every section/group: <sequence>_<name>, nested like the hierarchy."""
    for n in nodes:
        n["folder"] = f"{n['seq']}_{es.safe_name(name_of(n), 'Section')}"
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


def render_index(nb, nodes, utc):
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
            "Exported " + html.escape(es.fmt_dt(datetime.now(timezone.utc).isoformat(), utc))]
    return es.INDEX_TEMPLATE.substitute(
        title_text=html.escape(name), css=es.INDEX_CSS, title=html.escape(name),
        topbar="OneNote export", meta=" &middot; ".join(meta),
        toc=render(nodes, True) if nodes else "<p>No sections.</p>")


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Export a OneNote notebook to local files.")
    ap.add_argument("notebook_id", help="Notebook ID (use single quotes in zsh/bash)")
    ap.add_argument("-o", "--outdir", default=".", help="parent folder (default: current)")
    ap.add_argument("--utc", action="store_true",
                    help="show dates in UTC instead of local time")
    args = ap.parse_args()

    raw = os.environ.get("ONENOTE_TOKEN") or sys.exit("Set ONENOTE_TOKEN first")
    es.set_token(raw)

    # Notebook details and structure
    nid = args.notebook_id
    try:
        nb = es.get(f"{es.GRAPH}/me/onenote/notebooks/{nid}").json()
        nodes = fetch_nodes(f"{es.GRAPH}/me/onenote/notebooks/{nid}", None)
    except es.GraphError as e:
        sys.exit(f"Could not read notebook:\n{e}")

    nb_name = nb.get("displayName") or nid
    root = os.path.join(args.outdir, es.safe_name(nb_name, "Notebook"))
    meta_dir, sections_dir = os.path.join(root, "metadata"), os.path.join(root, "sections")
    os.makedirs(meta_dir, exist_ok=True)
    os.makedirs(sections_dir, exist_ok=True)

    number(nodes)
    flat = list(flatten(nodes))
    sections = [n for n in flat if n["kind"] == "section"]
    assign_paths(nodes, ["sections"])
    for n in flat:
        if n["kind"] == "group":
            os.makedirs(os.path.join(root, *n["parts"]), exist_ok=True)

    doc = dict(nb)
    doc["sections"] = [node_json(n) for n in nodes]
    es.write_json(os.path.join(meta_dir, es.safe_name(nb_name, "notebook") + ".json"), doc)
    print(f"Notebook: {nb_name} ({len(sections)} sections) -> {os.path.abspath(root)}")

    # Export every section with export_section.py
    aborted = False
    for i, n in enumerate(sections, 1):
        print(f"\n=== Section {i}/{len(sections)} ({n['seq']}): {name_of(n)} ===")
        try:
            # from sections/.../<folder>/html/ back up to the notebook's index.html
            backlink = "../" * (len(n["parts"]) + 1) + "index.html"
            res = es.export_section(n["meta"]["id"], os.path.join(root, *n["parts"][:-1]),
                                    args.utc, folder_name=n["folder"],
                                    notebook_link=(nb_name, backlink))
            n["result"] = res
            n["status"] = "ok" if res["totals"]["ok"] == res["totals"]["pages"] else "partial"
            aborted = res["aborted"]
        except es.TokenExpired as e:
            n["status"] = "failed"
            print(f"  ! {e}\n\nToken expired. Get a fresh one from Graph Explorer and re-run.")
            aborted = True
        except es.ExportError as e:
            n["status"] = "failed"
            print(f"  ! section failed: {e}")
        if aborted:
            break

    es.write_csv(os.path.join(meta_dir, "sections.csv"), sections_csv_rows(flat))
    es.write_text(os.path.join(root, "index.html"), render_index(nb, nodes, args.utc))

    ok = sum(1 for n in sections if n["status"] == "ok")
    print(f"\nDone: {ok} of {len(sections)} sections fully exported.")
    for n in sections:
        if n["status"] != "ok":
            print(f"  {status_of(n)}: {n['seq']} {name_of(n)}")
    print(f"Open: {os.path.abspath(os.path.join(root, 'index.html'))}")


if __name__ == "__main__":
    main()

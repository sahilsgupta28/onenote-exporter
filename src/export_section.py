#!/usr/bin/env python3
"""
List all OneNote sections (with their parent notebook) via Microsoft Graph
and save them to a CSV file.

Usage:
  pip3 install requests
  export ONENOTE_TOKEN="<access token from Graph Explorer>"
  python3 list_sections.py                         # all notebooks -> sections.csv
  python3 list_sections.py -o out.csv              # custom output file
  python3 list_sections.py -n "My Notebook"        # only notebooks whose name contains "My Notebook"
"""
import argparse
import csv
import os
import re
import sys

import requests

URL = (
    "https://graph.microsoft.com/v1.0/me/onenote/sections"
    "?$expand=parentNotebook($select=id,displayName)&$top=100"
)

ap = argparse.ArgumentParser()
ap.add_argument("-o", "--output", default="sections.csv")
ap.add_argument("-n", "--notebook", help="only include notebooks whose name contains this text")
args = ap.parse_args()

raw = os.environ.get("ONENOTE_TOKEN") or sys.exit("Set ONENOTE_TOKEN first")
token = re.sub(r"\s+", "", raw).strip("\"'")
if token.lower().startswith("bearer"):
    token = token[6:]
headers = {"Authorization": f"Bearer {token}"}

# Follow @odata.nextLink so notebooks with many sections are fully listed
sections = []
url = URL
while url:
    resp = requests.get(url, headers=headers, timeout=60)
    if not resp.ok:
        sys.exit(f"HTTP {resp.status_code}\n{resp.text}")
    data = resp.json()
    sections.extend(data.get("value", []))
    url = data.get("@odata.nextLink")

rows = []
for s in sections:
    nb = s.get("parentNotebook") or {}
    nb_name = nb.get("displayName", "")
    if args.notebook and args.notebook.lower() not in nb_name.lower():
        continue
    rows.append([
        nb_name,
        nb.get("id", ""),
        s.get("displayName", ""),
        s.get("id", ""),
        s.get("createdDateTime", ""),
        s.get("lastModifiedDateTime", ""),
    ])

rows.sort(key=lambda r: (r[0].lower(), r[2].lower()))

with open(args.output, "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["Number", "Notebook name", "Notebook ID", "Section Name",
                "Section ID", "createdDateTime", "lastModifiedDateTime"])
    for i, row in enumerate(rows, 1):
        w.writerow([i] + row)

print(f"Saved {len(rows)} sections to {os.path.abspath(args.output)}")

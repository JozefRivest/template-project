#!/usr/bin/env python3
"""
org_table.py - Render org/YAML reading notes as a table

Org file structure:
  * Section Name          <- becomes a full-width separator row in the table
  ** Author Year          <- becomes an entry (Author field)
  :PROPERTIES:
  :Date: 2015
  :Journal: JJSS
  :END:

YAML file structure:
  - category: Section Name
    entries:
      - author: Author Year
        date: 2015
        journal: JJSS
        ...
    subcategories:
      - subcategory: Subsection Name
        entries:
          - author: Author Year
            date: 2015
            ...

Usage:
    python3 org_table.py readings.org
    python3 org_table.py readings.yaml
    python3 org_table.py readings.org --html
    python3 org_table.py readings.org --live
    python3 org_table.py readings.org --columns Author Date Journal Claim
    python3 org_table.py readings.org --sort Date
    python3 org_table.py readings.org --filter Journal=APSR
"""

import argparse
import os
import re
import tempfile
import threading
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Optional

try:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    _OPENPYXL = True
except ImportError:
    _OPENPYXL = False

# ── Parsing ────────────────────────────────────────────────────────────────────


def parse_yaml(filepath: str) -> list[dict]:
    try:
        import yaml
    except ImportError:
        raise SystemExit("PyYAML is required for YAML files: pip install pyyaml")

    with open(filepath, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)

    if not isinstance(data, list):
        raise SystemExit("YAML file must be a list of category objects at the top level.")

    def parse_entries(entries, category, subcategory=""):
        rows = []
        for entry in entries:
            if not isinstance(entry, dict):
                continue
            row = {"_type": "entry", "_section": str(category), "_subsection": str(subcategory)}
            for k, v in entry.items():
                row[k.capitalize() if k == "author" else k] = str(v) if v is not None else ""
            if "Author" not in row:
                row["Author"] = ""
            rows.append(row)
        return rows

    items = []
    for block in data:
        if not isinstance(block, dict):
            continue
        category = block.get("category", "")
        if category:
            items.append({"_type": "section", "title": str(category)})
        for key, value in block.items():
            if key == "entries":
                items.extend(parse_entries(value or [], category))
            elif key == "subcategories":
                for subblock in (value or []):
                    if not isinstance(subblock, dict):
                        continue
                    subcategory = subblock.get("subcategory", "")
                    if subcategory:
                        items.append({"_type": "subsection", "title": str(subcategory)})
                    items.extend(parse_entries(subblock.get("entries", []), category, subcategory))
    return items


def parse_org(filepath: str) -> list[dict]:
    items = []
    current_entry = None
    in_properties = False
    current_section = None

    with open(filepath, "r", encoding="utf-8") as f:
        lines = f.readlines()

    for line in lines:
        line = line.rstrip("\n")

        if re.match(r"^\* [^*]", line):
            if current_entry is not None:
                items.append(current_entry)
                current_entry = None
            title = re.sub(r"^\* ", "", line).strip()
            current_section = title
            items.append({"_type": "section", "title": title})
            in_properties = False
            continue

        if re.match(r"^\*{2,} ", line):
            if current_entry is not None:
                items.append(current_entry)
            heading = re.sub(r"^\*+ ", "", line).strip()
            current_entry = {
                "_type": "entry",
                "_section": current_section,
                "Author": heading,
            }
            in_properties = False
            continue

        if current_entry is None:
            continue

        if line.strip() == ":PROPERTIES:":
            in_properties = True
            continue

        if line.strip() == ":END:":
            in_properties = False
            continue

        if in_properties:
            match = re.match(r"^:([^:]+):\s*(.*)", line.strip())
            if match:
                key, value = match.group(1).strip(), match.group(2).strip()
                current_entry[key] = value

    if current_entry is not None:
        items.append(current_entry)

    return items


# ── Helpers ────────────────────────────────────────────────────────────────────


def entries_only(items):
    return [i for i in items if i.get("_type") == "entry"]


def apply_filters(items, sort_by, filter_by):
    entries = entries_only(items)

    if not entries:
        print("No entries found.")
        return None, None

    if filter_by:
        key, _, value = filter_by.partition("=")
        entries = [e for e in entries if e.get(key, "").lower() == value.lower()]
        if not entries:
            print(f"No entries match {filter_by}")
            return None, None

    if filter_by or sort_by:
        if sort_by:
            entries = sorted(entries, key=lambda e: e.get(sort_by, ""))
        return entries, False

    return items, True


def resolve_columns(entries, columns):
    if columns:
        return columns
    seen = {}
    for e in entries:
        for k in e:
            if not k.startswith("_"):
                seen[k] = True
    return list(seen.keys())


def esc(text: str) -> str:
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def strip_inline(text: str) -> str:
    """Strip **bold** and *italic*/_italic_ markers from plain-text output."""
    s = str(text)
    s = re.sub(r'\*\*(.+?)\*\*', r'\1', s, flags=re.DOTALL)
    s = re.sub(r'\*(.+?)\*', r'\1', s, flags=re.DOTALL)
    s = re.sub(r'_(.+?)_', r'\1', s, flags=re.DOTALL)
    return s


def render_inline(text: str) -> str:
    """Convert **bold** and *italic*/_italic_ markers to HTML, escaping everything else."""
    parts = re.split(r'(\*\*[^*\n]+?\*\*|\*[^*\n]+?\*|_[^_\n]+?_)', str(text))
    result = []
    for part in parts:
        if part.startswith('**') and part.endswith('**') and len(part) > 4:
            result.append(f'<strong>{esc(part[2:-2])}</strong>')
        elif (part.startswith('*') and part.endswith('*') and len(part) > 2) or \
             (part.startswith('_') and part.endswith('_') and len(part) > 2):
            result.append(f'<em>{esc(part[1:-1])}</em>')
        else:
            result.append(esc(part))
    return ''.join(result)


# ── HTML builder ───────────────────────────────────────────────────────────────


def build_html(items, filepath, columns=None, sort_by=None, filter_by=None, live=False):
    result, sectioned = apply_filters(items, sort_by, filter_by)
    if result is None:
        result = []
        sectioned = False

    all_entries = entries_only(result) if sectioned else result
    cols = resolve_columns(all_entries, columns)
    title = Path(filepath).stem
    num_cols = len(cols)

    header_cells = "".join(f"<th>{esc(col)}</th>" for col in cols)

    body_rows = ""
    if sectioned:
        for item in result:
            if item["_type"] == "section":
                body_rows += (
                    f'<tr class="section-row">'
                    f'<td colspan="{num_cols}">{esc(item["title"])}</td>'
                    f"</tr>\n"
                )
            elif item["_type"] == "subsection":
                body_rows += (
                    f'<tr class="subsection-row">'
                    f'<td colspan="{num_cols}">{esc(item["title"])}</td>'
                    f"</tr>\n"
                )
            else:
                cells = "".join(f"<td>{render_inline(item.get(col, ''))}</td>" for col in cols)
                body_rows += f"<tr>{cells}</tr>\n"
    else:
        for e in result:
            cells = "".join(f"<td>{render_inline(e.get(col, ''))}</td>" for col in cols)
            body_rows += f"<tr>{cells}</tr>\n"

    entry_count = len(all_entries)

    live_script = (
        """
    const evtSource = new EventSource("/sse");
    evtSource.onmessage = function(e) {
      if (e.data === "reload") {
        fetch("/table")
          .then(r => r.text())
          .then(html => {
            const parser = new DOMParser();
            const doc = parser.parseFromString(html, "text/html");
            document.querySelector("tbody").innerHTML =
              doc.querySelector("tbody").innerHTML;
            document.getElementById("count").textContent =
              doc.getElementById("count").textContent;
          });
      }
    };
    """
        if live
        else ""
    )

    live_badge = '<span class="live-badge">● LIVE</span>' if live else ""

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>{esc(title)}</title>
  <style>
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      font-size: 18px;
      background: #f5f5f5;
      color: #222;
      padding: 2rem;
    }}
    .header {{
      display: flex;
      align-items: center;
      gap: 0.75rem;
      margin-bottom: 1rem;
    }}
    h1 {{ font-size: 1.4rem; font-weight: 600; color: #333; }}
    .live-badge {{
      font-size: 0.75rem;
      font-weight: 600;
      color: #2d6a4f;
      background: #d8f3dc;
      padding: 2px 8px;
      border-radius: 999px;
      animation: pulse 2s infinite;
    }}
    @keyframes pulse {{
      0%, 100% {{ opacity: 1; }}
      50% {{ opacity: 0.4; }}
    }}
    .controls {{ display: flex; gap: 0.75rem; margin-bottom: 1rem; }}
    input[type="text"] {{
      padding: 7px 12px;
      border: 1px solid #ccc;
      border-radius: 6px;
      font-size: 14px;
      width: 300px;
      outline: none;
    }}
    input[type="text"]:focus {{ border-color: #2d6a4f; }}
    .table-wrapper {{
      overflow-x: auto;
      border-radius: 8px;
      box-shadow: 0 1px 4px rgba(0,0,0,0.12);
    }}
    table {{ border-collapse: collapse; width: 100%; background: #fff; }}
    thead {{ background: #2d6a4f; color: #fff; }}
    th {{
      padding: 10px 14px;
      text-align: left;
      font-weight: 600;
      white-space: nowrap;
      cursor: pointer;
      user-select: none;
    }}
    th:hover {{ background: #245a42; }}
    th.sorted-asc::after  {{ content: " ▲"; font-size: 0.75em; }}
    th.sorted-desc::after {{ content: " ▼"; font-size: 0.75em; }}
    td {{
      padding: 9px 14px;
      border-bottom: 1px solid #eee;
      vertical-align: top;
      max-width: 300px;
      word-wrap: break-word;
    }}
    tr.section-row td {{
      background: #b7e4c7;
      color: #1b4332;
      font-weight: 700;
      font-size: 0.9rem;
      letter-spacing: 0.04em;
      text-transform: uppercase;
      padding: 7px 14px;
      border-bottom: 2px solid #74c69d;
    }}
    tr.subsection-row td {{
      background: #d8f3dc;
      color: #2d6a4f;
      font-weight: 600;
      font-size: 0.85rem;
      letter-spacing: 0.02em;
      padding: 5px 20px;
      border-bottom: 1px solid #95d5b2;
    }}
    tr:last-child td {{ border-bottom: none; }}
    tr:not(.section-row):nth-child(even) {{ background: #f9f9f9; }}
    tr:not(.section-row):hover {{ background: #eaf4ee; }}
    .count {{ margin-top: 0.75rem; font-size: 0.85rem; color: #888; }}
  </style>
</head>
<body>
  <div class="header">
    <h1>{esc(title)}</h1>
    {live_badge}
  </div>
  <div class="controls">
    <input type="text" id="search" placeholder="Search entries…" oninput="filterTable()" />
  </div>
  <div class="table-wrapper">
    <table id="main-table">
      <thead><tr>{header_cells}</tr></thead>
      <tbody>{body_rows}</tbody>
    </table>
  </div>
  <p class="count" id="count">{entry_count} entries</p>

  <script>
    function filterTable() {{
      const query = document.getElementById("search").value.toLowerCase();
      const rows = document.querySelectorAll("#main-table tbody tr");
      let visible = 0;
      let lastSection = null;
      rows.forEach(row => {{
        if (row.classList.contains("section-row")) {{
          lastSection = row;
          row.style.display = "none";
          return;
        }}
        const match = row.textContent.toLowerCase().includes(query);
        row.style.display = match ? "" : "none";
        if (match) {{
          visible++;
          if (lastSection) {{ lastSection.style.display = ""; lastSection = null; }}
        }}
      }});
      document.getElementById("count").textContent = visible + " entries";
    }}

    let sortCol = -1, sortAsc = true;
    document.querySelectorAll("#main-table thead th").forEach((th, idx) => {{
      th.addEventListener("click", () => {{
        const tbody = document.querySelector("#main-table tbody");
        const rows = Array.from(tbody.querySelectorAll("tr:not(.section-row)"));
        sortAsc = sortCol === idx ? !sortAsc : true;
        sortCol = idx;
        rows.sort((a, b) => {{
          const aText = a.cells[idx]?.textContent.trim() ?? "";
          const bText = b.cells[idx]?.textContent.trim() ?? "";
          return sortAsc ? aText.localeCompare(bText) : bText.localeCompare(aText);
        }});
        rows.forEach(r => tbody.appendChild(r));
        document.querySelectorAll("#main-table thead th").forEach(h =>
          h.classList.remove("sorted-asc", "sorted-desc"));
        th.classList.add(sortAsc ? "sorted-asc" : "sorted-desc");
      }});
    }});

    {live_script}
  </script>
</body>
</html>"""


# ── Terminal rendering ─────────────────────────────────────────────────────────


def render_table(items, columns=None, sort_by=None, filter_by=None, max_width=30):
    result, sectioned = apply_filters(items, sort_by, filter_by)
    if result is None:
        return

    all_entries = entries_only(result) if sectioned else result
    cols = resolve_columns(all_entries, columns)

    def truncate(text):
        return text[: max_width - 1] + "…" if len(text) > max_width else text

    all_rows = [[truncate(strip_inline(e.get(col, ""))) for col in cols] for e in all_entries]
    widths = [
        max(len(col), max((len(r[i]) for r in all_rows), default=0))
        for i, col in enumerate(cols)
    ]
    total_width = sum(widths) + 3 * len(widths) + 1

    def fmt_row(cells):
        return "| " + " | ".join(c.ljust(widths[i]) for i, c in enumerate(cells)) + " |"

    print(fmt_row(cols))
    print("|-" + "-+-".join("-" * w for w in widths) + "-|")

    if sectioned:
        for item in result:
            if item["_type"] == "section":
                label = f"  {item['title'].upper()}  "
                print("+" + label.center(total_width - 2, "-") + "+")
            elif item["_type"] == "subsection":
                label = f"  {item['title']}  "
                print("|" + label.center(total_width - 2, "·") + "|")
            else:
                print(fmt_row([truncate(strip_inline(item.get(col, ""))) for col in cols]))
    else:
        for e in result:
            print(fmt_row([truncate(strip_inline(e.get(col, ""))) for col in cols]))


# ── XLSX export ────────────────────────────────────────────────────────────────


def write_xlsx(
    items, filepath, out_path=None, columns=None, sort_by=None, filter_by=None
):
    if not _OPENPYXL:
        raise SystemExit(
            "openpyxl is required for xlsx export: sudo pacman -S python-openpyxl"
        )

    result, sectioned = apply_filters(items, sort_by, filter_by)
    if result is None:
        return

    all_entries = entries_only(result) if sectioned else result
    cols = resolve_columns(all_entries, columns)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = Path(filepath).stem[:31]

    header_fill = PatternFill("solid", fgColor="2D6A4F")
    header_font = Font(bold=True, color="FFFFFF")
    section_fill = PatternFill("solid", fgColor="B7E4C7")
    section_font = Font(bold=True, color="1B4332")
    subsection_fill = PatternFill("solid", fgColor="D8F3DC")
    subsection_font = Font(bold=True, color="2D6A4F")

    # Header row
    for col_idx, col in enumerate(cols, 1):
        cell = ws.cell(row=1, column=col_idx, value=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=False)

    row_idx = 2
    if sectioned:
        for item in result:
            if item["_type"] == "section":
                ws.merge_cells(
                    start_row=row_idx,
                    start_column=1,
                    end_row=row_idx,
                    end_column=len(cols),
                )
                cell = ws.cell(row=row_idx, column=1, value=item["title"].upper())
                cell.fill = section_fill
                cell.font = section_font
            elif item["_type"] == "subsection":
                ws.merge_cells(
                    start_row=row_idx,
                    start_column=1,
                    end_row=row_idx,
                    end_column=len(cols),
                )
                cell = ws.cell(row=row_idx, column=1, value=item["title"])
                cell.fill = subsection_fill
                cell.font = subsection_font
                cell.alignment = Alignment(indent=1)
            else:
                for col_idx, col in enumerate(cols, 1):
                    ws.cell(row=row_idx, column=col_idx, value=strip_inline(item.get(col, "")))
            row_idx += 1
    else:
        for entry in result:
            for col_idx, col in enumerate(cols, 1):
                ws.cell(row=row_idx, column=col_idx, value=strip_inline(entry.get(col, "")))
            row_idx += 1

    # Auto-fit column widths (capped at 60)
    for col_idx, col in enumerate(cols, 1):
        max_len = len(col)
        for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
            for cell in row:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 60)

    ws.freeze_panes = "A2"

    dest = out_path or Path(filepath).with_suffix(".xlsx")
    wb.save(dest)
    print(f"Saved: {dest}")


# ── Live server ────────────────────────────────────────────────────────────────


def start_live_server(filepath, columns, sort_by, filter_by, port=8765):
    sse_clients = []
    lock = threading.Lock()

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, format, *args):
            pass

        def do_GET(self):
            if self.path in ("/", "/table"):
                items = parse_yaml(filepath) if Path(filepath).suffix in (".yaml", ".yml") else parse_org(filepath)
                html = build_html(
                    items, filepath, columns, sort_by, filter_by, live=True
                )
                self._respond(200, "text/html", html.encode())

            elif self.path == "/sse":
                self.send_response(200)
                self.send_header("Content-Type", "text/event-stream")
                self.send_header("Cache-Control", "no-cache")
                self.send_header("Connection", "keep-alive")
                self.end_headers()
                with lock:
                    sse_clients.append(self.wfile)
                try:
                    while True:
                        time.sleep(1)
                except (BrokenPipeError, ConnectionResetError):
                    pass
                finally:
                    with lock:
                        if self.wfile in sse_clients:
                            sse_clients.remove(self.wfile)
            else:
                self._respond(404, "text/plain", b"Not found")

        def _respond(self, code, ctype, body):
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

    def watch_file():
        last_mtime = os.path.getmtime(filepath)
        while True:
            time.sleep(0.5)
            try:
                mtime = os.path.getmtime(filepath)
                if mtime != last_mtime:
                    last_mtime = mtime
                    msg = b"data: reload\n\n"
                    with lock:
                        dead = []
                        for client in sse_clients:
                            try:
                                client.write(msg)
                                client.flush()
                            except Exception:
                                dead.append(client)
                        for d in dead:
                            sse_clients.remove(d)
            except FileNotFoundError:
                pass

    # ThreadingHTTPServer handles each connection in its own thread,
    # so the SSE connection no longer blocks /table requests.
    for attempt in range(10):
        try:
            server = ThreadingHTTPServer(("localhost", port + attempt), Handler)
            port = port + attempt
            break
        except OSError:
            if attempt == 9:
                raise
            continue
    threading.Thread(target=watch_file, daemon=True).start()

    url = f"http://localhost:{port}"
    print(f"Serving at {url}  (Ctrl+C to stop)")
    webbrowser.open(url)

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")


# ── CLI ────────────────────────────────────────────────────────────────────────


def main():
    parser = argparse.ArgumentParser(description="Render org properties as a table.")
    parser.add_argument("file", help="Path to the org file")
    parser.add_argument("--columns", "-c", nargs="+", help="Columns to display")
    parser.add_argument("--sort", "-s", help="Sort by property (e.g. Date)")
    parser.add_argument(
        "--filter", "-f", help="Filter by property=value (e.g. Journal=APSR)"
    )
    parser.add_argument(
        "--max-width",
        "-w",
        type=int,
        default=30,
        help="Max cell width in terminal (default: 30)",
    )
    parser.add_argument(
        "--html", action="store_true", help="Open as styled HTML table in browser"
    )
    parser.add_argument(
        "--live", action="store_true", help="Live update in browser as you edit"
    )
    parser.add_argument(
        "--port", type=int, default=8765, help="Port for live server (default: 8765)"
    )
    parser.add_argument(
        "--xlsx",
        metavar="OUT",
        nargs="?",
        const="",
        help="Export to xlsx (optional output path)",
    )
    args = parser.parse_args()

    items = parse_yaml(args.file) if Path(args.file).suffix in (".yaml", ".yml") else parse_org(args.file)

    if args.xlsx is not None:
        out = args.xlsx if args.xlsx else None
        write_xlsx(items, args.file, out, args.columns, args.sort, args.filter)
    elif args.live:
        start_live_server(args.file, args.columns, args.sort, args.filter, args.port)
    elif args.html:
        html = build_html(items, args.file, args.columns, args.sort, args.filter)
        tmp = tempfile.NamedTemporaryFile(
            delete=False, suffix=".html", mode="w", encoding="utf-8"
        )
        tmp.write(html)
        tmp.close()
        webbrowser.open(f"file://{tmp.name}")
        print(f"Opened: {tmp.name}")
    else:
        render_table(items, args.columns, args.sort, args.filter, args.max_width)


if __name__ == "__main__":
    main()

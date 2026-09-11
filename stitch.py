"""
MICS Newsletter Stitcher
Runs AFTER weekly_market_snapshot_ExcelHTML.py.

Each week you drop your newsletter HTML (with the story written) into the
input/ folder. The file can have ANY name. This script:
  - picks up the newest .html in input/
  - drops in the fresh market table
  - sets the current Saturday date
  - sets the issue number automatically (it counts up on its own)
  - writes the finished file into the latest/ folder

The story text is never touched. Only the table, the date, and the issue
number are refreshed.
"""

import datetime
import glob
import os
import re
import shutil

INPUT_DIR = "input"
OUTPUT_DIR = "latest"
ISSUE_FILE = "issue.txt"   # remembers the last issue number between runs

# ── DATE HELPERS ────────────────────────────────────────────────────────────

def ordinal(n):
    if 11 <= (n % 100) <= 13:
        return f"{n}th"
    return f"{n}{ {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th') }"

def masthead_date(d):
    return f"{d.strftime('%A, %B')} {ordinal(d.day)} {d.year}"

def title_date(d):
    return f"{d.strftime('%B')} {ordinal(d.day)}, {d.year}"

# ── FIND THE INPUT NEWSLETTER (any name) ────────────────────────────────────

def find_newsletter():
    files = sorted(glob.glob(os.path.join(INPUT_DIR, "*.html")),
                   key=os.path.getmtime)
    if not files:
        raise SystemExit(
            f"No .html file found in the {INPUT_DIR}/ folder. Upload your "
            "newsletter there (any name) and run again."
        )
    return files[-1]   # newest wins if there is more than one

# ── FIND THE FRESH MARKET SNIPPET ───────────────────────────────────────────

def find_snippet():
    files = [f for f in glob.glob("MICS_Market_Snapshot_*.html")
             if "_PARTIAL" not in f]
    if not files:
        raise SystemExit(
            "No FINAL snapshot HTML found. Did the snapshot script run, and did "
            "it see today as Saturday? Check the TZ setting in the workflow."
        )
    return max(files, key=os.path.getmtime)

# ── ISSUE NUMBER ────────────────────────────────────────────────────────────

def next_issue(html):
    """Counts up on its own, ignoring whatever number is in the uploaded file."""
    if os.path.exists(ISSUE_FILE):
        last = int(open(ISSUE_FILE).read().strip())
    else:
        m = re.search(r"Issue &middot; (\d+)", html)   # seed from the file once
        last = int(m.group(1)) if m else 0
    nxt = last + 1
    with open(ISSUE_FILE, "w") as f:
        f.write(str(nxt))
    return nxt

# ── STITCH ──────────────────────────────────────────────────────────────────

def main():
    today = datetime.date.today()   # IST, because the workflow sets TZ

    newsletter_path = find_newsletter()
    with open(newsletter_path, "r", encoding="utf-8", newline="") as f:
        html = f.read()

    snippet_path = find_snippet()
    with open(snippet_path, "r", encoding="utf-8", newline="") as f:
        snippet = f.read().strip()
    snippet = snippet.replace("\r\n", "\n").replace("\n", "\r\n")

    # 1) swap the whole market-snapshot block (marker up to the STORY marker)
    block = re.compile(
        r"<!-- \u2550\u2550\u2550 MARKET SNAPSHOT.*?(?=<!-- \u2550\u2550\u2550 STORY)",
        re.DOTALL,
    )
    if not block.search(html):
        raise SystemExit(
            "Could not find the MARKET SNAPSHOT ... STORY markers in "
            f"{os.path.basename(newsletter_path)}. Those comment markers must stay."
        )
    html = block.sub(snippet + "\r\n\r\n    ", html)

    # 2) issue number, set automatically
    issue = next_issue(html)
    html, n_issue = re.subn(r"(Issue &middot; )(\d+)",
                            lambda m: f"{m.group(1)}{issue:03d}", html)

    # 3) masthead date
    date_div = re.compile(
        r"(color:rgba\(255,255,255,0\.90\);margin-top:4px;letter-spacing:0\.5px;\">\s*)"
        r"([^<]+?)(\s*</div>)",
        re.DOTALL,
    )
    html, n_date = date_div.subn(
        lambda m: f"{m.group(1)}{masthead_date(today)}{m.group(3)}", html)

    # 4) <title> date
    html, n_title = re.subn(
        r"(<title>GlobalMarketsBrief_).*?( - MICS Wealth</title>)",
        lambda m: f"{m.group(1)}{title_date(today)}{m.group(2)}",
        html,
    )

    # write the finished file into the one folder you download from
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    stable = os.path.join(OUTPUT_DIR, "Global_Market_Brief_latest.html")
    dated = os.path.join(OUTPUT_DIR, f"Global_Market_Brief_{today:%Y-%m-%d}.html")
    with open(stable, "w", encoding="utf-8", newline="") as f:
        f.write(html)
    shutil.copyfile(stable, dated)

    xlsx = sorted(glob.glob("MICS_Market_Snapshot_*.xlsx"), key=os.path.getmtime)
    if xlsx:
        shutil.copyfile(xlsx[-1],
                        os.path.join(OUTPUT_DIR, os.path.basename(xlsx[-1])))

    print(f"Input used   : {newsletter_path}")
    print(f"Snippet used : {snippet_path}")
    print(f"Issue set to : {issue:03d} ({n_issue} place)")
    print(f"Date set to  : {masthead_date(today)} ({n_date} masthead, {n_title} title)")
    print(f"Written to   : {stable}")
    print(f"Also saved   : {dated}")


if __name__ == "__main__":
    main()

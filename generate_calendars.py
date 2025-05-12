import sys
import csv
import re
import io
import os
import hashlib
import requests
import unicodedata
from datetime import datetime, timedelta
from pathlib import Path
from collections import defaultdict
from ics import Calendar, Event
from ics.grammar.parse import ContentLine

# --- Configuration ---
CSV_URL = os.getenv("SHEET_CSV_URL")
OUTPUT_DIR = Path(sys.argv[1] if len(sys.argv) > 1 else "docs")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

if not CSV_URL:
    print("❌ SHEET_CSV_URL environment variable is not set.")
    exit(1)

def clean(text):
    if not isinstance(text, str):
        text = str(text)
    return re.sub(r"\s+", " ", unicodedata.normalize("NFKC", text.strip()))

def generate_uid(date, team, content):
    return hashlib.md5(f"{date.isoformat()}-{team}-{content}".encode()).hexdigest()

def generate_index_html(output_dir: Path, teams: dict):
    output_path = output_dir / "index.html"
    links = [
        '<li><a href="{}.ics">{}</a></li>'.format(
            re.sub(r"[^\w]+", "_", team.lower()).strip("_"), team
        )
        for team in sorted(teams)
    ]
    html_list = "<ul>\n" + "\n".join(links) + "\n</ul>"
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(f"<html><body><h1>Team Calendars</h1>\n{html_list}\n</body></html>")
    print(f"📝 index.html written to {output_path}")

# --- Fetch CSV ---
print(f"🔄 Downloading schedule from:\n{CSV_URL}")
response = requests.get(CSV_URL)
response.raise_for_status()

try:
    decoded = response.content.decode("utf-8")
except UnicodeDecodeError as e:
    print("❌ UTF-8 decoding failed:", e)
    exit(1)

reader = csv.reader(io.StringIO(decoded))
rows = [[clean(cell) for cell in row] for row in reader]
print(f"✅ Loaded {len(rows)} rows")

# --- Detect calendar blocks with dynamic year tracking ---
blocks_by_row = {}
month_sequence = []
year_for_month = {}
base_year = datetime.now().year
current_year = base_year

for row_index, row in enumerate(rows):
    col = 0
    blocks = []
    while col < len(row) - 1:
        if re.fullmatch(r"\d{1,2}月", row[col]) and row[col + 1] == "":
            try:
                month = int(row[col].replace("月", ""))
            except ValueError:
                col += 1
                continue

            # Handle year rollover
            if month_sequence and month < month_sequence[-1]:
                current_year += 1
            if month not in year_for_month:
                year_for_month[month] = current_year
                month_sequence.append(month)

            team_cols = {}
            scan = col + 2
            while scan < len(row):
                if re.fullmatch(r"\d{1,2}月", row[scan]):
                    break
                name = clean(row[scan])
                if name:
                    team_cols[scan] = name
                scan += 1
            if team_cols:
                blocks.append((col, month, team_cols))
            col = scan
        else:
            col += 1
    if blocks:
        blocks_by_row[row_index] = blocks

# --- Parse event rows ---
active_blocks = []
events_by_team = defaultdict(list)
event_count = 0

for row_index, row in enumerate(rows):
    if row_index in blocks_by_row:
        active_blocks = blocks_by_row[row_index]
        continue

    for start_col, month, team_cols in active_blocks:
        if start_col >= len(row):
            continue
        day_str = clean(row[start_col])
        if not re.fullmatch(r"\d{1,2}", day_str):
            continue
        try:
            year = year_for_month.get(month, base_year)
            date = datetime(year, month, int(day_str))
        except ValueError:
            continue

        for col, team in team_cols.items():
            if col < len(row):
                content = clean(row[col])
                if content:
                    events_by_team[team].append((date, content))
                    event_count += 1

# --- Write ICS files ---
print(f"\n📊 Found {len(events_by_team)} teams and {event_count} total events.")
for team, events in events_by_team.items():
    cal = Calendar()
    cal.extra.append(ContentLine(name="X-WR-CALNAME", value=team))
    cal.extra.append(ContentLine(name="X-WR-TIMEZONE", value="Asia/Tokyo"))

    for date, desc in events:
        event = Event()
        event.name = desc
        event.begin = date.date()
        event.make_all_day()
        event.uid = generate_uid(date, team, desc)
        cal.events.add(event)

    team_slug = re.sub(r"[^\w]+", "_", team.lower()).strip("_")
    ics_path = OUTPUT_DIR / f"{team_slug}.ics"
    with open(ics_path, "w", encoding="utf-8") as f:
        f.writelines(cal)

    print(f"✅ {team}: {len(events)} events → {ics_path.name}")

generate_index_html(OUTPUT_DIR, events_by_team)
print("🎉 All calendars and index.html generated.")

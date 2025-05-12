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
from ics import Calendar, Event
from ics.grammar.parse import ContentLine

# --- Configuration ---
CURRENT_YEAR = datetime.now().year
CSV_URL = os.getenv("SHEET_CSV_URL")
OUTPUT_DIR = Path(sys.argv[1] if len(sys.argv) > 1 else "docs")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# --- Require SHEET_CSV_URL ---
if not CSV_URL:
    print("❌ SHEET_CSV_URL environment variable is not set.")
    exit(1)

# --- Helpers ---
def clean(text):
    if not isinstance(text, str):
        text = str(text)
    return re.sub(r"\s+", " ", unicodedata.normalize("NFKC", text.strip()))

def generate_uid(date, team, content):
    return hashlib.md5(f"{date.isoformat()}-{team}-{content}".encode()).hexdigest()

def generate_index_html(output_dir: Path, teams: dict):
    template_path = Path("index_template.html")
    if not template_path.exists():
        print("⚠️ index_template.html not found. Skipping index generation.")
        return

    with open(template_path, encoding="utf-8") as f:
        template = f.read()

    # Build list from teams (team slug and display name)
    links = [
        f'<li><a href="{re.sub(r"[^\w]+", "_", team.lower()).strip("_")}.ics">{team}</a></li>'
        for team in sorted(teams)
    ]
    html_list = "<ul>\n" + "\n".join(links) + "\n</ul>"

    result = template.replace("<!-- CALENDAR_LIST -->", html_list)
    output_path = output_dir / "index.html"

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(result)

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

if not rows or not rows[0]:
    print("⚠️ No data or header found.")
    exit(1)

# --- Detect header blocks ---
header = rows[0]
blocks = []
col = 0
while col < len(header):
    cell = header[col]
    if re.fullmatch(r"\d{1,2}月", cell):
        try:
            month = int(cell.replace("月", ""))
        except ValueError:
            col += 1
            continue

        team_columns = {}
        scan_col = col + 2
        while scan_col < len(header):
            if re.fullmatch(r"\d{1,2}月", header[scan_col]):
                break
            team = clean(header[scan_col])
            if team:
                team_columns[scan_col] = team
            scan_col += 1

        if team_columns:
            blocks.append((col, month, team_columns))
        col = scan_col
    else:
        col += 1

if not blocks:
    print("❌ No calendar blocks found.")
    exit(1)

# --- Collect events ---
events_by_team = {}
event_count = 0

for row in rows[1:]:
    for block_col, month, team_map in blocks:
        if block_col >= len(row):
            continue
        day_str = clean(row[block_col])
        if not re.fullmatch(r"\d{1,2}", day_str):
            continue
        try:
            date = datetime(CURRENT_YEAR, month, int(day_str))
        except ValueError:
            continue

        for col, team in team_map.items():
            if col >= len(row):
                continue
            content = clean(row[col])
            if not content:
                continue
            events_by_team.setdefault(team, []).append((date, content))
            event_count += 1

# --- Generate .ics files ---
print(f"\n📊 Found {len(events_by_team)} teams and {event_count} total events.")
for team, events in events_by_team.items():
    cal = Calendar()
    cal.extra.append(ContentLine(name="X-WR-CALNAME", value=team))
    cal.extra.append(ContentLine(name="X-WR-TIMEZONE", value="Asia/Tokyo"))

    for date, desc in events:
        event = Event()
        event.name = desc
        event.begin = date.isoformat()
        event.duration = timedelta(hours=1)
        event.uid = generate_uid(date, team, desc)
        cal.events.add(event)

    team_slug = re.sub(r"[^\w]+", "_", team.lower()).strip("_")
    ics_path = OUTPUT_DIR / f"{team_slug}.ics"
    with open(ics_path, "w", encoding="utf-8") as f:
        f.writelines(cal)

    print(f"✅ {team}: {len(events)} events → {ics_path.name}")

# --- Generate index.html ---
generate_index_html(OUTPUT_DIR, events_by_team)

print("🎉 All calendars and index.html generated.")

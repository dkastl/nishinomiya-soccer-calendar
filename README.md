# Nishinomoya Soccer School Team Calendar Generator

This project generates `.ics` calendar files for Nishinomiya Soccer School teams based on a shared Google Sheet.

## Features

- Pulls schedule from a public Google Sheet (published CSV)
- Extracts team-specific events from a calendar-style layout
- Outputs `.ics` files (iCalendar) for each team
- Designed to be hosted via GitHub Pages

## Requirements

- Python 3.12+
- `pip` for installing dependencies

## Setup

Clone the repository and install dependencies:

```bash
git clone https://github.com/dkastl/nishinomiya-soccer-calendar.git
cd nishinomiya-soccer-calendar
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

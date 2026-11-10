# snip-to-cal

Windows tool: press a hotkey, snip any text with dates (email, syllabus, announcement), and the events are added to Google Calendar with a 60-min popup reminder. A small card shows what was added, with Undo / Open / Close.

## Setup (Windows only)

1. `py -3.12 -m venv .venv`
2. `.venv\Scripts\pip install -r requirements.txt`
3. Create `.env` with `ANTHROPIC_API_KEY=...`
4. Put your Google OAuth desktop `credentials.json` in this folder.
5. `.venv\Scripts\python snip_to_cal.py --auth` (one-time Google sign-in)
6. Optional hotkey, in AutoHotkey v1:
   `!r::Run "<folder>\.venv\Scripts\pythonw.exe" "<folder>\snip_to_cal.py"`

## Run

- `python snip_to_cal.py`: grab the screen center (10% cropped off each side) and add events
- `python snip_to_cal.py --snip`: drag a manual snip instead
- `python snip_to_cal.py img.png`: test on an existing image

## Tech

Python, Claude Haiku (Anthropic API), Google Calendar API, Windows Snipping Tool (`ms-screenclip`), pywin32, tkinter, AutoHotkey.

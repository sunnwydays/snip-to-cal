"""Snip -> Google Calendar.

Run once per snip (AutoHotkey hotkey). Opens the Windows snip overlay, waits for
the image on the clipboard, asks Haiku for the events, adds them to Google
Calendar with a 60-min popup reminder, then shows a card with Undo.

python snip_to_cal.py            grab the screen center (10% cropped per side)
python snip_to_cal.py --snip     manual snip overlay instead
python snip_to_cal.py img.png    use an image file instead (testing)
python snip_to_cal.py --auth     sign in to Google once
"""
import base64
import ctypes
import datetime as dt
import io
import json
import os
import sys
import threading
import time
import traceback
import webbrowser
from pathlib import Path
from zoneinfo import ZoneInfo

ROOT = Path(__file__).resolve().parent
CAPTURES = ROOT / "captures"
LOG = ROOT / "log.jsonl"
TZ = "America/Toronto"		# event timezone
MODEL = "claude-haiku-5-5"	# reads the screenshot
REMINDER_MIN = 60			# popup reminder, min before start
SNIP_TIMEOUT_S = 60			# --snip wait limit, seconds
CROP_PCT = 0.10				# trimmed per screen edge (10%)
SCOPES = ["https://www.googleapis.com/auth/calendar.events"]
HIDE_TIME_EVENT_ADDED = 5	# card seconds after adding
HIDE_TIME_NO_EVENT = 5		# card seconds when nothing found

# ---------------------------------------------------------------- Google auth

def calendar_service():
    from google.auth.exceptions import RefreshError
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build

    token = ROOT / "token.json"
    creds = None
    if token.exists():
        creds = Credentials.from_authorized_user_file(str(token), SCOPES)
    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except RefreshError:
            creds = None  # token revoked/expired: sign in again
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(str(ROOT / "credentials.json"), SCOPES)
        creds = flow.run_local_server(port=0)
    token.write_text(creds.to_json())
    return build("calendar", "v3", credentials=creds, cache_discovery=False)

# ---------------------------------------------------------------- snip

def screen_image():
    """Grab the primary screen, minus CROP_PCT padding on every side."""
    from PIL import ImageGrab

    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # real pixels, not scaled
    except Exception:
        pass
    img = ImageGrab.grab()
    w, h = img.size
    dx, dy = int(w * CROP_PCT), int(h * CROP_PCT)
    return img.crop((dx, dy, w - dx, h - dy))


def snip_image():
    """Open the snip overlay and return the new clipboard image (or None)."""
    import win32clipboard
    from PIL import ImageGrab

    seq = win32clipboard.GetClipboardSequenceNumber()
    os.startfile("ms-screenclip:")
    end = time.time() + SNIP_TIMEOUT_S
    while time.time() < end:
        time.sleep(0.2)
        if win32clipboard.GetClipboardSequenceNumber() != seq:
            time.sleep(0.2)  # let the snip tool finish writing all formats
            for _ in range(10):
                try:
                    img = ImageGrab.grabclipboard()
                    break
                except OSError:  # clipboard busy
                    time.sleep(0.1)
            else:
                img = None
            if hasattr(img, "save"):
                return img
            seq = win32clipboard.GetClipboardSequenceNumber()  # text copied etc.: keep waiting
    return None

# ---------------------------------------------------------------- Haiku

EVENT_TOOL = {
    "name": "add_events",
    "description": "Calendar events found in the image.",
    "input_schema": {
        "type": "object",
        "properties": {
            "events": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "title": {"type": "string", "description": "Short, e.g. 'ECE345 Midterm' or 'Due: ECE302 PS7'"},
                        "date": {"type": "string", "description": "YYYY-MM-DD"},
                        "start": {"type": ["string", "null"], "description": "HH:MM 24h, null if all-day"},
                        "end": {"type": ["string", "null"], "description": "HH:MM 24h, null if unknown"},
                        "location": {"type": ["string", "null"]},
                        "notes": {"type": ["string", "null"], "description": "Useful details from the image, 1-2 lines"},
                        "unsure": {"type": ["string", "null"], "description": "What you could not read or had to guess, else null"},
                    },
                    "required": ["title", "date", "start", "end", "location", "notes", "unsure"],
                },
            },
            "none_reason": {"type": ["string", "null"], "description": "If no events: why"},
        },
        "required": ["events", "none_reason"],
    },
}


def extract_events(png: bytes):
    import anthropic
    from dotenv import load_dotenv

    load_dotenv(ROOT / ".env")
    now = dt.datetime.now(ZoneInfo(TZ))
    system = (
        f"You read a screenshot (email, syllabus, announcement, table) and pull out calendar events.\n"
        f"Now: {now:%A %Y-%m-%d %H:%M}, timezone {TZ}. Resolve relative dates ('next Friday', 'tomorrow') from this.\n"
        "Rules:\n"
        "- One event per dated item. Skip things with no date.\n"
        "- No year shown: use the next occurrence on or after today.\n"
        "- Deadline with a time ('due 11:59 pm'): title 'Due: <thing>', start = due time, end null.\n"
        "- No time at all: start and end null (all-day).\n"
        "- Course codes (e.g. ECE345) go at the start of the title.\n"
        "- Relative dates in a dated post/email resolve from that post's date, not today.\n"
        "- `unsure`: only when you genuinely could not read something or two readings are plausible "
        "(blurry digit, AM vs PM unclear, which week 'next Friday' means). Max 12 words. "
        "Picking the next occurrence for a missing year is normal, not unsure. Otherwise null.\n"
        "- Skip cancellations and non-events ('office hours cancelled'); mention them nowhere.\n"
        "- Ignore dates that are only context (e.g. 'posted on', email sent date).\n"
        "- Only use dates actually written in the image or derivable from it. If an item has no visible date "
        "(e.g. a calendar grid with no date headers), never assume today: skip it. "
        "Don't duplicate: the same title/time/place appearing twice is one event.\n"
        "- `location`: only for University of Toronto rooms, write '<CODE> <room>', e.g. 'MY 150', 'BA 4134' "
        "(no person names, parentheses or building names). Codes: Myhal MY, Galbraith GA, Bahen BA, "
        "Exam Centre EX, Sandford Fleming SF, Mechanical Engineering MC, Medical Sciences MS; other UofT "
        "buildings use their standard 2-3 letter code. Any other place (restaurant, address, non-UofT venue, "
        "Zoom link): write it exactly as shown, unchanged.\n"
        "- Use JSON null, never the string \"null\"."
    )
    client = anthropic.Anthropic()
    msg = client.messages.create(
        model=MODEL,
        max_tokens=2000,
        system=system,
        tools=[EVENT_TOOL],
        tool_choice={"type": "tool", "name": "add_events"},
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": "image/png",
                                          "data": base64.b64encode(png).decode()}},
            {"type": "text", "text": "Extract the events."},
        ]}],
    )
    out = next(b for b in msg.content if b.type == "tool_use").input
    for ev in out.get("events", []):
        for k, v in ev.items():
            if isinstance(v, str) and v.strip().lower() in ("", "null", "none"):
                ev[k] = None
    return out

# ---------------------------------------------------------------- Calendar

def to_body(ev):
    title = ev["title"] if not ev.get("unsure") else "?? " + ev["title"]
    notes = "\n".join(x for x in [ev.get("notes"), ev.get("unsure") and f"Unsure: {ev['unsure']}",
                                  "(added by snip-to-cal)"] if x)
    body = {
        "summary": title,
        "description": notes,
        "reminders": {"useDefault": False, "overrides": [{"method": "popup", "minutes": REMINDER_MIN}]},
    }
    if ev.get("location"):
        body["location"] = ev["location"]
    day = dt.date.fromisoformat(ev["date"])
    if ev.get("start"):
        start = dt.datetime.combine(day, dt.time.fromisoformat(ev["start"]))
        end = (dt.datetime.combine(day, dt.time.fromisoformat(ev["end"])) if ev.get("end") else None)
        if not end or end <= start:
            end = start + (dt.timedelta(minutes=30) if title.startswith(("Due", "?? Due")) else dt.timedelta(hours=1))
        body["start"] = {"dateTime": start.isoformat(), "timeZone": TZ}
        body["end"] = {"dateTime": end.isoformat(), "timeZone": TZ}
    else:
        body["start"] = {"date": day.isoformat()}
        body["end"] = {"date": (day + dt.timedelta(days=1)).isoformat()}
    return body


def when_text(ev):
    day = dt.date.fromisoformat(ev["date"])
    s = f"{day:%a %b} {day.day}"
    fmt = lambda t: dt.datetime.strptime(t, "%H:%M").strftime("%I:%M %p").lstrip("0")
    if ev.get("start"):
        s += f", {fmt(ev['start'])}" + (f"–{fmt(ev['end'])}" if ev.get("end") else "")
    else:
        s += " (all day)"
    return s

# ---------------------------------------------------------------- popup

BG, FG, DIM, WARN, ACC = "#1f1f23", "#ececf1", "#9a9aa5", "#f2b84b", "#7aa2f7"


class Card:
    def __init__(self):
        import tkinter as tk
        self.tk = tk
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
        self.root = tk.Tk()
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.configure(bg=BG, highlightthickness=1, highlightbackground="#3a3a42")
        self.frame = tk.Frame(self.root, bg=BG, padx=14, pady=10)
        self.frame.pack(fill="both")
        self.hover = False
        self.root.bind("<Enter>", lambda e: setattr(self, "hover", True))
        self.root.bind("<Leave>", lambda e: setattr(self, "hover", False))
        self.root.withdraw()

    def label(self, text, fg=FG, bold=False, size=10):
        self.tk.Label(self.frame, text=text, fg=fg, bg=BG, justify="left", anchor="w", wraplength=380,
                      font=("Segoe UI", size, "bold" if bold else "normal")).pack(fill="x", anchor="w")

    def show(self, build, close_after=None):
        for w in self.frame.winfo_children():
            w.destroy()
        build()
        self.root.deiconify()
        self.root.update_idletasks()
        w, h = self.root.winfo_reqwidth(), self.root.winfo_reqheight()
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        self.root.geometry(f"+{sw - w - 16}+{sh - h - 64}")
        if close_after:
            self.deadline = time.time() + close_after
            self.root.after(500, self._tick)

    def _tick(self):
        if self.hover:
            self.deadline = max(self.deadline, time.time() + 5)
        if time.time() >= self.deadline:
            self.root.destroy()
        else:
            self.root.after(500, self._tick)

    def buttons(self, specs):
        row = self.tk.Frame(self.frame, bg=BG)
        row.pack(fill="x", pady=(8, 0))
        for text, fn in specs:
            self.tk.Button(row, text=text, command=fn, fg=FG, bg="#33333b", activebackground="#44444e",
                           activeforeground=FG, relief="flat", padx=10, pady=2,
                           font=("Segoe UI", 9)).pack(side="left", padx=(0, 6))

# ---------------------------------------------------------------- main

def log(row):
    with LOG.open("a", encoding="utf-8") as f:
        f.write(json.dumps(row, ensure_ascii=False) + "\n")


def run(image_path=None, manual=False):
    if image_path:
        from PIL import Image
        img = Image.open(image_path)
    elif manual:
        img = snip_image()
    else:
        img = screen_image()
    if img is None:
        return  # cancelled / timed out: stay silent

    card = Card()
    card.show(lambda: card.label("Reading snip…", fg=DIM))
    state = {}

    def work():
        ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        row = {"ts": ts}
        try:
            CAPTURES.mkdir(exist_ok=True)
            cap = CAPTURES / f"{ts}.png"
            img.convert("RGB").save(cap)
            row["image"] = str(cap)
            buf = io.BytesIO()
            small = img.convert("RGB")
            small.thumbnail((1568, 1568))
            small.save(buf, "PNG")
            t0 = time.time()
            out = extract_events(buf.getvalue())
            row["extracted"] = out
            row["llm_s"] = round(time.time() - t0, 1)
            svc = calendar_service() if out["events"] else None
            created = []
            for ev in out["events"]:
                res = svc.events().insert(calendarId="primary", body=to_body(ev)).execute()
                created.append({"id": res["id"], "link": res.get("htmlLink"), "ev": ev})
            row["created"] = [c["id"] for c in created]
            state.update(created=created, none_reason=out.get("none_reason"), svc=svc)
        except Exception as e:
            row["error"] = traceback.format_exc()
            state["error"] = f"{type(e).__name__}: {e}"
        state["row"] = row
        log(row)
        card.root.after(0, done)

    def undo():
        svc = state["svc"]
        for c in state["created"]:
            try:
                svc.events().delete(calendarId="primary", eventId=c["id"]).execute()
            except Exception:
                pass
        log({"ts": state["row"]["ts"], "undo": [c["id"] for c in state["created"]]})
        card.show(lambda: card.label("Removed.", fg=DIM), close_after=2)

    def done():
        if "error" in state:
            card.show(lambda: (card.label("Snip to calendar failed", fg=WARN, bold=True),
                               card.label(state["error"][:300], fg=DIM),
                               card.buttons([("Close", card.root.destroy)])), close_after=30)
            return
        created = state["created"]
        if not created:
            card.show(lambda: (card.label("No events found", bold=True),
                               card.label(state.get("none_reason") or "", fg=DIM),
                               card.buttons([("Close", card.root.destroy)])), close_after=HIDE_TIME_NO_EVENT)
            return

        def build():
            n = len(created)
            card.label(f"Added {n} event{'s' if n > 1 else ''} · {REMINDER_MIN} min reminder", fg=DIM, size=9)
            for c in created:
                ev = c["ev"]
                card.label(ev["title"], fg=WARN if ev.get("unsure") else FG, bold=True, size=11)
                card.label(when_text(ev) + (f" · {ev['location']}" if ev.get("location") else ""), fg=DIM)
                if ev.get("unsure"):
                    card.label("?? " + ev["unsure"], fg=WARN, size=9)
            card.buttons([("Undo", undo),
                          ("Open", lambda: (webbrowser.open(created[0]["link"]), card.root.destroy())),
                          ("Close", card.root.destroy)])
        card.show(build, close_after=HIDE_TIME_EVENT_ADDED)

    threading.Thread(target=work, daemon=True).start()
    card.root.mainloop()


if __name__ == "__main__":
    if "--auth" in sys.argv:
        calendar_service()
        print("signed in, token.json saved")
    else:
        try:
            args = [a for a in sys.argv[1:] if not a.startswith("--")]
            run(args[0] if args else None, manual="--snip" in sys.argv)
        except Exception:  # pythonw has no console: keep the error somewhere visible
            (ROOT / "error.log").open("a", encoding="utf-8").write(
                f"--- {dt.datetime.now():%Y-%m-%d %H:%M:%S}\n{traceback.format_exc()}")
            raise

"""Tray icon for snip-to-cal. Click = snip the screen center; right-click for the menu.

pythonw tray.py
"""
import subprocess
import sys
from pathlib import Path

import pystray
from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parent
SCRIPT = ROOT / "snip_to_cal.py"
PYTHONW = ROOT / ".venv" / "Scripts" / "pythonw.exe"


def run(*args):
    """Each snip is its own process (tkinter card needs its own main loop)."""
    subprocess.Popen([str(PYTHONW), str(SCRIPT), *args], cwd=ROOT)


def make_icon():
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((4, 8, 60, 60), 8, fill="#7aa2f7")
    d.rectangle((4, 8, 60, 22), fill="#3b5bdb")
    d.rectangle((16, 2, 22, 14), fill="#ececf1")
    d.rectangle((42, 2, 48, 14), fill="#ececf1")
    d.rectangle((18, 32, 46, 48), outline="#1f1f23", width=3)
    return img


menu = pystray.Menu(
    pystray.MenuItem("Snip screen", lambda: run(), default=True),
    pystray.MenuItem("Manual snip", lambda: run("--snip")),
    pystray.MenuItem("Quit", lambda icon: icon.stop()),
)

if __name__ == "__main__":
    pystray.Icon("snip-to-cal", make_icon(), "Snip to Calendar", menu).run()

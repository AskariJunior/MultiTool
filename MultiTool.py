#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PANDA JUNIOR - System Info Dashboard
Borderless modern UI with live CPU / Memory / Disk monitoring,
WiFi password viewer, Office → PDF converter, Windows activation,
Office 365 (Ohook) activation, and a global progress bar showing
when tasks are running.
"""

import ctypes
from ctypes import wintypes
import http.client
import math
import os
import queue
import random
import shutil
import string
import struct
import subprocess
import sys
import tempfile
import threading
import time
import urllib.request

import tkinter as tk
from tkinter import filedialog

_HAS_WIN32COM = False
try:
    import win32com.client  # noqa: F401
    _HAS_WIN32COM = True
except Exception:
    win32com = None


# --------------------------------------------------------------------------
#  Never flash a console window when we spawn helper programs.
# --------------------------------------------------------------------------
if os.name == "nt":
    _orig_popen_init = subprocess.Popen.__init__

    def _popen_init_hidden(self, *args, **kwargs):
        kwargs["creationflags"] = (kwargs.get("creationflags", 0)
                                   | 0x08000000)          # CREATE_NO_WINDOW
        if kwargs.get("startupinfo") is None:
            si = subprocess.STARTUPINFO()
            si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            si.wShowWindow = 0                             # SW_HIDE
            kwargs["startupinfo"] = si
        _orig_popen_init(self, *args, **kwargs)

    subprocess.Popen.__init__ = _popen_init_hidden


# --------------------------------------------------------------------------
#  Config / palette
# --------------------------------------------------------------------------
APP_NAME  = "PANDA JUNIOR"
ICON_PATH = r"C:\Users\PANDA\Desktop\panda.ico"

BG           = "#080b12"
PANEL_BG     = "#0d1219"
CARD_BG      = "#111823"
CARD_BORDER  = "#1d2735"
SEP          = "#161e2a"

FG           = "#e8ecf2"
MUTED        = "#7b8ba1"
SUBTLE       = "#526175"
ACCENT       = "#38bdf8"
GREEN        = "#34d399"
YELLOW       = "#fbbf24"
RED          = "#ef4444"

FONT = "Segoe UI"
MONO = "Consolas"

UI_SCALE = 1.0

OUT_FONT_SIZE   = 14
OUT_VAL_SIZE    = 14
OUT_KEY_SIZE    = 11
OUT_HDR_SIZE    = 13
OUT_WARN_SIZE   = 14

DOWN_URL = "https://speed.cloudflare.com/__down?bytes={n}&r={r}"
PING_URL = "https://speed.cloudflare.com/__down?bytes=0&r={r}"
UP_HOST  = "speed.cloudflare.com"
UP_PATH  = "/__up?r={r}"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Accept": "*/*",
}

BTN_STYLES = {
    "":     ("#1a2330", "#c7d2e0", "#243044", "#ffffff", "#2a3648"),
    "all":  ("#0284c7", "#ffffff", "#0ea5e9", "#ffffff", ACCENT),
    "spd":  ("#065f46", "#a7f3d0", "#10b981", "#ffffff", "#34d399"),
    "pdf":  ("#7c2d12", "#fed7aa", "#c2410c", "#ffffff", "#fb923c"),
    "wifi": ("#4c1d95", "#ddd6fe", "#6d28d9", "#ffffff", "#a78bfa"),
    "act":  ("#312e81", "#c7d2fe", "#4338ca", "#ffffff", "#818cf8"),
    "off":  ("#134e4a", "#99f6e4", "#14b8a6", "#ffffff", "#5eead4"),
    "clean": ("#78350f", "#fde68a", "#d97706", "#ffffff", "#fbbf24"),
    "exit": ("#7f1d1d", "#fecaca", "#b91c1c", "#ffffff", "#ef4444"),
}


# --------------------------------------------------------------------------
#  Icon resolving
# --------------------------------------------------------------------------
def _extract_exe_icon(dest):
    """Rebuild a .ico file from the icon resource embedded in our own exe."""
    if os.name != "nt" or not getattr(sys, "frozen", False):
        return None
    try:
        from ctypes import wintypes
        k32 = ctypes.windll.kernel32
        k32.LoadLibraryExW.restype = wintypes.HMODULE
        k32.LoadLibraryExW.argtypes = [wintypes.LPCWSTR, wintypes.HANDLE,
                                       wintypes.DWORD]
        k32.FindResourceW.restype = wintypes.HANDLE
        k32.FindResourceW.argtypes = [wintypes.HMODULE, ctypes.c_void_p,
                                      ctypes.c_void_p]
        k32.LoadResource.restype = wintypes.HANDLE
        k32.LoadResource.argtypes = [wintypes.HMODULE, wintypes.HANDLE]
        k32.LockResource.restype = ctypes.c_void_p
        k32.LockResource.argtypes = [wintypes.HANDLE]
        k32.SizeofResource.restype = wintypes.DWORD
        k32.SizeofResource.argtypes = [wintypes.HMODULE, wintypes.HANDLE]
        k32.FreeLibrary.argtypes = [wintypes.HMODULE]
        k32.EnumResourceNamesW.argtypes = [wintypes.HMODULE, ctypes.c_void_p,
                                           ctypes.c_void_p, ctypes.c_void_p]

        RT_ICON, RT_GROUP_ICON = 3, 14
        hmod = k32.LoadLibraryExW(sys.executable, None, 0x2)
        if not hmod:
            return None

        def get_res(rtype, name):
            h = k32.FindResourceW(hmod, name, ctypes.c_void_p(rtype))
            if not h:
                return None
            ptr = k32.LoadResource(hmod, h)
            addr = k32.LockResource(ptr)
            size = k32.SizeofResource(hmod, h)
            if not addr or not size:
                return None
            return ctypes.string_at(addr, size)

        found = []
        CB = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HMODULE,
                                ctypes.c_void_p, ctypes.c_void_p,
                                ctypes.c_void_p)

        def cb(_m, _t, name, _l):
            found.append(name)
            return False

        k32.EnumResourceNamesW(hmod, ctypes.c_void_p(RT_GROUP_ICON),
                               CB(cb), None)
        if not found:
            k32.FreeLibrary(hmod)
            return None

        grp = get_res(RT_GROUP_ICON, found[0])
        if not grp or len(grp) < 6:
            k32.FreeLibrary(hmod)
            return None
        count = struct.unpack("<H", grp[4:6])[0]
        entries, blobs = [], []
        for i in range(count):
            e = grp[6 + i * 14: 6 + (i + 1) * 14]
            if len(e) < 14:
                break
            w, h, cc, rs, planes, bits, size, rid = struct.unpack("<BBBBHHIH", e)
            data = get_res(RT_ICON, rid)
            if data:
                entries.append((w, h, cc, planes, bits, len(data)))
                blobs.append(data)
        k32.FreeLibrary(hmod)
        if not blobs:
            return None

        out = bytearray(struct.pack("<HHH", 0, 1, len(blobs)))
        offset = 6 + 16 * len(blobs)
        for (w, h, cc, planes, bits, size) in entries:
            out += struct.pack("<BBBBHHII", w, h, cc, 0, planes, bits,
                               size, offset)
            offset += size
        for b in blobs:
            out += b
        with open(dest, "wb") as f:
            f.write(out)
        return dest
    except Exception:
        return None


def _resolve_icon_path(default_path):
    names = ("panda.ico", os.path.basename(default_path))
    dirs = []
    if getattr(sys, "_MEIPASS", None):
        dirs.append(sys._MEIPASS)
    if getattr(sys, "frozen", False):
        dirs.append(os.path.dirname(sys.executable))
    dirs.append(os.path.dirname(os.path.abspath(__file__)))
    for d in dirs:
        for n in names:
            cand = os.path.join(d, n)
            if os.path.isfile(cand):
                return cand
    if os.path.isfile(default_path):
        return default_path
    dest = os.path.join(tempfile.gettempdir(), "panda_junior_icon.ico")
    got = _extract_exe_icon(dest)
    return got or default_path


ICON_PATH = _resolve_icon_path(ICON_PATH)


# --------------------------------------------------------------------------
#  Windows activation data
# --------------------------------------------------------------------------
WIN_KEYS = {
    ("10", "Home"):       "TX9XD-98N7V-6WMQ6-BX7FG-H8Q99",
    ("10", "Pro"):        "W269N-WFGWX-YVC9B-4J6C9-T83GX",
    ("10", "Education"):  "W269N-WFGWX-YVC9B-4J6C9-T83GX",
    ("10", "Enterprise"): "NPPR9-FWDCX-D2C8J-H872K-2YT43",
    ("11", "Home"):       "TX9XD-98N7V-6WMQ6-BX7FG-H8Q99",
    ("11", "Pro"):        "W269N-WFGWX-YVC9B-4J6C9-T83GX",
    ("11", "Education"):  "W269N-WFGWX-YVC9B-4J6C9-T83GX",
    ("11", "Enterprise"): "NPPR9-FWDCX-D2C8J-H872K-2YT43",
}
KMS_SERVER = "kms8.msguides.com"

# --- Office 365 (Ohook) constants -----------------------------------------
OFFICE_DLL_NAME = "sppc64.dll"                                       # plain filename
OFFICE_DLL_SRC  = os.path.expandvars(r"%TEMP%\MyAppFiles\sppc64.dll")  # preferred drop location
OFFICE_KEY         = "H8DN8-Y2YP3-CR9JT-DHDR9-C7GP3"
OHOOK_DIR          = r"C:\ohook"
REG_RESILIENCY     = (r"Software\Microsoft\Office\16.0\Common"
                      r"\Licensing\Resiliency")
REG_HEARTBEAT_VAL  = "TimeOfLastHeartbeatFailure"
REG_HEARTBEAT_DAT  = "2040-01-01T00:00:00Z"


# --------------------------------------------------------------------------
#  Helpers
# --------------------------------------------------------------------------
def bar_text(pct, width=40):
    pct = max(0, min(100, int(pct)))
    filled = int(round(pct / 100.0 * width))
    return "\u2588" * filled + "\u2591" * (width - filled)


def set_window_icon(root, ico_path):
    if os.name != "nt" or not os.path.isfile(ico_path):
        return
    try:
        u32 = ctypes.windll.user32
        hwnd = u32.GetParent(root.winfo_id()) or root.winfo_id()
        IMAGE_ICON      = 1
        LR_LOADFROMFILE = 0x0010
        WM_SETICON      = 0x0080
        ICON_SMALL      = 0
        ICON_BIG        = 1
        hs = u32.LoadImageW(0, ico_path, IMAGE_ICON, 16, 16, LR_LOADFROMFILE)
        hb = u32.LoadImageW(0, ico_path, IMAGE_ICON, 32, 32, LR_LOADFROMFILE)
        if hs:
            u32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, hs)
        if hb:
            u32.SendMessageW(hwnd, WM_SETICON, ICON_BIG, hb)
        GWL_EXSTYLE      = -20
        WS_EX_APPWINDOW  = 0x00040000
        WS_EX_TOOLWINDOW = 0x00000080
        st = u32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        st = (st | WS_EX_APPWINDOW) & ~WS_EX_TOOLWINDOW
        u32.SetWindowLongW(hwnd, GWL_EXSTYLE, st)
    except Exception:
        pass


def enable_dpi_awareness():
    if os.name != "nt":
        return 1.0
    try:
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
    except Exception:
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass
    try:
        hdc = ctypes.windll.user32.GetDC(0)
        dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)
        ctypes.windll.user32.ReleaseDC(0, hdc)
        if dpi:
            return max(1.0, dpi / 96.0)
    except Exception:
        pass
    return 1.0


def round_rect(canvas, x0, y0, x1, y1, radius=14, **kwargs):
    radius = max(0, min(radius, (x1 - x0) / 2, (y1 - y0) / 2))
    points = [
        x0 + radius, y0,
        x1 - radius, y0,
        x1, y0,
        x1, y0 + radius,
        x1, y1 - radius,
        x1, y1,
        x1 - radius, y1,
        x0 + radius, y1,
        x0, y1,
        x0, y1 - radius,
        x0, y0 + radius,
        x0, y0,
    ]
    return canvas.create_polygon(points, smooth=True, splinesteps=24, **kwargs)


class RoundedCard(tk.Frame):
    def __init__(self, parent, bg=CARD_BG, border=CARD_BORDER, radius=16,
                 outer_bg=None, **kw):
        tk.Frame.__init__(self, parent, bg=(outer_bg or parent["bg"]), **kw)
        self._bg = bg
        self._border = border
        self._radius = radius
        self.canvas = tk.Canvas(self, bg=(outer_bg or parent["bg"]),
                                 highlightthickness=0, bd=0)
        self.canvas.pack(fill="both", expand=True)
        self.body = tk.Frame(self.canvas, bg=bg)
        self._win = self.canvas.create_window(0, 0, window=self.body, anchor="nw")
        self.canvas.bind("<Configure>", self._on_resize)

    def fit_height(self):
        try:
            self.update_idletasks()
            self.canvas.config(height=self.body.winfo_reqheight() + 2)
        except Exception:
            pass

    def _on_resize(self, event):
        w, h = event.width, event.height
        self.canvas.delete("card_bg")
        if w > 3 and h > 3:
            round_rect(self.canvas, 1, 1, w - 1, h - 1, self._radius,
                       fill=self._bg, outline=self._border, width=1,
                       tags="card_bg")
            self.canvas.tag_lower("card_bg")
            self.canvas.coords(self._win, 1, 1)
            self.canvas.itemconfig(self._win, width=w - 2, height=h - 2)


class RoundedButton(tk.Canvas):
    def __init__(self, parent, text, command=None, style="", radius=10,
                 font=None, height=36, **kw):
        bg, fg, hbg, hfg, border = BTN_STYLES.get(style, BTN_STYLES[""])
        kw.setdefault("height", height)
        tk.Canvas.__init__(self, parent, bg=parent["bg"], highlightthickness=0,
                           bd=0, cursor="hand2", **kw)
        self._bg, self._fg = bg, fg
        self._hbg, self._hfg = hbg, hfg
        self._border = border
        self._text = text
        self._radius = radius
        self._font = font or (FONT, 9, "bold")
        self._command = command
        self._hover = False
        self._pressed = False
        self.bind("<Configure>", self._redraw)
        self.bind("<Button-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    def _redraw(self, event=None):
        self.delete("all")
        w = self.winfo_width()
        h = self.winfo_height()
        if w < 4 or h < 4:
            return
        bg = self._hbg if (self._hover or self._pressed) else self._bg
        fg = self._hfg if (self._hover or self._pressed) else self._fg
        round_rect(self, 1, 1, w - 1, h - 1, self._radius,
                   fill=bg, outline=self._border, width=1)
        self.create_text(w // 2, h // 2, text=self._text, fill=fg,
                         font=self._font)

    def _on_press(self, event):
        self._pressed = True
        self._redraw()

    def _on_release(self, event):
        was_pressed = self._pressed
        self._pressed = False
        self._redraw()
        if was_pressed and self._command:
            x, y = event.x, event.y
            if 0 <= x <= self.winfo_width() and 0 <= y <= self.winfo_height():
                self._command()

    def _on_enter(self, event):
        self._hover = True
        self._redraw()

    def _on_leave(self, event):
        self._hover = False
        self._pressed = False
        self._redraw()


class ProgressBar(tk.Canvas):
    def __init__(self, parent, height=22, **kw):
        tk.Canvas.__init__(self, parent, height=height, bg=BG,
                           highlightthickness=0, bd=0, **kw)
        self._height = height
        self._mode = "idle"
        self._frac = 0.0
        self._phase = 0.0
        self._label = ""
        self._job = None
        self._color = ACCENT
        self.bind("<Configure>", lambda e: self._redraw())

    def start_spin(self, label="Working...", color=ACCENT):
        self._mode = "spin"
        self._label = label
        self._color = color
        self._phase = 0.0
        self._schedule()

    def set_fraction(self, frac, label=None, color=ACCENT):
        self._mode = "determinate"
        self._frac = max(0.0, min(1.0, float(frac)))
        if label is not None:
            self._label = label
        self._color = color
        self._redraw()

    def set_label(self, label):
        self._label = label
        if self._mode == "determinate":
            self._redraw()

    def stop(self):
        self._mode = "idle"
        if self._job is not None:
            try:
                self.after_cancel(self._job)
            except Exception:
                pass
            self._job = None
        self._redraw()

    def _schedule(self):
        if self._job is not None:
            try:
                self.after_cancel(self._job)
            except Exception:
                pass
        self._job = self.after(40, self._tick)

    def _tick(self):
        self._job = None
        if self._mode != "spin":
            return
        self._phase = (self._phase + 0.04) % 1.0
        self._redraw()
        self._schedule()

    def _redraw(self):
        self.delete("all")
        w = self.winfo_width()
        h = self.winfo_height()
        if w < 10 or h < 6:
            return

        round_rect(self, 0, 0, w, h, h / 2.0,
                   fill="#0f1722", outline=CARD_BORDER, width=1)

        if self._mode == "idle":
            self.create_text(w // 2, h // 2, text="Idle",
                             fill=SUBTLE, font=(FONT, 9, "bold"))
            return

        if self._mode == "spin":
            block = 0.30
            x_start = (self._phase * (1 + block)) - block
            if x_start < 0:
                x_start = 0
            if x_start + block > 1.0:
                x_end = 1.0
            else:
                x_end = x_start + block
            x0 = int(x_start * (w - 4)) + 2
            x1 = int(x_end * (w - 4)) + 2
            if x1 - x0 > 4:
                round_rect(self, x0, 3, x1, h - 3, (h - 6) / 2.0,
                           fill=self._color, outline="")
            if self._label:
                self.create_text(w // 2, h // 2, text=self._label,
                                 fill="#ffffff", font=(FONT, 9, "bold"))
            return

        fw = int((w - 4) * self._frac)
        if fw > 3:
            round_rect(self, 2, 3, max(2 + fw, 2 + (h - 6)),
                       h - 3, (h - 6) / 2.0,
                       fill=self._color, outline="")
        if self._label:
            self.create_text(w // 2, h // 2, text=self._label,
                             fill="#ffffff", font=(FONT, 9, "bold"))


# --------------------------------------------------------------------------
#  WiFi password retrieval
# --------------------------------------------------------------------------
def _run_netsh(args):
    if os.name != "nt":
        return ""
    try:
        result = subprocess.run(
            ["netsh"] + args,
            capture_output=True,
            text=True,
            encoding="oem",
            errors="replace",
            check=False,
        )
        return result.stdout or ""
    except Exception:
        return ""


def get_wifi_profiles():
    output = _run_netsh(["wlan", "show", "profiles"])
    profiles = []
    for line in output.splitlines():
        if "All User Profile" in line and ":" in line:
            profile = line.split(":", 1)[1].strip()
            if profile:
                profiles.append(profile)
    return profiles


def get_wifi_key(profile):
    output = _run_netsh(["wlan", "show", "profile", f"name={profile}", "key=clear"])
    key = ""
    for line in output.splitlines():
        if "Key Content" in line and ":" in line:
            key = line.split(":", 1)[1].strip()
    return key


def rows_wifi(progress_cb=None):
    if os.name != "nt":
        return [("row", "WiFi", "Only available on Windows.", ("warn",))]

    profiles = get_wifi_profiles()
    if not profiles:
        return [("row", "WiFi", "No saved WiFi profiles found.", ("warn",))]

    rows = []
    total = len(profiles)
    for i, name in enumerate(profiles):
        if progress_cb:
            try:
                progress_cb(i, total, name)
            except Exception:
                pass
        key = get_wifi_key(name)
        if key:
            rows.append(("row", f"Name: [{name}]", f"Password: [{key}]", ()))
    if progress_cb:
        try:
            progress_cb(total, total, "")
        except Exception:
            pass
    if not rows:
        rows.append(("row", "WiFi", "No saved WiFi passwords found.", ("warn",)))
    return rows


# --------------------------------------------------------------------------
#  psutil optional
# --------------------------------------------------------------------------
_HAS_PSUTIL = False
try:
    import psutil as _psutil
    _HAS_PSUTIL = True
except Exception:
    _psutil = None


class _FILETIME(ctypes.Structure):
    _fields_ = [("lo", ctypes.c_uint32), ("hi", ctypes.c_uint32)]


def _ft_to_int(ft):
    return (ft.hi << 32) | ft.lo


def _win_cpu_times():
    idle   = _FILETIME()
    kernel = _FILETIME()
    user   = _FILETIME()
    ctypes.windll.kernel32.GetSystemTimes(
        ctypes.byref(idle), ctypes.byref(kernel), ctypes.byref(user))
    return _ft_to_int(idle), _ft_to_int(kernel), _ft_to_int(user)


class _MEMORYSTATUSEX(ctypes.Structure):
    _fields_ = [
        ("dwLength", ctypes.c_ulong),
        ("dwMemoryLoad", ctypes.c_ulong),
        ("ullTotalPhys", ctypes.c_ulonglong),
        ("ullAvailPhys", ctypes.c_ulonglong),
        ("ullTotalPageFile", ctypes.c_ulonglong),
        ("ullAvailPageFile", ctypes.c_ulonglong),
        ("ullTotalVirtual", ctypes.c_ulonglong),
        ("ullAvailVirtual", ctypes.c_ulonglong),
        ("ullAvailExtendedVirtual", ctypes.c_ulonglong),
    ]


def _win_memory():
    st = _MEMORYSTATUSEX()
    st.dwLength = ctypes.sizeof(_MEMORYSTATUSEX)
    ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(st))
    return int(st.dwMemoryLoad), st.ullTotalPhys, st.ullAvailPhys


def _fixed_drives():
    if os.name != "nt":
        return ["/"]
    drives = []
    try:
        mask = ctypes.windll.kernel32.GetLogicalDrives()
    except Exception:
        return []
    for i, letter in enumerate(string.ascii_uppercase):
        if mask & (1 << i):
            root = letter + ":\\"
            try:
                if ctypes.windll.kernel32.GetDriveTypeW(root) == 3:
                    drives.append(root)
            except Exception:
                pass
    return drives


def get_disk_main():
    drives = _fixed_drives()
    if not drives:
        return 0, 0, 0
    d = drives[0]
    try:
        u = shutil.disk_usage(d)
        pct = int(u.used * 100 / u.total) if u.total else 0
        return pct, u.used, u.total
    except Exception:
        return 0, 0, 0


def get_cpu_sub():
    cores = os.cpu_count() or 1
    return f"{cores} cores"


# --------------------------------------------------------------------------
#  Live monitor
# --------------------------------------------------------------------------
class Monitor(object):
    HISTORY = 60

    def __init__(self, app):
        self.app = app
        self.stop_flag = False
        self.cpu_hist = [0.0] * self.HISTORY
        self.mem_hist = [0.0] * self.HISTORY

        if _HAS_PSUTIL:
            try:
                _psutil.cpu_percent(interval=None)
            except Exception:
                pass

        threading.Thread(target=self._run, daemon=True).start()

    def stop(self):
        self.stop_flag = True

    def _cpu_percent(self):
        if _HAS_PSUTIL:
            try:
                return float(_psutil.cpu_percent(interval=None))
            except Exception:
                pass
        try:
            i1, k1, u1 = _win_cpu_times()
            time.sleep(0.15)
            i2, k2, u2 = _win_cpu_times()
            idle = i2 - i1
            total = (k2 - k1) + (u2 - u1)
            if total <= 0:
                return 0.0
            return max(0.0, min(100.0, 100.0 * (1.0 - idle / total)))
        except Exception:
            return 0.0

    def _mem_percent(self):
        if _HAS_PSUTIL:
            try:
                vm = _psutil.virtual_memory()
                return float(vm.percent), vm.used, vm.total
            except Exception:
                pass
        try:
            pct, total, avail = _win_memory()
            return float(pct), total - avail, total
        except Exception:
            return 0.0, 0, 0

    def _run(self):
        while not self.stop_flag:
            cpu = self._cpu_percent()
            mem_pct, mem_used, mem_total = self._mem_percent()
            disk_pct, disk_used, disk_total = get_disk_main()

            self.cpu_hist.append(cpu)
            self.cpu_hist.pop(0)
            self.mem_hist.append(mem_pct)
            self.mem_hist.pop(0)

            self.app.post(
                "stats",
                cpu, mem_pct, mem_used, mem_total,
                disk_pct, disk_used, disk_total,
                list(self.cpu_hist), list(self.mem_hist),
            )
            for _ in range(10):
                if self.stop_flag:
                    return
                time.sleep(0.1)


# --------------------------------------------------------------------------
#  Speed test
# --------------------------------------------------------------------------
class _UploadSource(object):
    def __init__(self, test, idx):
        self.test = test
        self.idx = idx
        self.block = b"\x00" * 65536

    def read(self, size=-1):
        if self.test.abort:
            return b""
        n = 65536 if (size is None or size < 0) else min(size, 65536)
        data = self.block[:n]
        self.test._add(self.idx, len(data))
        return data


class SpeedTest(object):
    STREAMS = 4

    def __init__(self, app):
        self.app = app
        self.lock = threading.Lock()
        self.abort = False
        self.user_stop = False
        self.running = False
        self.counters = [0] * self.STREAMS
        self.total = 0

    def _add(self, idx, n):
        with self.lock:
            self.counters[idx] += n
            self.total += n

    def _down_worker(self, idx):
        while not self.abort:
            try:
                url = DOWN_URL.format(n=25000000, r=random.random())
                req = urllib.request.Request(url, headers=HEADERS)
                with urllib.request.urlopen(req, timeout=15) as resp:
                    while not self.abort:
                        chunk = resp.read(65536)
                        if not chunk:
                            break
                        self._add(idx, len(chunk))
            except Exception:
                if self.abort:
                    break
                time.sleep(0.3)

    def _up_worker(self, idx):
        while not self.abort:
            conn = None
            try:
                conn = http.client.HTTPSConnection(UP_HOST, timeout=15)
                src = _UploadSource(self, idx)
                conn.request("POST", UP_PATH.format(r=random.random()),
                             body=src,
                             headers={"Content-Type": "application/octet-stream",
                                      "User-Agent": HEADERS["User-Agent"]})
                conn.getresponse().read()
            except Exception:
                if self.abort:
                    break
                time.sleep(0.3)
            finally:
                if conn is not None:
                    try:
                        conn.close()
                    except Exception:
                        pass

    def measure_ping(self):
        best = None
        for _ in range(4):
            if self.abort:
                break
            try:
                t0 = time.time()
                url = PING_URL.format(r=random.random())
                req = urllib.request.Request(url, headers=HEADERS)
                with urllib.request.urlopen(req, timeout=8) as r:
                    r.read()
                dt = (time.time() - t0) * 1000.0
                if best is None or dt < best:
                    best = dt
            except Exception:
                pass
        return best

    def phase(self, kind, secs, label):
        self.abort = False
        with self.lock:
            self.counters = [0] * self.STREAMS
            self.total = 0

        t0 = time.time()
        target = self._down_worker if kind == "down" else self._up_worker
        threads = []
        for i in range(self.STREAMS):
            th = threading.Thread(target=target, args=(i,), daemon=True)
            th.start()
            threads.append(th)

        hist = []
        peak = 0.0
        baseline = None
        ok, msg = True, None

        while True:
            time.sleep(0.2)
            el = time.time() - t0
            with self.lock:
                b = self.total

            hist.append((el, b))
            while len(hist) > 1 and el - hist[0][0] > 1.2:
                hist.pop(0)

            dt = el - hist[0][0]
            inst = (b - hist[0][1]) * 8.0 / dt / 1e6 if dt > 0.05 else 0.0
            if inst > peak:
                peak = inst
            if baseline is None and el >= 1.0:
                baseline = (el, b)

            self.app.post("sample", inst)
            self.app.post("status", "{0} ... {1:.1f}s{2}".format(
                label, el, " / {}s".format(secs) if secs else ""))

            if secs:
                self.app.post("progress", min(1.0, el / secs))

            if el > 4 and b == 0:
                ok = False
                msg = "No data received. Check your connection or firewall."
                break
            if self.abort or (secs and el >= secs):
                break

        self.abort = True
        for th in threads:
            th.join(timeout=2.0)

        end_el = time.time() - t0
        with self.lock:
            end_b = self.total
        if baseline is None:
            baseline = (0.0, 0)
        span = end_el - baseline[0]
        avg = (end_b - baseline[1]) * 8.0 / span / 1e6 if span > 0.2 else 0.0
        return ok, avg, peak, msg

    def start(self, mode):
        if self.running:
            return
        self.running = True
        self.user_stop = False
        self.abort = False
        threading.Thread(target=self._run, args=(mode,), daemon=True).start()

    def stop(self):
        self.user_stop = True
        self.abort = True

    def _run(self, mode):
        post = self.app.post
        try:
            post("show_speed")
            post("title", "Live Speed Monitor" if mode == "live" else "Internet Speed Test")
            post("reset_speed")
            post("status", "Measuring latency...")
            post("progress_spin", "Measuring latency...")

            ms = self.measure_ping()
            if ms is None:
                post("status", "Cannot reach the speed server. Check your connection / firewall.")
                post("progress_stop")
                return
            post("ping", ms)

            if mode == "live":
                post("progress_spin", "Live download monitor")
                _ok, avg, peak, _msg = self.phase("down", 0, "Live download")
                post("status",
                     "Monitor stopped. Average {0:.1f} Mbps, peak {1:.1f} Mbps".format(avg, peak))
                post("progress_stop")
                return

            post("progress", 0.0, "Testing download")
            ok, avg, _peak, msg = self.phase("down", 10, "Testing download")
            if not ok:
                post("status", msg or "Download test failed.")
                post("progress_stop")
                return
            post("down", avg)
            if self.user_stop:
                post("status", "Stopped. Download {0:.1f} Mbps".format(avg))
                post("progress_stop")
                return
            post("clear_graph")
            post("big", 0.0)

            post("progress", 0.0, "Testing upload")
            ok2, avg2, _peak2, msg2 = self.phase("up", 8, "Testing upload")
            if ok2 and not self.user_stop:
                post("up", avg2)
                post("big", avg2)
                post("status",
                     "Done: Download {0:.1f} Mbps / Upload {1:.1f} Mbps / Ping {2:.0f} ms"
                     .format(avg, avg2, ms))
            elif ok2:
                post("up", avg2)
                post("status",
                     "Stopped. Download {0:.1f} Mbps / Upload {1:.1f} Mbps".format(avg, avg2))
            else:
                post("status",
                     "Upload test failed (download result is still valid). {}".format(msg2))
            post("progress_stop")
        except Exception as exc:
            post("status", "Error: {}".format(exc))
            post("progress_stop")
        finally:
            self.running = False
            self.user_stop = False
            self.abort = True


# --------------------------------------------------------------------------
#  Doc -> PDF converter
# --------------------------------------------------------------------------
WORD_EXTS = (".doc", ".docx")
PPT_EXTS  = (".ppt", ".pptx")
XLS_EXTS  = (".xls", ".xlsx")
OFFICE_EXTS = WORD_EXTS + PPT_EXTS + XLS_EXTS


class EmptyWorkbook(Exception):
    """Workbook has no data / printable content on any visible sheet."""


class Doc2Pdf(object):
    def __init__(self, app):
        self.app = app
        self.running = False
        self.empty_files = []

    def convert_path(self, path):
        if self.running:
            self.app.post("doc2pdf_status", "A conversion is already running.")
            return
        self.running = True
        threading.Thread(target=self._run, args=(path,), daemon=True).start()

    def _run(self, path):
        post = self.app.post
        try:
            if os.name != "nt" or not _HAS_WIN32COM:
                post("doc2pdf_status",
                     "Conversion requires Windows with Microsoft Office and "
                     "the 'pywin32' package installed (pip install pywin32).")
                post("progress_stop")
                return
            if not os.path.exists(path):
                post("doc2pdf_status", "Path not found: {}".format(path))
                post("progress_stop")
                return
            post("doc2pdf_status", "Converting... please keep Office documents closed.")
            post("progress_spin", "Converting documents...")

            files = []
            if os.path.isdir(path):
                for root_dir, _dirs, names in os.walk(path):
                    for name in names:
                        full = os.path.join(root_dir, name)
                        if os.path.splitext(full)[1].lower() in OFFICE_EXTS:
                            files.append(full)
            else:
                if os.path.splitext(path)[1].lower() in OFFICE_EXTS:
                    files.append(path)

            total_files = len(files)
            post("progress", 0.0,
                 "0 / {}".format(total_files) if total_files else "Scanning...")

            apps = {}
            total = 0
            skipped = 0
            errors = []
            done = 0
            self.empty_files = []
            try:
                for full in files:
                    done += 1
                    ok, did_skip, err = self._convert_one(full, apps)
                    if ok:
                        total += 1
                    elif did_skip:
                        skipped += 1
                    elif err:
                        errors.append(err)
                    post("doc2pdf_progress", total, full)
                    if total_files:
                        post("progress", done / float(total_files),
                             "{} / {}".format(done, total_files))
            finally:
                for com_app in apps.values():
                    try:
                        com_app.Quit()
                    except Exception:
                        pass

            msg = "Done. {} file(s) converted to PDF.".format(total)
            if skipped:
                msg += " {} skipped (PDF already existed).".format(skipped)
            if self.empty_files:
                msg += " {} empty workbook(s) skipped.".format(
                    len(self.empty_files))
            if errors:
                msg += " {} error(s).".format(len(errors))
            post("doc2pdf_status", msg)
            if errors:
                for e in errors[:50]:
                    post("doc2pdf_status_detail", e)
        except Exception as exc:
            post("doc2pdf_status", "Error: {}".format(exc))
        finally:
            post("progress_stop")
            self.running = False

    def _convert_one(self, path, apps):
        path = os.path.normpath(os.path.abspath(path))
        ext = os.path.splitext(path)[1].lower()
        if ext not in OFFICE_EXTS:
            return False, False, None
        parent = os.path.dirname(path)
        base = os.path.splitext(os.path.basename(path))[0]
        pdf_path = os.path.join(parent, base + ".pdf")
        if os.path.exists(pdf_path):
            return False, True, None
        try:
            if ext in WORD_EXTS:
                self._convert_word(path, pdf_path, apps)
            elif ext in PPT_EXTS:
                self._convert_ppt(path, pdf_path, apps)
            elif ext in XLS_EXTS:
                self._convert_xls(path, pdf_path, apps)
            return True, False, None
        except EmptyWorkbook:
            self.empty_files.append(path)
            return False, False, None
        except Exception as exc:
            return False, False, "{}: {}".format(os.path.basename(path), exc)

    def _get_app(self, apps, prog_id, visible_attr=None):
        if prog_id not in apps:
            app = win32com.client.DispatchEx(prog_id)
            if visible_attr:
                try:
                    setattr(app, visible_attr, False)
                except Exception:
                    pass
            apps[prog_id] = app
        return apps[prog_id]

    def _convert_word(self, path, pdf_path, apps):
        word = self._get_app(apps, "Word.Application", "Visible")
        try:
            word.WordBasic.DisableAutoMacros()
        except Exception:
            pass
        doc = word.Documents.Open(path)
        try:
            doc.SaveAs(pdf_path, FileFormat=17)
        finally:
            doc.Close(False)

    def _convert_ppt(self, path, pdf_path, apps):
        ppt = self._get_app(apps, "PowerPoint.Application")
        doc = ppt.Presentations.Open(path, False, False, False)
        try:
            doc.SaveAs(pdf_path, 32)
        finally:
            doc.Close()

    @staticmethod
    def _sheet_has_content(ws):
        """True if a worksheet holds any values, shapes, or charts."""
        try:
            ur = ws.UsedRange
            if ur.Rows.Count * ur.Columns.Count > 1:
                return True
            if str(ws.Range("A1").Formula or "") != "":
                return True
        except Exception:
            return True          # can't tell -> let Excel try
        try:
            if ws.Shapes.Count > 0 or ws.ChartObjects().Count > 0:
                return True
        except Exception:
            pass
        return False

    def _convert_xls(self, path, pdf_path, apps):
        xl = self._get_app(apps, "Excel.Application", "Visible")
        try:
            xl.DisplayAlerts = False
        except Exception:
            pass
        wb = xl.Workbooks.Open(path, UpdateLinks=0, ReadOnly=True)
        try:
            content = [ws for ws in wb.Worksheets
                       if ws.Visible == -1 and self._sheet_has_content(ws)]
            if not content:
                raise EmptyWorkbook()
            try:
                wb.ExportAsFixedFormat(Type=0, Filename=pdf_path, Quality=0,
                                       IncludeDocProperties=True,
                                       IgnorePrintAreas=False,
                                       OpenAfterPublish=False)
            except Exception:
                if len(content) == 1:
                    content[0].ExportAsFixedFormat(Type=0, Filename=pdf_path,
                                                   Quality=0,
                                                   OpenAfterPublish=False)
                else:
                    raise
            wb.Saved = True
        finally:
            wb.Close(False)


# --------------------------------------------------------------------------
#  Junk cleaner
# --------------------------------------------------------------------------
_REPARSE = 0x400      # FILE_ATTRIBUTE_REPARSE_POINT (symlinks / junctions)


def _fmt_size(n):
    n = float(n)
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024.0 or unit == "GB":
            return "{:.0f} {}".format(n, unit) if unit == "B" \
                else "{:.1f} {}".format(n, unit)
        n /= 1024.0


def _is_reparse(path):
    try:
        return bool(os.lstat(path).st_file_attributes & _REPARSE)
    except Exception:
        return os.path.islink(path)


def _dir_size(path):
    total = 0
    try:
        for entry in os.scandir(path):
            try:
                if entry.is_dir(follow_symlinks=False) \
                        and not _is_reparse(entry.path):
                    total += _dir_size(entry.path)
                elif not entry.is_dir(follow_symlinks=False):
                    total += entry.stat(follow_symlinks=False).st_size
            except Exception:
                pass
    except Exception:
        pass
    return total


def _lp(path):
    """Extended-length path prefix so paths over 260 chars still work."""
    if os.name != "nt":
        return path
    path = os.path.abspath(path)
    if path.startswith("\\\\?\\"):
        return path
    if path.startswith("\\\\"):
        return "\\\\?\\UNC\\" + path[2:]
    return "\\\\?\\" + path


def _plain(path):
    return path[4:] if path.startswith("\\\\?\\") else path


def _fix_acl(path):
    """Admin only: take ownership + grant Administrators full control."""
    p = _plain(path)
    try:
        subprocess.run(["takeown", "/f", p, "/d", "y"], timeout=20,
                       stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                       stderr=subprocess.DEVNULL)
        subprocess.run(["icacls", p, "/grant", "*S-1-5-32-544:F", "/c", "/q"],
                       timeout=20, stdin=subprocess.DEVNULL,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass


def _new_stats():
    return {"freed": 0, "removed": 0, "in_use": 0, "denied": 0,
            "other": 0, "acl_fixes": 0}


def _fails(st):
    return st["in_use"] + st["denied"] + st["other"]


def _remove(func, p, st, fix_acl, size=0, is_file=True):
    """Try to delete one item (with one retry after clearing read-only /
    fixing permissions). Records why it failed."""
    for attempt in (0, 1):
        try:
            func(p)
            if is_file:
                st["freed"] += size
                st["removed"] += 1
            return True
        except Exception as exc:
            code = getattr(exc, "winerror", None)
            if attempt == 0:
                try:
                    os.chmod(p, 0o777)          # clear read-only flag
                except Exception:
                    pass
                if code == 5 and fix_acl and st["acl_fixes"] < 400:
                    st["acl_fixes"] += 1
                    _fix_acl(p)
                continue
            if code in (32, 33):
                st["in_use"] += 1
            elif code == 5:
                st["denied"] += 1
            else:
                st["other"] += 1
            return False
    return False


def _clean_dir(path, st, fix_acl=False):
    """Delete everything INSIDE `path` (never the folder itself).
    Junctions / symlinks are unlinked, never followed."""
    try:
        entries = list(os.scandir(path))
    except Exception as exc:
        code = getattr(exc, "winerror", None)
        st["denied" if code == 5 else "other"] += 1
        return
    for entry in entries:
        p = entry.path
        try:
            is_dir = entry.is_dir(follow_symlinks=False)
        except Exception:
            st["other"] += 1
            continue
        if is_dir and _is_reparse(p):
            _remove(os.rmdir, p, st, fix_acl, is_file=False)
        elif is_dir:
            before = _fails(st)
            _clean_dir(p, st, fix_acl)
            if _fails(st) == before:
                _remove(os.rmdir, p, st, fix_acl, is_file=False)
        else:
            try:
                size = entry.stat(follow_symlinks=False).st_size
            except Exception:
                size = 0
            _remove(os.remove, p, st, fix_acl, size=size)


def _clean_target(path, fix_acl=False):
    """Clean one folder; retry once for anything that failed."""
    st = _new_stats()
    root = _lp(path)
    _clean_dir(root, st, fix_acl)
    if _fails(st):
        time.sleep(0.5)
        for k in ("in_use", "denied", "other"):
            st[k] = 0
        _clean_dir(root, st, fix_acl)
    return st


def _recycle_bin_info():
    """(size_bytes, item_count) of the Recycle Bin, or (0, 0)."""
    try:
        class SHQUERYRBINFO(ctypes.Structure):
            _fields_ = [("cbSize", wintypes.DWORD),
                        ("i64Size", ctypes.c_longlong),
                        ("i64NumItems", ctypes.c_longlong)]
        info = SHQUERYRBINFO()
        info.cbSize = ctypes.sizeof(info)
        if ctypes.windll.shell32.SHQueryRecycleBinW(None, ctypes.byref(info)) == 0:
            return int(info.i64Size), int(info.i64NumItems)
    except Exception:
        pass
    return 0, 0


def _empty_recycle_bin():
    try:
        # NOCONFIRMATION | NOPROGRESSUI | NOSOUND
        ctypes.windll.shell32.SHEmptyRecycleBinW(None, None, 0x7)
    except Exception:
        pass


def junk_targets():
    """[(label, path, needs_admin), ...] - only folders that exist."""
    env = os.environ
    local = env.get("LOCALAPPDATA", "")
    sysroot = env.get("SystemRoot", r"C:\Windows")
    pdata = env.get("ProgramData", r"C:\ProgramData")
    user_temp = tempfile.gettempdir()

    cand = [
        ("User Temp (%TEMP%)",       user_temp,                                         False),
        ("Local AppData Temp",       os.path.join(local, "Temp"),                       False),
        ("Crash Dumps",              os.path.join(local, "CrashDumps"),                 False),
        ("Error Reports (user)",     os.path.join(local, "Microsoft", "Windows", "WER"), False),
        ("Internet Cache",           os.path.join(local, "Microsoft", "Windows", "INetCache"), False),
        ("Chrome Cache",             os.path.join(local, "Google", "Chrome", "User Data", "Default", "Cache"), False),
        ("Chrome Code Cache",        os.path.join(local, "Google", "Chrome", "User Data", "Default", "Code Cache"), False),
        ("Edge Cache",               os.path.join(local, "Microsoft", "Edge", "User Data", "Default", "Cache"), False),
        ("Edge Code Cache",          os.path.join(local, "Microsoft", "Edge", "User Data", "Default", "Code Cache"), False),
        ("Windows Temp",             os.path.join(sysroot, "Temp"),                     True),
        ("Prefetch",                 os.path.join(sysroot, "Prefetch"),                 True),
        ("Windows Update Cache",     os.path.join(sysroot, "SoftwareDistribution", "Download"), True),
        ("Minidumps",                os.path.join(sysroot, "Minidump"),                 True),
        ("Error Reports (system)",   os.path.join(pdata, "Microsoft", "Windows", "WER", "ReportArchive"), True),
        ("Error Report Queue",       os.path.join(pdata, "Microsoft", "Windows", "WER", "ReportQueue"), True),
        ("Delivery Optimization",    os.path.join(sysroot, "ServiceProfiles", "NetworkService", "AppData",
                                                  "Local", "Microsoft", "Windows",
                                                  "DeliveryOptimization", "Cache"), True),
    ]
    seen, out = set(), []
    for label, path, adm in cand:
        if not path or not os.path.isdir(path):
            continue
        key = os.path.normcase(os.path.realpath(path))
        if key in seen:
            continue
        seen.add(key)
        out.append((label, path, adm))
    return out


class CleanDialog(tk.Toplevel):
    """Dark confirmation dialog (same look as ActivationDialog)."""

    def __init__(self, parent, on_confirm):
        tk.Toplevel.__init__(self, parent)
        self._parent = parent
        self._on_confirm = on_confirm
        self.configure(bg=BG)
        self.transient(parent)
        self.overrideredirect(True)

        parent.update_idletasks()
        w = int(470 * UI_SCALE)
        h = int(470 * UI_SCALE)
        x = parent.winfo_x() + (parent.winfo_width() - w) // 2
        y = parent.winfo_y() + (parent.winfo_height() - h) // 2
        self.geometry("{}x{}+{}+{}".format(w, h, x, y))

        outer = tk.Frame(self, bg=BG,
                         highlightbackground=CARD_BORDER, highlightthickness=1)
        outer.pack(fill="both", expand=True)
        wrap = tk.Frame(outer, bg=BG)
        wrap.pack(fill="both", expand=True, padx=24, pady=18)

        tk.Label(wrap, text="Clean Junk Files", bg=BG, fg=FG,
                 font=(FONT, 15, "bold")).pack(anchor="w")
        tk.Label(wrap, text="This permanently deletes the following. "
                            "Files in use are skipped.",
                 bg=BG, fg=SUBTLE, font=(FONT, 9), wraplength=int(410 * UI_SCALE),
                 justify="left").pack(anchor="w", pady=(2, 10))

        items = [
            "User Temp and %TEMP% folders",
            "Windows\\Temp",
            "Windows\\Prefetch (rebuilds itself; apps may open a little "
            "slower the first time)",
            "Windows Update download cache, Delivery Optimization cache",
            "Crash dumps, minidumps, Windows error reports",
            "Internet cache, Chrome and Edge cache",
            "Recycle Bin",
        ]
        for it in items:
            tk.Label(wrap, text="\u2022  " + it, bg=BG, fg="#c7d2e0",
                     font=(FONT, 10), anchor="w", justify="left",
                     wraplength=int(410 * UI_SCALE)).pack(fill="x", pady=1)

        tk.Label(wrap, text="Saved passwords, cookies and history are not touched.",
                 bg=BG, fg=MUTED, font=(FONT, 9), anchor="w").pack(
                     fill="x", pady=(10, 0))

        RoundedButton(wrap, "Clean Now", self._ok, style="clean", radius=10,
                      height=int(38 * UI_SCALE)).pack(fill="x", pady=(16, 4))
        RoundedButton(wrap, "Cancel", self._cancel, style="", radius=10,
                      height=int(34 * UI_SCALE)).pack(fill="x")

        self.bind("<Escape>", lambda e: self._cancel())
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self._dx = self._dy = 0
        for widget in (outer, wrap):
            widget.bind("<Button-1>", self._drag_start)
            widget.bind("<B1-Motion>", self._drag_move)
        self.after(50, self._focus)

    def _focus(self):
        try:
            self.lift()
            self.attributes("-topmost", True)
            self.focus_force()
            self.grab_set()
        except Exception:
            pass

    def _drag_start(self, event):
        self._dx = event.x_root - self.winfo_x()
        self._dy = event.y_root - self.winfo_y()

    def _drag_move(self, event):
        self.geometry("+{}+{}".format(event.x_root - self._dx,
                                      event.y_root - self._dy))

    def _close(self):
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()

    def _ok(self):
        cb, parent = self._on_confirm, self._parent
        self._close()
        parent.after(80, cb)

    def _cancel(self):
        self._close()


# --------------------------------------------------------------------------
#  Windows activation dialog
# --------------------------------------------------------------------------
class ActivationDialog(tk.Toplevel):
    EDITIONS = ("Home", "Pro", "Education", "Enterprise")

    def __init__(self, parent, win_version, on_pick=None):
        tk.Toplevel.__init__(self, parent)
        self.result = None
        self._parent = parent
        self._on_pick = on_pick
        self._win_version = win_version
        self.configure(bg=BG)
        self.transient(parent)
        self.overrideredirect(True)

        parent.update_idletasks()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_x()
        py = parent.winfo_y()
        w = int(440 * UI_SCALE)
        h = int(340 * UI_SCALE)
        x = px + (pw - w) // 2
        y = py + (ph - h) // 2
        self.geometry("{}x{}+{}+{}".format(w, h, x, y))

        outer = tk.Frame(self, bg=BG,
                         highlightbackground=CARD_BORDER, highlightthickness=1)
        outer.pack(fill="both", expand=True)

        wrap = tk.Frame(outer, bg=BG)
        wrap.pack(fill="both", expand=True, padx=24, pady=18)

        header = tk.Frame(wrap, bg=BG)
        header.pack(fill="x")
        tk.Label(header,
                 text="Activate Windows {}".format(win_version),
                 bg=BG, fg=FG, font=(FONT, 15, "bold")).pack(anchor="w")
        tk.Label(header,
                 text="Choose the edition installed on this PC",
                 bg=BG, fg=SUBTLE, font=(FONT, 9)).pack(anchor="w", pady=(2, 0))

        close = tk.Canvas(outer, width=28, height=28, bg=BG,
                          highlightthickness=0, bd=0, cursor="hand2")
        close.place(relx=1.0, x=-8, y=8, anchor="ne")
        close.create_text(14, 14, text="\u2715", fill="#5b6a80",
                          font=(FONT, 10), tags="x")
        close.bind("<Button-1>", lambda e: self._cancel())
        close.bind("<Enter>", lambda e: close.itemconfig("x", fill=RED))
        close.bind("<Leave>", lambda e: close.itemconfig("x", fill="#5b6a80"))

        btns = tk.Frame(wrap, bg=BG)
        btns.pack(fill="x", pady=(14, 0))
        for edition in self.EDITIONS:
            RoundedButton(
                btns,
                "Windows {} {}".format(win_version, edition),
                command=lambda e=edition: self._pick(e),
                style="act", radius=10, height=int(38 * UI_SCALE),
            ).pack(fill="x", pady=3)

        RoundedButton(
            wrap, "Cancel", self._cancel, style="", radius=10,
            height=int(34 * UI_SCALE),
        ).pack(fill="x", pady=(10, 0))

        self.bind("<Escape>", lambda e: self._cancel())
        self.protocol("WM_DELETE_WINDOW", self._cancel)

        self._dx = 0
        self._dy = 0
        for widget in (outer, wrap, header):
            widget.bind("<Button-1>", self._drag_start)
            widget.bind("<B1-Motion>", self._drag_move)

        self.after(50, self._focus)

    def _focus(self):
        try:
            self.lift()
            self.attributes("-topmost", True)
        except Exception:
            pass
        try:
            self.focus_force()
            self.grab_set()
        except Exception:
            pass

    def _drag_start(self, event):
        self._dx = event.x_root - self.winfo_x()
        self._dy = event.y_root - self.winfo_y()

    def _drag_move(self, event):
        self.geometry("+{}+{}".format(event.x_root - self._dx,
                                      event.y_root - self._dy))

    def _pick(self, edition):
        self.result = edition
        cb, parent = self._on_pick, self._parent
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()
        if cb:
            parent.after(80, lambda: cb(edition))

    def _cancel(self):
        self.result = None
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()


# --------------------------------------------------------------------------
#  Application
# --------------------------------------------------------------------------
class App(object):
    def __init__(self, root):
        self.root = root
        self.q = queue.Queue()
        self.samples = []
        self.mode = "out"
        self.cpu_hist = [0.0] * Monitor.HISTORY
        self.mem_hist = [0.0] * Monitor.HISTORY
        self.disk_pct = 0.0
        self.st = SpeedTest(self)
        self.doc2pdf = Doc2Pdf(self)
        self.monitor = None
        self._last_bar_pct = 0.0
        self._activating = False
        self._activating_office = False
        self._cleaning = False

        self._build()
        self.root.after(60, self._poll)
        self.root.after(200, self._start_monitor)

    def _start_monitor(self):
        self.monitor = Monitor(self)

    def post(self, kind, *args):
        self.q.put((kind, args))

    def _poll(self):
        try:
            while True:
                kind, args = self.q.get_nowait()
                try:
                    self._handle(kind, args)
                except Exception:
                    pass
        except queue.Empty:
            pass
        try:
            self.root.after(60, self._poll)
        except tk.TclError:
            pass

    # ======================================================================
    #  UI
    # ======================================================================
    def _build(self):
        r = self.root
        r.configure(bg=BG)
        r.overrideredirect(True)

        sw, sh = r.winfo_screenwidth(), r.winfo_screenheight()
        base_w = int(1040 * UI_SCALE)
        base_h = int(1180 * UI_SCALE)
        w = min(base_w, sw - 40)
        h = min(base_h, sh - 40)
        x, y = (sw - w) // 2, max(10, (sh - h) // 2)
        r.geometry("{}x{}+{}+{}".format(w, h, x, y))
        r.minsize(int(760 * UI_SCALE), int(680 * UI_SCALE))

        try:
            if os.path.isfile(ICON_PATH):
                r.iconbitmap(ICON_PATH)
        except Exception:
            pass

        outer = tk.Frame(r, bg=BG, highlightbackground=CARD_BORDER, highlightthickness=1)
        outer.pack(fill="both", expand=True)

        wrap = tk.Frame(outer, bg=BG)
        wrap.pack(fill="both", expand=True, padx=20, pady=12)

        self._build_header(wrap, outer)
        self._build_stats(wrap)
        self._build_actions(wrap)
        self._build_progress(wrap)
        self._build_panel(wrap)

        self.root.bind("<Escape>", lambda e: self.root.destroy())

        self.root.update_idletasks()
        if os.path.isfile(ICON_PATH):
            try:
                self.root.iconbitmap(ICON_PATH)
            except Exception:
                pass
            set_window_icon(self.root, ICON_PATH)

        try:
            self.root.focus_force()
        except Exception:
            pass

    # ---------------- header ----------------
    def _build_header(self, wrap, outer):
        header = tk.Frame(wrap, bg=BG)
        header.pack(fill="x")

        logo_img = self._load_logo(target=44)
        if logo_img is not None:
            lbl = tk.Label(header, image=logo_img, bg=BG, bd=0)
            lbl.image = logo_img
            lbl.pack(side="left", padx=(0, 14))
        else:
            logo_box = tk.Frame(header, bg="#0284c7", width=44, height=44)
            logo_box.pack(side="left", padx=(0, 14))
            logo_box.pack_propagate(False)
            tk.Label(logo_box, text="P", bg="#0284c7", fg="#ffffff",
                     font=(FONT, 20, "bold")).pack(expand=True)

        txt = tk.Frame(header, bg=BG)
        txt.pack(side="left", fill="x", expand=True)
        tk.Label(txt, text=APP_NAME, bg=BG, fg=FG,
                 font=(FONT, 20, "bold")).pack(anchor="w")
        tk.Label(txt, text="System Info Dashboard  \u00b7  live performance & network tools",
                 bg=BG, fg=SUBTLE, font=(FONT, 9)).pack(anchor="w", pady=(1, 0))

        close = tk.Canvas(outer, width=34, height=34, bg=BG,
                          highlightthickness=0, bd=0, cursor="hand2")
        close.place(relx=1.0, x=-14, y=12, anchor="ne")
        close.create_oval(4, 4, 30, 30, fill=BG, outline=BG, tags="bg")
        close.create_text(17, 17, text="\u2715", fill="#5b6a80",
                          font=(FONT, 11), tags="x")
        close.bind("<Button-1>", lambda e: self.root.destroy())
        close.bind("<Enter>", lambda e: (close.itemconfig("bg", fill="#7f1d1d"),
                                         close.itemconfig("x", fill="#ffffff")))
        close.bind("<Leave>", lambda e: (close.itemconfig("bg", fill=BG),
                                         close.itemconfig("x", fill="#5b6a80")))

        self._bind_drag(header)

    # ---------------- stat cards ----------------
    def _build_stats(self, wrap):
        row = tk.Frame(wrap, bg=BG)
        row.pack(fill="x", pady=(2, 0))
        for c in range(3):
            row.columnconfigure(c, weight=1, uniform="stat")

        self.cpu_val, self.cpu_sub, self.cpu_graph = self._mk_stat_card(
            row, 0, "CPU USAGE", ACCENT, "cpu")
        self.mem_val, self.mem_sub, self.mem_graph = self._mk_stat_card(
            row, 1, "MEMORY", GREEN, "mem")
        self.dsk_val, self.dsk_sub, self.dsk_graph = self._mk_stat_card(
            row, 2, "STORAGE", YELLOW, "disk")

    def _mk_stat_card(self, parent, col, title, color, kind):
        card = RoundedCard(parent, bg=CARD_BG, border=CARD_BORDER, radius=14,
                           outer_bg=BG)
        card.grid(row=0, column=col, sticky="nsew", padx=5)
        body = card.body

        line = tk.Frame(body, bg=CARD_BG)
        line.pack(fill="x", padx=10, pady=(4, 4))

        tk.Label(line, text=title, bg=CARD_BG, fg=MUTED,
                 font=(FONT, 8, "bold")).pack(side="left")

        val = tk.Label(line, text="--", bg=CARD_BG, fg=color,
                       font=(FONT, 12, "bold"), width=4, anchor="e")
        val.pack(side="left", padx=(6, 0))

        graph = tk.Canvas(line, width=1, height=int(18 * UI_SCALE),
                          bg=CARD_BG, highlightthickness=0, bd=0)
        graph.pack(side="left", fill="x", expand=True, padx=(6, 6))

        sub = tk.Label(line, text="", bg=CARD_BG, fg=SUBTLE, font=(FONT, 8))
        sub.pack(side="right")
        graph._color = color
        graph._kind = kind
        graph.bind("<Configure>", lambda e, c=graph: self._redraw_one(c))

        card.canvas.config(height=1)
        card.fit_height()
        card.after(150, card.fit_height)
        return val, sub, graph

    def _redraw_one(self, canvas):
        kind = getattr(canvas, "_kind", None)
        color = getattr(canvas, "_color", ACCENT)
        if kind == "cpu":
            self._draw_spark(canvas, self.cpu_hist, color)
        elif kind == "mem":
            self._draw_spark(canvas, self.mem_hist, color)
        elif kind == "disk":
            self._draw_bar(canvas, self.disk_pct, color)

    def _draw_spark(self, canvas, values, color):
        canvas.delete("all")
        try:
            w = canvas.winfo_width()
            h = canvas.winfo_height()
        except Exception:
            return
        if w < 12 or h < 8 or len(values) < 2:
            return
        m = max(max(values), 1.0)
        n = len(values)
        step = (w - 2) / (n - 1)

        line = []
        for i, v in enumerate(values):
            x = 1 + i * step
            y = h - 2 - (v / m) * (h - 6)
            line.extend([x, y])

        pts = [1, h - 1] + line + [w - 1, h - 1]
        canvas.create_polygon(pts, fill=color, outline="", stipple="gray75",
                              smooth=True, splinesteps=18)
        canvas.create_line(*line, fill=color, width=2, smooth=True,
                           splinesteps=18, capstyle=tk.ROUND,
                           joinstyle=tk.ROUND)

    def _draw_bar(self, canvas, pct, color):
        canvas.delete("all")
        try:
            w = canvas.winfo_width()
            h = canvas.winfo_height()
        except Exception:
            return
        if w < 12 or h < 8:
            return
        y0 = h // 2 - 4
        y1 = h // 2 + 4
        radius = (y1 - y0) / 2.0
        round_rect(canvas, 2, y0, w - 2, y1, radius,
                  fill="#1a2533", outline="")
        fw = int((w - 4) * max(0, min(100, pct)) / 100.0)
        if fw > 0:
            fill_x1 = max(2 + fw, 2 + radius * 2)
            round_rect(canvas, 2, y0, min(fill_x1, w - 2), y1, radius,
                      fill=color, outline="")

    # ---------------- action buttons ----------------
    def _build_actions(self, wrap):
        """Action buttons laid out A-Z in reading order (left to right,
        top to bottom), three per row."""
        grid = tk.Frame(wrap, bg=BG)
        grid.pack(fill="x", pady=(2, 4))
        for c in range(3):
            grid.columnconfigure(c, weight=1, uniform="btn")

        buttons = [
            ("Activate Windows 10",
             lambda: self.activate_windows("10"), "act"),
            ("Activate Windows 11",
             lambda: self.activate_windows("11"), "act"),
            ("Check Activation Status",
             self.show_activation_status, "act"),
            ("Clean Junk Files",
             self.clean_junk, "clean"),
            ("Convert File \u2192 PDF",
             self.doc2pdf_pick_file, "pdf"),
            ("Convert Folder \u2192 PDF",
             self.doc2pdf_pick_folder, "pdf"),
            ("Copy Output",
             self.copy_output, ""),
            ("Exit",
             self.root.destroy, "exit"),
            ("Live Speed Monitor",
             lambda: self.st.start("live"), "spd"),
            ("Office 365",
             self.activate_office, "off"),
            ("Speed Test (Down/Up/Ping)",
             lambda: self.st.start("full"), "spd"),
            ("WiFi Passwords",
             self.show_wifi, "wifi"),
        ]
        buttons.sort(key=lambda b: b[0].lower())
        for i, (text, cmd, style) in enumerate(buttons):
            self._mkbtn(grid, text, cmd, style, i // 3, i % 3)

    def _mkbtn(self, parent, text, command, style, row, col):
        b = RoundedButton(parent, text, command, style, radius=9,
                          font=(FONT, 9, "bold"), height=int(34 * UI_SCALE))
        b.grid(row=row, column=col, sticky="ew", padx=4, pady=2)
        return b

    # ---------------- progress bar ----------------
    def _build_progress(self, wrap):
        holder = tk.Frame(wrap, bg=BG)
        holder.pack(fill="x", pady=(2, 6))
        self.progress = ProgressBar(holder, height=int(22 * UI_SCALE))
        self.progress.pack(fill="x", padx=4)

    # ---------------- panel ----------------
    def _build_panel(self, wrap):
        self.title_lbl = tk.Label(wrap, text="WELCOME", bg=BG, fg=ACCENT,
                                  font=(FONT, 11, "bold"), anchor="w")
        self.title_lbl.pack(fill="x", padx=4, pady=(6, 4))

        area = tk.Frame(wrap, bg=BG)
        area.pack(fill="both", expand=True)
        self.panel_area = area

        self._build_out_panel(area)
        self._build_speed_panel(area)
        self._show_out()

    def _build_out_panel(self, parent):
        card = RoundedCard(parent, bg=PANEL_BG, border=CARD_BORDER, radius=16,
                           outer_bg=BG)
        inner = card.body

        sb = tk.Scrollbar(inner, orient="vertical", width=12,
                          bg="#16202f", troughcolor=PANEL_BG, bd=0,
                          activebackground="#26364f", relief="flat")
        sb.pack(side="right", fill="y")

        t = tk.Text(inner, height=4, bg=PANEL_BG, fg=FG, bd=0, highlightthickness=0,
                    wrap="word", padx=22, pady=16, font=(MONO, OUT_FONT_SIZE),
                    insertwidth=0, cursor="arrow", tabs=(220,),
                    yscrollcommand=sb.set, selectbackground="#1e3a5f",
                    spacing1=2, spacing3=2)
        t.pack(side="left", fill="both", expand=True)
        sb.config(command=t.yview)

        t.tag_configure("k", foreground=MUTED, font=(FONT, OUT_KEY_SIZE, "bold"),
                        spacing1=12, spacing3=8)
        t.tag_configure("v", foreground="#f8fafc",
                        font=(MONO, OUT_VAL_SIZE, "bold"),
                        spacing1=12, spacing3=8)
        t.tag_configure("warn", foreground=RED,
                        font=(FONT, OUT_WARN_SIZE, "bold"))
        t.tag_configure("ok", foreground=GREEN,
                        font=(FONT, OUT_WARN_SIZE, "bold"))
        t.tag_configure("bar", foreground=ACCENT,
                        font=(MONO, OUT_VAL_SIZE))
        t.tag_configure("hdr", foreground=ACCENT,
                        font=(FONT, OUT_HDR_SIZE, "bold"),
                        spacing1=18, spacing3=8)
        t.tag_configure("sep", foreground=SEP, font=(MONO, 7))

        t.config(state="disabled")
        self.out_frame = card
        self.txt = t

    def _build_speed_panel(self, parent):
        card = RoundedCard(parent, bg=PANEL_BG, border=CARD_BORDER, radius=16,
                           outer_bg=BG)
        inner = card.body

        hero = tk.Frame(inner, bg=PANEL_BG)
        hero.pack(fill="x", pady=(16, 0))

        bigrow = tk.Frame(hero, bg=PANEL_BG)
        bigrow.pack()
        self.big_lbl = tk.Label(bigrow, text="0.0", bg=PANEL_BG, fg=GREEN,
                                font=(FONT, 46, "bold"))
        self.big_lbl.pack(side="left")
        tk.Label(bigrow, text="Mbps", bg=PANEL_BG, fg=SUBTLE,
                 font=(FONT, 12, "bold")).pack(side="left", padx=(10, 0), pady=(30, 0))

        self.status_lbl = tk.Label(hero, text="Ready", bg=PANEL_BG,
                                   fg=YELLOW, font=(FONT, 10))
        self.status_lbl.pack(pady=(0, 6))

        barwrap = tk.Frame(inner, bg=PANEL_BG)
        barwrap.pack(fill="x", padx=20, pady=(4, 10))
        self.bar = tk.Canvas(barwrap, height=8, bg=PANEL_BG,
                             highlightthickness=0, bd=0)
        self.bar.pack(fill="x")
        self.bar_fill = None
        self.bar.bind("<Configure>", lambda e: self._draw_speed_bar_track())

        self.graph = tk.Canvas(inner, width=1, height=112, bg=BG,
                               highlightthickness=0, bd=0)
        self.graph.pack(fill="x", expand=True, padx=20, pady=(2, 8))
        self.graph.bind("<Configure>", lambda e: self._draw_speed_graph())

        boxrow = tk.Frame(inner, bg=PANEL_BG)
        boxrow.pack(pady=(4, 14))
        self.box_ping = self._mkbox(boxrow, "Ping",     YELLOW)
        self.box_down = self._mkbox(boxrow, "Download", ACCENT)
        self.box_up   = self._mkbox(boxrow, "Upload",   GREEN)

        self.speed_frame = card

    def _mkbox(self, parent, label, color):
        card = RoundedCard(parent, bg="#0f1723", border=CARD_BORDER, radius=12,
                           outer_bg=PANEL_BG)
        card.pack(side="left", padx=6)
        f = card.body
        tk.Label(f, text=label.upper(), bg="#0f1723", fg=MUTED,
                 font=(FONT, 8, "bold")).pack(anchor="w", padx=16, pady=(11, 0))
        v = tk.Label(f, text="--", bg="#0f1723", fg=color,
                     font=(MONO, 17, "bold"), width=11, anchor="w")
        v.pack(anchor="w", padx=16, pady=(0, 11))
        return v

    # ---------------- logo loader ----------------
    def _load_logo(self, target=44):
        if not os.path.isfile(ICON_PATH):
            return None
        try:
            from PIL import Image, ImageTk
            img = Image.open(ICON_PATH)
            try:
                img.size = (target, target)
            except Exception:
                pass
            img = img.convert("RGBA").resize((target, target), Image.LANCZOS)
            return ImageTk.PhotoImage(img)
        except Exception:
            pass
        try:
            return self._load_ico_manual(target)
        except Exception:
            return None

    def _load_ico_manual(self, target):
        with open(ICON_PATH, "rb") as f:
            data = f.read()
        if len(data) < 6 or data[:4] != b"\x00\x00\x01\x00":
            return None
        count = struct.unpack("<H", data[4:6])[0]
        entries = []
        for i in range(count):
            off = 6 + i * 16
            if off + 16 > len(data):
                break
            w  = data[off] or 256
            h  = data[off + 1] or 256
            sz = struct.unpack("<I", data[off + 8:off + 12])[0]
            of = struct.unpack("<I", data[off + 12:off + 16])[0]
            entries.append((w, h, sz, of))
        if not entries:
            return None
        pngs = [(w, h, data[of:of + sz]) for w, h, sz, of in entries
                if data[of:of + 8] == b"\x89PNG\r\n\x1a\n"]
        if pngs:
            pngs.sort(key=lambda t: abs(t[0] - target))
            img = self._image_from_bytes(pngs[0][2], ".png")
            if img is not None:
                while img.width() > target * 1.5 and img.width() >= 4:
                    img = img.subsample(2, 2)
                return img
        entries.sort(key=lambda t: abs(t[0] - target))
        for _w, _h, sz, of in entries:
            img = self._bmp_frame_to_image(data[of:of + sz])
            if img is not None:
                return img
        return None

    def _image_from_bytes(self, blob, suffix):
        fd, path = tempfile.mkstemp(suffix=suffix)
        try:
            with os.fdopen(fd, "wb") as f:
                f.write(blob)
            return tk.PhotoImage(file=path)
        except Exception:
            return None
        finally:
            try:
                os.remove(path)
            except Exception:
                pass

    def _bmp_frame_to_image(self, bmp):
        if len(bmp) < 40:
            return None
        bi_size     = struct.unpack("<I", bmp[0:4])[0]
        if bi_size < 40:
            return None
        width       = struct.unpack("<i", bmp[4:8])[0]
        height2     = struct.unpack("<i", bmp[8:12])[0]
        bitcount    = struct.unpack("<H", bmp[14:16])[0]
        compression = struct.unpack("<I", bmp[16:20])[0]
        if width <= 0 or height2 <= 0 or compression != 0 or bitcount not in (24, 32):
            return None
        height = height2 // 2
        bpp    = bitcount // 8
        row_raw = width * bpp
        row_padded = (row_raw + 3) & ~3
        pixel_offset = bi_size
        bg_r = int(BG[1:3], 16)
        bg_g = int(BG[3:5], 16)
        bg_b = int(BG[5:7], 16)
        rgb = bytearray(width * height * 3)
        for y in range(height):
            src_y = height - 1 - y
            start = pixel_offset + src_y * row_padded
            row = bmp[start:start + row_raw]
            if len(row) < row_raw:
                return None
            for x in range(width):
                b = row[x * bpp]
                g = row[x * bpp + 1]
                r = row[x * bpp + 2]
                a = row[x * bpp + 3] if bpp == 4 else 255
                if a < 255:
                    r = (r * a + bg_r * (255 - a)) // 255
                    g = (g * a + bg_g * (255 - a)) // 255
                    b = (b * a + bg_b * (255 - a)) // 255
                i = (y * width + x) * 3
                rgb[i] = r
                rgb[i + 1] = g
                rgb[i + 2] = b
        ppm = b"P6\n{} {}\n255\n".format(width, height) + bytes(rgb)
        return self._image_from_bytes(ppm, ".ppm")

    # ---------------- dragging ----------------
    def _bind_drag(self, widget):
        widget.bind("<Button-1>", self._drag_start)
        widget.bind("<B1-Motion>", self._drag_move)
        for child in widget.winfo_children():
            self._bind_drag(child)

    def _drag_start(self, event):
        self._dx = event.x_root - self.root.winfo_x()
        self._dy = event.y_root - self.root.winfo_y()

    def _drag_move(self, event):
        self.root.geometry("+{}+{}".format(event.x_root - self._dx,
                                           event.y_root - self._dy))

    # ---------------- panels ----------------
    def _show_out(self):
        self.mode = "out"
        self.speed_frame.pack_forget()
        self.out_frame.pack(fill="both", expand=True)

    def _show_speed(self):
        self.mode = "speed"
        self.out_frame.pack_forget()
        self.speed_frame.pack(fill="both", expand=True)

    # ---------------- info rendering ----------------
    def show_info(self, title, rows):
        try:
            self.st.stop()
        except Exception:
            pass
        self._show_out()
        self.title_lbl.config(text=title.upper())
        t = self.txt
        t.config(state="normal")
        t.delete("1.0", "end")
        for row in rows:
            if row[0] == "row":
                _, key, value, tags = row
                t.insert("end", key.upper() + "\t", ("k",))
                t.insert("end", str(value) + "\n", ("v",) + tuple(tags))
            elif row[0] == "bar":
                t.insert("end", bar_text(row[1]) + "\n", ("bar",))
            else:
                t.insert("end", row[1] + "\n", ("hdr",))
            t.insert("end", "\u2500" * 60 + "\n", ("sep",))
        t.config(state="disabled")
        t.yview_moveto(0.0)

    def replace_tag(self, tagname, text, extra_tag=None):
        t = self.txt
        t.config(state="normal")
        rng = t.tag_ranges(tagname)
        if rng:
            start, end = rng[0], rng[1]
            t.delete(start, end)
            tags = (tagname,) if not extra_tag else (tagname, extra_tag)
            t.insert(start, text, tags)
        t.config(state="disabled")

    # ---------------- WiFi ----------------
    def show_wifi(self):
        self.show_info("WiFi Passwords",
                       [("row", "WiFi", "Scanning saved profiles...", ())])
        self.progress.start_spin("Scanning WiFi profiles...", color="#a78bfa")
        threading.Thread(target=self._wifi_worker, daemon=True).start()

    def _wifi_worker(self):
        def progress_cb(i, total, name):
            if total:
                self.post("progress", i / float(total),
                          "WiFi profile {} / {}".format(i, total))
        try:
            rows = rows_wifi(progress_cb=progress_cb)
        finally:
            self.post("progress_stop")
        self.post("show_info", "WiFi Passwords", rows)

    # ---------------- Doc2Pdf ----------------
    def doc2pdf_pick_file(self):
        path = filedialog.askopenfilename(
            title="Choose a Word / PowerPoint / Excel file to convert",
            filetypes=[
                ("Office documents",
                 "*.doc *.docx *.ppt *.pptx *.xls *.xlsx"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self._doc2pdf_begin(path)

    def doc2pdf_pick_folder(self):
        path = filedialog.askdirectory(
            title="Choose a folder to convert (subfolders included)")
        if path:
            self._doc2pdf_begin(path)

    def _doc2pdf_begin(self, path):
        self.show_info("Doc → PDF Conversion",
                        [("row", "Source", path, ()),
                         ("row", "Status", "Starting...", ("pdfstatus",)),
                         ("row", "Converted", "0", ("pdfcount",))])
        self.progress.start_spin("Preparing conversion...", color="#fb923c")
        self.doc2pdf.convert_path(path)

    def _on_doc2pdf_progress(self, count, current_path):
        self.replace_tag("pdfcount", str(count))
        self.replace_tag("pdfstatus",
                          "Converting: {}".format(os.path.basename(current_path)))

    def _on_doc2pdf_status(self, message):
        self.replace_tag("pdfstatus", message)

    def _on_doc2pdf_status_detail(self, message):
        t = self.txt
        t.config(state="normal")
        t.insert("end", "  • " + message + "\n", ("warn",))
        t.config(state="disabled")

    # ---------------- Windows activation ----------------
    def activate_windows(self, win_version):
        if os.name != "nt":
            self.show_info("Windows Activation",
                           [("row", "Error",
                             "Activation is only available on Windows.",
                             ("warn",))])
            return
        if getattr(self, "_activating", False):
            self.show_info("Windows Activation",
                           [("row", "Status",
                             "An activation is already running. Please wait.",
                             ("warn",))])
            return
        ActivationDialog(
            self.root, win_version,
            on_pick=lambda edition: self._run_activation(win_version, edition))

    def _run_activation(self, win_version, edition):
        key = WIN_KEYS.get((win_version, edition))
        if not key:
            self.show_info("Windows Activation",
                           [("row", "Error",
                             "Unknown Windows version / edition combination.",
                             ("warn",))])
            return
        self._activating = True
        self.show_info("Windows Activation", [
            ("row", "Windows", "Windows {}".format(win_version), ()),
            ("row", "Edition", edition, ()),
            ("row", "Status", "Starting... approve the Windows (UAC) "
                              "prompt if it appears.", ()),
        ])
        self.progress.set_fraction(0.03, "Preparing activation...",
                                   color="#818cf8")
        threading.Thread(target=self._activation_worker,
                         args=(win_version, edition, key),
                         daemon=True).start()

    @staticmethod
    def _is_admin():
        try:
            return bool(ctypes.windll.shell32.IsUserAnAdmin())
        except Exception:
            return False

    @staticmethod
    def _launch_hidden_elevated(bat_path):
        class SHELLEXECUTEINFOW(ctypes.Structure):
            _fields_ = [
                ("cbSize", wintypes.DWORD),
                ("fMask", wintypes.ULONG),
                ("hwnd", wintypes.HWND),
                ("lpVerb", wintypes.LPCWSTR),
                ("lpFile", wintypes.LPCWSTR),
                ("lpParameters", wintypes.LPCWSTR),
                ("lpDirectory", wintypes.LPCWSTR),
                ("nShow", ctypes.c_int),
                ("hInstApp", wintypes.HINSTANCE),
                ("lpIDList", ctypes.c_void_p),
                ("lpClass", wintypes.LPCWSTR),
                ("hkeyClass", wintypes.HKEY),
                ("dwHotKey", wintypes.DWORD),
                ("hIconOrMonitor", wintypes.HANDLE),
                ("hProcess", wintypes.HANDLE),
            ]
        SEE_MASK_NOCLOSEPROCESS = 0x00000040
        SEE_MASK_NOASYNC        = 0x00000100
        sei = SHELLEXECUTEINFOW()
        sei.cbSize = ctypes.sizeof(sei)
        sei.fMask = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_NOASYNC
        sei.hwnd = None
        sei.lpVerb = "runas"
        sei.lpFile = "cmd.exe"
        sei.lpParameters = '/c ""{}""'.format(bat_path)
        sei.lpDirectory = None
        sei.nShow = 0
        ShellExecuteExW = ctypes.windll.shell32.ShellExecuteExW
        ShellExecuteExW.argtypes = [ctypes.POINTER(SHELLEXECUTEINFOW)]
        ShellExecuteExW.restype = wintypes.BOOL
        if not ShellExecuteExW(ctypes.byref(sei)):
            err = ctypes.GetLastError()
            if err == 1223:
                raise OSError("Administrator permission (UAC) was declined.")
            raise OSError("Could not start elevated process (error {}).".format(err))
        return sei.hProcess

    @staticmethod
    def _read_log(path):
        try:
            with open(path, "rb") as f:
                return f.read().decode("oem", errors="replace")
        except Exception:
            return ""

    def _activation_worker(self, win_version, edition, key):
        post = self.post
        bat_path = log_path = None
        rows = [
            ("row", "Windows", "Windows {}".format(win_version), ()),
            ("row", "Edition", edition, ()),
            ("row", "Product Key", key, ()),
            ("row", "KMS Server", KMS_SERVER, ()),
        ]
        try:
            system_root = os.environ.get("SystemRoot", r"C:\Windows")
            slmgr = os.path.join(system_root, "System32", "slmgr.vbs")

            fd, log_path = tempfile.mkstemp(suffix=".log", prefix="panda_act_")
            os.close(fd)
            fd, bat_path = tempfile.mkstemp(suffix=".bat", prefix="panda_act_")
            os.close(fd)

            lines = [
                "@echo off",
                'set "LOG={}"'.format(log_path),
                'echo ::STEP1:: > "%LOG%"',
                'cscript //nologo "{}" /ipk {} >> "%LOG%" 2>&1'.format(slmgr, key),
                'echo ::STEP2:: >> "%LOG%"',
                'cscript //nologo "{}" /skms {} >> "%LOG%" 2>&1'.format(slmgr, KMS_SERVER),
                'echo ::STEP3:: >> "%LOG%"',
                'cscript //nologo "{}" /ato >> "%LOG%" 2>&1'.format(slmgr),
                'echo ::DONE:: >> "%LOG%"',
            ]
            with open(bat_path, "wb") as f:
                f.write(("\r\n".join(lines) + "\r\n").encode("oem", errors="replace"))

            handle = proc = None
            if self._is_admin():
                proc = subprocess.Popen(
                    ["cmd.exe", "/c", bat_path],
                    stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    creationflags=0x08000000)
            else:
                post("progress", 0.05, "Waiting for administrator approval...")
                handle = self._launch_hidden_elevated(bat_path)

            k32 = ctypes.windll.kernel32
            k32.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
            k32.WaitForSingleObject.restype = wintypes.DWORD
            k32.CloseHandle.argtypes = [wintypes.HANDLE]

            labels = {
                0: ("Starting...",                       0.08, 0.14),
                1: ("[1/3] Installing product key...",   0.14, 0.40),
                2: ("[2/3] Setting KMS server...",       0.40, 0.58),
                3: ("[3/3] Activating (contacting KMS)...", 0.58, 0.97),
                4: ("Finished",                          1.00, 1.00),
            }
            t_start = stage_t = time.time()
            stage = -1
            finished = False
            timed_out = False
            while True:
                if proc is not None:
                    finished = proc.poll() is not None
                else:
                    finished = k32.WaitForSingleObject(handle, 200) == 0
                if proc is not None and not finished:
                    time.sleep(0.2)

                text = self._read_log(log_path)
                new_stage = 4 if "::DONE::" in text else (
                    3 if "::STEP3::" in text else (
                    2 if "::STEP2::" in text else (
                    1 if "::STEP1::" in text else 0)))
                if new_stage != stage:
                    stage, stage_t = new_stage, time.time()
                label, lo, hi = labels[stage]
                if stage < 4:
                    dt = time.time() - stage_t
                    frac = lo + (hi - lo) * (1.0 - math.exp(-dt / 6.0))
                else:
                    frac = 1.0
                post("progress", frac, label)

                if finished:
                    break
                if time.time() - t_start > 300:
                    timed_out = True
                    break

            if handle:
                k32.CloseHandle(handle)

            text = self._read_log(log_path)
            low = text.lower()
            out_lines = [l.strip() for l in text.splitlines()
                         if l.strip() and not l.strip().startswith("::")]

            if timed_out:
                status, tag = "Timed out waiting for activation to finish.", ("warn",)
            elif "activated successfully" in low or "activation succeeded" in low:
                status, tag = "Windows activated successfully.", ("ok",)
            elif "error" in low or "0xc004" in low or "0x8007" in low:
                status, tag = ("Activation failed. Check internet access "
                               "and try again."), ("warn",)
            elif "::done::" in low:
                status, tag = "Finished. Review the output below.", ()
            else:
                status, tag = "Activation did not complete.", ("warn",)

            rows.append(("row", "Status", status, tag))
            for i, line in enumerate(out_lines[:8], 1):
                rows.append(("row", "Output {}".format(i), line, ()))

            post("progress", 1.0, "Done" if tag == ("ok",) else "Finished")
            post("show_info", "Windows Activation", rows)
            time.sleep(1.2)
        except Exception as exc:
            rows.append(("row", "Status",
                         "Could not run activation: {}".format(exc), ("warn",)))
            post("show_info", "Windows Activation", rows)
        finally:
            post("progress_stop")
            self._activating = False
            for p in (bat_path, log_path):
                if p:
                    try:
                        os.remove(p)
                    except Exception:
                        pass

    def show_activation_status(self):
        if os.name != "nt":
            self.show_info("Activation Status",
                           [("row", "Error",
                             "Only available on Windows.", ("warn",))])
            return
        self.show_info("Activation Status",
                       [("row", "Status", "Querying...", ())])
        self.progress.start_spin("Checking activation status...", color="#818cf8")
        threading.Thread(target=self._activation_status_worker,
                         daemon=True).start()

    def _activation_status_worker(self):
        info = {}
        try:
            system_root = os.environ.get("SystemRoot", r"C:\Windows")
            slmgr_path  = os.path.join(system_root, "System32", "slmgr.vbs")
            r = subprocess.run(
                ["cscript", "//NoLogo", slmgr_path, "/dli"],
                capture_output=True, text=True, encoding="oem",
                errors="replace", timeout=20, stdin=subprocess.DEVNULL,
            )
            out = r.stdout or ""
            for line in out.splitlines():
                line = line.strip()
                if ":" not in line:
                    continue
                k, v = line.split(":", 1)
                k = k.strip().lower()
                v = v.strip()
                if not v:
                    continue
                if "name" in k:
                    info.setdefault("Name", v)
                elif "description" in k:
                    info.setdefault("Edition", v)
                elif "partial product key" in k:
                    info.setdefault("Partial Key", v)
                elif "license status" in k:
                    info.setdefault("License Status", v)
        except Exception as exc:
            info["Error"] = str(exc)

        rows = []
        for k in ("Name", "Edition", "Partial Key", "License Status", "Error"):
            if k in info:
                tag = ("ok",) if k == "License Status" and \
                       "licensed" in info[k].lower() else \
                       (("warn",) if k == "Error" else ())
                rows.append(("row", k, info[k], tag))
        if not rows:
            rows.append(("row", "Status", "No information returned.", ("warn",)))

        self.post("show_info", "Activation Status", rows)
        self.post("progress_stop")

    # ------------------------------------------------------------------
    #  Office 365 activation - pure Python (no .bat, no slmgr.vbs /ipk)
    # ------------------------------------------------------------------
    def activate_office(self):
        """Public entry point for the Office 365 button."""
        if os.name != "nt":
            self.show_info("Office 365", [
                ("row", "Product Name", "Office 365", ()),
                ("row", "Status",
                 "Office activation is only available on Windows.",
                 ("warn",)),
            ])
            return
        if getattr(self, "_activating_office", False):
            self.show_info("Office 365", [
                ("row", "Product Name", "Office 365", ()),
                ("row", "Status",
                 "An Office activation is already running. Please wait.",
                 ("warn",)),
            ])
            return

        # Look for sppc64.dll in a few sensible places.
        here = os.path.dirname(os.path.abspath(__file__))
        candidates = [
            OFFICE_DLL_SRC,                                              # %TEMP%\MyAppFiles\sppc64.dll
            os.path.join(os.environ.get("USERPROFILE", ""), "Desktop", OFFICE_DLL_NAME),
            os.path.join(here, OFFICE_DLL_NAME),                         # next to this script / exe
        ]
        src = next((p for p in candidates if p and os.path.isfile(p)), None)
        if src is None:
            listing = "\n  ".join(candidates)
            self.show_info("Office 365", [
                ("row", "Product Name", "Office 365", ()),
                ("row", "Status",
                 "{} was not found. Put it in one of these locations:\n  {}".format(
                     OFFICE_DLL_NAME, listing),
                 ("warn",)),
            ])
            return

        self._activating_office = True
        self.show_info("Office 365", [
            ("row", "Product Name", "Office 365", ()),
            ("row", "Source", src, ()),
            ("row", "Status", "Starting...", ()),
        ])
        self.progress.set_fraction(0.02, "Preparing Office activation...",
                                   color="#5eead4")
        threading.Thread(target=self._office_worker,
                         args=(src,), daemon=True).start()

    @staticmethod
    def _vfs_candidates():
        """Return every plausible Office <version>\\root\\vfs\\System path."""
        dirs = []
        for env in ("ProgramFiles", "ProgramFiles(x86)"):
            pf = os.environ.get(env)
            if pf:
                dirs.append(os.path.join(pf, r"Microsoft Office\root\vfs\System"))
        for env in ("ProgramFiles", "ProgramFiles(x86)"):
            pf = os.environ.get(env)
            if pf:
                for ver in ("Office16", "Office15", "Office14"):
                    dirs.append(os.path.join(pf, "Microsoft Office", ver,
                                             "vfs", "System"))
        return dirs

    @staticmethod
    def _office_ospp_candidates():
        dirs = []
        for env in ("ProgramFiles", "ProgramFiles(x86)"):
            pf = os.environ.get(env)
            if pf:
                for sub in (r"Microsoft Office\Office16",
                            r"Microsoft Office\root\Office16",
                            r"Microsoft Office\Office15",
                            r"Microsoft Office\root\Office15"):
                    dirs.append(os.path.join(pf, sub, "ospp.vbs"))
        return dirs

    @staticmethod
    def _run_slmgr_key_install(product_key):
        """Install the Office product key via the WMI
        SoftwareLicensingService class.

        Uses SWbemLocator instead of a WMI moniker string so we never
        have to escape backslashes. The previous moniker-based form
        raised (-2147221020, 'Invalid syntax').
        """
        if not _HAS_WIN32COM:
            raise RuntimeError(
                "pywin32 is required for Office activation. "
                "Install it with: pip install pywin32")

        locator = win32com.client.Dispatch("WbemScripting.SWbemLocator")
        svc = locator.ConnectServer(".", "root\\cimv2")
        svc.Get("SoftwareLicensingService").InstallProductKey(product_key)

    def _office_worker(self, src_dll):
        """Runs the entire Ohook flow, but only reports Product Name
        and a final Status line. All intermediate stages are shown by
        the progress bar labels."""
        post = self.post

        def stage(frac, label):
            post("progress", frac, label)

        try:
            # ---------- Step 1: stage the Ohook DLL -------------------
            stage(0.10, "[1/7] Staging Ohook DLL...")
            os.makedirs(OHOOK_DIR, exist_ok=True)
            ohook_dll = os.path.join(OHOOK_DIR, OFFICE_DLL_NAME)
            shutil.copyfile(src_dll, ohook_dll)

            # ---------- Step 2: locate Office VFS ---------------------
            stage(0.22, "[2/7] Locating Office VFS...")
            vfs = None
            for cand in self._vfs_candidates():
                if os.path.isdir(cand):
                    vfs = cand
                    break
            if vfs is None:
                post("show_info", "Office 365", [
                    ("row", "Product Name", "Office 365", ()),
                    ("row", "Status",
                     "Office VFS folder not found. Is Office installed?",
                     ("warn",)),
                ])
                post("progress_stop")
                return

            # ---------- Step 3: sppcs.dll → sppc.dll symlink ----------
            stage(0.34, "[3/7] Linking sppcs.dll ...")
            link_path = os.path.join(vfs, "sppcs.dll")
            target    = os.path.join(os.environ.get("windir", r"C:\Windows"),
                                     "System32", "sppc.dll")
            if not os.path.lexists(link_path):
                try:
                    os.symlink(target, link_path)
                except OSError:
                    subprocess.run(
                        ["cmd", "/c", "mklink", link_path, target],
                        check=True, creationflags=0x08000000,
                        stdin=subprocess.DEVNULL,
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                    )

            # ---------- Step 4: drop the Ohook DLL into VFS -----------
            stage(0.46, "[4/7] Placing Ohook DLL in VFS...")
            copy_dest = os.path.join(vfs, "sppc.dll")
            if os.path.lexists(copy_dest):
                os.remove(copy_dest)
            shutil.copyfile(ohook_dll, copy_dest)

            # ---------- Step 5: registry resiliency key ---------------
            stage(0.58, "[5/7] Writing registry keys...")
            import winreg
            key = winreg.CreateKeyEx(
                winreg.HKEY_CURRENT_USER, REG_RESILIENCY, 0,
                winreg.KEY_WRITE,
            )
            winreg.SetValueEx(
                key, REG_HEARTBEAT_VAL, 0,
                winreg.REG_SZ, REG_HEARTBEAT_DAT,
            )
            winreg.CloseKey(key)

            # ---------- Step 6: install product key (best effort) -----
            stage(0.72, "[6/7] Installing product key...")
            try:
                self._run_slmgr_key_install(OFFICE_KEY)
            except Exception:
                # Ohook itself performs the real activation; a failure
                # here is non-fatal so we just keep going.
                pass

            # ---------- Step 7: query Office license status -----------
            stage(0.86, "[7/7] Querying Office status...")
            ospp = next((p for p in self._office_ospp_candidates()
                         if os.path.isfile(p)), None)
            if ospp:
                try:
                    subprocess.run(
                        ["cscript", "//nologo", ospp, "/dstatus"],
                        capture_output=True, text=True, timeout=30,
                        creationflags=0x08000000,
                        stdin=subprocess.DEVNULL,
                    )
                except Exception:
                    pass

            # ---------- Cleanup ---------------------------------------
            stage(0.97, "[Done] Cleaning up...")
            if os.path.isdir(OHOOK_DIR):
                shutil.rmtree(OHOOK_DIR, ignore_errors=True)

            # ---------- Final report ----------------------------------
            post("progress", 1.0, "Done")
            post("show_info", "Office 365", [
                ("row", "Product Name", "Office 365", ("ok",)),
                ("row", "Status",
                 "Office activation steps completed.", ("ok",)),
            ])
            time.sleep(1.0)

        except Exception as exc:
            post("show_info", "Office 365", [
                ("row", "Product Name", "Office 365", ()),
                ("row", "Status",
                 "Activation failed: {}".format(exc), ("warn",)),
            ])
        finally:
            post("progress_stop")
            self._activating_office = False

    # ---------------- Junk cleaner ----------------
    def clean_junk(self):
        if os.name != "nt":
            self.show_info("Clean Junk Files", [
                ("row", "Status", "Junk cleaning is only available on Windows.",
                 ("warn",))])
            return
        if self._cleaning:
            return
        CleanDialog(self.root, self._start_clean)

    def _start_clean(self):
        if self._cleaning:
            return
        self._cleaning = True
        self.show_info("Clean Junk Files", [
            ("row", "Status", "Scanning and cleaning...", ()),
        ])
        self.progress.start_spin("Cleaning junk files...", color="#fbbf24")
        threading.Thread(target=self._clean_worker, daemon=True).start()

    def _clean_worker(self):
        post = self.post
        rows = []
        try:
            targets = junk_targets()
            is_admin = self._is_admin()
            user_t = [t for t in targets if not t[2]]
            admin_t = [t for t in targets if t[2]]
            results = {}               # label -> stats dict
            total_steps = len(targets) + 2
            step = 0

            # ---- user-level folders ----
            for label, path, _adm in user_t:
                post("progress", step / float(total_steps),
                     "Cleaning {}...".format(label))
                results[label] = _clean_target(path, fix_acl=False)
                step += 1

            # ---- admin-only folders ----
            admin_note = None
            if admin_t and is_admin:
                for label, path, _adm in admin_t:
                    post("progress", step / float(total_steps),
                         "Cleaning {}...".format(label))
                    results[label] = _clean_target(path, fix_acl=True)
                    step += 1
            elif admin_t:
                post("progress", step / float(total_steps),
                     "Waiting for administrator approval...")
                bat_path = log_path = None
                try:
                    fd, log_path = tempfile.mkstemp(suffix=".log",
                                                    prefix="panda_clean_")
                    os.close(fd)
                    fd, bat_path = tempfile.mkstemp(suffix=".bat",
                                                    prefix="panda_clean_")
                    os.close(fd)
                    count = ("for /f %%C in ('dir /a-d /s /b \"{p}\" 2^>nul "
                             "^| find /c /v \"\"') do set \"{v}=%%C\"")
                    purge = [
                        'del /f /s /q /a "{p}\\*" >nul 2>&1',
                        'for /d %%D in ("{p}\\*") do rd /s /q "%%D" >nul 2>&1',
                    ]
                    lines = ["@echo off", "setlocal enabledelayedexpansion"]
                    for i, (_l, path, _a) in enumerate(admin_t):
                        lines.append('set "B=0"')
                        lines.append(count.format(p=path, v="B"))
                        lines += [x.format(p=path) for x in purge]
                        lines.append('set "A=0"')
                        lines.append(count.format(p=path, v="A"))
                        lines.append('if not "!A!"=="0" (')
                        lines.append('  takeown /f "{p}" /r /d y >nul 2>&1'.format(p=path))
                        lines.append('  icacls "{p}" /grant *S-1-5-32-544:F /t /c /q >nul 2>&1'.format(p=path))
                        lines += ["  " + x.format(p=path) for x in purge]
                        lines.append('  set "A=0"')
                        lines.append("  " + count.format(p=path, v="A"))
                        lines.append(")")
                        lines.append('>>"{}" echo {}^|!B!^|!A!'.format(log_path, i))
                    with open(bat_path, "wb") as f:
                        f.write(("\r\n".join(lines) + "\r\n")
                                .encode("oem", errors="replace"))

                    handle = self._launch_hidden_elevated(bat_path)
                    post("progress_spin", "Cleaning system folders...")
                    k32 = ctypes.windll.kernel32
                    k32.WaitForSingleObject.argtypes = [wintypes.HANDLE,
                                                        wintypes.DWORD]
                    k32.CloseHandle.argtypes = [wintypes.HANDLE]
                    if handle:
                        k32.WaitForSingleObject(handle, 900000)
                        k32.CloseHandle(handle)

                    parsed = {}
                    for line in self._read_log(log_path).splitlines():
                        bits = line.strip().split("|")
                        if len(bits) == 3 and bits[0].isdigit():
                            try:
                                parsed[int(bits[0])] = (int(bits[1]), int(bits[2]))
                            except ValueError:
                                pass
                    if not parsed:
                        admin_note = ("The elevated cleanup did not report "
                                      "back. Try running the app as "
                                      "administrator.")
                    for i, (label, _p, _a) in enumerate(admin_t):
                        if i in parsed:
                            b, a = parsed[i]
                            st = _new_stats()
                            st["freed"] = None
                            st["removed"] = max(0, b - a)
                            st["in_use"] = a
                            results[label] = st
                except Exception as exc:
                    admin_note = str(exc)
                finally:
                    for pth in (bat_path, log_path):
                        if pth:
                            try:
                                os.remove(pth)
                            except Exception:
                                pass
            step = len(user_t) + len(admin_t)

            # ---- recycle bin ----
            post("progress", step / float(total_steps), "Emptying Recycle Bin...")
            rb_size, rb_items = _recycle_bin_info()
            _empty_recycle_bin()

            # ---- report ----
            post("progress", 1.0, "Done")
            grand = rb_size + sum(v["freed"] or 0 for v in results.values())
            files = sum(v["removed"] for v in results.values())
            rows.append(("row", "Total Freed",
                         _fmt_size(grand) if grand else "-", ("ok",)))
            rows.append(("row", "Files Deleted", str(files), ("ok",)))
            rows.append(("row", "Running As",
                         "Administrator" if is_admin else
                         "Standard user (UAC used for system folders)", ()))
            rows.append(("hdr", "Cleaned Locations"))
            for label, _path, _adm in targets:
                st = results.get(label)
                if st is None:
                    rows.append(("row", label, "Not cleaned (needs administrator)",
                                 ("warn",)))
                    continue
                parts = []
                if st["freed"] is not None:
                    parts.append(_fmt_size(st["freed"]))
                parts.append("{} file(s) removed".format(st["removed"]))
                notes = []
                if st["in_use"]:
                    notes.append("{} in use".format(st["in_use"]))
                if st["denied"]:
                    notes.append("{} access denied".format(st["denied"]))
                if st["other"]:
                    notes.append("{} other".format(st["other"]))
                text = ", ".join(parts)
                if notes:
                    text += "  [left: " + ", ".join(notes) + "]"
                rows.append(("row", label, text, ("warn",) if notes else ()))
            rows.append(("row", "Recycle Bin",
                         "{} ({} item(s))".format(_fmt_size(rb_size), rb_items),
                         ()))
            if admin_note:
                rows.append(("row", "Admin Folders",
                             "Not cleaned: {}".format(admin_note), ("warn",)))
            rows.append(("hdr", "Note"))
            rows.append(("row", "In use",
                         "Files held open by running programs cannot be "
                         "deleted; close apps and run again for more.", ()))
            time.sleep(0.4)
            post("show_info", "Clean Junk Files", rows)
        except Exception as exc:
            post("show_info", "Clean Junk Files", [
                ("row", "Status", "Cleanup failed: {}".format(exc), ("warn",)),
            ])
        finally:
            post("progress_stop")
            self._cleaning = False

    def copy_output(self):
        if self.mode == "speed":
            text = "\n".join([
                self.status_lbl.cget("text"),
                "Ping: " + self.box_ping.cget("text"),
                "Download: " + self.box_down.cget("text"),
                "Upload: " + self.box_up.cget("text"),
            ])
        else:
            text = self.txt.get("1.0", "end").strip()
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
        except Exception:
            pass

    # ---------------- speed drawing ----------------
    def _set_big(self, value):
        self.big_lbl.config(text="{:.1f}".format(value))
        pct = min(100.0, math.log(value + 1) / math.log(1001) * 100.0)
        self._last_bar_pct = pct
        self._draw_speed_bar_track()

    def _draw_speed_bar_track(self):
        c = self.bar
        c.delete("all")
        try:
            w = c.winfo_width()
            h = c.winfo_height()
        except Exception:
            return
        if w < 8 or h < 4:
            return
        pct = getattr(self, "_last_bar_pct", 0.0)
        round_rect(c, 0, 0, w, h, h / 2.0, fill="#1a2433", outline="")
        fw = int(w * max(0.0, min(100.0, pct)) / 100.0)
        if fw > 2:
            round_rect(c, 0, 0, fw, h, h / 2.0, fill=GREEN, outline="")

    def _draw_speed_graph(self):
        c = self.graph
        c.delete("all")
        try:
            w = c.winfo_width()
            h = c.winfo_height()
        except Exception:
            return
        if w < 20 or h < 10:
            return
        round_rect(c, 0, 0, w, h, 10, fill=BG, outline="")
        for i in range(1, 5):
            y = h * i / 5
            c.create_line(0, y, w, y, fill="#101720")
        a = self.samples
        if len(a) < 2:
            return
        m = max(a) * 1.15 or 1.0
        n = max(len(a), 60)
        step = w / (n - 1)
        line = []
        for i, v in enumerate(a):
            x = i * step
            y = h - (v / m) * (h - 8) - 4
            line.extend([x, y])
        pts = [0, h] + line + [(len(a) - 1) * step, h]
        c.create_polygon(pts, fill=GREEN, outline="", stipple="gray75",
                         smooth=True, splinesteps=18)
        c.create_line(*line, fill=GREEN, width=5, smooth=True, splinesteps=18,
                     capstyle=tk.ROUND, joinstyle=tk.ROUND, stipple="gray25")
        c.create_line(*line, fill=GREEN, width=2, smooth=True, splinesteps=18,
                     capstyle=tk.ROUND, joinstyle=tk.ROUND)
        lx, ly = line[-2], line[-1]
        c.create_oval(lx - 3, ly - 3, lx + 3, ly + 3,
                      fill=GREEN, outline=BG, width=1)

    # ---------------- queue ----------------
    def _handle(self, kind, args):
        if kind == "stats":
            self._on_stats(*args)
        elif kind == "show_speed":
            self._show_speed()
        elif kind == "show_info":
            self.show_info(args[0], args[1])
        elif kind == "title":
            self.title_lbl.config(text=str(args[0]).upper())
        elif kind == "status":
            self.status_lbl.config(text=str(args[0]))
        elif kind == "ping":
            self.box_ping.config(text="{:.0f} ms".format(args[0]))
        elif kind == "down":
            self.box_down.config(text="{:.1f} Mbps".format(args[0]))
        elif kind == "up":
            self.box_up.config(text="{:.1f} Mbps".format(args[0]))
        elif kind == "big":
            self._set_big(float(args[0]))
        elif kind == "sample":
            self.samples.append(float(args[0]))
            if len(self.samples) > 150:
                self.samples.pop(0)
            self._set_big(float(args[0]))
            self._draw_speed_graph()
        elif kind == "clear_graph":
            self.samples = []
            self._draw_speed_graph()
        elif kind == "reset_speed":
            self.samples = []
            self._set_big(0.0)
            self._draw_speed_graph()
            self.box_ping.config(text="--")
            self.box_down.config(text="--")
            self.box_up.config(text="--")
        elif kind == "replace_tag":
            self.replace_tag(args[0], args[1], args[2])
        elif kind == "doc2pdf_progress":
            self._on_doc2pdf_progress(args[0], args[1])
        elif kind == "doc2pdf_status":
            self._on_doc2pdf_status(args[0])
        elif kind == "doc2pdf_status_detail":
            self._on_doc2pdf_status_detail(args[0])
        elif kind == "progress":
            if len(args) >= 2 and args[1]:
                self.progress.set_fraction(args[0], args[1])
            else:
                self.progress.set_fraction(args[0])
        elif kind == "progress_spin":
            label = args[0] if args else "Working..."
            self.progress.start_spin(label)
        elif kind == "progress_stop":
            self.progress.stop()

    def _on_stats(self, cpu, mem_pct, mem_used, mem_total,
                  disk_pct, disk_used, disk_total,
                  cpu_hist, mem_hist):
        self.cpu_hist = cpu_hist
        self.mem_hist = mem_hist
        self.disk_pct = disk_pct

        cpu_color = GREEN if cpu < 60 else (YELLOW if cpu < 85 else RED)
        self.cpu_val.config(text="{:.0f}%".format(cpu), fg=cpu_color)
        self.cpu_sub.config(text=get_cpu_sub())
        self.cpu_graph._color = cpu_color
        self._draw_spark(self.cpu_graph, cpu_hist, cpu_color)

        mem_color = GREEN if mem_pct < 65 else (YELLOW if mem_pct < 85 else RED)
        self.mem_val.config(text="{:.0f}%".format(mem_pct), fg=mem_color)
        if mem_total > 0:
            self.mem_sub.config(
                text="{:.1f} / {:.1f} GB".format(mem_used / 1073741824.0,
                                                  mem_total / 1073741824.0))
        self.mem_graph._color = mem_color
        self._draw_spark(self.mem_graph, mem_hist, mem_color)

        disk_color = GREEN if disk_pct < 70 else (YELLOW if disk_pct < 90 else RED)
        self.dsk_val.config(text="{}%".format(int(disk_pct)), fg=disk_color)
        if disk_total > 0:
            self.dsk_sub.config(
                text="{:.0f} / {:.0f} GB".format(disk_used / 1073741824.0,
                                                  disk_total / 1073741824.0))
        self.dsk_graph._color = disk_color
        self._draw_bar(self.dsk_graph, disk_pct, disk_color)


# --------------------------------------------------------------------------
def main():
    global UI_SCALE
    UI_SCALE = enable_dpi_awareness()

    root = tk.Tk()
    try:
        root.tk.call("tk", "scaling", UI_SCALE * (96.0 / 72.0))
    except Exception:
        pass
    root.title(APP_NAME)
    app = App(root)

    def on_close():
        try:
            if app.monitor:
                app.monitor.stop()
            app.st.stop()
        except Exception:
            pass
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()


if __name__ == "__main__":
    main()
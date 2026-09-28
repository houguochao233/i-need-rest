"""轻量、无害的全屏蓝屏风格计时器与 .lnk 备份/恢复管理器。

仅适用于 Windows。它不会触发真实系统蓝屏或修改系统设置。
"""
from __future__ import annotations

import argparse
import ctypes
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import time
import threading
import uuid
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD

try:
    from win32com.client import Dispatch
    import pythoncom
    from PIL import Image, ImageDraw, ImageTk
except ImportError as exc:
    raise SystemExit("需要 pywin32 和 Pillow。请运行: pip install pywin32 pillow") from exc

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
PROFILES_DIR = DATA_DIR / "profiles"
BACKUPS_DIR = DATA_DIR / "backups"

APP_NAME = "I Need Rest"
APP_TITLE = "I Need Rest - \u4f2a\u84dd\u5c4f\u6478\u9c7c\u5de5\u5177"
DEFAULT_HOTKEY = "ctrl+shift+q"
ACCENT_BLUE = "#0078d7"  # Windows 10 stop-screen blue
THEMES = {
    "win11": {"label": "Windows 11", "bg": "#0078D4", "font": "Segoe UI"},
    "win10": {"label": "Windows 10", "bg": "#0078D7", "font": "Segoe UI"},
    "win7": {"label": "Windows 7", "bg": "#1B4F9C", "font": "Segoe UI"},
    "winxp": {"label": "Windows XP", "bg": "#0000AA", "font": "Lucida Console"},
}


def ensure_storage() -> None:
    PROFILES_DIR.mkdir(parents=True, exist_ok=True)
    BACKUPS_DIR.mkdir(parents=True, exist_ok=True)


def file_hash(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def save_profile(profile: dict) -> Path:
    ensure_storage()
    path = PROFILES_DIR / f"{profile['id']}.json"
    path.write_text(json.dumps(profile, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def read_profile(profile_id: str) -> dict:
    return json.loads((PROFILES_DIR / f"{profile_id}.json").read_text(encoding="utf-8"))


def all_profiles() -> list[dict]:
    ensure_storage()
    profiles = []
    for item in PROFILES_DIR.glob("*.json"):
        try:
            profiles.append(json.loads(item.read_text(encoding="utf-8")))
        except (OSError, json.JSONDecodeError):
            continue
    return sorted(profiles, key=lambda p: p.get("created_at", 0), reverse=True)


def shortcut_data(path: Path) -> dict:
    shell = Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(str(path))
    return {
        "target": shortcut.Targetpath or "",
        "arguments": shortcut.Arguments or "",
        "working_directory": shortcut.WorkingDirectory or "",
        "icon_location": shortcut.IconLocation or "",
        "description": shortcut.Description or "",
        "window_style": int(shortcut.WindowStyle or 1),
    }


def launcher_runtime() -> Path:
    pythonw = Path(sys.executable).with_name("pythonw.exe")
    return pythonw if pythonw.exists() else Path(sys.executable)


def configure_launcher(link_path: Path, profile_id: str, icon_location: str) -> None:
    """Configure a .lnk file as the local, no-console launcher."""
    shell = Dispatch("WScript.Shell")
    link = shell.CreateShortcut(str(link_path))
    link.Targetpath = str(launcher_runtime())
    link.Arguments = f'"{Path(__file__).resolve()}" --launch "{profile_id}"'
    link.WorkingDirectory = str(APP_DIR)
    link.Description = ""
    link.WindowStyle = 7
    # Keep the original icon resource, including a custom icon index.
    if icon_location:
        link.IconLocation = icon_location
    link.Save()
    # Do not force Explorer refresh; avoid desktop flicker.
    # Sending an extra SHChangeNotify here causes a visible desktop refresh.


def shortcut_icon_location(original: dict) -> str:
    """Return a usable icon resource, including the target fallback."""
    icon = (original.get("icon_location") or "").strip()
    # WScript.Shell reports an inherited icon as ",0". Passing that back
    # Do not force Explorer refresh; avoid desktop flicker.
    if (not icon or icon.startswith(",")) and original.get("target"):
        icon = f'{original["target"]},0'
    return icon


def replace_shortcut(path: Path, profile_id: str, original: dict) -> None:
    """Replace a .lnk in place, keeping its filename and icon resource."""
    configure_launcher(path, profile_id, shortcut_icon_location(original))


def create_shadow_shortcut(original_path: Path, profile_id: str, icon_source: Path | None = None) -> Path:
    """Create a same-looking launcher beside a moved regular file.

    Windows hides the final .lnk suffix by default, so `report.docx.lnk` is shown as
    `report.docx` and keeps the document's resolved shell icon.
    """
    proxy = original_path.with_name(original_path.name + ".lnk")
    configure_launcher(proxy, profile_id, f"{icon_source or original_path},0")
    return proxy


def profile_state(profile: dict) -> str:
    """Return a cheap, filesystem-only health state for a managed item."""
    if profile.get("restored_at"):
        return "restored"
    backup = Path(profile.get("backup_path", ""))
    current = Path(profile.get("shortcut_path", ""))
    destination = Path(profile.get("original_path", profile.get("shortcut_path", "")))
    if not backup.exists():
        return "missing_backup"
    if profile.get("source_kind", "shortcut") == "file" and destination.exists():
        return "conflict"
    if not current.exists():
        return "missing_proxy"
    return "active"


def profile_is_active(profile: dict) -> bool:
    return profile_state(profile) == "active"


def profile_needs_attention(profile: dict) -> bool:
    return profile_state(profile) in {"missing_backup", "missing_proxy", "conflict"}


def profile_can_restore(profile: dict) -> bool:
    return profile_state(profile) in {"active", "missing_proxy"}

def restore_profile(profile: dict, preserve_current: bool = False) -> tuple[bool, str]:
    source = Path(profile["backup_path"])
    if not source.exists():
        return False, "\u627e\u4e0d\u5230\u5907\u4efd\u6587\u4ef6\uff0c\u65e0\u6cd5\u6062\u590d\u3002"
    source_kind = profile.get("source_kind", "shortcut")
    destination = Path(profile.get("original_path", profile["shortcut_path"]))
    proxy = Path(profile["shortcut_path"])
    temp = destination.with_name(destination.name + ".bsod-restore-tmp")
    try:
        destination.parent.mkdir(parents=True, exist_ok=True)
        if source_kind == "file" and destination.exists():
            return False, f"\u6062\u590d\u51b2\u7a81\uff1a\u539f\u4f4d\u5df2\u5b58\u5728\u540c\u540d\u6587\u4ef6\uff1a{destination.name}\uff0c\u8bf7\u5148\u624b\u52a8\u5904\u7406\u3002"
        if preserve_current and proxy.exists():
            stamp = time.strftime("%Y%m%d-%H%M%S")
            conflict = proxy.with_name(f"{proxy.stem} - \u5f53\u524d\u7248\u672c {stamp}{proxy.suffix}")
            shutil.copy2(proxy, conflict)
        if temp.exists():
            temp.unlink()
        shutil.copy2(source, temp)
        if source_kind == "file" and proxy.exists():
            proxy.unlink()
        os.replace(temp, destination)
        profile["restored_at"] = time.time()
        save_profile(profile)
        message = "\u5df2\u6062\u590d\u539f\u59cb\u6587\u4ef6\u3002" if source_kind == "file" else "\u5df2\u6062\u590d\u539f\u59cb\u5feb\u6377\u65b9\u5f0f\u3002"
        return True, message
    except OSError as exc:
        try:
            if temp.exists():
                temp.unlink()
        except OSError:
            pass
        return False, f"\u6062\u590d\u5931\u8d25\uff1a{exc}"

def normalize_hotkey(value: str) -> str:
    value = value.strip().lower().replace(" ", "")
    aliases = {"control": "ctrl", "escape": "esc"}
    parts = [aliases.get(part, part) for part in value.split("+") if part]
    return "+".join(parts)


def tk_hotkey(value: str) -> str:
    keys = normalize_hotkey(value).split("+")
    modifiers = []
    key = "q"
    mapping = {"ctrl": "Control", "shift": "Shift", "alt": "Alt"}
    for item in keys:
        if item in mapping:
            modifiers.append(mapping[item])
        else:
            key = item
    return "<" + "-".join(modifiers + [key]) + ">"


def build_qr_image(size: int = 132) -> Image.Image:
    """A deterministic QR-like visual block; no scanning purpose or external dependency."""
    modules = 29
    pad = 4
    cell = max(1, (size - pad * 2) // modules)
    actual = cell * modules + pad * 2
    image = Image.new("RGB", (actual, actual), "white")
    draw = ImageDraw.Draw(image)

    def finder(x: int, y: int) -> None:
        draw.rectangle((pad + x*cell, pad + y*cell, pad + (x+7)*cell-1, pad + (y+7)*cell-1), fill="black")
        draw.rectangle((pad + (x+1)*cell, pad + (y+1)*cell, pad + (x+6)*cell-1, pad + (y+6)*cell-1), fill="white")
        draw.rectangle((pad + (x+2)*cell, pad + (y+2)*cell, pad + (x+5)*cell-1, pad + (y+5)*cell-1), fill="black")

    blocked = set()
    for x, y in ((0, 0), (modules - 7, 0), (0, modules - 7)):
        finder(x, y)
        for yy in range(y, y + 8):
            for xx in range(x, x + 8):
                blocked.add((xx, yy))
    seed = 0x53A9D17
    for y in range(modules):
        for x in range(modules):
            if (x, y) in blocked:
                continue
            # A stable pseudo-random matrix makes the block look like the standard stop-code QR area.
            bit = ((x * 17 + y * 31 + (x * y) * 7 + seed) ^ (x << 2) ^ (y << 1)) & 1
            if bit:
                draw.rectangle((pad+x*cell, pad+y*cell, pad+(x+1)*cell-1, pad+(y+1)*cell-1), fill="black")
    return image


class BsodWindow:
    def __init__(self, profile: dict):
        self.profile = profile
        self.root = tk.Tk()
        self.root.configure(bg=ACCENT_BLUE, cursor="none")
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.protocol("WM_DELETE_WINDOW", lambda: None)
        self.finished = False
        self.start = time.monotonic()
        self.duration = max(1, int(profile.get("duration_seconds", 1200)))
        self.start_percent = max(0, min(99, int(profile.get("start_percent", 0))))
        self.auto_mode = profile.get("restore_mode", "once")
        self.theme_id = profile.get("theme", "win10")
        if self.theme_id not in THEMES:
            self.theme_id = "win10"
        self.theme = THEMES[self.theme_id]
        self.root.configure(bg=self.theme["bg"])

        user32 = ctypes.windll.user32
        x = user32.GetSystemMetrics(76)  # SM_XVIRTUALSCREEN
        y = user32.GetSystemMetrics(77)  # SM_YVIRTUALSCREEN
        width = user32.GetSystemMetrics(78)  # SM_CXVIRTUALSCREEN
        height = user32.GetSystemMetrics(79)  # SM_CYVIRTUALSCREEN
        self.root.geometry(f"{width}x{height}{x:+d}{y:+d}")

        self.root.bind_all(tk_hotkey(profile.get("hotkey", DEFAULT_HOTKEY)), self.exit)
        self.root.bind_all("<Escape>", self.exit)  # documented emergency exit; avoids accidental lock-in
        self.build_layout(width, height)
        self.root.after(20, self.tick)

    def build_layout(self, width: int, height: int) -> None:
        scale = max(0.72, min(width / 1920, height / 1080))
        bg = self.theme["bg"]
        font = self.theme["font"]

        if self.theme_id in ("win10", "win11"):
            left = int(width * 0.158)
            top = int(height * 0.175)
            content = tk.Frame(self.root, bg=bg)
            content.place(x=left, y=top, anchor="nw")
            face_size = 106 if self.theme_id == "win10" else 104
            face = tk.Label(content, text=":(", fg="white", bg=bg,
                            font=(font, max(44, int(face_size * scale)), "normal"))
            face.pack(anchor="w", pady=(0, int(25 * scale)))
            if self.theme_id == "win11":
                message = "\u4f60\u7684\u8bbe\u5907\u9047\u5230\u95ee\u9898\uff0c\u9700\u8981\u91cd\u542f\u3002\n\u6211\u4eec\u53ea\u6536\u96c6\u67d0\u4e9b\u9519\u8bef\u4fe1\u606f\uff0c\u7136\u540e\u4e3a\u4f60\u91cd\u65b0\u542f\u52a8\u3002"
                detail_text = (
                    "\u6709\u5173\u6b64\u95ee\u9898\u7684\u8be6\u7ec6\u4fe1\u606f\u548c\u53ef\u80fd\u7684\u89e3\u51b3\u65b9\u6cd5\uff0c\u8bf7\u8bbf\u95ee\n"
                    "https://www.windows.com/stopcode\n\n"
                    "\u5982\u679c\u81f4\u7535\u652f\u6301\u4eba\u5458\uff0c\u8bf7\u5411\u4ed6\u4eec\u63d0\u4f9b\u4ee5\u4e0b\u4fe1\u606f\uff1a\n"
                    "\u7ec8\u6b62\u4ee3\u7801\uff1aCRITICAL_PROCESS_DIED"
                )
            else:
                message = "\u4f60\u7684\u8bbe\u5907\u9047\u5230\u95ee\u9898\uff0c\u9700\u8981\u91cd\u542f\u3002\n\u6211\u4eec\u53ea\u6536\u96c6\u67d0\u4e9b\u9519\u8bef\u4fe1\u606f\uff0c\u7136\u540e\u4e3a\u4f60\u91cd\u65b0\u542f\u52a8\u3002"
                detail_text = (
                    "\u6709\u5173\u6b64\u95ee\u9898\u7684\u8be6\u7ec6\u4fe1\u606f\u548c\u53ef\u80fd\u7684\u89e3\u51b3\u65b9\u6cd5\uff0c\u8bf7\u8bbf\u95ee\n"
                    "https://www.windows.com/stopcode\n\n"
                    "\u5982\u679c\u81f4\u7535\u652f\u6301\u4eba\u5458\uff0c\u8bf7\u5411\u4ed6\u4eec\u63d0\u4f9b\u4ee5\u4e0b\u4fe1\u606f\uff1a\n"
                    "\u7ec8\u6b62\u4ee3\u7801\uff1aCRITICAL_PROCESS_DIED"
                )
            tk.Label(content, text=message, fg="white", bg=bg, justify="left",
                     font=(font, max(17, int(29 * scale))), anchor="w").pack(anchor="w")
            self.progress_label = tk.Label(content, text="", fg="white", bg=bg,
                                           font=(font, max(16, int(27 * scale))), anchor="w")
            self.progress_label.pack(anchor="w", pady=(int(28 * scale), int(63 * scale)))
            detail = tk.Frame(content, bg=bg)
            detail.pack(anchor="w")
            qr = ImageTk.PhotoImage(build_qr_image(max(112, int(154 * scale))))
            self.root._qr_ref = qr
            tk.Label(detail, image=qr, bg="white", bd=0).pack(side="left", padx=(0, int(28 * scale)))
            tk.Label(detail, text=detail_text, fg="white", bg=bg, justify="left", anchor="w",
                     font=(font, max(10, int(15 * scale)))).pack(side="left")
            return

        if self.theme_id == "win7":
            # Classic Windows 7 stop screen: dark blue, Segoe UI, technical white text.
            left, top = int(width * 0.115), int(height * 0.13)
            content = tk.Frame(self.root, bg=bg)
            content.place(x=left, y=top, anchor="nw")
            tk.Label(content, text=":(", fg="white", bg=bg,
                     font=(font, max(42, int(92 * scale)))).pack(anchor="w", pady=(0, int(18 * scale)))
            text = (
                "Windows \u9047\u5230\u95ee\u9898\uff0c\u9700\u8981\u91cd\u65b0\u542f\u52a8\u3002\n\n"
                "\u5982\u679c\u8fd9\u662f\u4f60\u7b2c\u4e00\u6b21\u770b\u5230\u6b64\u505c\u6b62\u9519\u8bef\u5c4f\u5e55\uff0c\u8bf7\u91cd\u65b0\u542f\u52a8\u8ba1\u7b97\u673a\u3002\n"
                "\u5982\u679c\u518d\u6b21\u51fa\u73b0\u6b64\u5c4f\u5e55\uff0c\u8bf7\u6267\u884c\u4ee5\u4e0b\u6b65\u9aa4\uff1a\n\n"
                "\u68c0\u67e5\u662f\u5426\u5b89\u88c5\u4e86\u65b0\u7684\u786c\u4ef6\u6216\u8f6f\u4ef6\u3002\n"
                "\u5982\u679c\u9700\u8981\u4f7f\u7528\u5b89\u5168\u6a21\u5f0f\u5220\u9664\u6216\u7981\u7528\u7ec4\u4ef6\uff0c\u8bf7\u91cd\u65b0\u542f\u52a8\u8ba1\u7b97\u673a\uff0c\n"
                "\u6309 F8 \u9009\u62e9\u9ad8\u7ea7\u542f\u52a8\u9009\u9879\uff0c\u7136\u540e\u9009\u62e9\u5b89\u5168\u6a21\u5f0f\u3002\n\n"
                "\u6280\u672f\u4fe1\u606f\uff1a\n"
                "*** STOP: 0x0000007E (0xC0000005, 0xFFFFF800, 0x00000000, 0x00000000)"
            )
            tk.Label(content, text=text, fg="white", bg=bg, justify="left", anchor="w",
                     font=(font, max(12, int(19 * scale)))).pack(anchor="w")
            self.progress_label = tk.Label(content, text="", fg=bg, bg=bg,
                                           font=(font, 1))
            self.progress_label.pack()
            return

        # Classic Windows XP stop screen: solid navy, monospaced white text.
        content = tk.Frame(self.root, bg=bg)
        content.place(x=int(width * 0.055), y=int(height * 0.055), anchor="nw")
        text = (
            "A problem has been detected and Windows has been shut down to prevent damage\n"
            "to your computer.\n\n"
            "If this is the first time you've seen this Stop error screen,\n"
            "restart your computer. If this screen appears again, follow\n"
            "these steps:\n\n"
            "Check to make sure any new hardware or software is properly installed.\n"
            "If problems continue, disable or remove any newly installed hardware\n"
            "or software. Disable BIOS memory options such as caching or shadowing.\n\n"
            "Technical information:\n"
            "*** STOP: 0x0000007B (0xF78D2524, 0xC0000034, 0x00000000, 0x00000000)\n\n"
            "Beginning dump of physical memory\n"
            "Physical memory dump complete. Contact your system administrator or\n"
            "technical support group for further assistance."
        )
        tk.Label(content, text=text, fg="white", bg=bg, justify="left", anchor="nw",
                 font=(font, max(10, int(16 * scale)), "normal")).pack(anchor="nw")
        self.progress_label = tk.Label(content, text="", fg=bg, bg=bg,
                                       font=(font, 1))
        self.progress_label.pack()

    def tick(self) -> None:
        if self.finished:
            return
        elapsed = time.monotonic() - self.start
        ratio = min(1.0, elapsed / self.duration)
        percent = self.start_percent + round((100 - self.start_percent) * ratio)
        self.progress_label.configure(text=f"{percent}% 完成")
        if ratio >= 1:
            self.exit()
        else:
            self.root.after(220, self.tick)

    def exit(self, _event=None) -> None:
        if self.finished:
            return "break"
        self.finished = True
        if self.auto_mode == "after":
            restore_profile(self.profile)
        self.root.destroy()
        return "break"

    def run(self) -> None:
        self.root.mainloop()


class ShortcutManager(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        ensure_storage()
        self.title(APP_TITLE)
        self.resizable(False, False)
        self.configure(padx=14, pady=14)
        self.selected_path = tk.StringVar()
        self.duration_minutes = tk.IntVar(value=20)
        self.start_percent = tk.IntVar(value=0)
        self.restore_mode = tk.StringVar(value="once")
        self.theme = tk.StringVar(value="win10")
        self.hotkey = tk.StringVar(value=DEFAULT_HOTKEY)
        self.status = tk.StringVar(value="将一个 .lnk 快捷方式拖到下方区域，或点击选择。")
        self.build_ui()
        self.enable_drag_drop()
        self.refresh_profiles()
        self.after(250, self.check_startup_profiles)
        # The blue-screen launcher restores files in a separate process.
        self.after(1500, self.auto_refresh_profiles)

    def build_ui(self) -> None:
        header = ttk.Label(self, text=APP_NAME, font=("Segoe UI", 13, "bold"))
        header.grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(self, text="本地视觉模拟；不触发真实系统蓝屏。", foreground="#555555").grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(1, 10))

        self.drop = tk.Label(self, text="\u5c06\u6587\u4ef6\u6216 Windows \u5feb\u6377\u65b9\u5f0f\u62d6\u5230\u8fd9\u91cc\n\u6216\u70b9\u51fb\u9009\u62e9", relief="groove",
                             bd=1, width=48, height=4, bg="#f7f7f7", cursor="hand2", justify="center")
        self.drop.grid(row=2, column=0, columnspan=3, sticky="ew")
        self.drop.bind("<Button-1>", lambda _e: self.choose_shortcut())
        ttk.Label(self, textvariable=self.selected_path, wraplength=415, foreground="#444444").grid(
            row=3, column=0, columnspan=3, sticky="w", pady=(5, 9))

        ttk.Label(self, text="\u84dd\u5c4f\u65f6\u957f\uff08\u5206\u949f\uff09").grid(row=4, column=0, sticky="w")
        ttk.Spinbox(self, from_=1, to=180, textvariable=self.duration_minutes, width=8).grid(row=4, column=1, sticky="w")
        ttk.Label(self, text="\u5f00\u59cb\u8fdb\u5ea6\uff08%\uff09").grid(row=5, column=0, sticky="w", pady=(6, 0))
        ttk.Spinbox(self, from_=0, to=99, textvariable=self.start_percent, width=8).grid(row=5, column=1, sticky="w", pady=(6, 0))
        ttk.Label(self, text="\u84dd\u5c4f\u4e3b\u9898").grid(row=6, column=0, sticky="w", pady=(6, 0))
        self.theme_display_values = {f"{key}: {value['label']}": key for key, value in THEMES.items()}
        ttk.Combobox(self, textvariable=self.theme, state="readonly", width=18,
                     values=list(self.theme_display_values)).grid(row=6, column=1, columnspan=2, sticky="w", pady=(6, 0))
        self.theme.set("win10: Windows 10")
        ttk.Label(self, text="\u9000\u51fa\u70ed\u952e").grid(row=7, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(self, textvariable=self.hotkey, width=18).grid(row=7, column=1, columnspan=2, sticky="w", pady=(6, 0))

        modes = ttk.LabelFrame(self, text="恢复方式", padding=7)
        modes.grid(row=8, column=0, columnspan=3, sticky="ew", pady=(10, 8))
        ttk.Radiobutton(modes, text="一次后立即恢复（推荐）", variable=self.restore_mode, value="once").pack(anchor="w")
        ttk.Radiobutton(modes, text="蓝屏结束后恢复", variable=self.restore_mode, value="after").pack(anchor="w")
        ttk.Radiobutton(modes, text="手动恢复", variable=self.restore_mode, value="manual").pack(anchor="w")

        actions = ttk.Frame(self)
        actions.grid(row=9, column=0, columnspan=3, sticky="ew")
        ttk.Button(actions, text="预览蓝屏", command=self.preview).pack(side="left")
        ttk.Button(actions, text="生成蓝屏快捷方式", command=self.take_over).pack(side="right")
        ttk.Label(self, textvariable=self.status, foreground="#555555", wraplength=415).grid(
            row=10, column=0, columnspan=3, sticky="w", pady=(7, 11))

        managed = ttk.LabelFrame(self, text="已管理的快捷方式", padding=7)
        managed.grid(row=11, column=0, columnspan=3, sticky="ew")
        self.tree = ttk.Treeview(managed, columns=("name", "mode", "state"), show="headings", height=5, selectmode="browse")
        self.tree.heading("name", text="快捷方式")
        self.tree.heading("mode", text="恢复方式")
        self.tree.heading("state", text="状态")
        self.tree.column("name", width=220)
        self.tree.column("mode", width=100)
        self.tree.column("state", width=75)
        self.tree.pack(fill="x")
        lower = ttk.Frame(managed)
        lower.pack(fill="x", pady=(7, 0))
        ttk.Button(lower, text="恢复选中项", command=self.restore_selected).pack(side="left")
        ttk.Button(lower, text="删除记录", command=self.delete_selected).pack(side="left", padx=(6, 0))
        ttk.Button(lower, text="恢复全部", command=self.restore_all).pack(side="right")
        ttk.Button(lower, text="清理全部记录", command=self.delete_all_records).pack(side="right", padx=(0, 6))

    def choose_shortcut(self) -> None:
        value = filedialog.askopenfilename(
            title="\u9009\u62e9\u6587\u4ef6\u6216 Windows \u5feb\u6377\u65b9\u5f0f", filetypes=[("\u6240\u6709\u6587\u4ef6", "*.*")]
        )
        if value:
            self.set_shortcut(Path(value))

    def set_shortcut(self, path: Path) -> None:
        if not path.is_file():
            messagebox.showerror("\u4e0d\u652f\u6301\u7684\u9879\u76ee", "\u8bf7\u62d6\u5165\u6216\u9009\u62e9\u4e00\u4e2a\u6587\u4ef6\uff1b\u6587\u4ef6\u5939\u6682\u4e0d\u652f\u6301\u3002", parent=self)
            return
        self.selected_path.set(str(path))
        if path.suffix.lower() == ".lnk":
            try:
                info = shortcut_data(path)
                self.status.set("\u5df2\u9009\u62e9\u5feb\u6377\u65b9\u5f0f\uff1a{}  \u2192  {}".format(path.name, info['target'] or '\u672a\u8bbe\u7f6e\u76ee\u6807'))
            except Exception as exc:
                self.status.set(f"\u5df2\u9009\u62e9\uff1a{path.name}\uff08\u8bfb\u53d6\u5c5e\u6027\u5931\u8d25\uff1a{exc}\uff09")
        else:
            self.status.set(f"\u5df2\u9009\u62e9\u6587\u4ef6\uff1a{path.name}\u3002\u5c06\u5907\u4efd\u539f\u6587\u4ef6\u5e76\u5728\u539f\u4f4d\u7f6e\u5efa\u7acb\u540c\u56fe\u6807\u542f\u52a8\u5feb\u6377\u65b9\u5f0f\u3002")

    def validate_settings(self) -> bool:
        if not self.selected_path.get():
            messagebox.showwarning("尚未选择", "请先拖入或选择一个 .lnk 快捷方式。", parent=self)
            return False
        try:
            if not 1 <= int(self.duration_minutes.get()) <= 180:
                raise ValueError
            if not 0 <= int(self.start_percent.get()) <= 99:
                raise ValueError
        except (ValueError, tk.TclError):
            messagebox.showwarning("设置无效", "时长需为 1–180 分钟，开始进度需为 0–99%。", parent=self)
            return False
        try:
            tk_hotkey(self.hotkey.get())
        except Exception:
            messagebox.showwarning("热键无效", "热键示例：ctrl+shift+q", parent=self)
            return False
        return True

    def selected_theme(self) -> str:
        value = self.theme.get()
        return self.theme_display_values.get(value, value if value in THEMES else "win10")

    def preview(self) -> None:
        if not self.validate_settings():
            return
        profile = {
            "duration_seconds": int(self.duration_minutes.get()) * 60,
            "start_percent": int(self.start_percent.get()),
            "hotkey": normalize_hotkey(self.hotkey.get()),
            "restore_mode": "manual",
            "theme": self.selected_theme(),
        }
        self.withdraw()
        try:
            BsodWindow(profile).run()
        finally:
            self.deiconify()
            self.lift()

    def take_over(self) -> None:
        if getattr(self, "_busy", False):
            return
        if not self.validate_settings():
            return
        path = Path(self.selected_path.get())
        if not path.exists() or not path.is_file():
            messagebox.showerror("\u6587\u4ef6\u4e0d\u5b58\u5728", "\u9009\u5b9a\u7684\u6587\u4ef6\u5df2\u4e0d\u5b58\u5728\uff0c\u6216\u4e0d\u662f\u666e\u901a\u6587\u4ef6\u3002", parent=self)
            return
        is_shortcut = path.suffix.lower() == ".lnk"
        if not is_shortcut:
            answer = messagebox.askokcancel(
                "\u666e\u901a\u6587\u4ef6\u6a21\u5f0f",
                f"\u5c06\u628a\u539f\u6587\u4ef6\u79fb\u52a8\u5230\u672c\u5de5\u5177\u7684\u5907\u4efd\u76ee\u5f55\uff0c\u5e76\u5728\u539f\u4f4d\u7f6e\u751f\u6210\uff1a\n{path.name}.lnk\n\n\u662f\u5426\u7ee7\u7eed\uff1f", parent=self)
            if not answer:
                return
        settings = {
            "duration_seconds": int(self.duration_minutes.get()) * 60,
            "start_percent": int(self.start_percent.get()),
            "hotkey": normalize_hotkey(self.hotkey.get()),
            "restore_mode": self.restore_mode.get(),
            "theme": self.selected_theme(),
        }
        self._set_busy(True, "\u6b63\u5728\u521b\u5efa\uff0c\u8bf7\u7a0d\u5019\u2026")
        threading.Thread(target=self._take_over_task, args=(path, is_shortcut, settings), daemon=True).start()

    def _take_over_task(self, path: Path, is_shortcut: bool, settings: dict) -> None:
        pythoncom.CoInitialize()
        profile_id = uuid.uuid4().hex
        backup_dir = BACKUPS_DIR / profile_id
        profile: dict | None = None
        moved_original = False
        proxy_created = False
        try:
            for existing in all_profiles():
                if Path(existing.get("original_path", existing.get("shortcut_path", ""))) == path and profile_is_active(existing):
                    raise RuntimeError("\u8fd9\u4e2a\u6587\u4ef6\u5df2\u7ecf\u5904\u4e8e\u84dd\u5c4f\u6a21\u5f0f\u3002")
            # Do not scan desktop icon positions; keep Explorer communication minimal.
            backup_dir.mkdir(parents=True)
            if is_shortcut:
                original = shortcut_data(path)
                backup = backup_dir / "original.lnk"
                shutil.copy2(path, backup)
                profile = {"id": profile_id, "source_kind": "shortcut", "shortcut_path": str(path), "original_path": str(path), "backup_path": str(backup), "original": original, "created_at": time.time(), "desktop_position": None, **settings}
                save_profile(profile)
                replace_shortcut(path, profile_id, original)
                proxy_created = True
                done = "\u5df2\u5907\u4efd\u5e76\u63a5\u7ba1\u5feb\u6377\u65b9\u5f0f"
            else:
                backup = backup_dir / ("original" + path.suffix)
                shutil.move(str(path), str(backup))
                moved_original = True
                profile = {"id": profile_id, "source_kind": "file", "shortcut_path": str(path.with_name(path.name + ".lnk")), "original_path": str(path), "backup_path": str(backup), "original": {"name": path.name}, "created_at": time.time(), "desktop_position": None, **settings}
                save_profile(profile)
                create_shadow_shortcut(path, profile_id, backup)
                proxy_created = True
                done = "\u5df2\u5907\u4efd\u6587\u4ef6\u5e76\u521b\u5efa\u540c\u56fe\u6807\u84dd\u5c4f\u5feb\u6377\u65b9\u5f0f"
            self.after(0, lambda: self._take_over_done(True, f"{done}\uff1a{path.name}", profile))
        except Exception as exc:
            rollback_errors = []
            try:
                if profile and profile.get("source_kind") == "shortcut":
                    backup = Path(profile["backup_path"])
                    destination = Path(profile["original_path"])
                    temp = destination.with_name(destination.name + ".bsod-rollback-tmp")
                    if backup.exists():
                        if temp.exists():
                            temp.unlink()
                        shutil.copy2(backup, temp)
                        os.replace(temp, destination)
                elif profile:
                    proxy = Path(profile["shortcut_path"])
                    if proxy.exists():
                        proxy.unlink()
                if moved_original:
                    backup = backup_dir / ("original" + path.suffix)
                    if backup.exists() and not path.exists():
                        shutil.move(str(backup), str(path))
            except OSError as rollback_exc:
                rollback_errors.append(str(rollback_exc))
            try:
                profile_file = PROFILES_DIR / f"{profile_id}.json"
                if profile_file.exists():
                    profile_file.unlink()
                if backup_dir.exists():
                    shutil.rmtree(backup_dir)
            except OSError as cleanup_exc:
                rollback_errors.append(str(cleanup_exc))
            error_text = f"\u521b\u5efa\u5931\u8d25\uff1a{exc}"
            if rollback_errors:
                error_text += "\n\n\u56de\u6eda\u65f6\u9047\u5230\u95ee\u9898\uff1a" + "\n".join(rollback_errors)
            else:
                error_text += "\n\n\u5df2\u5c1d\u8bd5\u6062\u590d\u539f\u6587\u4ef6\u3002"
            self.after(0, lambda text=error_text: self._take_over_done(False, text, None))
        finally:
            pythoncom.CoUninitialize()

    def _take_over_done(self, ok: bool, text: str, profile: dict | None) -> None:
        self._set_busy(False)
        self.status.set(text)
        self.refresh_profiles()
        if ok and profile:
            messagebox.showinfo("\u5b8c\u6210", f"{text}\n\u9000\u51fa\u70ed\u952e\uff1a{profile['hotkey']}", parent=self)
        elif not ok:
            messagebox.showerror("\u64cd\u4f5c\u5931\u8d25", f"\u6ca1\u6709\u4fee\u6539\u6210\u529f\uff1a\n{text}", parent=self)

    def _set_busy(self, busy: bool, text: str = "") -> None:
        self._busy = busy
        if text:
            self.status.set(text)

    def selected_profile(self) -> dict | None:
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("未选择项目", "请先从列表中选择一个快捷方式。", parent=self)
            return None
        return read_profile(selected[0])

    def restore_selected(self) -> None:
        if getattr(self, "_busy", False):
            return
        profile = self.selected_profile()
        if not profile:
            return
        state = profile_state(profile)
        if state == "restored":
            messagebox.showinfo("\u5df2\u6062\u590d", "\u8fd9\u4e2a\u9879\u76ee\u5df2\u7ecf\u6062\u590d\u3002", parent=self)
            return
        if not profile_can_restore(profile):
            messagebox.showerror(
                "\u65e0\u6cd5\u6062\u590d",
                "\u8fd9\u4e2a\u9879\u76ee\u5b58\u5728\u5f02\u5e38\uff1a" + profile_state(profile) + "\u3002\n\n\u8bf7\u5148\u5904\u7406\u6587\u4ef6\u51b2\u7a81\u6216\u8865\u56de\u5907\u4efd\u6587\u4ef6\u3002",
                parent=self,
            )
            return
        if not messagebox.askyesno("\u6062\u590d\u9879\u76ee", "\u5c06\u6062\u590d\u8fd9\u4e2a\u9879\u76ee\u5e76\u4fdd\u7559\u5907\u4efd\u3002\u662f\u5426\u7ee7\u7eed\uff1f", parent=self):
            return
        self._set_busy(True, "\u6b63\u5728\u6062\u590d\uff0c\u8bf7\u7a0d\u5019\u2026")
        threading.Thread(target=self._restore_task, args=(profile, False), daemon=True).start()

    def _restore_task(self, profile: dict, preserve: bool) -> None:
        result = restore_profile(profile, preserve)
        self.after(0, lambda: self._restore_done(result))

    def _restore_done(self, result: tuple[bool, str]) -> None:
        self._set_busy(False)
        ok, text = result
        self.status.set(text)
        self.refresh_profiles()
        if not ok:
            messagebox.showerror("\u6062\u590d\u5931\u8d25", text, parent=self)

    def restore_all(self) -> None:
        if getattr(self, "_busy", False):
            return
        profiles = all_profiles()
        candidates = [p for p in profiles if profile_can_restore(p)]
        blocked = [p for p in profiles if not p.get("restored_at") and not profile_can_restore(p)]
        if not candidates:
            messagebox.showinfo("\u6ca1\u6709\u53ef\u6062\u590d\u9879\u76ee", "\u5f53\u524d\u6ca1\u6709\u53ef\u81ea\u52a8\u6062\u590d\u7684\u9879\u76ee\u3002", parent=self)
            return
        warning = f"\u5c06\u5c1d\u8bd5\u6062\u590d {len(candidates)} \u4e2a\u9879\u76ee\u3002"
        if blocked:
            warning += f"\n\n\u53e6\u6709 {len(blocked)} \u4e2a\u9879\u76ee\u5b58\u5728\u5f02\u5e38\uff0c\u4e0d\u4f1a\u88ab\u5f3a\u5236\u8986\u76d6\u3002"
        if not messagebox.askyesno("\u6062\u590d\u5168\u90e8", warning + "\n\n\u662f\u5426\u7ee7\u7eed\uff1f", parent=self):
            return
        self._set_busy(True, "\u6b63\u5728\u6062\u590d\u5168\u90e8\u9879\u76ee\uff0c\u8bf7\u7a0d\u5019\u2026")
        threading.Thread(target=self._restore_all_task, args=(candidates,), daemon=True).start()

    def _restore_all_task(self, profiles: list[dict]) -> None:
        count = 0
        errors = []
        for profile in profiles:
            ok, text = restore_profile(profile)
            count += int(ok)
            if not ok:
                errors.append(text)
        self.after(0, lambda: self._restore_all_done(count, errors))

    def _restore_all_done(self, count: int, errors: list[str]) -> None:
        self._set_busy(False)
        self.status.set(f"\u5df2\u6062\u590d {count} \u4e2a\u5feb\u6377\u65b9\u5f0f\u3002" + (" \u6709\u9879\u76ee\u5931\u8d25\u3002" if errors else ""))
        self.refresh_profiles()
        if errors:
            messagebox.showerror("\u6062\u590d\u5931\u8d25", "\n".join(errors), parent=self)

    def refresh_profiles(self) -> None:
        labels = {"once": "\u4e00\u6b21\u540e\u6062\u590d", "after": "\u7ed3\u675f\u540e\u6062\u590d", "manual": "\u624b\u52a8\u6062\u590d"}
        state_labels = {
            "active": "\u84dd\u5c4f\u4e2d",
            "restored": "\u5df2\u6062\u590d",
            "missing_backup": "\u5907\u4efd\u4e22\u5931",
            "missing_proxy": "\u4ee3\u7406\u4e22\u5931",
            "conflict": "\u6709\u6587\u4ef6\u51b2\u7a81",
        }
        rows = []
        for profile in all_profiles():
            state = profile_state(profile)
            name = Path(profile.get("shortcut_path", "\u672a\u77e5")).name
            rows.append((profile["id"], name, profile.get("restore_mode"), state))
        signature = tuple(rows)
        if signature == getattr(self, "_profile_signature", None):
            return
        self._profile_signature = signature
        for item in self.tree.get_children():
            self.tree.delete(item)
        for profile_id, name, mode, state in rows:
            self.tree.insert(
                "", "end", iid=profile_id,
                values=(name, labels.get(mode, "\u672a\u77e5"), state_labels.get(state, "\u672a\u77e5")),
            )

    def check_startup_profiles(self) -> None:
        """Find incomplete records left by a crash or interrupted operation."""
        attention = [profile for profile in all_profiles() if profile_needs_attention(profile)]
        if not attention:
            return
        names = []
        for profile in attention[:5]:
            name = Path(profile.get("original_path", profile.get("shortcut_path", "\u672a\u77e5"))).name
            names.append(f"- {name}: {profile_state(profile)}")
        more = "" if len(attention) <= 5 else f"\n...\u8fd8\u6709 {len(attention) - 5} \u4e2a\u9879\u76ee"
        prompt = (
            f"\u68c0\u6d4b\u5230 {len(attention)} \u4e2a\u9879\u76ee\u53ef\u80fd\u5728\u4e0a\u6b21\u64cd\u4f5c\u4e2d\u672a\u5b8c\u6210\u3002\n\n"
            + "\n".join(names)
            + more
            + "\n\n\u662f\u5426\u7acb\u5373\u5c1d\u8bd5\u6062\u590d\uff1f"
        )
        if not messagebox.askyesno("\u68c0\u6d4b\u5230\u672a\u5b8c\u6210\u9879\u76ee", prompt, parent=self):
            self.status.set("\u5df2\u53d1\u73b0\u672a\u5b8c\u6210\u9879\u76ee\uff0c\u8bf7\u5728\u5217\u8868\u4e2d\u68c0\u67e5\u3002")
            return
        candidates = [profile for profile in attention if profile_can_restore(profile)]
        if not candidates:
            self.status.set("\u53d1\u73b0\u5f02\u5e38\u9879\u76ee\uff0c\u4f46\u6ca1\u6709\u53ef\u81ea\u52a8\u6062\u590d\u7684\u5907\u4efd\u3002")
            return
        self._set_busy(True, "\u6b63\u5728\u68c0\u67e5\u5e76\u6062\u590d\u672a\u5b8c\u6210\u9879\u76ee\uff0c\u8bf7\u7a0d\u5019\u2026")
        threading.Thread(target=self._startup_recovery_task, args=(candidates,), daemon=True).start()

    def _startup_recovery_task(self, profiles: list[dict]) -> None:
        count = 0
        errors = []
        for profile in profiles:
            ok, text = restore_profile(profile)
            count += int(ok)
            if not ok:
                errors.append(text)
        self.after(0, lambda: self._startup_recovery_done(count, errors))

    def _startup_recovery_done(self, count: int, errors: list[str]) -> None:
        self._set_busy(False)
        self.refresh_profiles()
        if errors:
            self.status.set(f"\u5df2\u6062\u590d {count} \u4e2a\u9879\u76ee\uff0c\u4ecd\u6709 {len(errors)} \u4e2a\u9879\u76ee\u9700\u8981\u5904\u7406\u3002")
            messagebox.showwarning("\u6062\u590d\u68c0\u67e5\u5b8c\u6210", "\n".join(errors), parent=self)
        else:
            self.status.set(f"\u542f\u52a8\u68c0\u67e5\u5b8c\u6210\uff0c\u5df2\u6062\u590d {count} \u4e2a\u9879\u76ee\u3002")

    def auto_refresh_profiles(self) -> None:
        self.refresh_profiles()
        if self.winfo_exists():
            self.after(1500, self.auto_refresh_profiles)

    def delete_profile_record(self, profile: dict) -> None:
        """Remove one profile and its private backup after restoration is safe."""
        profile_file = PROFILES_DIR / f"{profile['id']}.json"
        backup_dir = BACKUPS_DIR / profile["id"]
        # Both paths are derived from the generated profile id and remain inside
        # the application's data directories.
        if profile_file.exists():
            profile_file.unlink()
        if backup_dir.exists():
            shutil.rmtree(backup_dir)

    def delete_selected(self) -> None:
        profile = self.selected_profile()
        if not profile:
            return
        state = profile_state(profile)
        if state != "restored":
            if not profile_can_restore(profile):
                messagebox.showerror("\u65e0\u6cd5\u5220\u9664", "\u8fd9\u4e2a\u9879\u76ee\u5b58\u5728\u5f02\u5e38\uff0c\u8bf7\u5148\u89e3\u51b3\u6062\u590d\u95ee\u9898\u3002", parent=self)
                return
            if not messagebox.askyesno("\u9879\u76ee\u5c1a\u672a\u6062\u590d", "\u5220\u9664\u8bb0\u5f55\u524d\u4f1a\u5148\u81ea\u52a8\u6062\u590d\u539f\u6587\u4ef6\uff0c\u662f\u5426\u7ee7\u7eed\uff1f", parent=self):
                return
            ok, text = restore_profile(profile)
            if not ok:
                messagebox.showerror("\u6062\u590d\u5931\u8d25", f"\u5220\u9664\u5df2\u53d6\u6d88\uff1a\n{text}", parent=self)
                return
        elif not messagebox.askyesno("\u5220\u9664\u8bb0\u5f55", "\u5220\u9664\u540e\u5c06\u540c\u65f6\u5220\u9664\u8be5\u9879\u76ee\u7684\u5907\u4efd\u8bb0\u5f55\uff0c\u662f\u5426\u7ee7\u7eed\uff1f", parent=self):
            return
        try:
            self.delete_profile_record(profile)
            self.status.set("\u5df2\u5220\u9664\u8bb0\u5f55\uff1a{}".format(Path(profile.get("shortcut_path", "\u672a\u77e5")).name))
            self.refresh_profiles()
        except OSError as exc:
            messagebox.showerror("\u5220\u9664\u5931\u8d25", str(exc), parent=self)

    def delete_all_records(self) -> None:
        profiles = all_profiles()
        if not profiles:
            messagebox.showinfo("\u6ca1\u6709\u8bb0\u5f55", "\u5f53\u524d\u6ca1\u6709\u53ef\u5220\u9664\u7684\u8bb0\u5f55\u3002", parent=self)
            return
        pending = [p for p in profiles if not p.get("restored_at")]
        blocked = [p for p in pending if not profile_can_restore(p)]
        warning = f"\u5171\u6709 {len(profiles)} \u6761\u8bb0\u5f55\u3002\n\u5176\u4e2d {len(pending)} \u6761\u5c1a\u672a\u6062\u590d\u3002"
        if blocked:
            warning += f"\n\u6709 {len(blocked)} \u6761\u5b58\u5728\u5f02\u5e38\uff0c\u9700\u8981\u5148\u624b\u52a8\u5904\u7406\u3002"
        warning += "\n\n\u5220\u9664\u524d\u4f1a\u81ea\u52a8\u6062\u590d\u53ef\u6062\u590d\u7684\u9879\u76ee\uff0c\u7136\u540e\u5220\u9664\u8bb0\u5f55\u3002\n\n\u662f\u5426\u7ee7\u7eed\uff1f"
        if not messagebox.askyesno("\u6e05\u7406\u5168\u90e8\u8bb0\u5f55", warning, parent=self):
            return
        if blocked:
            messagebox.showerror("\u65e0\u6cd5\u6e05\u7406", "\u5b58\u5728\u672a\u89e3\u51b3\u7684\u5f02\u5e38\u9879\u76ee\uff0c\u4e3a\u907f\u514d\u4e22\u5931\u5907\u4efd\uff0c\u5df2\u53d6\u6d88\u6e05\u7406\u3002", parent=self)
            return
        errors = []
        for profile in pending:
            ok, text = restore_profile(profile)
            if not ok:
                errors.append(text)
        if errors:
            messagebox.showerror("\u6062\u590d\u5931\u8d25", "\u6709\u9879\u76ee\u672a\u80fd\u6062\u590d\uff0c\u5df2\u53d6\u6d88\u5220\u9664\u3002\n\n" + "\n".join(errors), parent=self)
            self.refresh_profiles()
            return
        try:
            for profile in profiles:
                self.delete_profile_record(profile)
            self.status.set("\u5df2\u6062\u590d\u5e76\u6e05\u7406\u5168\u90e8\u8bb0\u5f55\u3002")
            self.refresh_profiles()
        except OSError as exc:
            messagebox.showerror("\u5220\u9664\u5931\u8d25", str(exc), parent=self)

    def enable_drag_drop(self) -> None:
        """Use TkDND rather than replacing Tk's window procedure (stable with Explorer drops)."""
        self.drop.drop_target_register(DND_FILES)
        self.drop.dnd_bind("<<Drop>>", self.on_drop)

    def on_drop(self, event) -> str:
        try:
            paths = self.tk.splitlist(event.data)
            if paths:
                self.set_shortcut(Path(paths[0]))
        except Exception as exc:
            self.status.set(f"\u65e0\u6cd5\u8bfb\u53d6\u62d6\u5165\u7684\u6587\u4ef6\uff1a{exc}")
        return "break"


def launch(profile_id: str) -> int:
    try:
        profile = read_profile(profile_id)
    except Exception:
        return 1
    # The one-shot choice restores the original link before any full-screen UI is shown.
    if profile.get("restore_mode") == "once":
        restore_profile(profile)
    BsodWindow(profile).run()
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--launch")
    args, _ = parser.parse_known_args()
    if args.launch:
        return launch(args.launch)
    app = ShortcutManager()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())



"""Best-effort Windows desktop icon position reader/writer.

The desktop is Explorer's SysListView32 control. Coordinates are meaningful only
when Explorer's automatic arrangement is disabled.
"""
from __future__ import annotations
import ctypes
from ctypes import wintypes

user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

LVM_FIRST = 0x1000
LVM_GETITEMCOUNT = LVM_FIRST + 4
LVM_GETITEMPOSITION = LVM_FIRST + 16
LVM_SETITEMPOSITION32 = LVM_FIRST + 49
LVM_GETITEMTEXTW = LVM_FIRST + 115
LVIF_TEXT = 0x0001
PROCESS_VM_OPERATION = 0x0008
PROCESS_VM_READ = 0x0010
PROCESS_VM_WRITE = 0x0020
MEM_COMMIT = 0x1000
MEM_RESERVE = 0x2000
MEM_RELEASE = 0x8000
PAGE_READWRITE = 0x04

PTR = ctypes.c_ssize_t
LPTR = ctypes.c_void_p

class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

class LVITEMW(ctypes.Structure):
    _fields_ = [
        ("mask", wintypes.UINT), ("iItem", ctypes.c_int), ("iSubItem", ctypes.c_int),
        ("state", wintypes.UINT), ("stateMask", wintypes.UINT), ("pszText", LPTR),
        ("cchTextMax", ctypes.c_int), ("iImage", ctypes.c_int), ("lParam", PTR),
        ("iIndent", ctypes.c_int), ("iGroupId", ctypes.c_int), ("cColumns", wintypes.UINT),
        ("puColumns", LPTR), ("piColFmt", LPTR), ("iGroup", ctypes.c_int),
    ]

user32.FindWindowW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR]
user32.FindWindowW.restype = wintypes.HWND
user32.FindWindowExW.argtypes = [wintypes.HWND, wintypes.HWND, wintypes.LPCWSTR, wintypes.LPCWSTR]
user32.FindWindowExW.restype = wintypes.HWND
user32.SendMessageTimeoutW.argtypes = [
    wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM,
    wintypes.UINT, wintypes.UINT, ctypes.POINTER(ctypes.c_size_t),
]
user32.SendMessageTimeoutW.restype = wintypes.BOOL

SMTO_ABORTIFHUNG = 0x0002
SMTO_BLOCK = 0x0001
MESSAGE_TIMEOUT_MS = 80

def _send(hwnd, message, wparam=0, lparam=0) -> tuple[bool, int]:
    result = ctypes.c_size_t()
    ok = user32.SendMessageTimeoutW(
        hwnd, message, wparam, lparam,
        SMTO_ABORTIFHUNG | SMTO_BLOCK, MESSAGE_TIMEOUT_MS, ctypes.byref(result),
    )
    return bool(ok), int(result.value)
user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
user32.GetWindowThreadProcessId.restype = wintypes.DWORD
kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
kernel32.OpenProcess.restype = wintypes.HANDLE
kernel32.VirtualAllocEx.argtypes = [wintypes.HANDLE, LPTR, ctypes.c_size_t, wintypes.DWORD, wintypes.DWORD]
kernel32.VirtualAllocEx.restype = LPTR
kernel32.VirtualFreeEx.argtypes = [wintypes.HANDLE, LPTR, ctypes.c_size_t, wintypes.DWORD]
kernel32.ReadProcessMemory.argtypes = [wintypes.HANDLE, LPTR, LPTR, ctypes.c_size_t, ctypes.POINTER(ctypes.c_size_t)]
kernel32.ReadProcessMemory.restype = wintypes.BOOL
kernel32.WriteProcessMemory.argtypes = [wintypes.HANDLE, LPTR, LPTR, ctypes.c_size_t, ctypes.POINTER(ctypes.c_size_t)]
kernel32.WriteProcessMemory.restype = wintypes.BOOL
kernel32.CloseHandle.argtypes = [wintypes.HANDLE]


def _desktop_listview() -> int:
    progman = user32.FindWindowW("Progman", None)
    if progman:
        view = user32.FindWindowExW(progman, 0, "SHELLDLL_DefView", None)
        if view:
            return user32.FindWindowExW(view, 0, "SysListView32", "FolderView")
    found = wintypes.HWND()
    CALLBACK = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    @CALLBACK
    def visit(hwnd, _):
        view = user32.FindWindowExW(hwnd, 0, "SHELLDLL_DefView", None)
        if view:
            candidate = user32.FindWindowExW(view, 0, "SysListView32", "FolderView")
            if candidate:
                found.value = candidate
                return False
        return True
    user32.EnumWindows.argtypes = [CALLBACK, wintypes.LPARAM]
    user32.EnumWindows(visit, 0)
    return found.value


def _with_remote(fn):
    hwnd = _desktop_listview()
    if not hwnd:
        return None
    pid = wintypes.DWORD()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    process = kernel32.OpenProcess(PROCESS_VM_OPERATION | PROCESS_VM_READ | PROCESS_VM_WRITE, False, pid.value)
    if not process:
        return None
    remote = kernel32.VirtualAllocEx(process, None, 4096, MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE)
    if not remote:
        kernel32.CloseHandle(process)
        return None
    try:
        return fn(hwnd, process, remote)
    finally:
        kernel32.VirtualFreeEx(process, remote, 0, MEM_RELEASE)
        kernel32.CloseHandle(process)


def _read(process, address, obj) -> bool:
    received = ctypes.c_size_t()
    return bool(kernel32.ReadProcessMemory(process, address, ctypes.byref(obj), ctypes.sizeof(obj), ctypes.byref(received)))


def _write(process, address, obj) -> bool:
    written = ctypes.c_size_t()
    return bool(kernel32.WriteProcessMemory(process, address, ctypes.byref(obj), ctypes.sizeof(obj), ctypes.byref(written)))


def _titles_and_positions(wanted: set[str] | None = None):
    def collect(hwnd, process, remote):
        ok, result = _send(hwnd, LVM_GETITEMCOUNT)
        if not ok:
            return False
        count = result
        items = []
        text_address = remote + 1024
        wanted_set = {item.casefold() for item in (wanted or set())}
        for index in range(count):
            item = LVITEMW(mask=LVIF_TEXT, iItem=index, iSubItem=0, pszText=text_address, cchTextMax=512)
            if not _write(process, remote, item):
                continue
            if not _send(hwnd, LVM_GETITEMTEXTW, index, remote)[0]:
                continue
            text = ctypes.create_unicode_buffer(512)
            if not _read(process, text_address, text):
                continue
            # Do not ask Explorer for every icon's coordinates. For capture we
            # only need the matching item, which substantially reduces the
            # amount of cross-process ListView traffic.
            if wanted_set and text.value.casefold() not in wanted_set:
                continue
            point = POINT()
            if _send(hwnd, LVM_GETITEMPOSITION, index, remote)[0] and _read(process, remote, point):
                items.append((text.value, int(point.x), int(point.y)))
                if wanted_set:
                    return items
        return items
    return _with_remote(collect) or []


def capture(candidates: list[str]) -> dict | None:
    wanted = {x.casefold() for x in candidates if x}
    for title, x, y in _titles_and_positions(wanted):
        if title.casefold() in wanted:
            return {"title": title, "x": x, "y": y}
    return None


def restore(position: dict | None, candidates: list[str]) -> bool:
    if not position:
        return False
    wanted = {x.casefold() for x in candidates if x}
    # Setting a ListView item needs its index. Obtain it in one remote session.
    def set_item(hwnd, process, remote):
        ok, result = _send(hwnd, LVM_GETITEMCOUNT)
        if not ok:
            return []
        count = result
        text_address = remote + 1024
        for index in range(count):
            item = LVITEMW(mask=LVIF_TEXT, iItem=index, iSubItem=0, pszText=text_address, cchTextMax=512)
            if not _write(process, remote, item):
                continue
            _send(hwnd, LVM_GETITEMTEXTW, index, remote)
            text = ctypes.create_unicode_buffer(512)
            if _read(process, text_address, text) and text.value.casefold() in wanted:
                packed = (int(position["x"]) & 0xFFFF) | ((int(position["y"]) & 0xFFFF) << 16)
                return _send(hwnd, LVM_SETITEMPOSITION32, index, packed)[0]
        return False
    return bool(_with_remote(set_item))

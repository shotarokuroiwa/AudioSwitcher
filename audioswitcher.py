import csv
import io
import sys
import subprocess
import tempfile
from pathlib import Path
import win32con
import win32gui

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent

SOUND_VOLUME_VIEW = BASE_DIR / "SoundVolumeView.exe"


def get_output_devices():
    result = subprocess.run(
        [SOUND_VOLUME_VIEW, "/scomma", ""],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="ignore",
    )

    if not result.stdout.strip():
        return []

    reader = csv.DictReader(io.StringIO(result.stdout))
    devices = []

    for row in reader:
        if row.get("Type") == "Device" and row.get("Direction") == "Render":
            did = row.get("Command-Line Friendly ID") or row.get("Item ID")
            if did:
                devices.append({
                    "name": row.get("Device Name", ""),
                    "id": did.strip(),
                })

    return devices


def set_default_output(device_id):
    subprocess.run(
        [SOUND_VOLUME_VIEW, "/SetDefault", device_id, "0"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def switch_by_index(index):
    devices = get_output_devices()
    if index < len(devices):
        dev = devices[index]
        set_default_output(dev["id"])


def show_hotkey_help():
    devices = get_output_devices()
    lines = [
        "",
        "=== Audio Switcher Hotkeys ===",
        *[f"Ctrl+Alt+{i+1} : {d['name']}" for i, d in enumerate(devices[:9])],
        "==============================",
        "Ctrl+Alt+H : ショートカット一覧",
        "",
    ]

    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".txt", delete=False, encoding="utf-8"
    ) as f:
        f.write("\n".join(lines))
        path = f.name

    ps_cmd = (
        f"Get-Content -LiteralPath '{path}' -Encoding UTF8; "
        f"Write-Host ''; "
        f"Read-Host 'Enterキーで閉じる'"
    )

    subprocess.Popen(
        [
            "powershell",
            "-NoProfile",
            "-Command",
            ps_cmd,
        ],
        creationflags=subprocess.CREATE_NEW_CONSOLE,
    )


HOTKEYS = {
    1: lambda: switch_by_index(0),
    2: lambda: switch_by_index(1),
    3: lambda: switch_by_index(2),
    4: lambda: switch_by_index(3),
    5: lambda: switch_by_index(4),
    6: lambda: switch_by_index(5),
    7: lambda: switch_by_index(6),
    8: lambda: switch_by_index(7),
    9: lambda: switch_by_index(8),
    10: show_hotkey_help,
}


def main():
    # 登録
    for hid, vk in enumerate(
        [ord(str(i)) for i in range(1, 10)] + [ord("H")], start=1
    ):
        win32gui.RegisterHotKey(
            None,
            hid,
            win32con.MOD_CONTROL | win32con.MOD_ALT,
            vk,
        )

    show_hotkey_help()

    try:
        while True:
            msg = win32gui.GetMessage(None, 0, 0)
            if msg[1][1] == win32con.WM_HOTKEY:
                hotkey_id = msg[1][2]
                action = HOTKEYS.get(hotkey_id)
                if action:
                    action()
    finally:
        for hid in HOTKEYS:
            win32gui.UnregisterHotKey(None, hid)

if __name__ == "__main__":
    main()

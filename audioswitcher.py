import csv
import io
import sys
import tempfile
import time
import subprocess
import keyboard
from pathlib import Path

if getattr(sys, "frozen", False):
    _base_dir = Path(sys.executable).parent
else:
    _base_dir = Path(__file__).parent
SOUND_VOLUME_VIEW = _base_dir / "SoundVolumeView.exe"


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

    for data in reader:
        if data.get("Type") == "Device" and data.get("Direction") == "Render":
            device_id = data.get("Command-Line Friendly ID") or data.get("Item ID")
            if not device_id:
                continue
            devices.append({
                "name": data.get("Device Name", ""),
                "id": device_id.strip(),
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
        print(f"切り替え → {dev['name']}")


def show_hotkey_help():
    devices = get_output_devices()
    lines = [
        "",
        "=== Audio Switcher Hotkeys ===",
        *[f"Ctrl+Alt+{i+1} : {d['name']}" for i, d in enumerate(devices[:9])],
        "==============================",
        "Ctrl+Alt+H : ショートカット一覧を表示",
        "",
    ]
    text = "\n".join(lines)

    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".txt", delete=False, encoding="utf-8"
    ) as f:
        f.write(text)
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

def main():
    if not SOUND_VOLUME_VIEW.exists():
        msg = "SoundVolumeView.exe が見つかりません"
        if getattr(sys, "frozen", False):
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, msg, "Audio Switcher", 0x10)
        else:
            print(msg)
        return

    for i in range(9):
        keyboard.add_hotkey(
            f"ctrl+alt+{i+1}",
            switch_by_index,
            args=(i,),
        )

    keyboard.add_hotkey("ctrl+alt+h", show_hotkey_help)

    show_hotkey_help()

    while True:
        time.sleep(1)


if __name__ == "__main__":
    main()

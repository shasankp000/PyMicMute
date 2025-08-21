# MicMuteApp: tray + GUI + hotkey + autostart
# Requirements:
#   pip install customtkinter pycaw comtypes keyboard pystray pillow winshell
# Build:
#   pyinstaller --onefile --noconsole --icon mic.ico --add-data "mic_on.ico;." --add-data "mic_off.ico;." --add-data "mic.ico;." app.py

import os
import sys
import json
import threading
import traceback
import atexit
from pathlib import Path
import webbrowser
try:
    from winotify import Notification, audio
except Exception:
    Notification = None
    audio = None

import customtkinter as ctk
import keyboard
import pystray
from PIL import Image
from ctypes import POINTER, cast, c_void_p, c_int, c_ulong, c_wchar_p, windll
from ctypes import wintypes
from comtypes import CLSCTX_ALL, GUID, HRESULT, IUnknown
from comtypes import CoCreateInstance
from comtypes import COMMETHOD

from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# Core Audio enums (avoid pycaw.constants version differences)
eRender = 0
eCapture = 1
eAll = 2
eConsole = 0
eMultimedia = 1
eCommunications = 2
DEVICE_STATE_ACTIVE = 0x00000001

# MMDeviceEnumerator GUID and minimal COM interfaces (avoid comtypes.gen)
CLSID_MMDeviceEnumerator_GUID = GUID('{BCDE0395-E52F-467C-8E3D-C4579291692E}')

class IMMDevice(IUnknown):
    _iid_ = GUID('{D666063F-1587-4E43-81F1-B948E807363F}')
    _methods_ = [
        COMMETHOD([], HRESULT, 'Activate',
                  (['in'], GUID, 'iid'),
                  (['in'], c_ulong, 'dwClsCtx'),
                  (['in'], c_void_p, 'pActivationParams'),
                  (['out'], POINTER(c_void_p), 'ppInterface')),
        COMMETHOD([], HRESULT, 'OpenPropertyStore',
                  (['in'], c_ulong, 'stgmAccess'),
                  (['out'], POINTER(c_void_p), 'ppProperties')),
        COMMETHOD([], HRESULT, 'GetId',
                  (['out'], POINTER(c_void_p), 'ppstrId')),
        COMMETHOD([], HRESULT, 'GetState',
                  (['out'], POINTER(c_ulong), 'pdwState')),
    ]

class IMMDeviceEnumerator(IUnknown):
    _iid_ = GUID('{A95664D2-9614-4F35-A746-DE8DB63617E6}')
    _methods_ = [
        COMMETHOD([], HRESULT, 'EnumAudioEndpoints',
                  (['in'], c_int, 'dataFlow'),
                  (['in'], c_ulong, 'dwStateMask'),
                  (['out'], POINTER(c_void_p), 'ppDevices')),
        COMMETHOD([], HRESULT, 'GetDefaultAudioEndpoint',
                  (['in'], c_int, 'dataFlow'),
                  (['in'], c_int, 'role'),
                  (['out'], POINTER(POINTER(IMMDevice)), 'ppDevice')),
        COMMETHOD([], HRESULT, 'GetDevice',
                  (['in'], c_wchar_p, 'pwstrId'),
                  (['out'], POINTER(POINTER(IMMDevice)), 'ppDevice')),
        COMMETHOD([], HRESULT, 'RegisterEndpointNotificationCallback',
                  (['in'], c_void_p, 'pClient')),
        COMMETHOD([], HRESULT, 'UnregisterEndpointNotificationCallback',
                  (['in'], c_void_p, 'pClient')),
    ]

CONFIG_DIR = Path(os.getenv("APPDATA", str(Path.home() / "AppData" / "Roaming"))) / "MicMuteApp"
CONFIG_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_FILE = CONFIG_DIR / "config.json"

DEFAULT_CONFIG = {
    "hotkey": "ctrl+alt+m",
    "device_id": None,   # IMMDevice ID string
    "autostart": False,
    "appearance": "dark",
    "last_muted": None  # persisted last mic state (True/False)
}

def load_config():
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return {**DEFAULT_CONFIG, **data}
        except Exception:
            print("Failed to load config, using defaults.", file=sys.stderr)
    return DEFAULT_CONFIG.copy()

def save_config(cfg):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except Exception as e:
        print("Failed to save config:", e, file=sys.stderr)

CONFIG = load_config()

# ------------- Core Audio Helpers -------------
def list_input_devices():
    # Returns list of (IMMDevice, friendly_name, device_id)
    devices = []
    try:
        enumerator = CoCreateInstance(CLSID_MMDeviceEnumerator_GUID, interface=IMMDeviceEnumerator, clsctx=CLSCTX_ALL)
        all_devs = AudioUtilities.GetAllDevices()
        get_flow = AudioUtilities.GetEndpointDataFlow
        for d in all_devs:
            try:
                flow = get_flow(d.id)
            except Exception:
                continue
            if str(flow) == 'eCapture' or flow == 1:
                name = getattr(d, 'FriendlyName', 'Unknown')
                try:
                    imm_dev = enumerator.GetDevice(d.id)
                except Exception:
                    continue
                try:
                    state = imm_dev.GetState()
                    if state != DEVICE_STATE_ACTIVE:
                        continue
                except Exception:
                    pass
                devices.append((imm_dev, name, d.id))
    except Exception:
        pass
    if not devices:
        # Fallback: include default device if available
        try:
            imm = get_default_input_device()
            devices.append((imm, 'Default microphone', None))
        except Exception:
            pass
    return devices

def get_device_by_id(dev_id):
    if not dev_id:
        return None
    try:
        enumerator = CoCreateInstance(CLSID_MMDeviceEnumerator_GUID, interface=IMMDeviceEnumerator, clsctx=CLSCTX_ALL)
        return enumerator.GetDevice(dev_id)
    except Exception:
        return None

def get_default_input_device():
    enumerator = CoCreateInstance(CLSID_MMDeviceEnumerator_GUID, interface=IMMDeviceEnumerator, clsctx=CLSCTX_ALL)
    return enumerator.GetDefaultAudioEndpoint(eCapture, eMultimedia)

def activate_endpoint_volume(device):
    return cast(device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None), POINTER(IAudioEndpointVolume))

def resolve_mic_endpoint():
    # priority: configured device -> default device
    dev = None
    if CONFIG.get("device_id"):
        dev = get_device_by_id(CONFIG["device_id"])
    if not dev:
        try:
            dev = get_default_input_device()
        except Exception:
            pass
    if not dev:
        devs = list_input_devices()
        if devs:
            dev = devs[0][0]
    return dev

MIC_DEVICE = resolve_mic_endpoint()
AUDIO_VOL = activate_endpoint_volume(MIC_DEVICE) if MIC_DEVICE else None

def iter_capture_volumes():
    for dev, _name, _id in list_input_devices():
        try:
            yield activate_endpoint_volume(dev)
        except Exception:
            continue

def is_muted():
    # Consider mic muted only if all active capture endpoints are muted
    states = []
    for vol in iter_capture_volumes():
        try:
            states.append(vol.GetMute())
        except Exception:
            continue
    if not states:
        return None
    return 1 if all(s == 1 for s in states) else 0

def set_muted(m):
    changed = False
    for vol in iter_capture_volumes():
        try:
            vol.SetMute(1 if m else 0, None)
            changed = True
        except Exception:
            continue
    if changed:
        # persist last known state
        CONFIG["last_muted"] = bool(m)
        save_config(CONFIG)
        # toast notification (single)
        try:
            if Notification is not None:
                icon_path = str(ICON_DIR / ("mic_off.ico" if m else "mic_on.ico"))
                toast = Notification(app_id="MicMuteApp",
                                     title="Microphone",
                                     msg=("Muted" if m else "Unmuted"),
                                     icon=icon_path)
                if audio is not None:
                    toast.set_audio(audio.Default, loop=False)
                toast.show()
        except Exception:
            pass

def toggle_mic():
    if AUDIO_VOL:
        set_muted(not is_muted())

# ------------- Autostart (Startup Shortcut) -------------
def startup_dir():
    return Path(os.getenv("APPDATA")) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"

def startup_shortcut_path():
    return startup_dir() / "MicMuteApp.lnk"

def exe_target_for_shortcut():
    # If frozen (PyInstaller), use the exe. Otherwise run pythonw with script.
    if getattr(sys, "frozen", False):
        return sys.executable
    else:
        # Use pythonw to avoid console window
        pyw = Path(sys.executable).with_name("pythonw.exe")
        if pyw.exists():
            return f'"{pyw}" "{Path(__file__).resolve()}"'
        return f'"{sys.executable}" "{Path(__file__).resolve()}"'

def add_to_startup():
    try:
        import winshell
        target = exe_target_for_shortcut()
        with winshell.shortcut(str(startup_shortcut_path())) as link:
            if getattr(sys, "frozen", False):
                link.path = sys.executable
                link.arguments = ""
            else:
                # When not frozen, link.path must be pythonw.exe and arguments = script path
                # exe_target_for_shortcut already includes both, but winshell separates fields
                parts = target.strip('"').split('" "')
                if len(parts) == 2:
                    link.path = parts[0]
                    link.arguments = parts[1]
                else:
                    link.path = sys.executable
                    link.arguments = str(Path(__file__).resolve())
            link.working_directory = str(Path(__file__).parent.resolve())
            link.description = "MicMuteApp"
        return True
    except Exception as e:
        print("Autostart failed:", e, file=sys.stderr)
        return False

def remove_from_startup():
    try:
        p = startup_shortcut_path()
        if p.exists():
            p.unlink()
        return True
    except Exception as e:
        print("Remove autostart failed:", e, file=sys.stderr)
        return False

def is_in_startup():
    return startup_shortcut_path().exists()

# ------------- Tray -------------
ICON_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
ICON_ON = Image.open(str(ICON_DIR / "mic_on.ico"))
ICON_OFF = Image.open(str(ICON_DIR / "mic_off.ico"))
ICON_APP = Image.open(str(ICON_DIR / "mic.ico"))

tray = None
hotkey_handle = None
SHUTDOWN_EVENT = threading.Event()
SINGLETON_HANDLE = None

def ensure_single_instance():
    # Named mutex to prevent multiple instances
    # Returns True if this is the only instance
    global SINGLETON_HANDLE
    try:
        name = "Local\\MicMuteApp"
        CreateMutex = windll.kernel32.CreateMutexW
        CreateMutex.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
        CreateMutex.restype = wintypes.HANDLE
        # Request initial ownership to avoid races
        handle = CreateMutex(None, True, name)
        if not handle:
            return False
        err = windll.kernel32.GetLastError()
        ERROR_ALREADY_EXISTS = 183
        if err == ERROR_ALREADY_EXISTS:
            windll.kernel32.CloseHandle(handle)
            return False
        SINGLETON_HANDLE = handle
        return True
    except Exception:
        return True

def release_single_instance():
    global SINGLETON_HANDLE
    try:
        if SINGLETON_HANDLE:
            windll.kernel32.CloseHandle(SINGLETON_HANDLE)
            SINGLETON_HANDLE = None
    except Exception:
        pass

atexit.register(release_single_instance)

def update_tray_icon():
    try:
        if tray:
            muted = is_muted()
            tray.icon = ICON_OFF if muted else ICON_ON
            tray.title = f"Mic: {'MUTED' if muted else 'ON'}"
    except Exception as e:
        print("Tray update failed:", e, file=sys.stderr)

def on_quit(icon, item):
    try:
        if hotkey_handle:
            keyboard.remove_hotkey(hotkey_handle)
    except Exception:
        pass
    try:
        icon.stop()
    finally:
        SHUTDOWN_EVENT.set()

def tray_thread():
    global tray
    menu = pystray.Menu(
        pystray.MenuItem("Settings", lambda icon, item: open_settings_window()),
        pystray.MenuItem("Toggle Mic", lambda icon, item: (toggle_mic(), update_tray_icon())),
        pystray.MenuItem("Quit", on_quit)
    )
    tray = pystray.Icon("MicMuteApp", ICON_APP, "Mic: ?", menu)
    update_tray_icon()
    tray.run()

# ------------- Hotkey Registration -------------
def register_hotkey(hotkey):
    global hotkey_handle
    try:
        if hotkey_handle:
            keyboard.remove_hotkey(hotkey_handle)
        hotkey_handle = keyboard.add_hotkey(hotkey, lambda: (toggle_mic(), update_tray_icon()))
        return True
    except Exception as e:
        print("Hotkey registration failed:", e, file=sys.stderr)
        return False

# ------------- GUI (customtkinter) -------------
def open_settings_window():
    # Create a new window each invocation (keeps tray responsive)
    ctk.set_appearance_mode(CONFIG.get("appearance", "dark"))
    ctk.set_default_color_theme("blue")
    win = ctk.CTk()
    win.title("MicMute Settings")
    win.geometry("420x480")
    try:
        win.minsize(420, 480)
    except Exception:
        pass
    try:
        win.iconbitmap(str(ICON_DIR / "mic.ico"))
    except Exception:
        pass

    # Scrollable content container
    scroll = ctk.CTkScrollableFrame(win, width=400, height=440)
    scroll.pack(fill="both", expand=True, padx=10, pady=10)

    # Header and link
    title_lbl = ctk.CTkLabel(scroll, text="PyMicMute v1.0.1", font=("Segoe UI", 22, "bold"))
    title_lbl.pack(pady=(0, 4))

    def open_author_link(_evt=None):
        try:
            webbrowser.open("https://www.github.com/shasankp000")
        except Exception:
            pass

    link_lbl = ctk.CTkLabel(
        scroll,
        text="by shasankp000 : www.github.com/shasankp000",
        text_color="#1e90ff",
        cursor="hand2",
    )
    link_lbl.bind("<Button-1>", open_author_link)
    link_lbl.pack(pady=(0, 10))

    # Status
    status = ctk.CTkLabel(scroll, text="Mic: ?", font=("Segoe UI", 20))
    status.pack(pady=12)

    def refresh_status():
        m = is_muted()
        status.configure(text=f"Mic: {'MUTED' if m else 'ON'}", text_color=("red" if m else "green"))
        update_tray_icon()

    # Toggle button
    toggle_btn = ctk.CTkButton(scroll, text="Toggle Mic", command=lambda: (toggle_mic(), refresh_status()))
    toggle_btn.pack(pady=8)

    # Device dropdown
    devices = list_input_devices()
    device_names = [name for (_dev, name, _id) in devices]
    selected_idx = 0
    if CONFIG.get("device_id"):
        for i, (_dev, _name, dev_id) in enumerate(devices):
            if dev_id == CONFIG["device_id"]:
                selected_idx = i
                break

    device_label = ctk.CTkLabel(scroll, text="Input device")
    device_label.pack(pady=(16, 0))

    device_combo = ctk.CTkComboBox(scroll, values=device_names, width=300)
    if device_names:
        if 0 <= selected_idx < len(device_names):
            device_combo.set(device_names[selected_idx])
        else:
            device_combo.set(device_names[0])
    device_combo.pack(pady=6)

    def apply_device():
        global MIC_DEVICE, AUDIO_VOL
        idx = device_names.index(device_combo.get())
        MIC_DEVICE = devices[idx][0]
        AUDIO_VOL = activate_endpoint_volume(MIC_DEVICE)
        CONFIG["device_id"] = devices[idx][2]
        save_config(CONFIG)
        # Re-apply last state to the newly selected device
        last = CONFIG.get("last_muted")
        if last is not None:
            try:
                set_muted(last)
            except Exception:
                pass
        refresh_status()

    device_apply = ctk.CTkButton(scroll, text="Use Selected Device", command=apply_device)
    device_apply.pack(pady=6)

    # Hotkey rebinding
    hotkey_label = ctk.CTkLabel(scroll, text=f"Hotkey: {CONFIG['hotkey']}")
    hotkey_label.pack(pady=(14, 6))

    def rebind_hotkey():
        hotkey_label.configure(text="Press new hotkey... (Esc to cancel)")
        try:
            combo = keyboard.read_hotkey(suppress=False)  # allow Esc to cancel
            if combo:
                if register_hotkey(combo):
                    CONFIG["hotkey"] = combo
                    save_config(CONFIG)
                    hotkey_label.configure(text=f"Hotkey: {combo}")
        except Exception:
            traceback.print_exc()
        finally:
            refresh_status()

    rebind_btn = ctk.CTkButton(scroll, text="Rebind Hotkey", command=rebind_hotkey)
    rebind_btn.pack(pady=6)

    # Autostart checkbox
    autostart_var = ctk.BooleanVar(value=is_in_startup())

    def toggle_autostart():
        want = autostart_var.get()
        ok = add_to_startup() if want else remove_from_startup()
        if ok:
            CONFIG["autostart"] = want
            save_config(CONFIG)

    autostart_chk = ctk.CTkCheckBox(scroll, text="Run at startup", variable=autostart_var, command=toggle_autostart)
    autostart_chk.pack(pady=10)

    # Appearance
    def set_appearance(mode):
        CONFIG["appearance"] = mode
        save_config(CONFIG)
        ctk.set_appearance_mode(mode)

    appearance_frame = ctk.CTkFrame(scroll)
    appearance_frame.pack(pady=10)
    ctk.CTkLabel(appearance_frame, text="Appearance").pack(side="left", padx=8)
    ctk.CTkButton(appearance_frame, text="Light", width=64, command=lambda: set_appearance("light")).pack(side="left", padx=4)
    ctk.CTkButton(appearance_frame, text="Dark", width=64, command=lambda: set_appearance("dark")).pack(side="left", padx=4)
    ctk.CTkButton(appearance_frame, text="System", width=64, command=lambda: set_appearance("system")).pack(side="left", padx=4)

    # Close btn
    close_btn = ctk.CTkButton(scroll, text="Close", fg_color="#444", hover_color="#333", command=win.destroy)
    close_btn.pack(pady=8)

    refresh_status()
    win.mainloop()


def main():
    # Enforce single instance
    if not ensure_single_instance():
        # Another instance is already running; notify and exit
        try:
            if Notification is not None:
                toast = Notification(app_id="MicMuteApp",
                                     title="MicMuteApp",
                                     msg="Already running",
                                     icon=str(ICON_DIR / "mic.ico"))
                toast.show()
        except Exception:
            pass
        return
    # Register hotkey from config
    register_hotkey(CONFIG["hotkey"])

    # Start tray thread
    t = threading.Thread(target=tray_thread, daemon=True)
    t.start()

    # Apply last known mic state on startup (if available)
    try:
        last = CONFIG.get("last_muted")
        if last is not None:
            set_muted(last)
            update_tray_icon()
    except Exception:
        pass

    # Keep the process alive for hotkey listening until quit
    try:
        SHUTDOWN_EVENT.wait()
    except KeyboardInterrupt:
        SHUTDOWN_EVENT.set()

if __name__ == "__main__":
    main()

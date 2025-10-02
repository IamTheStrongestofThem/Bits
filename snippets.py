import os
import subprocess
import difflib
import shlex
import psutil
import pygetwindow as gw
import ctypes
import json
from pathlib import Path
import platform

CONFIG_PATH = Path(os.getenv("APPDATA")) / "bits_config.json"

DEFAULT_ALIAS = {
    "chrome": r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    "brave": r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe",
    "edge": r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    "notepad": "notepad",
    "explorer": "explorer",
    "terminal": "cmd"
}

COMMANDS = [
    "open", "close", "closeall", "minimize", "maximize",
    "list", "shutdown", "restart", "lock", "sleep",
    "aliases", "addalias", "removealias", "help", "exit"
]

# ---------- Config ----------
def load_config():
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            return {"aliases": DEFAULT_ALIAS}
    else:
        cfg = {"aliases": DEFAULT_ALIAS}
        save_config(cfg)
        return cfg

def save_config(cfg):
    try:
        CONFIG_PATH.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
    except Exception as e:
        print("Warning: couldn't save config:", e)

cfg = load_config()

# ---------- Helpers ----------
def fuzzy_command(cmd):
    cmd_low = cmd.lower()
    if cmd_low in COMMANDS:
        return cmd_low
    alt = difflib.get_close_matches(cmd_low, COMMANDS, n=1, cutoff=0.6)
    return alt[0] if alt else None

def fuzzy_alias(name):
    names = list(cfg.get("aliases", {}).keys())
    if name in names:
        return name
    alt = difflib.get_close_matches(name, names, n=1, cutoff=0.6)
    return alt[0] if alt else None

# ---------- Window Functions ----------
def list_windows():
    wins = gw.getAllTitles()
    return [w for w in wins if w.strip()]

def find_window_titles(keyword):
    titles = list_windows()
    matches = [t for t in titles if keyword.lower() in t.lower()]
    if not matches:
        matches = difflib.get_close_matches(keyword, titles, n=5, cutoff=0.5)
    return matches

def close_windows(keyword, all_matches=False):
    matches = find_window_titles(keyword)
    if not matches:
        print(f"No windows match '{keyword}'")
        return
    if not all_matches and len(matches) > 1:
        print("Multiple matches:", matches)
        print("Use 'closeall' for all or refine your keyword.")
        return
    for title in matches:
        try:
            win = gw.getWindowsWithTitle(title)[0]
            win.close()
            print("Closed:", title)
        except Exception:
            print("Failed to close:", title)

def minimize_windows(keyword):
    matches = find_window_titles(keyword)
    for title in matches:
        try:
            win = gw.getWindowsWithTitle(title)[0]
            win.minimize()
            print("Minimized:", title)
        except Exception:
            print("Failed to minimize:", title)

def maximize_windows(keyword):
    matches = find_window_titles(keyword)
    for title in matches:
        try:
            win = gw.getWindowsWithTitle(title)[0]
            win.maximize()
            print("Maximized:", title)
        except Exception:
            print("Failed to maximize:", title)

# ---------- System Functions ----------
def open_app(name_or_path):
    alias = fuzzy_alias(name_or_path) or name_or_path
    cmd = cfg.get("aliases", {}).get(alias, alias)
    try:
        if os.path.isfile(cmd):
            os.startfile(cmd)
        else:
            subprocess.Popen(cmd, shell=True)
        print(f"Opening: {alias}")
    except Exception as e:
        print("Failed to open:", e)

def shutdown_system():
    subprocess.run("shutdown /s /t 0", shell=True)

def restart_system():
    subprocess.run("shutdown /r /t 0", shell=True)

def lock_system():
    ctypes.windll.user32.LockWorkStation()

def sleep_system():
    ctypes.windll.PowrProf.SetSuspendState(0, 1, 0)

# ---------- Alias Mgmt ----------
def show_aliases():
    for k, v in cfg.get("aliases", {}).items():
        print(f"{k} -> {v}")

def add_alias(name, command):
    cfg.setdefault("aliases", {})[name] = command
    save_config(cfg)
    print(f"Alias added: {name} -> {command}")

def remove_alias(name):
    if name in cfg.get("aliases", {}):
        cfg["aliases"].pop(name)
        save_config(cfg)
        print("Removed alias:", name)
    else:
        print("Alias not found:", name)

# ---------- Help ----------
def print_help():
    print("""
Commands:
  open <app>             Open an app or alias
  close <win>            Close a window
  closeall <kw>          Close all matching windows
  minimize <win>         Minimize window
  maximize <win>         Maximize window
  list                   List windows
  shutdown               Shutdown PC
  restart                Restart PC
  lock                   Lock PC
  sleep                  Sleep PC
  aliases                Show aliases
  addalias <n> <cmd>     Add alias
  removealias <n>        Remove alias
  help                   Show this help
  exit                   Quit Bits
""")

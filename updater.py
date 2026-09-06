"""
updater.py — GitHub-Releases-based auto-updater for ClipboardManagerPro.

Flow:
    1. Read the running app's version from version.py.
    2. Ask GitHub's public Releases API for the latest published release
       of the configured repo.
    3. If it's newer than the running version, prompt the user, download
       the installer asset attached to that release, launch it, and exit
       so the installer can overwrite the running exe.

All network work happens on background threads; every Tkinter call
(messagebox, launching the installer) is marshaled back onto the main
thread via window.after(), since Tkinter itself is not thread-safe.

Silent checks (app startup) never show a dialog unless an update is
actually found. Manual checks (a "Check for Updates" button) always show
a result, including "you're up to date" and network errors.
"""

import os
import re
import subprocess
import tempfile
import threading
from tkinter import messagebox

import requests

from version import __version__ as CURRENT_VERSION

# ---------------------------------------------------------------------
# Configuration — set these to your actual GitHub username/repo
# ---------------------------------------------------------------------
GITHUB_OWNER = "your-github-username"
GITHUB_REPO = "clipboard-manager-pro"

# Substring used to find the installer asset among a release's files.
# Matches the OutputBaseFilename pattern in installer.iss below.
INSTALLER_ASSET_HINT = "Setup"

API_URL = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
REQUEST_TIMEOUT = 6


def _parse_version(v):
    """'v1.2.3' or '1.2.3' -> (1, 2, 3) for tuple comparison."""
    v = v.strip().lstrip("vV")
    parts = re.findall(r"\d+", v)
    return tuple(int(p) for p in parts) if parts else (0,)


def _fetch_latest_release():
    try:
        resp = requests.get(
            API_URL, timeout=REQUEST_TIMEOUT,
            headers={"Accept": "application/vnd.github+json"}
        )
        if resp.status_code != 200:
            return None
        data = resp.json()
        if data.get("draft") or data.get("prerelease"):
            return None
        return data
    except requests.RequestException:
        return None


def _find_installer_asset(release_json):
    for asset in release_json.get("assets", []):
        name = asset.get("name", "")
        if INSTALLER_ASSET_HINT.lower() in name.lower() and name.lower().endswith(".exe"):
            return asset
    return None


def _download_asset(asset):
    url = asset["browser_download_url"]
    dest = os.path.join(tempfile.gettempdir(), asset["name"])
    with requests.get(url, stream=True, timeout=30) as resp:
        resp.raise_for_status()
        with open(dest, "wb") as f:
            for chunk in resp.iter_content(chunk_size=262144):
                if chunk:
                    f.write(chunk)
    return dest


def check_for_updates_async(window, silent=True):
    """Kick off a background version check. Call once at startup with
    silent=True, or from a manual 'Check for Updates' action with
    silent=False."""
    def worker():
        release = _fetch_latest_release()
        window.after(0, lambda: _on_release_fetched(window, release, silent))

    threading.Thread(target=worker, daemon=True).start()


def _on_release_fetched(window, release, silent):
    if release is None:
        if not silent:
            messagebox.showinfo(
                "Check for Updates",
                "Couldn't reach GitHub to check for updates.\n"
                "Check your internet connection and try again."
            )
        return

    latest_tag = release.get("tag_name", "")
    if _parse_version(latest_tag) <= _parse_version(CURRENT_VERSION):
        if not silent:
            messagebox.showinfo("Check for Updates", f"You're up to date (v{CURRENT_VERSION}).")
        return

    asset = _find_installer_asset(release)
    if asset is None:
        if not silent:
            messagebox.showwarning(
                "Check for Updates",
                f"A new version ({latest_tag}) is available, but no installer\n"
                "asset was found on the release. Download it manually from:\n"
                f"https://github.com/{GITHUB_OWNER}/{GITHUB_REPO}/releases"
            )
        return

    notes = (release.get("body") or "").strip()
    notes_preview = (notes[:400] + "…") if len(notes) > 400 else notes
    prompt = (
        f"A new version is available: {latest_tag} (you have v{CURRENT_VERSION}).\n\n"
        f"{notes_preview}\n\n"
        "Download and install it now? The app will close during installation."
    )
    if not messagebox.askyesno("Update Available", prompt):
        return

    _download_and_install(window, asset)


def _download_and_install(window, asset):
    def worker():
        try:
            installer_path = _download_asset(asset)
        except Exception as e:
            window.after(0, lambda: messagebox.showerror(
                "Update Failed",
                f"Couldn't download the update:\n{e}\n\n"
                f"You can download it manually from:\n"
                f"https://github.com/{GITHUB_OWNER}/{GITHUB_REPO}/releases"
            ))
            return
        window.after(0, lambda: _launch_installer(installer_path))

    threading.Thread(target=worker, daemon=True).start()


def _launch_installer(installer_path):
    try:
        subprocess.Popen([installer_path], close_fds=True)
    except Exception as e:
        messagebox.showerror("Update Failed", f"Couldn't launch the installer:\n{e}")
        return
    os._exit(0)  # Hard exit so the running exe isn't locked when the installer overwrites it


if __name__ == "__main__":
    # Manual smoke test: python updater.py — always shows a result dialog
    import tkinter as tk
    root = tk.Tk()
    root.withdraw()
    check_for_updates_async(root, silent=False)
    root.mainloop()
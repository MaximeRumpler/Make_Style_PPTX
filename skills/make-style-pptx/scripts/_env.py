"""Use the skill's private Python environment when the system Python lacks the packages.

Import this module first in every script that needs python-pptx, lxml or Pillow:

    import _env  # noqa: F401  (must stay the first project import)

On systems whose Python is "externally managed" (PEP 668: recent Ubuntu, Debian, Fedora,
Homebrew), pip refuses to install into it. `python3 scripts/_env.py --setup` builds a private
environment in a per-user folder (shared with Make_Clean_PPTX, whose `doctor.py --fix` does the
same); this module sends the script there when needed.
Nothing happens when the packages are importable, or when no private environment exists.
"""
from __future__ import annotations

import importlib.util
import os
import sys
from pathlib import Path

VENV_ENV = "MAKE_CLEAN_PPTX_VENV"


def venv_dir() -> Path:
    """Folder of the private environment (override with MAKE_CLEAN_PPTX_VENV)."""
    if os.environ.get(VENV_ENV):
        return Path(os.environ[VENV_ENV]).expanduser()
    if os.name == "nt":
        base = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
    else:
        base = Path(os.environ.get("XDG_CACHE_HOME") or Path.home() / ".cache")
    return base / "make-clean-pptx" / "venv"


def venv_python() -> Path:
    d = venv_dir()
    return d / ("Scripts/python.exe" if os.name == "nt" else "bin/python")


def _maybe_reexec() -> None:
    if "--setup" in sys.argv or os.environ.get("MAKE_CLEAN_PPTX_REEXEC") or importlib.util.find_spec("pptx") is not None:
        return
    py = venv_python()
    if py.is_file() and Path(sys.prefix).resolve() != venv_dir().resolve():  # a venv python is a symlink: compare prefixes
        os.environ["MAKE_CLEAN_PPTX_REEXEC"] = "1"  # never loop
        os.execv(str(py), [str(py), *sys.argv])


_maybe_reexec()


def setup() -> int:
    """Build the private environment from ../requirements.txt (what --setup does)."""
    import subprocess
    req = Path(__file__).resolve().parent.parent / "requirements.txt"
    py = venv_python()
    try:
        if not py.is_file():
            subprocess.run([sys.executable, "-m", "venv", str(venv_dir())], check=True)
        subprocess.run([str(py), "-m", "pip", "install", "--disable-pip-version-check", "-r", str(req)], check=True)
    except (OSError, subprocess.CalledProcessError) as exc:
        print(f"Could not set up the private environment: {exc}. On Debian/Ubuntu try: sudo apt install python3-venv", file=sys.stderr)
        return 1
    print(f"Ready: {venv_dir()}")
    return 0


if __name__ == "__main__":
    sys.exit(setup() if "--setup" in sys.argv else 0)

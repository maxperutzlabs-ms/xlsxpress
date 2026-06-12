"""Build the xlsxpress pyapp executable.

Steps: build wheel -> download/unpack pyapp source -> set env vars ->
cargo build -> copy renamed exe to dist/.

Requirements: uv, cargo (Rust), internet access (first run).
Usage: python scripts/build_app.py [--pyapp-version 0.28.0]
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tarfile
import urllib.request
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
BUILD_DIR = PROJECT_ROOT / "build" / "pyapp"
DIST_DIR = PROJECT_ROOT / "dist"
PYTHON_VERSION = "3.14"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--pyapp-version", default="latest")
    args = parser.parse_args()

    wheel = build_wheel()
    version = wheel.name.split("-")[1]
    pyapp_src = fetch_pyapp_source(args.pyapp_version)

    # Embedded files must live inside the pyapp source tree (relative paths).
    embedded_wheel = pyapp_src / wheel.name
    shutil.copy2(wheel, embedded_wheel)

    env = os.environ | {
        "PYAPP_PROJECT_PATH": wheel.name,  # embed wheel (offline install)
        "PYAPP_EXEC_SPEC": "xlsxpress.app.main:main",
        "PYAPP_PYTHON_VERSION": PYTHON_VERSION,
        "PYAPP_DISTRIBUTION_EMBED": "1",  # no download on first run
        "PYAPP_FULL_ISOLATION": "1",  # own distribution per install
        "PYAPP_PASS_LOCATION": "1",  # PYAPP env var = exe path
        "PYAPP_IS_GUI": "1",  # no console window on Windows
        "PYAPP_PIP_EXTRA_ARGS": "--only-binary :all:",  # also used by self update
    }
    subprocess.run(["cargo", "build", "--release"], cwd=pyapp_src, env=env, check=True)

    built = pyapp_src / "target" / "release" / exe_name("pyapp")
    DIST_DIR.mkdir(exist_ok=True)
    target = DIST_DIR / exe_name(f"xlsxpress-{version}")
    shutil.copy2(built, target)
    print(f"Built: {target}")
    return 0


def build_wheel() -> Path:
    subprocess.run(["uv", "build", "--wheel"], cwd=PROJECT_ROOT, check=True)
    wheels = sorted(DIST_DIR.glob("xlsxpress-*.whl"), key=lambda p: p.stat().st_mtime)
    return wheels[-1]


def fetch_pyapp_source(version: str) -> Path:
    """Download and unpack the pyapp source release (cached in build/)."""
    if version == "latest":
        url = "https://github.com/ofek/pyapp/releases/latest/download/source.tar.gz"
    else:
        url = (
            f"https://github.com/ofek/pyapp/releases/download/v{version}/source.tar.gz"
        )
    BUILD_DIR.mkdir(parents=True, exist_ok=True)
    archive = BUILD_DIR / f"pyapp-{version}.tar.gz"
    if not archive.exists():
        print(f"Downloading {url}")
        urllib.request.urlretrieve(url, archive)
        with tarfile.open(archive) as tar:
            tar.extractall(BUILD_DIR, filter="data")
    source_dirs = [
        d for d in BUILD_DIR.iterdir() if d.is_dir() and d.name.startswith("pyapp-")
    ]
    return sorted(source_dirs)[-1]


def exe_name(stem: str) -> str:
    return f"{stem}.exe" if sys.platform == "win32" else stem


if __name__ == "__main__":
    sys.exit(main())

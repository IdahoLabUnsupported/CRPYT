# Copyright 2026, Battelle Energy Alliance, LLC. All Rights Reserved.
from __future__ import annotations

import argparse
import datetime as dt
import re
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
CRPYT_PATH = REPO_ROOT / "tools" / "CRPYT.py"
README_PATH = REPO_ROOT / "README.md"
VERSION_RE = re.compile(r'^(BUILD_VERSION = ")([0-9]+\.[0-9]+\.[0-9]+)(")$', re.MULTILINE)
DATE_RE = re.compile(r'^(BUILD_DATE = ")([0-9]{4}-[0-9]{2}-[0-9]{2})(")$', re.MULTILINE)
README_VERSION_RE = re.compile(r'^(- CRPYT build version: )(.+)$', re.MULTILINE)
README_DATE_RE = re.compile(r'^(- CRPYT build date: )(.+)$', re.MULTILINE)


def _read_text_preserve_newlines(path: Path) -> str:
    with path.open("r", encoding="utf-8", newline="") as handle:
        return handle.read()


def _write_text(path: Path, text: str) -> None:
    with path.open("w", encoding="utf-8", newline="") as handle:
        handle.write(text)


def _read_current_version() -> str:
    text = _read_text_preserve_newlines(CRPYT_PATH)
    match = VERSION_RE.search(text)
    if not match:
        raise SystemExit(f"Could not find BUILD_VERSION in {CRPYT_PATH.relative_to(REPO_ROOT)}")
    return match.group(2)


def _bump_patch(version: str) -> str:
    major, minor, patch = (int(part) for part in version.split("."))
    return f"{major}.{minor}.{patch + 1}"


def _replace_once(pattern: re.Pattern[str], repl: str, text: str, label: str) -> str:
    updated, count = pattern.subn(repl, text, count=1)
    if count != 1:
        raise SystemExit(f"Could not update {label}")
    return updated


def update_build_metadata(version: str, build_date: str) -> None:
    crpyt_text = _read_text_preserve_newlines(CRPYT_PATH)
    crpyt_text = _replace_once(VERSION_RE, rf"\g<1>{version}\g<3>", crpyt_text, "BUILD_VERSION")
    crpyt_text = _replace_once(DATE_RE, rf"\g<1>{build_date}\g<3>", crpyt_text, "BUILD_DATE")
    _write_text(CRPYT_PATH, crpyt_text)

    readme_text = _read_text_preserve_newlines(README_PATH)
    readme_text = _replace_once(
        README_VERSION_RE,
        rf"\g<1>{version}",
        readme_text,
        "README build version",
    )
    readme_text = _replace_once(
        README_DATE_RE,
        rf"\g<1>{build_date}",
        readme_text,
        "README build date",
    )
    _write_text(README_PATH, readme_text)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Update CRPYT build version/date metadata.")
    version_group = parser.add_mutually_exclusive_group(required=True)
    version_group.add_argument("--version", help="Set an explicit semantic version, such as 1.0.1.")
    version_group.add_argument("--patch", action="store_true", help="Increment the current patch version.")
    parser.add_argument(
        "--date",
        default=dt.date.today().isoformat(),
        help="Build date in YYYY-MM-DD format. Defaults to today.",
    )
    return parser


def main() -> None:
    args = build_parser().parse_args()
    try:
        parsed_date = dt.date.fromisoformat(args.date)
    except ValueError as exc:
        raise SystemExit("Build date must use YYYY-MM-DD format.") from exc

    version = args.version or _bump_patch(_read_current_version())
    if not re.fullmatch(r"[0-9]+\.[0-9]+\.[0-9]+", version):
        raise SystemExit("Build version must use major.minor.patch format.")

    update_build_metadata(version, parsed_date.isoformat())
    print(f"CRPYT build metadata updated to {version} ({parsed_date.isoformat()})")


if __name__ == "__main__":
    main()
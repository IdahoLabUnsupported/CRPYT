# Copyright 2026, Battelle Energy Alliance, LLC. All Rights Reserved.
from __future__ import annotations
"""Helpers for loading the cached legacy CRPYT engine.

The JSON-backed runtime still reuses the original formula interpreter class.
That interpreter only exists as a compiled cache artifact in this workspace,
so this module centralizes the logic for finding and loading that `.pyc`
whether we are running from source or from a frozen PyInstaller executable.
"""

import importlib.machinery
import importlib.util
import sys
from functools import lru_cache
from pathlib import Path
from types import ModuleType


def _load_pyc(module_name: str, pyc_path: Path) -> ModuleType:
    """Import a compiled Python module directly from a specific `.pyc` path."""
    loader = importlib.machinery.SourcelessFileLoader(module_name, str(pyc_path))
    spec = importlib.util.spec_from_loader(loader.name, loader)
    if spec is None:
        raise ImportError(f"Unable to load module spec from {pyc_path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    loader.exec_module(module)
    return module


def _candidate_cache_dirs() -> list[Path]:
    """Return the cache locations to search in source and frozen layouts."""
    dirs = [Path(__file__).resolve().parent / "__pycache__"]
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        dirs.append(Path(meipass) / "__pycache__")
    return dirs


@lru_cache(maxsize=1)
def load_legacy_engine_module() -> ModuleType:
    """Load the cached legacy engine used by the JSON workbook emulator."""
    searched: list[Path] = []
    for cache_dir in _candidate_cache_dirs():
        searched.append(cache_dir)
        candidates = sorted(cache_dir.glob("CRPYT.cpython-312-*.pyc"))
        if candidates:
            return _load_pyc("_crpyt_legacy_engine_json", candidates[-1])
    raise FileNotFoundError(
        "Legacy CRPYT engine cache not found in "
        + ", ".join(str(path) for path in searched)
    )


@lru_cache(maxsize=1)
def load_cached_gui_wrapper_module() -> ModuleType:
    """Load the cached GUI wrapper used by older compatibility paths."""
    searched: list[Path] = []
    for cache_dir in _candidate_cache_dirs():
        cache_path = cache_dir / "CRPYT.cpython-312.pyc"
        searched.append(cache_path)
        if cache_path.exists():
            return _load_pyc("_crpyt_cached_wrapper", cache_path)
    raise FileNotFoundError(
        "Cached GUI wrapper not found at "
        + ", ".join(str(path) for path in searched)
    )

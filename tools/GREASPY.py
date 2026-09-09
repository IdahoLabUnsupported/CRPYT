#!/usr/bin/env python3
# Copyright 2026, Battelle Energy Alliance, LLC. All Rights Reserved.
"""
GREASPY: standalone GREASE/GREASY workbook replacement.

This script reproduces the primary workflow from:
- GREASY_Rev2.xlsx
- GREASE_Isotope_USER_common_rev3.xlsx
- GREASE_Isotope_USER.xlsx

Features:
- Select up to 10 isotopes and activities
- Set counting time, distance, background inclusion
- Optional shielding (for workbooks that include shielding coefficients)
- Compute total spectrum and key summary metrics
- Plot spectrum and isotope contributions
- Export PeakEasy-compatible .spe files
"""

from __future__ import annotations

import datetime as _dt
import json
import math
import os
import sys
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Optional, Tuple

try:
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
except Exception:
    plt = None
    FigureCanvasTkAgg = None
    NavigationToolbar2Tk = None

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string


DEAD_TIME_FACTOR = 0.0000580138947
NUM_BINS = 16384

SHIELDING_MATERIALS = [
    "Polyethylene",
    "Aluminum",
    "Stainless Steel 316",
    "Tungsten",
    "Pb",
    "Depleted Uranium (DU)",
]


def default_crpyt_workbook_candidates() -> List[Path]:
    """Return likely GREASY workbook locations for source and frozen runs."""
    names = ["GREASPY_CRPYT_Rev1.xlsx"]
    roots: List[Path] = []

    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        roots.append(Path(meipass))

    script_path = Path(__file__).resolve()
    roots.append(script_path.parent)
    roots.append(script_path.parent.parent)
    roots.append(Path.cwd())
    roots.append(Path.cwd().parent)

    out: List[Path] = []
    seen = set()
    for root in roots:
        for name in names:
            candidate = root / name
            key = str(candidate).lower()
            if key not in seen:
                seen.add(key)
                out.append(candidate)
    return out


@dataclass
class IsotopeMeta:
    name: str
    sheet: str
    col_letter: str
    comment: str = ""


@dataclass
class CalcInputs:
    isotopes: List[Tuple[str, float]]
    count_time_sec: float
    distance_cm: float
    include_background: bool
    shielding_cm: Dict[str, float]


@dataclass
class CrpytSpectrumInputs:
    activity_by_zaid_ci: Dict[int, float]
    count_time_sec: float
    distance_cm: float
    include_background: bool
    shielding_cm: Dict[str, float]


@dataclass
class CalcResults:
    energy_kev: List[float]
    total_counts: List[float]
    per_isotope_counts: Dict[str, List[float]]
    background_counts: List[float]
    attenuation: List[float]
    unshielded_dose_30cm: float
    unshielded_gammas_emitted_s: float
    unshielded_gammas_incident_s: float
    unshielded_dead_time_pct: float
    unshielded_real_time_sec: float
    unshielded_distance_20pct_cm: float
    shielded_dose_30cm: Optional[float]
    shielded_gammas_emitted_s: Optional[float]
    shielded_gammas_incident_s: Optional[float]
    shielded_dead_time_pct: Optional[float]
    shielded_real_time_sec: Optional[float]
    shielded_distance_20pct_cm: Optional[float]
    eff_a: float
    eff_b: float
    matched_isotope_count: int = 0
    missing_zaids: Tuple[int, ...] = ()


class GreaseDataModel:
    """Loads workbook data and computes spectra equivalent to GREASY/GREASE sheets."""

    def __init__(self, workbook_path: Path):
        # This is the original workbook-backed implementation used by the
        # standalone GREASPY utility. CRPYT prefers the JSON-backed model
        # below, but this path remains useful for workbook-side validation.
        self.workbook_path = workbook_path
        self.wb = load_workbook(workbook_path, data_only=True, read_only=True)

        self.user_ws = self.wb["USER"]
        self.fp_ws = self.wb["FP"]
        self.other_ws = self.wb["Other"]
        self.bg_ws = self.wb["BG"]

        self.has_shielding = "Shielding" in self.wb.sheetnames
        self.shielding_ws = self.wb["Shielding"] if self.has_shielding else None

        self.profile = self._detect_profile()
        self.isotope_map = self._build_isotope_map()

        self.energy_kev = self._read_energy_axis()
        self.bg_cps = self._read_background_cps()

        self.shield_mu = self._read_shielding_coeffs() if self.has_shielding else {}
        self.shield_density = self._read_shielding_density() if self.has_shielding else []

    def _detect_profile(self) -> str:
        if "INSTRUCTIONS" in self.wb.sheetnames:
            return "common"
        return "full"

    def _build_isotope_map(self) -> Dict[str, IsotopeMeta]:
        """Build the user-visible isotope lookup from the workbook's USER sheet."""
        iso_map: Dict[str, IsotopeMeta] = {}

        if self.profile == "common":
            start_row = 11
            iso_col, sheet_col, col_col, comment_col = 25, 26, 27, 28
        else:
            start_row = 10
            iso_col, sheet_col, col_col, comment_col = 24, 25, 26, 27

        for r in range(start_row, 5000):
            name = self.user_ws.cell(r, iso_col).value
            sheet = self.user_ws.cell(r, sheet_col).value
            col = self.user_ws.cell(r, col_col).value
            comment = self.user_ws.cell(r, comment_col).value if comment_col else ""

            if name is None and sheet is None and col is None:
                if r > start_row + 50:
                    break
                continue

            if not name or not sheet or not col:
                continue

            name = str(name).strip()
            sheet = str(sheet).strip()
            col = str(col).strip()
            comment = "" if comment is None else str(comment).strip()

            if sheet not in ("FP", "Other"):
                continue

            iso_map[name] = IsotopeMeta(name=name, sheet=sheet, col_letter=col, comment=comment)

        if not iso_map:
            raise RuntimeError("No isotope mapping data found in USER sheet.")

        return dict(sorted(iso_map.items(), key=lambda kv: kv[0].upper()))

    def _read_energy_axis(self) -> List[float]:
        axis = []
        for r in range(5, 5 + NUM_BINS):
            ev = self.fp_ws.cell(r, 1).value
            axis.append((float(ev) * 1000.0) if ev is not None else 0.0)
        return axis

    def _read_background_cps(self) -> List[float]:
        col_idx = 1 if self.bg_ws.cell(1, 1).value is not None else 2
        out = []
        for r in range(2, 2 + NUM_BINS):
            v = self.bg_ws.cell(r, col_idx).value
            out.append(float(v) if v is not None else 0.0)
        return out

    def _read_shielding_density(self) -> List[float]:
        if not self.shielding_ws:
            return []
        # Row 2, cols B:G
        vals = []
        for c in range(2, 8):
            v = self.shielding_ws.cell(2, c).value
            vals.append(float(v) if v is not None else 0.0)
        return vals

    def _read_shielding_coeffs(self) -> Dict[str, List[float]]:
        if not self.shielding_ws:
            return {}

        coeffs: Dict[str, List[float]] = {}
        for idx, material in enumerate(SHIELDING_MATERIALS):
            col = 2 + idx
            arr = []
            for r in range(4, 4 + NUM_BINS):
                v = self.shielding_ws.cell(r, col).value
                arr.append(float(v) if v is not None else 0.0)
            coeffs[material] = arr
        return coeffs

    def isotope_names(self) -> List[str]:
        return list(self.isotope_map.keys())

    def isotope_suggestion(self, name: str) -> str:
        meta = self.isotope_map.get(name)
        return meta.comment if meta else ""

    def _ws_for_meta(self, meta: IsotopeMeta):
        return self.fp_ws if meta.sheet == "FP" else self.other_ws

    def _read_isotope_constants(self, meta: IsotopeMeta) -> Tuple[float, float]:
        ws = self._ws_for_meta(meta)
        c = column_index_from_string(meta.col_letter)

        dose_per_ci = ws.cell(3, c).value
        gammas_per_ci = ws.cell(4, c).value

        return float(dose_per_ci or 0.0), float(gammas_per_ci or 0.0)

    def _read_isotope_spectrum_cps(self, meta: IsotopeMeta) -> List[float]:
        ws = self._ws_for_meta(meta)
        c = column_index_from_string(meta.col_letter)

        out = []
        for r in range(5, 5 + NUM_BINS):
            v = ws.cell(r, c).value
            out.append(float(v) if v is not None else 0.0)
        return out

    def compute(self, inp: CalcInputs) -> CalcResults:
        """Reproduce the GREASY workbook calculation path from direct isotope inputs."""
        if inp.count_time_sec <= 0:
            raise ValueError("Counting time must be > 0.")
        if inp.distance_cm <= 0:
            raise ValueError("Distance must be > 0.")

        active = [(name, act) for name, act in inp.isotopes if name and act > 0]
        if not active:
            raise ValueError("Select at least one isotope with activity > 0.")

        for name, _ in active:
            if name not in self.isotope_map:
                raise ValueError(f"Isotope '{name}' not found in loaded workbook mapping.")

        # Build attenuation curve (1.0 for unshielded or if workbook has no shielding sheet)
        attenuation = [1.0] * NUM_BINS
        total_shield_thickness = sum(max(0.0, inp.shielding_cm.get(m, 0.0)) for m in SHIELDING_MATERIALS)

        if self.has_shielding and total_shield_thickness > 0:
            total_mu_rho_x = [0.0] * NUM_BINS
            for i, material in enumerate(SHIELDING_MATERIALS):
                t_cm = max(0.0, inp.shielding_cm.get(material, 0.0))
                if t_cm <= 0:
                    continue
                density = self.shield_density[i]
                mu = self.shield_mu[material]
                for b in range(NUM_BINS):
                    total_mu_rho_x[b] += mu[b] * density * t_cm
            attenuation = [math.exp(-x) for x in total_mu_rho_x]

        background_counts = [
            (self.bg_cps[i] * inp.count_time_sec if inp.include_background else 0.0)
            for i in range(NUM_BINS)
        ]

        per_isotope_counts: Dict[str, List[float]] = {}
        signal_unshielded = [0.0] * NUM_BINS
        signal_shielded = [0.0] * NUM_BINS

        unshielded_dose_30cm = 0.0
        unshielded_gammas_emitted_s = 0.0

        for name, activity_ci in active:
            meta = self.isotope_map[name]
            dose_per_ci, gammas_per_ci = self._read_isotope_constants(meta)
            cps_vec = self._read_isotope_spectrum_cps(meta)

            # Excel scale: spectrum * activity * live_time * 100 / distance^2
            scaled_unshielded = [
                cps_vec[i] * activity_ci * inp.count_time_sec * 100.0 / (inp.distance_cm * inp.distance_cm)
                for i in range(NUM_BINS)
            ]
            scaled_shielded = [scaled_unshielded[i] * attenuation[i] for i in range(NUM_BINS)]

            per_isotope_counts[name] = scaled_shielded

            for i in range(NUM_BINS):
                signal_unshielded[i] += scaled_unshielded[i]
                signal_shielded[i] += scaled_shielded[i]

            unshielded_dose_30cm += activity_ci * dose_per_ci
            unshielded_gammas_emitted_s += activity_ci * gammas_per_ci

        # J5 behavior in GREASY: sum(background cps)
        bg_total_cps = sum(self.bg_cps) if inp.include_background else 0.0
        unshielded_gammas_emitted_s += bg_total_cps

        total_counts = [signal_shielded[i] + background_counts[i] for i in range(NUM_BINS)]

        unshielded_gammas_incident_s = unshielded_gammas_emitted_s / (inp.distance_cm * inp.distance_cm)
        unshielded_dead_time_pct = min(100.0, unshielded_gammas_incident_s * DEAD_TIME_FACTOR * 100.0)
        unshielded_real_time_sec = inp.count_time_sec / max(1e-12, (1.0 - unshielded_dead_time_pct / 100.0))
        unshielded_distance_20pct_cm = math.sqrt(DEAD_TIME_FACTOR * unshielded_gammas_emitted_s / 0.2)

        # Efficiency factors as in USER!G11/G12
        eff_a = self.energy_kev[0]
        eff_b = (self.energy_kev[-1] - self.energy_kev[0]) / NUM_BINS

        shielded_dose_30cm = None
        shielded_gammas_emitted_s = None
        shielded_gammas_incident_s = None
        shielded_dead_time_pct = None
        shielded_real_time_sec = None
        shielded_distance_20pct_cm = None

        if total_shield_thickness > 0:
            j6 = sum(signal_shielded)
            j7 = sum(signal_unshielded) + sum(background_counts)
            ratio = (j6 / j7) if j7 > 0 else 0.0

            shielded_dose_30cm = unshielded_dose_30cm * ratio
            shielded_gammas_emitted_s = unshielded_gammas_emitted_s
            shielded_gammas_incident_s = unshielded_gammas_incident_s * ratio
            shielded_dead_time_pct = min(100.0, shielded_gammas_incident_s * DEAD_TIME_FACTOR * 100.0)
            shielded_real_time_sec = inp.count_time_sec / max(1e-12, (1.0 - shielded_dead_time_pct / 100.0))
            shielded_distance_20pct_cm = math.sqrt(
                DEAD_TIME_FACTOR * shielded_gammas_emitted_s * ratio / 0.2
            )

        return CalcResults(
            energy_kev=self.energy_kev,
            total_counts=total_counts,
            per_isotope_counts=per_isotope_counts,
            background_counts=background_counts,
            attenuation=attenuation,
            unshielded_dose_30cm=unshielded_dose_30cm,
            unshielded_gammas_emitted_s=unshielded_gammas_emitted_s,
            unshielded_gammas_incident_s=unshielded_gammas_incident_s,
            unshielded_dead_time_pct=unshielded_dead_time_pct,
            unshielded_real_time_sec=unshielded_real_time_sec,
            unshielded_distance_20pct_cm=unshielded_distance_20pct_cm,
            shielded_dose_30cm=shielded_dose_30cm,
            shielded_gammas_emitted_s=shielded_gammas_emitted_s,
            shielded_gammas_incident_s=shielded_gammas_incident_s,
            shielded_dead_time_pct=shielded_dead_time_pct,
            shielded_real_time_sec=shielded_real_time_sec,
            shielded_distance_20pct_cm=shielded_distance_20pct_cm,
            eff_a=eff_a,
            eff_b=eff_b,
        )


class CrpytGreaseSpectrumModel:
    """Computes GREASPY_CRPYT detector spectra directly from CRPYT isotope activities."""

    def __init__(self, data_source: Path):
        # The runtime can load either from the extracted GREASY JSON bundle
        # (normal standalone path) or directly from the workbook as a fallback.
        self.data_source = Path(data_source)
        self.bundle_dir: Optional[Path] = None
        self.workbook_path: Optional[Path] = None
        self.wb = None
        self.greasy_user_ws = None
        self.fp_ws = None
        self.other_ws = None
        self.bg_ws = None
        self.shielding_ws = None
        self.defaults: Dict[str, object] = {}
        self.has_shielding = False

        self.energy_kev: Optional[List[float]] = None
        self.bg_cps: Optional[List[float]] = None
        self.shield_mu: Optional[Dict[str, List[float]]] = None
        self.shield_density: Optional[List[float]] = None
        self.zaid_column_map: Dict[int, Tuple[str, int]] = {}
        self.zaid_file_map: Dict[int, Path] = {}
        self._loaded_bundle_files: set[str] = set()
        self._entry_cache: Dict[int, Tuple[float, float, List[float]]] = {}

        if self.data_source.is_dir():
            self._load_bundle(self.data_source)
        elif self.data_source.is_file():
            self._load_workbook(self.data_source)
        else:
            raise FileNotFoundError(f"GREASY data source was not found: {self.data_source}")

    def _normalize_vector(self, values: object, length: int = NUM_BINS) -> List[float]:
        """Normalize serialized vectors to a fixed bin count of floats."""
        out: List[float] = []
        if isinstance(values, list):
            source = values
        elif isinstance(values, tuple):
            source = list(values)
        elif values is None:
            source = []
        else:
            source = list(values)

        for value in source[:length]:
            out.append(float(value) if value is not None else 0.0)
        if len(out) < length:
            out.extend([0.0] * (length - len(out)))
        return out

    def _load_bundle(self, bundle_dir: Path) -> None:
        """Load the pre-extracted GREASY JSON bundle into runtime caches."""
        manifest_path = bundle_dir / "manifest.json"
        if not manifest_path.exists():
            raise FileNotFoundError(f"GREASY bundle manifest was not found: {manifest_path}")

        with manifest_path.open("r", encoding="utf-8") as handle:
            manifest = json.load(handle)

        metadata_rel = manifest.get("metadata_file", "metadata.json")
        metadata_path = bundle_dir / metadata_rel
        if not metadata_path.exists():
            raise FileNotFoundError(f"GREASY bundle metadata was not found: {metadata_path}")

        with metadata_path.open("r", encoding="utf-8") as handle:
            metadata = json.load(handle)

        self.bundle_dir = bundle_dir
        self.defaults = metadata.get("defaults", {})
        self.energy_kev = self._normalize_vector(metadata.get("energy_kev"))
        self.bg_cps = self._normalize_vector(metadata.get("background_cps"))

        shielding = metadata.get("shielding") or {}
        coefficients = shielding.get("coefficients") or {}
        if coefficients:
            self.has_shielding = True
            self.shield_density = [
                float(value) if value is not None else 0.0
                for value in (shielding.get("density_g_cc") or [])
            ][: len(SHIELDING_MATERIALS)]
            if len(self.shield_density) < len(SHIELDING_MATERIALS):
                self.shield_density.extend([0.0] * (len(SHIELDING_MATERIALS) - len(self.shield_density)))
            self.shield_mu = {
                material: self._normalize_vector(coefficients.get(material))
                for material in SHIELDING_MATERIALS
            }
        else:
            self.has_shielding = False
            self.shield_density = []
            self.shield_mu = {}

        zaid_files = metadata.get("zaid_files") or {}
        for key, rel_path in zaid_files.items():
            try:
                zaid = int(key)
            except (TypeError, ValueError):
                continue
            self.zaid_file_map[zaid] = Path(rel_path)

    def _load_workbook(self, workbook_path: Path) -> None:
        """Fallback loader used when a live GREASY workbook is supplied."""
        self.workbook_path = workbook_path
        self.wb = load_workbook(workbook_path, data_only=True, read_only=True)
        self.greasy_user_ws = self.wb["GREASY USER"]
        self.fp_ws = self.wb["FP"]
        self.other_ws = self.wb["Other"]
        self.bg_ws = self.wb["BG"]
        self.has_shielding = "Shielding" in self.wb.sheetnames
        self.shielding_ws = self.wb["Shielding"] if self.has_shielding else None
        self.zaid_column_map = self._build_zaid_column_map()

    def _build_zaid_column_map(self) -> Dict[int, Tuple[str, int]]:
        mapping: Dict[int, Tuple[str, int]] = {}
        for sheet_name, ws in (("FP", self.fp_ws), ("Other", self.other_ws)):
            header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
            for col, value in enumerate(header_row[1:], start=2):
                try:
                    zaid = int(float(value))
                except Exception:
                    continue
                mapping[zaid] = (sheet_name, col)
        return mapping

    def _read_energy_axis(self) -> List[float]:
        axis = []
        for (energy_mev,) in self.fp_ws.iter_rows(
            min_row=6,
            max_row=6 + NUM_BINS - 1,
            min_col=1,
            max_col=1,
            values_only=True,
        ):
            axis.append((float(energy_mev) * 1000.0) if energy_mev is not None else 0.0)
        return axis

    def _read_background_cps(self) -> List[float]:
        out = []
        for (value,) in self.bg_ws.iter_rows(
            min_row=2,
            max_row=2 + NUM_BINS - 1,
            min_col=1,
            max_col=1,
            values_only=True,
        ):
            out.append(float(value) if value is not None else 0.0)
        return out

    def _read_shielding_density(self) -> List[float]:
        if not self.shielding_ws:
            return []
        row = next(
            self.shielding_ws.iter_rows(
                min_row=2,
                max_row=2,
                min_col=2,
                max_col=7,
                values_only=True,
            )
        )
        return [float(value) if value is not None else 0.0 for value in row]

    def _read_shielding_coeffs(self) -> Dict[str, List[float]]:
        if not self.shielding_ws:
            return {}
        coeffs: Dict[str, List[float]] = {}
        for index, material in enumerate(SHIELDING_MATERIALS):
            col = 2 + index
            arr = []
            for (value,) in self.shielding_ws.iter_rows(
                min_row=4,
                max_row=4 + NUM_BINS - 1,
                min_col=col,
                max_col=col,
                values_only=True,
            ):
                arr.append(float(value) if value is not None else 0.0)
            coeffs[material] = arr
        return coeffs

    def _ensure_energy_and_background_loaded(self) -> None:
        """Lazy-load the energy axis and background for workbook-backed runs."""
        if self.energy_kev is None and self.fp_ws is not None:
            self.energy_kev = self._read_energy_axis()
        if self.bg_cps is None and self.bg_ws is not None:
            self.bg_cps = self._read_background_cps()

    def _ensure_shielding_loaded(self) -> None:
        """Lazy-load shielding data only if the chosen source contains it."""
        if not self.has_shielding:
            self.shield_mu = {}
            self.shield_density = []
            return
        if self.shield_mu is None and self.shielding_ws is not None:
            self.shield_mu = self._read_shielding_coeffs()
        if self.shield_density is None and self.shielding_ws is not None:
            self.shield_density = self._read_shielding_density()

    def _load_isotope_entries(self, zaids: List[int]) -> None:
        """Populate the per-ZAID spectrum cache from JSON chunks or workbook columns."""
        if self.bundle_dir is not None:
            rel_paths: Dict[str, Path] = {}
            for zaid in zaids:
                if zaid in self._entry_cache:
                    continue
                rel_path = self.zaid_file_map.get(zaid)
                if rel_path is None:
                    continue
                rel_paths[str(rel_path)] = rel_path

            # Bundle files are chunked by ZAID group, so one file load may
            # satisfy many requested isotopes for the current run.
            for rel_path_key, rel_path in rel_paths.items():
                if rel_path_key in self._loaded_bundle_files:
                    continue
                entry_path = self.bundle_dir / rel_path
                if not entry_path.exists():
                    continue
                with entry_path.open("r", encoding="utf-8") as handle:
                    payload = json.load(handle)
                entries = payload.get("entries") or {}
                for zaid_key, entry_payload in entries.items():
                    try:
                        zaid = int(zaid_key)
                    except (TypeError, ValueError):
                        continue
                    if not isinstance(entry_payload, dict):
                        continue
                    self._entry_cache[zaid] = (
                        float(entry_payload.get("dose_per_ci") or 0.0),
                        float(entry_payload.get("gammas_per_ci") or 0.0),
                        self._normalize_vector(entry_payload.get("spectrum_cps")),
                    )
                self._loaded_bundle_files.add(rel_path_key)
            return

        pending_by_sheet: Dict[str, Dict[int, int]] = {"FP": {}, "Other": {}}
        for zaid in zaids:
            if zaid in self._entry_cache:
                continue
            source = self.zaid_column_map.get(zaid)
            if source is None:
                continue
            sheet_name, col = source
            pending_by_sheet[sheet_name][zaid] = col

        for sheet_name, columns_by_zaid in pending_by_sheet.items():
            if not columns_by_zaid:
                continue

            ws = self.fp_ws if sheet_name == "FP" else self.other_ws
            min_col = min(columns_by_zaid.values())
            max_col = max(columns_by_zaid.values())
            header_rows = list(
                ws.iter_rows(
                    min_row=1,
                    max_row=4,
                    min_col=min_col,
                    max_col=max_col,
                    values_only=True,
                )
            )
            dose_row = header_rows[2]
            gamma_row = header_rows[3]
            spectra_by_zaid: Dict[int, List[float]] = {
                zaid: [] for zaid in columns_by_zaid
            }
            indexed_columns = list(columns_by_zaid.items())

            for row_values in ws.iter_rows(
                min_row=6,
                max_row=6 + NUM_BINS - 1,
                min_col=min_col,
                max_col=max_col,
                values_only=True,
            ):
                for zaid, col in indexed_columns:
                    offset = col - min_col
                    value = row_values[offset] if offset < len(row_values) else None
                    spectra_by_zaid[zaid].append(float(value) if value is not None else 0.0)

            for zaid, col in indexed_columns:
                offset = col - min_col
                dose_per_ci = float(dose_row[offset] or 0.0)
                gammas_per_ci = float(gamma_row[offset] or 0.0)
                self._entry_cache[zaid] = (
                    dose_per_ci,
                    gammas_per_ci,
                    spectra_by_zaid[zaid],
                )

    def compute_from_activities(
        self,
        inp: CrpytSpectrumInputs,
        isotope_names_by_zaid: Optional[Dict[int, str]] = None,
    ) -> CalcResults:
        """Compute the detector spectrum from CRPYT isotope activities by ZAID."""
        if inp.count_time_sec <= 0:
            raise ValueError("Counting time must be > 0.")
        if inp.distance_cm <= 0:
            raise ValueError("Counting distance must be > 0.")

        # Collapse the current isotope inventory down to the active ZAIDs the
        # GREASY library can score.
        active = [
            (int(zaid), float(activity_ci))
            for zaid, activity_ci in inp.activity_by_zaid_ci.items()
            if float(activity_ci) > 0.0
        ]
        if not active:
            raise ValueError("No isotopes with activity > 0 are available for spectrum generation.")

        self._ensure_energy_and_background_loaded()
        assert self.energy_kev is not None
        assert self.bg_cps is not None
        self._load_isotope_entries([zaid for zaid, _activity_ci in active])

        # Attenuation is computed as an energy-dependent transmission curve
        # from the user-entered shielding stack.
        attenuation = [1.0] * NUM_BINS
        total_shield_thickness = sum(max(0.0, inp.shielding_cm.get(material, 0.0)) for material in SHIELDING_MATERIALS)
        if self.has_shielding and total_shield_thickness > 0.0:
            self._ensure_shielding_loaded()
            assert self.shield_mu is not None
            assert self.shield_density is not None
            total_mu_rho_x = [0.0] * NUM_BINS
            for index, material in enumerate(SHIELDING_MATERIALS):
                thickness_cm = max(0.0, inp.shielding_cm.get(material, 0.0))
                if thickness_cm <= 0.0:
                    continue
                density = self.shield_density[index]
                mu = self.shield_mu[material]
                for bin_index in range(NUM_BINS):
                    total_mu_rho_x[bin_index] += mu[bin_index] * density * thickness_cm
            attenuation = [math.exp(-value) for value in total_mu_rho_x]

        background_counts = [
            (self.bg_cps[index] * inp.count_time_sec if inp.include_background else 0.0)
            for index in range(NUM_BINS)
        ]

        per_isotope_counts: Dict[str, List[float]] = {}
        signal_unshielded = [0.0] * NUM_BINS
        signal_shielded = [0.0] * NUM_BINS
        unshielded_dose_30cm = 0.0
        unshielded_gammas_emitted_s = 0.0
        matched_isotope_count = 0
        missing_zaids: List[int] = []

        # Each library spectrum is scaled to the requested activity, count time,
        # and detector distance using the same factors as the workbook.
        for zaid, activity_ci in active:
            entry = self._entry_cache.get(zaid)
            if entry is None:
                missing_zaids.append(zaid)
                continue

            matched_isotope_count += 1
            dose_per_ci, gammas_per_ci, cps_vec = entry
            scaled_unshielded = [
                cps_vec[index] * activity_ci * inp.count_time_sec * 100.0 / (inp.distance_cm * inp.distance_cm)
                for index in range(NUM_BINS)
            ]
            scaled_shielded = [scaled_unshielded[index] * attenuation[index] for index in range(NUM_BINS)]

            label = isotope_names_by_zaid.get(zaid, str(zaid)) if isotope_names_by_zaid else str(zaid)
            if label in per_isotope_counts:
                label = f"{label} ({zaid})"
            per_isotope_counts[label] = scaled_shielded

            for index in range(NUM_BINS):
                signal_unshielded[index] += scaled_unshielded[index]
                signal_shielded[index] += scaled_shielded[index]

            unshielded_dose_30cm += activity_ci * dose_per_ci
            unshielded_gammas_emitted_s += activity_ci * gammas_per_ci

        if matched_isotope_count == 0:
            raise ValueError("No current isotopes are available in the GREASPY_CRPYT spectrum library.")

        bg_total_cps = sum(self.bg_cps) if inp.include_background else 0.0
        unshielded_gammas_emitted_s += bg_total_cps

        total_counts = [signal_shielded[index] + background_counts[index] for index in range(NUM_BINS)]
        unshielded_gammas_incident_s = unshielded_gammas_emitted_s / (inp.distance_cm * inp.distance_cm)
        unshielded_dead_time_pct = min(100.0, unshielded_gammas_incident_s * DEAD_TIME_FACTOR * 100.0)
        unshielded_real_time_sec = inp.count_time_sec / max(1e-12, (1.0 - unshielded_dead_time_pct / 100.0))
        unshielded_distance_20pct_cm = math.sqrt(DEAD_TIME_FACTOR * unshielded_gammas_emitted_s / 0.2)

        eff_a = self.energy_kev[0]
        eff_b = (self.energy_kev[-1] - self.energy_kev[0]) / NUM_BINS

        shielded_dose_30cm = None
        shielded_gammas_emitted_s = None
        shielded_gammas_incident_s = None
        shielded_dead_time_pct = None
        shielded_real_time_sec = None
        shielded_distance_20pct_cm = None

        if total_shield_thickness > 0.0:
            shielded_signal_total = sum(signal_shielded)
            reference_total = sum(signal_unshielded) + sum(background_counts)
            ratio = (shielded_signal_total / reference_total) if reference_total > 0.0 else 0.0

            shielded_dose_30cm = unshielded_dose_30cm * ratio
            shielded_gammas_emitted_s = unshielded_gammas_emitted_s
            shielded_gammas_incident_s = unshielded_gammas_incident_s * ratio
            shielded_dead_time_pct = min(100.0, shielded_gammas_incident_s * DEAD_TIME_FACTOR * 100.0)
            shielded_real_time_sec = inp.count_time_sec / max(1e-12, (1.0 - shielded_dead_time_pct / 100.0))
            shielded_distance_20pct_cm = math.sqrt(
                DEAD_TIME_FACTOR * shielded_gammas_emitted_s * ratio / 0.2
            )

        return CalcResults(
            energy_kev=self.energy_kev,
            total_counts=total_counts,
            per_isotope_counts=per_isotope_counts,
            background_counts=background_counts,
            attenuation=attenuation,
            unshielded_dose_30cm=unshielded_dose_30cm,
            unshielded_gammas_emitted_s=unshielded_gammas_emitted_s,
            unshielded_gammas_incident_s=unshielded_gammas_incident_s,
            unshielded_dead_time_pct=unshielded_dead_time_pct,
            unshielded_real_time_sec=unshielded_real_time_sec,
            unshielded_distance_20pct_cm=unshielded_distance_20pct_cm,
            shielded_dose_30cm=shielded_dose_30cm,
            shielded_gammas_emitted_s=shielded_gammas_emitted_s,
            shielded_gammas_incident_s=shielded_gammas_incident_s,
            shielded_dead_time_pct=shielded_dead_time_pct,
            shielded_real_time_sec=shielded_real_time_sec,
            shielded_distance_20pct_cm=shielded_distance_20pct_cm,
            eff_a=eff_a,
            eff_b=eff_b,
            matched_isotope_count=matched_isotope_count,
            missing_zaids=tuple(sorted(missing_zaids)),
        )


class GreaseApp(tk.Tk):
    def __init__(self):
        """Standalone GREASPY GUI preserved for direct workbook-side usage."""
        super().__init__()
        self.title("GREASPY - GREASE/GREASY Standalone")
        self.geometry("1600x900")
        self._closing = False
        self.fig = None
        self.ax = None
        self.canvas = None
        self.toolbar = None
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self.model: Optional[GreaseDataModel] = None
        self.results: Optional[CalcResults] = None

        self.workbook_var = tk.StringVar()
        self.count_time_var = tk.StringVar(value="300")
        self.distance_var = tk.StringVar(value="20")
        self.include_bg_var = tk.BooleanVar(value=True)
        self.log_scale_var = tk.BooleanVar(value=False)

        self.selected_isotope_vars = [tk.StringVar() for _ in range(10)]
        self.selected_activity_vars = [tk.StringVar(value="1e-6") for _ in range(10)]

        self.shield_vars = {m: tk.StringVar(value="0") for m in SHIELDING_MATERIALS}
        self.suggestion_vars = [tk.StringVar(value="") for _ in range(10)]

        self.summary_var = tk.StringVar(value="Load a workbook to begin.")

        self._build_ui()
        self._try_autoload_default()

    def _build_ui(self):
        """Construct the standalone GREASPY interface and plotting widgets."""
        main = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        main.pack(fill=tk.BOTH, expand=True)

        left = ttk.Frame(main, padding=8)
        right = ttk.Frame(main, padding=8)
        main.add(left, weight=3)
        main.add(right, weight=4)

        inputs = ttk.LabelFrame(left, text="Inputs", padding=8)
        inputs.pack(fill=tk.X)

        row = ttk.Frame(inputs)
        row.pack(fill=tk.X, pady=2)
        ttk.Label(row, text="Live Counting Time (sec)", width=28).pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=self.count_time_var, width=16).pack(side=tk.LEFT)

        row = ttk.Frame(inputs)
        row.pack(fill=tk.X, pady=2)
        ttk.Label(row, text="Counting Distance (cm)", width=28).pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=self.distance_var, width=16).pack(side=tk.LEFT)

        row = ttk.Frame(inputs)
        row.pack(fill=tk.X, pady=2)
        ttk.Checkbutton(row, text="Include Background", variable=self.include_bg_var).pack(side=tk.LEFT)

        iso_frame = ttk.LabelFrame(left, text="Isotopes (up to 10)", padding=8)
        iso_frame.pack(fill=tk.BOTH, expand=True, pady=8)

        hdr = ttk.Frame(iso_frame)
        hdr.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(hdr, text="#", width=3).pack(side=tk.LEFT)
        ttk.Label(hdr, text="Isotope", width=22).pack(side=tk.LEFT, padx=4)
        ttk.Label(hdr, text="Activity (Ci)", width=12).pack(side=tk.LEFT, padx=4)
        ttk.Label(hdr, text="Suggested Isotopes", width=34).pack(side=tk.LEFT, padx=4)
        ttk.Label(hdr, text="").pack(side=tk.LEFT, padx=4)

        for i in range(10):
            r = ttk.Frame(iso_frame)
            r.pack(fill=tk.X, pady=2)
            ttk.Label(r, text=f"{i+1}", width=3).pack(side=tk.LEFT)
            combo = ttk.Combobox(
                r,
                textvariable=self.selected_isotope_vars[i],
                values=[],
                width=22,
                state="disabled",
            )
            combo.pack(side=tk.LEFT, padx=4)
            combo.bind("<<ComboboxSelected>>", lambda _e, idx=i: self._update_suggestion(idx))
            ttk.Entry(r, textvariable=self.selected_activity_vars[i], width=12).pack(side=tk.LEFT, padx=4)
            ttk.Label(r, textvariable=self.suggestion_vars[i], width=34).pack(side=tk.LEFT, padx=4)
            ttk.Button(r, text="Clear", width=7, command=lambda idx=i: self._clear_isotope_slot(idx)).pack(
                side=tk.LEFT, padx=4
            )
            setattr(self, f"iso_combo_{i}", combo)

        iso_actions = ttk.Frame(iso_frame)
        iso_actions.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(iso_actions, text="Clear All Isotopes", command=self._clear_all_isotopes).pack(side=tk.LEFT)

        shield_frame = ttk.LabelFrame(left, text="Shielding Thickness (cm)", padding=8)
        shield_frame.pack(fill=tk.X)

        for material in SHIELDING_MATERIALS:
            r = ttk.Frame(shield_frame)
            r.pack(fill=tk.X, pady=1)
            ttk.Label(r, text=material, width=24).pack(side=tk.LEFT)
            ttk.Entry(r, textvariable=self.shield_vars[material], width=10).pack(side=tk.LEFT)

        actions = ttk.Frame(left)
        actions.pack(fill=tk.X, pady=8)
        ttk.Button(actions, text="Calculate", command=self._calculate).pack(side=tk.LEFT, padx=3)
        ttk.Button(actions, text="Export .SPE", command=self._export_spe).pack(side=tk.LEFT, padx=3)
        ttk.Button(actions, text="Export CSV", command=self._export_csv).pack(side=tk.LEFT, padx=3)
        ttk.Checkbutton(
            actions,
            text="Log Y-axis",
            variable=self.log_scale_var,
            command=self._refresh_plot,
        ).pack(side=tk.LEFT, padx=8)

        results_frame = ttk.LabelFrame(right, text="Summary", padding=8)
        results_frame.pack(fill=tk.X)
        ttk.Label(results_frame, textvariable=self.summary_var, justify=tk.LEFT).pack(anchor=tk.W)

        plot_wrap = ttk.LabelFrame(right, text="Spectrum", padding=8)
        plot_wrap.pack(fill=tk.BOTH, expand=True, pady=8)

        if plt is not None and FigureCanvasTkAgg is not None and NavigationToolbar2Tk is not None:
            self.fig, self.ax = plt.subplots(figsize=(9, 6), dpi=100)
            self.canvas = FigureCanvasTkAgg(self.fig, master=plot_wrap)
            self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            self.toolbar = NavigationToolbar2Tk(self.canvas, plot_wrap)
            self.toolbar.update()
        else:
            ttk.Label(plot_wrap, text="matplotlib not available - plotting disabled.").pack(anchor=tk.W)

    def _on_close(self):
        """Release Matplotlib/Tk resources so the process exits cleanly."""
        if self._closing:
            return
        self._closing = True

        if self.canvas is not None:
            try:
                self.canvas.get_tk_widget().destroy()
            except Exception:
                pass
            self.canvas = None

        if self.toolbar is not None:
            try:
                self.toolbar.destroy()
            except Exception:
                pass
            self.toolbar = None

        if plt is not None and self.fig is not None:
            try:
                plt.close(self.fig)
            except Exception:
                pass
        self.fig = None
        self.ax = None

        try:
            self.quit()
        finally:
            self.destroy()

    def _try_autoload_default(self):
        candidates = self._default_workbook_candidates()
        for p in candidates:
            if p.exists():
                self.workbook_var.set(str(p))
                try:
                    self._load_workbook()
                except Exception:
                    pass
                break

    def _default_workbook_candidates(self) -> List[Path]:
        names = [
            "GREASY_Rev2.xlsx",
            "GREASE_Isotope_USER_common_rev3.xlsx",
            "GREASE_Isotope_USER.xlsx",
        ]

        roots: List[Path] = []

        # PyInstaller one-file extraction directory
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            roots.append(Path(meipass))

        # Script directory and project root
        script_path = Path(__file__).resolve()
        roots.append(script_path.parent)
        if script_path.parent.name.lower() == "python greasy":
            roots.append(script_path.parent.parent)

        # Working directory as last fallback
        roots.append(Path.cwd())

        out: List[Path] = []
        seen = set()
        for root in roots:
            for n in names:
                p = root / n
                key = str(p).lower()
                if key not in seen:
                    seen.add(key)
                    out.append(p)
        return out

    def _browse_workbook(self):
        path = filedialog.askopenfilename(
            title="Select GREASE/GREASY workbook",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if path:
            self.workbook_var.set(path)

    def _set_shielding_enabled(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        for material in SHIELDING_MATERIALS:
            # We locate the entry widget by shared variable usage through children traversal.
            # Simple approach: leave values but gate at calculation.
            _ = state

    def _load_workbook(self):
        path = Path(self.workbook_var.get().strip())
        if not path.exists():
            raise FileNotFoundError(f"Workbook not found: {path}")

        self.summary_var.set("Loading workbook data...")
        self.update_idletasks()

        self.model = GreaseDataModel(path)

        names = self.model.isotope_names()
        combo_values = [""] + names
        for i in range(10):
            combo: ttk.Combobox = getattr(self, f"iso_combo_{i}")
            combo.configure(values=combo_values, state="readonly")
            combo.set("")
            self.suggestion_vars[i].set("")

        if not self.model.has_shielding:
            for m in SHIELDING_MATERIALS:
                self.shield_vars[m].set("0")

        self.summary_var.set(
            f"Loaded {path.name}\n"
            f"Profile: {self.model.profile} | Isotopes: {len(names)} | Shielding data: {self.model.has_shielding}"
        )

    def _update_suggestion(self, idx: int):
        if not self.model:
            return
        iso = self.selected_isotope_vars[idx].get().strip()
        tip = self.model.isotope_suggestion(iso)
        self.suggestion_vars[idx].set(tip)

    def _clear_isotope_slot(self, idx: int):
        self.selected_isotope_vars[idx].set("")
        self.suggestion_vars[idx].set("")

    def _clear_all_isotopes(self):
        for i in range(10):
            self._clear_isotope_slot(i)

    def _collect_inputs(self) -> CalcInputs:
        """Read and validate the current GUI inputs before calculation."""
        if not self.model:
            raise RuntimeError("Load a workbook first.")

        isotopes: List[Tuple[str, float]] = []
        for i in range(10):
            name = self.selected_isotope_vars[i].get().strip()
            if not name:
                continue
            act_txt = self.selected_activity_vars[i].get().strip()
            try:
                act = float(act_txt)
            except Exception as exc:
                raise ValueError(f"Invalid activity for isotope slot {i+1}: {act_txt}") from exc
            isotopes.append((name, act))

        count_time = float(self.count_time_var.get().strip())
        distance = float(self.distance_var.get().strip())

        shielding = {}
        for m in SHIELDING_MATERIALS:
            txt = self.shield_vars[m].get().strip()
            try:
                shielding[m] = float(txt) if txt else 0.0
            except Exception as exc:
                raise ValueError(f"Invalid shielding thickness for {m}: {txt}") from exc

        if not self.model.has_shielding:
            shielding = {m: 0.0 for m in SHIELDING_MATERIALS}

        return CalcInputs(
            isotopes=isotopes,
            count_time_sec=count_time,
            distance_cm=distance,
            include_background=self.include_bg_var.get(),
            shielding_cm=shielding,
        )

    def _calculate(self):
        """Run the workbook-backed GREASPY calculation and refresh the UI."""
        try:
            inp = self._collect_inputs()
            self.summary_var.set("Calculating spectrum...")
            self.update_idletasks()
            self.results = self.model.compute(inp) if self.model else None
            self._refresh_summary(inp)
            self._refresh_plot()
        except Exception as exc:
            messagebox.showerror("Calculation Error", str(exc))

    def _fmt(self, v: Optional[float]) -> str:
        if v is None:
            return "-"
        return f"{v:.6E}"

    def _refresh_summary(self, inp: CalcInputs):
        if not self.results:
            return
        r = self.results
        total_shield = sum(max(0.0, inp.shielding_cm.get(m, 0.0)) for m in SHIELDING_MATERIALS)

        text = [
            "Unshielded Values (Approximate)",
            f"Dose Rate at 30 cm (R/hr): {self._fmt(r.unshielded_dose_30cm)}",
            f"Gammas Emitted per Second: {self._fmt(r.unshielded_gammas_emitted_s)}",
            f"Gammas Incident on Detector per Second: {self._fmt(r.unshielded_gammas_incident_s)}",
            f"Dead Time (%): {r.unshielded_dead_time_pct:.4f}",
            f"Real Counting Time (sec): {r.unshielded_real_time_sec:.6f}",
            f"Distance for 20% Dead Time (cm): {r.unshielded_distance_20pct_cm:.6f}",
            "",
            "Shielded Values (Approximate)",
        ]

        if self.model and self.model.has_shielding and total_shield > 0:
            text.extend(
                [
                    f"Dose Rate at 30 cm (R/hr): {self._fmt(r.shielded_dose_30cm)}",
                    f"Gammas Emitted per Second: {self._fmt(r.shielded_gammas_emitted_s)}",
                    f"Gammas Incident on Detector per Second: {self._fmt(r.shielded_gammas_incident_s)}",
                    f"Dead Time (%): {r.shielded_dead_time_pct:.4f}",
                    f"Real Counting Time (sec): {r.shielded_real_time_sec:.6f}",
                    f"Distance for 20% Dead Time (cm): {r.shielded_distance_20pct_cm:.6f}",
                ]
            )
        else:
            text.append("No shielding applied.")

        text.extend(
            [
                "",
                f"PeakEasy Efficiency Cal Factors: A={r.eff_a:.6f}, B={r.eff_b:.9f}",
            ]
        )

        self.summary_var.set("\n".join(text))

    def _refresh_plot(self):
        """Redraw the detector spectrum and isotope-contribution curves."""
        if not self.results or self.ax is None or self.canvas is None:
            return

        r = self.results
        self.ax.clear()
        self.ax.plot(r.energy_kev, r.total_counts, label="Total", linewidth=1.2)

        # Draw active isotope contributions for context (max 10)
        for name, vec in r.per_isotope_counts.items():
            self.ax.plot(r.energy_kev, vec, linewidth=0.6, alpha=0.7, label=name)

        self.ax.set_title("Spectrum Counts vs Energy")
        self.ax.set_xlabel("Energy (keV)")
        self.ax.set_ylabel("Counts")
        self.ax.relim()
        self.ax.autoscale_view()
        self.ax.margins(x=0.01, y=0.05)

        if self.log_scale_var.get():
            self.ax.set_yscale("log")
            positives = [v for v in r.total_counts if v > 0]
            if positives:
                self.ax.set_ylim(bottom=max(min(positives) * 0.8, 1e-12))
        else:
            self.ax.set_yscale("linear")
        self.ax.grid(True, alpha=0.3)
        self.ax.legend(loc="upper right", fontsize=8, ncol=2)
        self.fig.tight_layout()
        self.canvas.draw()

    def _export_spe(self):
        if not self.results:
            messagebox.showinfo("Export .SPE", "Run Calculate first.")
            return

        try:
            inp = self._collect_inputs()
        except Exception as exc:
            messagebox.showerror("Export Error", str(exc))
            return

        save_path = filedialog.asksaveasfilename(
            title="Save PeakEasy .spe",
            defaultextension=".spe",
            filetypes=[("SPE file", "*.spe"), ("Text file", "*.txt")],
        )
        if not save_path:
            return

        lines = self._build_spe_lines(inp, self.results)
        with open(save_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")

        messagebox.showinfo("Export .SPE", f"Saved: {save_path}")

    def _build_spe_lines(self, inp: CalcInputs, r: CalcResults) -> List[str]:
        """Export the current detector spectrum in PeakEasy `.spe` text format."""
        shield_on = sum(max(0.0, inp.shielding_cm.get(m, 0.0)) for m in SHIELDING_MATERIALS) > 0

        dose = r.shielded_dose_30cm if shield_on else r.unshielded_dose_30cm
        dist20 = r.shielded_distance_20pct_cm if shield_on else r.unshielded_distance_20pct_cm
        dead = r.shielded_dead_time_pct if shield_on else r.unshielded_dead_time_pct
        emitted = r.shielded_gammas_emitted_s if shield_on else r.unshielded_gammas_emitted_s
        incident = r.shielded_gammas_incident_s if shield_on else r.unshielded_gammas_incident_s
        real_time = r.shielded_real_time_sec if shield_on else r.unshielded_real_time_sec

        now = _dt.datetime.now().strftime("%m/%d/%Y  %I:%M:%S %p")

        lines = [
            "$SPEC_ID:",
            f"Distance Measured = {inp.distance_cm:g} cm",
            f"Dose Rate (rem/hr) at 30 cm = {dose:.2E}",
            f"Distance (cm) Required for 20 Percent Dead Time = {dist20:.2E}",
            f"Dead Time (Percent) = {dead:.2E}",
            f"Total Gammas Emitted per Second by Source = {emitted:.2E}",
            f"Total Gammas Recorded per Second by Detector = {incident:.2E}",
        ]

        selected = inp.isotopes[:10]
        for i in range(10):
            if i < len(selected) and selected[i][0] and selected[i][1] > 0:
                name, act = selected[i]
                lines.append(f"Isotope {i+1}: {name} = {act:.2E} Ci")
            else:
                lines.append(f"Isotope {i+1}: None")

        lines.append(f"Background Included = {'YES' if inp.include_background else 'NO'}")
        for m in SHIELDING_MATERIALS:
            lines.append(f"Shielding: {m} = {inp.shielding_cm.get(m, 0.0):g} cm")

        lines.extend(
            [
                "$SPEC_REM:",
                "DET# 2",
                "DETDESC# GREASPY Standalone IDM-200-V",
                "AP# Generated by GREASPY",
                "$DATE_MEA:",
                now,
                "$MEAS_TIM:",
                f"{inp.count_time_sec:g} {real_time:.6f}",
                "$DATA:",
                "0 16383",
            ]
        )

        for v in r.total_counts:
            lines.append(f"{v:.6f}")

        lines.extend(
            [
                "$ENER_FIT:",
                f"{r.eff_a:.6f} {r.eff_b:.6f}",
            ]
        )

        return lines

    def _export_csv(self):
        if not self.results:
            messagebox.showinfo("Export CSV", "Run Calculate first.")
            return

        save_path = filedialog.asksaveasfilename(
            title="Save spectrum CSV",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
        )
        if not save_path:
            return

        r = self.results
        iso_names = list(r.per_isotope_counts.keys())

        with open(save_path, "w", encoding="utf-8") as f:
            header = ["energy_keV", "total_counts", "background_counts", "attenuation"] + iso_names
            f.write(",".join(header) + "\n")
            for i in range(NUM_BINS):
                row = [
                    f"{r.energy_kev[i]:.6f}",
                    f"{r.total_counts[i]:.6f}",
                    f"{r.background_counts[i]:.6f}",
                    f"{r.attenuation[i]:.9f}",
                ]
                for n in iso_names:
                    row.append(f"{r.per_isotope_counts[n][i]:.6f}")
                f.write(",".join(row) + "\n")

        messagebox.showinfo("Export CSV", f"Saved: {save_path}")


def main():
    app = GreaseApp()
    try:
        app.mainloop()
    finally:
        if plt is not None:
            plt.close("all")


if __name__ == "__main__":
    main()


# Copyright 2026, Battelle Energy Alliance, LLC. All Rights Reserved.
from __future__ import annotations
"""Extract the GREASPY_CRPYT workbook into the JSON bundle used at runtime."""

import argparse
import json
from pathlib import Path
from typing import Dict, List, Tuple

from openpyxl import load_workbook

from GREASPY import NUM_BINS, SHIELDING_MATERIALS, default_crpyt_workbook_candidates


BUNDLE_NAME = "greaspy_crpyt_rev1"
CHUNK_SIZE = 64


def _resolve_default_workbook_path() -> Path | None:
    """Pick the first GREASY workbook candidate that exists on disk."""
    for candidate in default_crpyt_workbook_candidates():
        if candidate.exists():
            return candidate
    return None


def _extract_energy_kev(fp_ws) -> List[float]:
    """Read the detector energy axis from the FP sheet."""
    energy_kev: List[float] = []
    for (value,) in fp_ws.iter_rows(
        min_row=6,
        max_row=6 + NUM_BINS - 1,
        min_col=1,
        max_col=1,
        values_only=True,
    ):
        energy_kev.append((float(value) * 1000.0) if value is not None else 0.0)
    return energy_kev


def _extract_background_cps(bg_ws) -> List[float]:
    """Read the background spectrum used by GREASY detector calculations."""
    background_cps: List[float] = []
    for (value,) in bg_ws.iter_rows(
        min_row=2,
        max_row=2 + NUM_BINS - 1,
        min_col=1,
        max_col=1,
        values_only=True,
    ):
        background_cps.append(float(value) if value is not None else 0.0)
    return background_cps


def _extract_shielding(shielding_ws) -> Dict[str, object]:
    """Read shielding densities and attenuation coefficients by material."""
    density_row = next(
        shielding_ws.iter_rows(
            min_row=2,
            max_row=2,
            min_col=2,
            max_col=7,
            values_only=True,
        )
    )
    densities = [float(value) if value is not None else 0.0 for value in density_row]

    coefficients: Dict[str, List[float]] = {}
    for index, material in enumerate(SHIELDING_MATERIALS):
        col = 2 + index
        coeffs: List[float] = []
        for (value,) in shielding_ws.iter_rows(
            min_row=4,
            max_row=4 + NUM_BINS - 1,
            min_col=col,
            max_col=col,
            values_only=True,
        ):
            coeffs.append(float(value) if value is not None else 0.0)
        coefficients[material] = coeffs

    return {
        "materials": SHIELDING_MATERIALS,
        "density_g_cc": densities,
        "coefficients": coefficients,
    }


def _extract_defaults(greasy_user_ws) -> Dict[str, object]:
    """Read the GREASY workbook defaults shown in the Spectrum tab."""
    count_time = greasy_user_ws["B2"].value
    distance = greasy_user_ws["B3"].value
    include_background = str(greasy_user_ws["B4"].value or "").strip().upper() == "YES"

    shielding_defaults: Dict[str, float] = {}
    for index, material in enumerate(SHIELDING_MATERIALS, start=7):
        value = greasy_user_ws.cell(index, 2).value
        shielding_defaults[material] = float(value) if value is not None else 0.0

    return {
        "count_time_sec": float(count_time) if count_time is not None else 300.0,
        "distance_cm": float(distance) if distance is not None else 2300.0,
        "include_background": include_background,
        "shielding_cm": shielding_defaults,
    }


def _iter_valid_columns(ws) -> List[Tuple[int, int]]:
    """Return the sheet columns whose headers contain valid ZAIDs."""
    header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    valid_columns: List[Tuple[int, int]] = []
    for col, value in enumerate(header_row[1:], start=2):
        try:
            zaid = int(float(value))
        except Exception:
            continue
        valid_columns.append((col, zaid))
    return valid_columns


def _extract_sheet_spectra(ws, output_dir: Path, zaid_files: Dict[str, str]) -> None:
    """Chunk one GREASY library sheet into JSON files keyed by ZAID."""
    valid_columns = _iter_valid_columns(ws)
    header_rows = list(ws.iter_rows(min_row=1, max_row=4, values_only=True))
    dose_row = header_rows[2]
    gamma_row = header_rows[3]

    spectra_dir = output_dir / "spectra"
    spectra_dir.mkdir(parents=True, exist_ok=True)

    # Chunking avoids thousands of tiny files while still letting runtime only
    # read the spectrum groups it actually needs for the current isotope set.
    for chunk_start in range(0, len(valid_columns), CHUNK_SIZE):
        chunk = valid_columns[chunk_start : chunk_start + CHUNK_SIZE]
        min_col = min(col for col, _zaid in chunk)
        max_col = max(col for col, _zaid in chunk)
        indexed_columns = list(chunk)
        spectra_by_zaid: Dict[int, List[float]] = {
            zaid: [] for _col, zaid in indexed_columns
        }

        for row_values in ws.iter_rows(
            min_row=6,
            max_row=6 + NUM_BINS - 1,
            min_col=min_col,
            max_col=max_col,
            values_only=True,
        ):
            for col, zaid in indexed_columns:
                offset = col - min_col
                value = row_values[offset] if offset < len(row_values) else None
                spectra_by_zaid[zaid].append(float(value) if value is not None else 0.0)

        entries: Dict[str, Dict[str, object]] = {}
        for col, zaid in indexed_columns:
            entries[str(zaid)] = {
                "dose_per_ci": float(dose_row[col - 1] or 0.0),
                "gammas_per_ci": float(gamma_row[col - 1] or 0.0),
                "spectrum_cps": spectra_by_zaid[zaid],
            }

        relative_path = Path("spectra") / f"{ws.title.lower()}_{(chunk_start // CHUNK_SIZE) + 1:03d}.json"
        payload = {"sheet": ws.title, "entries": entries}
        with (output_dir / relative_path).open("w", encoding="utf-8") as handle:
            json.dump(payload, handle, separators=(",", ":"))

        for _col, zaid in indexed_columns:
            zaid_files[str(zaid)] = str(relative_path)


def extract_bundle(workbook_path: Path, output_dir: Path) -> Path:
    """Write the GREASY bundle consumed by the standalone spectrum model."""
    output_dir.mkdir(parents=True, exist_ok=True)
    spectra_dir = output_dir / "spectra"
    spectra_dir.mkdir(parents=True, exist_ok=True)
    # Remove stale chunk files so reruns cannot leave mixed bundle layouts behind.
    for stale_path in spectra_dir.glob("*.json"):
        stale_path.unlink(missing_ok=True)

    wb = load_workbook(workbook_path, data_only=True, read_only=True)

    greasy_user_ws = wb["GREASY USER"]
    fp_ws = wb["FP"]
    other_ws = wb["Other"]
    bg_ws = wb["BG"]
    shielding_ws = wb["Shielding"] if "Shielding" in wb.sheetnames else None

    zaid_files: Dict[str, str] = {}
    _extract_sheet_spectra(fp_ws, output_dir, zaid_files)
    _extract_sheet_spectra(other_ws, output_dir, zaid_files)

    metadata = {
        "defaults": _extract_defaults(greasy_user_ws),
        "energy_kev": _extract_energy_kev(fp_ws),
        "background_cps": _extract_background_cps(bg_ws),
        "shielding": _extract_shielding(shielding_ws) if shielding_ws is not None else {},
        "zaid_files": zaid_files,
    }
    metadata_path = output_dir / "metadata.json"
    with metadata_path.open("w", encoding="utf-8") as handle:
        json.dump(metadata, handle, separators=(",", ":"))

    manifest = {
        "version": 1,
        "metadata_file": "metadata.json",
        "source_workbook": str(workbook_path),
        "spectra_dir": "spectra",
    }
    manifest_path = output_dir / "manifest.json"
    with manifest_path.open("w", encoding="utf-8") as handle:
        json.dump(manifest, handle, indent=2)
    return manifest_path


def ensure_bundle(workbook_path: Path | None, output_dir: Path) -> None:
    """Build the bundle on demand the first time the standalone app needs it."""
    manifest_path = output_dir / "manifest.json"
    if manifest_path.exists():
        return
    if workbook_path is None or not workbook_path.exists():
        raise FileNotFoundError(
            f"GREASPY_CRPYT workbook not found; cannot build spectrum bundle in {output_dir}."
        )
    extract_bundle(workbook_path, output_dir)


def main() -> None:
    """CLI entry point for regenerating the GREASY JSON bundle."""
    parser = argparse.ArgumentParser(description="Extract GREASPY_CRPYT workbook data into a JSON bundle.")
    parser.add_argument(
        "--workbook",
        type=Path,
        default=_resolve_default_workbook_path(),
        help="Path to GREASPY_CRPYT_Rev1.xlsx",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path(__file__).resolve().parent / "data" / BUNDLE_NAME,
        help="Directory to receive the extracted GREASY JSON bundle",
    )
    args = parser.parse_args()
    if args.workbook is None:
        raise FileNotFoundError("No GREASPY_CRPYT_Rev1.xlsx path was provided and no default workbook was found.")
    manifest_path = extract_bundle(args.workbook, args.output_dir)
    print(f"Wrote GREASY bundle manifest to {manifest_path}")


if __name__ == "__main__":
    main()

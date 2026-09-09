# Copyright 2026, Battelle Energy Alliance, LLC. All Rights Reserved.
from __future__ import annotations
"""Extract the CRPYT workbook into the JSON bundle used by the standalone app."""

import argparse
import json
import re
from pathlib import Path
from typing import Any, Dict, List

from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries



BUNDLE_NAME = "crpyt_irdff_rev1"
CACHED_RUNTIME_SHEETS = ("PR Coeff Irr", "PR Coeff Decay")
COMBINED_REACTOR_UI_LABEL = "Combined Spectrum"
COMBINED_REACTOR_WORKBOOK_VALUE = "Combo"
DEFAULT_COMBINED_SPECTRUM_NAME = "TRIGA"
DEFAULT_REACTOR = "TRIGA"
DEFAULT_FLUENCE_RATE = 1.0e12
DEFAULT_IRRADIATION_TIME = 3600.0
DEFAULT_DECAY_TIME = 86400.0
DEFAULT_ENERGY_MIN = 1.0e-9
DEFAULT_ENERGY_MAX = 20.0
DEFAULT_COVER_MATERIAL = 0.0
DEFAULT_ELEMENT_MASS = 0.0
DEFAULT_ELEMENTS_MIN_MASS = 1.0e-15
DEFAULT_ISOTOPE_MIN_ACTIVITY_CI = 1.0e-11
DEFAULT_GAMMA_MIN_RATE = 1.0


def _slugify(name: str) -> str:
    """Convert a sheet name into a stable JSON file stem."""
    return re.sub(r"[^a-z0-9]+", "_", name.lower()).strip("_")


def _normalize_formula_value(value: Any) -> Any:
    """Patch workbook-specific formula quirks into forms the parser accepts."""
    text = getattr(value, "text", value)
    if isinstance(text, str) and "_xlfn.IFNA" in text:
        return text.replace("_xlfn.IFNA", "IFERROR")
    return text


def _serialize_sheet(ws, *, patch_formulas: bool) -> Dict[str, Any]:
    """Serialize one worksheet into a compact row-oriented JSON payload."""
    rows: List[List[Any]] = []
    formula_cells: List[List[int]] = []
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        serialized_row: List[Any] = []
        last_non_empty = 0
        current_row_index = row[0].row if row else len(rows) + 1
        for cell in row:
            value = _normalize_formula_value(cell.value) if patch_formulas else cell.value
            serialized_row.append(value)
            if value is not None:
                last_non_empty = cell.column
            if isinstance(value, str) and value.startswith("=") and value.strip() != "=":
                formula_cells.append([current_row_index, cell.column])
        rows.append(serialized_row[:last_non_empty])
    return {
        "name": ws.title,
        "max_row": ws.max_row,
        "max_col": ws.max_column,
        "rows": rows,
        "formula_cells": formula_cells,
    }


def _extract_metadata(workbook_path: Path) -> Dict[str, Any]:
    """Pull out the reference data and GUI defaults used outside the formula engine."""
    wb_formula = load_workbook(workbook_path, data_only=False)
    wb = load_workbook(workbook_path, data_only=True)
    inputs_ws_formula = wb_formula["Choose Irradiation Parameters"]
    inputs_ws = wb["Choose Irradiation Parameters"]

    def resolve_list_values(ws, formula: str) -> List[str]:
        """Resolve an Excel data-validation list into concrete option values."""
        formula = formula.strip()
        if formula.startswith("="):
            formula = formula[1:]
        if "!" in formula:
            sheet_name, cell_range = formula.split("!", 1)
            sheet_name = sheet_name.strip("'\"")
            target_ws = ws.parent[sheet_name]
        else:
            cell_range = formula
            target_ws = ws
        min_col, min_row, max_col, max_row = range_boundaries(cell_range)
        values: List[str] = []
        for row in target_ws.iter_rows(
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
        ):
            for cell in row:
                if cell.value not in (None, ""):
                    values.append(str(cell.value))
        return values

    def fallback_reactor_options() -> List[str]:
        """Fallback scan used if validation metadata is missing or malformed."""
        options: List[str] = []
        seen = set()
        for row in range(4, 128):
            value = inputs_ws_formula.cell(row=row, column=90).value
            if value in (None, ""):
                continue
            option = str(value).strip()
            if option and option not in seen:
                seen.add(option)
                options.append(option)
        return options

    # Most GUI defaults mirror the workbook input sheet, so harvest them once
    # here and store them in metadata.json for the standalone app.
    reactor_options: List[str] = []
    for validation in inputs_ws_formula.data_validations.dataValidation:
        sqref_parts = str(validation.sqref).split()
        if "B2" in sqref_parts:
            reactor_options = resolve_list_values(inputs_ws_formula, validation.formula1)
            break
    if not reactor_options:
        reactor_options = fallback_reactor_options()
    ui_reactor_options = [
        COMBINED_REACTOR_UI_LABEL if option == COMBINED_REACTOR_WORKBOOK_VALUE else option
        for option in reactor_options
    ]
    spectrum_options = [
        option for option in ui_reactor_options if option != COMBINED_REACTOR_UI_LABEL
    ]

    composition_rows: List[Dict[str, Any]] = []
    for row in range(2, 81):
        atomic_number = inputs_ws.cell(row=row, column=5).value
        element = inputs_ws.cell(row=row, column=6).value
        if atomic_number in (None, "") or element in (None, ""):
            continue
        composition_rows.append(
            {
                "row": row,
                "atomic_number": atomic_number,
                "element": element,
                "default_mass": DEFAULT_ELEMENT_MASS,
            }
        )

    spectrum_rows: List[Dict[str, Any]] = []
    for row in range(21, 31):
        spectrum_rows.append(
            {
                "row": row,
                "label": inputs_ws.cell(row=row, column=1).value or f"Spectrum #{row - 20}",
                "default_name": inputs_ws.cell(row=row, column=2).value or DEFAULT_COMBINED_SPECTRUM_NAME,
                "default_fraction": float(inputs_ws.cell(row=row, column=3).value or 0.0),
            }
        )

    gui_config = {
        "workbook_path": str(workbook_path),
        "reactor_options": ui_reactor_options,
        "spectrum_options": spectrum_options,
        "default_reactor": DEFAULT_REACTOR,
        "default_fluence_rate": DEFAULT_FLUENCE_RATE,
        "default_irradiation_time": DEFAULT_IRRADIATION_TIME,
        "default_decay_time": DEFAULT_DECAY_TIME,
        "default_energy_min": DEFAULT_ENERGY_MIN,
        "default_energy_max": DEFAULT_ENERGY_MAX,
        "cover_materials": [
            {
                "cell": "B8",
                "label": str(inputs_ws.cell(row=8, column=1).value or "Cd Thickness (mm)"),
                "default": DEFAULT_COVER_MATERIAL,
            },
            {
                "cell": "B9",
                "label": str(inputs_ws.cell(row=9, column=1).value or "Gd Thickness (mm)"),
                "default": DEFAULT_COVER_MATERIAL,
            },
            {
                "cell": "B10",
                "label": str(inputs_ws.cell(row=10, column=1).value or "Boron Thickness (mm)"),
                "default": DEFAULT_COVER_MATERIAL,
            },
            {
                "cell": "B11",
                "label": str(inputs_ws.cell(row=11, column=1).value or "Boron Carbide (B4C) Thickness (mm)"),
                "default": DEFAULT_COVER_MATERIAL,
            },
        ],
        "spectrum_rows": spectrum_rows,
        "composition_rows": composition_rows,
        "result_thresholds": {
            "elements_min_mass": DEFAULT_ELEMENTS_MIN_MASS,
            "isotopes_min_atoms": float(wb["Results - Isotopes"]["B10"].value or 0.0),
            "isotopes_min_activity_ci": DEFAULT_ISOTOPE_MIN_ACTIVITY_CI,
            "gamma_min_rate": DEFAULT_GAMMA_MIN_RATE,
        },
    }

    branching = wb["Branching Ratios"]
    isotope_properties = wb["Isotope Properties"]
    gamma_summary = wb["Gammas per Decay"]
    beta_summary = wb["Betas per Decay"]
    gamma_lines_ws = wb["Gamma"]
    results_elements = wb["Results - Elements"]

    isotope_to_element: Dict[str, str] = {}
    half_life_by_zaid: Dict[str, float] = {}
    for row in range(2, branching.max_row + 1):
        parent_name = branching.cell(row, 2).value
        parent_zaid = branching.cell(row, 1).value
        half_life = branching.cell(row, 5).value
        element = branching.cell(row, 12).value
        if parent_name is not None and element is not None and str(parent_name) not in isotope_to_element:
            isotope_to_element[str(parent_name)] = str(element)
        if parent_zaid is not None and half_life is not None:
            key = str(int(parent_zaid))
            half_life_by_zaid.setdefault(key, float(half_life))

    molar_mass_by_zaid: Dict[str, float] = {}
    for row in range(2, isotope_properties.max_row + 1):
        zaid = isotope_properties.cell(row, 1).value
        mass = isotope_properties.cell(row, 4).value
        if zaid is not None and mass is not None:
            molar_mass_by_zaid[str(int(zaid))] = float(mass)

    gamma_per_ci_by_zaid: Dict[str, float] = {}
    gamma_dose_per_ci_by_zaid: Dict[str, float] = {}
    for row in range(2, gamma_summary.max_row + 1):
        zaid = gamma_summary.cell(row, 1).value
        if zaid is None:
            continue
        key = str(int(zaid))
        gamma_per_ci_by_zaid[key] = float(gamma_summary.cell(row, 9).value or 0.0)
        gamma_dose_per_ci_by_zaid[key] = float(gamma_summary.cell(row, 14).value or 0.0)

    beta_per_ci_by_zaid: Dict[str, float] = {}
    beta_dose_per_ci_by_zaid: Dict[str, float] = {}
    for row in range(2, beta_summary.max_row + 1):
        zaid = beta_summary.cell(row, 1).value
        if zaid is None:
            continue
        key = str(int(zaid))
        beta_per_ci_by_zaid[key] = float(beta_summary.cell(row, 7).value or 0.0)
        beta_dose_per_ci_by_zaid[key] = float(beta_summary.cell(row, 11).value or 0.0)

    gamma_lines: List[List[Any]] = []
    for row in range(2, gamma_lines_ws.max_row + 1):
        zaid = gamma_lines_ws.cell(row, 1).value
        isotope = gamma_lines_ws.cell(row, 2).value
        energy = gamma_lines_ws.cell(row, 3).value
        intensity = gamma_lines_ws.cell(row, 5).value
        if zaid is None or intensity is None:
            continue
        gamma_lines.append([int(zaid), str(isotope), float(energy), float(intensity)])

    element_order: List[str] = []
    seen_elements = set()
    for row in range(3, 121):
        element = results_elements.cell(row, 14).value
        if element and element not in seen_elements:
            seen_elements.add(element)
            element_order.append(str(element))

    return {
        "gui_config": gui_config,
        "lookups": {
            "isotope_to_element": isotope_to_element,
            "half_life_by_zaid": half_life_by_zaid,
            "molar_mass_by_zaid": molar_mass_by_zaid,
            "gamma_per_ci_by_zaid": gamma_per_ci_by_zaid,
            "gamma_dose_per_ci_by_zaid": gamma_dose_per_ci_by_zaid,
            "beta_per_ci_by_zaid": beta_per_ci_by_zaid,
            "beta_dose_per_ci_by_zaid": beta_dose_per_ci_by_zaid,
        },
        "gamma_lines": gamma_lines,
        "element_order": element_order,
    }


def extract_bundle(workbook_path: Path, output_dir: Path) -> Path:
    """Write the CRPYT workbook bundle consumed by the standalone runtime."""
    workbook_path = workbook_path.resolve()
    output_dir = output_dir.resolve()
    formula_dir = output_dir / "formula"
    cached_dir = output_dir / "cached"
    output_dir.mkdir(parents=True, exist_ok=True)
    formula_dir.mkdir(parents=True, exist_ok=True)
    cached_dir.mkdir(parents=True, exist_ok=True)

    wb_formula = load_workbook(workbook_path, data_only=False)
    wb_cached = load_workbook(workbook_path, data_only=True)

    # Keep the runtime bundle split into formula sheets, cached sheets, and
    # metadata so the engine can instantiate only what it needs.
    formula_sheet_files: List[str] = []
    for ws in wb_formula.worksheets:
        payload = _serialize_sheet(ws, patch_formulas=True)
        filename = f"{_slugify(ws.title)}.json"
        relative_path = Path("formula") / filename
        with (output_dir / relative_path).open("w", encoding="utf-8") as handle:
            json.dump(payload, handle, separators=(",", ":"))
        formula_sheet_files.append(relative_path.as_posix())

    cached_sheet_files: List[str] = []
    for sheet_name in CACHED_RUNTIME_SHEETS:
        ws = wb_cached[sheet_name]
        payload = _serialize_sheet(ws, patch_formulas=False)
        payload["formula_cells"] = []
        filename = f"{_slugify(ws.title)}.json"
        relative_path = Path("cached") / filename
        with (output_dir / relative_path).open("w", encoding="utf-8") as handle:
            json.dump(payload, handle, separators=(",", ":"))
        cached_sheet_files.append(relative_path.as_posix())

    metadata = _extract_metadata(workbook_path)
    metadata_path = output_dir / "metadata.json"
    with metadata_path.open("w", encoding="utf-8") as handle:
        json.dump(metadata, handle, separators=(",", ":"))

    manifest = {
        "version": 1,
        "template_workbook": "../CRPYT_IRDFF_Rev1.xlsx",
        "metadata_file": "metadata.json",
        "formula_sheets": formula_sheet_files,
        "cached_sheets": cached_sheet_files,
    }
    manifest_path = output_dir / "manifest.json"
    with manifest_path.open("w", encoding="utf-8") as handle:
        json.dump(manifest, handle, indent=2)
    return manifest_path


def main() -> None:
    """CLI entry point for regenerating the CRPYT JSON bundle."""
    parser = argparse.ArgumentParser(description="Extract the CRPYT IRDFF workbook into JSON bundles.")
    parser.add_argument(
        "--workbook",
        type=Path,
        default=Path(__file__).resolve().with_name("CRPYT_IRDFF_Rev1.xlsx"),
        help="Path to CRPYT_IRDFF_Rev1.xlsx",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path(__file__).resolve().parent / "data" / BUNDLE_NAME,
        help="Directory to receive the extracted JSON bundle",
    )
    args = parser.parse_args()
    manifest_path = extract_bundle(args.workbook, args.output_dir)
    print(f"Wrote bundle manifest to {manifest_path}")


if __name__ == "__main__":
    main()

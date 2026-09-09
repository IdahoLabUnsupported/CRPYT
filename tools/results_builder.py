# Copyright 2026, Battelle Energy Alliance, LLC. All Rights Reserved.
from __future__ import annotations
"""Native result aggregation for the standalone CRPYT runtime.

The original workbook computes result sheets with thousands of formulas. This
module replaces the slowest result tabs with direct Python aggregation while
preserving the same output structure expected by the GUI.
"""

import json
import math
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List

from openpyxl.utils import column_index_from_string


AVOGADRO = 6.022e23
CI_TO_BQ = 3.7e10
STABLE_HALF_LIFE_SENTINEL = 1.0e30
PR_COEFF_IRR_SHEET = "PR Coeff Irr"
PR_COEFF_DECAY_SHEET = "PR Coeff Decay"
ISOTOPE_PROPERTIES_SHEET = "Isotope Properties"
STABLE_ISOTOPE_PROPERTIES_SHEET = "Stable Isotope Properties"
BRANCHING_RATIOS_SHEET = "Branching Ratios"
PR_COEFF_DATA_START_ROW = 5
PR_COEFF_IRR_BOI_PAIR = (
    column_index_from_string("A"),
    column_index_from_string("B"),
)
PR_COEFF_DECAY_AVERAGE_PAIRS = [
    (column_index_from_string("AP"), column_index_from_string("AQ")),
    (column_index_from_string("AR"), column_index_from_string("AS")),
    (column_index_from_string("AT"), column_index_from_string("AU")),
    (column_index_from_string("AV"), column_index_from_string("AW")),
    (column_index_from_string("AX"), column_index_from_string("AY")),
    (column_index_from_string("AZ"), column_index_from_string("BA")),
    (column_index_from_string("BB"), column_index_from_string("BC")),
    (column_index_from_string("BD"), column_index_from_string("BE")),
    (column_index_from_string("BF"), column_index_from_string("BG")),
    (column_index_from_string("BH"), column_index_from_string("BI")),
    (column_index_from_string("BJ"), column_index_from_string("BK")),
    (column_index_from_string("BL"), column_index_from_string("BM")),
    (column_index_from_string("BN"), column_index_from_string("BO")),
    (column_index_from_string("BP"), column_index_from_string("BQ")),
    (column_index_from_string("BR"), column_index_from_string("BS")),
    (column_index_from_string("BT"), column_index_from_string("BU")),
    (column_index_from_string("BV"), column_index_from_string("BW")),
]


@dataclass(frozen=True)
class ReferenceMetadata:
    """Reference data extracted once from the workbook bundle metadata."""
    gui_config: Dict[str, Any]
    isotope_to_element: Dict[str, str]
    half_life_by_zaid: Dict[int, float]
    molar_mass_by_zaid: Dict[int, float]
    gamma_per_ci_by_zaid: Dict[int, float]
    gamma_dose_per_ci_by_zaid: Dict[int, float]
    beta_per_ci_by_zaid: Dict[int, float]
    beta_dose_per_ci_by_zaid: Dict[int, float]
    gamma_lines: List[List[Any]]
    element_order: List[str]


def load_reference_metadata(metadata_path: Path) -> ReferenceMetadata:
    """Load the extracted metadata JSON and normalize ZAID-keyed lookups."""
    with metadata_path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)
    return ReferenceMetadata(
        gui_config=payload["gui_config"],
        isotope_to_element=payload["lookups"]["isotope_to_element"],
        half_life_by_zaid={int(key): value for key, value in payload["lookups"]["half_life_by_zaid"].items()},
        molar_mass_by_zaid={int(key): value for key, value in payload["lookups"]["molar_mass_by_zaid"].items()},
        gamma_per_ci_by_zaid={int(key): value for key, value in payload["lookups"]["gamma_per_ci_by_zaid"].items()},
        gamma_dose_per_ci_by_zaid={int(key): value for key, value in payload["lookups"]["gamma_dose_per_ci_by_zaid"].items()},
        beta_per_ci_by_zaid={int(key): value for key, value in payload["lookups"]["beta_per_ci_by_zaid"].items()},
        beta_dose_per_ci_by_zaid={int(key): value for key, value in payload["lookups"]["beta_dose_per_ci_by_zaid"].items()},
        gamma_lines=payload["gamma_lines"],
        element_order=payload["element_order"],
    )


def _safe_float(value: Any, default: float = 0.0) -> float:
    try:
        return float(value)
    except Exception:
        return default


def _safe_int(value: Any) -> int | None:
    try:
        return int(float(value))
    except Exception:
        return None


def _passes_activity_filter(activity_ci: float, min_activity_ci: float) -> bool:
    """Match Excel's zero-threshold behavior for stable isotope rows."""
    if activity_ci > min_activity_ci:
        return True
    return min_activity_ci <= 0.0 and activity_ci == 0.0


def _build_isotope_name_lookup(engine: Any) -> Dict[int, str]:
    """Merge the workbook's isotope naming sources into one ZAID lookup."""
    lookup: Dict[int, str] = {}

    isotope_properties = engine.wb_formula[ISOTOPE_PROPERTIES_SHEET]
    for row in range(2, isotope_properties.max_row + 1):
        zaid = _safe_int(isotope_properties._get_value(row, 1))
        name = isotope_properties._get_value(row, 2)
        if zaid is not None and name not in (None, "") and zaid not in lookup:
            lookup[zaid] = str(name)

    stable_properties = engine.wb_formula[STABLE_ISOTOPE_PROPERTIES_SHEET]
    for row in range(3, stable_properties.max_row + 1):
        zaid = _safe_int(stable_properties._get_value(row, 1))
        name = stable_properties._get_value(row, 3)
        if zaid is not None and name not in (None, "") and zaid not in lookup:
            lookup[zaid] = str(name)

    branching = engine.wb_formula[BRANCHING_RATIOS_SHEET]
    for row in range(2, branching.max_row + 1):
        zaid = _safe_int(branching._get_value(row, 1))
        name = branching._get_value(row, 2)
        if zaid is not None and name not in (None, "") and zaid not in lookup:
            lookup[zaid] = str(name)

    return lookup


def _accumulate_average_pair(
    engine: Any,
    sheet_name: str,
    key_col: int,
    value_col: int,
    results: Dict[int, float],
) -> None:
    """Average repeated ZAID/value pairs from one workbook column pair."""
    sums: Dict[int, float] = {}
    counts: Dict[int, int] = {}
    worksheet = engine.wb_formula[sheet_name]
    for row in range(PR_COEFF_DATA_START_ROW, worksheet.max_row + 1):
        zaid = _safe_int(engine._get_cell(sheet_name, row, key_col))
        if zaid is None:
            continue
        value = _safe_float(engine._get_cell(sheet_name, row, value_col))
        sums[zaid] = sums.get(zaid, 0.0) + value
        counts[zaid] = counts.get(zaid, 0) + 1
    for zaid, total in sums.items():
        count = counts.get(zaid, 0)
        if count:
            results[zaid] = results.get(zaid, 0.0) + (total / count)


def _accumulate_average_pairs(
    engine: Any,
    sheet_name: str,
    pairs: List[tuple[int, int]],
    results: Dict[int, float],
) -> None:
    """Apply the same averaging logic across many workbook column pairs."""
    for key_col, value_col in pairs:
        _accumulate_average_pair(engine, sheet_name, key_col, value_col, results)


def _aggregate_isotope_atoms(engine: Any) -> tuple[Dict[int, float], Dict[int, float]]:
    """Build BOI and EOID atom inventories from the PR coefficient sheets."""
    boi_atoms_by_zaid: Dict[int, float] = {}
    eoid_atoms_by_zaid: Dict[int, float] = {}

    _accumulate_average_pair(
        engine,
        PR_COEFF_IRR_SHEET,
        PR_COEFF_IRR_BOI_PAIR[0],
        PR_COEFF_IRR_BOI_PAIR[1],
        boi_atoms_by_zaid,
    )
    _accumulate_average_pairs(engine, PR_COEFF_DECAY_SHEET, PR_COEFF_DECAY_AVERAGE_PAIRS, eoid_atoms_by_zaid)

    return boi_atoms_by_zaid, eoid_atoms_by_zaid


def _build_isotope_rows(
    engine: Any,
    metadata: ReferenceMetadata,
    min_atoms: float,
    min_activity_ci: float,
) -> Dict[str, Any]:
    """Convert aggregated atoms into isotope rows and summary totals."""
    all_isotope_entries: List[Dict[str, Any]] = []
    activity_by_zaid: Dict[int, float] = {}
    isotope_name_by_zaid = _build_isotope_name_lookup(engine)
    boi_atoms_by_zaid, eoid_atoms_by_zaid = _aggregate_isotope_atoms(engine)

    total_activity_ci = 0.0
    total_gamma_rate = 0.0
    total_gamma_dose = 0.0
    total_beta_rate = 0.0
    total_beta_dose = 0.0

    # Each isotope row is built directly from the aggregated atom totals rather
    # than by replaying the workbook's Results - Isotopes formulas.
    for zaid_int, eoid_value in eoid_atoms_by_zaid.items():
        if not (eoid_value > min_atoms):
            continue

        isotope_name = isotope_name_by_zaid.get(zaid_int)
        if isotope_name in (None, ""):
            continue

        boi_value = boi_atoms_by_zaid.get(zaid_int, 0.0)
        element = metadata.isotope_to_element.get(isotope_name, "")
        molar_mass = metadata.molar_mass_by_zaid.get(zaid_int, 0.0)
        half_life = metadata.half_life_by_zaid.get(zaid_int, 1.0e99)

        boi_mass = boi_value * molar_mass / AVOGADRO if molar_mass else 0.0
        eoid_mass = eoid_value * molar_mass / AVOGADRO if molar_mass else 0.0
        activity_ci = 0.0 if half_life > STABLE_HALF_LIFE_SENTINEL else eoid_value * math.log(2.0) / half_life / CI_TO_BQ
        gamma_rate = metadata.gamma_per_ci_by_zaid.get(zaid_int, 0.0) * activity_ci
        gamma_dose = metadata.gamma_dose_per_ci_by_zaid.get(zaid_int, 0.0) * activity_ci
        beta_rate = metadata.beta_per_ci_by_zaid.get(zaid_int, 0.0) * activity_ci
        beta_dose = metadata.beta_dose_per_ci_by_zaid.get(zaid_int, 0.0) * activity_ci

        activity_by_zaid[zaid_int] = activity_ci
        total_activity_ci += activity_ci
        total_gamma_rate += gamma_rate
        total_gamma_dose += gamma_dose
        total_beta_rate += beta_rate
        total_beta_dose += beta_dose

        all_isotope_entries.append(
            {
                "activity_ci": activity_ci,
                "zaid": zaid_int,
                "reference_row": [zaid_int, isotope_name, boi_value, eoid_value],
                "main_row": [
                    element,
                    zaid_int,
                    isotope_name,
                    boi_value,
                    eoid_value,
                    boi_mass,
                    eoid_mass,
                    activity_ci,
                    gamma_rate,
                    gamma_dose,
                    beta_rate,
                    beta_dose,
                ],
            }
        )

    all_isotope_entries.sort(key=lambda entry: (-entry["activity_ci"], entry["zaid"]))
    filtered_entries = [
        entry
        for entry in all_isotope_entries
        if _passes_activity_filter(_safe_float(entry["activity_ci"]), min_activity_ci)
    ]
    reference_rows = [entry["reference_row"] for entry in filtered_entries]
    main_rows = [entry["main_row"] for entry in filtered_entries]

    return {
        "rows": main_rows,
        "reference_rows": reference_rows,
        "all_rows": [entry["main_row"] for entry in all_isotope_entries],
        "activity_by_zaid": activity_by_zaid,
        "summary": {
            "Total Activity (Ci)": total_activity_ci,
            "Total Activity (Bq)": total_activity_ci * CI_TO_BQ,
            "Gamma-Ray Emission Rate (Bq)": total_gamma_rate,
            "Gamma Dose Rate (R/hr) at 30 cm": total_gamma_dose,
            "Beta Emission Rate (Bq)": total_beta_rate,
            "Beta Dose Rate (R/hr) at Contact": total_beta_dose,
        },
    }


def _build_element_rows(
    isotope_rows: List[List[Any]],
    metadata: ReferenceMetadata,
    min_mass: float,
) -> Dict[str, List[List[Any]]]:
    """Roll isotope masses up to element totals for BOI and EOID."""
    boi_by_element: Dict[str, float] = {}
    eoid_by_element: Dict[str, float] = {}
    for row in isotope_rows:
        element = str(row[0])
        boi_by_element[element] = boi_by_element.get(element, 0.0) + _safe_float(row[5])
        eoid_by_element[element] = eoid_by_element.get(element, 0.0) + _safe_float(row[6])

    all_rows: List[List[Any]] = []
    for element in metadata.element_order:
        all_rows.append(
            [
                element,
                boi_by_element.get(element, 0.0),
                eoid_by_element.get(element, 0.0),
            ]
        )
    filtered_rows = [row for row in all_rows if row[2] > min_mass]
    return {
        "filtered_rows": filtered_rows,
        "all_rows": all_rows,
    }


def _build_gamma_rows(
    metadata: ReferenceMetadata,
    activity_by_zaid: Dict[int, float],
    total_gamma_rate: float,
    min_rate: float,
) -> Dict[str, Any]:
    """Filter and sort the gamma-line library against the current activities."""
    gamma_entries: List[Dict[str, Any]] = []
    filtered_total = 0.0
    for zaid, isotope, energy_kev, intensity in metadata.gamma_lines:
        zaid_int = int(zaid)
        rate = _safe_float(intensity) * CI_TO_BQ * activity_by_zaid.get(zaid_int, 0.0)
        if rate > min_rate:
            filtered_total += rate
            gamma_entries.append(
                {
                    "rate": rate,
                    "zaid": zaid_int,
                    "energy": _safe_float(energy_kev),
                    "row": [zaid_int, str(isotope), _safe_float(energy_kev), rate],
                }
            )

    gamma_entries.sort(key=lambda entry: (-entry["rate"], entry["energy"], entry["zaid"]))
    filtered_rows = [entry["row"] for entry in gamma_entries]

    fraction = filtered_total / total_gamma_rate if total_gamma_rate else 0.0
    return {
        "rows": filtered_rows,
        "summary": {
            "Total Gammas Emitted Per Second": total_gamma_rate,
            "Minimum Gamma Ray Emission Rate (Bq)": min_rate,
            "Fraction Above Threshold": fraction,
        },
    }


def build_results(
    engine: Any,
    metadata: ReferenceMetadata,
    *,
    elements_min_mass: float,
    isotopes_min_atoms: float,
    isotopes_min_activity_ci: float,
    gamma_min_rate: float,
) -> Dict[str, Any]:
    """Build the GUI-ready result payload from one configured engine state."""
    isotope_payload = _build_isotope_rows(
        engine,
        metadata,
        isotopes_min_atoms,
        isotopes_min_activity_ci,
    )
    element_payload = _build_element_rows(
        isotope_payload["all_rows"],
        metadata,
        elements_min_mass,
    )
    gamma_payload = _build_gamma_rows(
        metadata,
        isotope_payload["activity_by_zaid"],
        isotope_payload["summary"]["Gamma-Ray Emission Rate (Bq)"],
        gamma_min_rate,
    )

    return {
        "isotopes_summary": isotope_payload["summary"],
        "gamma_summary": gamma_payload["summary"],
        "elements_filtered": element_payload["filtered_rows"],
        "elements_all": element_payload["all_rows"],
        "isotopes_all": isotope_payload["all_rows"],
        "isotopes_main": isotope_payload["rows"],
        "isotopes_reference": isotope_payload["reference_rows"],
        "gamma_table": gamma_payload["rows"],
    }


def aggregate_results(
    payloads: List[Dict[str, Any]],
    metadata: ReferenceMetadata,
    *,
    elements_min_mass: float,
    isotopes_min_atoms: float,
    isotopes_min_activity_ci: float,
    gamma_min_rate: float,
) -> Dict[str, Any]:
    """Combine many independent CRPYT runs into one GUI-ready result payload.

    Element-by-element covers are represented by splitting the active elements
    into independent engine runs, each with one cover combination. The
    irradiation and decay network is linear in this model, so the final isotope
    inventories can be added back together before the usual element and gamma
    summaries are built.
    """

    combined_by_zaid: Dict[int, Dict[str, Any]] = {}
    total_activity_ci = 0.0
    total_gamma_rate = 0.0
    total_gamma_dose = 0.0
    total_beta_rate = 0.0
    total_beta_dose = 0.0

    for payload in payloads:
        for row in payload.get("isotopes_all", payload.get("isotopes_main", [])):
            zaid = _safe_int(row[1])
            if zaid is None:
                continue

            entry = combined_by_zaid.setdefault(
                zaid,
                {
                    "element": str(row[0]),
                    "isotope": str(row[2]),
                    "boi_atoms": 0.0,
                    "eoid_atoms": 0.0,
                    "boi_mass": 0.0,
                    "eoid_mass": 0.0,
                    "activity_ci": 0.0,
                    "gamma_rate": 0.0,
                    "gamma_dose": 0.0,
                    "beta_rate": 0.0,
                    "beta_dose": 0.0,
                },
            )
            entry["boi_atoms"] += _safe_float(row[3])
            entry["eoid_atoms"] += _safe_float(row[4])
            entry["boi_mass"] += _safe_float(row[5])
            entry["eoid_mass"] += _safe_float(row[6])
            entry["activity_ci"] += _safe_float(row[7])
            entry["gamma_rate"] += _safe_float(row[8])
            entry["gamma_dose"] += _safe_float(row[9])
            entry["beta_rate"] += _safe_float(row[10])
            entry["beta_dose"] += _safe_float(row[11])

    all_isotope_entries: List[Dict[str, Any]] = []
    activity_by_zaid: Dict[int, float] = {}
    for zaid, entry in combined_by_zaid.items():
        activity_ci = entry["activity_ci"]
        gamma_rate = entry["gamma_rate"]
        gamma_dose = entry["gamma_dose"]
        beta_rate = entry["beta_rate"]
        beta_dose = entry["beta_dose"]
        total_activity_ci += activity_ci
        total_gamma_rate += gamma_rate
        total_gamma_dose += gamma_dose
        total_beta_rate += beta_rate
        total_beta_dose += beta_dose
        activity_by_zaid[zaid] = activity_ci
        all_isotope_entries.append(
            {
                "activity_ci": activity_ci,
                "zaid": zaid,
                "reference_row": [
                    zaid,
                    entry["isotope"],
                    entry["boi_atoms"],
                    entry["eoid_atoms"],
                ],
                "main_row": [
                    entry["element"],
                    zaid,
                    entry["isotope"],
                    entry["boi_atoms"],
                    entry["eoid_atoms"],
                    entry["boi_mass"],
                    entry["eoid_mass"],
                    activity_ci,
                    gamma_rate,
                    gamma_dose,
                    beta_rate,
                    beta_dose,
                ],
            }
        )

    all_isotope_entries.sort(key=lambda entry: (-entry["activity_ci"], entry["zaid"]))
    filtered_entries = [
        entry
        for entry in all_isotope_entries
        if _safe_float(entry["main_row"][4]) > isotopes_min_atoms
        and _passes_activity_filter(
            _safe_float(entry["main_row"][7]),
            isotopes_min_activity_ci,
        )
    ]

    isotope_rows = [entry["main_row"] for entry in filtered_entries]
    isotope_reference_rows = [entry["reference_row"] for entry in filtered_entries]
    element_payload = _build_element_rows(
        [entry["main_row"] for entry in all_isotope_entries],
        metadata,
        elements_min_mass,
    )
    gamma_payload = _build_gamma_rows(
        metadata,
        activity_by_zaid,
        total_gamma_rate,
        gamma_min_rate,
    )

    return {
        "isotopes_summary": {
            "Total Activity (Ci)": total_activity_ci,
            "Total Activity (Bq)": total_activity_ci * CI_TO_BQ,
            "Gamma-Ray Emission Rate (Bq)": total_gamma_rate,
            "Gamma Dose Rate (R/hr) at 30 cm": total_gamma_dose,
            "Beta Emission Rate (Bq)": total_beta_rate,
            "Beta Dose Rate (R/hr) at Contact": total_beta_dose,
        },
        "gamma_summary": gamma_payload["summary"],
        "elements_filtered": element_payload["filtered_rows"],
        "elements_all": element_payload["all_rows"],
        "isotopes_all": [entry["main_row"] for entry in all_isotope_entries],
        "isotopes_main": isotope_rows,
        "isotopes_reference": isotope_reference_rows,
        "gamma_table": gamma_payload["rows"],
    }

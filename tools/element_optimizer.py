# Copyright 2026, Battelle Energy Alliance, LLC. All Rights Reserved.
from __future__ import annotations
"""Element-set optimization utilities for CRPYT coverage studies.

This module answers a different question than the main GUI: instead of
reporting results for a user-supplied composition, it estimates which elements
would make a good dosimetry set for a given irradiation spectrum.

Current optimization model:
- every candidate element is evaluated at the same trial mass
- only reactions whose estimated activity clears a user-defined minimum are
  eligible for selection
- at most one reaction is chosen per element
- a greedy search adds the element/reaction pair that improves energy coverage
  the most at each step

The activity split across competing production paths uses the same
spectrum-weighted cross-section proxy that drives the CRPYT heatmap.
"""

import math
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Sequence, Tuple

from openpyxl.utils import column_index_from_string

from extract_crpyt_irdff_data import (
    BUNDLE_NAME,
    DEFAULT_DECAY_TIME,
    DEFAULT_ENERGY_MAX,
    DEFAULT_ENERGY_MIN,
    DEFAULT_FLUENCE_RATE,
    DEFAULT_IRRADIATION_TIME,
    DEFAULT_REACTOR,
)
from json_engine import JsonEngineFactory
from results_builder import ReferenceMetadata, build_results, load_reference_metadata


COMBINED_REACTOR_UI_LABEL = "Combined Spectrum"
COMBINED_REACTOR_WORKBOOK_VALUE = "Combo"
IMPORT_SPECTRUM_UI_LABEL = "Import Spectrum"
ENERGY_SHEET = "Choose Irradiation Parameters"
PRODUCTION_SHEET = "Production"
OTHER_MT_SHEET = "Other MT"
ISOTOPICS_SHEET = "Isotopics"
ISOTOPE_PROPERTIES_SHEET = "Isotope Properties"
STABLE_ISOTOPE_PROPERTIES_SHEET = "Stable Isotope Properties"


@dataclass(frozen=True)
class SpectrumMixEntry:
    """One row of the combined-spectrum input table."""

    name: str
    fraction: float


@dataclass(frozen=True)
class OptimizationInputs:
    """User-controlled settings for the element-set optimization pass."""

    reactor: str = DEFAULT_REACTOR
    fluence_rate: float = DEFAULT_FLUENCE_RATE
    irradiation_time_sec: float = DEFAULT_IRRADIATION_TIME
    decay_time_sec: float = DEFAULT_DECAY_TIME
    energy_min_mev: float = DEFAULT_ENERGY_MIN
    energy_max_mev: float = DEFAULT_ENERGY_MAX
    trial_mass_g: float = 1.0
    min_reaction_activity_ci: float = 1.0e-11
    min_recommended_mass_g: float = 0.0
    max_recommended_mass_g: float | None = None
    candidate_cover_materials: Tuple[str, ...] = ()
    cover_thickness: float = 0.0
    max_elements: int = 8
    coverage_mode: str = "log"
    major_paths_only: bool = True
    path_activity_share_threshold: float = 0.01
    included_elements: Tuple[str, ...] = ()
    excluded_elements: Tuple[str, ...] = ()
    combined_spectra: Tuple[SpectrumMixEntry, ...] = ()
    imported_spectrum: Tuple[float, ...] = ()


@dataclass(frozen=True)
class ReactionCandidate:
    """One eligible production path that can represent an element in the search."""

    element: str
    atomic_number: int
    cover_material: str | None
    label: str
    parent_zaid: int
    daughter_zaid: int
    estimated_activity_ci: float
    recommended_mass_g: float
    product_activity_ci: float
    activity_share: float
    weighted_xs: float
    total_reaction_rate: float
    centroid_energy_mev: float
    normalized_profile: Tuple[float, ...]


@dataclass(frozen=True)
class ElementPoolStatus:
    """Status shown for each candidate element after optimization filtering."""

    element: str
    status: str


@dataclass(frozen=True)
class ElementSelection:
    """One greedy optimizer pick."""

    rank: int
    element: str
    cover_material: str | None
    reaction_label: str
    recommended_mass_g: float
    estimated_activity_ci: float
    total_reaction_rate: float
    centroid_energy_mev: float
    incremental_coverage: float
    cumulative_coverage: float


@dataclass(frozen=True)
class OptimizationResult:
    """Top-level optimization report returned by the backend."""

    inputs: OptimizationInputs
    selected: Tuple[ElementSelection, ...]
    candidate_pool: Tuple[ElementPoolStatus, ...]
    total_candidate_elements: int
    eligible_elements: Tuple[str, ...]
    ineligible_elements: Tuple[str, ...]
    total_candidate_reactions: int
    qualifying_reactions: int
    final_coverage: float
    coverage_mode: str


@dataclass(frozen=True)
class _ReactionLibrary:
    """Internal representation of the one-run reaction library."""

    reactions: Tuple[ReactionCandidate, ...]
    total_candidate_elements: int
    energies_mev: Tuple[float, ...]


def _default_bundle_dir() -> Path:
    """Resolve the CRPYT JSON bundle next to this source file."""

    return Path(__file__).resolve().parent / "data" / BUNDLE_NAME


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


def _normalize_reactor_name(value: str) -> str:
    text = value.strip()
    if text == COMBINED_REACTOR_UI_LABEL:
        return COMBINED_REACTOR_WORKBOOK_VALUE
    return text


def _normalize_element_names(values: Sequence[str]) -> set[str]:
    return {value.strip().upper() for value in values if value.strip()}


def _build_isotope_name_lookup(engine: Any) -> Dict[int, str]:
    """Merge the workbook naming sources used throughout the GUI."""

    lookup: Dict[int, str] = {}
    isotope_ws = engine.wb_formula[ISOTOPE_PROPERTIES_SHEET]
    for row in range(2, isotope_ws.max_row + 1):
        zaid = _safe_int(isotope_ws._get_value(row, 1))
        name = isotope_ws._get_value(row, 2)
        if zaid is not None and name not in (None, "") and zaid not in lookup:
            lookup[zaid] = str(name)

    stable_ws = engine.wb_formula[STABLE_ISOTOPE_PROPERTIES_SHEET]
    for row in range(3, stable_ws.max_row + 1):
        zaid = _safe_int(stable_ws._get_value(row, 1))
        name = stable_ws._get_value(row, 3)
        if zaid is not None and name not in (None, "") and zaid not in lookup:
            lookup[zaid] = str(name)
    return lookup


def _build_other_mt_lookup(engine: Any) -> Dict[tuple[int, str], str]:
    """Map each parent/MT pair to its Other MT cross-section column."""

    lookup: Dict[tuple[int, str], str] = {}
    ws = engine.wb_formula[ISOTOPICS_SHEET]
    for row in range(2, ws.max_row + 1):
        zaid = _safe_int(ws._get_value(row, 9))
        mt_meta = ws._get_value(row, 10)
        if zaid is None or mt_meta in (None, ""):
            continue
        try:
            column_name = str(engine._get_cell(ISOTOPICS_SHEET, row, 12))
        except Exception:
            continue
        if not column_name or column_name == "CLB":
            continue
        lookup[(zaid, str(mt_meta))] = column_name
    return lookup


def _format_mt_meta(mt_value: Any, daughter_zaid: int | None) -> str:
    text = str(int(float(mt_value))) if mt_value not in (None, "") else ""
    if daughter_zaid is not None and daughter_zaid % 1000 > 400:
        return f"{text}m"
    return text


def _build_linear_energy_edges(energies: Sequence[float]) -> List[float]:
    """Convert point-sampled energies into bin edges for integration weights."""

    if not energies:
        return []
    if len(energies) == 1:
        energy = energies[0]
        return [max(energy * 0.5, 1.0e-30), energy * 1.5]

    edges: List[float] = []
    first_step = energies[1] - energies[0]
    edges.append(max(energies[0] - first_step / 2.0, 1.0e-30))
    for left, right in zip(energies[:-1], energies[1:]):
        edges.append((left + right) / 2.0)
    last_step = energies[-1] - energies[-2]
    edges.append(energies[-1] + last_step / 2.0)
    return edges


def _split_cell_ref(cell_ref: str) -> tuple[int, int]:
    """Convert an A1-style cell reference into row and column numbers."""

    column_text = "".join(ch for ch in cell_ref if ch.isalpha())
    row_text = "".join(ch for ch in cell_ref if ch.isdigit())
    if not column_text or not row_text:
        raise ValueError(f"Invalid workbook cell reference: {cell_ref}")
    return int(row_text), column_index_from_string(column_text)


def _cover_attenuation_for_spectrum_row(
    engine: Any,
    gui_config: Mapping[str, Any],
    spectrum_row: int,
) -> float:
    """Calculate the cover attenuation factor used by built-in spectra."""

    attenuation_exponent = 0.0
    cover_xs_row = spectrum_row + 4
    for column_offset, entry in enumerate(gui_config["cover_materials"], start=2):
        cell_row, cell_column = _split_cell_ref(str(entry["cell"]))
        thickness = _safe_float(engine._get_cell(ENERGY_SHEET, cell_row, cell_column))
        if thickness <= 0.0:
            continue
        coefficient = _safe_float(engine._get_cell("Cover XS", cover_xs_row, column_offset))
        if coefficient > 0.0:
            attenuation_exponent += thickness * coefficient / 10.0

    if attenuation_exponent <= 0.0:
        return 1.0
    if attenuation_exponent > 745.0:
        return 0.0
    return math.exp(-attenuation_exponent)


def _build_imported_spectrum_overrides(
    engine: Any,
    gui_config: Mapping[str, Any],
    imported_spectrum: Sequence[float],
) -> Dict[str, float]:
    """Build J-column overrides for imported optimizer spectra."""

    if not imported_spectrum:
        return {}

    target_rows: List[int] = []
    for row in range(3, 1003):
        energy = _safe_float(engine._get_cell(ENERGY_SHEET, row, 9), default=-1.0)
        if energy > 0.0:
            target_rows.append(row)
    if len(imported_spectrum) != len(target_rows):
        raise ValueError("Imported spectrum does not match the CRPYT reactor spectrum grid.")

    overrides: Dict[str, float] = {}
    for row, value in zip(target_rows, imported_spectrum):
        attenuation = _cover_attenuation_for_spectrum_row(engine, gui_config, row)
        overrides[f"{ENERGY_SHEET}!J{row}"] = max(float(value) * attenuation, 0.0)
    return overrides

def _build_bounded_spectrum_overrides(
    engine: Any,
    min_energy: float,
    max_energy: float,
) -> Dict[str, float]:
    """Freeze the irradiation spectrum to the requested energy window."""

    overrides: Dict[str, float] = {}
    for row in range(3, 1003):
        energy = _safe_float(engine._get_cell(ENERGY_SHEET, row, 9), default=-1.0)
        if energy <= 0.0:
            continue
        spectrum_value = _safe_float(engine._get_cell(ENERGY_SHEET, row, 10))
        bounded_value = 0.0
        if min_energy <= energy <= max_energy:
            bounded_value = max(spectrum_value, 0.0)
        overrides[f"{ENERGY_SHEET}!J{row}"] = bounded_value
    return overrides


def _extract_bounded_spectrum(engine: Any) -> tuple[List[float], List[float]]:
    """Read the bounded spectrum directly from the engine's live state."""

    energies: List[float] = []
    spectrum: List[float] = []
    for row in range(3, 1003):
        energy = _safe_float(engine._get_cell(ENERGY_SHEET, row, 9), default=-1.0)
        value = _safe_float(engine._get_cell(ENERGY_SHEET, row, 10), default=-1.0)
        if energy <= 0.0 or value < 0.0:
            continue
        energies.append(energy)
        spectrum.append(value)
    return energies, spectrum


def _validate_inputs(inputs: OptimizationInputs) -> None:
    """Catch invalid optimization settings before the engine work starts."""

    if inputs.fluence_rate <= 0.0:
        raise ValueError("Fluence rate must be greater than 0.")
    if inputs.irradiation_time_sec <= 0.0:
        raise ValueError("Irradiation time must be greater than 0.")
    if inputs.decay_time_sec < 0.0:
        raise ValueError("Decay time must be greater than or equal to 0.")
    if inputs.energy_min_mev < 0.0 or inputs.energy_max_mev <= 0.0:
        raise ValueError("Energy bounds must be nonnegative and increasing.")
    if inputs.energy_min_mev >= inputs.energy_max_mev:
        raise ValueError("Minimum energy must be less than maximum energy.")
    if inputs.trial_mass_g <= 0.0:
        raise ValueError("Trial mass must be greater than 0.")
    if inputs.min_reaction_activity_ci < 0.0:
        raise ValueError("Minimum reaction activity must be greater than or equal to 0.")
    if inputs.min_recommended_mass_g < 0.0:
        raise ValueError("Minimum recommended mass must be greater than or equal to 0.")
    if inputs.max_recommended_mass_g is not None and inputs.max_recommended_mass_g <= 0.0:
        raise ValueError("Maximum recommended mass must be greater than 0 when provided.")
    if (
        inputs.max_recommended_mass_g is not None
        and inputs.min_recommended_mass_g > inputs.max_recommended_mass_g
    ):
        raise ValueError("Minimum recommended mass must be less than or equal to maximum recommended mass.")
    if inputs.cover_thickness < 0.0:
        raise ValueError("Cover thickness must be greater than or equal to 0.")
    if inputs.max_elements <= 0:
        raise ValueError("Max elements must be greater than 0.")
    if inputs.coverage_mode not in {"log", "linear"}:
        raise ValueError("Coverage mode must be 'log' or 'linear'.")
    if not (0.0 <= inputs.path_activity_share_threshold <= 1.0):
        raise ValueError("Path activity share threshold must be between 0 and 1.")
    if inputs.reactor == IMPORT_SPECTRUM_UI_LABEL and not inputs.imported_spectrum:
        raise ValueError("Imported spectrum values are required when Import Spectrum is selected.")


def _select_candidate_rows(
    metadata: ReferenceMetadata,
    inputs: OptimizationInputs,
) -> List[Dict[str, Any]]:
    """Filter the GUI element list down to the optimization candidate set."""

    all_rows = [dict(entry) for entry in metadata.gui_config["composition_rows"]]
    include_names = _normalize_element_names(inputs.included_elements)
    exclude_names = _normalize_element_names(inputs.excluded_elements)
    available_names = {str(entry["element"]).upper() for entry in all_rows}

    unknown_includes = sorted(include_names - available_names)
    if unknown_includes:
        raise ValueError(f"Unknown included element(s): {', '.join(unknown_includes)}")

    unknown_excludes = sorted(exclude_names - available_names)
    if unknown_excludes:
        raise ValueError(f"Unknown excluded element(s): {', '.join(unknown_excludes)}")

    selected: List[Dict[str, Any]] = []
    for entry in all_rows:
        name = str(entry["element"]).upper()
        if include_names and name not in include_names:
            continue
        if name in exclude_names:
            continue
        selected.append(entry)

    if not selected:
        raise ValueError("No candidate elements remain after include/exclude filters.")
    return selected


def _select_cover_materials(
    metadata: ReferenceMetadata,
    inputs: OptimizationInputs,
) -> tuple[str, ...]:
    """Validate and normalize the optional optimizer cover scenarios."""

    available = {
        str(entry["label"]).strip().upper(): str(entry["label"])
        for entry in metadata.gui_config["cover_materials"]
    }
    selected: List[str] = []
    seen = set()
    for raw in inputs.candidate_cover_materials:
        key = raw.strip().upper()
        if not key:
            continue
        if key not in available:
            raise ValueError(f"Unknown cover material: {raw}")
        if key in seen:
            continue
        seen.add(key)
        selected.append(available[key])
    return tuple(selected)


def _apply_inputs_to_engine(
    engine: Any,
    metadata: ReferenceMetadata,
    inputs: OptimizationInputs,
    candidate_rows: Sequence[Mapping[str, Any]],
    applied_cover_material: str | None,
) -> None:
    """Apply one optimization scenario to a fresh CRPYT engine instance."""

    gui_config = metadata.gui_config
    engine.set_cell(f"{ENERGY_SHEET}!B2", _normalize_reactor_name(inputs.reactor))
    engine.set_cell(f"{ENERGY_SHEET}!B3", float(inputs.fluence_rate))
    engine.set_cell(f"{ENERGY_SHEET}!B4", float(inputs.irradiation_time_sec))
    engine.set_cell(f"{ENERGY_SHEET}!B5", float(inputs.decay_time_sec))

    # Optimization cover scenarios are evaluated independently of the main
    # Inputs-tab cover settings: every run is either bare or uses exactly one
    # selected cover at the requested thickness.
    for entry in gui_config["cover_materials"]:
        engine.set_cell(
            f"{ENERGY_SHEET}!{entry['cell']}",
            float(inputs.cover_thickness)
            if applied_cover_material == str(entry["label"])
            else 0.0,
        )

    spectrum_entries = list(inputs.combined_spectra)
    for index, entry in enumerate(gui_config["spectrum_rows"]):
        row = entry["row"]
        if index < len(spectrum_entries):
            engine.set_cell(f"{ENERGY_SHEET}!B{row}", spectrum_entries[index].name)
            engine.set_cell(f"{ENERGY_SHEET}!C{row}", float(spectrum_entries[index].fraction))
        else:
            engine.set_cell(f"{ENERGY_SHEET}!B{row}", entry["default_name"])
            engine.set_cell(f"{ENERGY_SHEET}!C{row}", float(entry["default_fraction"]))

    for ref, value in _build_imported_spectrum_overrides(
        engine,
        gui_config,
        inputs.imported_spectrum,
    ).items():
        engine.set_cell(ref, value)

    candidate_rows_by_row = {int(entry["row"]): entry for entry in candidate_rows}
    for entry in gui_config["composition_rows"]:
        row = int(entry["row"])
        mass = float(inputs.trial_mass_g) if row in candidate_rows_by_row else 0.0
        engine.set_cell(f"{ENERGY_SHEET}!G{row}", mass)

    for ref, value in _build_bounded_spectrum_overrides(
        engine,
        inputs.energy_min_mev,
        inputs.energy_max_mev,
    ).items():
        engine.set_cell(ref, value)


def _build_reaction_library(
    engine: Any,
    metadata: ReferenceMetadata,
    candidate_rows: Sequence[Mapping[str, Any]],
    inputs: OptimizationInputs,
    applied_cover_material: str | None,
) -> _ReactionLibrary:
    """Generate the one-pass reaction library used by the greedy optimizer."""

    results = build_results(
        engine,
        metadata,
        elements_min_mass=0.0,
        isotopes_min_atoms=0.0,
        isotopes_min_activity_ci=0.0,
        gamma_min_rate=0.0,
    )
    energies, spectrum = _extract_bounded_spectrum(engine)
    isotope_names = _build_isotope_name_lookup(engine)
    other_lookup = _build_other_mt_lookup(engine)
    production_ws = engine.wb_formula[PRODUCTION_SHEET]
    other_mt_ws = engine.wb_formula[OTHER_MT_SHEET]

    atomic_to_element = {
        int(entry["atomic_number"]): str(entry["element"])
        for entry in candidate_rows
    }
    candidate_atomic_numbers = set(atomic_to_element)
    isotope_activity_by_zaid = {
        _safe_int(row[1]): _safe_float(row[7])
        for row in results["isotopes_main"]
        if _safe_int(row[1]) is not None
    }
    included_isotopes = set(isotope_activity_by_zaid)

    base_reactions: List[Dict[str, Any]] = []
    seen = set()
    for row in range(2, production_ws.max_row + 1):
        parent_zaid = _safe_int(production_ws._get_value(row, 2))
        mt_value = production_ws._get_value(row, 4)
        reaction = production_ws._get_value(row, 5)
        daughter_zaid = _safe_int(production_ws._get_value(row, 10))
        if parent_zaid is None or daughter_zaid is None:
            continue
        if mt_value in (None, "") or reaction in (None, ""):
            continue
        atomic_number = parent_zaid // 1000
        if atomic_number not in candidate_atomic_numbers:
            continue
        if daughter_zaid not in included_isotopes:
            continue

        mt_meta = _format_mt_meta(mt_value, daughter_zaid)
        column_name = other_lookup.get((parent_zaid, mt_meta))
        if not column_name:
            continue

        key = (parent_zaid, daughter_zaid, mt_meta, column_name)
        if key in seen:
            continue
        seen.add(key)

        column_index = column_index_from_string(column_name)
        cross_section_values: List[float] = []
        for mt_row in range(5, 1005):
            raw_value = _safe_float(other_mt_ws._get_value(mt_row, column_index))
            cross_section_values.append(max(raw_value, 0.0))
        if max(cross_section_values, default=0.0) <= 0.0:
            continue

        parent_name = isotope_names.get(parent_zaid, str(parent_zaid))
        daughter_name = isotope_names.get(daughter_zaid, str(daughter_zaid))
        label = f"{parent_name} {reaction}"
        if daughter_name:
            label = f"{label} -> {daughter_name}"

        base_reactions.append(
            {
                "element": atomic_to_element[atomic_number],
                "atomic_number": atomic_number,
                "label": label,
                "parent_zaid": parent_zaid,
                "daughter_zaid": daughter_zaid,
                "product_activity_ci": isotope_activity_by_zaid.get(daughter_zaid, 0.0),
                "weighted_xs": _safe_float(engine._get_cell(OTHER_MT_SHEET, 1, column_index)),
                "cross_section_values": cross_section_values,
            }
        )

    reactions: List[ReactionCandidate] = []
    grouped: Dict[int, List[Dict[str, Any]]] = {}
    for reaction in base_reactions:
        grouped.setdefault(reaction["daughter_zaid"], []).append(reaction)

    # Competing production paths that land on the same daughter isotope split
    # that isotope's activity in proportion to the spectrum-weighted cross
    # section, matching the plot-side approximation already used in CRPYT.
    for group in grouped.values():
        total_weighted_xs = sum(max(item["weighted_xs"], 0.0) for item in group)
        split_count = len(group)
        for item in group:
            if total_weighted_xs > 0.0:
                share = max(item["weighted_xs"], 0.0) / total_weighted_xs
            else:
                share = 1.0 / split_count if split_count else 0.0

            rate_values = [
                max(spectrum_value, 0.0) * max(cross_section_value, 0.0)
                for spectrum_value, cross_section_value in zip(
                    spectrum,
                    item["cross_section_values"],
                )
            ]
            total_reaction_rate = sum(rate_values)
            if total_reaction_rate <= 0.0:
                continue

            estimated_activity_ci = item["product_activity_ci"] * share
            if inputs.major_paths_only:
                minimum_activity = (
                    max(item["product_activity_ci"], 0.0)
                    * inputs.path_activity_share_threshold
                )
                if estimated_activity_ci + 1.0e-30 < minimum_activity:
                    continue
            if estimated_activity_ci < inputs.min_reaction_activity_ci:
                continue
            recommended_mass_g = _recommended_mass_from_activity(estimated_activity_ci, inputs)

            peak_value = max(rate_values, default=0.0)
            centroid = sum(
                energy * value for energy, value in zip(energies, rate_values)
            ) / total_reaction_rate
            reactions.append(
                ReactionCandidate(
                    element=item["element"],
                    atomic_number=int(item["atomic_number"]),
                    cover_material=applied_cover_material,
                    label=str(item["label"]),
                    parent_zaid=int(item["parent_zaid"]),
                    daughter_zaid=int(item["daughter_zaid"]),
                    estimated_activity_ci=float(estimated_activity_ci),
                    recommended_mass_g=float(recommended_mass_g),
                    product_activity_ci=float(item["product_activity_ci"]),
                    activity_share=float(share),
                    weighted_xs=float(item["weighted_xs"]),
                    total_reaction_rate=float(total_reaction_rate),
                    centroid_energy_mev=float(centroid),
                    normalized_profile=tuple(
                        (value / peak_value) if peak_value > 0.0 else 0.0
                        for value in rate_values
                    ),
                )
            )

    reactions.sort(
        key=lambda item: (
            item.element,
            -item.total_reaction_rate,
            -item.estimated_activity_ci,
            item.label,
        )
    )
    return _ReactionLibrary(
        reactions=tuple(reactions),
        total_candidate_elements=len(candidate_rows),
        energies_mev=tuple(energies),
    )


def _build_coverage_weights(
    energies_mev: Sequence[float],
    coverage_mode: str,
) -> List[float]:
    """Build per-bin weights for the coverage integration metric."""

    edges = _build_linear_energy_edges(energies_mev)
    weights: List[float] = []
    for left, right in zip(edges[:-1], edges[1:]):
        if coverage_mode == "linear":
            weights.append(max(right - left, 0.0))
            continue
        lower = max(left, 1.0e-30)
        upper = max(right, lower * (1.0 + 1.0e-12))
        weights.append(max(math.log10(upper) - math.log10(lower), 0.0))
    return weights


def _coverage_score(profile: Sequence[float], weights: Sequence[float]) -> float:
    total_weight = sum(weights)
    if total_weight <= 0.0:
        return 0.0
    return sum(
        max(0.0, min(value, 1.0)) * weight
        for value, weight in zip(profile, weights)
    ) / total_weight


def _recommended_mass_from_activity(
    estimated_activity_ci: float,
    inputs: OptimizationInputs,
) -> float:
    """Back-calculate the minimum mass needed to satisfy the activity target.

    The optimizer evaluates each candidate at a shared trial mass. For the
    selected path, activity scales linearly with that mass in the current
    model, so the minimum mass required to hit the requested activity target
    can be inferred directly from the selected reaction's activity at the trial
    mass.
    """

    if inputs.min_reaction_activity_ci <= 0.0:
        return float(inputs.trial_mass_g)
    if estimated_activity_ci <= 0.0:
        return float(inputs.trial_mass_g)
    return float(inputs.trial_mass_g) * (
        float(inputs.min_reaction_activity_ci) / float(estimated_activity_ci)
    )


def _mass_within_bounds(mass_g: float, inputs: OptimizationInputs) -> bool:
    """Return True when the inferred mass satisfies the user's bounds."""

    if mass_g < inputs.min_recommended_mass_g:
        return False
    if (
        inputs.max_recommended_mass_g is not None
        and mass_g > inputs.max_recommended_mass_g
    ):
        return False
    return True


def _cover_display_name(cover_material: str | None) -> str:
    """Render bare and single-cover scenarios consistently for reports."""

    return "Bare" if cover_material in (None, "") else str(cover_material)


def optimize_element_set(
    inputs: OptimizationInputs,
    bundle_dir: Path | None = None,
) -> OptimizationResult:
    """Recommend a dosimetry element set for the requested irradiation case."""

    _validate_inputs(inputs)
    bundle_path = (bundle_dir or _default_bundle_dir()).resolve()
    factory = JsonEngineFactory(bundle_path)
    metadata = load_reference_metadata(factory.metadata_path)
    candidate_rows = _select_candidate_rows(metadata, inputs)
    selected_covers = _select_cover_materials(metadata, inputs)
    scenario_cover_materials: List[str | None] = [None]
    if inputs.cover_thickness > 0.0:
        scenario_cover_materials.extend(selected_covers)

    all_reactions: List[ReactionCandidate] = []
    total_candidate_reactions = 0
    energies_mev: Tuple[float, ...] = ()
    for applied_cover_material in scenario_cover_materials:
        engine = factory.create_engine()
        _apply_inputs_to_engine(
            engine,
            metadata,
            inputs,
            candidate_rows,
            applied_cover_material,
        )
        reaction_library = _build_reaction_library(
            engine,
            metadata,
            candidate_rows,
            inputs,
            applied_cover_material,
        )
        if not energies_mev:
            energies_mev = reaction_library.energies_mev
        total_candidate_reactions += len(reaction_library.reactions)
        all_reactions.extend(reaction_library.reactions)

    all_reactions_by_element: Dict[str, List[ReactionCandidate]] = {}
    for reaction in all_reactions:
        all_reactions_by_element.setdefault(reaction.element, []).append(reaction)

    reactions_by_element: Dict[str, List[ReactionCandidate]] = {}
    for element, reactions in all_reactions_by_element.items():
        bounded_reactions = [
            reaction
            for reaction in reactions
            if _mass_within_bounds(reaction.recommended_mass_g, inputs)
        ]
        if bounded_reactions:
            bounded_reactions.sort(
                key=lambda reaction: (
                    -reaction.total_reaction_rate,
                    -reaction.estimated_activity_ci,
                    _cover_display_name(reaction.cover_material),
                    reaction.label,
                )
            )
            reactions_by_element[element] = bounded_reactions

    eligible_elements = tuple(sorted(reactions_by_element))
    ineligible_elements = tuple(
        sorted(
            str(entry["element"])
            for entry in candidate_rows
            if str(entry["element"]) not in reactions_by_element
        )
    )

    weights = _build_coverage_weights(energies_mev, inputs.coverage_mode)
    current_coverage = [0.0] * len(weights)
    selected: List[ElementSelection] = []
    used_elements: set[str] = set()

    # Greedy selection is a good fit here because we care about incremental
    # coverage of a shared energy axis and want a transparent, inspectable
    # ranking rather than a black-box optimum.
    for rank in range(1, min(inputs.max_elements, len(eligible_elements)) + 1):
        best_reaction: ReactionCandidate | None = None
        best_gain = 0.0
        best_coverage: List[float] | None = None

        for element, reactions in reactions_by_element.items():
            if element in used_elements:
                continue
            for reaction in reactions:
                candidate_coverage = [
                    max(current, reaction_value)
                    for current, reaction_value in zip(
                        current_coverage,
                        reaction.normalized_profile,
                    )
                ]
                gain = _coverage_score(candidate_coverage, weights) - _coverage_score(
                    current_coverage,
                    weights,
                )
                if best_reaction is None or (
                    gain,
                    reaction.total_reaction_rate,
                    reaction.estimated_activity_ci,
                    reaction.element,
                    reaction.label,
                ) > (
                    best_gain,
                    best_reaction.total_reaction_rate,
                    best_reaction.estimated_activity_ci,
                    best_reaction.element,
                    best_reaction.label,
                ):
                    best_reaction = reaction
                    best_gain = gain
                    best_coverage = candidate_coverage

        if best_reaction is None or best_coverage is None:
            break
        if best_gain <= 0.0 and selected:
            break

        current_coverage = best_coverage
        used_elements.add(best_reaction.element)
        selected.append(
            ElementSelection(
                rank=rank,
                element=best_reaction.element,
                cover_material=best_reaction.cover_material,
                reaction_label=best_reaction.label,
                recommended_mass_g=best_reaction.recommended_mass_g,
                estimated_activity_ci=best_reaction.estimated_activity_ci,
                total_reaction_rate=best_reaction.total_reaction_rate,
                centroid_energy_mev=best_reaction.centroid_energy_mev,
                incremental_coverage=best_gain,
                cumulative_coverage=_coverage_score(current_coverage, weights),
            )
        )

    selected_by_element = {selection.element: selection for selection in selected}
    candidate_pool: List[ElementPoolStatus] = []
    for entry in sorted(candidate_rows, key=lambda item: str(item["element"])):
        element = str(entry["element"])
        if element in selected_by_element:
            status = f"Selected: {_cover_display_name(selected_by_element[element].cover_material)}"
        elif element in reactions_by_element:
            best_reaction = reactions_by_element[element][0]
            status = f"Eligible: {_cover_display_name(best_reaction.cover_material)}"
        else:
            candidate_reactions = all_reactions_by_element.get(element, [])
            if not candidate_reactions:
                status = "No qualifying reaction"
            else:
                # When an element has qualifying reactions at the trial mass
                # but none survive the requested mass window, surface the
                # specific bound that excluded it in the candidate pool.
                below_min = any(
                    reaction.recommended_mass_g < inputs.min_recommended_mass_g
                    for reaction in candidate_reactions
                )
                above_max = any(
                    inputs.max_recommended_mass_g is not None
                    and reaction.recommended_mass_g > inputs.max_recommended_mass_g
                    for reaction in candidate_reactions
                )
                if below_min and not above_max:
                    status = "Omitted: below minimum mass"
                elif above_max and not below_min:
                    status = "Omitted: above maximum mass"
                else:
                    status = "Omitted: outside mass bounds"
        candidate_pool.append(ElementPoolStatus(element=element, status=status))

    return OptimizationResult(
        inputs=inputs,
        selected=tuple(selected),
        candidate_pool=tuple(candidate_pool),
        total_candidate_elements=len(candidate_rows),
        eligible_elements=eligible_elements,
        ineligible_elements=ineligible_elements,
        total_candidate_reactions=total_candidate_reactions,
        qualifying_reactions=sum(len(reactions) for reactions in reactions_by_element.values()),
        final_coverage=_coverage_score(current_coverage, weights),
        coverage_mode=inputs.coverage_mode,
    )


def optimization_result_to_dict(result: OptimizationResult) -> Dict[str, Any]:
    """Convert the nested dataclasses into a JSON-friendly structure."""

    return asdict(result)

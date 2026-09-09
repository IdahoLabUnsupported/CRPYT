# Copyright 2026, Battelle Energy Alliance, LLC. All Rights Reserved.
from __future__ import annotations
"""Command-line entry point for the CRPYT element-set optimizer."""

import argparse
import json
from pathlib import Path
from typing import List

from element_optimizer import (
    COMBINED_REACTOR_UI_LABEL,
    COMBINED_REACTOR_WORKBOOK_VALUE,
    OptimizationInputs,
    SpectrumMixEntry,
    optimization_result_to_dict,
    optimize_element_set,
)


def _parse_spectrum_entries(values: List[str]) -> tuple[SpectrumMixEntry, ...]:
    """Parse repeated NAME=FRACTION combined-spectrum arguments."""

    entries = []
    for raw in values:
        if "=" not in raw:
            raise ValueError(f"Spectrum entries must use NAME=FRACTION format: {raw}")
        name, fraction = raw.split("=", 1)
        entries.append(SpectrumMixEntry(name=name.strip(), fraction=float(fraction)))
    return tuple(entries)


def build_parser() -> argparse.ArgumentParser:
    """Create the CLI argument parser for optimization studies."""

    parser = argparse.ArgumentParser(
        description="Recommend a set of CRPYT elements that covers the current irradiation spectrum.",
    )
    parser.add_argument("--reactor", default="TRIGA")
    parser.add_argument("--fluence", type=float, default=1.0e12, help="Neutron fluence rate (n/cm^2/sec)")
    parser.add_argument("--irradiation", type=float, default=3600.0, help="Irradiation time (sec)")
    parser.add_argument("--decay", type=float, default=86400.0, help="Decay time (sec)")
    parser.add_argument("--energy-min", type=float, default=1.0e-9, help="Minimum energy (MeV)")
    parser.add_argument("--energy-max", type=float, default=20.0, help="Maximum energy (MeV)")
    parser.add_argument("--trial-mass", type=float, default=1.0, help="Trial mass per candidate element (g)")
    parser.add_argument("--min-reaction-activity", type=float, default=1.0e-11, help="Minimum qualifying reaction activity (Ci)")
    parser.add_argument("--min-mass", type=float, default=0.0, help="Minimum allowed recommended element mass (g)")
    parser.add_argument("--max-mass", type=float, default=None, help="Maximum allowed recommended element mass (g)")
    parser.add_argument("--max-elements", type=int, default=8, help="Maximum number of elements to recommend")
    parser.add_argument("--coverage-mode", choices=["log", "linear"], default="log")
    parser.add_argument("--include", nargs="*", default=[], help="Optional element symbol allow-list")
    parser.add_argument("--exclude", nargs="*", default=[], help="Optional element symbol block-list")
    parser.add_argument(
        "--cover-material",
        action="append",
        default=[],
        help="Cover material label to evaluate as a single-cover scenario; repeat for multiple covers",
    )
    parser.add_argument(
        "--cover-thickness",
        type=float,
        default=0.0,
        help="Thickness applied to each selected single-cover scenario",
    )
    parser.add_argument("--spectrum", action="append", default=[], help="Combined spectrum row in NAME=FRACTION format")
    parser.add_argument("--all-paths", action="store_true", help="Keep minor production paths instead of applying the 1%% cutoff")
    parser.add_argument("--path-share-threshold", type=float, default=0.01, help="Minimum fraction of product isotope activity for a path to count")
    parser.add_argument("--bundle-dir", type=Path, default=None, help="Override the CRPYT JSON bundle location")
    parser.add_argument("--json-out", type=Path, default=None, help="Optional path to write the full result as JSON")
    return parser


def main() -> None:
    """Parse CLI arguments, run the optimizer, and print a readable summary."""

    parser = build_parser()
    args = parser.parse_args()

    reactor = args.reactor
    if reactor == COMBINED_REACTOR_UI_LABEL:
        reactor = COMBINED_REACTOR_WORKBOOK_VALUE

    inputs = OptimizationInputs(
        reactor=reactor,
        fluence_rate=args.fluence,
        irradiation_time_sec=args.irradiation,
        decay_time_sec=args.decay,
        energy_min_mev=args.energy_min,
        energy_max_mev=args.energy_max,
        trial_mass_g=args.trial_mass,
        min_reaction_activity_ci=args.min_reaction_activity,
        min_recommended_mass_g=args.min_mass,
        max_recommended_mass_g=args.max_mass,
        candidate_cover_materials=tuple(args.cover_material),
        cover_thickness=args.cover_thickness,
        max_elements=args.max_elements,
        coverage_mode=args.coverage_mode,
        major_paths_only=not args.all_paths,
        path_activity_share_threshold=args.path_share_threshold,
        included_elements=tuple(args.include),
        excluded_elements=tuple(args.exclude),
        combined_spectra=_parse_spectrum_entries(args.spectrum),
    )

    result = optimize_element_set(inputs, bundle_dir=args.bundle_dir)
    payload = optimization_result_to_dict(result)

    print(f"Coverage mode: {result.coverage_mode}")
    print(f"Final coverage score: {result.final_coverage:.6f}")
    print(f"Candidate elements: {result.total_candidate_elements}")
    print(f"Eligible elements: {len(result.eligible_elements)}")
    print(f"Ineligible elements: {len(result.ineligible_elements)}")
    print(f"Qualifying reactions: {result.qualifying_reactions}")
    print("")
    print("Recommended set:")
    if not result.selected:
        print("  No elements met the requested activity threshold.")
    for selection in result.selected:
        cover_name = selection.cover_material or "Bare"
        print(
            f"  {selection.rank}. {selection.element:>2s} | "
            f"{cover_name} | "
            f"{selection.reaction_label} | "
            f"mass={selection.recommended_mass_g:.6e} g | "
            f"activity={selection.estimated_activity_ci:.6e} Ci | "
            f"gain={selection.incremental_coverage:.6f} | "
            f"coverage={selection.cumulative_coverage:.6f}"
        )

    if args.json_out is not None:
        args.json_out.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        print("")
        print(f"Wrote JSON report to {args.json_out}")


if __name__ == "__main__":
    main()

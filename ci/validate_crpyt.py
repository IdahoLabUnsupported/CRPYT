# Copyright 2026, Battelle Energy Alliance, LLC. All Rights Reserved.
from __future__ import annotations

import compileall
import datetime as dt
import json
import struct
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
TOOLS_DIR = REPO_ROOT / "tools"
ARTIFACT_DIR = REPO_ROOT / "ci_artifacts"
SUMMARY_PATH = ARTIFACT_DIR / "validation-summary.txt"


@dataclass
class DummyInputs:
    isotopes: list[tuple[str, float]]
    count_time_sec: float
    distance_cm: float
    include_background: bool
    shielding_cm: dict[str, float]


@dataclass
class DummyResults:
    total_counts: list[float]
    unshielded_dose_30cm: float
    unshielded_gammas_emitted_s: float
    unshielded_gammas_incident_s: float
    unshielded_dead_time_pct: float
    unshielded_real_time_sec: float
    unshielded_distance_20pct_cm: float
    shielded_dose_30cm: float | None
    shielded_gammas_emitted_s: float | None
    shielded_gammas_incident_s: float | None
    shielded_dead_time_pct: float | None
    shielded_real_time_sec: float | None
    shielded_distance_20pct_cm: float | None
    eff_a: float
    eff_b: float


def record(message: str) -> None:
    ARTIFACT_DIR.mkdir(exist_ok=True)
    with SUMMARY_PATH.open("a", encoding="utf-8") as handle:
        handle.write(f"{message}\n")


def require(path: Path, label: str) -> None:
    if not path.exists():
        raise AssertionError(f"Missing {label}: {path.relative_to(REPO_ROOT)}")


def validate_required_files() -> None:
    required = [
        (TOOLS_DIR / "CRPYT.py", "main CRPYT GUI"),
        (TOOLS_DIR / "GREASPY.py", "GREASPY spectrum dependency"),
        (TOOLS_DIR / "CRPYT_export.py", "spectrum exporter"),
        (TOOLS_DIR / "spectrum_import.py", "spectrum import helper"),
        (TOOLS_DIR / "spectrum_import_template.csv", "spectrum import template"),
        (TOOLS_DIR / "json_engine.py", "JSON engine"),
        (TOOLS_DIR / "results_builder.py", "results builder"),
        (TOOLS_DIR / "element_optimizer.py", "element optimizer"),
        (TOOLS_DIR / "extract_crpyt_irdff_data.py", "CRPYT data extractor"),
        (TOOLS_DIR / "extract_greaspy_crpyt_data.py", "GREASPY/CRPYT spectrum data extractor"),
        (TOOLS_DIR / "data" / "crpyt_irdff_rev1" / "manifest.json", "CRPYT data manifest"),
        (TOOLS_DIR / "data" / "greaspy_crpyt_rev1" / "manifest.json", "GREASPY/CRPYT data manifest"),
        (TOOLS_DIR / "__pycache__" / "CRPYT.cpython-312-INL430125.pyc", "legacy CRPYT engine cache"),
        (REPO_ROOT / "bin" / "CRPYT.exe", "Windows executable"),
    ]
    for path, label in required:
        require(path, label)

    exe_size = (REPO_ROOT / "bin" / "CRPYT.exe").stat().st_size
    if exe_size < 100_000_000:
        raise AssertionError(f"bin/CRPYT.exe is unexpectedly small ({exe_size} bytes)")

    workbook_files = sorted(
        path.relative_to(REPO_ROOT)
        for suffix in ("*.xlsx", "*.xlsm", "*.xls")
        for path in REPO_ROOT.rglob(suffix)
    )
    if workbook_files:
        joined = "\n".join(f"- {path}" for path in workbook_files)
        raise AssertionError(f"Workbook files should not be tracked in this repo:\n{joined}")


def resolve_manifest_path(bundle_dir: Path, rel_path: str) -> Path:
    return bundle_dir.joinpath(*str(rel_path).replace("\\", "/").split("/"))


def validate_json_bundle(bundle_name: str) -> None:
    bundle_dir = TOOLS_DIR / "data" / bundle_name
    manifest_path = bundle_dir / "manifest.json"
    metadata_path = bundle_dir / "metadata.json"
    require(manifest_path, f"{bundle_name} manifest")
    require(metadata_path, f"{bundle_name} metadata")

    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
    if not isinstance(manifest, dict):
        raise AssertionError(f"{bundle_name} manifest must be a JSON object")
    if not isinstance(metadata, dict):
        raise AssertionError(f"{bundle_name} metadata must be a JSON object")

    missing = []
    for key in ("formula_sheets", "cached_sheets"):
        for rel_path in manifest.get(key, []):
            if not resolve_manifest_path(bundle_dir, rel_path).exists():
                missing.append(rel_path)
    spectra_dir = manifest.get("spectra_dir")
    if spectra_dir:
        spectra_path = resolve_manifest_path(bundle_dir, str(spectra_dir))
        if not spectra_path.exists():
            missing.append(str(spectra_dir))
        elif not any(spectra_path.glob("*.json")):
            missing.append(f"{spectra_dir}/*.json")
    for rel_path in (manifest.get("metadata_file", "metadata.json"),):
        if not resolve_manifest_path(bundle_dir, rel_path).exists():
            missing.append(rel_path)
    if missing:
        joined = "\n".join(f"- {path}" for path in missing)
        raise AssertionError(f"{bundle_name} manifest references missing files:\n{joined}")


def validate_compile() -> None:
    ok = compileall.compile_dir(str(TOOLS_DIR), quiet=1, force=True)
    if not ok:
        raise AssertionError("Python source compilation failed")


def validate_workbook_export() -> bool:
    if sys.version_info[:2] != (3, 12):
        record("validation workbook export skipped because the legacy cache requires Python 3.12")
        return False

    sys.path.insert(0, str(TOOLS_DIR))
    from json_engine import JsonEngineFactory
    from openpyxl import load_workbook

    factory = JsonEngineFactory(TOOLS_DIR / "data" / "crpyt_irdff_rev1")
    overrides = {
        "Choose Irradiation Parameters!B2": "D-T",
        "Choose Irradiation Parameters!B3": 1.0e12,
        "Choose Irradiation Parameters!B4": 600.0,
        "Choose Irradiation Parameters!B5": 300.0,
        "Results - Elements!F4": 1.0e-15,
        "Results - Isotopes!B10": 0.0,
        "Results - Gamma!E2": 0.0,
    }

    with tempfile.TemporaryDirectory(dir=ARTIFACT_DIR) as tmp:
        out_path = Path(tmp) / "crpyt_validation_export_smoke.xlsx"
        factory.export_validation_workbook(overrides, out_path)
        if not out_path.exists() or out_path.stat().st_size == 0:
            raise AssertionError("Validation workbook export smoke file was not written")

        workbook = load_workbook(out_path, data_only=False, read_only=True)
        try:
            required_sheets = {
                "Choose Irradiation Parameters",
                "Results - Elements",
                "Results - Isotopes",
                "Results - Gamma",
            }
            missing = required_sheets.difference(workbook.sheetnames)
            if missing:
                raise AssertionError(f"Validation workbook export is missing sheets: {sorted(missing)}")

            inputs = workbook["Choose Irradiation Parameters"]
            if inputs["B2"].value != "D-T" or inputs["B3"].value != 1.0e12:
                raise AssertionError("Validation workbook input overrides were not applied")
            if not isinstance(inputs["C31"].value, str) or not inputs["C31"].value.startswith("="):
                raise AssertionError("Validation workbook formulas were not preserved")
        finally:
            workbook.close()
    return True


def validate_export_formats() -> None:
    sys.path.insert(0, str(TOOLS_DIR))
    from CRPYT_export import build_chn_bytes, build_spc_bytes, write_spectrum_file

    channels = 64
    inp = DummyInputs(
        isotopes=[("Co60", 1.0e-6), ("Fe59", 2.0e-7)],
        count_time_sec=300.0,
        distance_cm=5.0,
        include_background=True,
        shielding_cm={
            "Polyethylene": 0.0,
            "Aluminum": 0.0,
            "Stainless Steel 316": 0.0,
            "Tungsten": 0.0,
            "Pb": 0.0,
            "Depleted Uranium (DU)": 0.0,
        },
    )
    res = DummyResults(
        total_counts=[float(index + 1) for index in range(channels)],
        unshielded_dose_30cm=1.0e-6,
        unshielded_gammas_emitted_s=1.0e5,
        unshielded_gammas_incident_s=4.0e3,
        unshielded_dead_time_pct=2.5,
        unshielded_real_time_sec=307.69,
        unshielded_distance_20pct_cm=17.0,
        shielded_dose_30cm=None,
        shielded_gammas_emitted_s=None,
        shielded_gammas_incident_s=None,
        shielded_dead_time_pct=None,
        shielded_real_time_sec=None,
        shielded_distance_20pct_cm=None,
        eff_a=0.0,
        eff_b=1.0,
    )
    measured_at = dt.datetime(2026, 6, 9, 12, 0, 0)

    chn = build_chn_bytes(inp, res, measurement_datetime=measured_at)
    if struct.unpack_from("<h", chn, 0)[0] != -1:
        raise AssertionError("CHN header marker mismatch")
    if struct.unpack_from("<h", chn, 30)[0] != channels:
        raise AssertionError("CHN channel count mismatch")

    spc = build_spc_bytes(inp, res, measurement_datetime=measured_at)
    if len(spc) % 128 != 0:
        raise AssertionError("SPC output must be record aligned")
    if struct.unpack_from("<h", spc, 0)[0] != 1 or struct.unpack_from("<h", spc, 2)[0] != 5:
        raise AssertionError("SPC header mismatch")

    with tempfile.TemporaryDirectory(dir=ARTIFACT_DIR) as tmp:
        tmpdir = Path(tmp)
        for fmt in ("spe", "chn", "n42", "spc"):
            out_path = tmpdir / f"export_smoke.{fmt}"
            written = write_spectrum_file(out_path, inp, res, fmt, measurement_datetime=measured_at)
            if written != fmt:
                raise AssertionError(f"Expected {fmt} writer, got {written}")
            if not out_path.exists() or out_path.stat().st_size == 0:
                raise AssertionError(f"{fmt} export smoke file was not written")


def validate_spectrum_import() -> None:
    sys.path.insert(0, str(TOOLS_DIR))
    from spectrum_import import read_imported_spectrum_csv, rebase_spectrum_to_target_grid

    series = read_imported_spectrum_csv(TOOLS_DIR / "spectrum_import_template.csv")
    if len(series) != 1:
        raise AssertionError("Spectrum import template should contain exactly one spectrum")

    _name, points = series[0]
    target_energies = [energy for energy, _value in points]
    source_values = [value for _energy, value in points]
    rebinned_values, source_total, target_total = rebase_spectrum_to_target_grid(
        points,
        target_energies,
    )

    if len(rebinned_values) != len(source_values):
        raise AssertionError("Imported spectrum changed bin count during same-grid rebin")
    max_delta = max(abs(source - rebinned) for source, rebinned in zip(source_values, rebinned_values))
    if max_delta > 1.0e-12:
        raise AssertionError(f"Same-grid spectrum import changed bin weights by {max_delta:.3e}")
    if abs(sum(source_values) - target_total) > 1.0e-12:
        raise AssertionError("Same-grid spectrum import did not preserve the spectrum sum")
    if abs(source_total - target_total) > 1.0e-12:
        raise AssertionError("Spectrum import source/target conservation check failed")

def main() -> None:
    ARTIFACT_DIR.mkdir(exist_ok=True)
    validate_required_files()
    record("required file validation passed")
    validate_json_bundle("crpyt_irdff_rev1")
    record("CRPYT JSON bundle validation passed")
    validate_json_bundle("greaspy_crpyt_rev1")
    record("GREASPY/CRPYT spectrum bundle validation passed")
    validate_compile()
    record("Python source compilation passed")
    if validate_workbook_export():
        record("validation workbook export validation passed")
    validate_spectrum_import()
    record("spectrum importer validation passed")
    validate_export_formats()
    record("spectrum exporter validation passed")
    print("CRPYT CI validation passed.")


if __name__ == "__main__":
    main()

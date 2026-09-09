<!-- Copyright 2026, Battelle Energy Alliance, LLC. All Rights Reserved. -->

# CRPYT

Python tools and a Windows executable for CRPYT activation, gamma-line, element-optimization, and GREASPY detector-spectrum workflows.

CRPYT runs from packaged JSON data bundles. The original Excel workbooks are not required for application use and are intentionally excluded from this repository.

## Version Information

- CRPYT build version: 1.0.3
- CRPYT build date: 2026-08-04
- JSON bundle schema: version 1
- Included executable: `bin/CRPYT.exe`

## Authors

- Dr. Michael Reichenber, Idaho National Laboratory
- Dr. Tommy Holschuh, Idaho National Laboratory

## Repository Contents

- `bin/CRPYT.exe` - Windows executable.
- `tools/CRPYT.py` - main Tkinter desktop application.
- `tools/GREASPY.py` - GREASPY spectrum model used by the CRPYT Spectrum tab.
- `tools/CRPYT_export.py` - shared `.spe`, `.chn`, `.n42`, and `.spc` spectrum exporters.
- `tools/json_engine.py` - JSON-backed workbook emulator and validation-workbook exporter for the CRPYT calculation engine.
- `tools/results_builder.py` - result aggregation and display-table builder.
- `tools/element_optimizer.py` and `tools/optimize_elements.py` - element-set optimization logic and CLI.
- `tools/extract_crpyt_irdff_data.py` and `tools/extract_greaspy_crpyt_data.py` - maintenance scripts for rebuilding JSON bundles from local workbook sources.
- `tools/data/` - packaged CRPYT and GREASPY/CRPYT JSON data bundles.
- `tools/__pycache__/CRPYT.cpython-312-INL430125.pyc` - required cached legacy CRPYT formula engine.
- `ci/validate_crpyt.py` - CI validation harness.
- `ci/package_release.py` - release archive builder.

## Requirements

- Python 3.12 is recommended for source runs because the legacy formula engine cache is Python 3.12 bytecode.
- Windows is required for the tracked executable.
- Python packages listed in `requirements.txt`.

Validation-workbook export is generated directly from the bundled JSON data with `openpyxl`; no local Excel installation or workbook template is required.

## Usage

Run the tracked executable:

```powershell
.\bin\CRPYT.exe
```

Run from source:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python .\tools\CRPYT.py
```

## Import Reactor Spectra

Custom reactor spectra can be imported from the Reactor Type dropdown by choosing Import Spectrum. Use `tools/spectrum_import_template.csv` as a starting format. CRPYT accepts `.csv`, `.xlsx`, and `.xlsm` files where the first column is energy in MeV and each later column is a nonnegative spectrum. Column headers become the imported spectrum names.

A single import file can contain multiple spectra. Imported spectra are added as Reactor Type choices and can also be selected in Combined Spectrum rows.

During import, CRPYT conservatively rebases each source spectrum onto the internal 1000-bin log-space CRPYT energy grid and checks that the spectrum area is preserved before using it.

When project settings are saved, CRPYT records the imported spectrum source file path in `CRPYT.user_settings.json`. On later launches or project loads, the app reopens that saved location and restores the imported spectra, so you do not need to import them every time as long as the source file has not moved and remains readable.

The Spectrum tab can export detector spectra in `.spe`, `.chn`, `.n42`, and `.spc` formats after a calculation has populated the detector spectrum. The Export Workbook action writes an `.xlsx` reconstructed from the JSON bundle and current run inputs. The Generate PDF Report action prompts for a `.pdf` path and writes a shareable report with run inputs, summary outputs, isotope results grouped by generating element, and the reaction plot.

CRPYT stores named project states in `CRPYT.user_settings.json` when settings are saved. Use the Project selector in the bottom action bar to type or choose a project name, then Save State to capture the current run parameters, filters, cover settings, element masses, spectrum settings, and optimizer controls. Load State restores the selected project, and Delete State removes a saved non-default project. The active project is also saved when CRPYT closes so run parameters carry forward between sessions. The settings file is ignored by Git because it stores per-user GUI state; older single-state settings files load as the `Default` project.

## CI/CD

GitLab CI validates the repository without installing packages on the shell runner. It checks required files, JSON bundle manifests, workbook exclusion, source compilation, the required legacy engine cache, validation-workbook export, and spectrum exporter smoke tests.

Pipelines on `main` and tags also build `dist/crpyt-release.zip`. To avoid GitLab upload timeouts on the large Windows binary, CI publishes that archive as smaller `crpyt-release.zip.part-*` artifacts with a SHA256 file and reassembly instructions. The archive is the portable Windows binary distribution: it contains `bin/CRPYT.exe` plus the project README, changelog, and license. The JSON data bundles and required legacy cache are embedded in the executable by PyInstaller; clone the repository to run or inspect the Python source directly.

## Rebuild Data Bundles

The source workbooks stay outside the repository. To rebuild a bundle from a local workbook, pass the workbook path explicitly:

```powershell
python .\tools\extract_crpyt_irdff_data.py --workbook path\to\CRPYT_IRDFF_Rev1.xlsx --output .\tools\data\crpyt_irdff_rev1
python .\tools\extract_greaspy_crpyt_data.py --workbook path\to\GREASPY_CRPYT_Rev1.xlsx --output .\tools\data\greaspy_crpyt_rev1
```

## Rebuild Executable

Install runtime and build dependencies:

```powershell
python -m pip install -r requirements.txt
python -m pip install -r requirements-build.txt
```

Before building a merge-ready executable, bump the tracked build metadata. Use `--patch` for the usual merge increment, or pass `--version` for an explicit release:

```powershell
python .\ci\bump_build.py --patch
```

Build from the repository root:

```powershell
python -m PyInstaller --noconfirm --clean --distpath bin --workpath build\pyinstaller-crpyt tools\CRPYT.spec
```

## What Stays Out

The repo excludes original Excel workbooks, legacy validation folders, generated spectrum files, generated CSV files, PyInstaller scratch output, and ordinary Python cache files. The one tracked `.pyc` is a required runtime dependency for the cached legacy CRPYT formula engine.

## License

CRPYT is available under the MIT License. See `LICENSE` for details.

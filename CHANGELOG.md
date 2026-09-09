<!-- Copyright 2026, Battelle Energy Alliance, LLC. All Rights Reserved. -->

# Changelog

## 1.0.3 - 2026-08-04

- Report imported reactor spectra by their selected spectrum name on generated PDF input pages instead of the generic import token.

## 1.0.2 - 2026-07-28

- Add named project save states for retaining CRPYT run parameters, filters, cover settings, element masses, spectrum inputs, and optimizer settings between sessions.
- Keep older single-state `CRPYT.user_settings.json` files loadable as the `Default` project state.

## 1.0.0 - 2026-06-09

- Replace the scaffolded GREASPY-derived implementation with the actual CRPYT Python application from the OOS CRPYT source tree.
- Add packaged CRPYT and GREASPY/CRPYT JSON data bundles for workbook-free runtime calculations.
- Include the required cached legacy CRPYT formula engine bytecode used by the JSON-backed calculation path.
- Include the Windows executable and default `CRPYT.user_settings.json` distribution settings.
- Add shared spectrum exporters for `.spe`, `.chn`, `.n42`, and ORTEC real-format `.spc` files.
- Add a Spectrum-tab export action for detector-spectrum file export.
- Generate validation workbooks directly from the JSON bundle instead of requiring a local Excel template.
- Add CI validation for source compilation, bundle integrity, required runtime artifacts, workbook exclusion, and spectrum exporters.

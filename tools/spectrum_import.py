# Copyright 2026, Battelle Energy Alliance, LLC. All Rights Reserved.
"""Headless helpers for importing and rebinding reactor spectra."""

from __future__ import annotations

import csv
import math
from pathlib import Path
from typing import Any, Sequence


def read_imported_spectrum_file(path: Path) -> list[tuple[str, list[tuple[float, float]]]]:
    """Read CSV/XLSX spectra where column 1 is energy and later columns are spectra."""
    path = Path(path)
    if not path.exists():
        raise ValueError(f"Spectrum file does not exist: {path}")

    suffix = path.suffix.lower()
    if suffix == ".csv":
        return read_imported_spectrum_csv(path)
    if suffix in {".xlsx", ".xlsm"}:
        return read_imported_spectrum_workbook(path)
    raise ValueError("Imported spectrum files must be .csv, .xlsx, or .xlsm files.")


def read_imported_spectrum_csv(path: Path) -> list[tuple[str, list[tuple[float, float]]]]:
    """Read a CSV spectrum table, accepting one optional header row."""
    raw_rows: list[tuple[int, list[Any]]] = []
    try:
        with Path(path).open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.reader(handle)
            for line_number, row in enumerate(reader, start=1):
                raw_rows.append((line_number, list(row)))
    except OSError as exc:
        raise ValueError(f"Could not read spectrum CSV: {exc}") from exc
    return parse_imported_spectrum_rows(Path(path), raw_rows)


def read_imported_spectrum_workbook(path: Path) -> list[tuple[str, list[tuple[float, float]]]]:
    """Read the first sheet of an Excel spectrum workbook."""
    try:
        from openpyxl import load_workbook

        workbook = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        raise ValueError(f"Could not read spectrum workbook: {exc}") from exc

    try:
        worksheet = workbook.worksheets[0]
        raw_rows = [
            (row_number, list(row))
            for row_number, row in enumerate(worksheet.iter_rows(values_only=True), start=1)
        ]
    finally:
        workbook.close()
    return parse_imported_spectrum_rows(Path(path), raw_rows)


def cell_is_blank(value: Any) -> bool:
    """Return True for empty cells from CSV or workbook readers."""
    return value is None or (isinstance(value, str) and value.strip() == "")


def parse_imported_spectrum_rows(
    path: Path,
    raw_rows: Sequence[tuple[int, Sequence[Any]]],
) -> list[tuple[str, list[tuple[float, float]]]]:
    """Normalize raw table rows into one named point list per spectrum column."""
    rows: list[tuple[int, list[Any]]] = []
    for row_number, raw_row in raw_rows:
        cells = list(raw_row)
        while cells and cell_is_blank(cells[-1]):
            cells.pop()
        if not cells:
            continue
        first_cell = cells[0]
        if isinstance(first_cell, str) and first_cell.strip().startswith("#"):
            continue
        rows.append((row_number, cells))

    if not rows:
        raise ValueError("Imported spectrum file did not contain any data rows.")

    def can_parse_float(value: Any) -> bool:
        if cell_is_blank(value):
            return False
        try:
            parsed = float(str(value).strip())
        except Exception:
            return False
        return math.isfinite(parsed)

    first_cells = rows[0][1]
    first_row_is_data = (
        len(first_cells) >= 2
        and can_parse_float(first_cells[0])
        and all(can_parse_float(cell) for cell in first_cells[1:])
    )
    has_header = not first_row_is_data
    header_cells = rows[0][1] if has_header else []
    data_rows = rows[1:] if has_header else rows
    max_columns = max((len(cells) for _row_number, cells in rows), default=0)
    if max_columns < 2:
        raise ValueError("Imported spectrum file must contain energy plus at least one spectrum column.")

    spectrum_count = max_columns - 1
    raw_names: list[str] = []
    for column_index in range(1, max_columns):
        header_name = ""
        if has_header and column_index < len(header_cells) and not cell_is_blank(header_cells[column_index]):
            header_name = str(header_cells[column_index]).strip()
        if header_name:
            raw_names.append(header_name)
        elif spectrum_count == 1:
            raw_names.append(path.stem)
        else:
            raw_names.append(f"{path.stem} {column_index}")

    numeric_rows: list[tuple[float, list[float]]] = []
    for row_number, cells in data_rows:
        if len(cells) > max_columns:
            raise ValueError(f"Row {row_number} has more columns than the header row.")
        if not can_parse_float(cells[0]):
            raise ValueError(f"Row {row_number} has an invalid energy value.")
        energy = float(str(cells[0]).strip())
        if energy < 0.0:
            raise ValueError(f"Row {row_number} has an invalid energy value.")

        values: list[float] = []
        for column_index in range(1, max_columns):
            cell = cells[column_index] if column_index < len(cells) else None
            if not can_parse_float(cell):
                raise ValueError(
                    f"Row {row_number}, column {column_index + 1} must contain a numeric spectrum value."
                )
            value = float(str(cell).strip())
            if value < 0.0:
                raise ValueError(
                    f"Row {row_number}, column {column_index + 1} has an invalid spectrum value."
                )
            values.append(value)
        numeric_rows.append((energy, values))

    if len(numeric_rows) < 2:
        raise ValueError("Imported spectrum file must contain at least two numeric data rows.")

    numeric_rows.sort(key=lambda item: item[0])
    for previous, current in zip(numeric_rows[:-1], numeric_rows[1:]):
        if current[0] <= previous[0]:
            raise ValueError("Imported spectrum energies must be unique after sorting.")

    series: list[tuple[str, list[tuple[float, float]]]] = []
    for value_index, name in enumerate(raw_names):
        points = [(energy, values[value_index]) for energy, values in numeric_rows]
        series.append((name, points))
    return series


def build_linear_energy_edges(energies: Sequence[float]) -> list[float]:
    """Convert sampled energy centers into linear bin edges."""
    energy_values = list(energies)
    if not energy_values:
        return []
    if len(energy_values) == 1:
        energy = energy_values[0]
        return [max(energy * 0.5, 1.0e-30), energy * 1.5]

    edges: list[float] = []
    first_step = energy_values[1] - energy_values[0]
    edges.append(max(energy_values[0] - first_step / 2.0, 1.0e-30))
    for left, right in zip(energy_values[:-1], energy_values[1:]):
        edges.append((left + right) / 2.0)
    last_step = energy_values[-1] - energy_values[-2]
    edges.append(energy_values[-1] + last_step / 2.0)
    return edges


def rebase_spectrum_to_target_grid(
    points: Sequence[tuple[float, float]],
    target_energies: Sequence[float],
) -> tuple[list[float], float, float]:
    """Conservatively rebin imported spectrum weights onto CRPYT energy bins."""
    source_energies = [float(point[0]) for point in points]
    source_values = [float(point[1]) for point in points]
    source_edges = build_linear_energy_edges(source_energies)
    target_edges = build_linear_energy_edges(list(target_energies))

    if len(source_edges) != len(source_values) + 1:
        raise ValueError("Imported spectrum grid has invalid bin edges.")
    if len(target_edges) != len(target_energies) + 1:
        raise ValueError("CRPYT target spectrum grid has invalid bin edges.")
    for left_edge, right_edge in zip(source_edges[:-1], source_edges[1:]):
        if right_edge <= left_edge:
            raise ValueError("Imported spectrum bins must have positive width.")
    for left_edge, right_edge in zip(target_edges[:-1], target_edges[1:]):
        if right_edge <= left_edge:
            raise ValueError("CRPYT target spectrum bins must have positive width.")

    rebinned_values: list[float] = []
    target_min = target_edges[0]
    target_max = target_edges[-1]
    source_integral = 0.0
    source_index = 0
    for left_edge, right_edge in zip(target_edges[:-1], target_edges[1:]):
        while source_index < len(source_values) and source_edges[source_index + 1] <= left_edge:
            source_index += 1

        value = 0.0
        cursor = source_index
        while cursor < len(source_values) and source_edges[cursor] < right_edge:
            overlap_left = max(left_edge, source_edges[cursor])
            overlap_right = min(right_edge, source_edges[cursor + 1])
            if overlap_right > overlap_left:
                source_width = source_edges[cursor + 1] - source_edges[cursor]
                contribution = source_values[cursor] * (overlap_right - overlap_left) / source_width
                value += contribution
            if source_edges[cursor + 1] >= right_edge:
                break
            cursor += 1
        rebinned_values.append(max(value, 0.0))

    for left_edge, right_edge, source_value in zip(source_edges[:-1], source_edges[1:], source_values):
        overlap_left = max(left_edge, target_min)
        overlap_right = min(right_edge, target_max)
        if overlap_right > overlap_left:
            source_integral += source_value * (overlap_right - overlap_left) / (right_edge - left_edge)

    target_integral = sum(rebinned_values)
    return rebinned_values, source_integral, target_integral

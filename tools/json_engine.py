# Copyright 2026, Battelle Energy Alliance, LLC. All Rights Reserved.
from __future__ import annotations
"""JSON-backed workbook emulation for the standalone CRPYT runtime.

The original CRPYT calculation engine expects an openpyxl workbook and the
legacy formula interpreter that was written around it. This module recreates a
minimal workbook API on top of extracted JSON bundles so the same interpreter
can run without opening Excel files during normal calculations.
"""

import json
import math
from dataclasses import dataclass
from decimal import Decimal, localcontext
from functools import lru_cache
from pathlib import Path
from typing import Any, Dict, Iterable, List, Tuple

from openpyxl import Workbook
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.utils import column_index_from_string, get_column_letter

from _legacy_loader import load_legacy_engine_module


@dataclass(frozen=True)
class SheetBundle:
    """Immutable representation of one extracted worksheet."""
    name: str
    max_row: int
    max_col: int
    rows: Tuple[Tuple[Any, ...], ...]
    formula_cells: Tuple[Tuple[int, int], ...] = ()

    def instantiate(self) -> "JsonWorksheet":
        """Create a mutable worksheet instance from the stored base rows."""
        return JsonWorksheet(
            title=self.name,
            base_rows=self.rows,
            max_row=self.max_row,
            max_col=self.max_col,
        )


@dataclass(frozen=True)
class WorkbookBundle:
    """Container for all extracted sheets plus the validation template path."""
    bundle_dir: Path
    template_workbook: Path
    formula_sheets: Dict[str, SheetBundle]
    cached_sheets: Dict[str, SheetBundle]
    metadata_path: Path

    def instantiate_formula_workbook(self) -> "JsonWorkbook":
        """Build a fresh workbook containing the formula sheets."""
        sheets = [self.formula_sheets[name].instantiate() for name in self.formula_sheets]
        return JsonWorkbook(sheets)

    def instantiate_cached_workbook(self) -> "JsonWorkbook":
        """Build a fresh workbook containing cached runtime-only sheets."""
        sheets = [self.cached_sheets[name].instantiate() for name in self.cached_sheets]
        return JsonWorkbook(sheets)


class JsonCell:
    """Tiny cell wrapper used by the legacy engine's openpyxl-style accessors."""
    __slots__ = ("_worksheet", "row", "column")

    def __init__(self, worksheet: "JsonWorksheet", row: int, column: int) -> None:
        self._worksheet = worksheet
        self.row = row
        self.column = column

    @property
    def coordinate(self) -> str:
        return f"{get_column_letter(self.column)}{self.row}"

    @property
    def value(self) -> Any:
        return self._worksheet._get_value(self.row, self.column)

    @value.setter
    def value(self, new_value: Any) -> None:
        self._worksheet._set_value(self.row, self.column, new_value)


class JsonWorksheet:
    """Worksheet facade backed by immutable rows plus in-memory overrides."""
    __slots__ = ("title", "_base_rows", "_row_overrides", "max_row", "max_column")

    def __init__(
        self,
        title: str,
        base_rows: Tuple[Tuple[Any, ...], ...],
        max_row: int,
        max_col: int,
    ) -> None:
        self.title = title
        self._base_rows = base_rows
        self._row_overrides: Dict[int, List[Any]] = {}
        self.max_row = max_row
        self.max_column = max_col

    def _get_row(self, row: int) -> Tuple[Any, ...] | List[Any]:
        """Return the overridden row if present, otherwise the serialized row."""
        override = self._row_overrides.get(row)
        if override is not None:
            return override
        if 1 <= row <= len(self._base_rows):
            return self._base_rows[row - 1]
        return ()

    def _get_value(self, row: int, column: int) -> Any:
        if row < 1 or column < 1:
            return None
        row_data = self._get_row(row)
        if column <= len(row_data):
            return row_data[column - 1]
        return None

    def _set_value(self, row: int, column: int, value: Any) -> None:
        """Apply a local override without mutating the extracted JSON payload."""
        if row < 1 or column < 1:
            raise ValueError("Worksheet rows and columns are 1-based.")
        row_data = self._row_overrides.get(row)
        if row_data is None:
            base = self._base_rows[row - 1] if 1 <= row <= len(self._base_rows) else ()
            row_data = list(base)
            self._row_overrides[row] = row_data
        if column > len(row_data):
            row_data.extend([None] * (column - len(row_data)))
        row_data[column - 1] = value
        if row > self.max_row:
            self.max_row = row
        if column > self.max_column:
            self.max_column = column

    def cell(self, row: int, column: int) -> JsonCell:
        return JsonCell(self, row, column)

    def iter_rows(
        self,
        min_row: int = 1,
        max_row: int | None = None,
        min_col: int = 1,
        max_col: int | None = None,
    ) -> Iterable[Tuple[JsonCell, ...]]:
        upper_row = self.max_row if max_row is None else max_row
        upper_col = self.max_column if max_col is None else max_col
        for row in range(min_row, upper_row + 1):
            yield tuple(self.cell(row, col) for col in range(min_col, upper_col + 1))


class JsonWorkbook:
    """Minimal workbook wrapper exposing sheet lookup by name."""
    __slots__ = ("_sheets_by_name", "worksheets")

    def __init__(self, sheets: Iterable[JsonWorksheet]) -> None:
        worksheet_list = list(sheets)
        self.worksheets = worksheet_list
        self._sheets_by_name = {sheet.title: sheet for sheet in worksheet_list}

    def __getitem__(self, sheet_name: str) -> JsonWorksheet:
        return self._sheets_by_name[sheet_name]


def _parse_workbook_ref(ref: str) -> tuple[str, int, int]:
    """Split a workbook-style `Sheet!A1` reference into sheet, row, and column."""
    if "!" not in ref:
        raise ValueError(f"Workbook reference must include a sheet name: {ref}")
    sheet_name, coordinate = ref.split("!", 1)
    column_text, row = coordinate_from_string(coordinate.replace("$", ""))
    return sheet_name, int(row), column_index_from_string(column_text)


def _populate_openpyxl_sheet(ws, sheet_bundle: SheetBundle) -> None:
    """Copy one JSON sheet payload into an openpyxl worksheet."""
    for row_index, row_values in enumerate(sheet_bundle.rows, start=1):
        for column_index, value in enumerate(row_values, start=1):
            if value is not None:
                ws.cell(row=row_index, column=column_index, value=value)


def build_validation_workbook(bundle: "WorkbookBundle", overrides: Dict[str, Any]) -> Workbook:
    """Reconstruct a validation workbook from JSON sheet payloads."""
    workbook = Workbook()
    workbook.properties.title = "CRPYT Validation Workbook"
    workbook.properties.creator = "CRPYT"
    workbook.calculation.calcMode = "auto"
    workbook.calculation.calcOnSave = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.fullCalcOnLoad = True

    default_sheet = workbook.active
    for index, sheet_bundle in enumerate(bundle.formula_sheets.values()):
        worksheet = default_sheet if index == 0 else workbook.create_sheet()
        worksheet.title = sheet_bundle.name
        _populate_openpyxl_sheet(worksheet, sheet_bundle)

    for ref, value in overrides.items():
        sheet_name, row, column = _parse_workbook_ref(ref)
        if sheet_name not in workbook.sheetnames:
            raise KeyError(f"Unknown worksheet in validation export override: {sheet_name}")
        workbook[sheet_name].cell(row=row, column=column, value=value)

    return workbook


def _read_sheet_bundle(sheet_path: Path) -> SheetBundle:
    """Deserialize one extracted sheet file into an immutable bundle object."""
    with sheet_path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)
    rows = tuple(tuple(row) for row in payload["rows"])
    formula_cells = tuple(tuple(pair) for pair in payload.get("formula_cells", []))
    return SheetBundle(
        name=payload["name"],
        max_row=payload["max_row"],
        max_col=payload["max_col"],
        rows=rows,
        formula_cells=formula_cells,
    )


def _manifest_relative_path(bundle_path: Path, rel_path: str) -> Path:
    """Resolve manifest paths written on either Windows or POSIX systems."""
    return bundle_path.joinpath(*str(rel_path).replace("\\", "/").split("/"))


@lru_cache(maxsize=4)
def load_bundle(bundle_dir: str) -> WorkbookBundle:
    """Load the workbook manifest and all referenced sheet payloads."""
    bundle_path = Path(bundle_dir).resolve()
    manifest_path = bundle_path / "manifest.json"
    with manifest_path.open("r", encoding="utf-8") as handle:
        manifest = json.load(handle)

    formula_sheets: Dict[str, SheetBundle] = {}
    for rel_path in manifest["formula_sheets"]:
        sheet_bundle = _read_sheet_bundle(_manifest_relative_path(bundle_path, rel_path))
        formula_sheets[sheet_bundle.name] = sheet_bundle

    cached_sheets: Dict[str, SheetBundle] = {}
    for rel_path in manifest.get("cached_sheets", []):
        sheet_bundle = _read_sheet_bundle(_manifest_relative_path(bundle_path, rel_path))
        cached_sheets[sheet_bundle.name] = sheet_bundle

    template_workbook = _manifest_relative_path(bundle_path.parent, manifest["template_workbook"])
    metadata_path = _manifest_relative_path(bundle_path, manifest["metadata_file"])
    return WorkbookBundle(
        bundle_dir=bundle_path,
        template_workbook=template_workbook,
        formula_sheets=formula_sheets,
        cached_sheets=cached_sheets,
        metadata_path=metadata_path,
    )


LEGACY = load_legacy_engine_module()

PR_COEFF_IRR_SHEET = "PR Coeff Irr"
PR_COEFF_DECAY_SHEET = "PR Coeff Decay"
PR_COEFF_STABLE_VALUE_COLUMNS = {
    PR_COEFF_IRR_SHEET: {
        column_index_from_string("AT"): {
            "node_columns": ("D", "E", "F"),
            "factor_columns": ("W", "B", "X"),
        },
        column_index_from_string("AV"): {
            "node_columns": ("D", "E", "F", "G"),
            "factor_columns": ("W", "B", "X", "Y"),
        },
        column_index_from_string("AX"): {
            "node_columns": ("D", "E", "F", "G", "H"),
            "factor_columns": ("W", "B", "X", "Y", "Z"),
        },
        column_index_from_string("AZ"): {
            "node_columns": ("D", "E", "F", "G", "H", "I"),
            "factor_columns": ("W", "B", "X", "Y", "Z", "AA"),
        },
    },
    PR_COEFF_DECAY_SHEET: {
        column_index_from_string("AS"): {
            "node_columns": ("D", "E"),
            "factor_columns": ("W", "B"),
        },
        column_index_from_string("AU"): {
            "node_columns": ("D", "E", "F"),
            "factor_columns": ("W", "B", "X"),
        },
        column_index_from_string("AW"): {
            "node_columns": ("D", "E", "F", "G"),
            "factor_columns": ("W", "B", "X", "Y"),
        },
        column_index_from_string("AY"): {
            "node_columns": ("D", "E", "F", "G", "H"),
            "factor_columns": ("W", "B", "X", "Y", "Z"),
        },
        column_index_from_string("BA"): {
            "node_columns": ("D", "E", "F", "G", "H", "I"),
            "factor_columns": ("W", "B", "X", "Y", "Z", "AA"),
        },
    },
}


class JsonBackedEngine(LEGACY.CryptWorkbookEngine):
    """Legacy formula engine wired to JSON-backed workbooks instead of Excel files."""

    def __init__(self, bundle: WorkbookBundle) -> None:
        self.path = bundle.template_workbook
        self.wb_formula = bundle.instantiate_formula_workbook()
        self.wb_cached = bundle.instantiate_cached_workbook()
        self.formulas: Dict[Any, str] = {}
        # Rebuild the formula map that the legacy interpreter expects.
        for sheet_name, sheet_bundle in bundle.formula_sheets.items():
            worksheet = self.wb_formula[sheet_name]
            for row, col in sheet_bundle.formula_cells:
                raw = worksheet._get_value(row, col)
                if isinstance(raw, str) and raw.startswith("=") and raw.strip() != "=":
                    self.formulas[LEGACY.CellRef(sheet_name, row, col)] = raw[1:]
        self.cell_cache: Dict[Any, Any] = {}
        self.parse_cache: Dict[str, Any] = {}
        self.range_cache: Dict[Any, Any] = {}
        # COUNTIF/SUMIF dominate runtime in the original workbook formulas, so
        # they get dedicated aggregate caches instead of rescanning ranges.
        self.countif_cache: Dict[Any, Dict[Any, int]] = {}
        self.sumif_cache: Dict[Any, Dict[Any, float]] = {}
        self.evaluating = set()

    def set_cell(self, ref: str, value: Any) -> None:
        """Delegate to the legacy engine and invalidate aggregate scan caches."""
        super().set_cell(ref, value)
        self.countif_cache.clear()
        self.sumif_cache.clear()

    def _get_cell(self, sheet: str, row: int, col: int) -> Any:
        """Override unstable PR coefficient value cells with stable fallbacks."""
        stable_specs = PR_COEFF_STABLE_VALUE_COLUMNS.get(sheet)
        if stable_specs and col in stable_specs:
            ref = LEGACY.CellRef(sheet, row, col)
            cached = self.cell_cache.get(ref)
            if cached is not None:
                return cached

            raw_value = LEGACY.CryptWorkbookEngine._get_cell(self, sheet, row, col)
            stable_value = self._compute_stable_pr_coeff_value(sheet, row, col)
            if self._should_replace_pr_coeff_value(sheet, raw_value, stable_value):
                self.cell_cache[ref] = stable_value
                return stable_value
            return raw_value

        return LEGACY.CryptWorkbookEngine._get_cell(self, sheet, row, col)

    def _should_replace_pr_coeff_value(self, sheet: str, raw_value: Any, stable_value: Any) -> bool:
        """Swap in the stable fallback only when the legacy replay is clearly unstable."""
        try:
            raw_numeric = float(raw_value)
            stable_numeric = float(stable_value)
        except Exception:
            return False
        if not math.isfinite(raw_numeric) or not math.isfinite(stable_numeric):
            return False

        magnitude = max(abs(raw_numeric), abs(stable_numeric))
        if magnitude <= 1.0e-30:
            return False
        if raw_numeric == stable_numeric:
            return False
        if raw_numeric == 0.0 or stable_numeric == 0.0:
            return True

        ratio = abs(raw_numeric / stable_numeric)
        sign_flip = raw_numeric * stable_numeric < 0.0

        if sheet == PR_COEFF_IRR_SHEET and raw_numeric < 0.0 < stable_numeric:
            return True
        return sign_flip or ratio > 10.0 or ratio < 0.1

    def _compute_stable_pr_coeff_value(self, sheet: str, row: int, col: int) -> float | None:
        """Evaluate the direct PR coefficient value columns at high precision."""
        stable_specs = PR_COEFF_STABLE_VALUE_COLUMNS.get(sheet, {})
        spec = stable_specs.get(col)
        if spec is None:
            return None

        try:
            nodes = tuple(
                self._decimal_from_cell(sheet, row, column_index_from_string(col_text))
                for col_text in spec["node_columns"]
            )
            factors = tuple(
                self._decimal_from_cell(sheet, row, column_index_from_string(col_text))
                for col_text in spec["factor_columns"]
            )
        except Exception:
            return None

        try:
            with localcontext() as ctx:
                ctx.prec = 80
                elapsed_time = self._decimal_from_cell(
                    sheet,
                    2,
                    column_index_from_string("B"),
                )
                anchored = self._exp_divided_difference(nodes, elapsed_time)
                product = anchored
                for factor in factors:
                    product *= factor
                return float(product)
        except Exception:
            return None

    def _decimal_from_cell(self, sheet: str, row: int, col: int) -> Decimal:
        value = LEGACY.CryptWorkbookEngine._get_cell(self, sheet, row, col)
        if isinstance(value, bool):
            return Decimal(1 if value else 0)
        return Decimal(str(value))

    def _exp_divided_difference(
        self,
        nodes: Tuple[Decimal, ...],
        elapsed_time: Decimal,
    ) -> Decimal:
        if len(nodes) < 1:
            raise ValueError("At least one node is required.")
        if len(nodes) == 1:
            return (nodes[0] * elapsed_time).exp()
        if len(nodes) == 2:
            left, right = nodes
            denominator = left - right
            if denominator == 0:
                raise ZeroDivisionError("Duplicate exponential nodes.")
            return ((left * elapsed_time).exp() - (right * elapsed_time).exp()) / denominator

        denominator = nodes[0] - nodes[1]
        if denominator == 0:
            raise ZeroDivisionError("Duplicate anchored nodes.")
        left = self._exp_divided_difference((nodes[0], *nodes[2:]), elapsed_time)
        right = self._exp_divided_difference((nodes[1], *nodes[2:]), elapsed_time)
        return (left - right) / denominator

    def _resolve_range_node(
        self,
        node: Any,
        current_sheet: str,
    ) -> tuple[str, int, int, int, int] | None:
        """Normalize a parsed range node into sheet/min/max row/column bounds."""
        if not isinstance(node, tuple) or len(node) != 3 or node[0] != "range":
            return None
        start_node, end_node = node[1], node[2]
        if (
            not isinstance(start_node, tuple)
            or len(start_node) != 4
            or start_node[0] != "cell"
            or not isinstance(end_node, tuple)
            or len(end_node) != 4
            or end_node[0] != "cell"
        ):
            return None
        start_sheet = start_node[1] or current_sheet
        end_sheet = end_node[1] or start_sheet
        if start_sheet != end_sheet:
            return None
        min_row, max_row = sorted((int(start_node[2]), int(end_node[2])))
        min_col, max_col = sorted((int(start_node[3]), int(end_node[3])))
        return (start_sheet, min_row, min_col, max_row, max_col)

    def _normalize_criteria_key(self, value: Any) -> Any:
        """Collapse equivalent numeric criteria values onto one cache key."""
        if value in (None, ""):
            return None
        if isinstance(value, (int, float)):
            numeric = float(value)
            return int(numeric) if numeric.is_integer() else numeric
        try:
            numeric = float(value)
        except Exception:
            return value
        return int(numeric) if numeric.is_integer() else numeric

    def _build_countif_map(
        self,
        range_node: Any,
        current_sheet: str,
    ) -> Dict[Any, int] | None:
        """Scan a range once and cache the frequency of each normalized value."""
        resolved = self._resolve_range_node(range_node, current_sheet)
        if resolved is None:
            return None
        cached = self.countif_cache.get(resolved)
        if cached is not None:
            return cached

        sheet, min_row, min_col, max_row, max_col = resolved
        counts: Dict[Any, int] = {}
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                try:
                    raw_value = self._get_cell(sheet, row, col)
                except Exception:
                    raw_value = None
                key = self._normalize_criteria_key(raw_value)
                counts[key] = counts.get(key, 0) + 1
        self.countif_cache[resolved] = counts
        return counts

    def _build_sumif_map(
        self,
        criteria_range_node: Any,
        sum_range_node: Any,
        current_sheet: str,
    ) -> Dict[Any, float] | None:
        """Scan paired ranges once and cache the SUMIF totals by criteria value."""
        criteria_range = self._resolve_range_node(criteria_range_node, current_sheet)
        sum_range = self._resolve_range_node(sum_range_node, current_sheet)
        if criteria_range is None or sum_range is None:
            return None

        criteria_sheet, crit_min_row, crit_min_col, crit_max_row, crit_max_col = criteria_range
        sum_sheet, sum_min_row, sum_min_col, sum_max_row, sum_max_col = sum_range
        if (
            (crit_max_row - crit_min_row) != (sum_max_row - sum_min_row)
            or (crit_max_col - crit_min_col) != (sum_max_col - sum_min_col)
        ):
            return None

        cache_key = (criteria_range, sum_range)
        cached = self.sumif_cache.get(cache_key)
        if cached is not None:
            return cached

        sums: Dict[Any, float] = {}
        for row_offset in range(crit_max_row - crit_min_row + 1):
            for col_offset in range(crit_max_col - crit_min_col + 1):
                try:
                    crit_value = self._get_cell(
                        criteria_sheet,
                        crit_min_row + row_offset,
                        crit_min_col + col_offset,
                    )
                except Exception:
                    crit_value = None
                crit_key = self._normalize_criteria_key(crit_value)
                try:
                    sum_value = self._get_cell(
                        sum_sheet,
                        sum_min_row + row_offset,
                        sum_min_col + col_offset,
                    )
                except Exception:
                    sum_value = 0.0
                try:
                    numeric_sum = float(sum_value)
                except Exception:
                    numeric_sum = 0.0
                sums[crit_key] = sums.get(crit_key, 0.0) + numeric_sum
        self.sumif_cache[cache_key] = sums
        return sums

    def _eval_func(self, fn: str, arg_nodes: List[Any], current_sheet: str, current_row: int) -> Any:
        """Intercept expensive aggregate functions before falling back to legacy logic."""
        normalized_fn = fn.upper()
        if normalized_fn == "COUNTIF" and len(arg_nodes) == 2:
            counts = self._build_countif_map(arg_nodes[0], current_sheet)
            if counts is not None:
                criterion = self._normalize_criteria_key(self._eval(arg_nodes[1], current_sheet, current_row))
                return counts.get(criterion, 0)
        if normalized_fn == "SUMIF" and len(arg_nodes) in (2, 3):
            sum_range_node = arg_nodes[2] if len(arg_nodes) == 3 else arg_nodes[0]
            sums = self._build_sumif_map(arg_nodes[0], sum_range_node, current_sheet)
            if sums is not None:
                criterion = self._normalize_criteria_key(self._eval(arg_nodes[1], current_sheet, current_row))
                return sums.get(criterion, 0.0)
        return super()._eval_func(fn, arg_nodes, current_sheet, current_row)


class JsonEngineFactory:
    """Small factory that owns the loaded bundle and creates fresh engines."""
    def __init__(self, bundle_dir: Path) -> None:
        self.bundle = load_bundle(str(bundle_dir))
        self.legacy = LEGACY

    @property
    def template_workbook(self) -> Path:
        return self.bundle.template_workbook

    @property
    def metadata_path(self) -> Path:
        return self.bundle.metadata_path

    def create_engine(self) -> JsonBackedEngine:
        """Return a new engine so each calculation starts from clean workbook state."""
        return JsonBackedEngine(self.bundle)

    def export_validation_workbook(self, overrides: Dict[str, Any], out_path: Path) -> None:
        """Write a workbook reconstructed from the bundled JSON sheets."""
        out_path.parent.mkdir(parents=True, exist_ok=True)
        workbook = build_validation_workbook(self.bundle, overrides)
        try:
            workbook.save(out_path)
        finally:
            workbook.close()

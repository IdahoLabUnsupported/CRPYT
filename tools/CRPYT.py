# Copyright 2026, Battelle Energy Alliance, LLC. All Rights Reserved.
from __future__ import annotations
"""Main standalone CRPYT GUI.

This module ties together the JSON-backed CRPYT engine, the GREASY spectrum
model, and the Tk/Matplotlib front end. The goal is to preserve the workbook
workflow while moving the actual calculation path into Python.
"""

from dataclasses import dataclass
from datetime import datetime
import math
import threading
import traceback
import tkinter as tk
import json
import os
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk
from typing import Any, Dict, List, Sequence, Tuple
import sys

import matplotlib
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.cm import ScalarMappable
from matplotlib.colors import LinearSegmentedColormap, Normalize
from matplotlib.figure import Figure
from openpyxl.utils import column_index_from_string

from GREASPY import (
    SHIELDING_MATERIALS,
    CalcInputs as GreaspyCalcInputs,
    CalcResults as GreasyCalcResults,
    CrpytGreaseSpectrumModel,
    CrpytSpectrumInputs,
    default_crpyt_workbook_candidates,
)
from CRPYT_export import infer_spectrum_format, write_spectrum_file
from spectrum_import import (
    build_linear_energy_edges,
    read_imported_spectrum_csv,
    read_imported_spectrum_file,
    read_imported_spectrum_workbook,
    rebase_spectrum_to_target_grid,
)
from extract_crpyt_irdff_data import (
    BUNDLE_NAME,
    DEFAULT_COMBINED_SPECTRUM_NAME,
    DEFAULT_COVER_MATERIAL,
    DEFAULT_DECAY_TIME,
    DEFAULT_ENERGY_MAX,
    DEFAULT_ENERGY_MIN,
    DEFAULT_FLUENCE_RATE,
    DEFAULT_IRRADIATION_TIME,
    DEFAULT_REACTOR,
    extract_bundle,
)
from extract_greaspy_crpyt_data import (
    BUNDLE_NAME as GREASY_BUNDLE_NAME,
    ensure_bundle as ensure_greasy_bundle,
)
from json_engine import JsonEngineFactory
from element_optimizer import (
    OptimizationInputs,
    OptimizationResult,
    SpectrumMixEntry,
    optimize_element_set,
)
from results_builder import aggregate_results, build_results, load_reference_metadata


COMBINED_REACTOR_UI_LABEL = "Combined Spectrum"
COMBINED_REACTOR_WORKBOOK_VALUE = "Combo"
IMPORT_SPECTRUM_UI_LABEL = "Import Spectrum"
DEFAULT_WINDOW_SIZE = "1680x980"
DEFAULT_ELEMENT_MASS_VALUE = 0.0
DEFAULT_ELEMENT_MASS_TEXT = "0"
DEFAULT_FLUENCE_RATE_TEXT = "1e12"
DEFAULT_ELEMENTS_MIN_MASS_VALUE = 1.0e-15
DEFAULT_ELEMENTS_MIN_MASS_TEXT = "1e-15"
DEFAULT_ISOTOPE_MIN_ACTIVITY_CI_VALUE = 1.0e-11
DEFAULT_ISOTOPE_MIN_ACTIVITY_CI_TEXT = "1e-11"
DEFAULT_GAMMA_MIN_RATE_VALUE = 1.0
DEFAULT_GAMMA_MIN_RATE_TEXT = "1"
DEFAULT_ENERGY_MIN_TEXT = "1e-9"
DEFAULT_ENERGY_MAX_TEXT = "20"
PLOT_REACTION_ACTIVITY_SHARE_THRESHOLD = 0.01
DEFAULT_SPECTRUM_COUNT_TIME_TEXT = "300"
DEFAULT_SPECTRUM_DISTANCE_TEXT = "5"
DEFAULT_OPTIMIZATION_TRIAL_MASS_TEXT = "1"
DEFAULT_OPTIMIZATION_MIN_ACTIVITY_TEXT = DEFAULT_ISOTOPE_MIN_ACTIVITY_CI_TEXT
DEFAULT_OPTIMIZATION_MIN_MASS_TEXT = "0"
DEFAULT_OPTIMIZATION_MAX_MASS_TEXT = ""
DEFAULT_OPTIMIZATION_MAX_ELEMENTS_TEXT = "8"
SPECTRUM_MIN_DISPLAY_ENERGY_KEV = 25.0
SPECTRUM_LOG_YMIN = 0.01
SPECTRUM_LINEAR_YMIN = 0.0
SPECTRUM_ENTRY_WIDTH = 12
COMBOBOX_WIDTH = 28
IRRADIATION_ENTRY_WIDTH = 16
FILTER_ENTRY_WIDTH = 14
COVER_ENTRY_WIDTH = 10
MASS_ENTRY_WIDTH = 10
LEGEND_REACTION_LIMIT = 20
HEATMAP_ZERO_COLOR = "#f2f2f2"
HEATMAP_ZERO_TRANSITION = 0.01
BUILD_VERSION = "1.0.3"
BUILD_DATE = "2026-08-04"
USER_SETTINGS_FILENAME = "CRPYT.user_settings.json"
USER_SETTINGS_SCHEMA_VERSION = 2
DEFAULT_PROJECT_STATE_NAME = "Default"
REPORT_PAGE_SIZE = (8.5, 11.0)
REPORT_LANDSCAPE_PAGE_SIZE = (11.0, 8.5)
AUTHORS = (
    "Dr. Michael Reichenber",
    "Dr. Tommy Holschuh",
)
AUTHOR_AFFILIATION = "Idaho National Laboratory"
# Nominal alloy recipes inferred from standard 304/316 stainless chemistry and
# normalized over the elemental inputs already supported by the CRPYT bundle.
STAINLESS_STEEL_RECIPE_SPECS = (
    {
        "id": "SS304",
        "label": "SS304",
        "components": {
            "Cr": 18.0,
            "Ni": 8.0,
            "Mn": 2.0,
            "Si": 1.0,
            "P": 0.04,
            "S": 0.03,
        },
    },
    {
        "id": "SS316",
        "label": "SS316",
        "components": {
            "Cr": 17.0,
            "Ni": 12.0,
            "Mo": 2.5,
            "Mn": 2.0,
            "Si": 1.0,
            "P": 0.04,
            "S": 0.03,
        },
    },
)
ABOUT_FEATURES = (
    "Standalone CRPYT calculations backed by extracted JSON data instead of live workbook formulas.",
    "GUI inputs for reactor type, fluence rate, irradiation time, decay time, global cover thicknesses, element-by-element cover assignment, and element masses.",
    "Live result filtering for element mass, isotope atoms, isotope activity, gamma rate, and energy bounds.",
    "Interactive reaction heatmap plot with zoom/pan, log or linear scaling, energy-window display, and minor-path filtering.",
    "Reaction-band layout toggle for reaction-rate weighting or even vertical spacing.",
    "GREASPY detector spectrum generation with counting time, distance, background, shielding inputs, and .spe/.chn/.n42/.spc export.",
    "Tabbed Summary, Elements, Isotopes, Spectrum, and Optimization views.",
    "Validation workbook export reconstructed from bundled JSON data for comparison studies.",
    "Element-set optimization to recommend candidate dosimetry materials that improve spectrum coverage.",
)

ELEMENT_HEADERS = ("Element", "Mass at BOI (g)", "Mass at EOID (g)")
ISOTOPE_HEADERS = (
    "Element",
    "ZAID",
    "Isotope",
    "BOI Atoms",
    "EOID Atoms",
    "BOI Mass (g)",
    "EOID Mass (g)",
    "Activity (Ci)",
    "Gamma Rate (Bq)",
    "Gamma Dose (R/hr)",
    "Beta Rate (Bq)",
    "Beta Dose (R/hr)",
)
ISOTOPE_REFERENCE_HEADERS = ("ZAID", "Isotope", "BOI Atoms", "EOID Atoms")
GAMMA_HEADERS = ("ZAID", "Isotope", "Energy (keV)", "Gammas / sec")
OPTIMIZATION_HEADERS = (
    "Rank",
    "Element",
    "Cover",
    "Reaction",
    "Recommended Mass (g)",
    "Estimated Activity (Ci)",
    "Reaction Rate",
    "Centroid Energy (MeV)",
    "Incremental Coverage",
    "Cumulative Coverage",
)
OPTIMIZATION_POOL_HEADERS = ("Element", "Status")


def _default_workbook_path() -> Path:
    """Return the conventional workbook path used when regenerating JSON data."""
    script_dir = Path(__file__).resolve().parent
    candidates = [
        script_dir / "CRPYT_IRDFF_Rev1.xlsx",
        script_dir.parent / "CRPYT_IRDFF_Rev1.xlsx",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0]


def _bundle_dir() -> Path:
    """Return the extracted CRPYT JSON bundle directory."""
    return Path(__file__).resolve().parent / "data" / BUNDLE_NAME


def _default_greasy_workbook_path() -> Path | None:
    """Locate the source GREASY workbook if the JSON bundle needs to be built."""
    for candidate in default_crpyt_workbook_candidates():
        if candidate.exists():
            return candidate
    return None


def _greasy_bundle_dir() -> Path:
    """Return the extracted GREASY JSON bundle directory."""
    return Path(__file__).resolve().parent / "data" / GREASY_BUNDLE_NAME


def _preferred_settings_dir() -> Path:
    """Return the first-choice directory for the persisted GUI settings file."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _fallback_settings_dir() -> Path:
    """Return a per-user writable fallback directory for GUI settings."""
    appdata = os.environ.get("APPDATA")
    if appdata:
        return Path(appdata) / "CRPYT"
    return Path.home() / ".crpyt"


def ensure_bundle(workbook_path: Path, bundle_dir: Path) -> None:
    """Build the CRPYT JSON bundle on first run if it does not already exist."""
    manifest = bundle_dir / "manifest.json"
    if manifest.exists():
        return
    if not workbook_path.exists():
        raise FileNotFoundError(
            f"Workbook not found at {workbook_path}; cannot build the JSON bundle."
        )
    extract_bundle(workbook_path, bundle_dir)


@dataclass(frozen=True)
class ImportedSpectrum:
    """One user-imported spectrum rebinned onto the CRPYT reactor grid."""

    name: str
    path: Path
    values: Tuple[float, ...]
    source_points: int
    integral: float


class CrpytGuiApp:
    """Tk application that mirrors the workbook UI with Python-native calculations."""

    def __init__(self, root: tk.Tk, bundle_dir: Path, workbook_path: Path | None = None) -> None:
        workbook_path = workbook_path or _default_workbook_path()
        ensure_bundle(workbook_path, bundle_dir)
        self.root = root
        self.bundle_dir = bundle_dir
        self.engine_factory = JsonEngineFactory(bundle_dir)
        # Load the extracted CRPYT bundle and the reference metadata used by the
        # native result builder.
        self.metadata = load_reference_metadata(self.engine_factory.metadata_path)
        self.gui_config = self._prepare_gui_config(self.metadata.gui_config)
        self.base_reactor_options = tuple(self.gui_config["reactor_options"])
        self.base_spectrum_options = tuple(self.gui_config["spectrum_options"])
        self.greasy_bundle_dir = _greasy_bundle_dir()
        self.spectrum_model: CrpytGreaseSpectrumModel | None = None
        self.spectrum_model_error: str | None = None
        # GREASY spectrum generation is optional at startup, but we attempt to
        # make the bundle available up front so the Spectrum tab can work
        # without touching the workbook at runtime.
        try:
            ensure_greasy_bundle(_default_greasy_workbook_path(), self.greasy_bundle_dir)
            self.spectrum_model = CrpytGreaseSpectrumModel(self.greasy_bundle_dir)
        except Exception as exc:
            self.spectrum_model_error = str(exc)

        self.current_results: Dict[str, Any] | None = None
        self.current_overrides: Dict[str, Any] | None = None
        self.current_engine: Any | None = None
        self.current_group_runs: List[Dict[str, Any]] = []
        self.current_export_unavailable_reason: str | None = None
        self.current_plot_payload: Dict[str, Any] | None = None
        self.current_spectrum_results: GreasyCalcResults | None = None
        self.current_optimization_result: OptimizationResult | None = None
        self.current_run_overrides: Dict[str, Any] | None = None
        self.current_run_reactor_display: str | None = None
        self.current_run_thresholds: Dict[str, float] | None = None
        self.current_run_input_specs: List[Dict[str, Any]] = []
        self.run_thread: threading.Thread | None = None
        self.export_thread: threading.Thread | None = None
        self.optimize_thread: threading.Thread | None = None

        self.root.title("CRPYT Standalone")
        self.root.geometry(DEFAULT_WINDOW_SIZE)
        self._closing = False
        self._configure_styles()

        # Tk state variables are kept here so the GUI, live filters, and plots
        # all observe the same source of truth.
        self.status_var = tk.StringVar(value="Ready")
        self.project_var = tk.StringVar(value=DEFAULT_PROJECT_STATE_NAME)
        self.reactor_var = tk.StringVar(value=self.gui_config["default_reactor"])
        self.fluence_var = tk.StringVar(value=DEFAULT_FLUENCE_RATE_TEXT)
        self.irradiation_var = tk.StringVar(value=str(self.gui_config["default_irradiation_time"]))
        self.decay_var = tk.StringVar(value=str(self.gui_config["default_decay_time"]))
        self.energy_min_var = tk.StringVar(value=DEFAULT_ENERGY_MIN_TEXT)
        self.energy_max_var = tk.StringVar(value=DEFAULT_ENERGY_MAX_TEXT)
        self.irradiation_days_var = tk.StringVar(value="--")
        self.decay_days_var = tk.StringVar(value="--")
        self.total_fluence_var = tk.StringVar(value="--")
        self.total_mass_var = tk.StringVar(value="--")

        thresholds = self.gui_config["result_thresholds"]
        self.elements_threshold_var = tk.StringVar(value=DEFAULT_ELEMENTS_MIN_MASS_TEXT)
        self.isotopes_threshold_var = tk.StringVar(value=str(thresholds["isotopes_min_atoms"]))
        self.isotopes_activity_threshold_var = tk.StringVar(
            value=str(thresholds["isotopes_min_activity_ci"])
        )
        self.gamma_threshold_var = tk.StringVar(value=DEFAULT_GAMMA_MIN_RATE_TEXT)
        self.gamma_scale_var = tk.StringVar(value="log")
        self.spectrum_count_time_var = tk.StringVar(value=DEFAULT_SPECTRUM_COUNT_TIME_TEXT)
        self.spectrum_distance_var = tk.StringVar(value=DEFAULT_SPECTRUM_DISTANCE_TEXT)
        self.spectrum_include_bg_var = tk.BooleanVar(value=True)
        self.sensitivity_scale_var = tk.StringVar(value="log")
        self.sensitivity_major_paths_only_var = tk.BooleanVar(value=True)
        self.sensitivity_band_scale_var = tk.StringVar(value="reaction_rate")
        self.sensitivity_note_var = tk.StringVar(value="Run a calculation to populate the plot.")
        self.spectrum_summary_var = tk.StringVar(value=self._default_spectrum_summary_text())
        self.optimization_trial_mass_var = tk.StringVar(value=DEFAULT_OPTIMIZATION_TRIAL_MASS_TEXT)
        self.optimization_min_activity_var = tk.StringVar(value=DEFAULT_OPTIMIZATION_MIN_ACTIVITY_TEXT)
        self.optimization_min_mass_var = tk.StringVar(value=DEFAULT_OPTIMIZATION_MIN_MASS_TEXT)
        self.optimization_max_mass_var = tk.StringVar(value=DEFAULT_OPTIMIZATION_MAX_MASS_TEXT)
        self.optimization_max_elements_var = tk.StringVar(value=DEFAULT_OPTIMIZATION_MAX_ELEMENTS_TEXT)
        self.optimization_coverage_mode_var = tk.StringVar(value="log")
        self.optimization_major_paths_only_var = tk.BooleanVar(value=True)
        self.optimization_include_var = tk.StringVar(value="")
        self.optimization_include_menu_summary_var = tk.StringVar(value="All elements")
        self.optimization_exclude_var = tk.StringVar(value="")
        self.optimization_exclude_menu_summary_var = tk.StringVar(value="Select elements...")
        self.optimization_cover_thickness_var = tk.StringVar(value="0")
        self.optimization_cover_menu_summary_var = tk.StringVar(value="No covers")
        self.optimization_summary_var = tk.StringVar(
            value=(
                "Run optimization to rank candidate elements. "
                "The optimizer uses irradiation settings from the Inputs tab. "
                "Optional cover scenarios are configured on the Optimization tab. "
                "Leave Include Elements blank to consider all available elements."
            )
        )

        self.cover_vars: Dict[str, tk.StringVar] = {
            entry["cell"]: tk.StringVar(value=str(entry["default"]))
            for entry in self.gui_config["cover_materials"]
        }
        mass_input_entries = self._all_mass_input_entries()
        self.mass_vars: Dict[Any, tk.StringVar] = {
            entry["row"]: tk.StringVar(value=DEFAULT_ELEMENT_MASS_TEXT)
            for entry in mass_input_entries
        }
        self.mass_bare_vars: Dict[Any, tk.BooleanVar] = {
            entry["row"]: tk.BooleanVar(value=True)
            for entry in mass_input_entries
        }
        self.mass_cover_vars: Dict[Tuple[Any, str], tk.BooleanVar] = {
            (entry["row"], str(cover_entry["cell"])): tk.BooleanVar(value=False)
            for entry in mass_input_entries
            for cover_entry in self.gui_config["cover_materials"]
        }
        self.spectrum_name_vars: Dict[int, tk.StringVar] = {
            entry["row"]: tk.StringVar(value=entry["default_name"])
            for entry in self.gui_config["spectrum_rows"]
        }
        self.spectrum_fraction_vars: Dict[int, tk.StringVar] = {
            entry["row"]: tk.StringVar(value=str(entry["default_fraction"]))
            for entry in self.gui_config["spectrum_rows"]
        }
        self.spectrum_shield_vars: Dict[str, tk.StringVar] = {
            material: tk.StringVar(value="0")
            for material in SHIELDING_MATERIALS
        }
        self.optimization_exclude_vars: Dict[str, tk.BooleanVar] = {
            str(entry["element"]): tk.BooleanVar(value=False)
            for entry in self.gui_config["composition_rows"]
        }
        self.optimization_include_vars: Dict[str, tk.BooleanVar] = {
            str(entry["element"]): tk.BooleanVar(value=False)
            for entry in self.gui_config["composition_rows"]
        }
        self.optimization_cover_vars: Dict[str, tk.BooleanVar] = {
            str(entry["label"]): tk.BooleanVar(value=False)
            for entry in self.gui_config["cover_materials"]
        }
        self.mass_indicator_labels: Dict[Any, ttk.Label] = {}
        self.mass_element_labels: Dict[Any, ttk.Label] = {}
        self.mass_cover_buttons: Dict[Tuple[Any, str], ttk.Checkbutton] = {}

        self.summary_labels: Dict[str, ttk.Label] = {}
        self.elements_tree = None
        self.elements_all_tree = None
        self.isotopes_tree = None
        self.isotopes_ref_tree = None
        self.gamma_tree = None
        self.optimization_tree = None
        self.optimization_pool_tree = None
        self.reactor_combobox: ttk.Combobox | None = None
        self.spectrum_comboboxes: List[ttk.Combobox] = []

        self.elements_axes = None
        self.elements_canvas = None
        self.isotopes_axes = None
        self.isotopes_canvas = None
        self.gamma_axes = None
        self.gamma_canvas = None
        self.gamma_toolbar = None
        self.sensitivity_axes = None
        self.sensitivity_secondary_axes = None
        self.sensitivity_canvas = None
        self.sensitivity_toolbar = None
        self.sensitivity_colorbar = None
        self.sensitivity_colorbar_axis = None
        self.sensitivity_figure = None
        self.sensitivity_note_label = None
        self.optimize_button = None
        self.report_button = None
        self.apply_optimization_button = None
        self.optimization_include_menu_button = None
        self.optimization_include_menu = None
        self.optimization_exclude_menu_button = None
        self.optimization_exclude_menu = None
        self.optimization_cover_menu_button = None
        self.optimization_cover_menu = None
        self.about_window: tk.Toplevel | None = None
        self.project_combobox: ttk.Combobox | None = None
        self.active_project_name = DEFAULT_PROJECT_STATE_NAME
        self.project_states: Dict[str, Dict[str, Any]] = {}
        self.settings_path: Path | None = None
        self.imported_spectra: Dict[str, ImportedSpectrum] = {}
        self.imported_spectrum_path: Path | None = None
        self.imported_spectrum_values: List[float] = []
        self.imported_spectrum_source_points = 0
        self.imported_spectrum_integral: float | None = None
        self.imported_spectrum_summary_var = tk.StringVar(value="")
        self._builtin_spectrum_values_cache: Dict[str, Tuple[float, ...]] | None = None
        self._reactor_import_prompt_pending = False
        self._last_non_import_reactor = self.gui_config["default_reactor"]
        self._spectrum_energy_grid: List[float] | None = None

        self._build_ui()
        self._load_user_settings()
        self._bind_updates()
        self._update_optimization_include_summary()
        self._update_optimization_cover_summary()
        self._update_optimization_exclude_summary()
        self._update_mass_cover_states()
        self._update_mass_row_indicators()
        self._update_derived_metrics()
        self._update_spectrum_visibility()
        self._draw_sensitivity_plot()
        self.root.protocol("WM_DELETE_WINDOW", self._handle_close)

    def _build_ui(self) -> None:
        """Construct the notebook layout and the shared bottom action bar."""
        outer = ttk.Frame(self.root, padding=10)
        outer.pack(fill="both", expand=True)

        self.notebook = ttk.Notebook(outer)
        self.notebook.pack(fill="both", expand=True)

        self.inputs_tab = ttk.Frame(self.notebook)
        self.plot_tab = ttk.Frame(self.notebook)
        self.optimization_tab = ttk.Frame(self.notebook)
        self.summary_tab = ttk.Frame(self.notebook)
        self.elements_tab = ttk.Frame(self.notebook)
        self.isotopes_tab = ttk.Frame(self.notebook)
        self.gamma_tab = ttk.Frame(self.notebook)

        self.notebook.add(self.inputs_tab, text="Inputs")
        self.notebook.add(self.plot_tab, text="Plot")
        self.notebook.add(self.summary_tab, text="Summary")
        self.notebook.add(self.elements_tab, text="Elements")
        self.notebook.add(self.isotopes_tab, text="Isotopes")
        self.notebook.add(self.gamma_tab, text="Spectrum")
        self.notebook.add(self.optimization_tab, text="Optimization")

        self._build_inputs_tab()
        self._build_plot_tab()
        self._build_summary_tab()
        self._build_elements_tab()
        self._build_isotopes_tab()
        self._build_gamma_tab()
        self._build_optimization_tab()

        bottom = ttk.Frame(outer)
        bottom.pack(fill="x", pady=(10, 0))
        self.run_button = ttk.Button(bottom, text="Run Calculation", command=self._start_run)
        self.run_button.pack(side="left", padx=(0, 6))
        self.export_button = ttk.Button(
            bottom,
            text="Export Workbook",
            command=self._start_export,
        )
        self.export_button.pack(side="left")
        self.report_button = ttk.Button(
            bottom,
            text="Generate PDF Report",
            command=self._start_report,
        )
        self.report_button.pack(side="left", padx=(6, 0))
        self.about_button = ttk.Button(bottom, text="About", command=self._show_about_dialog)
        self.about_button.pack(side="left", padx=(6, 0))
        ttk.Separator(bottom, orient="vertical").pack(side="left", fill="y", padx=(12, 8))
        ttk.Label(bottom, text="Project").pack(side="left", padx=(0, 4))
        self.project_combobox = ttk.Combobox(
            bottom,
            textvariable=self.project_var,
            values=self._project_option_names(),
            width=24,
        )
        self.project_combobox.pack(side="left", padx=(0, 6))
        self.project_combobox.configure(postcommand=self._refresh_project_options)
        ttk.Button(bottom, text="Save State", command=self._save_active_project_state).pack(
            side="left",
            padx=(0, 4),
        )
        ttk.Button(bottom, text="Load State", command=self._load_selected_project_state).pack(
            side="left",
            padx=(0, 4),
        )
        ttk.Button(bottom, text="Delete State", command=self._delete_selected_project_state).pack(
            side="left",
        )
        if not self._validation_export_available():
            self.export_button.state(["disabled"])
        ttk.Label(bottom, textvariable=self.status_var).pack(side="right")

    def _validation_export_available(self) -> bool:
        """Return True when the JSON bundle can produce validation workbooks."""
        return True

    def _project_name_from_value(self, value: Any) -> str:
        """Normalize a user-entered project state name for display and storage."""
        raw_value = "" if value is None else str(value)
        normalized = " ".join(raw_value.split())
        return normalized or DEFAULT_PROJECT_STATE_NAME

    def _project_option_names(self) -> List[str]:
        """Return saved project names with the default project listed first."""
        names = {
            self._project_name_from_value(name)
            for name, payload in self.project_states.items()
            if isinstance(payload, dict)
        }
        names.add(self._project_name_from_value(self.project_var.get()))
        ordered = sorted(names, key=str.casefold)
        if DEFAULT_PROJECT_STATE_NAME in ordered:
            ordered.remove(DEFAULT_PROJECT_STATE_NAME)
            ordered.insert(0, DEFAULT_PROJECT_STATE_NAME)
        return ordered

    def _refresh_project_options(self) -> None:
        """Refresh the project-state dropdown with the currently known names."""
        if self.project_combobox is not None:
            self.project_combobox.configure(values=self._project_option_names())

    def _save_current_project_state(self, project_name: str | None = None) -> str:
        """Capture the visible GUI state under the active project name."""
        normalized_name = self._project_name_from_value(
            self.active_project_name if project_name is None else project_name
        )
        self.project_var.set(normalized_name)
        self.active_project_name = normalized_name
        self.project_states[normalized_name] = self._build_project_state_payload()
        self._refresh_project_options()
        return normalized_name

    def _project_states_from_payload(self, payload: Dict[str, Any]) -> tuple[Dict[str, Dict[str, Any]], str]:
        """Read schema-v2 project states, treating schema-v1 files as Default."""
        raw_projects = payload.get("projects")
        if not isinstance(raw_projects, dict):
            return {DEFAULT_PROJECT_STATE_NAME: payload}, DEFAULT_PROJECT_STATE_NAME

        projects: Dict[str, Dict[str, Any]] = {}
        for raw_name, raw_state in raw_projects.items():
            if isinstance(raw_state, dict):
                projects[self._project_name_from_value(raw_name)] = raw_state
        if not projects:
            projects[DEFAULT_PROJECT_STATE_NAME] = {}

        active_project = self._project_name_from_value(payload.get("active_project"))
        if active_project not in projects:
            if DEFAULT_PROJECT_STATE_NAME in projects:
                active_project = DEFAULT_PROJECT_STATE_NAME
            else:
                active_project = sorted(projects, key=str.casefold)[0]
        return projects, active_project

    def _apply_settings_file_payload(self, payload: Dict[str, Any]) -> None:
        """Load the project-state container from disk and apply the active state."""
        self.project_states, active_project = self._project_states_from_payload(payload)
        self.active_project_name = active_project
        self.project_var.set(active_project)
        self._refresh_project_options()
        self._apply_user_settings(self.project_states[active_project])

    def _build_user_settings_payload(self) -> Dict[str, Any]:
        """Build the schema-v2 settings container written to disk."""
        active_project = self._save_current_project_state()
        projects = {
            name: state
            for name, state in sorted(self.project_states.items(), key=lambda item: item[0].casefold())
            if isinstance(state, dict)
        }
        return {
            "schema_version": USER_SETTINGS_SCHEMA_VERSION,
            "build_version": BUILD_VERSION,
            "saved_at": datetime.now().isoformat(timespec="seconds"),
            "active_project": active_project,
            "projects": projects,
        }

    def _refresh_after_project_state_change(self) -> None:
        """Refresh derived labels and plots after manually loading a project state."""
        self._update_optimization_include_summary()
        self._update_optimization_cover_summary()
        self._update_optimization_exclude_summary()
        self._update_mass_cover_states()
        self._update_mass_row_indicators()
        self._update_derived_metrics()
        self._update_spectrum_visibility()
        self._refresh_filtered_views()
        self._draw_sensitivity_plot()
        self._refresh_gamma_spectrum()

    def _save_active_project_state(self) -> None:
        """Persist the current GUI settings under the selected project name."""
        try:
            project_name = self._save_current_project_state(self.project_var.get())
            self._save_user_settings()
        except Exception as exc:
            messagebox.showwarning(
                "Settings Save Failed",
                f"The project state could not be saved.\n\n{exc}",
            )
            return
        self.status_var.set(f"Saved project state: {project_name}")

    def _load_selected_project_state(self) -> None:
        """Apply a saved project state selected in the project dropdown."""
        project_name = self._project_name_from_value(self.project_var.get())
        if project_name not in self.project_states:
            messagebox.showwarning(
                "Load State",
                f"No saved state named '{project_name}'.\n\nUse Save State to create it.",
            )
            self.project_var.set(self.active_project_name)
            self._refresh_project_options()
            return

        if project_name != self.active_project_name:
            self.project_states[self.active_project_name] = self._build_project_state_payload()
        self.project_var.set(project_name)
        self.active_project_name = project_name
        self._apply_user_settings(self.project_states[project_name])
        self._refresh_after_project_state_change()
        self.status_var.set(f"Loaded project state: {project_name}")

    def _delete_selected_project_state(self) -> None:
        """Delete the selected saved project state after confirmation."""
        project_name = self._project_name_from_value(self.project_var.get())
        if project_name not in self.project_states:
            messagebox.showwarning("Delete State", f"No saved state named '{project_name}'.")
            self.project_var.set(self.active_project_name)
            self._refresh_project_options()
            return
        if project_name == DEFAULT_PROJECT_STATE_NAME:
            messagebox.showwarning("Delete State", "The Default project state cannot be deleted.")
            return
        if not messagebox.askyesno("Delete State", f"Delete saved state '{project_name}'?"):
            self.project_var.set(self.active_project_name)
            return

        was_active = project_name == self.active_project_name
        del self.project_states[project_name]
        if not self.project_states:
            self.project_states[DEFAULT_PROJECT_STATE_NAME] = self._build_project_state_payload()
        if was_active:
            next_project = (
                DEFAULT_PROJECT_STATE_NAME
                if DEFAULT_PROJECT_STATE_NAME in self.project_states
                else sorted(self.project_states, key=str.casefold)[0]
            )
            self.project_var.set(next_project)
            self.active_project_name = next_project
            self._apply_user_settings(self.project_states[next_project])
            self._refresh_after_project_state_change()
        else:
            self.project_var.set(self.active_project_name)
        self._refresh_project_options()

        try:
            self._save_user_settings()
        except Exception as exc:
            messagebox.showwarning(
                "Settings Save Failed",
                f"The project state was removed in memory, but the settings file could not be saved.\n\n{exc}",
            )
            return
        self.status_var.set(f"Deleted project state: {project_name}")

    def _settings_path_candidates(self) -> List[Path]:
        """Return the ordered list of locations used for loading/saving settings."""
        candidates: List[Path] = []
        if self.settings_path is not None:
            candidates.append(self.settings_path)
        for base_dir in (_preferred_settings_dir(), _fallback_settings_dir()):
            candidate = base_dir / USER_SETTINGS_FILENAME
            if candidate not in candidates:
                candidates.append(candidate)
        return candidates

    def _build_project_state_payload(self) -> Dict[str, Any]:
        """Capture the current user-editable GUI state as JSON-serializable data."""
        selected_tab_id = self.notebook.select()
        selected_tab = self.notebook.tab(selected_tab_id, "text") if selected_tab_id else None
        return {
            "schema_version": 1,
            "build_version": BUILD_VERSION,
            "window_geometry": self.root.geometry(),
            "selected_tab": selected_tab,
            "inputs": {
                "reactor": self.reactor_var.get(),
                "fluence_rate": self.fluence_var.get(),
                "irradiation_time": self.irradiation_var.get(),
                "decay_time": self.decay_var.get(),
                "energy_min": self.energy_min_var.get(),
                "energy_max": self.energy_max_var.get(),
                "imported_spectrum_path": str(self.imported_spectrum_path) if self.imported_spectrum_path else "",
            },
            "filters": {
                "elements_min_mass": self.elements_threshold_var.get(),
                "isotopes_min_atoms": self.isotopes_threshold_var.get(),
                "isotopes_min_activity_ci": self.isotopes_activity_threshold_var.get(),
                "gamma_min_rate": self.gamma_threshold_var.get(),
            },
            "plot": {
                "gamma_scale": self.gamma_scale_var.get(),
                "sensitivity_scale": self.sensitivity_scale_var.get(),
                "sensitivity_major_paths_only": self.sensitivity_major_paths_only_var.get(),
                "sensitivity_band_scale": self.sensitivity_band_scale_var.get(),
            },
            "spectrum": {
                "count_time_sec": self.spectrum_count_time_var.get(),
                "distance_cm": self.spectrum_distance_var.get(),
                "include_background": self.spectrum_include_bg_var.get(),
                "names": {
                    str(row): variable.get()
                    for row, variable in self.spectrum_name_vars.items()
                },
                "fractions": {
                    str(row): variable.get()
                    for row, variable in self.spectrum_fraction_vars.items()
                },
                "shielding_cm": {
                    material: variable.get()
                    for material, variable in self.spectrum_shield_vars.items()
                },
            },
            "optimization": {
                "trial_mass_g": self.optimization_trial_mass_var.get(),
                "min_activity_ci": self.optimization_min_activity_var.get(),
                "min_mass_g": self.optimization_min_mass_var.get(),
                "max_mass_g": self.optimization_max_mass_var.get(),
                "max_elements": self.optimization_max_elements_var.get(),
                "coverage_mode": self.optimization_coverage_mode_var.get(),
                "major_paths_only": self.optimization_major_paths_only_var.get(),
                "include_elements": self.optimization_include_var.get(),
                "exclude_elements": self.optimization_exclude_var.get(),
                "cover_thickness": self.optimization_cover_thickness_var.get(),
                "included_checks": {
                    element: variable.get()
                    for element, variable in self.optimization_include_vars.items()
                },
                "excluded_checks": {
                    element: variable.get()
                    for element, variable in self.optimization_exclude_vars.items()
                },
                "cover_checks": {
                    cover_label: variable.get()
                    for cover_label, variable in self.optimization_cover_vars.items()
                },
            },
            "covers": {
                str(cell): variable.get()
                for cell, variable in self.cover_vars.items()
            },
            "masses": {
                str(row): variable.get()
                for row, variable in self.mass_vars.items()
            },
            "mass_profiles": {
                "bare": {
                    str(row): variable.get()
                    for row, variable in self.mass_bare_vars.items()
                },
                "covers": {
                    str(row): {
                        str(cell): self.mass_cover_vars[(row, cell)].get()
                        for entry in self.gui_config["cover_materials"]
                        for cell in (str(entry["cell"]),)
                    }
                    for row in self.mass_bare_vars
                },
            },
        }

    def _set_string_var(
        self,
        variable: tk.StringVar,
        values: Dict[str, Any],
        key: str,
        allowed: set[str] | None = None,
    ) -> None:
        """Restore a string-valued setting when the saved value is valid."""
        if key not in values or values[key] is None:
            return
        text = str(values[key])
        if allowed is not None and text not in allowed:
            return
        variable.set(text)

    def _set_bool_var(
        self,
        variable: tk.BooleanVar,
        values: Dict[str, Any],
        key: str,
    ) -> None:
        """Restore a boolean setting when present in the saved payload."""
        if key not in values:
            return
        variable.set(bool(values[key]))

    def _select_notebook_tab(self, tab_label: str) -> None:
        """Restore the previously active notebook tab by display label."""
        for tab_id in self.notebook.tabs():
            if self.notebook.tab(tab_id, "text") == tab_label:
                self.notebook.select(tab_id)
                return

    def _apply_user_settings(self, payload: Dict[str, Any]) -> None:
        """Restore persisted GUI settings, ignoring keys not valid for this build."""
        inputs = payload.get("inputs", {})
        filters = payload.get("filters", {})
        plot = payload.get("plot", {})
        spectrum = payload.get("spectrum", {})
        optimization = payload.get("optimization", {})
        covers = payload.get("covers", {})
        masses = payload.get("masses", {})
        mass_profiles = payload.get("mass_profiles", {})

        if not isinstance(inputs, dict):
            inputs = {}
        if not isinstance(filters, dict):
            filters = {}
        if not isinstance(plot, dict):
            plot = {}
        if not isinstance(spectrum, dict):
            spectrum = {}
        if not isinstance(optimization, dict):
            optimization = {}
        if not isinstance(covers, dict):
            covers = {}
        if not isinstance(masses, dict):
            masses = {}
        if not isinstance(mass_profiles, dict):
            mass_profiles = {}

        imported_path = inputs.get("imported_spectrum_path")
        if isinstance(imported_path, str) and imported_path.strip():
            try:
                self._load_imported_spectrum_file(Path(imported_path.strip()))
            except Exception as exc:
                self.imported_spectrum_summary_var.set(f"Imported spectrum not restored: {exc}")

        if inputs.get("reactor") == IMPORT_SPECTRUM_UI_LABEL and not self.imported_spectrum_values:
            inputs = dict(inputs)
            inputs["reactor"] = self.gui_config["default_reactor"]

        self._set_string_var(
            self.reactor_var,
            inputs,
            "reactor",
            allowed=set(self.gui_config["reactor_options"]),
        )
        if self.reactor_var.get() != IMPORT_SPECTRUM_UI_LABEL:
            self._last_non_import_reactor = self.reactor_var.get()
        self._set_string_var(self.fluence_var, inputs, "fluence_rate")
        self._set_string_var(self.irradiation_var, inputs, "irradiation_time")
        self._set_string_var(self.decay_var, inputs, "decay_time")
        self._set_string_var(self.energy_min_var, inputs, "energy_min")
        self._set_string_var(self.energy_max_var, inputs, "energy_max")

        self._set_string_var(self.elements_threshold_var, filters, "elements_min_mass")
        self._set_string_var(self.isotopes_threshold_var, filters, "isotopes_min_atoms")
        self._set_string_var(self.isotopes_activity_threshold_var, filters, "isotopes_min_activity_ci")
        self._set_string_var(self.gamma_threshold_var, filters, "gamma_min_rate")

        self._set_string_var(self.gamma_scale_var, plot, "gamma_scale", allowed={"log", "linear"})
        self._set_string_var(self.sensitivity_scale_var, plot, "sensitivity_scale", allowed={"log", "linear"})
        self._set_bool_var(self.sensitivity_major_paths_only_var, plot, "sensitivity_major_paths_only")
        self._set_string_var(
            self.sensitivity_band_scale_var,
            plot,
            "sensitivity_band_scale",
            allowed={"reaction_rate", "even"},
        )

        self._set_string_var(self.spectrum_count_time_var, spectrum, "count_time_sec")
        self._set_string_var(self.spectrum_distance_var, spectrum, "distance_cm")
        self._set_bool_var(self.spectrum_include_bg_var, spectrum, "include_background")

        spectrum_names = spectrum.get("names", {})
        if isinstance(spectrum_names, dict):
            allowed_spectrum_names = set(self._current_spectrum_options())
            for row, variable in self.spectrum_name_vars.items():
                self._set_string_var(
                    variable,
                    spectrum_names,
                    str(row),
                    allowed=allowed_spectrum_names,
                )

        spectrum_fractions = spectrum.get("fractions", {})
        if isinstance(spectrum_fractions, dict):
            for row, variable in self.spectrum_fraction_vars.items():
                self._set_string_var(variable, spectrum_fractions, str(row))

        spectrum_shielding = spectrum.get("shielding_cm", {})
        if isinstance(spectrum_shielding, dict):
            for material, variable in self.spectrum_shield_vars.items():
                self._set_string_var(variable, spectrum_shielding, material)

        self._set_string_var(self.optimization_trial_mass_var, optimization, "trial_mass_g")
        self._set_string_var(self.optimization_min_activity_var, optimization, "min_activity_ci")
        self._set_string_var(self.optimization_min_mass_var, optimization, "min_mass_g")
        self._set_string_var(self.optimization_max_mass_var, optimization, "max_mass_g")
        self._set_string_var(self.optimization_max_elements_var, optimization, "max_elements")
        self._set_string_var(
            self.optimization_coverage_mode_var,
            optimization,
            "coverage_mode",
            allowed={"log", "linear"},
        )
        self._set_bool_var(self.optimization_major_paths_only_var, optimization, "major_paths_only")
        self._set_string_var(self.optimization_include_var, optimization, "include_elements")
        self._set_string_var(self.optimization_exclude_var, optimization, "exclude_elements")
        self._set_string_var(self.optimization_cover_thickness_var, optimization, "cover_thickness")

        optimization_included = optimization.get("included_checks", {})
        if isinstance(optimization_included, dict):
            for element, variable in self.optimization_include_vars.items():
                self._set_bool_var(variable, optimization_included, element)

        optimization_excluded = optimization.get("excluded_checks", {})
        if isinstance(optimization_excluded, dict):
            for element, variable in self.optimization_exclude_vars.items():
                self._set_bool_var(variable, optimization_excluded, element)

        optimization_covers = optimization.get("cover_checks", {})
        if isinstance(optimization_covers, dict):
            for cover_label, variable in self.optimization_cover_vars.items():
                self._set_bool_var(variable, optimization_covers, cover_label)

        for cell, variable in self.cover_vars.items():
            self._set_string_var(variable, covers, str(cell))
        for row, variable in self.mass_vars.items():
            self._set_string_var(variable, masses, str(row))

        bare_profiles = mass_profiles.get("bare", {})
        if isinstance(bare_profiles, dict):
            for row, variable in self.mass_bare_vars.items():
                self._set_bool_var(variable, bare_profiles, str(row))

        cover_profiles = mass_profiles.get("covers", {})
        if isinstance(cover_profiles, dict):
            for (row, cell), variable in self.mass_cover_vars.items():
                row_settings = cover_profiles.get(str(row), {})
                if isinstance(row_settings, dict):
                    self._set_bool_var(variable, row_settings, str(cell))

        geometry = payload.get("window_geometry")
        if geometry not in (None, ""):
            self.root.geometry(str(geometry))

        selected_tab = payload.get("selected_tab")
        if isinstance(selected_tab, str) and selected_tab:
            self._select_notebook_tab(selected_tab)

    def _load_user_settings(self) -> None:
        """Load the most recent persisted settings file if one is available."""
        load_errors: List[tuple[Path, Exception]] = []
        for path in self._settings_path_candidates():
            if not path.exists():
                continue
            try:
                payload = json.loads(path.read_text(encoding="utf-8"))
                if not isinstance(payload, dict):
                    raise ValueError("Settings file must contain a JSON object.")
                self._apply_settings_file_payload(payload)
                self.settings_path = path
                return
            except Exception as exc:
                load_errors.append((path, exc))

        if load_errors:
            path, exc = load_errors[0]
            messagebox.showwarning(
                "Settings Load Failed",
                f"Could not load saved settings from:\n{path}\n\nUsing built-in defaults instead.\n\n{exc}",
            )

    def _save_user_settings(self) -> None:
        """Persist the current GUI state to JSON using the first writable path."""
        payload = self._build_user_settings_payload()
        save_errors: List[tuple[Path, Exception]] = []
        for path in self._settings_path_candidates():
            try:
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
                self.settings_path = path
                return
            except Exception as exc:
                save_errors.append((path, exc))

        if save_errors:
            path, exc = save_errors[0]
            raise OSError(f"Could not save settings to {path}: {exc}") from exc

    def _handle_close(self) -> None:
        """Save settings on exit, then close the application window."""
        if self._closing:
            return
        self._closing = True

        try:
            self._save_user_settings()
        except Exception as exc:
            messagebox.showwarning(
                "Settings Save Failed",
                f"The application could not save your settings.\n\n{exc}",
            )
        self._shutdown_matplotlib()
        try:
            self.root.quit()
        finally:
            try:
                self.root.destroy()
            except tk.TclError:
                pass

    def _shutdown_matplotlib(self) -> None:
        """Release embedded Matplotlib widgets before Tk shuts down."""
        for canvas_attr in (
            "elements_canvas",
            "isotopes_canvas",
            "gamma_canvas",
            "sensitivity_canvas",
        ):
            canvas = getattr(self, canvas_attr, None)
            if canvas is None:
                continue
            figure = getattr(canvas, "figure", None)
            try:
                canvas.get_tk_widget().destroy()
            except Exception:
                pass
            if figure is not None:
                try:
                    figure.clear()
                except Exception:
                    pass
            setattr(self, canvas_attr, None)

        for toolbar_attr in ("gamma_toolbar", "sensitivity_toolbar"):
            toolbar = getattr(self, toolbar_attr, None)
            if toolbar is None:
                continue
            try:
                toolbar.destroy()
            except Exception:
                pass
            setattr(self, toolbar_attr, None)

        self.elements_axes = None
        self.isotopes_axes = None
        self.gamma_axes = None
        self.sensitivity_axes = None
        self.sensitivity_secondary_axes = None
        self.sensitivity_colorbar = None
        self.sensitivity_colorbar_axis = None
        self.sensitivity_figure = None

    def _configure_styles(self) -> None:
        """Define the small style variants used to flag active mass inputs."""
        style = ttk.Style(self.root)
        style.configure("MassInputActive.TLabel", foreground="#0f5c2f")
        style.configure("MassInputInactive.TLabel", foreground="#7a7a7a")

    def _bind_updates(self) -> None:
        """Wire Tk variables to the lightweight live-refresh behaviors."""
        tracked = [
            self.reactor_var,
            self.fluence_var,
            self.irradiation_var,
            self.decay_var,
            self.elements_threshold_var,
            self.isotopes_threshold_var,
            self.isotopes_activity_threshold_var,
            self.gamma_threshold_var,
            *self.cover_vars.values(),
            *self.mass_vars.values(),
            *self.spectrum_name_vars.values(),
            *self.spectrum_fraction_vars.values(),
        ]
        for variable in tracked:
            variable.trace_add("write", lambda *_args: self._update_derived_metrics())
        for variable in self.cover_vars.values():
            variable.trace_add("write", lambda *_args: self._update_mass_cover_states())
        self.reactor_var.trace_add("write", lambda *_args: self._handle_reactor_selection())
        self.gamma_scale_var.trace_add("write", lambda *_args: self._draw_gamma_plot())
        self.sensitivity_scale_var.trace_add("write", lambda *_args: self._draw_sensitivity_plot())
        self.sensitivity_major_paths_only_var.trace_add("write", lambda *_args: self._refresh_filtered_views())
        self.sensitivity_band_scale_var.trace_add("write", lambda *_args: self._draw_sensitivity_plot())
        self.energy_min_var.trace_add("write", lambda *_args: self._draw_sensitivity_plot())
        self.energy_max_var.trace_add("write", lambda *_args: self._draw_sensitivity_plot())
        self.spectrum_count_time_var.trace_add("write", lambda *_args: self._refresh_gamma_spectrum())
        self.spectrum_distance_var.trace_add("write", lambda *_args: self._refresh_gamma_spectrum())
        self.spectrum_include_bg_var.trace_add("write", lambda *_args: self._refresh_gamma_spectrum())
        self.optimization_include_var.trace_add("write", lambda *_args: self._update_optimization_include_summary())
        self.optimization_exclude_var.trace_add("write", lambda *_args: self._update_optimization_exclude_summary())
        for variable in self.mass_vars.values():
            variable.trace_add("write", lambda *_args: self._update_mass_row_indicators())
        for variable in self.mass_bare_vars.values():
            variable.trace_add("write", lambda *_args: self._update_mass_row_indicators())
        for variable in self.mass_cover_vars.values():
            variable.trace_add("write", lambda *_args: self._update_mass_row_indicators())
        for variable in self.spectrum_shield_vars.values():
            variable.trace_add("write", lambda *_args: self._refresh_gamma_spectrum())
        for variable in self.optimization_include_vars.values():
            variable.trace_add("write", lambda *_args: self._update_optimization_include_summary())
        for variable in self.optimization_exclude_vars.values():
            variable.trace_add("write", lambda *_args: self._update_optimization_exclude_summary())
        for variable in self.optimization_cover_vars.values():
            variable.trace_add("write", lambda *_args: self._update_optimization_cover_summary())
        for variable in (
            self.elements_threshold_var,
            self.isotopes_threshold_var,
            self.isotopes_activity_threshold_var,
            self.gamma_threshold_var,
        ):
            variable.trace_add("write", lambda *_args: self._refresh_filtered_views())

    def _build_inputs_tab(self) -> None:
        """Create the input form that replaces the workbook's entry sheet."""
        parent = self.inputs_tab
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(3, weight=1)

        general = ttk.LabelFrame(parent, text="Irradiation Parameters", padding=10)
        general.grid(row=0, column=0, sticky="ew", padx=8, pady=8)
        general.columnconfigure(4, weight=1)

        ttk.Label(general, text="Reactor Type").grid(row=0, column=0, sticky="w")
        self.reactor_combobox = ttk.Combobox(
            general,
            textvariable=self.reactor_var,
            values=self._current_reactor_options(),
            state="readonly",
            width=COMBOBOX_WIDTH,
        )
        self.reactor_combobox.grid(row=0, column=1, sticky="w", padx=(6, 18))
        ttk.Label(
            general,
            textvariable=self.imported_spectrum_summary_var,
        ).grid(row=0, column=4, sticky="w")

        self._add_numeric_entry(
            general,
            "Neutron Fluence Rate (n/cm^2/sec)",
            self.fluence_var,
            1,
            width=IRRADIATION_ENTRY_WIDTH,
            sticky="w",
        )
        self._add_numeric_entry(
            general,
            "Irradiation Time (sec)",
            self.irradiation_var,
            2,
            width=IRRADIATION_ENTRY_WIDTH,
            sticky="w",
        )
        self._add_numeric_entry(
            general,
            "Decay Time (sec)",
            self.decay_var,
            3,
            width=IRRADIATION_ENTRY_WIDTH,
            sticky="w",
        )

        self._add_readonly_metric(general, "Irradiation Time (days)", self.irradiation_days_var, 2, 2)
        self._add_readonly_metric(general, "Decay Time (days)", self.decay_days_var, 3, 2)
        self._add_readonly_metric(general, "Total Fluence (n/cm^2)", self.total_fluence_var, 1, 2)
        self._add_readonly_metric(general, "Total Mass (g)", self.total_mass_var, 0, 2)

        filters = ttk.LabelFrame(parent, text="Result Filters", padding=10)
        filters.grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 8))
        filters.columnconfigure(2, weight=1)
        self._add_numeric_entry(
            filters,
            "Minimum Mass at EOID (g)",
            self.elements_threshold_var,
            0,
            width=FILTER_ENTRY_WIDTH,
            sticky="w",
        )
        self._add_numeric_entry(
            filters,
            "Minimum Atoms of Isotope Provided",
            self.isotopes_threshold_var,
            1,
            width=FILTER_ENTRY_WIDTH,
            sticky="w",
        )
        self._add_numeric_entry(
            filters,
            "Minimum Activity (Ci)",
            self.isotopes_activity_threshold_var,
            2,
            width=FILTER_ENTRY_WIDTH,
            sticky="w",
        )
        self._add_numeric_entry(
            filters,
            "Minimum Gamma Rate (Bq)",
            self.gamma_threshold_var,
            3,
            width=FILTER_ENTRY_WIDTH,
            sticky="w",
        )
        self._add_numeric_entry(
            filters,
            "Minimum Energy (MeV)",
            self.energy_min_var,
            4,
            width=FILTER_ENTRY_WIDTH,
            sticky="w",
        )
        self._add_numeric_entry(
            filters,
            "Maximum Energy (MeV)",
            self.energy_max_var,
            5,
            width=FILTER_ENTRY_WIDTH,
            sticky="w",
        )

        lower = ttk.Panedwindow(parent, orient="horizontal")
        lower.grid(row=2, column=0, sticky="nsew", padx=8, pady=(0, 8))
        self.lower_paned = lower

        covers = ttk.LabelFrame(lower, text="Cover Materials", padding=10)
        lower.add(covers, weight=1)
        covers.columnconfigure(2, weight=1)
        for idx, entry in enumerate(self.gui_config["cover_materials"]):
            ttk.Label(covers, text=entry["label"]).grid(row=idx, column=0, sticky="w", pady=2)
            ttk.Entry(
                covers,
                textvariable=self.cover_vars[entry["cell"]],
                width=COVER_ENTRY_WIDTH,
            ).grid(row=idx, column=1, sticky="w", padx=(6, 0), pady=2)

        self.spectrum_frame = ttk.LabelFrame(lower, text="Combined Spectrum", padding=10)
        lower.add(self.spectrum_frame, weight=1)
        spectrum_options = self._current_spectrum_options()
        ttk.Label(self.spectrum_frame, text="Spectrum").grid(row=0, column=0, sticky="w")
        ttk.Label(self.spectrum_frame, text="Fraction").grid(row=0, column=1, sticky="w")
        for idx, entry in enumerate(self.gui_config["spectrum_rows"], start=1):
            row_number = entry["row"]
            ttk.Label(self.spectrum_frame, text=entry["label"]).grid(row=idx, column=0, sticky="w", pady=2)
            spectrum_combobox = ttk.Combobox(
                self.spectrum_frame,
                textvariable=self.spectrum_name_vars[row_number],
                values=spectrum_options,
                state="readonly",
                width=COMBOBOX_WIDTH,
            )
            spectrum_combobox.grid(row=idx, column=1, sticky="w", padx=(6, 12), pady=2)
            self.spectrum_comboboxes.append(spectrum_combobox)
            ttk.Entry(
                self.spectrum_frame,
                textvariable=self.spectrum_fraction_vars[row_number],
                width=12,
            ).grid(row=idx, column=2, sticky="w", pady=2)
        self.spectrum_frame.columnconfigure(0, weight=1)

        masses_frame = ttk.LabelFrame(parent, text="Element Masses", padding=10)
        masses_frame.grid(row=3, column=0, sticky="nsew", padx=8, pady=(0, 8))
        masses_frame.columnconfigure(0, weight=1)
        masses_frame.rowconfigure(1, weight=1)
        ttk.Label(
            masses_frame,
            text=(
                "Checked Bare adds a bare run; checked cover columns apply the nonzero cover profile above. "
                "SS304 and SS316 are nominal recipe rows that expand into the supported elemental inputs."
            ),
            justify="left",
            wraplength=1100,
        ).grid(row=0, column=0, sticky="w", pady=(0, 6))

        canvas = tk.Canvas(masses_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(masses_frame, orient="vertical", command=canvas.yview)
        interior = ttk.Frame(canvas)

        interior.bind(
            "<Configure>",
            lambda _event: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas_window = canvas.create_window((0, 0), window=interior, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind(
            "<Configure>",
            lambda event: canvas.itemconfigure(canvas_window, width=event.width),
        )
        canvas.grid(row=1, column=0, sticky="nsew")
        scrollbar.grid(row=1, column=1, sticky="ns")

        filler_column = 5 + len(self.gui_config["cover_materials"])
        interior.columnconfigure(filler_column, weight=1)

        ttk.Label(interior, text="Set").grid(row=0, column=0, sticky="w")
        ttk.Label(interior, text="Z").grid(row=0, column=1, sticky="w")
        ttk.Label(interior, text="Element").grid(row=0, column=2, sticky="w")
        ttk.Label(interior, text="Mass (g)").grid(row=0, column=3, sticky="w")
        ttk.Label(interior, text="Bare").grid(row=0, column=4, sticky="w", padx=(6, 0))
        for cover_index, entry in enumerate(self.gui_config["cover_materials"], start=5):
            ttk.Label(
                interior,
                text=self._cover_column_label(str(entry["label"])),
            ).grid(row=0, column=cover_index, sticky="w", padx=(6, 0))
        for idx, entry in enumerate(self._all_mass_input_entries(), start=1):
            self._add_mass_input_row(interior, idx, entry)

    def _add_mass_input_row(
        self,
        parent: ttk.Frame,
        table_row: int,
        entry: Dict[str, Any],
    ) -> None:
        """Render one real-element or synthetic-alloy row in the mass table."""
        input_key = entry["row"]
        indicator_label = ttk.Label(
            parent,
            text="",
            width=4,
            style="MassInputInactive.TLabel",
        )
        indicator_label.grid(row=table_row, column=0, sticky="w", pady=2)
        self.mass_indicator_labels[input_key] = indicator_label

        ttk.Label(parent, text=str(entry["atomic_number"])).grid(row=table_row, column=1, sticky="w", pady=2)
        element_label = ttk.Label(parent, text=str(entry["element"]))
        element_label.grid(row=table_row, column=2, sticky="w", padx=(8, 8), pady=2)
        self.mass_element_labels[input_key] = element_label
        ttk.Entry(
            parent,
            textvariable=self.mass_vars[input_key],
            width=MASS_ENTRY_WIDTH,
        ).grid(row=table_row, column=3, sticky="w", padx=(0, 8), pady=2)
        ttk.Checkbutton(
            parent,
            variable=self.mass_bare_vars[input_key],
            takefocus=False,
        ).grid(row=table_row, column=4, sticky="w", padx=(6, 0), pady=2)

        # Each row can opt into a bare run and any nonzero global cover
        # thickness profile by checking the corresponding columns here.
        for cover_index, cover_entry in enumerate(self.gui_config["cover_materials"], start=5):
            key = (input_key, str(cover_entry["cell"]))
            checkbox = ttk.Checkbutton(
                parent,
                variable=self.mass_cover_vars[key],
                takefocus=False,
            )
            checkbox.grid(row=table_row, column=cover_index, sticky="w", padx=(6, 0), pady=2)
            self.mass_cover_buttons[key] = checkbox

    def _build_plot_tab(self) -> None:
        """Create the reaction-rate heatmap tab and its toolbar controls."""
        parent = self.plot_tab
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        controls = ttk.Frame(parent, padding=(8, 8, 8, 0))
        controls.grid(row=0, column=0, sticky="ew")
        ttk.Label(controls, text="Y Axis").pack(side="left")
        ttk.Radiobutton(
            controls,
            text="Log",
            value="log",
            variable=self.sensitivity_scale_var,
        ).pack(side="left", padx=(8, 4))
        ttk.Radiobutton(
            controls,
            text="Linear",
            value="linear",
            variable=self.sensitivity_scale_var,
        ).pack(side="left")
        ttk.Checkbutton(
            controls,
            text="Hide paths <1% of isotope activity",
            variable=self.sensitivity_major_paths_only_var,
        ).pack(side="left", padx=(14, 0))
        ttk.Label(controls, text="Band Height").pack(side="left", padx=(14, 0))
        ttk.Radiobutton(
            controls,
            text="Reaction Rate",
            value="reaction_rate",
            variable=self.sensitivity_band_scale_var,
        ).pack(side="left", padx=(8, 4))
        ttk.Radiobutton(
            controls,
            text="Even",
            value="even",
            variable=self.sensitivity_band_scale_var,
        ).pack(side="left")

        plot_outer = ttk.LabelFrame(parent, text="Input Spectrum and Active Reaction Sensitivities", padding=8)
        plot_outer.grid(row=1, column=0, sticky="nsew", padx=8, pady=8)
        plot_outer.columnconfigure(0, weight=1)
        plot_outer.rowconfigure(0, weight=1)

        figure = Figure(figsize=(8.8, 5.0))
        self.sensitivity_figure = figure
        self._reset_sensitivity_axes(figure)
        self.sensitivity_canvas = FigureCanvasTkAgg(figure, master=plot_outer)
        self.sensitivity_canvas.get_tk_widget().grid(row=0, column=0, sticky="nsew")
        self.sensitivity_note_label = ttk.Label(
            plot_outer,
            textvariable=self.sensitivity_note_var,
            justify="left",
            anchor="w",
        )
        self.sensitivity_note_label.grid(row=1, column=0, sticky="ew", pady=(8, 6))
        self.sensitivity_toolbar = NavigationToolbar2Tk(self.sensitivity_canvas, plot_outer, pack_toolbar=False)
        self.sensitivity_toolbar.grid(row=2, column=0, sticky="ew")
        self.sensitivity_toolbar.update()
        plot_outer.bind("<Configure>", self._handle_sensitivity_plot_resize, add="+")
        self._update_sensitivity_note_wrap()

    def _build_optimization_tab(self) -> None:
        """Build the element-set optimization tab around the reusable backend."""
        parent = self.optimization_tab
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        controls = ttk.LabelFrame(parent, text="Optimization Controls", padding=10)
        controls.grid(row=0, column=0, sticky="ew", padx=8, pady=8)
        controls.columnconfigure(7, weight=1)

        ttk.Label(
            controls,
            text=(
                "Uses reactor, fluence, irradiation time, decay time, energy bounds, "
                "and combined-spectrum settings from the Inputs tab. "
                "Optimizer cover scenarios below are evaluated independently."
            ),
            justify="left",
        ).grid(row=0, column=0, columnspan=6, sticky="w", pady=(0, 8))

        ttk.Label(controls, text="Trial Mass Per Element (g)").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(
            controls,
            textvariable=self.optimization_trial_mass_var,
            width=FILTER_ENTRY_WIDTH,
        ).grid(row=1, column=1, sticky="w", padx=(6, 16), pady=4)

        ttk.Label(controls, text="Minimum Reaction Activity (Ci)").grid(row=1, column=2, sticky="w", pady=4)
        ttk.Entry(
            controls,
            textvariable=self.optimization_min_activity_var,
            width=FILTER_ENTRY_WIDTH,
        ).grid(row=1, column=3, sticky="w", padx=(6, 16), pady=4)

        ttk.Label(controls, text="Minimum Mass (g)").grid(row=1, column=4, sticky="w", pady=4)
        ttk.Entry(
            controls,
            textvariable=self.optimization_min_mass_var,
            width=FILTER_ENTRY_WIDTH,
        ).grid(row=1, column=5, sticky="w", padx=(6, 16), pady=4)

        ttk.Label(controls, text="Maximum Mass (g; blank = none)").grid(row=1, column=6, sticky="w", pady=4)
        ttk.Entry(
            controls,
            textvariable=self.optimization_max_mass_var,
            width=FILTER_ENTRY_WIDTH,
        ).grid(row=1, column=7, sticky="w", padx=(6, 0), pady=4)

        ttk.Label(controls, text="Coverage Weighting").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Radiobutton(
            controls,
            text="Log",
            value="log",
            variable=self.optimization_coverage_mode_var,
        ).grid(row=2, column=1, sticky="w", padx=(6, 4), pady=4)
        ttk.Radiobutton(
            controls,
            text="Linear",
            value="linear",
            variable=self.optimization_coverage_mode_var,
        ).grid(row=2, column=2, sticky="w", pady=4)
        ttk.Label(controls, text="Max Elements").grid(row=2, column=3, sticky="w", pady=4)
        ttk.Entry(
            controls,
            textvariable=self.optimization_max_elements_var,
            width=8,
        ).grid(row=2, column=4, sticky="w", padx=(6, 16), pady=4)
        ttk.Checkbutton(
            controls,
            text="Major paths only",
            variable=self.optimization_major_paths_only_var,
        ).grid(row=2, column=5, sticky="w", pady=4)
        self.apply_optimization_button = ttk.Button(
            controls,
            text="Apply Recommended Set to Inputs and Run",
            command=self._apply_optimization_to_inputs_and_run,
        )
        self.apply_optimization_button.grid(row=2, column=6, sticky="e", padx=(6, 6), pady=4)
        self.apply_optimization_button.state(["disabled"])
        self.optimize_button = ttk.Button(
            controls,
            text="Optimize Element Set",
            command=self._start_optimization,
        )
        self.optimize_button.grid(row=2, column=7, sticky="e", pady=4)

        ttk.Label(controls, text="Include Elements (check and/or type; blank = all)").grid(row=3, column=0, sticky="w", pady=4)
        self.optimization_include_menu_button = ttk.Menubutton(
            controls,
            textvariable=self.optimization_include_menu_summary_var,
            direction="below",
        )
        self.optimization_include_menu_button.grid(row=3, column=1, sticky="ew", padx=(6, 12), pady=4)
        include_menu = tk.Menu(self.optimization_include_menu_button, tearoff=False)
        include_menu.add_command(label="Clear All", command=self._clear_optimization_inclusions)
        include_menu.add_separator()
        for element in sorted(self.optimization_include_vars):
            include_menu.add_checkbutton(
                label=element,
                variable=self.optimization_include_vars[element],
            )
        self.optimization_include_menu_button["menu"] = include_menu
        self.optimization_include_menu = include_menu
        ttk.Entry(
            controls,
            textvariable=self.optimization_include_var,
            width=42,
        ).grid(row=3, column=2, columnspan=2, sticky="ew", padx=(6, 16), pady=4)

        ttk.Label(controls, text="Exclude Elements (check and/or type)").grid(row=4, column=0, sticky="w", pady=4)
        self.optimization_exclude_menu_button = ttk.Menubutton(
            controls,
            textvariable=self.optimization_exclude_menu_summary_var,
            direction="below",
        )
        self.optimization_exclude_menu_button.grid(row=4, column=1, sticky="ew", padx=(6, 12), pady=4)
        exclude_menu = tk.Menu(self.optimization_exclude_menu_button, tearoff=False)
        exclude_menu.add_command(label="Clear All", command=self._clear_optimization_exclusions)
        exclude_menu.add_separator()
        for element in sorted(self.optimization_exclude_vars):
            exclude_menu.add_checkbutton(
                label=element,
                variable=self.optimization_exclude_vars[element],
            )
        self.optimization_exclude_menu_button["menu"] = exclude_menu
        self.optimization_exclude_menu = exclude_menu
        ttk.Entry(
            controls,
            textvariable=self.optimization_exclude_var,
            width=42,
        ).grid(row=4, column=2, columnspan=6, sticky="ew", padx=(6, 0), pady=4)

        ttk.Label(controls, text="Optimizer Covers").grid(row=5, column=0, sticky="w", pady=4)
        self.optimization_cover_menu_button = ttk.Menubutton(
            controls,
            textvariable=self.optimization_cover_menu_summary_var,
            direction="below",
        )
        self.optimization_cover_menu_button.grid(row=5, column=1, columnspan=2, sticky="ew", padx=(6, 16), pady=4)
        cover_menu = tk.Menu(self.optimization_cover_menu_button, tearoff=False)
        cover_menu.add_command(label="Clear All", command=self._clear_optimization_cover_scenarios)
        cover_menu.add_separator()
        for cover_label in self.optimization_cover_vars:
            cover_menu.add_checkbutton(
                label=cover_label,
                variable=self.optimization_cover_vars[cover_label],
            )
        self.optimization_cover_menu_button["menu"] = cover_menu
        self.optimization_cover_menu = cover_menu

        ttk.Label(controls, text="Cover Thickness (mm)").grid(row=5, column=3, sticky="w", pady=4)
        ttk.Entry(
            controls,
            textvariable=self.optimization_cover_thickness_var,
            width=FILTER_ENTRY_WIDTH,
        ).grid(row=5, column=4, sticky="w", padx=(6, 16), pady=4)
        ttk.Label(
            controls,
            text="Evaluates bare plus one run per selected cover at this thickness.",
        ).grid(row=5, column=5, columnspan=3, sticky="w", pady=4)

        ttk.Label(
            controls,
            textvariable=self.optimization_summary_var,
            justify="left",
            wraplength=1200,
        ).grid(row=6, column=0, columnspan=8, sticky="w", pady=(8, 0))

        lower = ttk.Panedwindow(parent, orient="horizontal")
        lower.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))

        selected_frame = ttk.LabelFrame(lower, text="Recommended Set", padding=8)
        lower.add(selected_frame, weight=3)
        self.optimization_tree = self._create_tree(selected_frame, OPTIMIZATION_HEADERS)
        self.optimization_tree.column("Reaction", width=360, anchor="w")
        self.optimization_tree.column("Element", width=90, anchor="w")
        self.optimization_tree.column("Cover", width=220, anchor="w")

        pool_frame = ttk.LabelFrame(lower, text="Candidate Pool", padding=8)
        lower.add(pool_frame, weight=1)
        self.optimization_pool_tree = self._create_tree(pool_frame, OPTIMIZATION_POOL_HEADERS)
        self.optimization_pool_tree.column("Element", width=120, anchor="w")
        self.optimization_pool_tree.column("Status", width=240, anchor="w")

    def _reset_sensitivity_axes(self, figure: Figure) -> None:
        """Recreate the plot, reaction-label, and colorbar axes from scratch."""
        figure.clear()
        grid = figure.add_gridspec(1, 2, width_ratios=[30, 1.8], wspace=0.22)
        self.sensitivity_axes = figure.add_subplot(grid[0])
        self.sensitivity_secondary_axes = self.sensitivity_axes.twinx()
        self.sensitivity_colorbar_axis = figure.add_subplot(grid[1])
        self.sensitivity_colorbar_axis.set_visible(False)
        self.sensitivity_colorbar = None

    def _handle_sensitivity_plot_resize(self, event: tk.Event) -> None:
        """Keep the sensitivity note wrapped to the visible plot width."""
        width = getattr(event, "width", 0)
        self._update_sensitivity_note_wrap(width)

    def _update_sensitivity_note_wrap(self, width: int | None = None) -> None:
        """Wrap the sensitivity note so it stays inside the plot frame."""
        if self.sensitivity_note_label is None:
            return
        if width is None or width <= 1:
            width = self.sensitivity_note_label.winfo_width()
        wraplength = max(int(width) - 24, 240)
        self.sensitivity_note_label.configure(wraplength=wraplength)

    def _ensure_sensitivity_axes(self) -> bool:
        """Recover detached Matplotlib axes after redraws or widget recreation."""
        if self.sensitivity_canvas is None:
            return False
        figure = self.sensitivity_canvas.figure
        if (
            self.sensitivity_axes is None
            or self.sensitivity_secondary_axes is None
            or self.sensitivity_colorbar_axis is None
            or self.sensitivity_axes.figure is not figure
            or self.sensitivity_secondary_axes.figure is not figure
            or self.sensitivity_colorbar_axis.figure is not figure
        ):
            self._reset_sensitivity_axes(figure)
        self.sensitivity_figure = figure
        return True

    def _build_summary_tab(self) -> None:
        """Build the compact summary page for the headline workbook metrics."""
        parent = self.summary_tab
        summary = ttk.Frame(parent, padding=20)
        summary.pack(fill="both", expand=True)
        keys = [
            "Total Activity (Ci)",
            "Total Activity (Bq)",
            "Gamma-Ray Emission Rate (Bq)",
            "Gamma Dose Rate (R/hr) at 30 cm",
            "Beta Emission Rate (Bq)",
            "Beta Dose Rate (R/hr) at Contact",
            "Total Gammas Emitted Per Second",
            "Minimum Gamma Ray Emission Rate (Bq)",
            "Fraction Above Threshold",
        ]
        for row, key in enumerate(keys):
            ttk.Label(summary, text=key).grid(row=row, column=0, sticky="w", pady=4)
            label = ttk.Label(summary, text="--")
            label.grid(row=row, column=1, sticky="w", padx=(16, 0), pady=4)
            self.summary_labels[key] = label
        summary.columnconfigure(1, weight=1)

    def _build_elements_tab(self) -> None:
        """Build the Elements tab with a filtered view and the full element table."""
        parent = self.elements_tab
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        plot_frame = ttk.Frame(parent, padding=(8, 8, 8, 0))
        plot_frame.grid(row=0, column=0, sticky="nsew")
        figure = Figure(figsize=(7.5, 3.2), tight_layout=True)
        self.elements_axes = figure.add_subplot(111)
        self.elements_canvas = FigureCanvasTkAgg(figure, master=plot_frame)
        self.elements_canvas.get_tk_widget().pack(fill="both", expand=True)

        lower = ttk.Panedwindow(parent, orient="horizontal")
        lower.grid(row=1, column=0, sticky="nsew", padx=8, pady=8)

        filtered_frame = ttk.LabelFrame(lower, text="Filtered Results", padding=8)
        lower.add(filtered_frame, weight=1)
        self.elements_tree = self._create_tree(filtered_frame, ELEMENT_HEADERS)

        all_frame = ttk.LabelFrame(lower, text="All Elements", padding=8)
        lower.add(all_frame, weight=1)
        self.elements_all_tree = self._create_tree(all_frame, ELEMENT_HEADERS)

    def _build_isotopes_tab(self) -> None:
        """Build the isotope tab with plots plus filtered and reference tables."""
        parent = self.isotopes_tab
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        plot_frame = ttk.Frame(parent, padding=(8, 8, 8, 0))
        plot_frame.grid(row=0, column=0, sticky="nsew")
        figure = Figure(figsize=(7.5, 3.2), tight_layout=True)
        self.isotopes_axes = figure.add_subplot(111)
        self.isotopes_canvas = FigureCanvasTkAgg(figure, master=plot_frame)
        self.isotopes_canvas.get_tk_widget().pack(fill="both", expand=True)

        lower = ttk.Panedwindow(parent, orient="horizontal")
        lower.grid(row=1, column=0, sticky="nsew", padx=8, pady=8)

        main_frame = ttk.LabelFrame(lower, text="Isotope Results", padding=8)
        lower.add(main_frame, weight=3)
        self.isotopes_tree = self._create_tree(main_frame, ISOTOPE_HEADERS)

        ref_frame = ttk.LabelFrame(lower, text="Reference Rows", padding=8)
        lower.add(ref_frame, weight=2)
        self.isotopes_ref_tree = self._create_tree(ref_frame, ISOTOPE_REFERENCE_HEADERS)

    def _build_gamma_tab(self) -> None:
        """Build the GREASPY spectrum tab and the gamma-line table below it."""
        parent = self.gamma_tab
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)

        inputs_outer = ttk.LabelFrame(parent, text="Spectrum Inputs", padding=8)
        inputs_outer.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 0))
        inputs_outer.columnconfigure(5, weight=1)

        ttk.Label(inputs_outer, text="Counting Time (sec)").grid(row=0, column=0, sticky="w")
        ttk.Entry(
            inputs_outer,
            textvariable=self.spectrum_count_time_var,
            width=SPECTRUM_ENTRY_WIDTH,
        ).grid(row=0, column=1, sticky="w", padx=(6, 12))
        ttk.Label(inputs_outer, text="Detector Distance (cm)").grid(row=0, column=2, sticky="w")
        ttk.Entry(
            inputs_outer,
            textvariable=self.spectrum_distance_var,
            width=SPECTRUM_ENTRY_WIDTH,
        ).grid(row=0, column=3, sticky="w", padx=(6, 12))
        ttk.Checkbutton(
            inputs_outer,
            text="Include Background",
            variable=self.spectrum_include_bg_var,
        ).grid(row=0, column=4, sticky="w")
        ttk.Label(inputs_outer, text="Y Axis").grid(row=0, column=5, sticky="e")
        ttk.Radiobutton(inputs_outer, text="Log", value="log", variable=self.gamma_scale_var).grid(
            row=0,
            column=6,
            sticky="w",
            padx=(8, 4),
        )
        ttk.Radiobutton(inputs_outer, text="Linear", value="linear", variable=self.gamma_scale_var).grid(
            row=0,
            column=7,
            sticky="w",
        )
        ttk.Button(
            inputs_outer,
            text="Export Spectrum",
            command=self._export_detector_spectrum,
        ).grid(row=0, column=8, sticky="e", padx=(12, 0))

        shielding_frame = ttk.Frame(inputs_outer)
        shielding_frame.grid(row=1, column=0, columnspan=9, sticky="ew", pady=(8, 4))
        for index, material in enumerate(SHIELDING_MATERIALS):
            col = (index % 3) * 2
            row = index // 3
            ttk.Label(shielding_frame, text=material).grid(row=row, column=col, sticky="w", pady=2)
            ttk.Entry(
                shielding_frame,
                textvariable=self.spectrum_shield_vars[material],
                width=8,
            ).grid(row=row, column=col + 1, sticky="w", padx=(6, 16), pady=2)

        ttk.Label(
            inputs_outer,
            textvariable=self.spectrum_summary_var,
            justify="left",
        ).grid(row=2, column=0, columnspan=9, sticky="w", pady=(6, 0))

        plot_outer = ttk.LabelFrame(parent, text="Detected Gamma Spectrum (GREASPY)", padding=8)
        plot_outer.grid(row=1, column=0, sticky="nsew", padx=8, pady=(8, 0))
        plot_outer.columnconfigure(0, weight=1)
        plot_outer.rowconfigure(0, weight=1)

        figure = Figure(figsize=(8.5, 3.5), tight_layout=True)
        self.gamma_axes = figure.add_subplot(111)
        self.gamma_canvas = FigureCanvasTkAgg(figure, master=plot_outer)
        self.gamma_canvas.get_tk_widget().grid(row=0, column=0, sticky="nsew")
        self.gamma_toolbar = NavigationToolbar2Tk(self.gamma_canvas, plot_outer, pack_toolbar=False)
        self.gamma_toolbar.grid(row=1, column=0, sticky="ew")
        self.gamma_toolbar.update()

        table_frame = ttk.LabelFrame(parent, text="Gamma Lines Above Threshold", padding=8)
        table_frame.grid(row=2, column=0, sticky="nsew", padx=8, pady=8)
        self.gamma_tree = self._create_tree(table_frame, GAMMA_HEADERS)

    def _create_tree(self, parent: ttk.Widget, headers: tuple[str, ...]) -> ttk.Treeview:
        """Create a scrollable Treeview with consistent column sizing."""
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(frame, columns=headers, show="headings")
        y_scroll = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        x_scroll = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        tree.pack(side="left", fill="both", expand=True)
        y_scroll.pack(side="right", fill="y")
        x_scroll.pack(side="bottom", fill="x")
        for header in headers:
            tree.heading(header, text=header)
            tree.column(header, width=max(110, len(header) * 10), anchor="w")
        return tree

    def _add_numeric_entry(
        self,
        parent: ttk.Widget,
        label: str,
        variable: tk.StringVar,
        row: int,
        *,
        width: int | None = None,
        sticky: str = "ew",
    ) -> None:
        """Place a label plus numeric entry using the shared two-column pattern."""
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=4)
        entry_kwargs: Dict[str, Any] = {"textvariable": variable}
        if width is not None:
            entry_kwargs["width"] = width
        ttk.Entry(parent, **entry_kwargs).grid(
            row=row,
            column=1,
            sticky=sticky,
            padx=(6, 18),
            pady=4,
        )

    def _add_readonly_metric(
        self,
        parent: ttk.Widget,
        label: str,
        variable: tk.StringVar,
        row: int,
        column: int,
    ) -> None:
        """Place a derived read-only metric beside the editable inputs."""
        ttk.Label(parent, text=label).grid(row=row, column=column, sticky="w", pady=4)
        ttk.Label(parent, textvariable=variable).grid(row=row, column=column + 1, sticky="w", padx=(6, 18), pady=4)

    def _cover_column_label(self, label: str) -> str:
        """Compress verbose cover labels so the element table stays readable."""
        cleaned = label.replace(" Thickness (mm)", "").strip()
        if cleaned == "Boron Carbide (B4C)":
            return "B4C"
        return cleaned

    def _get_row_cover_profile(
        self,
        row: Any,
        active_cover_cells: tuple[str, ...] | None = None,
    ) -> tuple[str, ...]:
        """Return the active covered profile assigned to one element row."""
        if active_cover_cells is None:
            active_cover_cells = tuple(
                str(entry["cell"])
                for entry in self.gui_config["cover_materials"]
                if (self._float_or_none(self.cover_vars[str(entry["cell"])].get()) or 0.0) > 0.0
            )

        active_cells: List[str] = []
        for cell in active_cover_cells:
            if self.mass_cover_vars[(row, cell)].get():
                active_cells.append(cell)
        return tuple(active_cells)

    def _get_row_run_profiles(
        self,
        row: Any,
        active_cover_cells: tuple[str, ...] | None = None,
    ) -> tuple[tuple[str, ...], ...]:
        """Return every scenario that should be evaluated for one element row."""
        cover_profile = self._get_row_cover_profile(row, active_cover_cells)
        profiles: List[tuple[str, ...]] = []

        # If no cover is selected, keep the legacy behavior and run the element
        # bare even if the user toggled the explicit Bare checkbox off.
        if self.mass_bare_vars[row].get() or not cover_profile:
            profiles.append(tuple())
        if cover_profile:
            profiles.append(tuple(cover_profile))
        return tuple(profiles)

    def _cover_profile_display(self, cover_cells: tuple[str, ...]) -> str:
        """Format one element's cover combination for plot labels and messages."""
        if not cover_cells:
            return "Bare"
        labels = []
        for entry in self.gui_config["cover_materials"]:
            cell = str(entry["cell"])
            if cell in cover_cells:
                labels.append(self._cover_column_label(str(entry["label"])))
        return "+".join(labels) if labels else "Bare"

    def _update_mass_cover_states(self) -> None:
        """Enable cover checkboxes only when that global cover thickness is nonzero."""
        for entry in self.gui_config["cover_materials"]:
            cell = str(entry["cell"])
            is_active = (self._float_or_none(self.cover_vars[cell].get()) or 0.0) > 0.0
            for composition in self._all_mass_input_entries():
                key = (composition["row"], cell)
                button = self.mass_cover_buttons.get(key)
                if button is None:
                    continue
                if is_active:
                    button.state(["!disabled"])
                else:
                    button.state(["disabled"])

    def _update_mass_row_indicators(self) -> None:
        """Highlight element rows whose mass input is currently nonzero."""
        for row, variable in self.mass_vars.items():
            numeric_value = self._float_or_none(variable.get())
            is_active = numeric_value is not None and abs(numeric_value) > 0.0

            indicator_label = self.mass_indicator_labels.get(row)
            if indicator_label is not None:
                indicator_label.configure(
                    text="set" if is_active else "",
                    style="MassInputActive.TLabel" if is_active else "MassInputInactive.TLabel",
                )

            element_label = self.mass_element_labels.get(row)
            if element_label is not None:
                element_label.configure(
                    style="MassInputActive.TLabel" if is_active else "TLabel",
                )

    def _prepare_gui_config(self, gui_config: Dict[str, Any]) -> Dict[str, Any]:
        """Normalize bundle metadata into the exact defaults used by the GUI."""
        prepared = dict(gui_config)

        reactor_options = [str(option).strip() for option in gui_config.get("reactor_options", []) if option not in (None, "")]
        if not reactor_options:
            reactor_options = self._bundle_reactor_options()

        normalized_reactors: List[str] = []
        seen = set()
        for option in reactor_options:
            ui_option = (
                COMBINED_REACTOR_UI_LABEL
                if option == COMBINED_REACTOR_WORKBOOK_VALUE
                else option
            )
            if ui_option and ui_option not in seen:
                seen.add(ui_option)
                normalized_reactors.append(ui_option)
        if COMBINED_REACTOR_UI_LABEL not in seen:
            seen.add(COMBINED_REACTOR_UI_LABEL)
            normalized_reactors.append(COMBINED_REACTOR_UI_LABEL)
        if IMPORT_SPECTRUM_UI_LABEL not in seen:
            normalized_reactors.append(IMPORT_SPECTRUM_UI_LABEL)

        spectrum_options = [
            option
            for option in normalized_reactors
            if option not in {COMBINED_REACTOR_UI_LABEL, IMPORT_SPECTRUM_UI_LABEL}
        ]
        if not spectrum_options:
            spectrum_options = [DEFAULT_COMBINED_SPECTRUM_NAME]

        default_reactor = DEFAULT_REACTOR
        if default_reactor not in normalized_reactors:
            default_reactor = normalized_reactors[0] if normalized_reactors else COMBINED_REACTOR_UI_LABEL

        prepared["reactor_options"] = normalized_reactors
        prepared["spectrum_options"] = spectrum_options
        prepared["default_reactor"] = default_reactor
        prepared["default_fluence_rate"] = DEFAULT_FLUENCE_RATE
        prepared["default_irradiation_time"] = DEFAULT_IRRADIATION_TIME
        prepared["default_decay_time"] = DEFAULT_DECAY_TIME
        prepared["default_energy_min"] = DEFAULT_ENERGY_MIN
        prepared["default_energy_max"] = DEFAULT_ENERGY_MAX

        prepared["spectrum_rows"] = []
        for entry in gui_config["spectrum_rows"]:
            updated = dict(entry)
            default_name = str(updated.get("default_name") or "").strip()
            if default_name not in spectrum_options:
                updated["default_name"] = spectrum_options[0]
            prepared["spectrum_rows"].append(updated)

        prepared["cover_materials"] = [
            {**entry, "default": DEFAULT_COVER_MATERIAL}
            for entry in gui_config["cover_materials"]
        ]

        prepared["composition_rows"] = [
            {**entry, "default_mass": DEFAULT_ELEMENT_MASS_VALUE}
            for entry in gui_config["composition_rows"]
        ]
        prepared["material_recipes"] = self._build_material_recipes(prepared["composition_rows"])

        thresholds = dict(gui_config.get("result_thresholds", {}))
        thresholds["elements_min_mass"] = DEFAULT_ELEMENTS_MIN_MASS_VALUE
        thresholds["isotopes_min_activity_ci"] = DEFAULT_ISOTOPE_MIN_ACTIVITY_CI_VALUE
        thresholds["gamma_min_rate"] = DEFAULT_GAMMA_MIN_RATE_VALUE
        prepared["result_thresholds"] = thresholds
        return prepared

    def _build_material_recipes(self, composition_rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Create synthetic alloy rows that expand into supported elemental inputs."""
        rows_by_element = {
            str(entry["element"]).strip().upper(): dict(entry)
            for entry in composition_rows
        }
        fe_entry = rows_by_element.get("FE")
        recipes: List[Dict[str, Any]] = []
        for spec in STAINLESS_STEEL_RECIPE_SPECS:
            raw_components = dict(spec["components"])
            specified = sum(float(value) for value in raw_components.values())
            if fe_entry is not None and specified < 100.0:
                raw_components["Fe"] = 100.0 - specified

            supported_components: List[Dict[str, Any]] = []
            supported_total = 0.0
            for element_name, weight_percent in raw_components.items():
                if weight_percent <= 0.0:
                    continue
                row_entry = rows_by_element.get(str(element_name).strip().upper())
                if row_entry is None:
                    continue
                supported_total += float(weight_percent)
                supported_components.append(
                    {
                        "row": int(row_entry["row"]),
                        "element": str(row_entry["element"]),
                        "atomic_number": int(row_entry["atomic_number"]),
                        "weight_percent": float(weight_percent),
                    }
                )

            if supported_total <= 0.0:
                continue

            components = [
                {
                    **component,
                    "mass_fraction": component["weight_percent"] / supported_total,
                }
                for component in supported_components
            ]
            recipes.append(
                {
                    "row": str(spec["id"]),
                    "atomic_number": "mix",
                    "element": str(spec["label"]),
                    "default_mass": DEFAULT_ELEMENT_MASS_VALUE,
                    "components": components,
                }
            )
        return recipes

    def _all_mass_input_entries(self) -> List[Dict[str, Any]]:
        """Return the real element rows plus synthetic alloy recipe rows."""
        return [
            *self.gui_config["composition_rows"],
            *self.gui_config.get("material_recipes", []),
        ]

    def _bundle_reactor_options(self) -> List[str]:
        """Fall back to reactor names stored on the extracted input sheet."""
        workbook = self.engine_factory.bundle.instantiate_formula_workbook()
        try:
            ws = workbook["Choose Irradiation Parameters"]
        except KeyError:
            return []

        options: List[str] = []
        seen = set()
        for row in range(4, 128):
            value = ws._get_value(row, 90)
            if value in (None, ""):
                continue
            option = str(value).strip()
            if option and option not in seen:
                seen.add(option)
                options.append(option)
        return options

    def _float_or_none(self, value: Any) -> float | None:
        """Best-effort float conversion for workbook-style cell values."""
        if value in (None, ""):
            return None
        try:
            return float(value)
        except Exception:
            return None

    def _int_or_none(self, value: Any) -> int | None:
        """Convert numeric workbook values to ints when present."""
        numeric = self._float_or_none(value)
        if numeric is None:
            return None
        return int(numeric)

    def _format_mt_meta(self, mt_value: Any, daughter_zaid: int | None) -> str:
        """Format MT values the same way the workbook annotates metastable products."""
        text = str(int(float(mt_value))) if mt_value not in (None, "") else ""
        if daughter_zaid is not None and daughter_zaid % 1000 > 400:
            return f"{text}m"
        return text

    def _build_isotope_name_map(self, engine: Any | None = None) -> Dict[int, str]:
        """Build a ZAID-to-name lookup from both radioactive and stable sheets."""
        engine = engine or self.current_engine
        if engine is None:
            return {}
        ws = engine.wb_formula["Isotope Properties"]
        lookup: Dict[int, str] = {}
        for row in range(2, ws.max_row + 1):
            zaid = self._int_or_none(ws._get_value(row, 1))
            name = ws._get_value(row, 2)
            if zaid is not None and name not in (None, "") and zaid not in lookup:
                lookup[zaid] = str(name)
        stable_ws = engine.wb_formula["Stable Isotope Properties"]
        for row in range(3, stable_ws.max_row + 1):
            zaid = self._int_or_none(stable_ws._get_value(row, 1))
            name = stable_ws._get_value(row, 3)
            if zaid is not None and name not in (None, "") and zaid not in lookup:
                lookup[zaid] = str(name)
        return lookup

    def _build_other_mt_lookup(self, engine: Any | None = None) -> Dict[tuple[int, str], str]:
        """Map each parent/MT pair to the cross-section column used in Other MT."""
        engine = engine or self.current_engine
        if engine is None:
            return {}
        lookup: Dict[tuple[int, str], str] = {}
        ws = engine.wb_formula["Isotopics"]
        for row in range(2, ws.max_row + 1):
            zaid = self._int_or_none(ws._get_value(row, 9))
            mt_meta = ws._get_value(row, 10)
            if zaid is None or mt_meta in (None, ""):
                continue
            try:
                column_name = str(engine._get_cell("Isotopics", row, 12))
            except Exception:
                continue
            if not column_name or column_name == "CLB":
                continue
            lookup[(zaid, str(mt_meta))] = column_name
        return lookup

    def _build_sensitivity_plot_payload(self) -> Dict[str, Any]:
        """Assemble the heatmap payload from the current run plus live filters."""
        if self.current_engine is None or self.current_results is None:
            return {"energies": [], "spectrum": [], "reactions": [], "note": "Run a calculation to populate the plot."}

        # The black spectrum line shows the bounded incident spectrum before any
        # element-specific cover profile is applied.
        energies: List[float] = []
        spectrum: List[float] = []
        for row in range(3, 1003):
            energy = self._float_or_none(self.current_engine._get_cell("Choose Irradiation Parameters", row, 9))
            value = self._float_or_none(self.current_engine._get_cell("Choose Irradiation Parameters", row, 10))
            if energy is None or value is None or energy <= 0.0:
                continue
            energies.append(energy)
            spectrum.append(max(value, 0.0))

        active_group_runs = [
            group_run
            for group_run in self.current_group_runs
            if group_run.get("rows")
        ]
        active_input_count = len(
            {
                str(entry["row"])
                for group_run in active_group_runs
                for entry in group_run["rows"]
            }
        )
        if active_input_count == 0:
            return {
                "energies": energies,
                "spectrum": spectrum,
                "reactions": [],
                "note": "No reaction sensitivities to display. Enter one or more element masses above 0 and run the calculation.",
            }

        isotope_rows, _reference_rows = self._get_filtered_isotope_results()
        included_isotope_zaids = {
            self._int_or_none(row[1])
            for row in isotope_rows
        }
        included_isotope_zaids.discard(None)
        isotope_activity_by_zaid = {
            self._int_or_none(row[1]): self._float_or_none(row[7]) or 0.0
            for row in isotope_rows
            if self._int_or_none(row[1]) is not None
        }
        if not included_isotope_zaids:
            return {
                "energies": energies,
                "spectrum": spectrum,
                "reactions": [],
                "note": "No reaction sensitivities to display. No isotopes met the isotope-results thresholds.",
            }

        isotope_names = self._build_isotope_name_map(self.current_engine)
        reactions: List[Dict[str, Any]] = []
        for group_run in active_group_runs:
            group_engine = group_run["engine"]
            group_atomic_numbers = {
                int(value)
                for value in group_run.get("atomic_numbers", ())
            }
            if not group_atomic_numbers:
                continue

            other_lookup = self._build_other_mt_lookup(group_engine)
            production_ws = group_engine.wb_formula["Production"]
            other_mt_ws = group_engine.wb_formula["Other MT"]
            group_spectrum: List[float] = []
            for spectrum_row in range(3, 1003):
                value = self._float_or_none(group_engine._get_cell("Choose Irradiation Parameters", spectrum_row, 10))
                group_spectrum.append(0.0 if value is None else max(value, 0.0))

            seen = set()
            for row in range(2, production_ws.max_row + 1):
                parent_zaid = self._int_or_none(production_ws._get_value(row, 2))
                mt_value = production_ws._get_value(row, 4)
                reaction = production_ws._get_value(row, 5)
                daughter_zaid = self._int_or_none(production_ws._get_value(row, 10))
                if parent_zaid is None or mt_value in (None, "") or reaction in (None, ""):
                    continue
                if parent_zaid // 1000 not in group_atomic_numbers:
                    continue
                if daughter_zaid not in included_isotope_zaids:
                    continue

                mt_meta = self._format_mt_meta(mt_value, daughter_zaid)
                column_name = other_lookup.get((parent_zaid, mt_meta))
                if not column_name:
                    continue

                key = (group_run["cover_display"], parent_zaid, mt_meta, column_name)
                if key in seen:
                    continue
                seen.add(key)

                column_index = column_index_from_string(column_name)
                values = []
                for mt_row in range(5, 1005):
                    raw_value = self._float_or_none(other_mt_ws._get_value(mt_row, column_index))
                    values.append(0.0 if raw_value is None else max(raw_value, 0.0))
                if max(values, default=0.0) <= 0.0:
                    continue

                weighted_xs = self._float_or_none(group_engine._get_cell("Other MT", 1, column_index)) or 0.0
                parent_name = isotope_names.get(parent_zaid, str(parent_zaid))
                daughter_name = (
                    isotope_names.get(daughter_zaid, str(daughter_zaid))
                    if daughter_zaid not in (None, 0)
                    else ""
                )
                label = f"{parent_name} {reaction}"
                if daughter_name:
                    label = f"{label} -> {daughter_name}"
                if group_run["cover_display"] != "Bare":
                    label = f"{label} [{group_run['cover_display']}]"

                reactions.append(
                    {
                        "label": label,
                        "parent_zaid": parent_zaid,
                        "daughter_zaid": daughter_zaid,
                        "mt_meta": mt_meta,
                        "weighted_xs": weighted_xs,
                        "product_activity_ci": isotope_activity_by_zaid.get(daughter_zaid, 0.0),
                        "cross_section_values": values,
                        "spectrum_values": group_spectrum,
                    }
                )

        # Split each isotope's activity across its candidate production paths
        # using the cover-specific spectrum-weighted cross sections as a proxy.
        grouped: Dict[int | None, List[Dict[str, Any]]] = {}
        for reaction in reactions:
            grouped.setdefault(reaction["daughter_zaid"], []).append(reaction)
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
                        item["spectrum_values"],
                        item["cross_section_values"],
                    )
                ]
                peak_value = max(rate_values, default=0.0)
                intensity_sum = sum(rate_values)
                centroid = (
                    sum(energy * value for energy, value in zip(energies, rate_values)) / intensity_sum
                    if intensity_sum > 0.0
                    else 0.0
                )
                item["activity_share"] = share
                item["estimated_activity_ci"] = item["product_activity_ci"] * share
                item["total_reaction_rate"] = intensity_sum
                item["rate_values"] = rate_values
                item["normalized_values"] = [
                    (value / peak_value) if peak_value > 0.0 else 0.0
                    for value in rate_values
                ]
                item["centroid_energy"] = centroid

        # The optional "major paths only" toggle is purely a plotting filter;
        # isotope tables still keep the full filtered isotope inventory.
        total_reaction_count = len(reactions)
        hidden_minor_paths = 0
        if self.sensitivity_major_paths_only_var.get():
            filtered_reactions: List[Dict[str, Any]] = []
            for item in reactions:
                minimum_activity = max(item["product_activity_ci"], 0.0) * PLOT_REACTION_ACTIVITY_SHARE_THRESHOLD
                if item["estimated_activity_ci"] + 1.0e-30 < minimum_activity:
                    hidden_minor_paths += 1
                    continue
                filtered_reactions.append(item)
            reactions = filtered_reactions
            if not reactions:
                return {
                    "energies": energies,
                    "spectrum": spectrum,
                    "reactions": [],
                    "note": (
                        "No reaction sensitivities to display. "
                        f"All {total_reaction_count} matching production paths fell below the 1% isotope-activity cutoff."
                    ),
                }

        reactions.sort(
            key=lambda item: (
                -item["estimated_activity_ci"],
                -item["product_activity_ci"],
                -item["weighted_xs"],
                item["label"],
            )
        )
        note = (
            f"{len(reactions)} heat-map bands matched isotopes currently shown in the isotope results, "
            f"from {active_input_count} nonzero element inputs across {len(active_group_runs)} cover configuration(s). "
            "Color shows the relative spectrum-weighted reaction rate."
        )
        if any(group_run["cover_display"] != "Bare" for group_run in active_group_runs):
            note = f"{note} The black line shows the bounded incident spectrum before element-specific covers."
        if self.sensitivity_major_paths_only_var.get():
            note = (
                f"{note} Hidden {hidden_minor_paths} path(s) below "
                f"{PLOT_REACTION_ACTIVITY_SHARE_THRESHOLD * 100:.0f}% of their product isotope activity."
            )
        return {
            "energies": energies,
            "spectrum": spectrum,
            "reactions": reactions,
            "note": note,
        }

    def _build_linear_energy_edges(self, energies: List[float]) -> List[float]:
        """Convert sampled energy centers into linear bin edges for pcolormesh."""
        return build_linear_energy_edges(energies)

    def _build_sensitivity_heatmap_cmap(self) -> LinearSegmentedColormap:
        """Blend the heatmap's zero level into a light grey background."""
        base_cmap = matplotlib.colormaps["jet"]
        zero_rgba = matplotlib.colors.to_rgba(HEATMAP_ZERO_COLOR)
        blue_rgba = base_cmap(0.0)
        low_end_samples = [0.0, 0.2, 0.4, 0.6, 0.8, 1.0]
        control_points: List[tuple[float, Any]] = []
        for fraction in low_end_samples:
            interpolated = tuple(
                zero_component + (blue_component - zero_component) * fraction
                for zero_component, blue_component in zip(zero_rgba, blue_rgba)
            )
            control_points.append((HEATMAP_ZERO_TRANSITION * fraction, interpolated))
        for index in range(1, 256):
            normalized = index / 255.0
            mapped_position = HEATMAP_ZERO_TRANSITION + normalized * (1.0 - HEATMAP_ZERO_TRANSITION)
            control_points.append((mapped_position, base_cmap(normalized)))
        return LinearSegmentedColormap.from_list(
            "crpyt_sensitivity_heatmap",
            control_points,
            N=256,
        )

    def _current_spectrum_options(self) -> tuple[str, ...]:
        """Return built-in plus imported spectra for combined-spectrum rows."""
        options: List[str] = []
        seen = set()
        for option in self.base_spectrum_options:
            text = str(option).strip()
            if text and text not in seen:
                seen.add(text)
                options.append(text)
        for name in self.imported_spectra:
            if name not in seen:
                seen.add(name)
                options.append(name)
        return tuple(options or (DEFAULT_COMBINED_SPECTRUM_NAME,))

    def _current_reactor_options(self) -> tuple[str, ...]:
        """Return reactor choices with imported spectra before the import action."""
        options: List[str] = []
        seen = set()
        include_import_action = False
        for option in self.base_reactor_options:
            text = str(option).strip()
            if not text:
                continue
            if text == IMPORT_SPECTRUM_UI_LABEL:
                include_import_action = True
                continue
            if text not in seen:
                seen.add(text)
                options.append(text)
        for name in self.imported_spectra:
            if name not in seen:
                seen.add(name)
                options.append(name)
        if include_import_action and IMPORT_SPECTRUM_UI_LABEL not in seen:
            options.append(IMPORT_SPECTRUM_UI_LABEL)
        return tuple(options)

    def _sync_imported_spectrum_options(self) -> None:
        """Refresh reactor and combined-spectrum dropdowns after imports change."""
        reactor_options = self._current_reactor_options()
        spectrum_options = self._current_spectrum_options()
        self.gui_config["reactor_options"] = list(reactor_options)
        self.gui_config["spectrum_options"] = list(spectrum_options)

        if self.reactor_combobox is not None:
            self.reactor_combobox.configure(values=reactor_options)
        for combobox in self.spectrum_comboboxes:
            combobox.configure(values=spectrum_options)

        if self.reactor_var.get() not in reactor_options:
            self.reactor_var.set(self.gui_config["default_reactor"])
        for variable in self.spectrum_name_vars.values():
            if variable.get() not in spectrum_options:
                variable.set(spectrum_options[0])

    def _unique_imported_spectrum_name(self, raw_name: str, used_names: set[str]) -> str:
        """Preserve imported headers while avoiding ambiguous dropdown names."""
        base_name = str(raw_name).strip() or "Imported Spectrum"
        candidate = base_name
        if candidate in used_names:
            candidate = f"{base_name} (imported)"
        suffix = 2
        while candidate in used_names:
            candidate = f"{base_name} (imported {suffix})"
            suffix += 1
        used_names.add(candidate)
        return candidate

    def _selected_imported_spectrum(self, reactor_selection: str | None = None) -> ImportedSpectrum | None:
        """Return the imported spectrum selected as the reactor, if any."""
        selected = self.reactor_var.get() if reactor_selection is None else reactor_selection
        if selected in self.imported_spectra:
            return self.imported_spectra[selected]
        if selected == IMPORT_SPECTRUM_UI_LABEL and self.imported_spectra:
            return next(iter(self.imported_spectra.values()))
        return None

    def _using_imported_spectrum(self) -> bool:
        """Return True when the current reactor selection uses imported values."""
        return self._selected_imported_spectrum() is not None or self.reactor_var.get() == IMPORT_SPECTRUM_UI_LABEL

    def _handle_reactor_selection(self) -> None:
        """React to reactor dropdown changes, including the import action."""
        selected = self.reactor_var.get()
        if selected != IMPORT_SPECTRUM_UI_LABEL:
            if selected:
                self._last_non_import_reactor = selected
            self._update_spectrum_visibility()
            return

        self._update_spectrum_visibility()
        if self._reactor_import_prompt_pending:
            return
        self._reactor_import_prompt_pending = True
        self.root.after_idle(self._prompt_for_imported_spectrum)

    def _prompt_for_imported_spectrum(self) -> None:
        """Open a spectrum file picker and load user-provided reactor spectra."""
        self._reactor_import_prompt_pending = False
        path_text = filedialog.askopenfilename(
            title="Import Reactor Spectrum",
            filetypes=(
                ("Spectrum files", "*.csv *.xlsx *.xlsm"),
                ("CSV files", "*.csv"),
                ("Excel workbooks", "*.xlsx *.xlsm"),
                ("All files", "*.*"),
            ),
        )
        if not path_text:
            self.reactor_var.set(self._last_non_import_reactor)
            return

        path = Path(path_text)
        try:
            imported_names = self._load_imported_spectrum_file(path)
        except Exception as exc:
            messagebox.showerror("Import Spectrum", str(exc))
            self.reactor_var.set(self._last_non_import_reactor)
            return

        if imported_names:
            self.reactor_var.set(imported_names[0])
            self._last_non_import_reactor = imported_names[0]
        self.status_var.set(f"Imported {len(imported_names)} spectrum(s): {path.name}")

    def _load_imported_spectrum_file(self, path: Path) -> List[str]:
        """Parse and conservatively rebin imported spectra onto CRPYT's grid."""
        series = self._read_imported_spectrum_file(path)
        target_energies = self._target_spectrum_energies()
        used_names = set(self.base_reactor_options) | set(self.base_spectrum_options)
        used_names.update({COMBINED_REACTOR_UI_LABEL, IMPORT_SPECTRUM_UI_LABEL})

        loaded_spectra: Dict[str, ImportedSpectrum] = {}
        for raw_name, points in series:
            rebinned_values, source_integral, target_integral = self._rebase_spectrum_to_target_grid(
                points,
                target_energies,
            )

            if abs(source_integral - target_integral) > max(1.0e-12, abs(source_integral) * 1.0e-9):
                raise ValueError(
                    f"Imported spectrum '{raw_name}' failed its conservation check: "
                    f"source area {source_integral:.12e}, target area {target_integral:.12e}."
                )

            name = self._unique_imported_spectrum_name(raw_name, used_names)
            loaded_spectra[name] = ImportedSpectrum(
                name=name,
                path=path,
                values=tuple(rebinned_values),
                source_points=len(points),
                integral=target_integral,
            )

        if not loaded_spectra:
            raise ValueError("Imported spectrum file did not contain any spectra.")

        self.imported_spectra = loaded_spectra
        first_spectrum = next(iter(loaded_spectra.values()))
        self.imported_spectrum_path = path
        self.imported_spectrum_values = list(first_spectrum.values)
        self.imported_spectrum_source_points = first_spectrum.source_points
        self.imported_spectrum_integral = first_spectrum.integral
        self._sync_imported_spectrum_options()

        names = list(loaded_spectra)
        displayed_names = ", ".join(names[:3])
        if len(names) > 3:
            displayed_names = f"{displayed_names}, +{len(names) - 3} more"
        self.imported_spectrum_summary_var.set(
            f"Imported {len(names)} spectrum(s) from {path.name}: {displayed_names} "
            f"({first_spectrum.source_points} pts -> {len(first_spectrum.values)} bins)"
        )
        return names

    def _read_imported_spectrum_file(self, path: Path) -> List[tuple[str, List[tuple[float, float]]]]:
        """Read CSV/XLSX spectra where column 1 is energy and later columns are spectra."""
        return read_imported_spectrum_file(path)

    def _read_imported_spectrum_csv(self, path: Path) -> List[tuple[str, List[tuple[float, float]]]]:
        """Read a CSV spectrum table, accepting one optional header row."""
        return read_imported_spectrum_csv(path)

    def _read_imported_spectrum_workbook(self, path: Path) -> List[tuple[str, List[tuple[float, float]]]]:
        """Read the first sheet of an Excel spectrum workbook."""
        return read_imported_spectrum_workbook(path)

    def _target_spectrum_energies(self) -> List[float]:
        """Return the CRPYT reactor spectrum energy grid used by the runtime."""
        if self._spectrum_energy_grid is not None:
            return list(self._spectrum_energy_grid)

        engine = self.engine_factory.create_engine()
        energies: List[float] = []
        for row in range(3, 1003):
            energy = self._float_or_none(engine._get_cell("Choose Irradiation Parameters", row, 9))
            if energy is None or energy <= 0.0:
                continue
            energies.append(energy)

        if len(energies) < 2:
            raise ValueError("CRPYT spectrum grid could not be loaded from the bundled reactor spectra.")
        for left, right in zip(energies[:-1], energies[1:]):
            if right <= left:
                raise ValueError("CRPYT spectrum grid energies must be strictly increasing.")

        self._spectrum_energy_grid = list(energies)
        return energies

    def _rebase_spectrum_to_target_grid(
        self,
        points: Sequence[tuple[float, float]],
        target_energies: Sequence[float],
    ) -> tuple[List[float], float, float]:
        """Conservatively rebin imported spectrum weights onto CRPYT energy bins."""
        return rebase_spectrum_to_target_grid(points, target_energies)

    def _builtin_spectrum_values_by_name(self) -> Dict[str, Tuple[float, ...]]:
        """Load built-in reactor spectrum columns keyed by workbook display name."""
        if self._builtin_spectrum_values_cache is not None:
            return dict(self._builtin_spectrum_values_cache)

        engine = self.engine_factory.create_engine()
        sheet_name = "Choose Irradiation Parameters"
        target_energies = self._target_spectrum_energies()
        values_by_name: Dict[str, Tuple[float, ...]] = {}

        lookup_name_column = column_index_from_string("CL")
        lookup_letter_column = column_index_from_string("CM")
        worksheet = engine.wb_formula[sheet_name]
        for row in range(3, worksheet.max_row + 1):
            raw_name = engine._get_cell(sheet_name, row, lookup_name_column)
            raw_column = engine._get_cell(sheet_name, row, lookup_letter_column)
            if raw_name in (None, "") or raw_column in (None, ""):
                continue
            name = str(raw_name).strip()
            if not name or name == "Spectrum" or name in values_by_name:
                continue
            try:
                value_column = column_index_from_string(str(raw_column).strip())
            except Exception:
                continue
            values: List[float] = []
            for spectrum_row in range(3, 1003):
                value = self._float_or_none(engine._get_cell(sheet_name, spectrum_row, value_column))
                values.append(0.0 if value is None else max(value, 0.0))
            if len(values) == len(target_energies):
                values_by_name[name] = tuple(values)

        self._builtin_spectrum_values_cache = dict(values_by_name)
        return values_by_name

    def _spectrum_values_for_name(self, name: str) -> Tuple[float, ...]:
        """Resolve a combined-spectrum row name to raw unattenuated bin values."""
        if name in self.imported_spectra:
            return self.imported_spectra[name].values
        builtin_values = self._builtin_spectrum_values_by_name()
        if name in builtin_values:
            return builtin_values[name]
        raise ValueError(f"Unknown spectrum selected for Combined Spectrum: {name}")

    def _combine_spectrum_entries(self, entries: Sequence[SpectrumMixEntry]) -> Tuple[float, ...]:
        """Build a raw combined spectrum from built-in and imported entries."""
        target_count = len(self._target_spectrum_energies())
        combined = [0.0] * target_count
        for entry in entries:
            name = entry.name.strip()
            fraction = float(entry.fraction)
            if not name or fraction == 0.0:
                continue
            values = self._spectrum_values_for_name(name)
            if len(values) != target_count:
                raise ValueError(f"Spectrum '{name}' does not match the CRPYT reactor spectrum grid.")
            for index, value in enumerate(values):
                combined[index] += max(float(value), 0.0) * fraction
        return tuple(max(value, 0.0) for value in combined)

    def _runtime_spectrum_values(
        self,
        reactor_selection: str,
        combined_spectra: Sequence[SpectrumMixEntry],
    ) -> Tuple[float, ...]:
        """Return J-column source values when the workbook lookup should be bypassed."""
        if reactor_selection == COMBINED_REACTOR_UI_LABEL:
            return self._combine_spectrum_entries(combined_spectra)
        imported_spectrum = self._selected_imported_spectrum(reactor_selection)
        if imported_spectrum is not None:
            return imported_spectrum.values
        if reactor_selection == IMPORT_SPECTRUM_UI_LABEL:
            raise ValueError("Choose a spectrum file before running with Import Spectrum selected.")
        return ()

    def _split_cell_ref(self, cell_ref: str) -> tuple[int, int]:
        """Convert an A1-style cell reference into row and column numbers."""
        column_text = "".join(ch for ch in cell_ref if ch.isalpha())
        row_text = "".join(ch for ch in cell_ref if ch.isdigit())
        if not column_text or not row_text:
            raise ValueError(f"Invalid workbook cell reference: {cell_ref}")
        return int(row_text), column_index_from_string(column_text)

    def _cover_attenuation_for_spectrum_row(self, engine: Any, spectrum_row: int) -> float:
        """Calculate the cover attenuation factor used by built-in spectra."""
        attenuation_exponent = 0.0
        cover_xs_row = spectrum_row + 4
        for column_offset, entry in enumerate(self.gui_config["cover_materials"], start=2):
            cell_row, cell_column = self._split_cell_ref(str(entry["cell"]))
            thickness = self._float_or_none(
                engine._get_cell("Choose Irradiation Parameters", cell_row, cell_column)
            ) or 0.0
            if thickness <= 0.0:
                continue
            coefficient = self._float_or_none(engine._get_cell("Cover XS", cover_xs_row, column_offset)) or 0.0
            if coefficient > 0.0:
                attenuation_exponent += thickness * coefficient / 10.0

        if attenuation_exponent <= 0.0:
            return 1.0
        if attenuation_exponent > 745.0:
            return 0.0
        return math.exp(-attenuation_exponent)

    def _build_imported_spectrum_overrides(
        self,
        engine: Any,
        imported_spectrum_values: Sequence[float],
    ) -> Dict[str, float]:
        """Build J-column overrides for an imported spectrum on this engine."""
        if not imported_spectrum_values:
            return {}
        target_energies = self._target_spectrum_energies()
        if len(imported_spectrum_values) != len(target_energies):
            raise ValueError("Imported spectrum does not match the CRPYT reactor spectrum grid.")

        overrides: Dict[str, float] = {}
        for row, value in zip(range(3, 1003), imported_spectrum_values):
            attenuation = self._cover_attenuation_for_spectrum_row(engine, row)
            overrides[f"Choose Irradiation Parameters!J{row}"] = max(float(value) * attenuation, 0.0)
        return overrides

    def _update_spectrum_visibility(self) -> None:
        """Show combined-spectrum inputs only when the Combo reactor is selected."""
        panes = set(str(pane) for pane in self.lower_paned.panes())
        frame_path = str(self.spectrum_frame)
        if self.reactor_var.get() == COMBINED_REACTOR_UI_LABEL:
            if frame_path not in panes:
                self.lower_paned.add(self.spectrum_frame, weight=1)
        else:
            if frame_path in panes:
                self.lower_paned.forget(self.spectrum_frame)

    def _update_derived_metrics(self) -> None:
        """Refresh the lightweight derived metrics shown beside the inputs."""
        try:
            fluence = self._parse_float(self.fluence_var.get(), "Neutron Fluence Rate")
            irradiation = self._parse_float(self.irradiation_var.get(), "Irradiation Time")
            decay = self._parse_float(self.decay_var.get(), "Decay Time")
            total_mass = sum(self._parse_float(var.get(), "Mass") for var in self.mass_vars.values())
        except ValueError:
            self.irradiation_days_var.set("--")
            self.decay_days_var.set("--")
            self.total_fluence_var.set("--")
            self.total_mass_var.set("--")
            return

        self.irradiation_days_var.set(self._format_value(irradiation / 86400.0))
        self.decay_days_var.set(self._format_value(decay / 86400.0))
        self.total_fluence_var.set(self._format_value(fluence * irradiation))
        self.total_mass_var.set(self._format_value(total_mass))

    def _default_spectrum_summary_text(self) -> str:
        """Return the placeholder text used before GREASPY results exist."""
        if self.spectrum_model_error:
            return f"GREASPY spectrum unavailable: {self.spectrum_model_error}"
        return "Run a calculation to populate the GREASPY detector spectrum."

    def _collect_spectrum_settings(self) -> tuple[float, float, bool, Dict[str, float]]:
        """Parse the GREASPY-specific detector settings from the Spectrum tab."""
        count_time_sec = self._parse_float(self.spectrum_count_time_var.get(), "Counting Time")
        distance_cm = self._parse_float(self.spectrum_distance_var.get(), "Detector Distance")
        if count_time_sec <= 0.0:
            raise ValueError("Counting Time must be greater than 0.")
        if distance_cm <= 0.0:
            raise ValueError("Detector Distance must be greater than 0.")

        shielding_cm: Dict[str, float] = {}
        for material in SHIELDING_MATERIALS:
            value = self._parse_float(self.spectrum_shield_vars[material].get(), f"{material} shielding")
            if value < 0.0:
                raise ValueError(f"{material} shielding must be greater than or equal to 0.")
            shielding_cm[material] = value

        return count_time_sec, distance_cm, self.spectrum_include_bg_var.get(), shielding_cm

    def _build_spectrum_inputs_from_results(
        self,
        results: Dict[str, Any],
        count_time_sec: float,
        distance_cm: float,
        include_background: bool,
        shielding_cm: Dict[str, float],
    ) -> tuple[CrpytSpectrumInputs, Dict[int, str], int]:
        """Translate CRPYT isotope results into GREASPY activity inputs."""
        if self.spectrum_model is None:
            raise ValueError(self._default_spectrum_summary_text())

        activity_by_zaid_ci: Dict[int, float] = {}
        isotope_names_by_zaid: Dict[int, str] = {}
        for row in results["isotopes_main"]:
            zaid = self._int_or_none(row[1])
            activity_ci = self._float_or_none(row[7]) or 0.0
            if zaid is None or activity_ci <= 0.0:
                continue
            activity_by_zaid_ci[zaid] = activity_by_zaid_ci.get(zaid, 0.0) + activity_ci
            isotope_names_by_zaid.setdefault(zaid, str(row[2]))

        if not activity_by_zaid_ci:
            raise ValueError("No isotope activities are available for GREASPY spectrum generation.")

        return (
            CrpytSpectrumInputs(
                activity_by_zaid_ci=activity_by_zaid_ci,
                count_time_sec=count_time_sec,
                distance_cm=distance_cm,
                include_background=include_background,
                shielding_cm=shielding_cm,
            ),
            isotope_names_by_zaid,
            len(activity_by_zaid_ci),
        )

    def _collect_spectrum_inputs(self) -> tuple[CrpytSpectrumInputs, Dict[int, str], int]:
        """Build GREASPY inputs from the most recent CRPYT calculation."""
        if self.current_results is None:
            raise ValueError("Run a calculation to populate the GREASPY detector spectrum.")
        settings = self._collect_spectrum_settings()
        return self._build_spectrum_inputs_from_results(self.current_results, *settings)

    def _build_spectrum_summary(
        self,
        spectrum_inputs: CrpytSpectrumInputs,
        spectrum_results: GreasyCalcResults,
        total_active_isotopes: int,
    ) -> str:
        """Summarize the GREASPY run in a single status line for the Spectrum tab."""
        total_counts = sum(spectrum_results.total_counts)
        total_shield = sum(max(0.0, value) for value in spectrum_inputs.shielding_cm.values())
        lines = [
            f"Matched isotopes in GREASPY: {spectrum_results.matched_isotope_count} of {total_active_isotopes}",
            f"Total counts: {self._format_value(total_counts)}",
            f"Background included: {'Yes' if spectrum_inputs.include_background else 'No'}",
            f"Unshielded dead time (%): {self._format_value(spectrum_results.unshielded_dead_time_pct)}",
            f"Real counting time (sec): {self._format_value(spectrum_results.unshielded_real_time_sec)}",
            f"Dose at 30 cm (R/hr): {self._format_value(spectrum_results.unshielded_dose_30cm)}",
        ]
        if total_shield > 0.0 and spectrum_results.shielded_dead_time_pct is not None:
            lines.extend(
                [
                    f"Shielded dead time (%): {self._format_value(spectrum_results.shielded_dead_time_pct)}",
                    f"Shielded dose at 30 cm (R/hr): {self._format_value(spectrum_results.shielded_dose_30cm)}",
                ]
            )
        if spectrum_results.missing_zaids:
            lines.append(f"Missing library spectra: {len(spectrum_results.missing_zaids)} isotope(s)")
        return " | ".join(lines)

    def _refresh_gamma_spectrum(self) -> None:
        """Recompute the GREASPY detector spectrum when its inputs change."""
        if self.gamma_axes is None or self.gamma_canvas is None:
            return
        if self.current_results is None:
            self.current_spectrum_results = None
            self.spectrum_summary_var.set(self._default_spectrum_summary_text())
            self._draw_gamma_plot()
            return
        try:
            spectrum_inputs, isotope_names_by_zaid, total_active_isotopes = self._collect_spectrum_inputs()
            if self.spectrum_model is None:
                raise ValueError(self._default_spectrum_summary_text())
            self.current_spectrum_results = self.spectrum_model.compute_from_activities(
                spectrum_inputs,
                isotope_names_by_zaid,
            )
            self.spectrum_summary_var.set(
                self._build_spectrum_summary(
                    spectrum_inputs,
                    self.current_spectrum_results,
                    total_active_isotopes,
                )
            )
        except ValueError as exc:
            self.current_spectrum_results = None
            self.spectrum_summary_var.set(str(exc))
        except Exception as exc:
            self.current_spectrum_results = None
            self.spectrum_summary_var.set(f"GREASPY spectrum failed: {exc}")
        self._draw_gamma_plot()

    def _build_spectrum_export_inputs(
        self,
        spectrum_inputs: CrpytSpectrumInputs,
        isotope_names_by_zaid: Dict[int, str],
    ) -> GreaspyCalcInputs:
        """Convert CRPYT activity-by-ZAID inputs into the shared exporter shape."""
        isotopes = [
            (isotope_names_by_zaid.get(int(zaid), str(int(zaid))), float(activity_ci))
            for zaid, activity_ci in sorted(spectrum_inputs.activity_by_zaid_ci.items())
            if float(activity_ci) > 0.0
        ]
        return GreaspyCalcInputs(
            isotopes=isotopes,
            count_time_sec=spectrum_inputs.count_time_sec,
            distance_cm=spectrum_inputs.distance_cm,
            include_background=spectrum_inputs.include_background,
            shielding_cm=spectrum_inputs.shielding_cm,
        )

    def _export_detector_spectrum(self) -> None:
        """Export the current GREASPY detector spectrum in a supported file format."""
        if self.current_results is None:
            messagebox.showinfo("No Spectrum", "Run a calculation before exporting a detector spectrum.")
            return

        try:
            spectrum_inputs, isotope_names_by_zaid, _total_active_isotopes = self._collect_spectrum_inputs()
            if self.spectrum_model is None:
                raise ValueError(self._default_spectrum_summary_text())
            if self.current_spectrum_results is None:
                self.current_spectrum_results = self.spectrum_model.compute_from_activities(
                    spectrum_inputs,
                    isotope_names_by_zaid,
                )
            export_inputs = self._build_spectrum_export_inputs(spectrum_inputs, isotope_names_by_zaid)
        except Exception as exc:
            messagebox.showerror("Export Spectrum", str(exc))
            return

        save_path = filedialog.asksaveasfilename(
            title="Export Detector Spectrum",
            defaultextension=".spe",
            filetypes=[
                ("Spectrum files", "*.spe *.chn *.n42 *.spc"),
                ("SPE file", "*.spe"),
                ("ORTEC CHN file", "*.chn"),
                ("ANSI N42 file", "*.n42"),
                ("ORTEC SPC file", "*.spc"),
            ],
        )
        if not save_path:
            return

        try:
            save_file = Path(save_path)
            spectrum_format = infer_spectrum_format(save_file)
            write_spectrum_file(
                save_file,
                export_inputs,
                self.current_spectrum_results,
                spectrum_format,
                detector_description="CRPYT Detector Spectrum IDM-200-V",
                analysis_program="Generated by CRPYT",
                shielding_materials=SHIELDING_MATERIALS,
            )
        except Exception as exc:
            messagebox.showerror("Export Spectrum", str(exc))
            return

        messagebox.showinfo("Export Spectrum", f"Saved {spectrum_format.upper()}: {save_path}")

    def _get_filtered_isotope_results(self) -> tuple[List[List[Any]], List[List[Any]]]:
        """Apply the live isotope display thresholds to the cached run results."""
        if self.current_results is None:
            return [], []
        min_atoms = self._parse_float(
            self.isotopes_threshold_var.get(),
            "Minimum Atoms of Isotope Provided",
        )
        min_activity_ci = self._parse_float(
            self.isotopes_activity_threshold_var.get(),
            "Minimum Activity",
        )
        min_gamma_rate = self._parse_float(
            self.gamma_threshold_var.get(),
            "Minimum Gamma Rate",
        )
        filtered_main: List[List[Any]] = []
        filtered_reference: List[List[Any]] = []
        for main_row, reference_row in zip(
            self.current_results.get("isotopes_main", []),
            self.current_results.get("isotopes_reference", []),
        ):
            if (
                self._float_or_none(main_row[4]) or 0.0
            ) <= min_atoms:
                continue
            if (
                self._float_or_none(main_row[7]) or 0.0
            ) <= min_activity_ci:
                continue
            if (
                self._float_or_none(main_row[8]) or 0.0
            ) <= min_gamma_rate:
                continue
            filtered_main.append(main_row)
            filtered_reference.append(reference_row)
        return filtered_main, filtered_reference

    def _get_filtered_element_results(self) -> List[List[Any]]:
        """Apply the live element mass threshold to the cached element rows."""
        if self.current_results is None:
            return []
        min_mass = self._parse_float(
            self.elements_threshold_var.get(),
            "Minimum Mass at EOID",
        )
        return [
            row
            for row in self.current_results.get("elements_all", [])
            if (self._float_or_none(row[2]) or 0.0) > min_mass
        ]

    def _refresh_filtered_views(self) -> None:
        """Redraw only the views that depend on live UI thresholds and toggles."""
        if self.current_results is None:
            return
        try:
            element_rows = self._get_filtered_element_results()
            isotope_rows, reference_rows = self._get_filtered_isotope_results()
        except ValueError:
            return

        self.current_plot_payload = self._build_sensitivity_plot_payload()
        self._fill_tree(self.elements_tree, ELEMENT_HEADERS, element_rows)
        self._fill_tree(self.isotopes_tree, ISOTOPE_HEADERS, isotope_rows)
        self._fill_tree(self.isotopes_ref_tree, ISOTOPE_REFERENCE_HEADERS, reference_rows)
        self._draw_elements_plot()
        self._draw_isotopes_plot()
        self._draw_sensitivity_plot()

    def _parse_float(self, text: str, field_name: str) -> float:
        """Parse a numeric entry field using workbook-friendly empty-as-zero rules."""
        cleaned = text.strip()
        if not cleaned:
            return 0.0
        try:
            return float(cleaned)
        except ValueError as exc:
            raise ValueError(f"{field_name} must be numeric.") from exc

    def _parse_int(self, text: str, field_name: str) -> int:
        """Parse an integer entry while rejecting non-integral numeric input."""
        cleaned = text.strip()
        if not cleaned:
            return 0
        try:
            return int(cleaned)
        except ValueError:
            numeric = self._parse_float(cleaned, field_name)
            if not float(numeric).is_integer():
                raise ValueError(f"{field_name} must be an integer.")
            return int(numeric)

    def _parse_optional_float(self, text: str, field_name: str) -> float | None:
        """Parse an optional float entry, returning None when left blank."""
        cleaned = text.strip()
        if not cleaned:
            return None
        return self._parse_float(cleaned, field_name)

    def _parse_element_list(self, text: str) -> tuple[str, ...]:
        """Parse comma-separated element names without splitting internal spaces."""
        parts = []
        for raw in text.replace(";", ",").split(","):
            cleaned = raw.strip()
            if cleaned:
                parts.append(cleaned)
        return tuple(parts)

    def _clear_optimization_inclusions(self) -> None:
        """Clear all include-element checkbox selections from the dropdown."""
        for variable in self.optimization_include_vars.values():
            variable.set(False)

    def _clear_optimization_exclusions(self) -> None:
        """Clear all exclude-element checkbox selections from the dropdown."""
        for variable in self.optimization_exclude_vars.values():
            variable.set(False)

    def _clear_optimization_cover_scenarios(self) -> None:
        """Clear all optimizer cover-scenario selections from the dropdown."""
        for variable in self.optimization_cover_vars.values():
            variable.set(False)

    def _get_optimization_included_elements(self) -> tuple[str, ...]:
        """Merge checked inclusions with any manually typed inclusions."""
        selected = [
            element
            for element, variable in self.optimization_include_vars.items()
            if variable.get()
        ]
        manual = list(self._parse_element_list(self.optimization_include_var.get()))
        merged: List[str] = []
        seen = set()
        for value in selected + manual:
            key = value.strip().upper()
            if key and key not in seen:
                seen.add(key)
                merged.append(value)
        return tuple(merged)

    def _get_optimization_excluded_elements(self) -> tuple[str, ...]:
        """Merge checked exclusions with any manually typed exclusions."""
        selected = [
            element
            for element, variable in self.optimization_exclude_vars.items()
            if variable.get()
        ]
        manual = list(self._parse_element_list(self.optimization_exclude_var.get()))
        merged: List[str] = []
        seen = set()
        for value in selected + manual:
            key = value.strip().upper()
            if key and key not in seen:
                seen.add(key)
                merged.append(value)
        return tuple(merged)

    def _update_optimization_include_summary(self) -> None:
        """Keep the include dropdown label short while preserving blank=all behavior."""
        checked = [
            element
            for element, variable in self.optimization_include_vars.items()
            if variable.get()
        ]
        manual_values = self._parse_element_list(self.optimization_include_var.get())
        manual_only_count = 0
        seen_checked = {value.strip().upper() for value in checked}
        for value in manual_values:
            key = value.strip().upper()
            if key and key not in seen_checked:
                manual_only_count += 1

        if not checked and manual_only_count == 0:
            summary = "All elements"
        elif len(checked) <= 2 and manual_only_count == 0:
            summary = ", ".join(checked)
        else:
            total_count = len(checked) + manual_only_count
            summary = f"{total_count} selected"
        self.optimization_include_menu_summary_var.set(summary)

    def _update_optimization_exclude_summary(self) -> None:
        """Keep the exclude dropdown label short but informative."""
        checked = [
            element
            for element, variable in self.optimization_exclude_vars.items()
            if variable.get()
        ]
        manual_values = self._parse_element_list(self.optimization_exclude_var.get())
        manual_only_count = 0
        seen_checked = {value.strip().upper() for value in checked}
        for value in manual_values:
            key = value.strip().upper()
            if key and key not in seen_checked:
                manual_only_count += 1

        if not checked and manual_only_count == 0:
            summary = "Select elements..."
        elif len(checked) <= 2 and manual_only_count == 0:
            summary = ", ".join(checked)
        else:
            total_count = len(checked) + manual_only_count
            summary = f"{total_count} selected"
        self.optimization_exclude_menu_summary_var.set(summary)

    def _get_optimization_cover_materials(self) -> tuple[str, ...]:
        """Return the checked optimizer cover scenarios in display order."""
        return tuple(
            cover_label
            for cover_label, variable in self.optimization_cover_vars.items()
            if variable.get()
        )

    def _update_optimization_cover_summary(self) -> None:
        """Keep the cover dropdown label compact while still showing selection state."""
        selected = list(self._get_optimization_cover_materials())
        if not selected:
            summary = "No covers"
        elif len(selected) <= 2:
            summary = ", ".join(selected)
        else:
            summary = f"{len(selected)} covers selected"
        self.optimization_cover_menu_summary_var.set(summary)

    def _format_value(self, value: Any) -> str:
        """Format numbers the same way the GUI tables and labels expect them."""
        if value in (None, ""):
            return ""
        if isinstance(value, float):
            abs_value = abs(value)
            if abs_value >= 1.0e4 or (abs_value > 0.0 and abs_value < 1.0e-4):
                return f"{value:.6e}"
            return f"{value:.6f}".rstrip("0").rstrip(".")
        return str(value)

    def _parse_energy_bounds(self) -> tuple[float, float]:
        """Validate the energy window shared by the calculation and plot."""
        min_energy = self._parse_float(self.energy_min_var.get(), "Minimum Energy")
        max_energy = self._parse_float(self.energy_max_var.get(), "Maximum Energy")
        if min_energy < 0.0:
            raise ValueError("Minimum Energy must be greater than or equal to 0.")
        if max_energy <= 0.0:
            raise ValueError("Maximum Energy must be greater than 0.")
        if min_energy >= max_energy:
            raise ValueError("Minimum Energy must be less than Maximum Energy.")
        return min_energy, max_energy

    def _resolve_sensitivity_plot_energy_limits(self, energies: List[float]) -> tuple[float, float] | None:
        """Choose safe x-limits for the plot when the entry fields are invalid."""
        positive_energies = [energy for energy in energies if energy > 0.0]
        if not positive_energies:
            return None
        try:
            min_energy, max_energy = self._parse_energy_bounds()
        except ValueError:
            return min(positive_energies), max(positive_energies)

        lower = min_energy if min_energy > 0.0 else min(positive_energies)
        upper = max_energy
        if upper <= lower:
            return min(positive_energies), max(positive_energies)
        return lower, upper

    def _build_bounded_spectrum_overrides(
        self,
        engine: Any,
        min_energy: float,
        max_energy: float,
    ) -> Dict[str, float]:
        """Zero out spectrum bins outside the requested calculation energy window."""
        overrides: Dict[str, float] = {}
        for row in range(3, 1003):
            energy = self._float_or_none(engine._get_cell("Choose Irradiation Parameters", row, 9))
            if energy is None or energy <= 0.0:
                continue
            spectrum_value = self._float_or_none(engine._get_cell("Choose Irradiation Parameters", row, 10))
            bounded_value = 0.0
            if min_energy <= energy <= max_energy and spectrum_value is not None:
                bounded_value = max(spectrum_value, 0.0)
            overrides[f"Choose Irradiation Parameters!J{row}"] = bounded_value
        return overrides

    def _collect_element_cover_groups(
        self,
        overrides: Dict[str, Any],
        input_specs: List[Dict[str, Any]],
    ) -> tuple[List[Dict[str, Any]], Dict[str, Any] | None, str | None]:
        """Group active elements by their selected cover combination.

        The extracted workbook logic only supports one global cover setup at a
        time. To support element-by-element covers, the runtime groups elements
        that share the same selected cover profile, runs each group separately,
        and adds the results back together.
        """

        cover_thickness_by_cell: Dict[str, float] = {}
        for entry in self.gui_config["cover_materials"]:
            cell = str(entry["cell"])
            ref = f"Choose Irradiation Parameters!{cell}"
            thickness = self._float_or_none(overrides.get(ref)) or 0.0
            if thickness > 0.0:
                cover_thickness_by_cell[cell] = thickness

        grouped_row_masses: Dict[tuple[str, ...], Dict[int, float]] = {}
        grouped_rows: Dict[tuple[str, ...], List[Dict[str, Any]]] = {}
        active_profiles: List[tuple[str, ...]] = []
        for spec in input_specs:
            for cover_profile in spec["run_profiles"]:
                row_masses = grouped_row_masses.setdefault(cover_profile, {})
                for row, mass in spec["row_masses"].items():
                    row_masses[row] = row_masses.get(row, 0.0) + float(mass)
                grouped_rows.setdefault(cover_profile, []).append(
                    {
                        "row": spec["row"],
                        "element": spec["element"],
                        "atomic_numbers": tuple(spec["atomic_numbers"]),
                    }
                )
                active_profiles.append(cover_profile)

        if not grouped_rows:
            grouped_rows[tuple()] = []
            grouped_row_masses[tuple()] = {}

        group_specs: List[Dict[str, Any]] = []
        for cover_profile, rows in sorted(
            grouped_rows.items(),
            key=lambda item: (len(item[0]), self._cover_profile_display(item[0])),
        ):
            atomic_numbers = sorted(
                {
                    int(value)
                    for entry in rows
                    for value in entry.get("atomic_numbers", ())
                }
            )
            group_specs.append(
                {
                    "cover_cells": tuple(cover_profile),
                    "cover_display": self._cover_profile_display(tuple(cover_profile)),
                    "rows": [dict(entry) for entry in rows],
                    "row_masses": dict(grouped_row_masses.get(cover_profile, {})),
                    "atomic_numbers": tuple(atomic_numbers),
                }
            )

        export_overrides: Dict[str, Any] | None = dict(overrides)
        export_error: str | None = None
        unique_profiles = {profile for profile in active_profiles}
        if len(unique_profiles) > 1:
            export_overrides = None
            export_error = (
                "Validation workbook export is unavailable when different elements use "
                "different cover combinations. The workbook export only supports one "
                "global cover configuration."
            )
        else:
            export_profile = next(iter(unique_profiles), tuple())
            export_row_masses = grouped_row_masses.get(export_profile, {})
            for entry in self.gui_config["cover_materials"]:
                cell = str(entry["cell"])
                ref = f"Choose Irradiation Parameters!{cell}"
                export_overrides[ref] = cover_thickness_by_cell.get(cell, 0.0) if cell in export_profile else 0.0
            for entry in self.gui_config["composition_rows"]:
                row = int(entry["row"])
                ref = f"Choose Irradiation Parameters!G{row}"
                export_overrides[ref] = float(export_row_masses.get(row, 0.0))
        return group_specs, export_overrides, export_error

    def _build_group_overrides(
        self,
        base_overrides: Dict[str, Any],
        cover_cells: tuple[str, ...],
        row_masses: Dict[int, float],
    ) -> Dict[str, Any]:
        """Project the shared GUI inputs into one cover-profile engine run."""

        scenario_overrides = dict(base_overrides)
        for entry in self.gui_config["cover_materials"]:
            cell = str(entry["cell"])
            ref = f"Choose Irradiation Parameters!{cell}"
            scenario_overrides[ref] = base_overrides[ref] if cell in cover_cells else 0.0
        for entry in self.gui_config["composition_rows"]:
            row = int(entry["row"])
            ref = f"Choose Irradiation Parameters!G{row}"
            scenario_overrides[ref] = float(row_masses.get(row, 0.0))
        return scenario_overrides

    def _run_grouped_calculation(
        self,
        overrides: Dict[str, Any],
        thresholds: Dict[str, float],
        input_specs: List[Dict[str, Any]],
        imported_spectrum_values: Sequence[float],
    ) -> tuple[Any, Dict[str, Any] | None, str | None, List[Dict[str, Any]], Dict[str, Any]]:
        """Execute one CRPYT run per active cover profile and combine the results."""

        group_specs, export_overrides, export_error = self._collect_element_cover_groups(overrides, input_specs)

        # The base engine keeps the bounded incident spectrum and static lookup
        # sheets used by the plot, independent of any element-specific covers.
        base_engine = self.engine_factory.create_engine()
        base_overrides = dict(overrides)
        for entry in self.gui_config["cover_materials"]:
            base_overrides[f"Choose Irradiation Parameters!{entry['cell']}"] = 0.0
        for entry in self.gui_config["composition_rows"]:
            base_overrides[f"Choose Irradiation Parameters!G{entry['row']}"] = 0.0
        for ref, value in base_overrides.items():
            base_engine.set_cell(ref, value)
        for ref, value in self._build_imported_spectrum_overrides(base_engine, imported_spectrum_values).items():
            base_engine.set_cell(ref, value)
        bounded_base_overrides = self._build_bounded_spectrum_overrides(
            base_engine,
            thresholds["energy_min_mev"],
            thresholds["energy_max_mev"],
        )
        for ref, value in bounded_base_overrides.items():
            base_engine.set_cell(ref, value)

        grouped_payloads: List[Dict[str, Any]] = []
        group_runs: List[Dict[str, Any]] = []
        for group_spec in group_specs:
            engine = self.engine_factory.create_engine()
            scenario_overrides = self._build_group_overrides(
                overrides,
                tuple(group_spec["cover_cells"]),
                dict(group_spec["row_masses"]),
            )
            for ref, value in scenario_overrides.items():
                engine.set_cell(ref, value)
            for ref, value in self._build_imported_spectrum_overrides(engine, imported_spectrum_values).items():
                engine.set_cell(ref, value)
            bounded_spectrum_overrides = self._build_bounded_spectrum_overrides(
                engine,
                thresholds["energy_min_mev"],
                thresholds["energy_max_mev"],
            )
            for ref, value in bounded_spectrum_overrides.items():
                engine.set_cell(ref, value)

            grouped_payloads.append(
                build_results(
                    engine,
                    self.metadata,
                    elements_min_mass=0.0,
                    isotopes_min_atoms=0.0,
                    isotopes_min_activity_ci=0.0,
                    gamma_min_rate=0.0,
                )
            )
            group_runs.append(
                {
                    "engine": engine,
                    "cover_cells": tuple(group_spec["cover_cells"]),
                    "cover_display": str(group_spec["cover_display"]),
                    "rows": tuple(dict(entry) for entry in group_spec["rows"]),
                    "atomic_numbers": tuple(group_spec["atomic_numbers"]),
                }
            )

        if export_overrides is not None:
            export_engine = self.engine_factory.create_engine()
            for ref, value in export_overrides.items():
                export_engine.set_cell(ref, value)
            export_imported_overrides = self._build_imported_spectrum_overrides(
                export_engine,
                imported_spectrum_values,
            )
            for ref, value in export_imported_overrides.items():
                export_engine.set_cell(ref, value)
            export_overrides.update(export_imported_overrides)
            export_bounded_overrides = self._build_bounded_spectrum_overrides(
                export_engine,
                thresholds["energy_min_mev"],
                thresholds["energy_max_mev"],
            )
            export_overrides.update(export_bounded_overrides)

        results = aggregate_results(
            grouped_payloads,
            self.metadata,
            elements_min_mass=thresholds["elements_min_mass"],
            isotopes_min_atoms=0.0,
            isotopes_min_activity_ci=0.0,
            gamma_min_rate=thresholds["gamma_min_rate"],
        )
        return base_engine, export_overrides, export_error, group_runs, results

    def _collect_optimization_inputs(self) -> OptimizationInputs:
        """Collect optimizer-specific settings plus shared irradiation inputs."""
        min_energy, max_energy = self._parse_energy_bounds()
        combined_spectra = tuple(
            SpectrumMixEntry(
                name=self.spectrum_name_vars[entry["row"]].get().strip(),
                fraction=self._parse_float(
                    self.spectrum_fraction_vars[entry["row"]].get(),
                    f"{entry['label']} fraction",
                ),
            )
            for entry in self.gui_config["spectrum_rows"]
        )
        runtime_spectrum_values = self._runtime_spectrum_values(
            self.reactor_var.get(),
            combined_spectra,
        )

        return OptimizationInputs(
            reactor=self.reactor_var.get(),
            fluence_rate=self._parse_float(self.fluence_var.get(), "Neutron Fluence Rate"),
            irradiation_time_sec=self._parse_float(self.irradiation_var.get(), "Irradiation Time"),
            decay_time_sec=self._parse_float(self.decay_var.get(), "Decay Time"),
            energy_min_mev=min_energy,
            energy_max_mev=max_energy,
            trial_mass_g=self._parse_float(
                self.optimization_trial_mass_var.get(),
                "Trial Mass Per Element",
            ),
            min_reaction_activity_ci=self._parse_float(
                self.optimization_min_activity_var.get(),
                "Minimum Reaction Activity",
            ),
            min_recommended_mass_g=self._parse_float(
                self.optimization_min_mass_var.get(),
                "Minimum Mass",
            ),
            max_recommended_mass_g=self._parse_optional_float(
                self.optimization_max_mass_var.get(),
                "Maximum Mass",
            ),
            max_elements=self._parse_int(
                self.optimization_max_elements_var.get(),
                "Max Elements",
            ),
            coverage_mode=self.optimization_coverage_mode_var.get(),
            major_paths_only=self.optimization_major_paths_only_var.get(),
            path_activity_share_threshold=PLOT_REACTION_ACTIVITY_SHARE_THRESHOLD,
            included_elements=self._get_optimization_included_elements(),
            excluded_elements=self._get_optimization_excluded_elements(),
            candidate_cover_materials=self._get_optimization_cover_materials(),
            cover_thickness=self._parse_float(
                self.optimization_cover_thickness_var.get(),
                "Cover Thickness",
            ),
            combined_spectra=combined_spectra,
            imported_spectrum=runtime_spectrum_values,
        )

    def _collect_mass_input_specs(self, overrides: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Snapshot the logical element/alloy inputs and their cover selections."""
        active_cover_cells = tuple(
            str(entry["cell"])
            for entry in self.gui_config["cover_materials"]
            if (self._float_or_none(overrides.get(f"Choose Irradiation Parameters!{entry['cell']}")) or 0.0) > 0.0
        )

        specs: List[Dict[str, Any]] = []
        for entry in self.gui_config["composition_rows"]:
            row = int(entry["row"])
            mass = self._float_or_none(overrides.get(f"Choose Irradiation Parameters!G{row}")) or 0.0
            if mass <= 0.0:
                continue
            specs.append(
                {
                    "row": row,
                    "element": str(entry["element"]),
                    "input_mass": mass,
                    "atomic_numbers": (int(entry["atomic_number"]),),
                    "run_profiles": self._get_row_run_profiles(row, active_cover_cells),
                    "row_masses": {row: mass},
                }
            )

        for entry in self.gui_config.get("material_recipes", []):
            input_key = entry["row"]
            mass = self._parse_float(self.mass_vars[input_key].get(), str(entry["element"]))
            if mass <= 0.0:
                continue
            row_masses = {
                int(component["row"]): mass * float(component["mass_fraction"])
                for component in entry["components"]
            }
            specs.append(
                {
                    "row": input_key,
                    "element": str(entry["element"]),
                    "input_mass": mass,
                    "atomic_numbers": tuple(
                        int(component["atomic_number"])
                        for component in entry["components"]
                    ),
                    "run_profiles": self._get_row_run_profiles(input_key, active_cover_cells),
                    "row_masses": row_masses,
                }
            )
        return specs

    def _collect_overrides(self) -> tuple[Dict[str, Any], Dict[str, float], List[Dict[str, Any]]]:
        """Collect workbook-style cell overrides from the current GUI state."""
        min_energy, max_energy = self._parse_energy_bounds()
        reactor_selection = self.reactor_var.get()
        selected_imported_spectrum = self._selected_imported_spectrum(reactor_selection)
        if reactor_selection == IMPORT_SPECTRUM_UI_LABEL and selected_imported_spectrum is None:
            raise ValueError("Choose a spectrum file before running with Import Spectrum selected.")
        if reactor_selection == COMBINED_REACTOR_UI_LABEL:
            workbook_reactor = COMBINED_REACTOR_WORKBOOK_VALUE
        elif selected_imported_spectrum is not None:
            workbook_reactor = IMPORT_SPECTRUM_UI_LABEL
        else:
            workbook_reactor = reactor_selection
        overrides: Dict[str, Any] = {
            "Choose Irradiation Parameters!B2": workbook_reactor,
            "Choose Irradiation Parameters!B3": self._parse_float(self.fluence_var.get(), "Neutron Fluence Rate"),
            "Choose Irradiation Parameters!B4": self._parse_float(self.irradiation_var.get(), "Irradiation Time"),
            "Choose Irradiation Parameters!B5": self._parse_float(self.decay_var.get(), "Decay Time"),
            "Results - Elements!F4": self._parse_float(self.elements_threshold_var.get(), "Minimum Mass at EOID"),
            "Results - Isotopes!B10": self._parse_float(self.isotopes_threshold_var.get(), "Minimum Atoms of Isotope Provided"),
            "Results - Gamma!E2": self._parse_float(self.gamma_threshold_var.get(), "Minimum Gamma Ray Emission Rate"),
        }

        # These loops preserve the original workbook addressing so the same
        # values can be applied to the JSON engine and workbook export path.
        for entry in self.gui_config["cover_materials"]:
            overrides[f"Choose Irradiation Parameters!{entry['cell']}"] = self._parse_float(
                self.cover_vars[entry["cell"]].get(),
                entry["label"],
            )
        for entry in self.gui_config["spectrum_rows"]:
            row = entry["row"]
            overrides[f"Choose Irradiation Parameters!B{row}"] = self.spectrum_name_vars[row].get().strip()
            overrides[f"Choose Irradiation Parameters!C{row}"] = self._parse_float(
                self.spectrum_fraction_vars[row].get(),
                f"{entry['label']} fraction",
            )
        for entry in self.gui_config["composition_rows"]:
            row = entry["row"]
            overrides[f"Choose Irradiation Parameters!G{row}"] = self._parse_float(
                self.mass_vars[row].get(),
                f"Mass row {row}",
            )

        thresholds = {
            "elements_min_mass": float(overrides["Results - Elements!F4"]),
            "isotopes_min_atoms": float(overrides["Results - Isotopes!B10"]),
            "isotopes_min_activity_ci": self._parse_float(
                self.isotopes_activity_threshold_var.get(),
                "Minimum Activity",
            ),
            "gamma_min_rate": float(overrides["Results - Gamma!E2"]),
            "energy_min_mev": min_energy,
            "energy_max_mev": max_energy,
        }
        input_specs = self._collect_mass_input_specs(overrides)
        return overrides, thresholds, input_specs

    def _current_reactor_report_display(self, reactor_selection: str) -> str:
        """Return the user-facing spectrum/reactor label for report inputs."""
        imported_spectrum = self._selected_imported_spectrum(reactor_selection)
        if imported_spectrum is not None:
            return imported_spectrum.name
        return self._report_reactor_display(reactor_selection)

    def _has_active_worker(self) -> bool:
        """Return True when a calculation, export, or optimization is running."""
        return any(
            thread is not None and thread.is_alive()
            for thread in (self.run_thread, self.export_thread, self.optimize_thread)
        )

    def _update_optimization_action_state(self) -> None:
        """Enable the optimizer feedback action only when a set is available."""
        if self.apply_optimization_button is None:
            return
        has_selection = bool(
            self.current_optimization_result is not None
            and self.current_optimization_result.selected
        )
        if self._has_active_worker() or not has_selection:
            self.apply_optimization_button.state(["disabled"])
        else:
            self.apply_optimization_button.state(["!disabled"])

    def _set_running_state(self, is_running: bool, status_text: str = "Running Python calculation...") -> None:
        """Enable/disable long-running actions while a worker thread is active."""
        if is_running:
            self.status_var.set(status_text)
            self.run_button.state(["disabled"])
            self.export_button.state(["disabled"])
            if self.report_button is not None:
                self.report_button.state(["disabled"])
            if self.optimize_button is not None:
                self.optimize_button.state(["disabled"])
            if self.apply_optimization_button is not None:
                self.apply_optimization_button.state(["disabled"])
        else:
            self.run_button.state(["!disabled"])
            if self._validation_export_available():
                self.export_button.state(["!disabled"])
            if self.report_button is not None:
                self.report_button.state(["!disabled"])
            if self.optimize_button is not None:
                self.optimize_button.state(["!disabled"])
            self._update_optimization_action_state()

    def _start_run(self) -> None:
        """Launch the CRPYT calculation on a background thread."""
        if self._has_active_worker():
            return
        try:
            overrides, thresholds, input_specs = self._collect_overrides()
            combined_spectra = tuple(
                SpectrumMixEntry(
                    name=str(overrides[f"Choose Irradiation Parameters!B{entry['row']}"]).strip(),
                    fraction=float(overrides[f"Choose Irradiation Parameters!C{entry['row']}"]),
                )
                for entry in self.gui_config["spectrum_rows"]
            )
            runtime_spectrum_values = self._runtime_spectrum_values(
                self.reactor_var.get(),
                combined_spectra,
            )
            reactor_report_display = self._current_reactor_report_display(self.reactor_var.get())
        except ValueError as exc:
            messagebox.showerror("Invalid Input", str(exc))
            return
        try:
            spectrum_settings = self._collect_spectrum_settings()
            spectrum_settings_error = None
        except ValueError as exc:
            spectrum_settings = None
            spectrum_settings_error = str(exc)

        self._set_running_state(True)

        def worker() -> None:
            # Each run groups active elements by cover profile, evaluates those
            # groups independently, then combines the result inventories back
            # into one GUI payload.
            try:
                engine, export_overrides, export_error, group_runs, results = self._run_grouped_calculation(
                    overrides,
                    thresholds,
                    input_specs,
                    runtime_spectrum_values,
                )
                spectrum_results = None
                spectrum_summary = self._default_spectrum_summary_text()
                # GREASPY is optional. A CRPYT calculation can still complete
                # even if the detector-spectrum model is unavailable.
                if spectrum_settings_error is not None:
                    spectrum_summary = spectrum_settings_error
                elif self.spectrum_model is not None and spectrum_settings is not None:
                    try:
                        spectrum_inputs, isotope_names_by_zaid, total_active_isotopes = self._build_spectrum_inputs_from_results(
                            results,
                            *spectrum_settings,
                        )
                        spectrum_results = self.spectrum_model.compute_from_activities(
                            spectrum_inputs,
                            isotope_names_by_zaid,
                        )
                        spectrum_summary = self._build_spectrum_summary(
                            spectrum_inputs,
                            spectrum_results,
                            total_active_isotopes,
                        )
                    except Exception as exc:
                        spectrum_summary = f"GREASPY spectrum failed: {exc}"
                payload = (
                    "success",
                    engine,
                    export_overrides,
                    export_error,
                    group_runs,
                    results,
                    dict(overrides),
                    reactor_report_display,
                    dict(thresholds),
                    [dict(spec) for spec in input_specs],
                    spectrum_results,
                    spectrum_summary,
                )
            except Exception:
                payload = ("error", traceback.format_exc())
            self.root.after(0, lambda: self._finish_run(payload))

        self.run_thread = threading.Thread(target=worker, daemon=True)
        self.run_thread.start()

    def _finish_run(self, payload: tuple[Any, ...]) -> None:
        """Consume the worker payload and refresh the GUI on the main thread."""
        self._set_running_state(False)
        self.run_thread = None
        if payload[0] == "error":
            self.status_var.set("Calculation failed.")
            messagebox.showerror("Calculation Failed", payload[1])
            return

        (
            _status,
            engine,
            overrides,
            export_error,
            group_runs,
            results,
            run_overrides,
            reactor_report_display,
            run_thresholds,
            run_input_specs,
            spectrum_results,
            spectrum_summary,
        ) = payload
        self.current_engine = engine
        self.current_group_runs = list(group_runs)
        self.current_overrides = overrides
        self.current_export_unavailable_reason = export_error
        self.current_results = results
        self.current_run_overrides = dict(run_overrides)
        self.current_run_reactor_display = str(reactor_report_display)
        self.current_run_thresholds = dict(run_thresholds)
        self.current_run_input_specs = [dict(spec) for spec in run_input_specs]
        self.current_spectrum_results = spectrum_results
        self.spectrum_summary_var.set(spectrum_summary)
        self.status_var.set("Calculation complete.")
        self._render_all_results()

    def _render_all_results(self) -> None:
        """Push the latest calculation results into every tab."""
        if self.current_results is None:
            return

        summary = self.current_results["isotopes_summary"] | self.current_results["gamma_summary"]
        for key, label in self.summary_labels.items():
            label.configure(text=self._format_value(summary.get(key, "")))

        element_rows = self._get_filtered_element_results()
        isotope_rows, reference_rows = self._get_filtered_isotope_results()
        self.current_plot_payload = self._build_sensitivity_plot_payload()
        self._fill_tree(self.elements_tree, ELEMENT_HEADERS, element_rows)
        self._fill_tree(self.elements_all_tree, ELEMENT_HEADERS, self.current_results["elements_all"])
        self._fill_tree(self.isotopes_tree, ISOTOPE_HEADERS, isotope_rows)
        self._fill_tree(
            self.isotopes_ref_tree,
            ISOTOPE_REFERENCE_HEADERS,
            reference_rows,
        )
        self._fill_tree(self.gamma_tree, GAMMA_HEADERS, self.current_results["gamma_table"])

        self._draw_elements_plot()
        self._draw_isotopes_plot()
        self._draw_gamma_plot()
        self._draw_sensitivity_plot()

    def _fill_tree(
        self,
        tree: ttk.Treeview | None,
        headers: tuple[str, ...],
        rows: List[List[Any]],
    ) -> None:
        """Replace the contents of a Treeview using preformatted result rows."""
        if tree is None:
            return
        tree.delete(*tree.get_children())
        for row in rows:
            tree.insert("", "end", values=[self._format_value(value) for value in row[: len(headers)]])

    def _render_optimization_results(self) -> None:
        """Render the latest optimizer result into the Optimization tab."""
        result = self.current_optimization_result
        if result is None:
            self._fill_tree(self.optimization_tree, OPTIMIZATION_HEADERS, [])
            self._fill_tree(self.optimization_pool_tree, OPTIMIZATION_POOL_HEADERS, [])
            self.optimization_summary_var.set(
                "Run optimization to rank candidate elements. "
                "The optimizer uses irradiation settings from the Inputs tab. "
                "Optional cover scenarios are configured on the Optimization tab. "
                "Leave Include Elements blank to consider all available elements."
            )
            self._update_optimization_action_state()
            return

        selected_rows = [
            [
                selection.rank,
                selection.element,
                selection.cover_material or "Bare",
                selection.reaction_label,
                selection.recommended_mass_g,
                selection.estimated_activity_ci,
                selection.total_reaction_rate,
                selection.centroid_energy_mev,
                selection.incremental_coverage,
                selection.cumulative_coverage,
            ]
            for selection in result.selected
        ]
        pool_rows = [
            [entry.element, entry.status]
            for entry in result.candidate_pool
        ]
        omitted_by_mass_bounds = sum(
            1 for entry in result.candidate_pool
            if entry.status.startswith("Omitted:")
        )

        summary = (
            f"Selected {len(result.selected)} element(s); final coverage = "
            f"{self._format_value(result.final_coverage)} ({result.coverage_mode} weighting). "
            f"{len(result.eligible_elements)} eligible candidate element(s), "
            f"{len(result.ineligible_elements)} ineligible, "
            f"{result.qualifying_reactions} mass-bounded qualifying reaction path(s). "
            f"{omitted_by_mass_bounds} element(s) omitted by mass bounds. "
            "The optimizer evaluates the bare case plus each selected single-cover scenario. "
            "Applying the set back to Inputs uses the recommended per-element masses "
            "needed to meet the optimizer's minimum activity target."
        )
        if not result.selected:
            summary = (
                f"No element set satisfied the requested optimization threshold. {summary}"
            )

        self.optimization_summary_var.set(summary)
        self._fill_tree(self.optimization_tree, OPTIMIZATION_HEADERS, selected_rows)
        self._fill_tree(self.optimization_pool_tree, OPTIMIZATION_POOL_HEADERS, pool_rows)
        self._update_optimization_action_state()

    def _draw_elements_plot(self) -> None:
        """Render the top filtered element masses as a horizontal bar chart."""
        if self.current_results is None or self.elements_axes is None or self.elements_canvas is None:
            return
        try:
            rows = self._get_filtered_element_results()
        except ValueError:
            return
        ax = self.elements_axes
        ax.clear()
        if not rows:
            ax.set_title("No element results above threshold")
            self.elements_canvas.draw()
            return
        sorted_rows = sorted(rows, key=lambda row: abs(float(row[2])), reverse=True)[:12]
        labels = [str(row[0]) for row in sorted_rows][::-1]
        values = [float(row[2]) for row in sorted_rows][::-1]
        ax.barh(labels, values, color="#2f6f8f")
        ax.set_title("Top Element Masses at EOID")
        ax.set_xlabel("Mass (g)")
        self.elements_canvas.draw()

    def _draw_isotopes_plot(self) -> None:
        """Render the top filtered isotope activities as a bar chart."""
        if self.current_results is None or self.isotopes_axes is None or self.isotopes_canvas is None:
            return
        try:
            rows, _reference_rows = self._get_filtered_isotope_results()
        except ValueError:
            return
        ax = self.isotopes_axes
        ax.clear()
        if not rows:
            ax.set_title("No isotope results above threshold")
            self.isotopes_canvas.draw()
            return
        sorted_rows = sorted(rows, key=lambda row: float(row[7]), reverse=True)[:12]
        labels = [str(row[2]) for row in sorted_rows]
        values = [float(row[7]) for row in sorted_rows]
        ax.bar(labels, values, color="#c25b2a")
        ax.set_title("Top Isotopes by Activity (Ci)")
        ax.set_ylabel("Activity (Ci)")
        ax.tick_params(axis="x", rotation=45)
        self.isotopes_canvas.draw()

    def _draw_gamma_plot(self) -> None:
        """Render the GREASPY detector spectrum with the current display scale."""
        if self.gamma_axes is None or self.gamma_canvas is None:
            return
        ax = self.gamma_axes
        ax.clear()
        spectrum_results = self.current_spectrum_results
        if spectrum_results is None:
            ax.set_title("GREASPY detector spectrum unavailable")
            ax.set_xlabel("Energy (keV)")
            ax.set_ylabel("Counts")
            ax.text(
                0.5,
                0.5,
                self.spectrum_summary_var.get(),
                ha="center",
                va="center",
                transform=ax.transAxes,
                wrap=True,
            )
            self.gamma_canvas.draw()
            return

        filtered_points = [
            (energy, total, background)
            for energy, total, background in zip(
                spectrum_results.energy_kev,
                spectrum_results.total_counts,
                spectrum_results.background_counts,
            )
            if energy >= SPECTRUM_MIN_DISPLAY_ENERGY_KEV
        ]
        if not filtered_points:
            ax.set_title("GREASPY detector spectrum unavailable")
            ax.set_xlabel("Energy (keV)")
            ax.set_ylabel("Counts")
            ax.text(
                0.5,
                0.5,
                f"No spectrum bins are available at or above {SPECTRUM_MIN_DISPLAY_ENERGY_KEV:g} keV.",
                ha="center",
                va="center",
                transform=ax.transAxes,
                wrap=True,
            )
            self.gamma_canvas.draw()
            return

        energies_kev = [point[0] for point in filtered_points]
        total_counts_raw = [point[1] for point in filtered_points]
        background_counts_raw = [point[2] for point in filtered_points]
        total_counts = total_counts_raw
        background_counts = background_counts_raw
        if self.gamma_scale_var.get() == "log":
            total_counts = [value if value > 0.0 else float("nan") for value in total_counts_raw]
            background_counts = [value if value > 0.0 else float("nan") for value in background_counts_raw]

        ax.plot(
            energies_kev,
            total_counts,
            color="#1f1f1f",
            linewidth=1.1,
            label="Total counts",
        )
        if any(value > 0.0 for value in background_counts_raw):
            ax.plot(
                energies_kev,
                background_counts,
                color="#8a8a8a",
                linewidth=0.9,
                linestyle="--",
                label="Background",
            )
        ax.set_title("GREASPY Detector Spectrum at EOID")
        ax.set_xlabel("Energy (keV)")
        ax.set_ylabel("Counts")
        ax.set_yscale(self.gamma_scale_var.get())
        positive_counts = [value for value in total_counts_raw if value > 0.0]
        if self.gamma_scale_var.get() == "log":
            ymax = max(positive_counts) * 1.25 if positive_counts else 1.0
            ax.set_ylim(SPECTRUM_LOG_YMIN, max(ymax, SPECTRUM_LOG_YMIN * 10.0))
        else:
            ymax = max(total_counts_raw) * 1.25 if total_counts_raw else 1.0
            ax.set_ylim(SPECTRUM_LINEAR_YMIN, max(ymax, SPECTRUM_LINEAR_YMIN + 1.0))
        xmax = max(energies_kev)
        if xmax <= SPECTRUM_MIN_DISPLAY_ENERGY_KEV:
            xmax = SPECTRUM_MIN_DISPLAY_ENERGY_KEV * 1.01
        ax.set_xlim(SPECTRUM_MIN_DISPLAY_ENERGY_KEV, xmax)
        ax.grid(True, which="both", alpha=0.22)
        ax.legend(loc="upper right", fontsize=8)
        self.gamma_canvas.draw()

    def _draw_sensitivity_plot(self) -> None:
        """Render the input spectrum plus reaction-rate heatmap bands."""
        if (
            not self._ensure_sensitivity_axes()
            or self.sensitivity_axes is None
            or self.sensitivity_secondary_axes is None
        ):
            return

        ax = self.sensitivity_axes
        ax2 = self.sensitivity_secondary_axes
        ax.set_xscale("linear")
        ax2.set_xscale("linear")
        ax.clear()
        ax2.clear()
        self.sensitivity_colorbar = None
        if self.sensitivity_colorbar_axis is not None:
            self.sensitivity_colorbar_axis.clear()
            self.sensitivity_colorbar_axis.set_visible(False)

        payload = self.current_plot_payload
        if payload is None or not payload.get("energies"):
            ax.set_title("Run a calculation to generate the plot")
            ax.set_xlabel("Energy (MeV)")
            ax.set_ylabel("Input Spectrum")
            ax.set_xscale("log")
            ax.set_facecolor(HEATMAP_ZERO_COLOR)
            ax.grid(True, which="both", alpha=0.22)
            self.sensitivity_note_var.set("Run a calculation to populate the plot.")
            self.sensitivity_canvas.draw()
            return

        energies = payload["energies"]
        spectrum = payload["spectrum"]
        reactions = payload["reactions"]
        heatmap_cmap = self._build_sensitivity_heatmap_cmap()
        heatmap_norm = Normalize(vmin=0.0, vmax=1.0)

        # The left axis carries the input spectrum, while the twin right axis
        # is used as a labeled canvas for the heatmap bands.
        ax.plot(energies, spectrum, color="#111111", linewidth=2.2, label="Input Spectrum")
        ax.set_title("Input Spectrum and Active Reaction Rates")
        ax.set_xlabel("Energy (MeV)")
        ax.set_ylabel("Input Spectrum")
        ax.set_xscale("log")
        ax.set_yscale(self.sensitivity_scale_var.get())
        ax.grid(True, which="both", alpha=0.22)
        ax.set_zorder(2)
        ax.set_facecolor(HEATMAP_ZERO_COLOR)
        ax.patch.set_alpha(0.0)

        ax2.set_facecolor(HEATMAP_ZERO_COLOR)
        ax2.set_zorder(1)
        ax2.set_xscale("log")
        ax2.set_ylabel("Reaction")
        ax2.yaxis.set_label_position("right")
        ax2.yaxis.tick_right()
        ax2.set_ylim(0.0, 1.0)
        ax2.tick_params(axis="y", colors="#8c2f39", labelsize=6, length=0, pad=4)
        ax2.spines["right"].set_color("#8c2f39")
        ax2.yaxis.label.set_color("#8c2f39")
        ax2.spines["left"].set_visible(False)
        ax2.spines["top"].set_visible(False)
        ax2.spines["bottom"].set_visible(False)
        ax2.grid(False)

        # Heatmap bands can either be weighted by their integrated reaction
        # rate or split evenly; the toggle only changes layout, not chemistry.
        x_edges = self._build_linear_energy_edges(energies)
        band_gap = 0.006 if len(reactions) > 1 else 0.0
        top_margin = 0.02
        bottom_margin = 0.02
        available_height = max(0.0, 1.0 - top_margin - bottom_margin - band_gap * max(len(reactions) - 1, 0))
        band_height_mode = self.sensitivity_band_scale_var.get()
        band_weights = [max(float(reaction.get("total_reaction_rate", 0.0)), 0.0) for reaction in reactions]
        total_band_weight = sum(band_weights)
        use_even_bands = (
            band_height_mode == "even"
            or not reactions
            or total_band_weight <= 0.0
        )
        min_band_height = min(0.012, available_height / max(len(reactions), 1) * 0.75) if reactions else 0.0
        extra_height = max(0.0, available_height - min_band_height * len(reactions))

        tick_positions: List[float] = []
        tick_labels: List[str] = []
        current_top = 1.0 - top_margin
        for index, reaction in enumerate(reactions):
            if use_even_bands:
                band_fraction = 1.0 / len(reactions) if reactions else 0.0
            else:
                band_fraction = band_weights[index] / total_band_weight
            band_height = min_band_height + extra_height * band_fraction
            band_bottom = current_top - band_height
            if band_bottom < bottom_margin:
                band_bottom = bottom_margin
            ax2.pcolormesh(
                x_edges,
                [band_bottom, current_top],
                [reaction["normalized_values"]],
                cmap=heatmap_cmap,
                norm=heatmap_norm,
                shading="flat",
                alpha=0.88,
                rasterized=True,
            )
            tick_positions.append((band_bottom + current_top) / 2.0)
            tick_labels.append(reaction["label"])
            current_top = band_bottom - band_gap
            if current_top <= bottom_margin:
                break

        if tick_positions:
            ax2.set_yticks(tick_positions, labels=tick_labels)
        else:
            ax2.set_yticks([])

        positive_spectrum = [value for value in spectrum if value > 0.0]
        if positive_spectrum and self.sensitivity_scale_var.get() == "log":
            ax.set_ylim(min(positive_spectrum) * 0.8, max(positive_spectrum) * 1.25)
        energy_limits = self._resolve_sensitivity_plot_energy_limits(energies)
        if energy_limits is not None:
            ax.set_xlim(*energy_limits)
            ax2.set_xlim(*energy_limits)

        if reactions and self.sensitivity_colorbar_axis is not None:
            self.sensitivity_colorbar_axis.set_visible(True)
            self.sensitivity_colorbar = ax.figure.colorbar(
                ScalarMappable(norm=heatmap_norm, cmap=heatmap_cmap),
                cax=self.sensitivity_colorbar_axis,
                orientation="vertical",
            )
            self.sensitivity_colorbar.ax.tick_params(labelsize=7, length=2, pad=2)
            self.sensitivity_colorbar.set_ticks([0.0, 0.5, 1.0])
            self.sensitivity_colorbar.outline.set_visible(False)
            self.sensitivity_colorbar_axis.set_facecolor(HEATMAP_ZERO_COLOR)
            self.sensitivity_colorbar.ax.yaxis.set_label_position("right")
            self.sensitivity_colorbar.set_label(
                r"Relative reaction rate ($\sigma \times \phi$)",
                fontsize=8,
                rotation=270,
                labelpad=14,
            )

        layout_note = (
            "Band height mode: evenly distributed."
            if use_even_bands
            else "Band height mode: scaled by integrated reaction rate."
        )
        self.sensitivity_note_var.set(f"{payload['note']} {layout_note}")
        self.sensitivity_canvas.draw()

    def _apply_optimization_to_inputs_and_run(self) -> None:
        """Push the recommended set into Inputs using the optimizer's mass targets."""
        result = self.current_optimization_result
        if result is None or not result.selected:
            messagebox.showinfo(
                "No Recommended Set",
                "Run optimization and generate at least one recommended element before applying it to Inputs.",
            )
            return
        if self._has_active_worker():
            return

        selected_masses = {
            selection.element: selection.recommended_mass_g
            for selection in result.selected
        }
        selected_cover_by_element = {
            selection.element: selection.cover_material
            for selection in result.selected
        }
        selected_covers = {
            selection.cover_material
            for selection in result.selected
            if selection.cover_material not in (None, "")
        }

        # Optimizer write-back uses one shared thickness value for all selected
        # cover scenarios, then applies those covers row by row in the Inputs
        # table using the new per-element cover checkboxes.
        for entry in self.gui_config["cover_materials"]:
            cell = entry["cell"]
            cover_value = float(result.inputs.cover_thickness) if str(entry["label"]) in selected_covers else 0.0
            self.cover_vars[cell].set(f"{cover_value:.16g}" if cover_value else "0")
        for variable in self.mass_bare_vars.values():
            variable.set(True)
        for variable in self.mass_cover_vars.values():
            variable.set(False)
        for entry in self.gui_config.get("material_recipes", []):
            self.mass_vars[entry["row"]].set(DEFAULT_ELEMENT_MASS_TEXT)

        # The optimizer now returns the inferred minimum mass needed for each
        # selected element to satisfy the requested activity target.
        for entry in self.gui_config["composition_rows"]:
            row = int(entry["row"])
            element = str(entry["element"])
            mass_value = selected_masses.get(element)
            self.mass_vars[row].set(
                f"{mass_value:.16g}" if mass_value is not None else DEFAULT_ELEMENT_MASS_TEXT
            )
            selected_cover = selected_cover_by_element.get(element)
            if selected_cover not in (None, ""):
                for cover_entry in self.gui_config["cover_materials"]:
                    if str(cover_entry["label"]) == str(selected_cover):
                        self.mass_cover_vars[(row, str(cover_entry["cell"]))].set(True)
                        break

        self.notebook.select(self.inputs_tab)
        self.status_var.set("Applied optimized element set to Inputs. Running calculation...")
        self._start_run()

    def _start_optimization(self) -> None:
        """Launch the element-set optimizer on a background thread."""
        if self._has_active_worker():
            return
        try:
            optimization_inputs = self._collect_optimization_inputs()
        except ValueError as exc:
            messagebox.showerror("Invalid Optimization Input", str(exc))
            return

        self.current_optimization_result = None
        self._fill_tree(self.optimization_tree, OPTIMIZATION_HEADERS, [])
        self._fill_tree(self.optimization_pool_tree, OPTIMIZATION_POOL_HEADERS, [])
        self.optimization_summary_var.set("Running element optimization...")
        self._set_running_state(True, "Running element optimization...")

        def worker() -> None:
            try:
                result = optimize_element_set(optimization_inputs, bundle_dir=self.bundle_dir)
                payload = ("success", result)
            except Exception:
                payload = ("error", traceback.format_exc())
            self.root.after(0, lambda: self._finish_optimization(payload))

        self.optimize_thread = threading.Thread(target=worker, daemon=True)
        self.optimize_thread.start()

    def _finish_optimization(self, payload: tuple[Any, ...]) -> None:
        """Handle optimizer completion and refresh the Optimization tab."""
        self._set_running_state(False)
        self.optimize_thread = None
        if payload[0] == "error":
            self.status_var.set("Element optimization failed.")
            self.optimization_summary_var.set("Element optimization failed.")
            messagebox.showerror("Optimization Failed", payload[1])
            return

        _status, result = payload
        self.current_optimization_result = result
        self.status_var.set("Element optimization complete.")
        self._render_optimization_results()

    def _start_report(self) -> None:
        """Prompt for a PDF path and write a report for the current results."""
        if self.current_results is None:
            messagebox.showinfo("No Results", "Run a calculation before generating a PDF report.")
            return
        if self._has_active_worker():
            return

        path = filedialog.asksaveasfilename(
            title="Generate PDF Report",
            defaultextension=".pdf",
            initialfile="CRPYT_report.pdf",
            filetypes=[("PDF Report", "*.pdf")],
        )
        if not path:
            return

        out_path = Path(path)
        if self.report_button is not None:
            self.report_button.state(["disabled"])
        self.status_var.set("Generating PDF report...")
        self.root.update_idletasks()
        try:
            self._write_pdf_report(out_path)
        except Exception:
            self.status_var.set("PDF report failed.")
            messagebox.showerror("Report Failed", traceback.format_exc())
        else:
            self.status_var.set(f"Generated PDF report: {out_path.name}")
            messagebox.showinfo("Report Complete", f"PDF report saved to:\n{out_path}")
        finally:
            if self.report_button is not None and not self._has_active_worker():
                self.report_button.state(["!disabled"])

    def _report_number(self, value: Any) -> str:
        """Format report table numbers compactly and consistently."""
        numeric = self._float_or_none(value)
        if numeric is None:
            return self._format_value(value)
        if numeric == 0.0:
            return "0"
        abs_value = abs(numeric)
        if abs_value >= 1.0e4 or abs_value < 1.0e-3:
            return f"{numeric:.4e}"
        return f"{numeric:.6f}".rstrip("0").rstrip(".")

    def _report_key_value(self, label: str, value: Any) -> str:
        """Return a fixed-width key/value line for text-based report pages."""
        return f"{label:<38} {value}"

    def _report_reactor_display(self, value: Any) -> str:
        """Convert the workbook reactor token back to the GUI/report label."""
        text = str(value or "").strip()
        if text == COMBINED_REACTOR_WORKBOOK_VALUE:
            return COMBINED_REACTOR_UI_LABEL
        if text == IMPORT_SPECTRUM_UI_LABEL:
            imported_spectrum = self._selected_imported_spectrum(self.reactor_var.get())
            if imported_spectrum is not None:
                return imported_spectrum.name
        return text or self.reactor_var.get()

    def _build_report_input_lines(self) -> List[str]:
        """Summarize the run inputs that produced the current result payload."""
        overrides = self.current_run_overrides or {}
        thresholds = self.current_run_thresholds or {}
        specs = self.current_run_input_specs

        reactor = self.current_run_reactor_display or self._report_reactor_display(
            overrides.get("Choose Irradiation Parameters!B2", self.reactor_var.get())
        )
        fluence_rate = self._float_or_none(
            overrides.get("Choose Irradiation Parameters!B3", self.fluence_var.get())
        ) or 0.0
        irradiation_sec = self._float_or_none(
            overrides.get("Choose Irradiation Parameters!B4", self.irradiation_var.get())
        ) or 0.0
        decay_sec = self._float_or_none(
            overrides.get("Choose Irradiation Parameters!B5", self.decay_var.get())
        ) or 0.0
        total_mass = sum(float(spec.get("input_mass", 0.0) or 0.0) for spec in specs)

        lines = [
            "## Irradiation Parameters",
            self._report_key_value("Reactor Type", reactor),
            self._report_key_value("Neutron Fluence Rate (n/cm^2/sec)", self._report_number(fluence_rate)),
            self._report_key_value("Irradiation Time (sec)", self._report_number(irradiation_sec)),
            self._report_key_value("Irradiation Time (days)", self._report_number(irradiation_sec / 86400.0)),
            self._report_key_value("Decay Time (sec)", self._report_number(decay_sec)),
            self._report_key_value("Decay Time (days)", self._report_number(decay_sec / 86400.0)),
            self._report_key_value("Total Fluence (n/cm^2)", self._report_number(fluence_rate * irradiation_sec)),
            self._report_key_value("Energy Minimum (MeV)", self._report_number(thresholds.get("energy_min_mev", self.energy_min_var.get()))),
            self._report_key_value("Energy Maximum (MeV)", self._report_number(thresholds.get("energy_max_mev", self.energy_max_var.get()))),
            self._report_key_value("Total Input Mass (g)", self._report_number(total_mass)),
            "",
            "## Nonzero Cover Materials",
        ]

        cover_lines: List[str] = []
        for entry in self.gui_config["cover_materials"]:
            cell = str(entry["cell"])
            thickness = self._float_or_none(
                overrides.get(f"Choose Irradiation Parameters!{cell}", self.cover_vars[cell].get())
            ) or 0.0
            if thickness > 0.0:
                cover_lines.append(self._report_key_value(str(entry["label"]), f"{self._report_number(thickness)} mm"))
        lines.extend(cover_lines or ["None"])

        lines.extend(["", "## Nonzero Elements and Masses"])
        if specs:
            for spec in specs:
                profiles = ", ".join(
                    self._cover_profile_display(tuple(profile))
                    for profile in spec.get("run_profiles", (tuple(),))
                )
                if not profiles:
                    profiles = "Bare"
                mass = self._report_number(spec.get("input_mass", 0.0))
                lines.append(f"{str(spec.get('element', '')):<16} {mass:>14} g   Profiles: {profiles}")
        else:
            lines.append("None")

        if reactor == COMBINED_REACTOR_UI_LABEL:
            lines.extend(["", "## Combined Spectrum Fractions"])
            added = False
            for entry in self.gui_config["spectrum_rows"]:
                row = entry["row"]
                name = overrides.get(
                    f"Choose Irradiation Parameters!B{row}",
                    self.spectrum_name_vars[row].get(),
                )
                fraction = self._float_or_none(
                    overrides.get(
                        f"Choose Irradiation Parameters!C{row}",
                        self.spectrum_fraction_vars[row].get(),
                    )
                ) or 0.0
                if fraction > 0.0:
                    added = True
                    lines.append(self._report_key_value(str(name), self._report_number(fraction)))
            if not added:
                lines.append("None")

        return lines

    def _build_report_summary_lines(self) -> List[str]:
        """Summarize the headline outputs requested from the Summary tab."""
        if self.current_results is None:
            return []
        summary = self.current_results["isotopes_summary"] | self.current_results["gamma_summary"]
        keys = [
            "Total Activity (Ci)",
            "Total Activity (Bq)",
            "Gamma-Ray Emission Rate (Bq)",
            "Gamma Dose Rate (R/hr) at 30 cm",
            "Beta Emission Rate (Bq)",
            "Beta Dose Rate (R/hr) at Contact",
        ]
        lines = ["## Summary Outputs"]
        for key in keys:
            lines.append(self._report_key_value(key, self._report_number(summary.get(key, ""))))
        return lines

    def _element_label_for_atomic_number(self, atomic_number: int) -> str:
        """Resolve an atomic number to the GUI element label when possible."""
        for entry in self.gui_config["composition_rows"]:
            if int(entry["atomic_number"]) == int(atomic_number):
                return str(entry["element"])
        return f"Z={atomic_number}"

    def _report_source_label(self, atomic_number: int, cover_display: str) -> str:
        """Format a report grouping label for a parent element and cover profile."""
        label = self._element_label_for_atomic_number(atomic_number)
        if cover_display and cover_display != "Bare":
            return f"{label} [{cover_display}]"
        return label

    def _build_report_source_map(
        self,
        isotope_zaids: set[int],
    ) -> Dict[int, Dict[str, Dict[str, float]]]:
        """Estimate product-isotope source groups from production-path metadata."""
        source_map: Dict[int, Dict[str, Dict[str, float]]] = {}
        active_group_runs = [group for group in self.current_group_runs if group.get("rows")]
        if not active_group_runs:
            return source_map

        for group_run in active_group_runs:
            group_engine = group_run["engine"]
            group_atomic_numbers = {
                int(value)
                for value in group_run.get("atomic_numbers", ())
            }
            if not group_atomic_numbers:
                continue

            other_lookup = self._build_other_mt_lookup(group_engine)
            production_ws = group_engine.wb_formula["Production"]
            cover_display = str(group_run.get("cover_display", "Bare"))
            seen = set()
            for row in range(2, production_ws.max_row + 1):
                parent_zaid = self._int_or_none(production_ws._get_value(row, 2))
                mt_value = production_ws._get_value(row, 4)
                reaction = production_ws._get_value(row, 5)
                daughter_zaid = self._int_or_none(production_ws._get_value(row, 10))
                if parent_zaid is None or daughter_zaid is None:
                    continue
                if mt_value in (None, "") or reaction in (None, ""):
                    continue
                parent_atomic_number = parent_zaid // 1000
                if parent_atomic_number not in group_atomic_numbers:
                    continue
                if daughter_zaid not in isotope_zaids:
                    continue

                mt_meta = self._format_mt_meta(mt_value, daughter_zaid)
                column_name = other_lookup.get((parent_zaid, mt_meta))
                if not column_name:
                    continue
                key = (cover_display, parent_zaid, daughter_zaid, mt_meta, column_name)
                if key in seen:
                    continue
                seen.add(key)

                column_index = column_index_from_string(column_name)
                weighted_xs = self._float_or_none(group_engine._get_cell("Other MT", 1, column_index)) or 0.0
                source_label = self._report_source_label(parent_atomic_number, cover_display)
                source_entry = source_map.setdefault(daughter_zaid, {}).setdefault(
                    source_label,
                    {"weight": 0.0, "count": 0.0},
                )
                source_entry["weight"] += max(weighted_xs, 0.0)
                source_entry["count"] += 1.0
        return source_map

    def _build_report_isotope_groups(self) -> List[Dict[str, Any]]:
        """Return isotope report rows grouped by generating parent element."""
        isotope_rows, _reference_rows = self._get_filtered_isotope_results()
        isotope_zaids = {
            zaid
            for row in isotope_rows
            for zaid in (self._int_or_none(row[1]),)
            if zaid is not None
        }
        source_map = self._build_report_source_map(isotope_zaids)
        grouped: Dict[str, Dict[str, Any]] = {}

        for row in isotope_rows:
            zaid = self._int_or_none(row[1])
            if zaid is None:
                continue
            source_entries = source_map.get(zaid, {})
            if source_entries:
                total_weight = sum(entry["weight"] for entry in source_entries.values())
                if total_weight > 0.0:
                    contributions = [
                        (label, entry["weight"] / total_weight)
                        for label, entry in source_entries.items()
                    ]
                else:
                    total_count = sum(entry["count"] for entry in source_entries.values()) or 1.0
                    contributions = [
                        (label, entry["count"] / total_count)
                        for label, entry in source_entries.items()
                    ]
            else:
                contributions = [("Unattributed", 1.0)]

            for source_label, share in contributions:
                if share <= 0.0:
                    continue
                activity = (self._float_or_none(row[7]) or 0.0) * share
                gamma_rate = (self._float_or_none(row[8]) or 0.0) * share
                gamma_dose = (self._float_or_none(row[9]) or 0.0) * share
                beta_dose = (self._float_or_none(row[11]) or 0.0) * share
                group = grouped.setdefault(
                    source_label,
                    {"label": source_label, "total_activity": 0.0, "rows": []},
                )
                group["total_activity"] += activity
                group["rows"].append(
                    {
                        "isotope": str(row[2]),
                        "zaid": zaid,
                        "share": share,
                        "activity": activity,
                        "gamma_rate": gamma_rate,
                        "gamma_dose": gamma_dose,
                        "beta_dose": beta_dose,
                    }
                )

        groups = list(grouped.values())
        for group in groups:
            group["rows"].sort(key=lambda item: (-item["activity"], item["isotope"], item["zaid"]))
        groups.sort(key=lambda item: (-item["total_activity"], item["label"]))
        return groups

    def _write_report_lines(
        self,
        pdf: PdfPages,
        title: str,
        lines: List[str],
        *,
        landscape: bool = False,
    ) -> None:
        """Write one or more simple text pages to the PDF report."""
        page_size = REPORT_LANDSCAPE_PAGE_SIZE if landscape else REPORT_PAGE_SIZE
        max_lines = 34 if landscape else 43
        line_height = 0.024 if landscape else 0.020
        chunks = [lines[index:index + max_lines] for index in range(0, len(lines), max_lines)] or [[]]
        total_pages = len(chunks)

        for page_index, chunk in enumerate(chunks, start=1):
            fig = Figure(figsize=page_size)
            fig.text(0.5, 0.965, title, ha="center", va="top", fontsize=14, weight="bold")
            if total_pages > 1:
                fig.text(
                    0.94,
                    0.965,
                    f"Page {page_index} of {total_pages}",
                    ha="right",
                    va="top",
                    fontsize=8,
                )
            y = 0.925
            for line in chunk:
                if line.startswith("## "):
                    y -= line_height * 0.35
                    fig.text(0.06, y, line[3:], ha="left", va="top", fontsize=10, weight="bold")
                else:
                    fig.text(
                        0.06,
                        y,
                        line,
                        ha="left",
                        va="top",
                        fontsize=7.5 if landscape else 8,
                        family="DejaVu Sans Mono",
                    )
                y -= line_height
            pdf.savefig(fig)
            fig.clear()

    def _write_report_isotope_pages(self, pdf: PdfPages) -> None:
        """Write isotope results grouped by generating parent element."""
        groups = self._build_report_isotope_groups()
        if not groups:
            self._write_report_lines(
                pdf,
                "Isotope Results by Generating Element",
                ["No isotope rows met the current isotope filters."],
                landscape=True,
            )
            return

        lines: List[str] = [
            "Rows are grouped by parent element and cover profile. Shares are estimated from production-path weighting when multiple sources feed one isotope.",
            "",
        ]
        header = (
            f"{'Product Isotope':<20} {'ZAID':>8} {'Share':>8} "
            f"{'Activity (Ci)':>14} {'Gamma Rate (Bq)':>16} "
            f"{'Gamma Dose (R/hr)':>18} {'Beta Dose (R/hr)':>17}"
        )
        separator = "-" * len(header)
        for group in groups:
            lines.append(
                f"## Generated by: {group['label']}  "
                f"(est. activity total {self._report_number(group['total_activity'])} Ci)"
            )
            lines.append(header)
            lines.append(separator)
            for row in group["rows"]:
                lines.append(
                    f"{row['isotope']:<20.20} {row['zaid']:>8} "
                    f"{row['share'] * 100.0:>7.1f}% "
                    f"{self._report_number(row['activity']):>14} "
                    f"{self._report_number(row['gamma_rate']):>16} "
                    f"{self._report_number(row['gamma_dose']):>18} "
                    f"{self._report_number(row['beta_dose']):>17}"
                )
            lines.append("")

        self._write_report_lines(
            pdf,
            "Isotope Results by Generating Element",
            lines,
            landscape=True,
        )

    def _write_report_plot_page(self, pdf: PdfPages) -> None:
        """Append the current reaction sensitivity plot as the final report page."""
        if self.sensitivity_figure is None or self.sensitivity_canvas is None:
            self._write_report_lines(
                pdf,
                "Reaction Plot",
                ["The reaction plot is not available in this session."],
                landscape=True,
            )
            return
        self._draw_sensitivity_plot()
        pdf.savefig(self.sensitivity_figure, bbox_inches="tight")

    def _write_pdf_report(self, out_path: Path) -> None:
        """Build a PDF report from the current run inputs, outputs, table, and plot."""
        overview_lines = [
            self._report_key_value("Generated", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            self._report_key_value("CRPYT Build", f"{BUILD_VERSION} ({BUILD_DATE})"),
            "",
            *self._build_report_input_lines(),
            "",
            *self._build_report_summary_lines(),
        ]
        with PdfPages(out_path) as pdf:
            metadata = pdf.infodict()
            metadata["Title"] = "CRPYT Report"
            metadata["Author"] = ", ".join(AUTHORS)
            metadata["Subject"] = "CRPYT calculation inputs and results"
            metadata["CreationDate"] = datetime.now()
            self._write_report_lines(pdf, "CRPYT Report", overview_lines)
            self._write_report_isotope_pages(pdf)
            self._write_report_plot_page(pdf)

    def _start_export(self) -> None:
        """Export the current GUI state back into the validation workbook format."""
        if self.current_results is None:
            messagebox.showinfo("No Results", "Run a calculation before exporting a validation workbook.")
            return
        if self.current_export_unavailable_reason:
            messagebox.showinfo("Export Unavailable", self.current_export_unavailable_reason)
            return
        if self.current_overrides is None:
            messagebox.showinfo("Export Unavailable", "This result set does not have an exportable workbook configuration.")
            return
        if self._has_active_worker():
            return
        if not self._validation_export_available():
            messagebox.showerror(
                "Export Unavailable",
                "Validation workbook export is unavailable because the CRPYT JSON bundle could not be loaded.",
            )
            return

        path = filedialog.asksaveasfilename(
            title="Export Workbook",
            defaultextension=".xlsx",
            initialfile="CRPYT_IRDFF_Rev1_validation.xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not path:
            return

        self.status_var.set("Exporting validation workbook...")
        self.export_button.state(["disabled"])
        if self.report_button is not None:
            self.report_button.state(["disabled"])
        overrides = dict(self.current_overrides)
        out_path = Path(path)

        def worker() -> None:
            try:
                self.engine_factory.export_validation_workbook(overrides, out_path)
                payload = ("success", out_path)
            except Exception:
                payload = ("error", traceback.format_exc())
            self.root.after(0, lambda: self._finish_export(payload))

        self.export_thread = threading.Thread(target=worker, daemon=True)
        self.export_thread.start()

    def _finish_export(self, payload: tuple[Any, ...]) -> None:
        """Handle export completion and surface any workbook generation failures."""
        self.export_thread = None
        if self._validation_export_available():
            self.export_button.state(["!disabled"])
        if self.report_button is not None:
            self.report_button.state(["!disabled"])
        if payload[0] == "error":
            self.status_var.set("Validation export failed.")
            messagebox.showerror("Export Failed", payload[1])
            return
        self.status_var.set(f"Exported validation workbook: {payload[1].name}")
        messagebox.showinfo("Export Complete", f"Validation workbook saved to:\n{payload[1]}")

    def _format_about_path(self, path: Path) -> str:
        """Show short bundle paths in About when they live under a known app root."""
        resolved = path.resolve()
        candidate_roots = [
            Path.cwd(),
            Path(__file__).resolve().parent.parent,
            Path(__file__).resolve().parent,
        ]
        for root in candidate_roots:
            try:
                return str(resolved.relative_to(root.resolve()))
            except ValueError:
                continue
        return str(resolved)

    def _build_about_text(self) -> str:
        """Build the long-form About text shown in the modal dialog."""
        runtime_mode = "Executable" if getattr(sys, "frozen", False) else "Source"
        python_version = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
        greasy_status = "Available" if self.spectrum_model is not None else f"Unavailable ({self.spectrum_model_error})"
        lines = [
            "CRPYT Standalone",
            "",
            "Authors:",
            *(f"- {author}, {AUTHOR_AFFILIATION}" for author in AUTHORS),
            "",
            "Build Information:",
            f"- Version: {BUILD_VERSION}",
            f"- Build Date: {BUILD_DATE}",
            f"- Runtime Mode: {runtime_mode}",
            f"- Python: {python_version}",
            f"- Matplotlib: {matplotlib.__version__}",
            f"- CRPYT Bundle: {self._format_about_path(self.bundle_dir)}",
            f"- GREASPY Bundle: {self._format_about_path(self.greasy_bundle_dir)}",
            f"- GREASPY Status: {greasy_status}",
            f"- Validation Workbook Export Available: {'Yes' if self._validation_export_available() else 'No'}",
            "",
            "Features:",
            *(f"- {feature}" for feature in ABOUT_FEATURES),
            "",
            "Notes:",
            "- Standard calculations run from the extracted JSON data bundle rather than the original CRPYT workbook.",
            "- Validation workbook export reconstructs an Excel workbook from the bundled JSON data.",
            "- The Optimization tab reuses the irradiation setup from the Inputs tab and ranks candidate elements with a greedy coverage search.",
        ]
        return "\n".join(lines)

    def _show_about_dialog(self) -> None:
        """Open a modal About dialog that documents the app and current build."""
        if self.about_window is not None and self.about_window.winfo_exists():
            self.about_window.lift()
            self.about_window.focus_force()
            return

        window = tk.Toplevel(self.root)
        self.about_window = window
        window.title("About CRPYT")
        window.geometry("820x620")
        window.transient(self.root)
        window.grab_set()

        def close_dialog() -> None:
            self.about_window = None
            window.destroy()

        window.protocol("WM_DELETE_WINDOW", close_dialog)
        container = ttk.Frame(window, padding=10)
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        # A scrolled read-only text widget keeps the feature list readable
        # without forcing the main window to grow.
        text = scrolledtext.ScrolledText(container, wrap="word", height=30)
        text.grid(row=0, column=0, sticky="nsew")
        text.insert("1.0", self._build_about_text())
        text.configure(state="disabled")

        ttk.Button(container, text="Close", command=close_dialog).grid(
            row=1,
            column=0,
            sticky="e",
            pady=(8, 0),
        )
        window.focus_force()


def main() -> None:
    """Run the desktop GUI entry point."""
    bundle_dir = _bundle_dir()
    root = tk.Tk()
    app = CrpytGuiApp(root, bundle_dir)
    try:
        root.mainloop()
    finally:
        app._shutdown_matplotlib()
        try:
            root.destroy()
        except tk.TclError:
            pass


if __name__ == "__main__":
    main()

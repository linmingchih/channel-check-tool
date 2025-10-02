import json
import json
import re
import sys
import traceback
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

try:
    from PySide6.QtCore import Qt, QSettings, QObject, QThread, Signal
    from PySide6.QtGui import QColor, QBrush, QTextCursor
    from PySide6.QtWidgets import (
        QApplication,
        QFileDialog,
        QCheckBox,
        QGroupBox,
        QHBoxLayout,
        QLabel,
        QListWidget,
        QListWidgetItem,
        QMessageBox,
        QPushButton,
        QFormLayout,
        QDoubleSpinBox,
        QVBoxLayout,
        QWidget,
        QAbstractItemView,
        QComboBox,
        QTabWidget,
        QLineEdit,
        QTableWidget,
        QTableWidgetItem,
        QHeaderView,
        QProgressBar,
        QPlainTextEdit,
        QSizePolicy,
    )
    QT_LIB = "PySide6"
except ImportError:  # pragma: no cover - fallback when PySide6 is unavailable
    from PySide2.QtCore import Qt, QSettings, QObject, QThread, Signal
    from PySide2.QtGui import QColor, QBrush, QTextCursor
    from PySide2.QtWidgets import (
        QApplication,
        QFileDialog,
        QCheckBox,
        QGroupBox,
        QHBoxLayout,
        QLabel,
        QListWidget,
        QListWidgetItem,
        QMessageBox,
        QPushButton,
        QFormLayout,
        QDoubleSpinBox,
        QVBoxLayout,
        QWidget,
        QAbstractItemView,
        QComboBox,
        QTabWidget,
        QLineEdit,
        QTableWidget,
        QTableWidgetItem,
        QHeaderView,
        QProgressBar,
        QPlainTextEdit,
        QSizePolicy,
    )
    QT_LIB = "PySide2"

from pyedb import Edb

ROOT_DIR = Path(__file__).resolve().parent
if ROOT_DIR.name == 'src':
    ROOT_DIR = ROOT_DIR.parent
SRC_DIR = ROOT_DIR / 'src'
if SRC_DIR.exists():
    src_str = str(SRC_DIR)
    if src_str not in sys.path:
        sys.path.append(src_str)

try:  # pragma: no cover - optional dependency at runtime
    from cct import CCT, load_port_metadata, prefix_port_name, DEFAULT_CIRCUIT_VERSION
except ImportError:  # pragma: no cover - allow GUI without CCT backend
    CCT = None
    load_port_metadata = None
    DEFAULT_CIRCUIT_VERSION = "2025.1"

    def prefix_port_name(name: str, sequence: int) -> str:
        base = str(name or '')
        match = re.match(r'^\d+_(.*)$', base)
        if match:
            base = match.group(1)
        base = base.strip()
        return f"{sequence}_{base}" if base else str(sequence)

_COMPONENT_PATTERN = re.compile(r"^U\d+", re.IGNORECASE)

try:
    DIFF_ROW_BRUSH = QBrush(QColor(230, 240, 255))
except NameError:  # pragma: no cover - Qt import failed
    DIFF_ROW_BRUSH = None

SETTINGS_ORG = "Ansys"
SETTINGS_APP = "AEDBNetExplorer"

DEFAULT_EDB_VERSION = "2025.2"
EDB_VERSION_OPTIONS = [
    "None",
    DEFAULT_EDB_VERSION,
    "2025.1",
    "2024.2",
    "2024.1",
    "2023.2",
    "2023.1",
]

DEFAULT_CUTOUT_EXPANSION = 0.002
SWEEP_TYPE_OPTIONS = [
    "linear count",
    "log scale",
    "linear scale",
]
DEFAULT_FREQUENCY_SWEEPS = [
    ["linear count", "0", "1kHz", "1"],
    ["log scale", "1kHz", "0.1GHz", "10"],
    ["linear scale", "0.1GHz", "10GHz", "0.1GHz"],
]


class _CctWorker(QObject):
    progress = Signal(int)
    message = Signal(str)
    finished = Signal(str)
    failed = Signal(str, object)

    def __init__(
        self,
        touchstone_path: Path,
        metadata_path: Path,
        output_path: Optional[Path],
        workdir: Path,
        settings: Dict[str, Dict[str, object]],
        mode: str = 'run',
    ) -> None:
        super().__init__()
        self._touchstone_path = Path(touchstone_path)
        self._metadata_path = Path(metadata_path)
        self._output_path = Path(output_path) if output_path is not None else None
        self._workdir = Path(workdir)
        self._settings = settings
        self._mode = mode

    def run(self) -> None:
        try:
            self.message.emit('Preparing CCT inputs...')
            self.progress.emit(0)
            self._workdir.mkdir(parents=True, exist_ok=True)

            options = self._settings.get('options') or self._settings.get('prune', {})
            threshold_raw = options.get('threshold_db') if isinstance(options, dict) else None
            try:
                threshold_value = float(threshold_raw) if threshold_raw is not None else None
            except (TypeError, ValueError):
                threshold_value = None

            circuit_version = None
            if isinstance(options, dict):
                version_candidate = options.get('circuit_version')
                if version_candidate is not None:
                    circuit_version = str(version_candidate).strip() or None

            cct = CCT(
                str(self._touchstone_path),
                str(self._metadata_path),
                workdir=self._workdir,
                threshold_db=threshold_value,
                circuit_version=circuit_version,
            )

            self.message.emit('Configuring transmit settings...')
            self.progress.emit(1)
            tx = self._settings.get('tx', {})
            cct.set_txs(
                vhigh=tx.get('vhigh', ''),
                t_rise=tx.get('t_rise', ''),
                ui=tx.get('ui', ''),
                res_tx=tx.get('res_tx', ''),
                cap_tx=tx.get('cap_tx', ''),
            )

            self.message.emit('Configuring receive settings...')
            self.progress.emit(2)
            rx = self._settings.get('rx', {})
            cct.set_rxs(
                res_rx=rx.get('res_rx', ''),
                cap_rx=rx.get('cap_rx', ''),
            )

            if self._mode == 'prerun':
                self.message.emit('Running pre-run threshold analysis...')
                self.progress.emit(3)
                summaries = cct.pre_run()
                summary_text = self._summarize_prerun(summaries, threshold_value)
                self.progress.emit(4)
                self.finished.emit(summary_text)
                return

            self.message.emit('Running transient simulation...')
            self.progress.emit(3)
            run_params = self._settings.get('run', {})
            cct.run(
                tstep=run_params.get('tstep', ''),
                tstop=run_params.get('tstop', ''),
            )

            self.message.emit('Generating CCT report...')
            self.progress.emit(4)
            if self._output_path is None:
                raise RuntimeError('Output path not provided for CCT run')
            cct.calculate(output_path=str(self._output_path))
        except ImportError as exc:  # pragma: no cover - runtime feedback path
            self.failed.emit('dependency', exc)
            return
        except Exception as exc:  # pragma: no cover - runtime feedback path
            self.failed.emit('failure', exc)
            return

        self.message.emit(f"CCT results saved to {self._output_path}")
        self.finished.emit(str(self._output_path))

    @staticmethod
    def _summarize_prerun(summaries: List[Dict[str, object]], threshold_value: Optional[float]) -> str:
        if not summaries:
            if threshold_value is None:
                return 'Pre-run complete. No transmitters evaluated.'
            return f'Pre-run complete at threshold {threshold_value:.1f} dB. No transmitters evaluated.'

        lines: List[str] = []
        if threshold_value is None:
            lines.append('Pre-run complete. Using full network (no threshold applied).')
        else:
            lines.append(f'Pre-run complete at threshold {threshold_value:.1f} dB.')

        port_ratios: List[float] = []
        rx_ratios: List[float] = []
        for stats in summaries:
            total_ports = int(stats.get('total_port_count', 0) or 0)
            kept_ports = int(stats.get('kept_port_count', 0) or 0)
            port_ratio = (kept_ports / total_ports) if total_ports else 0.0
            port_ratios.append(port_ratio)

            total_rx = int(stats.get('total_rx_port_count', 0) or 0)
            kept_rx = int(stats.get('kept_rx_port_count', 0) or 0)
            rx_ratio = (kept_rx / total_rx) if total_rx else None
            if rx_ratio is not None:
                rx_ratios.append(rx_ratio)

            label = str(stats.get('tx_label', 'tx'))
            line = f"{label}: ports {kept_ports}/{total_ports}"
            if total_ports:
                line += f" ({port_ratio:.1%})"
            if total_rx:
                line += f", rx {kept_rx}/{total_rx} ({rx_ratio:.1%})"
            lines.append(line)

        if port_ratios:
            avg_port = sum(port_ratios) / len(port_ratios)
            lines.insert(1, f"Average kept ports: {avg_port:.1%}")
        if rx_ratios:
            avg_rx = sum(rx_ratios) / len(rx_ratios)
            insert_at = 2 if port_ratios else 1
            lines.insert(insert_at, f"Average kept RX ports: {avg_rx:.1%}")

        return "\n".join(lines)



class CheckableListWidget(QListWidget):
    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setSelectionBehavior(QAbstractItemView.SelectItems)
        try:
            self.setSelectionRectVisible(True)
        except AttributeError:
            pass

    def keyPressEvent(self, event) -> None:  # noqa: D401 - Qt signature
        if event.key() == Qt.Key_Space:
            items = self.selectedItems()
            if not items:
                current = self.currentItem()
                if current is not None:
                    items = [current]
            if items:
                for item in items:
                    if item.checkState() == Qt.Checked:
                        item.setCheckState(Qt.Unchecked)
                    else:
                        item.setCheckState(Qt.Checked)
            event.accept()
            return
        super().keyPressEvent(event)

DEFAULT_CCT_SETTINGS: Dict[str, float] = {
    "vhigh": 0.8,
    "t_rise": 30.0,
    "ui": 133.0,
    "res_tx": 40.0,
    "cap_tx": 1.0,
    "res_rx": 30.0,
    "cap_rx": 1.8,
    "tstep": 100.0,
    "tstop": 3.0,
    "threshold_db": -60.0,
}

DEFAULT_CCT_TEXT_SETTINGS: Dict[str, str] = {
    "circuit_version": DEFAULT_CIRCUIT_VERSION,
}

DEFAULT_CCT_ALL_SETTINGS: Dict[str, object] = {
    **DEFAULT_CCT_SETTINGS,
    **DEFAULT_CCT_TEXT_SETTINGS,
}

CCT_PARAM_GROUPS = {
    "tx": ["vhigh", "t_rise", "ui", "res_tx", "cap_tx"],
    "rx": ["res_rx", "cap_rx"],
    "transient": ["tstep", "tstop"],
    "options": ["circuit_version", "threshold_db"],
}

CCT_GROUP_ALIASES = {
    "prune": "options",
}



class EdbGui(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("AEDB CCT Calculator")
        self.resize(1000, 650)

        self._edb: Edb | None = None
        self._components: Dict[str, object] = {}
        self._component_nets: Dict[str, set[str]] = {}
        self._aedb_path: Optional[Path] = None
        self._aedb_source_path: Optional[Path] = None
        self._edb_version: Optional[str] = DEFAULT_EDB_VERSION
        self._cct_metadata: Optional[dict] = None
        self._cct_port_entries: List[dict] = []
        self._settings = QSettings(SETTINGS_ORG, SETTINGS_APP)
        self._cct_param_spins: Dict[str, QDoubleSpinBox] = {}
        self._cct_text_fields: Dict[str, QLineEdit] = {}
        self._active_cct_mode: Optional[str] = None
        self._cct_progress_steps = 4
        self.cutout_enable_checkbox: Optional[QCheckBox] = None
        self.cutout_expansion_spin: Optional[QDoubleSpinBox] = None
        self.sweep_table: Optional[QTableWidget] = None
        self.simulation_apply_button: Optional[QPushButton] = None
        self.sim_signal_label: Optional[QLabel] = None
        self.sim_reference_label: Optional[QLabel] = None
        self._loading_simulation_settings = False
        self._simulation_signal_nets: List[str] = []
        self._simulation_reference_net: Optional[str] = None
        self._cct_thread: Optional[QThread] = None
        self._cct_worker: Optional[_CctWorker] = None

        self._build_ui()
        self._restore_edb_version()
        self._restore_simulation_settings()
        self._restore_cct_settings()

    def _build_ui(self) -> None:
        main_layout = QVBoxLayout(self)

        header_layout = QHBoxLayout()

        header_layout.addWidget(QLabel("EDB version:"))
        self.edb_version_combo = QComboBox()
        self.edb_version_combo.setEditable(True)
        self.edb_version_combo.setInsertPolicy(QComboBox.NoInsert)
        for version in EDB_VERSION_OPTIONS:
            self.edb_version_combo.addItem(version)
        self.edb_version_combo.currentTextChanged.connect(self._on_edb_version_changed)
        line_edit = self.edb_version_combo.lineEdit()
        if line_edit is not None:
            line_edit.setPlaceholderText("Leave blank for auto-detect")
        header_layout.addWidget(self.edb_version_combo)

        self.open_button = QPushButton("Open .aedb?")
        self.open_button.clicked.connect(self._prompt_for_aedb)
        header_layout.addWidget(self.open_button)

        self.file_label = QLabel("No design loaded")
        self.file_label.setMinimumWidth(400)
        header_layout.addWidget(self.file_label, stretch=1)
        main_layout.addLayout(header_layout)

        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs, stretch=1)

        port_tab = QWidget()
        self._build_port_tab(port_tab)
        self.tabs.addTab(port_tab, "Port Setup")

        simulation_tab = QWidget()
        self._build_simulation_tab(simulation_tab)
        self.tabs.addTab(simulation_tab, "Simulation")

        cct_tab = QWidget()
        self._build_cct_tab(cct_tab)
        self.tabs.addTab(cct_tab, "CCT")

        self.status_output = QPlainTextEdit()
        self.status_output.setReadOnly(True)
        self.status_output.setTabChangesFocus(True)
        self.status_output.setLineWrapMode(QPlainTextEdit.WidgetWidth)
        self.status_output.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.status_output.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.status_output.setMaximumHeight(90)
        self.status_output.setPlainText("")
        self.status_output.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        main_layout.addWidget(self.status_output)

    def _build_port_tab(self, container: QWidget) -> None:
        layout = QVBoxLayout(container)

        selector_layout = QHBoxLayout()

        self.controller_group = QGroupBox("Controller Components")
        controller_layout = QVBoxLayout(self.controller_group)
        self.controller_list = QListWidget()
        self.controller_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.controller_list.itemSelectionChanged.connect(self._update_results)
        controller_layout.addWidget(self.controller_list)

        self.dram_group = QGroupBox("DRAM Components")
        dram_layout = QVBoxLayout(self.dram_group)
        self.dram_list = QListWidget()
        self.dram_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.dram_list.itemSelectionChanged.connect(self._update_results)
        dram_layout.addWidget(self.dram_list)

        selector_layout.addWidget(self.controller_group, stretch=1)
        selector_layout.addWidget(self.dram_group, stretch=1)
        layout.addLayout(selector_layout)

        reference_layout = QHBoxLayout()
        reference_layout.addWidget(QLabel("Reference net:"))
        self.reference_combo = QComboBox()
        self.reference_combo.setEnabled(False)
        self.reference_combo.setFixedWidth(220)
        self.reference_combo.addItem("Select reference net...")
        reference_layout.addWidget(self.reference_combo)

        reference_layout.addStretch(1)
        self.port_count_label = QLabel("Checked nets: 0 | Ports: 0")
        self.port_count_label.setMinimumWidth(200)
        self.port_count_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        reference_layout.addWidget(self.port_count_label)
        layout.addLayout(reference_layout)

        nets_layout = QHBoxLayout()

        self.single_group = QGroupBox("Single-Ended Nets")
        single_layout = QVBoxLayout(self.single_group)
        self.single_list = CheckableListWidget()
        self.single_list.itemChanged.connect(self._update_action_state)
        single_layout.addWidget(self.single_list)
        nets_layout.addWidget(self.single_group, stretch=1)

        self.diff_group = QGroupBox("Differential Pairs")
        diff_layout = QVBoxLayout(self.diff_group)
        self.diff_list = CheckableListWidget()
        self.diff_list.itemChanged.connect(self._update_action_state)
        diff_layout.addWidget(self.diff_list)
        nets_layout.addWidget(self.diff_group, stretch=1)

        layout.addLayout(nets_layout, stretch=1)

        actions_layout = QHBoxLayout()
        actions_layout.addStretch(1)
        self.apply_button = QPushButton("Apply")
        self.apply_button.setEnabled(False)
        self.apply_button.clicked.connect(self._apply_changes)
        actions_layout.addWidget(self.apply_button)
        layout.addLayout(actions_layout)

        self.reference_combo.currentIndexChanged.connect(self._update_action_state)

    def _build_simulation_tab(self, container: QWidget) -> None:
        layout = QVBoxLayout(container)

        cutout_group = QGroupBox("Cutout")
        cutout_form = QFormLayout(cutout_group)
        self.cutout_enable_checkbox = QCheckBox("Enable cutout")
        self.cutout_enable_checkbox.setChecked(True)
        self.cutout_enable_checkbox.toggled.connect(self._on_cutout_enabled_changed)
        cutout_form.addRow(self.cutout_enable_checkbox)

        self.cutout_expansion_spin = QDoubleSpinBox()
        self.cutout_expansion_spin.setRange(0.0, 1000.0)
        self.cutout_expansion_spin.setDecimals(6)
        self.cutout_expansion_spin.setSingleStep(0.001)
        self.cutout_expansion_spin.setValue(DEFAULT_CUTOUT_EXPANSION)
        self.cutout_expansion_spin.valueChanged.connect(self._persist_simulation_settings)
        cutout_form.addRow("Expansion size (m)", self.cutout_expansion_spin)

        self.sim_signal_label = QLabel("(not set)")
        self.sim_signal_label.setWordWrap(True)
        cutout_form.addRow("Signal nets", self.sim_signal_label)

        self.sim_reference_label = QLabel("(not set)")
        cutout_form.addRow("Reference net", self.sim_reference_label)

        layout.addWidget(cutout_group)

        sweeps_group = QGroupBox("Frequency Sweeps")
        sweeps_layout = QVBoxLayout(sweeps_group)

        self.sweep_table = QTableWidget(0, 4)
        self.sweep_table.setHorizontalHeaderLabels(["Sweep Type", "Start", "Stop", "Step/Count"])
        self.sweep_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.sweep_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.sweep_table.verticalHeader().setVisible(False)
        header = self.sweep_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.sweep_table.itemChanged.connect(self._on_sweep_cell_changed)
        sweeps_layout.addWidget(self.sweep_table, stretch=1)

        sweep_buttons = QHBoxLayout()
        add_button = QPushButton("Add Sweep")
        add_button.clicked.connect(lambda: self._add_sweep_row())
        sweep_buttons.addWidget(add_button)
        remove_button = QPushButton("Remove Selected")
        remove_button.clicked.connect(lambda: self._remove_selected_sweeps())
        sweep_buttons.addWidget(remove_button)
        sweep_buttons.addStretch(1)
        sweeps_layout.addLayout(sweep_buttons)

        layout.addWidget(sweeps_group, stretch=1)

        actions_layout = QHBoxLayout()
        actions_layout.addStretch(1)
        self.simulation_apply_button = QPushButton("Apply Simulation")
        self.simulation_apply_button.setEnabled(False)
        self.simulation_apply_button.clicked.connect(self._apply_simulation_settings)
        actions_layout.addWidget(self.simulation_apply_button)
        layout.addLayout(actions_layout)


    def _add_sweep_row(self, data: Optional[Iterable[object]] = None, persist: bool = True) -> None:
        table = self.sweep_table
        if table is None:
            return

        values = ["", "", "", ""]
        if data is not None:
            seq = list(data)
            for idx in range(min(4, len(seq))):
                values[idx] = "" if seq[idx] is None else str(seq[idx])

        previous = self._loading_simulation_settings
        self._loading_simulation_settings = True
        try:
            row = table.rowCount()
            table.insertRow(row)

            combo = QComboBox()
            for option in SWEEP_TYPE_OPTIONS:
                combo.addItem(option)
            sweep_type = values[0] or SWEEP_TYPE_OPTIONS[0]
            if sweep_type not in SWEEP_TYPE_OPTIONS:
                combo.addItem(sweep_type)
            combo.setCurrentText(sweep_type)
            combo.currentTextChanged.connect(self._on_sweep_combo_changed)
            table.setCellWidget(row, 0, combo)

            for column in range(1, 4):
                item = QTableWidgetItem(values[column])
                table.setItem(row, column, item)
        finally:
            self._loading_simulation_settings = previous

        if persist and not previous:
            self._persist_simulation_settings()
        self._update_simulation_ui_state()

    def _remove_selected_sweeps(self) -> None:
        table = self.sweep_table
        if table is None:
            return
        selection = table.selectionModel()
        if selection is None:
            return
        rows = sorted({index.row() for index in selection.selectedRows()}, reverse=True)
        if not rows:
            return

        previous = self._loading_simulation_settings
        self._loading_simulation_settings = True
        try:
            for row in rows:
                table.removeRow(row)
        finally:
            self._loading_simulation_settings = previous

        self._persist_simulation_settings()
        self._update_simulation_ui_state()

    def _on_sweep_cell_changed(self, item: QTableWidgetItem) -> None:
        if self._loading_simulation_settings:
            return
        self._persist_simulation_settings()

    def _on_sweep_combo_changed(self, _text: str) -> None:
        if self._loading_simulation_settings:
            return
        self._persist_simulation_settings()

    def _collect_sweep_rows(self) -> List[List[str]]:
        table = self.sweep_table
        if table is None:
            return []
        rows: List[List[str]] = []
        for row in range(table.rowCount()):
            widget = table.cellWidget(row, 0)
            sweep_type = widget.currentText().strip() if isinstance(widget, QComboBox) else ""
            start_item = table.item(row, 1)
            stop_item = table.item(row, 2)
            step_item = table.item(row, 3)
            start = start_item.text().strip() if start_item is not None else ""
            stop = stop_item.text().strip() if stop_item is not None else ""
            step = step_item.text().strip() if step_item is not None else ""
            rows.append([sweep_type, start, stop, step])
        return rows

    @staticmethod
    def _coerce_sweep_value(value: str) -> object:
        text = (value or "").strip()
        if not text:
            return text
        if re.fullmatch(r"[+-]?\d+", text):
            try:
                return int(text)
            except ValueError:
                return text
        if re.fullmatch(r"[+-]?(?:\d+\.\d*|\.\d+)(?:[eE][+-]?\d+)?", text):
            try:
                return float(text)
            except ValueError:
                return text
        return text

    def _normalized_sweep_rows(self) -> List[List[object]]:
        normalized: List[List[object]] = []
        for sweep_type, start, stop, step in self._collect_sweep_rows():
            if not sweep_type:
                continue
            normalized.append([sweep_type, start, stop, self._coerce_sweep_value(step)])
        return normalized

    def _persist_simulation_settings(self) -> None:
        if self._loading_simulation_settings or not getattr(self, '_settings', None):
            return
        expansion = DEFAULT_CUTOUT_EXPANSION
        if self.cutout_expansion_spin is not None:
            expansion = float(self.cutout_expansion_spin.value())
        cutout_enabled = True
        if self.cutout_enable_checkbox is not None:
            cutout_enabled = self.cutout_enable_checkbox.isChecked()
        sweeps = self._collect_sweep_rows()
        try:
            sweeps_payload = json.dumps(sweeps)
        except TypeError:
            sweeps_payload = json.dumps(DEFAULT_FREQUENCY_SWEEPS)
        self._settings.setValue('simulation/expansion_size', expansion)
        self._settings.setValue('simulation/cutout_enabled', cutout_enabled)
        self._settings.setValue('simulation/sweeps', sweeps_payload)
        self._settings.sync()
        self._update_simulation_ui_state()

    def _restore_simulation_settings(self) -> None:
        if not getattr(self, '_settings', None):
            return

        expansion_raw = self._settings.value('simulation/expansion_size', DEFAULT_CUTOUT_EXPANSION)
        try:
            expansion = float(expansion_raw)
        except (TypeError, ValueError):
            expansion = DEFAULT_CUTOUT_EXPANSION

        cutout_enabled_raw = self._settings.value('simulation/cutout_enabled', True)
        if isinstance(cutout_enabled_raw, bool):
            cutout_enabled = cutout_enabled_raw
        else:
            cutout_enabled = str(cutout_enabled_raw).strip().lower() not in ('false', '0', 'no', '')

        sweeps_raw = self._settings.value('simulation/sweeps', None)
        parsed_sweeps: List[List[str]] = []
        if sweeps_raw:
            try:
                data = json.loads(str(sweeps_raw))
                if isinstance(data, list):
                    for entry in data:
                        if not isinstance(entry, (list, tuple)) or len(entry) < 4:
                            continue
                        parsed_sweeps.append([
                            "" if entry[i] is None else str(entry[i]) if i < len(entry) else ""
                            for i in range(4)
                        ])
            except (TypeError, ValueError):
                parsed_sweeps = []
        if not parsed_sweeps:
            parsed_sweeps = DEFAULT_FREQUENCY_SWEEPS.copy()

        if self.cutout_enable_checkbox is not None:
            blocker_cb = self.cutout_enable_checkbox.blockSignals(True)
            self.cutout_enable_checkbox.setChecked(cutout_enabled)
            self.cutout_enable_checkbox.blockSignals(blocker_cb)

        if self.cutout_expansion_spin is not None:
            blocker_spin = self.cutout_expansion_spin.blockSignals(True)
            self.cutout_expansion_spin.setValue(expansion)
            self.cutout_expansion_spin.setEnabled(cutout_enabled)
            self.cutout_expansion_spin.blockSignals(blocker_spin)

        table = self.sweep_table
        if table is None:
            self._persist_simulation_settings()
            return

        previous = self._loading_simulation_settings
        self._loading_simulation_settings = True
        try:
            table.setRowCount(0)
            for row_data in parsed_sweeps:
                self._add_sweep_row(row_data, persist=False)
        finally:
            self._loading_simulation_settings = previous

        if table.rowCount() == 0:
            previous = self._loading_simulation_settings
            self._loading_simulation_settings = True
            try:
                for row_data in DEFAULT_FREQUENCY_SWEEPS:
                    self._add_sweep_row(row_data, persist=False)
            finally:
                self._loading_simulation_settings = previous

        self._persist_simulation_settings()

    def _update_simulation_ui_state(self) -> None:
        button = getattr(self, 'simulation_apply_button', None)
        if button is None:
            return
        has_edb = self._edb is not None
        has_sweeps = bool(self._normalized_sweep_rows())
        cutout_enabled = True
        if self.cutout_enable_checkbox is not None:
            cutout_enabled = self.cutout_enable_checkbox.isChecked()
        has_nets = bool(self._simulation_signal_nets) if cutout_enabled else True
        has_reference = self._simulation_reference_net is not None if cutout_enabled else True
        button.setEnabled(has_edb and has_sweeps and has_nets and has_reference)

    def _on_cutout_enabled_changed(self, enabled: bool) -> None:
        if self.cutout_expansion_spin is not None:
            self.cutout_expansion_spin.setEnabled(enabled)
        self._persist_simulation_settings()
        self._update_simulation_ui_state()

    def _set_simulation_sources(self, signal_nets: Iterable[str], reference_net: Optional[str]) -> None:
        unique: List[str] = []
        seen: set[str] = set()
        for net in signal_nets:
            cleaned = str(net).strip() if net is not None else ''
            if not cleaned or cleaned in seen:
                continue
            seen.add(cleaned)
            unique.append(cleaned)

        self._simulation_signal_nets = unique
        self._simulation_reference_net = reference_net or None

        if self.sim_signal_label is not None:
            self.sim_signal_label.setText(', '.join(unique) if unique else '(not set)')
        if self.sim_reference_label is not None:
            self.sim_reference_label.setText(reference_net or '(not set)')

        self._update_simulation_ui_state()

    def _apply_simulation_settings(self) -> None:
        if self._edb is None:
            QMessageBox.warning(self, 'No design', 'Load an AEDB design before applying simulation settings.')
            return

        cutout_enabled = True
        if self.cutout_enable_checkbox is not None:
            cutout_enabled = self.cutout_enable_checkbox.isChecked()

        signal_nets = list(self._simulation_signal_nets)
        reference_net = self._simulation_reference_net

        if cutout_enabled:
            if not signal_nets:
                QMessageBox.warning(self, 'No nets selected', 'Run Apply on the Port Setup tab to choose the nets to include in the cutout.')
                return
            if reference_net is None:
                QMessageBox.warning(self, 'No reference net', 'Run Apply on the Port Setup tab to choose a reference net for the cutout.')
                return

        sweeps = self._normalized_sweep_rows()
        if not sweeps:
            QMessageBox.warning(self, 'No sweeps', 'Add at least one frequency sweep before applying simulation settings.')
            return

        expansion = DEFAULT_CUTOUT_EXPANSION
        if self.cutout_expansion_spin is not None:
            expansion = float(self.cutout_expansion_spin.value())

        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            if cutout_enabled:
                self._edb.cutout(signal_nets, [reference_net], expansion_size=expansion, extent_type='Bounding')

            setup = self._edb.create_siwave_syz_setup()
            setup.add_frequency_sweep('mysetup', sweeps)
            saved_path = self._save_modified_design()
        except Exception as exc:
            self._show_error('Failed to apply simulation settings', exc)
            return
        finally:
            try:
                QApplication.restoreOverrideCursor()
            except Exception:
                pass

        if saved_path is None:
            QMessageBox.information(self, 'Simulation applied', 'Simulation settings applied, but the design was not saved because no target path was available.')
            status_msg = 'Simulation settings applied without saving a new AEDB copy.'
        else:
            QMessageBox.information(self, 'Simulation applied', f'Simulation setup created and saved to {saved_path}')
            status_msg = f'Simulation setup saved to {saved_path}'
        if not cutout_enabled:
            status_msg += ' (cutout skipped)'
        self._set_status_message(status_msg)
        self._update_simulation_ui_state()

    def _set_edb_version_from_text(self, text: str) -> None:
        value = (text or '').strip()
        if not value or value.lower() == 'none':
            self._edb_version = None
        else:
            self._edb_version = value

    def _on_edb_version_changed(self, text: str) -> None:
        self._set_edb_version_from_text(text)
        self._persist_edb_version()

    def _persist_edb_version(self) -> None:
        if not getattr(self, '_settings', None):
            return
        combo = getattr(self, 'edb_version_combo', None)
        raw_text = ''
        if combo is not None:
            raw_text = combo.currentText().strip()
        self._settings.setValue('general/edb_version', raw_text)
        self._settings.sync()

    def _restore_edb_version(self) -> None:
        if not getattr(self, '_settings', None):
            return
        raw = self._settings.value('general/edb_version', DEFAULT_EDB_VERSION)
        if raw is None:
            text_value = DEFAULT_EDB_VERSION
        else:
            text_value = str(raw).strip()
        combo = getattr(self, 'edb_version_combo', None)
        if combo is not None:
            was_blocked = combo.blockSignals(True)
            combo.setEditText(text_value)
            combo.blockSignals(was_blocked)
        self._set_edb_version_from_text(text_value)


    def _set_status_message(self, message: str) -> None:
        widget = getattr(self, 'status_output', None)
        if widget is None:
            return
        widget.setPlainText(message or '')
        widget.moveCursor(QTextCursor.End)
        widget.ensureCursorVisible()


    def _build_cct_tab(self, container: QWidget) -> None:
        layout = QVBoxLayout(container)

        file_layout = QVBoxLayout()

        touchstone_layout = QHBoxLayout()
        touchstone_layout.addWidget(QLabel("Touchstone (.sNp):"))
        self.cct_touchstone_edit = QLineEdit()
        self.cct_touchstone_edit.textChanged.connect(self._update_cct_ui_state)
        touchstone_layout.addWidget(self.cct_touchstone_edit, stretch=1)
        browse_touchstone = QPushButton("Browse")
        browse_touchstone.clicked.connect(self._browse_touchstone)
        touchstone_layout.addWidget(browse_touchstone)
        file_layout.addLayout(touchstone_layout)

        json_layout = QHBoxLayout()
        json_layout.addWidget(QLabel("Port metadata (.json):"))
        self.cct_json_edit = QLineEdit()
        self.cct_json_edit.textChanged.connect(self._update_cct_ui_state)
        json_layout.addWidget(self.cct_json_edit, stretch=1)
        browse_json = QPushButton("Browse")
        browse_json.clicked.connect(self._browse_metadata)
        json_layout.addWidget(browse_json)
        file_layout.addLayout(json_layout)

        layout.addLayout(file_layout)

        params_row = QHBoxLayout()

        def _make_group(title: str) -> QFormLayout:
            group = QGroupBox(title)
            form = QFormLayout(group)
            params_row.addWidget(group)
            return form

        tx_form = _make_group("TX Settings")
        rx_form = _make_group("RX Settings")
        transient_form = _make_group("Transient Settings")
        option_form = _make_group("Options")

        def _add_param(
            target_form: QFormLayout,
            name: str,
            label: str,
            suffix: str,
            minimum: float,
            maximum: float,
            step: float,
            decimals: int,
        ) -> None:
            spin = QDoubleSpinBox()
            spin.setRange(minimum, maximum)
            spin.setDecimals(decimals)
            spin.setSingleStep(step)
            if suffix:
                spin.setSuffix(f" {suffix}")
            default_value = DEFAULT_CCT_SETTINGS.get(name, 0.0)
            spin.setValue(default_value)
            spin.valueChanged.connect(self._persist_cct_settings)
            target_form.addRow(label, spin)
            self._cct_param_spins[name] = spin

        _add_param(tx_form, "vhigh", "TX Vhigh", "V", 0.0, 10.0, 0.05, 3)
        _add_param(tx_form, "t_rise", "TX Rise Time", "ps", 0.0, 5000.0, 1.0, 3)
        _add_param(tx_form, "ui", "Unit Interval", "ps", 0.0, 5000.0, 1.0, 3)
        _add_param(tx_form, "res_tx", "TX Resistance", "ohm", 0.0, 1000.0, 1.0, 3)
        _add_param(tx_form, "cap_tx", "TX Capacitance", "pF", 0.0, 100.0, 0.1, 3)

        _add_param(rx_form, "res_rx", "RX Resistance", "ohm", 0.0, 1000.0, 1.0, 3)
        _add_param(rx_form, "cap_rx", "RX Capacitance", "pF", 0.0, 100.0, 0.1, 3)

        version_edit = QLineEdit()
        default_version = DEFAULT_CCT_TEXT_SETTINGS.get("circuit_version", "")
        version_edit.setText(default_version)
        version_edit.setPlaceholderText("e.g. 2025.1")
        version_edit.editingFinished.connect(self._persist_cct_settings)
        option_form.addRow("AEDT Version", version_edit)
        self._cct_text_fields["circuit_version"] = version_edit

        _add_param(transient_form, "tstep", "Transient Step", "ps", 0.0, 1_000_000.0, 10.0, 3)
        _add_param(transient_form, "tstop", "Transient Stop", "ns", 0.0, 1_000_000.0, 0.1, 3)
        _add_param(option_form, "threshold_db", "Threshold", "dB", -200.0, 0.0, 1.0, 1)

        params_row.addStretch(1)

        button_container = QWidget()
        button_column = QVBoxLayout(button_container)
        button_column.setContentsMargins(0, 0, 0, 0)
        button_column.addStretch(1)
        self.cct_save_button = QPushButton("Save Config")
        self.cct_save_button.clicked.connect(self._save_cct_config)
        button_column.addWidget(self.cct_save_button)
        self.cct_load_button = QPushButton("Load Config")
        self.cct_load_button.clicked.connect(self._load_cct_config)
        button_column.addWidget(self.cct_load_button)
        self.cct_reset_button = QPushButton("Reset Defaults")
        self.cct_reset_button.clicked.connect(self._reset_cct_config)
        button_column.addWidget(self.cct_reset_button)

        params_row.addWidget(button_container, alignment=Qt.AlignBottom | Qt.AlignRight)
        layout.addLayout(params_row)

        self.cct_table = QTableWidget(0, 4)
        self.cct_table.setHorizontalHeaderLabels(["TX Port", "RX Port", "Type", "Pair"])
        header = self.cct_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        layout.addWidget(self.cct_table, stretch=1)

        cct_actions = QHBoxLayout()
        cct_actions.addStretch(1)
        self.cct_prerun_button = QPushButton("Pre-run")
        self.cct_prerun_button.setEnabled(False)
        self.cct_prerun_button.clicked.connect(self._run_cct_prerun)
        cct_actions.addWidget(self.cct_prerun_button)
        self.cct_calculate_button = QPushButton("Calculate")
        self.cct_calculate_button.setEnabled(False)
        self.cct_calculate_button.clicked.connect(self._run_cct_calculation)
        cct_actions.addWidget(self.cct_calculate_button)
        layout.addLayout(cct_actions)

        self.cct_progress = QProgressBar()
        self.cct_progress.setMinimum(0)
        self.cct_progress.setMaximum(0)
        self.cct_progress.setVisible(False)
        self.cct_progress.setTextVisible(False)
        layout.addWidget(self.cct_progress)

        self._update_cct_ui_state()

    def _browse_touchstone(self) -> None:
        start_dir = self._aedb_path.parent if self._aedb_path else ROOT_DIR
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Touchstone File",
            str(start_dir),
            "Touchstone (*.s*p *.S*p);;All Files (*.*)",
        )
        if file_path:
            self.cct_touchstone_edit.setText(file_path)
            self._set_status_message(f"Touchstone set to {file_path}")

    def _browse_metadata(self) -> None:
        start_dir = ROOT_DIR
        if self._aedb_path:
            start_dir = self._aedb_path.parent
        elif self._aedb_source_path:
            start_dir = self._aedb_source_path.parent

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Port Metadata JSON",
            str(start_dir),
            "JSON Files (*.json);;All Files (*.*)",
        )
        if file_path:
            self.cct_json_edit.setText(file_path)
            self._auto_load_cct_metadata()

    def _load_cct_metadata(self) -> None:
        if load_port_metadata is None:
            QMessageBox.warning(
                self,
                "Unavailable",
                "Port metadata loading requires the CCT module. Ensure dependencies are installed.",
            )
            return

        path = self.cct_json_edit.text().strip()
        if not path:
            QMessageBox.warning(self, "Metadata path missing", "Enter a JSON metadata path first.")
            return

        metadata_path = Path(path)
        if not metadata_path.exists():
            QMessageBox.warning(
                self,
                "Metadata not found",
                f"File does not exist:\n{metadata_path}",
            )
            return

        try:
            entries, raw = load_port_metadata(metadata_path)
        except Exception as exc:  # pragma: no cover - runtime feedback path
            self._show_error("Failed to load metadata", exc)
            return

        self._cct_metadata = {"entries": entries, "raw": raw, "path": metadata_path}
        self._populate_cct_table(entries)
        self._set_status_message(
            f"Loaded {len(entries)} ports from {metadata_path.name}"
        )
        self._update_cct_ui_state()

    def _auto_load_cct_metadata(self) -> None:
        path_str = self.cct_json_edit.text().strip()
        if not path_str:
            if self._cct_metadata is not None or self._cct_port_entries:
                self._cct_metadata = None
                self._cct_port_entries = []
                self.cct_table.setRowCount(0)
            return

        if load_port_metadata is None:
            return

        metadata_path = Path(path_str)
        if not metadata_path.exists():
            return

        current_path = self._cct_metadata.get("path") if self._cct_metadata else None
        if current_path and Path(current_path) == metadata_path:
            return

        try:
            entries, raw = load_port_metadata(metadata_path)
        except Exception as exc:  # pragma: no cover - runtime feedback path
            self._show_error("Failed to load metadata", exc)
            return

        self._cct_metadata = {"entries": entries, "raw": raw, "path": metadata_path}
        self._populate_cct_table(entries)
        self._set_status_message(
            f"Loaded {len(entries)} ports from {metadata_path.name}"
        )

    def _populate_cct_table(self, entries: Iterable[object]) -> None:
        rows = self._build_cct_rows(entries)
        self._cct_port_entries = list(entries)
        self.cct_table.setRowCount(len(rows))

        for row_index, row in enumerate(rows):
            tx_item = QTableWidgetItem(row["tx_display"])
            rx_item = QTableWidgetItem(row["rx_display"])
            type_item = QTableWidgetItem(row["type"])
            pair_item = QTableWidgetItem(row.get("pair", ""))

            for item in (tx_item, rx_item, type_item, pair_item):
                item.setFlags(item.flags() ^ Qt.ItemIsEditable)

            self.cct_table.setItem(row_index, 0, tx_item)
            self.cct_table.setItem(row_index, 1, rx_item)
            self.cct_table.setItem(row_index, 2, type_item)
            self.cct_table.setItem(row_index, 3, pair_item)

            if row.get("is_diff") and DIFF_ROW_BRUSH is not None:
                for column in range(4):
                    self.cct_table.item(row_index, column).setBackground(DIFF_ROW_BRUSH)

        if not rows:
            self.cct_table.clearContents()

    def _build_cct_rows(self, entries: Iterable[object]) -> List[Dict[str, object]]:
        singles_ctrl: Dict[str, List[object]] = {}
        singles_dram: Dict[str, List[object]] = {}
        diff_ctrl: Dict[Tuple[str, str], Dict[str, object]] = {}
        diff_dram: Dict[Tuple[str, str], Dict[str, object]] = {}

        for entry in entries:
            role = getattr(entry, "component_role", "")
            net_type = getattr(entry, "net_type", "single")
            net_name = getattr(entry, "net", "")

            if net_type == "differential":
                pair_label = getattr(entry, "pair", None) or net_name
                key = (getattr(entry, "component", ""), pair_label)
                target = diff_ctrl if role == "controller" else diff_dram if role == "dram" else None
                if target is None:
                    continue
                mapping = target.setdefault(key, {})
                polarity = getattr(entry, "polarity", None) or (
                    "positive" if "positive" not in mapping else "negative"
                )
                mapping[polarity] = entry
            else:
                target_single = singles_ctrl if role == "controller" else singles_dram if role == "dram" else None
                if target_single is None:
                    continue
                target_single.setdefault(net_name, []).append(entry)

        rows: List[Dict[str, object]] = []

        all_single_nets = sorted(set(list(singles_ctrl.keys()) + list(singles_dram.keys())))
        for net in all_single_nets:
            ctrl_entries = singles_ctrl.get(net, [])
            dram_entries = singles_dram.get(net, [])
            if ctrl_entries and dram_entries:
                for ctrl in ctrl_entries:
                    for dram in dram_entries:
                        rows.append(
                            {
                                "type": "Single",
                                "tx_display": getattr(ctrl, "name", ""),
                                "rx_display": getattr(dram, "name", ""),
                                "pair": net,
                                "is_diff": False,
                            }
                        )
            elif ctrl_entries:
                for ctrl in ctrl_entries:
                    rows.append(
                        {
                            "type": "Single",
                            "tx_display": getattr(ctrl, "name", ""),
                            "rx_display": "(none)",
                            "pair": net,
                            "is_diff": False,
                        }
                    )
            else:
                for dram in dram_entries:
                    rows.append(
                        {
                            "type": "Single",
                            "tx_display": "(none)",
                            "rx_display": getattr(dram, "name", ""),
                            "pair": net,
                            "is_diff": False,
                        }
                    )

        def build_diff_map(source: Dict[Tuple[str, str], Dict[str, object]]) -> Dict[Tuple[str, str], List[Dict[str, object]]]:
            result: Dict[Tuple[str, str], List[Dict[str, object]]] = {}
            for (_, pair_label), mapping in source.items():
                pos_entry = mapping.get("positive")
                neg_entry = mapping.get("negative")
                if not pos_entry or not neg_entry:
                    continue
                signature = tuple(sorted([getattr(pos_entry, "net", ""), getattr(neg_entry, "net", "")]))
                label = getattr(pos_entry, "pair", None) or getattr(neg_entry, "pair", None)
                if not label:
                    label = f"{getattr(pos_entry, 'name', '')}/{getattr(neg_entry, 'name', '')}"
                result.setdefault(signature, []).append(
                    {
                        "label": label,
                        "positive": pos_entry,
                        "negative": neg_entry,
                    }
                )
            return result

        ctrl_pairs = build_diff_map(diff_ctrl)
        dram_pairs = build_diff_map(diff_dram)
        all_signatures = sorted(set(ctrl_pairs.keys()) | set(dram_pairs.keys()))

        for signature in all_signatures:
            ctrl_list = ctrl_pairs.get(signature, [])
            dram_list = dram_pairs.get(signature, [])
            pair_label = ctrl_list[0]["label"] if ctrl_list else (dram_list[0]["label"] if dram_list else "/".join(signature))

            if ctrl_list and dram_list:
                for ctrl in ctrl_list:
                    for dram in dram_list:
                        rows.append(
                            {
                                "type": "Differential",
                                "tx_display": f"{getattr(ctrl['positive'], 'name', '')} / {getattr(ctrl['negative'], 'name', '')}",
                                "rx_display": f"{getattr(dram['positive'], 'name', '')} / {getattr(dram['negative'], 'name', '')}",
                                "pair": pair_label,
                                "is_diff": True,
                            }
                        )
            elif ctrl_list:
                for ctrl in ctrl_list:
                    rows.append(
                        {
                            "type": "Differential",
                            "tx_display": f"{getattr(ctrl['positive'], 'name', '')} / {getattr(ctrl['negative'], 'name', '')}",
                            "rx_display": "(none)",
                            "pair": pair_label,
                            "is_diff": True,
                        }
                    )
            else:
                for dram in dram_list:
                    rows.append(
                        {
                            "type": "Differential",
                            "tx_display": "(none)",
                            "rx_display": f"{getattr(dram['positive'], 'name', '')} / {getattr(dram['negative'], 'name', '')}",
                            "pair": pair_label,
                            "is_diff": True,
                        }
                    )

        rows.sort(key=lambda row: (row["type"], row.get("pair", ""), row["tx_display"], row["rx_display"]))
        return rows

    def _persist_cct_settings(self) -> None:
        if not getattr(self, "_settings", None):
            return
        for key, spin in self._cct_param_spins.items():
            self._settings.setValue(f"cct/{key}", spin.value())
        for key, field in self._cct_text_fields.items():
            self._settings.setValue(f"cct/{key}", field.text().strip())
        self._settings.sync()

    def _restore_cct_settings(self) -> None:
        if not getattr(self, "_settings", None):
            return
        stored: Dict[str, object] = {}
        for key, default in DEFAULT_CCT_SETTINGS.items():
            value = self._settings.value(f"cct/{key}", default)
            try:
                stored[key] = float(value)
            except (TypeError, ValueError):
                stored[key] = default
        for key, default in DEFAULT_CCT_TEXT_SETTINGS.items():
            value = self._settings.value(f"cct/{key}", default)
            if value is None:
                stored[key] = default
            else:
                stored[key] = str(value)
        self._apply_cct_values(stored, persist=False)

    def _current_cct_settings(self) -> Dict[str, object]:
        values: Dict[str, object] = {}
        for key, default in DEFAULT_CCT_SETTINGS.items():
            spin = self._cct_param_spins.get(key)
            values[key] = float(spin.value()) if spin is not None else default
        for key, default in DEFAULT_CCT_TEXT_SETTINGS.items():
            field = self._cct_text_fields.get(key)
            if field is None:
                values[key] = default
                continue
            text = field.text().strip()
            values[key] = text if text else default
        return values

    @staticmethod
    def _format_with_unit(value: float, unit: str) -> str:
        return f"{value:g}{unit}"

    def _apply_cct_values(self, values: Dict[str, object], persist: bool = True) -> None:
        for key, raw_value in values.items():
            spin = self._cct_param_spins.get(key)
            if spin is not None:
                default_numeric = DEFAULT_CCT_SETTINGS.get(key, 0.0)
                try:
                    numeric = float(raw_value)
                except (TypeError, ValueError):
                    numeric = default_numeric
                was_blocked = spin.blockSignals(True)
                spin.setValue(numeric)
                spin.blockSignals(was_blocked)
                continue

            field = self._cct_text_fields.get(key)
            if field is not None:
                default_text = DEFAULT_CCT_TEXT_SETTINGS.get(key, "")
                if raw_value is None:
                    text_value = default_text
                else:
                    text_value = str(raw_value).strip()
                    if not text_value:
                        text_value = default_text
                was_blocked = field.blockSignals(True)
                field.setText(text_value)
                field.blockSignals(was_blocked)
        if persist:
            self._persist_cct_settings()

    def _save_cct_config(self) -> None:
        params = self._current_cct_settings()
        grouped = {
            group: {key: params[key] for key in keys if key in params}
            for group, keys in CCT_PARAM_GROUPS.items()
        }

        start_dir = self._aedb_path.parent if self._aedb_path else ROOT_DIR
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Save CCT Parameters",
            str(start_dir),
            "JSON Files (*.json);;All Files (*.*)",
        )
        if not filename:
            return

        try:
            with open(filename, "w", encoding="utf-8") as handle:
                json.dump(grouped, handle, indent=2)
        except Exception as exc:
            self._show_error("Failed to save CCT config", exc)
            return

        self._set_status_message(f"Saved CCT config to {filename}")

    def _load_cct_config(self) -> None:
        start_dir = self._aedb_path.parent if self._aedb_path else ROOT_DIR
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Load CCT Parameters",
            str(start_dir),
            "JSON Files (*.json);;All Files (*.*)",
        )
        if not filename:
            return

        try:
            with open(filename, "r", encoding="utf-8") as handle:
                data = json.load(handle)
        except Exception as exc:
            self._show_error("Failed to load CCT config", exc)
            return

        if not isinstance(data, dict):
            QMessageBox.warning(self, "Invalid config", "Selected JSON does not contain an object.")
            return

        extracted: Dict[str, object] = {}
        groups_to_check = dict(CCT_PARAM_GROUPS)
        for alias, target in CCT_GROUP_ALIASES.items():
            if alias not in groups_to_check and target in CCT_PARAM_GROUPS:
                groups_to_check[alias] = CCT_PARAM_GROUPS[target]

        if any(group in data for group in groups_to_check):
            for group, keys in groups_to_check.items():
                section = data.get(group, {})
                if not isinstance(section, dict):
                    continue
                for key in keys:
                    if key in section:
                        extracted[key] = section[key]
        else:
            allowed_keys = set(DEFAULT_CCT_SETTINGS) | set(DEFAULT_CCT_TEXT_SETTINGS)
            for key in allowed_keys:
                if key in data:
                    extracted[key] = data[key]

        if not extracted:
            QMessageBox.warning(self, "No parameters", "No recognized CCT parameters were found in the file.")
            return

        self._apply_cct_values(extracted, persist=True)
        self._set_status_message(f"Loaded CCT config from {filename}")

    def _reset_cct_config(self) -> None:
        self._apply_cct_values(DEFAULT_CCT_ALL_SETTINGS, persist=True)
        self._set_status_message("CCT parameters reset to defaults")

    def _update_cct_ui_state(self) -> None:
        button = getattr(self, "cct_calculate_button", None)
        if button is None:
            return

        self._auto_load_cct_metadata()

        touchstone_path = self.cct_touchstone_edit.text().strip()
        metadata_path = self.cct_json_edit.text().strip()
        metadata_loaded = bool(self._cct_metadata)
        backend_ready = CCT is not None and load_port_metadata is not None
        thread_running = self._cct_thread is not None and self._cct_thread.isRunning()
        enabled = (
            bool(touchstone_path)
            and bool(metadata_path)
            and metadata_loaded
            and backend_ready
            and not thread_running
        )
        button.setEnabled(enabled)
        prerun_button = getattr(self, "cct_prerun_button", None)
        if prerun_button is not None:
            prerun_button.setEnabled(enabled)

    def _validate_cct_environment(self) -> bool:
        if CCT is None:
            QMessageBox.warning(
                self,
                "Unavailable",
                "CCT backend is not available. Ensure required dependencies are installed.",
            )
            return False
        if self._cct_thread is not None and self._cct_thread.isRunning():
            QMessageBox.information(
                self,
                "CCT in progress",
                "A CCT calculation is already running. Please wait for it to finish before starting a new one.",
            )
            return False
        return True

    def _collect_cct_paths(self) -> Optional[Tuple[Path, Path]]:
        touchstone_path = Path(self.cct_touchstone_edit.text().strip())
        metadata_path = Path(self.cct_json_edit.text().strip())
        if not touchstone_path.exists():
            QMessageBox.warning(
                self,
                "Touchstone missing",
                f"Touchstone file not found:\n{touchstone_path}",
            )
            return None
        if not metadata_path.exists():
            QMessageBox.warning(
                self,
                "Metadata missing",
                f"Port metadata file not found:\n{metadata_path}",
            )
            return None
        return touchstone_path, metadata_path

    def _build_cct_settings_payload(self, params: Dict[str, object]) -> Dict[str, Dict[str, object]]:
        version_value = params.get("circuit_version", "")
        if version_value is None:
            version_str = DEFAULT_CCT_TEXT_SETTINGS.get("circuit_version", "")
        else:
            version_str = str(version_value).strip() or DEFAULT_CCT_TEXT_SETTINGS.get("circuit_version", "")

        return {
            "tx": {
                "vhigh": self._format_with_unit(params.get("vhigh", 0.0), "V"),
                "t_rise": self._format_with_unit(params.get("t_rise", 0.0), "ps"),
                "ui": self._format_with_unit(params.get("ui", 0.0), "ps"),
                "res_tx": self._format_with_unit(params.get("res_tx", 0.0), "ohm"),
                "cap_tx": self._format_with_unit(params.get("cap_tx", 0.0), "pF"),
            },
            "rx": {
                "res_rx": self._format_with_unit(params.get("res_rx", 0.0), "ohm"),
                "cap_rx": self._format_with_unit(params.get("cap_rx", 0.0), "pF"),
            },
            "run": {
                "tstep": self._format_with_unit(params.get("tstep", 0.0), "ps"),
                "tstop": self._format_with_unit(params.get("tstop", 0.0), "ns"),
            },
            "options": {
                "threshold_db": params.get("threshold_db"),
                "circuit_version": version_str,
            },
        }

    def _start_cct_worker(
        self,
        *,
        mode: str,
        touchstone_path: Path,
        metadata_path: Path,
        output_path: Optional[Path],
        workdir: Path,
        settings_payload: Dict[str, Dict[str, object]],
    ) -> None:
        QApplication.setOverrideCursor(Qt.WaitCursor)
        self.cct_progress.setVisible(True)
        self.cct_progress.setMinimum(0)
        self.cct_progress.setMaximum(0)
        self.cct_progress.setValue(0)
        self.cct_progress.setTextVisible(True)
        if mode == "prerun":
            self.cct_progress.setFormat("Pre-run...")
            self._set_status_message("Starting pre-run threshold check...")
        else:
            self.cct_progress.setFormat("Working...")
            self._set_status_message("Starting CCT calculation...")
        self._active_cct_mode = mode
        self._cct_progress_steps = 4
        self._update_cct_ui_state()

        worker = _CctWorker(
            touchstone_path=touchstone_path,
            metadata_path=metadata_path,
            output_path=output_path,
            workdir=workdir,
            settings=settings_payload,
            mode=mode,
        )
        thread = QThread(self)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.progress.connect(self._on_cct_progress)
        worker.message.connect(self._set_status_message)
        worker.finished.connect(self._on_cct_finished)
        worker.failed.connect(self._on_cct_failed)
        worker.finished.connect(thread.quit)
        worker.failed.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        worker.failed.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        thread.finished.connect(self._cleanup_cct_thread)

        self._cct_worker = worker
        self._cct_thread = thread
        thread.start()
        self._update_cct_ui_state()

    def _run_cct_calculation(self) -> None:
        if not self._validate_cct_environment():
            return
        paths = self._collect_cct_paths()
        if not paths:
            return
        touchstone_path, metadata_path = paths
        params = self._current_cct_settings()
        settings_payload = self._build_cct_settings_payload(params)
        output_path = metadata_path.with_name(f"{metadata_path.stem}_cct.csv")
        workdir = metadata_path.parent / "cct_work"
        self._start_cct_worker(
            mode="run",
            touchstone_path=touchstone_path,
            metadata_path=metadata_path,
            output_path=output_path,
            workdir=workdir,
            settings_payload=settings_payload,
        )

    def _run_cct_prerun(self) -> None:
        if not self._validate_cct_environment():
            return
        paths = self._collect_cct_paths()
        if not paths:
            return
        touchstone_path, metadata_path = paths
        params = self._current_cct_settings()
        settings_payload = self._build_cct_settings_payload(params)
        workdir = metadata_path.parent / "cct_work"
        self._start_cct_worker(
            mode="prerun",
            touchstone_path=touchstone_path,
            metadata_path=metadata_path,
            output_path=None,
            workdir=workdir,
            settings_payload=settings_payload,
        )

    def _on_cct_progress(self, step: int) -> None:
        if self.cct_progress.maximum() == 0:
            steps = max(1, getattr(self, "_cct_progress_steps", 4))
            self.cct_progress.setMaximum(steps)
            self.cct_progress.setValue(0)
            self.cct_progress.setFormat('Step %v / %m')
        clamped = max(0, min(step, self.cct_progress.maximum()))
        self.cct_progress.setValue(clamped)

    def _on_cct_finished(self, payload: str) -> None:
        QApplication.restoreOverrideCursor()
        self._finalize_cct_feedback()
        mode = self._active_cct_mode or "run"
        if mode == "prerun":
            summary = payload.strip() if payload else "Pre-run complete."
            QMessageBox.information(
                self,
                'Pre-run complete',
                summary,
            )
            self._set_status_message('Pre-run complete')
        else:
            path = Path(payload)
            QMessageBox.information(
                self,
                'CCT complete',
                f'CCT calculation finished.\nResults saved to {path}',
            )
            self._set_status_message(f'CCT results saved to {path}')
        self._active_cct_mode = None
        self._update_cct_ui_state()

    def _on_cct_failed(self, kind: str, exc: object) -> None:
        QApplication.restoreOverrideCursor()
        self._finalize_cct_feedback()
        error = exc if isinstance(exc, Exception) else Exception(str(exc))
        mode = self._active_cct_mode or "run"
        if kind == 'dependency':
            self._show_error('CCT dependency missing', error)
            self._set_status_message('CCT dependencies not available')
        else:
            title = 'Pre-run failed' if mode == 'prerun' else 'CCT calculation failed'
            self._show_error(title, error)
            self._set_status_message(title)
        self._active_cct_mode = None
        self._update_cct_ui_state()

    def _finalize_cct_feedback(self) -> None:
        self.cct_progress.setVisible(False)
        self.cct_progress.setTextVisible(False)
        self.cct_progress.setFormat('%p%')
        self.cct_progress.setMinimum(0)
        self.cct_progress.setMaximum(0)
        self.cct_progress.setValue(0)

    def _cleanup_cct_thread(self) -> None:
        self._cct_thread = None
        self._cct_worker = None
        self._active_cct_mode = None
        self._update_cct_ui_state()

    def _prompt_for_aedb(self) -> None:
        start_dir = str(Path.cwd())
        dialog = QFileDialog(self, "Select AEDB Directory")
        dialog.setFileMode(QFileDialog.Directory)
        dialog.setOption(QFileDialog.ShowDirsOnly, True)
        dialog.setDirectory(start_dir)
        if dialog.exec():
            selected = dialog.selectedFiles()
            if selected:
                self._load_aedb(Path(selected[0]))

    def _load_aedb(self, path: Path) -> None:
        try:
            if not path.exists():
                raise FileNotFoundError(f"AEDB path does not exist: {path}")
            self._close_edb()
            self._set_status_message(f"Loading AEDB using {QT_LIB}?")
            QApplication.processEvents()
            self._edb = Edb(str(path), edbversion=self._edb_version)
            self._aedb_path = path
            self._aedb_source_path = path
            self.file_label.setText(str(path))
            self._populate_components()
            self._update_simulation_ui_state()
            self._set_status_message(
                f"Loaded design with {len(self._components)} components (Uxxx)"
            )
        except Exception as exc:  # pragma: no cover - GUI feedback path
            self._show_error("Failed to load AEDB", exc)
            self._set_status_message("Failed to load design")
            self._close_edb()

    def _close_edb(self) -> None:
        if self._edb is not None:
            try:
                self._edb.close()
            except Exception:
                pass
            finally:
                self._edb = None
                self._components.clear()
                self.controller_list.clear()
                self.dram_list.clear()
                self.single_list.clear()
                self.diff_list.clear()
                self._update_reference_combo([])
                self._aedb_path = None
                self._aedb_source_path = None
                self._component_nets.clear()
                self._cct_metadata = None
                self._cct_port_entries = []
                if hasattr(self, "cct_table"):
                    self.cct_table.setRowCount(0)
                self._set_simulation_sources([], None)
                self._update_action_state()
                self._update_simulation_ui_state()
                self._update_cct_ui_state()

    def _populate_components(self) -> None:
        self.controller_list.blockSignals(True)
        self.dram_list.blockSignals(True)
        self.controller_list.clear()
        self.dram_list.clear()
        self._component_nets.clear()

        if self._edb is None:
            self._update_reference_combo([])
            return

        components = getattr(self._edb.components, "components", {})
        filtered = [
            (name, comp, self._pin_count(comp))
            for name, comp in components.items()
            if self._matches_component_pattern(name)
        ]
        filtered.sort(key=lambda item: (-item[2], item[0]))

        self._components = {name: comp for name, comp, _ in filtered}
        for name, _, pin_count in filtered:
            nets = self._extract_net_names(self._components[name])
            self._component_nets[name] = nets
            label = f"{name} ({pin_count})"
            for list_widget in (self.controller_list, self.dram_list):
                item = QListWidgetItem(label)
                item.setData(Qt.UserRole, name)
                list_widget.addItem(item)

        self.controller_list.blockSignals(False)
        self.dram_list.blockSignals(False)
        self._update_results()

    def _update_results(self) -> None:
        self.single_list.clear()
        self.diff_list.clear()

        if self._edb is None:
            self._update_reference_combo([])
            self._set_status_message("Load an AEDB to see nets")
            self._update_action_state()
            return

        controller_names = self._selected_component_names(self.controller_list)
        dram_names = self._selected_component_names(self.dram_list)

        if not controller_names or not dram_names:
            self._update_reference_combo([])
            self._set_status_message(
                "Select at least one controller and one DRAM component"
            )
            self._update_action_state()
            return

        controller_nets = self._collect_net_names(controller_names)
        dram_nets = self._collect_net_names(dram_names)
        shared_nets = sorted(controller_nets & dram_nets)

        all_selected = controller_names + dram_names
        common_nets = self._nets_in_all_components(all_selected)
        self._update_reference_combo(common_nets)

        diff_pairs = self._shared_diff_pairs(shared_nets)
        diff_net_names = {name for _, pos, neg in diff_pairs for name in (pos, neg)}
        filtered_single_nets = [net for net in shared_nets if net not in diff_net_names]

        for net_name in filtered_single_nets:
            item = QListWidgetItem(net_name)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)
            self.single_list.addItem(item)

        for pair_name, pos_name, neg_name in diff_pairs:
            label = f"{pair_name}: {pos_name} / {neg_name}"
            item = QListWidgetItem(label)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)
            self.diff_list.addItem(item)

        single_count = len(filtered_single_nets)
        diff_count = len(diff_pairs)
        self._set_status_message(
            f"Controllers: {len(controller_names)} | DRAMs: {len(dram_names)} | "
            f"Shared nets: {single_count} | Shared differential pairs: {diff_count}"
        )
        self._update_action_state()

    def _collect_net_names(self, component_names: Iterable[str]) -> set[str]:
        names: set[str] = set()
        for name in component_names:
            component = self._components.get(name)
            if component is None:
                continue
            cached = self._component_nets.get(name)
            if cached is None:
                cached = self._extract_net_names(component)
                self._component_nets[name] = cached
            names.update(cached)
        return names

    def _checked_diff_entries(self) -> List[Dict[str, str]]:
        entries: List[Dict[str, str]] = []
        for index in range(self.diff_list.count()):
            item = self.diff_list.item(index)
            if item is None or item.checkState() != Qt.Checked:
                continue
            text = item.text()
            if ": " in text:
                pair_name, payload = text.split(": ", 1)
            else:
                pair_name, payload = text, text
            parts = [token.strip() for token in payload.split("/")]
            if len(parts) != 2:
                continue
            pos_name, neg_name = parts
            if pos_name:
                entries.append(
                    {
                        "net": pos_name,
                        "pair": pair_name.strip(),
                        "polarity": "positive",
                    }
                )
            if neg_name:
                entries.append(
                    {
                        "net": neg_name,
                        "pair": pair_name.strip(),
                        "polarity": "negative",
                    }
                )
        return entries

    def _checked_single_nets(self) -> List[str]:
        nets: List[str] = []
        for index in range(self.single_list.count()):
            item = self.single_list.item(index)
            if item is not None and item.checkState() == Qt.Checked:
                nets.append(item.text())
        return nets

    def _checked_diff_nets(self) -> List[str]:
        return [entry["net"] for entry in self._checked_diff_entries()]

    def _checked_net_names(self) -> List[str]:
        ordered = self._checked_single_nets() + self._checked_diff_nets()
        seen: set[str] = set()
        unique: List[str] = []
        for net in ordered:
            if net not in seen:
                seen.add(net)
                unique.append(net)
        return unique

    def _checked_net_metadata(self) -> Dict[str, Dict[str, Optional[str]]]:
        metadata: Dict[str, Dict[str, Optional[str]]] = {}
        for net in self._checked_single_nets():
            metadata[net] = {"type": "single"}
        for entry in self._checked_diff_entries():
            metadata[entry["net"]] = {
                "type": "differential",
                "pair": entry.get("pair"),
                "polarity": entry.get("polarity"),
            }
        return metadata

    def _selected_reference_net(self) -> Optional[str]:
        if not self.reference_combo.isEnabled():
            return None
        text = self.reference_combo.currentText().strip()
        if not text or text == "Select reference net...":
            return None
        return text

    def _update_action_state(self) -> None:
        button = getattr(self, "apply_button", None)
        if button is None or self.port_count_label is None:
            return

        checked_nets = self._checked_net_names()
        controllers = self._selected_component_names(self.controller_list)
        drams = self._selected_component_names(self.dram_list)
        selected_components = controllers + drams
        estimated_ports = self._estimate_port_count(checked_nets, selected_components)
        self.port_count_label.setText(
            f"Checked nets: {len(checked_nets)} | Ports: {estimated_ports}"
        )

        has_edb = self._edb is not None
        has_reference = self._selected_reference_net() is not None
        has_nets = bool(checked_nets)
        has_controller = bool(controllers)
        has_dram = bool(drams)

        button.setEnabled(has_edb and has_reference and has_nets and has_controller and has_dram)
        self._update_simulation_ui_state()
        self._update_cct_ui_state()

    def _estimate_port_count(
        self, net_names: Iterable[str], component_names: Iterable[str]
    ) -> int:
        if not net_names or not component_names:
            return 0

        total = 0
        net_set = set(net_names)
        for component_name in component_names:
            nets = self._component_nets.get(component_name)
            if nets is None:
                component = self._components.get(component_name)
                if component is None:
                    continue
                nets = self._extract_net_names(component)
                self._component_nets[component_name] = nets
            total += len(net_set & nets)
        return total

    def _write_port_metadata(
        self,
        metadata: List[Dict[str, Optional[str]]],
        aedb_path: Path,
        reference_net: str,
        controllers: List[str],
        drams: List[str],
    ) -> Path:
        if not metadata:
            raise ValueError("No port metadata available to write.")

        normalized = []
        for index, entry in enumerate(metadata, 1):
            normalized.append(
                {
                    "sequence": index,
                    "name": entry.get("name"),
                    "component": entry.get("component"),
                    "component_role": entry.get("component_role"),
                    "net": entry.get("net"),
                    "net_type": entry.get("net_type"),
                    "pair": entry.get("pair"),
                    "polarity": entry.get("polarity"),
                    "reference_net": reference_net,
                }
            )

        data = {
            "aedb_path": str(aedb_path),
            "reference_net": reference_net,
            "controller_components": controllers,
            "dram_components": drams,
            "ports": normalized,
        }

        json_path = aedb_path.parent / f"{aedb_path.stem}_ports.json"
        with json_path.open("w", encoding="utf-8") as handle:
            json.dump(data, handle, indent=2)

        return json_path

    def _after_metadata_saved(self, json_path: Path) -> None:
        self.cct_json_edit.setText(str(json_path))
        if load_port_metadata is not None:
            self._load_cct_metadata()
        else:
            self._update_cct_ui_state()

    def _apply_changes(self) -> None:
        if self._edb is None:
            QMessageBox.warning(self, "No design", "Load an AEDB design before applying changes.")
            return

        nets = self._checked_net_names()
        if not nets:
            QMessageBox.warning(self, "No nets selected", "Check at least one net before applying changes.")
            return

        reference_net = self._selected_reference_net()
        if reference_net is None:
            QMessageBox.warning(self, "No reference net", "Select a reference net before applying changes.")
            return

        controller_names = self._selected_component_names(self.controller_list)
        dram_names = self._selected_component_names(self.dram_list)
        if not controller_names or not dram_names:
            QMessageBox.warning(
                self,
                "Incomplete selection",
                "Select at least one controller and one DRAM component before applying changes.",
            )
            return

        components = controller_names + dram_names
        component_roles = {name: "controller" for name in controller_names}
        component_roles.update({name: "dram" for name in dram_names})
        net_metadata = self._checked_net_metadata()

        try:
            port_metadata = self._create_ports_for_nets(
                nets,
                reference_net,
                components,
                component_roles,
                net_metadata,
            )
        except Exception as exc:  # pragma: no cover - runtime feedback path
            self._show_error("Failed to create ports", exc)
            return

        ports_created = len(port_metadata)
        saved_path: Optional[Path] = None
        if ports_created > 0:
            try:
                saved_path = self._save_modified_design()
            except Exception as exc:  # pragma: no cover - runtime feedback path
                self._show_error("Failed to save AEDB", exc)
                self._set_status_message("Failed to save AEDB; see error details.")
                self._update_action_state()
                return

        json_path: Optional[Path] = None
        if ports_created > 0 and saved_path is not None:
            try:
                json_path = self._write_port_metadata(
                    port_metadata,
                    saved_path,
                    reference_net,
                    controller_names,
                    dram_names,
                )
                self._after_metadata_saved(json_path)
            except Exception as exc:  # pragma: no cover - runtime feedback path
                self._show_error("Failed to write port metadata", exc)
                self._set_status_message("Failed to write port metadata; see error details.")
                self._update_action_state()
                return

        if ports_created == 0:
            QMessageBox.information(
                self,
                "Apply complete",
                "No matching component pins were found for the checked nets.",
            )
        elif saved_path is None:
            QMessageBox.information(
                self,
                "Apply complete",
                f"Created {ports_created} ports referencing {reference_net}.\n"
                "Design was not saved to a new AEDB.",
            )
        else:
            QMessageBox.information(
                self,
                "Apply complete",
                f"Created {ports_created} ports referencing {reference_net}.\n"
                f"Saved to {saved_path}" + (
                    f"\nPort metadata: {json_path}" if json_path else ""
                ),
            )

        if saved_path is not None:
            if json_path is not None:
                self._set_status_message(
                    f"Created {ports_created} ports, saved to {saved_path}, metadata {json_path}"
                )
            else:
                self._set_status_message(
                    f"Created {ports_created} ports and saved to {saved_path}"
                )
        else:
            self._set_status_message(
                f"Created {ports_created} ports for {len(nets)} nets using {reference_net}"
            )

        self._set_simulation_sources(nets, reference_net)
        self._update_action_state()

    def _default_output_path(self) -> Optional[Path]:
        source = self._aedb_source_path or self._aedb_path
        if source is None:
            return None
        stem = source.stem
        stem = re.sub(r"(_applied)+$", "", stem)
        if not stem:
            stem = source.stem
        return source.with_name(f"{stem}_applied.aedb")

    def _save_modified_design(self) -> Optional[Path]:
        if self._edb is None:
            return None

        target = self._default_output_path()
        if target is None:
            return None

        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            self._edb.save_edb_as(str(target))
            try:
                self._edb.close_edb()
            except Exception:
                pass

            try:
                new_edb = Edb(str(target), edbversion=self._edb_version)
            except Exception:
                self._edb = None
                self._aedb_path = None
                self._components.clear()
                self.controller_list.clear()
                self.dram_list.clear()
                self.single_list.clear()
                self.diff_list.clear()
                self._update_reference_combo([])
                raise
        finally:
            QApplication.restoreOverrideCursor()

        self._edb = new_edb
        self._aedb_path = target
        self.file_label.setText(str(target))
        self._populate_components()
        self._update_simulation_ui_state()
        return target

    def _create_ports_for_nets(
        self,
        net_names: Iterable[str],
        reference_net: str,
        component_names: Iterable[str],
        component_roles: Dict[str, str],
        net_metadata: Dict[str, Dict[str, Optional[str]]],
    ) -> List[Dict[str, Optional[str]]]:
        if self._edb is None:
            return []

        edb = self._edb
        nets_container = getattr(edb.nets, "nets", {})
        if isinstance(nets_container, dict):
            nets_lookup = nets_container
        else:
            nets_lookup = dict(getattr(nets_container, "items", lambda: [])())

        component_set = set(component_names)
        component_order = {name: idx for idx, name in enumerate(component_names)}
        reference_terminals: Dict[str, object] = {}
        metadata: List[Dict[str, Optional[str]]] = []

        for net_name in net_names:
            net_obj = nets_lookup.get(net_name)
            if net_obj is None:
                continue

            net_components = getattr(net_obj, "components", {})
            if isinstance(net_components, dict):
                component_iterable = list(net_components.keys())
            else:
                component_iterable = list(net_components)

            component_iterable.sort(key=lambda name: component_order.get(name, len(component_order)))

            for component_name in component_iterable:
                if component_name not in component_set:
                    continue
                if component_name not in self._components:
                    continue

                reference_terminal = reference_terminals.get(component_name)
                if reference_terminal is None:
                    reference_terminal = self._create_reference_terminal(component_name, reference_net)
                    if reference_terminal is None:
                        continue
                    reference_terminals[component_name] = reference_terminal

                signal_terminal = self._create_signal_terminal(component_name, net_name)
                if signal_terminal is None:
                    continue

                base_port_name = f"{component_name}_{net_name}"
                sequence = len(metadata) + 1
                port_name = prefix_port_name(base_port_name, sequence)
                if hasattr(signal_terminal, "SetName"):
                    signal_terminal.SetName(port_name)
                if hasattr(signal_terminal, "SetReferenceTerminal"):
                    signal_terminal.SetReferenceTerminal(reference_terminal)
                net_info = net_metadata.get(net_name, {"type": "single"})
                metadata.append(
                    {
                        "sequence": sequence,
                        "name": port_name,
                        "component": component_name,
                        "component_role": component_roles.get(component_name, "unknown"),
                        "net": net_name,
                        "net_type": net_info.get("type", "single"),
                        "pair": net_info.get("pair"),
                        "polarity": net_info.get("polarity"),
                    }
                )

        return metadata

    def _create_reference_terminal(self, component_name: str, reference_net: str):
        group_name = self._sanitized_group_name(component_name, reference_net, suffix="ref")
        pin_group = self._ensure_pin_group(component_name, reference_net, group_name)
        if pin_group is None:
            return None

        terminal = pin_group.create_port_terminal(50)
        if hasattr(terminal, "SetName"):
            terminal.SetName(f"ref;{component_name};{reference_net}")
        return terminal

    def _create_signal_terminal(self, component_name: str, net_name: str):
        group_name = self._sanitized_group_name(component_name, net_name)
        pin_group = self._ensure_pin_group(component_name, net_name, group_name)
        if pin_group is None:
            return None
        return pin_group.create_port_terminal(50)

    def _ensure_pin_group(self, component_name: str, net_name: str, group_name: str):
        if self._edb is None:
            return None

        result = self._edb.core_siwave.create_pin_group_on_net(component_name, net_name, group_name)
        return self._pin_group_from_result(result)

    @staticmethod
    def _pin_group_from_result(result):
        candidates = []
        if isinstance(result, (list, tuple)):
            candidates.extend(result)
        else:
            candidates.append(result)

        for entry in reversed(candidates):
            if hasattr(entry, "create_port_terminal"):
                return entry
        return None

    @staticmethod
    def _sanitized_group_name(component_name: str, net_name: str, suffix: Optional[str] = None) -> str:
        base = f"{component_name}_{net_name}"
        safe = re.sub(r"[^A-Za-z0-9_]+", "_", base)
        safe = re.sub(r"_+", "_", safe).strip("_")
        if suffix:
            safe = f"{safe}_{suffix}" if safe else suffix
        if not safe:
            fallback = re.sub(r"[^A-Za-z0-9_]+", "_", component_name).strip("_") or "comp"
            suffix_part = suffix or "pg"
            safe = f"{fallback}_{suffix_part}"
        return safe

    def _nets_in_all_components(self, component_names: Iterable[str]) -> List[str]:
        valid_names = [name for name in component_names if self._components.get(name) is not None]
        if not valid_names:
            return []

        common: set[str] | None = None
        for name in valid_names:
            nets = self._extract_net_names(self._components[name])
            if common is None:
                common = set(nets)
            else:
                common &= nets
            if not common:
                break

        return sorted(common or [])

    def _update_reference_combo(self, common_nets: Iterable[str]) -> None:
        nets = sorted(set(common_nets))
        self.reference_combo.blockSignals(True)
        self.reference_combo.clear()

        if not nets:
            self.reference_combo.addItem("Select reference net...")
            self.reference_combo.setEnabled(False)
            self.reference_combo.blockSignals(False)
            self._update_action_state()
            return

        self.reference_combo.setEnabled(True)
        self.reference_combo.addItem("Select reference net...")
        for net in nets:
            self.reference_combo.addItem(net)

        default_net = next((net for net in nets if "gnd" in net.lower()), None)
        if default_net:
            index = self.reference_combo.findText(default_net)
            if index != -1:
                self.reference_combo.setCurrentIndex(index)
        else:
            self.reference_combo.setCurrentIndex(0)

        self.reference_combo.blockSignals(False)
        self._update_action_state()

    def _shared_diff_pairs(self, shared_nets: Iterable[str]) -> List[Tuple[str, str, str]]:
        if self._edb is None:
            return []

        shared_set = set(shared_nets)
        diff_container = getattr(self._edb, "differential_pairs", None)
        if diff_container is None:
            return []

        try:
            items_attr = diff_container.items
        except AttributeError:
            return []

        if callable(items_attr):
            candidates = items_attr()
            if isinstance(candidates, dict):
                iterator = candidates.items()
            else:
                iterator = candidates
        elif isinstance(items_attr, dict):
            iterator = items_attr.items()
        else:
            iterator = []

        matches: List[Tuple[str, str, str]] = []
        for name, diff in iterator:
            pos = getattr(diff, "positive_net", None)
            neg = getattr(diff, "negative_net", None)
            pos_name = getattr(pos, "name", None)
            neg_name = getattr(neg, "name", None)
            if pos_name in shared_set and neg_name in shared_set:
                matches.append((name, pos_name, neg_name))

        return sorted(matches, key=lambda item: item[0])

    @staticmethod
    def _extract_net_names(component: object) -> set[str]:
        names: set[str] = set()
        nets = getattr(component, "nets", [])
        if isinstance(nets, dict):
            iterable = nets.values()
        else:
            iterable = nets
        for net in iterable:
            if isinstance(net, str):
                names.add(net)
            else:
                name = getattr(net, "name", None)
                if name:
                    names.add(name)
        return names

    @staticmethod
    def _selected_component_names(list_widget: QListWidget) -> List[str]:
        names: List[str] = []
        for item in list_widget.selectedItems():
            name = item.data(Qt.UserRole)
            if isinstance(name, str) and name:
                names.append(name)
            else:
                names.append(item.text())
        return names

    @staticmethod
    def _pin_count(component: object) -> int:
        pins = getattr(component, "pins", None)
        if pins is None:
            return 0
        try:
            keys = getattr(pins, "keys", None)
            if callable(keys):
                return len(list(keys()))
        except Exception:
            pass
        try:
            return len(pins)
        except TypeError:
            return sum(1 for _ in getattr(pins, "values", lambda: [])())

    @staticmethod
    def _matches_component_pattern(name: str) -> bool:
        return bool(_COMPONENT_PATTERN.match(name))

    def _show_error(self, title: str, exc: Exception) -> None:
        details = "".join(traceback.format_exception(exc))
        QMessageBox.critical(self, title, f"{exc}\n\n{details}")

    def closeEvent(self, event) -> None:  # noqa: D401 - Qt signature
        if self._cct_thread is not None and self._cct_thread.isRunning():
            QMessageBox.warning(
                self,
                'CCT running',
                'Please wait for the current CCT calculation to finish before closing the application.',
            )
            event.ignore()
            return

        self._close_edb()
        super().closeEvent(event)


def main() -> None:
    app = QApplication(sys.argv)
    widget = EdbGui()
    widget.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()





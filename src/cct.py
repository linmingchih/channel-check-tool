import json
import math
import re
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

try:  # pragma: no cover - optional dependency
    from ansys.aedt.core import Circuit
    from ansys.aedt.core.generic.constants import Setups
except ImportError:  # pragma: no cover - allow metadata utilities without AEDT
    Circuit = None
    Setups = None

try:  # pragma: no cover - optional dependency for pruning
    import skrf as rf
except ImportError:  # pragma: no cover - pruning falls back to no-op
    rf = None

import numpy as np

ROOT_DIR = Path(__file__).resolve().parents[1]
NETLIST_DEBUG_DIR = ROOT_DIR / "data" / "netlist"
TRIMMED_TOUCHSTONE_DIRNAME = "trimmed_touchstone"
DEFAULT_CIRCUIT_VERSION = "2025.1"

def integrate_nonuniform(x_list, y_list):
    integral = 0.0
    for i in range(len(x_list) - 1):
        dx = x_list[i + 1] - x_list[i]
        integral += 0.5 * (y_list[i] + y_list[i + 1]) * dx
    return integral


def get_sig_isi(time_list, voltage_list, unit_interval):
    """Compute signal and ISI metrics over the provided waveform."""
    t = np.asarray(time_list, dtype=float)
    v = np.asarray(voltage_list, dtype=float)
    if t.ndim != 1 or v.ndim != 1 or t.size != v.size:
        raise ValueError("time_list and voltage_list must be 1-D and of equal length")
    if unit_interval <= 0:
        raise ValueError("unit_interval must be positive")

    order = np.argsort(t)
    t = t[order]
    v = v[order]

    if t[-1] - t[0] < unit_interval:
        raise ValueError("Waveform duration is shorter than unit interval")

    dt = np.diff(t)
    trap = np.concatenate([[0.0], np.cumsum((v[:-1] + v[1:]) * 0.5 * dt)])
    trap_abs = np.concatenate([[0.0], np.cumsum((np.abs(v[:-1]) + np.abs(v[1:])) * 0.5 * dt)])
    total_abs = trap_abs[-1]

    n = len(t)
    last_i = np.searchsorted(t, t[-1] - unit_interval, side="right") - 1
    if last_i < 0:
        raise ValueError("No valid integration window of length unit_interval")

    sig_max = -np.inf
    best_i = best_j = 0
    best_t_end = None
    j = 0

    for i in range(last_i + 1):
        t_end = t[i] + unit_interval
        while j + 1 < n and t[j + 1] <= t_end:
            j += 1

        integ = trap[j] - trap[i]
        if j + 1 < n and t[j] < t_end < t[j + 1]:
            v_end = v[j] + (v[j + 1] - v[j]) * (t_end - t[j]) / (t[j + 1] - t[j])
            integ += 0.5 * (v[j] + v_end) * (t_end - t[j])

        if integ > sig_max:
            sig_max = integ
            best_i, best_j, best_t_end = i, j, t_end

    i, j, t_end = best_i, best_j, best_t_end
    integ_abs = trap_abs[j] - trap_abs[i]
    if j + 1 < n and t[j] < t_end < t[j + 1]:
        v_end = v[j] + (v[j + 1] - v[j]) * (t_end - t[j]) / (t[j + 1] - t[j])
        integ_abs += 0.5 * (abs(v[j]) + abs(v_end)) * (t_end - t[j])

    sig = float(sig_max)
    isi = float(total_abs - integ_abs)
    return sig, isi


@dataclass
class PortMetadata:
    sequence: int
    name: str
    component: str
    component_role: str
    net: str
    net_type: str
    pair: Optional[str] = None
    polarity: Optional[str] = None


@dataclass
class PruneResult:
    kept_sequences: List[int]
    trimmed_metadata: List[PortMetadata]
    touchstone_path: Path
    txs: List[object]
    rxs: List[object]
    tx_lookup: Dict[Tuple[str, str], object]
    stats: Dict[str, object]


def _normalize_role(value: Optional[str]) -> str:
    if not value:
        return "unknown"
    value = str(value).lower()
    if value in {"controller", "ctrl", "host"}:
        return "controller"
    if value in {"dram", "memory", "mem"}:
        return "dram"
    return value


def _normalize_net_type(value: Optional[str]) -> str:
    if not value:
        return "single"
    value = str(value).lower()
    if value in {"diff", "differential"}:
        return "differential"
    return "single"


def _normalize_polarity(value: Optional[str]) -> Optional[str]:
    if value is None:
        return None
    value = str(value).lower()
    if value in {"positive", "pos", "+", "p"}:
        return "positive"
    if value in {"negative", "neg", "-", "n"}:
        return "negative"
    return value


def prefix_port_name(name: str, sequence: int) -> str:
    base = str(name or '')
    match = re.match(r'^\d+_(.*)$', base)
    if match:
        base = match.group(1)
    base = base.strip()
    return f"{sequence}_{base}" if base else str(sequence)


def load_port_metadata(path: str | Path) -> Tuple[List[PortMetadata], Dict[str, object]]:
    meta_path = Path(path)
    with meta_path.open("r", encoding="utf-8") as handle:
        raw: Dict[str, object] = json.load(handle)

    entries: List[PortMetadata] = []
    for index, item in enumerate(raw.get("ports", []), 1):
        entries.append(
            PortMetadata(
                sequence=int(item.get("sequence", index)),
                name=str(item.get("name", "")),
                component=str(item.get("component", "")),
                component_role=_normalize_role(item.get("component_role")),
                net=str(item.get("net", "")),
                net_type=_normalize_net_type(item.get("net_type")),
                pair=item.get("pair"),
                polarity=_normalize_polarity(item.get("polarity")),
            )
        )

    if not entries:
        raise ValueError(f"No ports found in metadata file: {meta_path}")

    entries.sort(key=lambda entry: entry.sequence)
    for idx, entry in enumerate(entries, 1):
        entry.sequence = idx
        entry.name = prefix_port_name(entry.name, idx)

    return entries, raw


def _clone_port(entry: PortMetadata, new_sequence: int) -> PortMetadata:
    return PortMetadata(
        sequence=new_sequence,
        name=prefix_port_name(entry.name, new_sequence),
        component=entry.component,
        component_role=entry.component_role,
        net=entry.net,
        net_type=entry.net_type,
        pair=entry.pair,
        polarity=entry.polarity,
    )

class Tx:
    def __init__(self, meta: PortMetadata, vhigh, t_rise, ui, res_tx, cap_tx):
        self.meta = meta
        self.pid = meta.sequence
        self.sequence = meta.sequence
        self.label = meta.name
        self.active = [
            f"V{self.pid} netb_{self.pid} 0 PULSE(0 {vhigh} 1e-10 {t_rise} {t_rise} {ui} 1.5e+100)",
            f"R{self.pid} netb_{self.pid} net_{self.pid} {res_tx}",
            f"C{self.pid} netb_{self.pid} 0 {cap_tx}",
        ]
        self.passive = [
            f"R{self.pid} netb_{self.pid} net_{self.pid} {res_tx}",
            f"C{self.pid} netb_{self.pid} 0 {cap_tx}",
        ]
        self.kind = 'single'
        self.key = meta.net

    def get_netlist(self, active: bool = True) -> List[str]:
        return self.active if active else self.passive


class Tx_diff:
    def __init__(
        self,
        positive: PortMetadata,
        negative: PortMetadata,
        vhigh,
        t_rise,
        ui,
        res_tx,
        cap_tx,
    ) -> None:
        self.pos = positive
        self.neg = negative
        self.pid_pos = positive.sequence
        self.pid_neg = negative.sequence
        self.sequence = min(positive.sequence, negative.sequence)
        self.label = positive.pair or f"{positive.name}/{negative.name}"
        self.active = [
            f"V{self.pid_pos} netb_{self.pid_pos} 0 PULSE(0 {vhigh} 1e-10 {t_rise} {t_rise} {ui} 1.5e+100)",
            f"R{self.pid_pos} netb_{self.pid_pos} net_{self.pid_pos} {res_tx}",
            f"C{self.pid_pos} netb_{self.pid_pos} 0 {cap_tx}",
            f"V{self.pid_neg} netb_{self.pid_neg} 0 PULSE(0 -{vhigh} 1e-10 {t_rise} {t_rise} {ui} 1.5e+100)",
            f"R{self.pid_neg} netb_{self.pid_neg} net_{self.pid_neg} {res_tx}",
            f"C{self.pid_neg} netb_{self.pid_neg} 0 {cap_tx}",
        ]
        self.passive = [
            f"R{self.pid_pos} netb_{self.pid_pos} net_{self.pid_pos} {res_tx}",
            f"C{self.pid_pos} netb_{self.pid_pos} 0 {cap_tx}",
            f"R{self.pid_neg} netb_{self.pid_neg} net_{self.pid_neg} {res_tx}",
            f"C{self.pid_neg} netb_{self.pid_neg} 0 {cap_tx}",
        ]
        self.kind = 'diff'
        self.key = tuple(sorted([positive.net, negative.net]))

    def get_netlist(self, active: bool = True) -> List[str]:
        return self.active if active else self.passive


class Rx:
    def __init__(self, meta: PortMetadata, res_rx, cap_rx):
        self.meta = meta
        self.pid = meta.sequence
        self.sequence = meta.sequence
        self.label = meta.name
        self.netlist = [
            f"R{self.pid} net_{self.pid} 0 {res_rx}",
            f"C{self.pid} net_{self.pid} 0 {cap_rx}",
        ]
        self.waveforms: Dict[object, Tuple[List[float], List[float]]] = {}
        self.expected_tx: Optional[object] = None
        self.kind = 'single'
        self.key = meta.net

    def get_netlist(self) -> List[str]:
        return self.netlist


class Rx_diff:
    def __init__(self, positive: PortMetadata, negative: PortMetadata, res_rx, cap_rx):
        self.pos = positive
        self.neg = negative
        self.pid_pos = positive.sequence
        self.pid_neg = negative.sequence
        self.label = positive.pair or f"{positive.name}/{negative.name}"
        self.netlist = [
            f"R{self.pid_pos} net_{self.pid_pos} 0 {res_rx}",
            f"C{self.pid_pos} net_{self.pid_pos} 0 {cap_rx}",
            f"R{self.pid_neg} net_{self.pid_neg} 0 {res_rx}",
            f"C{self.pid_neg} net_{self.pid_neg} 0 {cap_rx}",
        ]
        self.waveforms: Dict[object, Tuple[List[float], List[float]]] = {}
        self.expected_tx: Optional[object] = None
        self.kind = 'diff'
        self.key = tuple(sorted([positive.net, negative.net]))

    def get_netlist(self) -> List[str]:
        return self.netlist


class Design:
    def __init__(self, workdir: Path, tstep='100ps', tstop='3ns', version: Optional[str] = None):
        if Circuit is None or Setups is None:
            raise ImportError("ansys.aedt.core is required to run CCT simulations")

        self.workdir = Path(workdir)
        self.workdir.mkdir(parents=True, exist_ok=True)

        self.netlist_path = self.workdir / f"{uuid.uuid4()}.cir"
        self.netlist_path.write_text('', encoding='utf-8')

        version_str = (str(version).strip() if version is not None else '') or DEFAULT_CIRCUIT_VERSION
        self.circuit_version = version_str

        self.circuit = circuit = Circuit(
            version=self.circuit_version,
            non_graphical=True,
            close_on_exit=True,
        )

        circuit.add_netlist_datablock(str(self.netlist_path))
        self.setup = circuit.create_setup('myTransient', Setups.NexximTransient)
        self.setup.props['TransientData'] = [tstep, tstop]
        self.circuit.save_project()

    def run(self, netlist):
        with open(self.netlist_path, 'w') as f:
            f.write(netlist)

        self.circuit.odesign.InvalidateSolution('myTransient')
        self.circuit.save_project()
        self.circuit.analyze('myTransient')
        self.circuit.save_project()

        result = {}
        for v in self.circuit.post.available_report_quantities():
            data = self.circuit.post.get_solution_data(v, domain='Time')
            x = [1e3 * i for i in data.primary_sweep_values]
            y = [1e-3 * i for i in data.data_real()]
            m = re.search(r'net_(\d+)', v)
            if m:
                number = int(m.group(1))
                result[number] = (x, y)
        return result

class CCT:
    def __init__(
        self,
        snp_path: str | Path,
        port_metadata_path: str | Path,
        workdir: Optional[str | Path] = None,
        threshold_db: Optional[float] = None,
        circuit_version: Optional[str] = None,
    ):
        self.snp_path = str(snp_path)
        self.port_metadata, self.metadata_info = load_port_metadata(port_metadata_path)
        self.reference_net = self.metadata_info.get("reference_net")
        self.controller_components = self.metadata_info.get("controller_components", [])
        self.dram_components = self.metadata_info.get("dram_components", [])

        metadata_dir = Path(port_metadata_path).resolve().parent
        if workdir is None:
            workdir = metadata_dir / "cct_work"
        self.workdir = Path(workdir)
        self.workdir.mkdir(parents=True, exist_ok=True)

        self.output_dir = metadata_dir
        NETLIST_DEBUG_DIR.mkdir(parents=True, exist_ok=True)

        self.threshold_db = threshold_db
        version_candidate = circuit_version if circuit_version is not None else self.metadata_info.get("circuit_version")
        version_str = (str(version_candidate).strip() if version_candidate is not None else '') or DEFAULT_CIRCUIT_VERSION
        self.circuit_version = version_str
        self._prune_cache: Dict[Tuple[str, str], PruneResult] = {}
        self._prerun_summaries: List[Dict[str, object]] = []
        self._prune_warning_emitted = False

        self._metadata_by_sequence = {entry.sequence: entry for entry in self.port_metadata}

        self.txs: List[object] = []
        self.rxs: List[object] = []
        self.tx_single_map: Dict[str, Tx] = {}
        self.tx_diff_map: Dict[Tuple[str, str], Tx_diff] = {}
        self.rx_single_map: Dict[str, Rx] = {}
        self.rx_diff_map: Dict[Tuple[str, str], Rx_diff] = {}
        self._tx_lookup: Dict[Tuple[str, str], object] = {}
        self._rx_lookup: Dict[Tuple[str, str], object] = {}
        self.tx_config: Optional[Dict[str, str]] = None
        self.rx_config: Optional[Dict[str, str]] = None

        self._classify_ports()

        nets = ' '.join([f'net_{entry.sequence}' for entry in self.port_metadata])
        self.netlist = [
            self._channel_model_line(self.snp_path),
            f'S1 {nets} FQMODEL="Channel"',
        ]

        self._network = None
        if rf is not None:
            try:
                self._network = rf.Network(self.snp_path)
            except Exception:
                self._network = None

        self._trim_dir = self.workdir / TRIMMED_TOUCHSTONE_DIRNAME

    @staticmethod
    def _channel_model_line(tstone_path: str | Path) -> str:
        return (
            f'.model "Channel" S TSTONEFILE="{tstone_path}" '
            'INTERPOLATION=LINEAR INTDATTYP=MA HIGHPASS=10 LOWPASS=10 '
            'convolution=1 enforce_passivity=0 Noisemodel=External'
        )

    def _classify_port_groups(self, entries: Iterable[PortMetadata]):
        entries = list(entries)
        tx_single_entries = [
            entry for entry in entries if entry.component_role == "controller" and entry.net_type == "single"
        ]
        rx_single_entries = [
            entry for entry in entries if entry.component_role == "dram" and entry.net_type == "single"
        ]
        tx_single_entries.sort(key=lambda entry: entry.sequence)
        rx_single_entries.sort(key=lambda entry: entry.sequence)
        tx_diff_entries = self._group_differential("controller", entries)
        rx_diff_entries = self._group_differential("dram", entries)
        return tx_single_entries, rx_single_entries, tx_diff_entries, rx_diff_entries

    def _classify_ports(self) -> None:
        (
            self.tx_single_entries,
            self.rx_single_entries,
            self.tx_diff_entries,
            self.rx_diff_entries,
        ) = self._classify_port_groups(self.port_metadata)

        controller_sequences = [entry.sequence for entry in self.tx_single_entries]
        for pos_entry, neg_entry in self.tx_diff_entries:
            controller_sequences.extend([pos_entry.sequence, neg_entry.sequence])
        self._controller_sequences = sorted(set(controller_sequences))

        self._rx_total_groups = len(self.rx_single_entries) + len(self.rx_diff_entries)
        self._rx_total_ports = len(self.rx_single_entries) + 2 * len(self.rx_diff_entries)

    def _group_differential(
        self,
        role: str,
        entries: Optional[Iterable[PortMetadata]] = None,
    ) -> List[Tuple[PortMetadata, PortMetadata]]:
        source = list(entries) if entries is not None else self.port_metadata
        groups: Dict[Tuple[str, str], Dict[str, PortMetadata]] = {}
        for entry in source:
            if entry.component_role != role or entry.net_type != "differential":
                continue
            key = (entry.component, entry.pair or entry.net)
            suggested = "positive" if "positive" not in groups.get(key, {}) else "negative"
            polarity = entry.polarity or suggested
            groups.setdefault(key, {})[polarity] = entry

        pairs: List[Tuple[PortMetadata, PortMetadata]] = []
        for (_component, _pair_name), mapping in groups.items():
            pos = mapping.get("positive")
            neg = mapping.get("negative")
            if not pos or not neg:
                continue
            pairs.append((pos, neg))

        pairs.sort(key=lambda item: min(item[0].sequence, item[1].sequence))
        return pairs

    @staticmethod
    def _diff_identifier(positive: PortMetadata, negative: PortMetadata) -> Tuple[str, str]:
        return tuple(sorted([positive.net, negative.net]))
    def _create_tx_objects(
        self,
        tx_single_entries: Iterable[PortMetadata],
        tx_diff_entries: Iterable[Tuple[PortMetadata, PortMetadata]],
        *,
        vhigh,
        t_rise,
        ui,
        res_tx,
        cap_tx,
    ) -> Tuple[List[object], Dict[str, Tx], Dict[Tuple[str, str], Tx_diff]]:
        txs: List[object] = []
        tx_single_map: Dict[str, Tx] = {}
        tx_diff_map: Dict[Tuple[str, str], Tx_diff] = {}

        for entry in tx_single_entries:
            tx = Tx(entry, vhigh, t_rise, ui, res_tx, cap_tx)
            txs.append(tx)
            tx_single_map[entry.net] = tx

        for pos_entry, neg_entry in tx_diff_entries:
            txd = Tx_diff(pos_entry, neg_entry, vhigh, t_rise, ui, res_tx, cap_tx)
            txs.append(txd)
            identifier = self._diff_identifier(pos_entry, neg_entry)
            txd.key = identifier
            tx_diff_map[identifier] = txd

        return txs, tx_single_map, tx_diff_map

    def _create_rx_objects(
        self,
        rx_single_entries: Iterable[PortMetadata],
        rx_diff_entries: Iterable[Tuple[PortMetadata, PortMetadata]],
        *,
        res_rx,
        cap_rx,
        tx_single_map: Dict[str, Tx],
        tx_diff_map: Dict[Tuple[str, str], Tx_diff],
    ) -> Tuple[List[object], Dict[str, Rx], Dict[Tuple[str, str], Rx_diff]]:
        rxs: List[object] = []
        rx_single_map: Dict[str, Rx] = {}
        rx_diff_map: Dict[Tuple[str, str], Rx_diff] = {}

        for entry in rx_single_entries:
            rx = Rx(entry, res_rx, cap_rx)
            rx.expected_tx = tx_single_map.get(entry.net)
            rxs.append(rx)
            rx_single_map[entry.net] = rx

        for pos_entry, neg_entry in rx_diff_entries:
            rx = Rx_diff(pos_entry, neg_entry, res_rx, cap_rx)
            identifier = self._diff_identifier(pos_entry, neg_entry)
            rx.expected_tx = tx_diff_map.get(identifier)
            rxs.append(rx)
            rx_diff_map[identifier] = rx

        return rxs, rx_single_map, rx_diff_map

    def set_threshold(self, threshold_db: Optional[float]) -> None:
        self.threshold_db = threshold_db
        self._prune_cache.clear()
        self._prerun_summaries.clear()

    def set_txs(self, vhigh, t_rise, ui, res_tx, cap_tx):
        self.ui = ui
        self.tx_config = {
            "vhigh": vhigh,
            "t_rise": t_rise,
            "ui": ui,
            "res_tx": res_tx,
            "cap_tx": cap_tx,
        }

        self.txs, self.tx_single_map, self.tx_diff_map = self._create_tx_objects(
            self.tx_single_entries,
            self.tx_diff_entries,
            vhigh=vhigh,
            t_rise=t_rise,
            ui=ui,
            res_tx=res_tx,
            cap_tx=cap_tx,
        )

        self._tx_lookup = {self._tx_to_key(tx): tx for tx in self.txs}
        self._prune_cache.clear()
        self._prerun_summaries.clear()

    def set_rxs(self, res_rx, cap_rx):
        if self.tx_config is None:
            raise RuntimeError("set_txs must be called before set_rxs")

        self.rx_config = {
            "res_rx": res_rx,
            "cap_rx": cap_rx,
        }

        self.rxs, self.rx_single_map, self.rx_diff_map = self._create_rx_objects(
            self.rx_single_entries,
            self.rx_diff_entries,
            res_rx=res_rx,
            cap_rx=cap_rx,
            tx_single_map=self.tx_single_map,
            tx_diff_map=self.tx_diff_map,
        )

        self._rx_lookup = {self._rx_to_key(rx): rx for rx in self.rxs}
        for rx in self.rxs:
            rx.waveforms.clear()

        self._prune_cache.clear()
        self._prerun_summaries.clear()

    def _tx_to_key(self, tx: object) -> Tuple[str, str]:
        if isinstance(tx, Tx_diff):
            identifier = self._diff_identifier(tx.pos, tx.neg)
            return ("diff", "::".join(identifier))
        if isinstance(tx, Tx):
            return ("single", tx.meta.net)
        raise TypeError(f"Unsupported TX type: {type(tx)!r}")

    def _rx_to_key(self, rx: object) -> Tuple[str, str]:
        if isinstance(rx, Rx_diff):
            identifier = self._diff_identifier(rx.pos, rx.neg)
            return ("diff", "::".join(identifier))
        if isinstance(rx, Rx):
            return ("single", rx.meta.net)
        raise TypeError(f"Unsupported RX type: {type(rx)!r}")

    def _ensure_prune_result(self, tx: object) -> PruneResult:
        key = self._tx_to_key(tx)
        cached = self._prune_cache.get(key)
        if cached is not None:
            return cached
        prune_result = self._compute_prune_result(tx)
        self._prune_cache[key] = prune_result
        return prune_result

    def _compute_prune_result(self, tx: object) -> PruneResult:
        if self.tx_config is None or self.rx_config is None:
            raise RuntimeError("set_txs and set_rxs must be called before running pruning")

        total_port_count = len(self.port_metadata)
        kept_sequences = set(self._controller_sequences)

        if isinstance(tx, Tx_diff):
            tx_sequences = [tx.pid_pos, tx.pid_neg]
        elif isinstance(tx, Tx):
            tx_sequences = [tx.pid]
        else:
            raise TypeError(f"Unsupported TX type: {type(tx)!r}")

        if self.threshold_db is not None and self._network is None and not self._prune_warning_emitted:
            print('[prune] scikit-rf not available; pruning disabled for this run')
            self._prune_warning_emitted = True
        if self.threshold_db is None or self._network is None:
            kept_sequences.update(range(1, total_port_count + 1))
        else:
            tx_indices = [seq - 1 for seq in tx_sequences]
            threshold = float(self.threshold_db)

            for entry in self.rx_single_entries:
                base_rx = self.rx_single_map.get(entry.net)
                keep = False
                if base_rx is not None and base_rx.expected_tx is tx:
                    keep = True
                else:
                    rx_idx = entry.sequence - 1
                    peak = 0.0
                    for tx_idx in tx_indices:
                        data = np.abs(self._network.s[:, rx_idx, tx_idx])
                        if data.size:
                            peak = max(peak, float(np.max(data)))
                    peak_db = 20 * math.log10(peak) if peak > 0 else float('-inf')
                    keep = peak_db >= threshold
                if keep:
                    kept_sequences.add(entry.sequence)

            for pos_entry, neg_entry in self.rx_diff_entries:
                identifier = self._diff_identifier(pos_entry, neg_entry)
                base_rx = self.rx_diff_map.get(identifier)
                keep = False
                if base_rx is not None and base_rx.expected_tx is tx:
                    keep = True
                else:
                    rx_indices = [pos_entry.sequence - 1, neg_entry.sequence - 1]
                    peak = 0.0
                    for rx_idx in rx_indices:
                        for tx_idx in tx_indices:
                            data = np.abs(self._network.s[:, rx_idx, tx_idx])
                            if data.size:
                                peak = max(peak, float(np.max(data)))
                    peak_db = 20 * math.log10(peak) if peak > 0 else float('-inf')
                    keep = peak_db >= threshold
                if keep:
                    kept_sequences.update([pos_entry.sequence, neg_entry.sequence])

            if not kept_sequences.issuperset(self._controller_sequences):
                kept_sequences.update(self._controller_sequences)

        kept_sequences_sorted = sorted(kept_sequences)
        trimmed_metadata: List[PortMetadata] = []
        for new_sequence, original_sequence in enumerate(kept_sequences_sorted, 1):
            original_entry = self._metadata_by_sequence[original_sequence]
            trimmed_metadata.append(_clone_port(original_entry, new_sequence))

        (
            tx_single_entries,
            rx_single_entries,
            tx_diff_entries,
            rx_diff_entries,
        ) = self._classify_port_groups(trimmed_metadata)

        txs, tx_single_map, tx_diff_map = self._create_tx_objects(
            tx_single_entries,
            tx_diff_entries,
            **self.tx_config,
        )
        rxs, rx_single_map, rx_diff_map = self._create_rx_objects(
            rx_single_entries,
            rx_diff_entries,
            res_rx=self.rx_config["res_rx"],
            cap_rx=self.rx_config["cap_rx"],
            tx_single_map=tx_single_map,
            tx_diff_map=tx_diff_map,
        )

        tx_lookup = {self._tx_to_key(tx_obj): tx_obj for tx_obj in txs}

        kept_rx_group_count = len(rx_single_entries) + len(rx_diff_entries)
        kept_rx_port_count = len(rx_single_entries) + 2 * len(rx_diff_entries)

        touchstone_path = Path(self.snp_path)
        if self.threshold_db is not None and self._network is not None and kept_rx_group_count < self._rx_total_groups:
            self._trim_dir.mkdir(parents=True, exist_ok=True)
            port_indices = [seq - 1 for seq in kept_sequences_sorted]
            trimmed_network = self._network.subnetwork(port_indices)
            base_label = getattr(tx, 'label', 'tx')
            label = self._sanitize_label(base_label)
            port_count = len(kept_sequences_sorted)
            filename = f"{Path(self.snp_path).stem}_{label}_{port_count}p"
            trimmed_network.write_touchstone(filename=filename, dir=str(self._trim_dir))
            touchstone_path = self._trim_dir / f"{filename}.s{port_count}p"

        stats = {
            "tx_label": getattr(tx, 'label', 'tx'),
            "threshold_db": self.threshold_db,
            "kept_port_count": len(kept_sequences_sorted),
            "total_port_count": total_port_count,
            "kept_rx_port_count": kept_rx_port_count,
            "total_rx_port_count": self._rx_total_ports,
            "kept_rx_group_count": kept_rx_group_count,
            "total_rx_group_count": self._rx_total_groups,
            "touchstone_path": str(touchstone_path),
        }

        prune_result = PruneResult(
            kept_sequences=kept_sequences_sorted,
            trimmed_metadata=trimmed_metadata,
            touchstone_path=touchstone_path,
            txs=txs,
            rxs=rxs,
            tx_lookup=tx_lookup,
            stats=stats,
        )
        return prune_result

    @staticmethod
    def _sanitize_label(label: str) -> str:
        sanitized = re.sub(r'[^A-Za-z0-9_.-]+', '_', label).strip('_')
        return sanitized or 'tx'

    def pre_run(self, threshold_db: Optional[float] = None) -> List[Dict[str, object]]:
        if threshold_db is not None:
            self.set_threshold(threshold_db)
        if not self.txs or not self.rxs:
            raise RuntimeError("set_txs and set_rxs must be called before pre_run")
        summaries: List[Dict[str, object]] = []
        for tx in self.txs:
            prune_result = self._ensure_prune_result(tx)
            stats = dict(prune_result.stats)
            self._log_prune_stats(stats)
            summaries.append(stats)
        if summaries:
            ratio = sum(s["kept_port_count"] / s["total_port_count"] for s in summaries) / len(summaries)
            rx_ratio = 0.0
            if self._rx_total_ports:
                rx_ratio = sum(s["kept_rx_port_count"] / s["total_rx_port_count"] for s in summaries) / len(summaries)
            print(
                f"[prune] Average kept ports: {ratio:.1%}; average kept RX ports: {rx_ratio:.1%}"
            )
        self._prerun_summaries = summaries
        return summaries

    def _log_prune_stats(self, stats: Dict[str, object]) -> None:
        kept = stats.get("kept_port_count", 0)
        total = stats.get("total_port_count", 1)
        kept_ratio = kept / total if total else 0.0
        rx_kept = stats.get("kept_rx_port_count", 0)
        rx_total = stats.get("total_rx_port_count", 0)
        rx_ratio = rx_kept / rx_total if rx_total else 0.0
        threshold = stats.get("threshold_db")
        msg = (
            f"[prune] Tx {stats.get('tx_label', 'tx')}: "
            f"ports {kept}/{total} ({kept_ratio:.1%})"
        )
        if rx_total:
            msg += f", rx ports {rx_kept}/{rx_total} ({rx_ratio:.1%})"
        if threshold is not None:
            msg += f", threshold {threshold} dB"
        print(msg)

    def run(self, tstep='100ps', tstop='3ns'):
        if not self.txs or not self.rxs:
            raise RuntimeError("set_txs and set_rxs must be called before run")

        design = Design(self.workdir, tstep, tstop, version=self.circuit_version)
        for rx in self.rxs:
            rx.waveforms.clear()

        for tx in self.txs:
            prune_result = self._ensure_prune_result(tx)
            if not self._prerun_summaries:
                self._log_prune_stats(prune_result.stats)

            netlist_lines = self._build_netlist(prune_result, tx)
            netlist_text = '\n'.join(netlist_lines)
            self._write_debug_netlist(tx, netlist_text)
            result = design.run(netlist_text)
            self._store_waveforms(prune_result, result, tx)

    def _build_netlist(self, prune_result: PruneResult, active_tx: object) -> List[str]:
        nets = ' '.join([f'net_{entry.sequence}' for entry in prune_result.trimmed_metadata])
        netlist = [
            self._channel_model_line(prune_result.touchstone_path),
            f'S1 {nets} FQMODEL="Channel"',
        ]
        active_key = self._tx_to_key(active_tx)
        trimmed_active_tx = prune_result.tx_lookup.get(active_key)
        for tx in prune_result.txs:
            netlist.extend(tx.get_netlist(tx is trimmed_active_tx))
        for rx in prune_result.rxs:
            netlist.extend(rx.get_netlist())
        return netlist

    def _store_waveforms(self, prune_result: PruneResult, result: Dict[int, Tuple[List[float], List[float]]], base_tx: object) -> None:
        for rx in prune_result.rxs:
            base_rx = self._rx_lookup.get(self._rx_to_key(rx))
            if base_rx is None:
                continue
            if isinstance(rx, Rx):
                if rx.sequence in result:
                    base_rx.waveforms[base_tx] = result[rx.sequence]
            elif isinstance(rx, Rx_diff):
                if rx.pid_pos in result and rx.pid_neg in result:
                    time_pos, waveform_pos = result[rx.pid_pos]
                    _, waveform_neg = result[rx.pid_neg]
                    new_result = (
                        time_pos,
                        [vpos - vneg for vpos, vneg in zip(waveform_pos, waveform_neg)],
                    )
                    base_rx.waveforms[base_tx] = new_result

    def calculate(self, output_path):
        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)

        ui = float(self.ui.replace('ps', ''))

        result = []
        for rx in self.rxs:
            if not getattr(rx, 'waveforms', None):
                continue
            primary_tx = getattr(rx, 'expected_tx', None)
            if primary_tx is None:
                continue
            waveform_primary = rx.waveforms.get(primary_tx)
            if waveform_primary is None:
                continue
            sig = isi = 0.0
            xtalk = 0.0
            for tx, waveform in rx.waveforms.items():
                time, voltage = waveform
                if tx == primary_tx:
                    sig, isi = get_sig_isi(time, voltage, ui)
                else:
                    xtalk += integrate_nonuniform(time, [abs(v) for v in voltage])
            pseudo_eye = sig - isi - xtalk
            denom = isi + xtalk
            p_ratio = sig / denom if denom else float('inf')

            tx_label = getattr(primary_tx, 'label', getattr(primary_tx, 'pid', 'unknown'))
            rx_label = getattr(rx, 'label', str(getattr(rx, 'pid', 'unknown')))
            result.append(
                f'{tx_label}, {rx_label}, {sig:.3f}, {isi:.3f}, {xtalk:.3f}, {pseudo_eye:.3f}, {p_ratio:.3f}'
            )

        with output_file.open('w') as f:
            f.writelines('tx_name, rx_name, sig(V*ps), isi(V*ps), xtalk(V*ps), pseudo_eye(V*ps), power_ratio\n')
            f.write('\n'.join(result))

    def _write_debug_netlist(self, tx_obj: object, netlist_text: str) -> None:
        if not netlist_text:
            return
        sequence = getattr(tx_obj, 'sequence', None)
        label = getattr(tx_obj, 'label', f"tx_{sequence if sequence is not None else 'unknown'}")
        sanitized = re.sub(r'[^A-Za-z0-9_.-]+', '_', label).strip('_') or 'tx'
        if sequence is not None:
            filename = f"netlist_{sequence:03d}_{sanitized}.cir"
        else:
            filename = f"netlist_{sanitized}.cir"
        path = NETLIST_DEBUG_DIR / filename
        path.write_text(netlist_text, encoding='utf-8')


if __name__ == '__main__':
    import sys

    if len(sys.argv) >= 3:
        touchstone_path = sys.argv[1]
        metadata_path = sys.argv[2]
        output_csv = sys.argv[3] if len(sys.argv) >= 4 else str(Path(metadata_path).with_name(f"{Path(metadata_path).stem}_cct.csv"))
        threshold_arg = sys.argv[4] if len(sys.argv) >= 5 else None
        version_arg = sys.argv[5] if len(sys.argv) >= 6 else None
    else:
        touchstone_path = r"D:\OneDrive - ANSYS, Inc\a-client-repositories\quanta-cct-circuit-202508\data\Sweep1_DV3.s88p"
        metadata_path = r"D:\OneDrive - ANSYS, Inc\a-client-repositories\quanta-cct-circuit-202508\output\Sweep1_DV3_ports.json"
        output_csv = str(Path(metadata_path).with_name(f"{Path(metadata_path).stem}_cct.csv"))
        threshold_arg = None
        version_arg = None

    threshold_db = None
    if threshold_arg is not None:
        try:
            threshold_db = float(threshold_arg)
        except ValueError:
            threshold_db = None

    cct = CCT(touchstone_path, metadata_path, threshold_db=threshold_db, circuit_version=version_arg)
    cct.set_txs(vhigh="0.8V", t_rise="30ps", ui="133ps", res_tx="40ohm", cap_tx="1pF")
    cct.set_rxs(res_rx="30ohm", cap_rx="1.8pF")
    if threshold_db is not None:
        cct.pre_run()
    cct.run()
    cct.calculate(output_path=output_csv)

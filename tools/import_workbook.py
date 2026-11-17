"""Build a model input YAML from the VSM Excel workbook (INPUT + TS_* sheets).

Usage::

    python tools/import_workbook.py VSM_EVAP.xlsx config/line_parameters_workbook.yaml

What is taken from the workbook
-------------------------------
* Per-EVA cycle time of every process for models A-D: the "Cycle time used" row
  of each time-study sheet (measured laps where they exist, otherwise the
  whiteboard rate x driver count + fixed minutes). The fin-press value already
  includes the punch OEE of each model.
* Standard time = cycle x rating / (1 - allowance), as on the workbook.
  With allowance 15 % this equals dividing by the 85 % worker efficiency, so
  both readings of the workbook give the same simulated times.
* Product mix from the ORDERS table (units per model). The model simulates
  one EVA type, so each process time is the mix-weighted mean, and its CV is
  the CV of the mixture (within-model CV from the base YAML plus the spread
  between models).
* Operators during each time study (``crew_basis``) and the crew size on INPUT
  (resource ``crew``): e.g. steeking timed with 2 people but run by 1 takes
  twice as long, as on the workbook. Crew sizes, the worker allocation
  ("Allocated to") -> shared resources,
  lead times of procured/outsourced steps, the working week and the demand
  (units in the planning period / planning weeks).

Everything the workbook does not contain (engineering release, test failure
rate, within-model CVs, distribution families, arrival variability, ...) is
copied from the base YAML and left as it is.
"""

from __future__ import annotations

import argparse
import copy
import math
from pathlib import Path

import openpyxl
import yaml

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT / "config" / "line_parameters.yaml"
MODELS = ("A", "B", "C", "D")

# process ID -> (model operation key, time-study sheet)
PROCESS = {
    "E01": ("fin_press", "TS_FIN_PRESS"),
    "E02": ("boxing", "TS_BOXING"),
    "E03": ("trim_tubes", "TS_TRIM_TUBES"),
    "E04": ("weld_hairpins", "TS_WELDING_HAIRPINS"),
    "E05": ("steeking", "TS_STEEKING"),
    "E06": ("weld_return_bends", "TS_RETURN_BENDS"),
    "E07": ("expansion", "TS_EXPANDING"),
    "E08": ("prepare_headers", "TS_PREPARE_HEADERS"),
    "E09": ("header_welding", "TS_WELDING_HEADERS"),
    "E11": ("wash", "TS_WASH_BAY"),
    "E10": ("pressure_test", "TS_PRESSURE_TEST"),
    "A01": ("complete_assembly", "TS_COMPLETE_ASSY"),
    "S01": ("press_brake", "TS_PLENUM_PRESS_BRAKE"),
    "S02": ("flat_welding", "TS_PLENUM_FLAT_WELD"),
    "S03": ("plenum_assembly", "TS_PLENUM_ASSEMBLY"),
    "S04": ("drip_tray_weld", "TS_DRIP_TRAY_WELD"),
    "S05": ("dye_penetration_test", "TS_DYE_PEN_TEST"),
    "S06": ("drip_tray_assembly", "TS_DRIP_TRAY_ASSY"),
}
LEAD = {"P01": "procurement_tubing", "P02": "procurement_sheet_metal",
        "O03": "drip_tray_external", "O04": "coating"}
DEPT = {"L1 Evaporator": "evaporator", "L2 Sheet metal": "sheet_metal",
        "Final assembly": "assembly"}


def slug(name: str) -> str:
    return "".join(c if c.isalnum() else "_" for c in str(name).strip().lower()).strip("_")


def find_row(ws, col: str, text: str, start: int = 1) -> int:
    for r in range(start, ws.max_row + 1):
        if str(ws[f"{col}{r}"].value or "").strip() == text:
            return r
    raise KeyError(f"'{text}' not found in column {col} of {ws.title}")


def mixture(means: list[float], weights: list[float], cv_within: float) -> tuple[float, float]:
    """Mean and CV of a mixture of models (each with the same within-model CV)."""
    m = sum(w * x for w, x in zip(weights, means, strict=True))
    second = sum(w * x * x * (1 + cv_within ** 2) for w, x in zip(weights, means, strict=True))
    return m, math.sqrt(max(second - m * m, 0.0)) / m


def read_workbook(path: Path) -> dict:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["INPUT"]
    v = lambda c: ws[c].value  # noqa: E731
    d: dict = {
        "hours_mon_thu": float(v("D6")), "hours_fri": float(v("D7")),
        "hours_week": float(v("D8")), "days_week": int(v("D9")),
        "period_weeks": float(v("D10")), "allowance": float(v("D12")),
        "rating": float(v("D13")),
    }
    units = {m: float(ws.cell(7, 10 + k).value or 0) for k, m in enumerate(MODELS)}
    d["units"], d["orders"] = units, float(v("N8"))
    # order list (due dates)
    r0 = find_row(ws, "I", "Order no.") + 1
    d["order_list"] = [(ws[f"I{r}"].value, ws[f"J{r}"].value, ws[f"K{r}"].value, ws[f"L{r}"].value)
                       for r in range(r0, r0 + 20) if ws[f"I{r}"].value]
    # processes, crew, lead times, allocation
    r0 = find_row(ws, "B", "Process") + 1
    procs = {}
    for r in range(r0, r0 + 40):
        pid = ws[f"C{r}"].value
        if not pid:
            break
        procs[pid] = {"name": ws[f"B{r}"].value, "type": ws[f"D{r}"].value,
                      "crew": ws[f"E{r}"].value, "lead_days": ws[f"F{r}"].value,
                      "line": ws[f"G{r}"].value, "worker": ws[f"H{r}"].value}
    d["processes"] = procs
    # workers (hours per week)
    rw = find_row(ws, "I", "Worker") + 1
    d["workers"] = {}
    for r in range(rw, rw + 12):
        name = ws[f"I{r}"].value
        if not name or not isinstance(ws[f"K{r}"].value, (int, float)):
            break
        d["workers"][name] = {"line": ws[f"J{r}"].value, "hours_week": float(ws[f"K{r}"].value),
                              "efficiency": float(ws[f"L{r}"].value)}
    # cycle times per model from the time-study sheets
    d["cycle"], d["data"], d["specs"] = {}, {}, {}
    for pid, (_, sheet) in PROCESS.items():
        ts = wb[sheet]
        r = find_row(ts, "G", "Cycle time used [min/EVA]")
        d["cycle"][pid] = [float(ts.cell(r, 8 + k).value) for k in range(4)]
        d["data"][pid] = str(ts[f"H{find_row(ts, 'G', 'Source')}"].value)
        d.setdefault("study_ops", {})[pid] = int(
            ts[f"H{find_row(ts, 'G', 'Operators during the time study')}"].value or 1)
    r = find_row(ws, "B", "Fins")
    d["specs"]["fins"] = [float(ws.cell(r, 4 + k).value) for k in range(4)]
    return d


def anonymise(d: dict) -> None:
    """Replace people's names by "Operator n" (stations such as "FA Bay 1" are kept)."""
    people = [n for n in d["workers"] if "bay" not in str(n).lower()]
    alias = {n: f"Operator {k}" for k, n in enumerate(people, 1)}
    d["workers"] = {alias.get(n, n): w for n, w in d["workers"].items()}
    for p in d["processes"].values():
        if p["worker"] in alias:
            p["worker"] = alias[p["worker"]]


def build(d: dict, base: dict) -> dict:
    cfg = copy.deepcopy(base)
    tot = sum(d["units"].values())
    w = [d["units"][m] / tot for m in MODELS]
    std = d["rating"] / (1 - d["allowance"])     # cycle -> standard time
    day_min = d["hours_week"] * 60 / d["days_week"]
    demand = tot / d["period_weeks"]
    mix = ", ".join(f"{int(d['units'][m])} x {m}" for m in MODELS if d["units"][m])

    cfg["meta"].update({
        "model_name": "Evaporator production system - imported from VSM workbook",
        "config_version": "3.0.0-workbook",
        "data_status": "Imported from the VSM Excel workbook (INPUT and time-study sheets)",
        "product_mix": mix,
        "standard_time_factor": round(std, 6),
    })
    cfg["calendar"] = {"minutes_per_day": round(day_min, 3), "days_per_week": d["days_week"]}
    cfg["demand"].update({
        "eva_per_week": demand,
        "due_date_weeks": d["period_weeks"],
        "sweep_eva_per_week": [round(demand * f, 3) for f in (1, 1.5, 2, 2.5, 3, 3.5, 4)],
        "scenario_eva_per_week": [demand, 2 * demand],
    })
    lot = int(round(tot))
    cfg["orders"] = {"lot_size": lot, "transfer_batch": lot}

    # resources = workers (shared between their stations); FA bay gets one
    # server per 37.75 h of weekly hours
    resources = {}
    for name, wk in d["workers"].items():
        servers = max(1, int(round(wk["hours_week"] / d["hours_week"])))
        resources[slug(name)] = {"label": str(name), "servers": servers, "crew": 1}
    cfg["resources"] = resources
    cfg["crew_efficiency"] = {1: 1.0, 2: 1.0, 3: 0.85, 4: 0.75}

    ops = {}
    for pid, (op, _) in PROCESS.items():
        p = d["processes"][pid]
        crew = int(p["crew"] or 1)            # operators on the station (INPUT)
        basis = d["study_ops"][pid]          # operators during the time study
        res = slug(p["worker"])
        if crew > resources[res]["crew"]:
            resources[res]["crew"] = crew
        means = [x * std for x in d["cycle"][pid]]
        if op == "fin_press":
            m, _ = mixture(means, w, 0.0)
            fins = sum(wi * f for wi, f in zip(w, d["specs"]["fins"], strict=True))
            rate = 7.0
            cfg["fin_press"].update({
                "source": "estimate", "nominal_rate_fpm": rate,
                "good_fins_per_eva": round(fins, 3),
                # the workbook's per-model punch OEE is folded into P; no separate
                # setup, stops, scrap or reel changes (they are inside the OEE)
                "performance": round(fins / (rate * m), 6), "quality": 1.0,
                "setup_min": 0.0, "setup_cv": 0.0, "mtbf_min": 1.0e9, "mttr_min": 1.0,
                "reel_fins": 1.0e9, "reload_min": 0.0, "resource": res,
            })
            continue
        old = base["operations"].get(op, {})
        cv_in = float(old.get("cv", 0.15))
        m, cv = mixture(means, w, cv_in)
        src = "time study" if d["data"][pid].startswith("Measured") else "estimate"
        o = {"step": old.get("step", 14), "label": str(p["name"]),
             "department": DEPT.get(p["line"], "evaporator"), "resource": res,
             "mean_min": round(m, 3), "crew_basis": basis, "cv": round(cv, 4),
             "source": src, "workbook_id": pid,
             "workbook_cycle_min_by_model": dict(zip(MODELS, [round(x, 2) for x in
                                                          d["cycle"][pid]], strict=True))}
        if "distribution" in old:
            o["distribution"] = old["distribution"]
        ops[op] = o
    cfg["operations"] = ops

    for pid, key in LEAD.items():
        cfg["lead_times"][key].update({"mean_days": float(d["processes"][pid]["lead_days"]),
                                       "source": "estimate"})
        cfg["lead_times"][key].pop("brief_mean_days", None)
        cfg["lead_times"][key].pop("calibrated", None)
    cfg["rework"]["resource"] = ops["header_welding"]["resource"]

    st = ops["steeking"]["mean_min"]
    cfg["input_uncertainty"]["steeking_triangular_min"] = [round(st * f, 1)
                                                          for f in (360 / 420, 1, 600 / 420)]
    cfg["input_uncertainty"]["steeking_cases_hours"] = [round(st * f / 60, 2)
                                                        for f in (6 / 7, 1, 16 / 7, 20 / 7)]
    return cfg


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("workbook", type=Path)
    ap.add_argument("out", type=Path)
    ap.add_argument("--base", type=Path, default=BASE)
    ap.add_argument("--keep-names", action="store_true",
                    help="keep workers' names (default: Operator 1, 2, ...)")
    a = ap.parse_args()
    base = yaml.safe_load(a.base.read_text())
    d = read_workbook(a.workbook)
    if not a.keep_names:
        anonymise(d)
    cfg = build(d, base)
    head = ("# Generated by tools/import_workbook.py from "
            f"{a.workbook.name} - do not edit by hand; edit the workbook and re-import.\n")
    a.out.write_text(head + yaml.safe_dump(cfg, sort_keys=False, width=100))
    print(f"wrote {a.out}: demand {cfg['demand']['eva_per_week']} EVA/week, "
          f"mix {cfg['meta']['product_mix']}, {len(cfg['resources'])} workers")
    for op, o in cfg["operations"].items():
        print(f"  {op:22s} {o['mean_min']:8.1f} min/EVA  cv {o['cv']:.2f}  {o['resource']}")


if __name__ == "__main__":
    main()

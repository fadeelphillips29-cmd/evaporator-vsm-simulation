"""Demand sweep and one-piece-flow what-if for the workbook inputs.

Usage: PYTHONPATH=src python tools/workbook_sweep.py
Writes results/workbook/demand_sweep.csv (100 replications per case).
"""

import pandas as pd

from evapsim.config import apply_overrides, load_config, with_demand
from evapsim.metrics import summarize
from evapsim.runner import resource_summary, run_replications


def main() -> None:
    c = load_config("config/line_parameters_workbook.yaml")
    top2 = list(resource_summary(run_replications(c, 10, "probe"))
                .sort_values("busy", ascending=False)["resource"][:2])
    cases = [(d, "lot 3, transfer batch 3", {}) for d in (1.0, 1.5, 1.75, 1.85, 2.0)]
    cases.append((1.5, "transfer batch 1 (one-piece flow)", {"orders": {"transfer_batch": 1}}))
    rows = []
    for d, lab, ov in cases:
        run = run_replications(apply_overrides(with_demand(c, d), ov), 100, lab)
        k = summarize(run.reps).set_index("metric")["mean"]
        rs = resource_summary(run).set_index("resource")["busy"]
        rows.append({"demand": d, "case": lab, "TH": k["throughput_eva_week"],
                     "LT_mean_d": k["lt_mean_days"], "LT_P90_d": k["lt_p90_days"],
                     "WIP": k["wip_total"], **{r: rs[r] for r in top2}})
    df = pd.DataFrame(rows).round(3)
    print(df.to_string(index=False))
    df.to_csv("results/workbook/demand_sweep.csv", index=False)


if __name__ == "__main__":
    main()

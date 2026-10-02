# Databricks notebook source
# MAGIC %md
# MAGIC # Forecast Performance Analysis — Internal Model vs. Vendor
# MAGIC
# MAGIC **Refactored version.** Every cell that changed opens with a `REFACTOR NOTE` comment
# MAGIC block naming the recommendation it implements:
# MAGIC
# MAGIC | # | Recommendation | Where |
# MAGIC |---|----------------|-------|
# MAGIC | 1 | Parameterize the INTERNAL table name | Cell 2 (config), Cell 6 (reader) |
# MAGIC | 2 | Replace `scope.collect()` + `.isin()` with a broadcast join | Cell 7 |
# MAGIC | 3 | Consolidate duplicated WMAPE aggregation into `compute_wmape_metrics()` | Cell 3, Cell 10 |
# MAGIC | 4 | Consolidate duplicated chart code into `plot_wmape_chart()` | Cell 3, Cells 16–17 |
# MAGIC | 5 | Move Excel/PNG exports to a UC Volume | Cell 2, Cells 16–18 |
# MAGIC | 6 | Add data quality guards + automated profiling thresholds | Cell 3, Cell 4, Cell 7 |
# MAGIC | 7 | Stop mutating the source INTERNAL table — write wins to a lookup table | Cell 12 |
# MAGIC | 8 | Auto-generated markdown summary | Cell 19 |

# COMMAND ----------

# MAGIC %md
# MAGIC ### REFACTOR NOTE (supporting change for Rec 5)
# MAGIC `%pip install openpyxl` used to sit at cell 14, in the middle of the notebook.
# MAGIC On DBR 13.x+ `%pip` automatically restarts the Python interpreter, which wipes
# MAGIC every variable defined in earlier cells — so a top-to-bottom run would lose all
# MAGIC of its state at that point. It now runs first, before anything is defined.
# MAGIC (On DBR < 13.0, add `dbutils.library.restartPython()` after the install.)

# COMMAND ----------

# DBTITLE 1,Install openpyxl
# MAGIC %pip install openpyxl

# COMMAND ----------

# DBTITLE 1,Configuration
# ─────────────────────────────────────────────────────────────────────────────
# REFACTOR NOTE (Rec 1, 5, 6, 7)
#
# Rec 1 — INTERNAL table name is no longer hardcoded. Two supported modes:
#         * USE_COMPILED_INTERNAL = True  → read the single compiled table and
#           filter on its source_date partition column (preferred, no manual
#           edits, no table sprawl).
#         * USE_COMPILED_INTERNAL = False → derive the dated table name from
#           FCST_DATE (`..._08_09`) so the only thing to change per run is
#           FCST_DATE. INTERNAL_TABLE_OVERRIDE is the escape hatch for one-offs.
#
# Rec 5 — EXPORT_VOLUME added. All .xlsx/.png output now lands on a UC Volume
#         instead of the ephemeral driver filesystem.
#
# Rec 6 — DQ_* thresholds added so the data quality checks are configuration,
#         not hand-inspection of cell 2's profiling output.
#
# Rec 7 — WINS_LOOKUP_TABLE added. forecast_wins is written here instead of
#         being MERGEd back into the INTERNAL input table.
# ─────────────────────────────────────────────────────────────────────────────
from datetime import datetime

# ============ Run parameters — change these once per run ============
FCST_DATE  = "2026-08-09"   # vendor fcstdate filter; also drives the INTERNAL table suffix
START_DATE = "2026-08-30"   # INTERNAL forecast-Sunday filter
LAG_VALUE  = 3              # BY lag filter

# Weeks to exclude from charts (known bad-data weeks)
EXCLUDE_WEEKS = ["2026-08-16"]

# ============ Source tables ============
VENDOR_TABLE     = "my_catalog.vendor.vendor_forecast_perf"
DEMAND_TABLE = "my_catalog.internal_model.demand_classification"
SCOPE_TABLE  = "my_catalog.forecasting.scope_list"

# ---- Rec 1: INTERNAL table resolution -----------------------------------------
INTERNAL_SCHEMA          = "my_catalog.internal_model"
INTERNAL_TABLE_PREFIX    = "demand_forecast_detailed_output"
INTERNAL_DATE_COLUMN     = "source_date"          # partition column on the compiled table
USE_COMPILED_INTERNAL    = False                  # flip to True once the compiled table is authoritative
INTERNAL_TABLE_OVERRIDE  = None                   # e.g. "my_catalog.internal_model.some_adhoc_table"

COMPILED_TABLE = f"{INTERNAL_SCHEMA}.{INTERNAL_TABLE_PREFIX}_compiled"

def _dated_internal_table(fcst_date: str) -> str:
    """Build `..._MM_DD` from an ISO date so the suffix never has to be typed by hand."""
    suffix = datetime.strptime(fcst_date, "%Y-%m-%d").strftime("%m_%d")
    return f"{INTERNAL_SCHEMA}.{INTERNAL_TABLE_PREFIX}_{suffix}"

if INTERNAL_TABLE_OVERRIDE:
    INTERNAL_TABLE = INTERNAL_TABLE_OVERRIDE
elif USE_COMPILED_INTERNAL:
    INTERNAL_TABLE = COMPILED_TABLE
else:
    INTERNAL_TABLE = _dated_internal_table(FCST_DATE)

# ============ Output targets ============
SCHEMA         = "my_catalog.forecasting"
WINNERS_TABLE  = f"{INTERNAL_SCHEMA}.comparison_dfu_75pct_winners"

# Rec 7: forecast_wins lands here instead of being merged into INTERNAL_TABLE
WINS_LOOKUP_TABLE = f"{INTERNAL_SCHEMA}.comparison_dfu_forecast_wins"
# Seed the lookup from weeks already stamped in the compiled table so the
# ≥75% win-rate history survives the cutover. Set to False once backfilled.
WINS_BACKFILL_FROM_COMPILED = True

# Rec 5: UC Volume for durable, shareable exports
EXPORT_VOLUME = "/Volumes/my_catalog/forecasting/exports"

# ============ Rec 6: data quality thresholds ============
DQ_FAIL_ON_ERROR        = True   # False = warn and continue instead of raising
DQ_MIN_JOIN_RETENTION   = 0.95   # joined rows must be >= 95% of the INTERNAL input
DQ_MIN_VOLUME_RATIO     = 0.80   # actuals volume vs. median of prior weeks
DQ_MAX_VOLUME_RATIO     = 1.25
DQ_PROFILE_WEEKS        = 6      # how many prior fcstdates to profile against
DQ_MIN_SCOPE_MATCH      = 0.50   # >= 50% of the scope list should appear in the join

print(f"INTERNAL_TABLE        : {INTERNAL_TABLE}")
print(f"  resolution mode  : {'override' if INTERNAL_TABLE_OVERRIDE else ('compiled' if USE_COMPILED_INTERNAL else 'dated-from-FCST_DATE')}")
print(f"VENDOR_TABLE           : {VENDOR_TABLE}")
print(f"FCST_DATE / START  : {FCST_DATE} / {START_DATE}   (lag={LAG_VALUE})")
print(f"EXPORT_VOLUME      : {EXPORT_VOLUME}")

# COMMAND ----------

# DBTITLE 1,Shared Helpers
# ─────────────────────────────────────────────────────────────────────────────
# REFACTOR NOTE (Rec 3, 4, 5, 6)  — NEW CELL
#
# Everything reusable lives here so the analysis cells stay short:
#   * compute_wmape_metrics()  — Rec 3: the 9-metric aggregation that used to be
#     copy-pasted five times in cell 8. One definition, one place to fix a bug.
#   * plot_wmape_chart()       — Rec 4: the ~200-line chart body that used to be
#     duplicated across cells 15 and 16.
#   * dq_check / dq_row_retention / profile_volume — Rec 6: assertion helpers so
#     data quality is enforced rather than eyeballed.
#   * ensure_volume / save_to_volume — Rec 5: Volume-safe file writes.
# ─────────────────────────────────────────────────────────────────────────────
import os
import shutil
import warnings

import numpy as np
import pandas as pd

from pyspark.sql import DataFrame
from pyspark.sql.functions import (
    col, lit, when, count, broadcast,
    sum as spark_sum, abs as spark_abs, round as spark_round,
)

# Column-name constants — single place to change if the source schema drifts.
ACTUALS_COL   = "basehist"
OUR_FCST_COL  = "Our Base Forecast"
VENDOR_FCST_COL   = "VENDOR_Basefcst"
OUR_ERR_COL   = "Our Absolute error"
VENDOR_ERR_COL    = "Vendor Absolute error"


# ============================================================
# Rec 6 — Data quality guards
# ============================================================
class DataQualityWarning(UserWarning):
    """Raised as a warning when a DQ check fails and DQ_FAIL_ON_ERROR is False."""


def dq_check(label, passed, detail, fail_on_error=None):
    """Log a DQ result; raise (or warn) when it fails.

    Centralizing this means every check reports identically and the
    fail-vs-warn behaviour is controlled by one config flag.
    """
    fail_on_error = DQ_FAIL_ON_ERROR if fail_on_error is None else fail_on_error
    if passed:
        print(f"  [PASS] {label}: {detail}")
        return True
    message = f"[DQ FAIL] {label}: {detail}"
    if fail_on_error:
        raise AssertionError(message)
    print(f"  {message}")
    warnings.warn(message, DataQualityWarning)
    return False


def dq_row_retention(label, before, after, min_ratio=None, fail_on_error=None):
    """Assert that a join/filter retained at least `min_ratio` of its input rows."""
    min_ratio = DQ_MIN_JOIN_RETENTION if min_ratio is None else min_ratio
    ratio = (after / before) if before else 0.0
    return dq_check(
        label,
        ratio >= min_ratio,
        f"{after:,} of {before:,} rows retained ({ratio:.1%}); floor is {min_ratio:.0%}",
        fail_on_error=fail_on_error,
    )


def dq_range(label, value, low, high, unit="", fail_on_error=None):
    """Assert a value sits inside an expected band (used for volume drift)."""
    shown = f"{value:.3f}{unit}" if value is not None else "n/a"
    return dq_check(
        label,
        (value is not None) and (low <= value <= high),
        f"{shown} expected within [{low:.3f}, {high:.3f}]",
        fail_on_error=fail_on_error,
    )


# ============================================================
# Rec 5 — UC Volume helpers
# ============================================================
def ensure_volume(path=None):
    """Create the export directory on the Volume if it does not exist."""
    path = EXPORT_VOLUME if path is None else path
    try:
        os.makedirs(path, exist_ok=True)
        return path
    except Exception as exc:  # pragma: no cover - environment dependent
        print(f"  [WARN] could not create {path}: {exc}")
        return None


def save_to_volume(local_path, target_name=None, subdir=None):
    """Copy a driver-local file onto the Volume and return the final path.

    Written locally first, then copied: several writers (openpyxl, matplotlib)
    need random-access/seekable handles that Volumes do not always provide.
    """
    target_name = target_name or os.path.basename(local_path)
    dest_dir = os.path.join(EXPORT_VOLUME, subdir) if subdir else EXPORT_VOLUME
    if ensure_volume(dest_dir) is None:
        print(f"  [WARN] falling back to driver path: {local_path}")
        return local_path
    dest = os.path.join(dest_dir, target_name)
    shutil.copyfile(local_path, dest)
    print(f"  ✓ saved to Volume: {dest}")
    return dest


# ============================================================
# Rec 3 — One WMAPE aggregation, reused everywhere
# ============================================================
def compute_wmape_metrics(df: DataFrame, group_cols=None, total_volume=None,
                          drop_null_groups=True) -> DataFrame:
    """WMAPE / bias / win-rate metrics for any grouping.

    Replaces five near-identical `.agg(...)` blocks. Pass `group_cols=None`
    for the grand total, a string for one column, or a list for several.

    Metrics returned (unchanged from the original hand-written blocks):
        Our_WMAPE, VENDOR_WMAPE, Our_Bias, VENDOR_Bias, Total_Volume,
        Pct_of_Total_Volume, Our_Win_Rate_Pct, VENDOR_Win_Rate_Pct,
        WMAPE_Improvement_Percentage_Points

    Args:
        df:               rows already filtered to basehist > 0.
        group_cols:       column name, list of names, or None for overall.
        total_volume:     denominator for Pct_of_Total_Volume. Computed from
                          `df` when omitted — pass the shared value so every
                          breakdown is a percentage of the SAME denominator.
        drop_null_groups: drop rows whose grouping key is null (the original
                          code did this for demand_class / dmdgroup / Brand Code).
    """
    if group_cols is None:
        group_cols = []
    elif isinstance(group_cols, str):
        group_cols = [group_cols]
    else:
        group_cols = list(group_cols)

    if total_volume is None:
        total_volume = df.select(spark_sum(col(ACTUALS_COL))).first()[0]

    scoped = df
    if drop_null_groups:
        for c in group_cols:
            scoped = scoped.filter(col(c).isNotNull())

    aggregations = [
        (spark_sum(col(OUR_ERR_COL)) / spark_sum(col(ACTUALS_COL)) * 100).alias("Our_WMAPE"),
        (spark_sum(col(VENDOR_ERR_COL)) / spark_sum(col(ACTUALS_COL)) * 100).alias("VENDOR_WMAPE"),
        ((spark_sum(col(OUR_FCST_COL)) / spark_sum(col(ACTUALS_COL)) - 1.0) * 100).alias("Our_Bias"),
        ((spark_sum(col(VENDOR_FCST_COL)) / spark_sum(col(ACTUALS_COL)) - 1.0) * 100).alias("VENDOR_Bias"),
        spark_sum(col(ACTUALS_COL)).alias("Total_Volume"),
        (spark_sum(col(ACTUALS_COL)) / lit(total_volume) * 100).alias("Pct_of_Total_Volume"),
        (spark_sum(when(col(OUR_ERR_COL) < col(VENDOR_ERR_COL), 1).otherwise(0)) / count("*") * 100).alias("Our_Win_Rate_Pct"),
        (spark_sum(when(col(VENDOR_ERR_COL) < col(OUR_ERR_COL), 1).otherwise(0)) / count("*") * 100).alias("VENDOR_Win_Rate_Pct"),
        ((spark_sum(col(VENDOR_ERR_COL)) - spark_sum(col(OUR_ERR_COL))) / spark_sum(col(ACTUALS_COL)) * 100).alias("WMAPE_Improvement_Percentage_Points"),
    ]

    # groupBy() with no columns is a valid global aggregation in Spark,
    # so the grand-total case needs no special branch.
    return scoped.groupBy(*group_cols).agg(*aggregations)


# ============================================================
# Rec 4 — One chart function, reused by both chart cells
# ============================================================
def plot_wmape_chart(table_name, title, filter_expr=None, png_name=None,
                     display_weeks=13, exclude_weeks=None, save=True):
    """Render the two-panel WMAPE / improvement chart from a persisted Delta table.

    Cells 15 and 16 of the original notebook were ~200 near-identical lines each.
    The only real differences were the source table, the row filter and the title,
    so those are the arguments here.

    Args:
        table_name:    fully-qualified Delta table with the weekly metrics.
        title:         figure title (week count is appended automatically).
        filter_expr:   optional SQL string, e.g. "Scope = 'In'".
        png_name:      output filename; defaults to a slug of the title.
        display_weeks: width of the display axis (actual weeks + future placeholders).
        exclude_weeks: list of start_date strings to drop; defaults to EXCLUDE_WEEKS.
        save:          write the PNG to the Volume (Rec 5).

    Returns:
        (fig, pandas.DataFrame) so the caller can annotate further if needed.
    """
    import matplotlib.pyplot as plt
    from matplotlib.ticker import FixedLocator, FormatStrFormatter

    exclude_weeks = EXCLUDE_WEEKS if exclude_weeks is None else exclude_weeks

    # ---------- 1. Load ----------
    sdf = spark.read.table(table_name)
    if filter_expr:
        sdf = sdf.filter(filter_expr)
    if exclude_weeks:
        sdf = sdf.filter(~col("start_date").isin(exclude_weeks))

    df = sdf.orderBy("start_date").toPandas()
    if df.empty:
        raise ValueError(f"No rows in {table_name} for filter={filter_expr!r} after excluding {exclude_weeks}")

    df["start_date"] = pd.to_datetime(df["start_date"])
    df = df.sort_values("start_date").reset_index(drop=True)

    # ---------- 2. Series ----------
    x_data     = np.arange(len(df))
    our        = df["Our_WMAPE"].astype(float).to_numpy()
    by         = df["VENDOR_WMAPE"].astype(float).to_numpy()
    improvement = df["WMAPE_Improvement_Percentage_Points"].astype(float).to_numpy()
    volume     = df["Total_Volume"].astype(float).to_numpy()

    # ---------- 3. Display axis: actual weeks + future placeholders ----------
    # No future WMAPE values are invented; only the x-axis is extended.
    actual_dates = df["start_date"].tolist()
    remaining = display_weeks - len(actual_dates)
    if remaining > 0:
        future = pd.date_range(
            start=actual_dates[-1] + pd.Timedelta(days=7),
            periods=remaining,
            freq="7D",
        ).tolist()
        weeks = actual_dates + future
    else:
        weeks = actual_dates[:display_weeks]
    x_all = np.arange(len(weeks))
    date_labels = [f"{d.month}/{d.day}" for d in weeks]

    # ---------- 4. WMAPE y-axis: round out to multiples of 5 ----------
    wmape_tick_interval = 5
    wmape_data_min = min(np.nanmin(our), np.nanmin(by))
    wmape_data_max = max(np.nanmax(our), np.nanmax(by))
    wmape_y_min = np.floor(wmape_data_min / wmape_tick_interval) * wmape_tick_interval
    wmape_y_max = np.ceil(wmape_data_max / wmape_tick_interval) * wmape_tick_interval
    if wmape_y_max == wmape_data_max:
        wmape_y_max += wmape_tick_interval
    wmape_ticks = np.arange(wmape_y_min, wmape_y_max + wmape_tick_interval, wmape_tick_interval)

    # ---------- 5. Improvement y-axis: multiples of 2, 10% headroom for labels ----------
    # BUGFIX vs. original: the floor was hardcoded to 0, which clipped the bars
    # in any week where Vendor beat us (negative improvement).
    improvement_tick_interval = 2
    improvement_data_min = float(np.nanmin(improvement))
    improvement_data_max = float(np.nanmax(improvement))
    improvement_y_min = min(
        0.0,
        np.floor((improvement_data_min * 1.10) / improvement_tick_interval) * improvement_tick_interval,
    )
    improvement_y_max = np.ceil((improvement_data_max * 1.10) / improvement_tick_interval) * improvement_tick_interval
    if improvement_y_max <= improvement_data_max:
        improvement_y_max += improvement_tick_interval
    improvement_ticks = np.arange(
        improvement_y_min, improvement_y_max + improvement_tick_interval, improvement_tick_interval
    )

    # ---------- 6. Figure ----------
    fig = plt.figure(figsize=(15, 8.5), facecolor="white")
    ax_top = fig.add_axes([0.075, 0.49, 0.89, 0.40])              # [left, bottom, width, height]
    ax_bottom = fig.add_axes([0.075, 0.12, 0.89, 0.25], sharex=ax_top)

    # ---------- 7. Colors ----------
    blue, gray, green = "#1F66D1", "#747D84", "#198754"
    grid, text = "#DDE2E6", "#222222"

    # ---------- 8. Top chart: WMAPE lines + volume on the right axis ----------
    ax_top.fill_between(x_data, our, by, color="#8DB79A", alpha=0.25, zorder=1)
    ax_top.plot(x_data, by, color=gray, linewidth=2.8, marker="o", markersize=7,
                label="Vendor", zorder=3)
    ax_top.plot(x_data, our, color=blue, linewidth=2.8, marker="o", markersize=7,
                label="Our Model", zorder=4)

    ax_top_right = ax_top.twinx()
    ax_top_right.plot(x_data, volume, color="#D32F2F", linewidth=2.2, linestyle=":",
                      marker="D", markersize=5, label="Total Volume", zorder=2)
    ax_top_right.set_ylabel("Total Volume", fontsize=12, labelpad=12, color="#D32F2F")

    vol_min, vol_max = np.nanmin(volume), np.nanmax(volume)
    vol_padding = (vol_max - vol_min) * 0.3 or max(vol_max * 0.05, 1.0)  # guard single-week runs
    ax_top_right.set_ylim(vol_min - vol_padding, vol_max + vol_padding)
    ax_top_right.tick_params(axis="y", labelsize=10, colors="#D32F2F")
    ax_top_right.spines["top"].set_visible(False)
    ax_top_right.spines["left"].set_visible(False)
    ax_top_right.spines["right"].set_color("#D32F2F")

    # ---------- 9. Top axis formatting ----------
    ax_top.set_xlim(-0.5, display_weeks - 0.5)
    ax_top.set_ylim(wmape_y_min, wmape_y_max)
    ax_top.yaxis.set_major_locator(FixedLocator(wmape_ticks))
    ax_top.yaxis.set_major_formatter(FormatStrFormatter("%d"))
    ax_top.set_ylabel("WMAPE (%)", fontsize=12, labelpad=12)
    ax_top.grid(axis="y", color=grid, linewidth=0.9)

    lines_1, labels_1 = ax_top.get_legend_handles_labels()
    lines_2, labels_2 = ax_top_right.get_legend_handles_labels()
    ax_top.legend(lines_1 + lines_2, labels_1 + labels_2,
                  loc="upper right", frameon=False, fontsize=11)

    # ---------- 10. Date labels ----------
    ax_top.set_xticks(x_all)
    ax_top.set_xticklabels(date_labels, rotation=45, ha="right", fontsize=9)
    ax_top.tick_params(axis="x", bottom=True, labelbottom=True, pad=6, length=4,
                       color="#777777", labelcolor=text)
    ax_top.set_xlabel("Week Starting", fontsize=11, labelpad=8)

    # ---------- 11-12. Bottom chart: improvement bars + labels ----------
    bars = ax_bottom.bar(x_data, improvement, width=0.58, color=green,
                         edgecolor="#106A40", linewidth=0.8, zorder=3)
    for bar, value in zip(bars, improvement):
        ax_bottom.annotate(
            f"{value:.2f}%",
            xy=(bar.get_x() + bar.get_width() / 2, bar.get_height()),
            xytext=(0, 7 if value >= 0 else -16),
            textcoords="offset points",
            ha="center", va="bottom", fontsize=11, fontweight="bold",
            color=text, clip_on=False,
        )

    # ---------- 13. Bottom axis formatting ----------
    ax_bottom.set_ylim(improvement_y_min, improvement_y_max)
    ax_bottom.yaxis.set_major_locator(FixedLocator(improvement_ticks))
    ax_bottom.yaxis.set_major_formatter(FormatStrFormatter("%d"))
    ax_bottom.set_ylabel("Improvement (pp)", fontsize=12, labelpad=12)
    ax_bottom.grid(axis="y", color=grid, linewidth=0.9)
    ax_bottom.axhline(0, color="#777777", linewidth=0.9)
    ax_bottom.tick_params(axis="x", bottom=False, labelbottom=False)

    # ---------- 14. Borders ----------
    for ax in (ax_top, ax_bottom):
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color("#A7ADB3")
        ax.spines["bottom"].set_color("#A7ADB3")
        ax.tick_params(axis="y", labelsize=10, colors=text)

    # ---------- 15. Title ----------
    fig.suptitle(f"{title} ({len(df)} Weeks)", fontsize=20, fontweight="bold", y=0.965)
    fig.text(0.5, 0.925,
             f'Data through {df["start_date"].max().strftime("%B %d, %Y")}',
             ha="center", va="center", fontsize=11.5, color="#555555")

    # ---------- 16. Save (Rec 5: Volume, not the driver filesystem) ----------
    if save:
        png_name = png_name or ("WMAPE_" + "".join(
            ch if ch.isalnum() else "_" for ch in title
        ).strip("_") + f"_{FCST_DATE}.png")
        local_png = f"/tmp/{png_name}"
        fig.savefig(local_png, dpi=220, facecolor="white", bbox_inches=None)
        save_to_volume(local_png, png_name, subdir="charts")

    plt.show()
    return fig, df


print("Helpers loaded: compute_wmape_metrics, plot_wmape_chart, dq_check, dq_row_retention, save_to_volume")

# COMMAND ----------

# DBTITLE 1,Profile vendor_forecast_perf (automated thresholds)
# ─────────────────────────────────────────────────────────────────────────────
# REFACTOR NOTE (Rec 6)
# WAS: a hand-typed list of fcstdates and a printed summary table that someone
#      had to read and judge by eye.
# NOW: the comparison window is derived from FCST_DATE (DQ_PROFILE_WEEKS prior
#      weeks), and the actuals-volume drift check is asserted against
#      DQ_MIN_VOLUME_RATIO / DQ_MAX_VOLUME_RATIO instead of being eyeballed.
#      The display() calls are kept — automation added, visibility not removed.
# ─────────────────────────────────────────────────────────────────────────────
from datetime import timedelta
from pyspark.sql.functions import (
    avg, stddev, length,
    min as spark_min, max as spark_max,
)

# Rec 6: profile window derived from FCST_DATE rather than hardcoded
_anchor = datetime.strptime(FCST_DATE, "%Y-%m-%d")
profile_dates = [(_anchor - timedelta(weeks=i)).strftime("%Y-%m-%d")
                 for i in range(DQ_PROFILE_WEEKS)]
print(f"Profiling fcstdates: {sorted(profile_dates)}")

vendor_table = (
    spark.read.table(VENDOR_TABLE)
    .filter((col("__is_current") == True) & (col("__is_deleted") == False))
    .filter(col("lag") == LAG_VALUE)
    .filter(length(col("loc")) == 3)
)

print("=== Summary stats by fcstdate (lag=%d) ===" % LAG_VALUE)
summary = (
    vendor_table
    .filter(col("fcstdate").isin(profile_dates))
    .groupBy("fcstdate")
    .agg(
        count("*").alias("row_count"),
        spark_round(avg("basefcst"), 2).alias("avg_basefcst"),
        spark_round(stddev("basefcst"), 2).alias("std_basefcst"),
        spark_round(spark_min("basefcst"), 2).alias("min_basefcst"),
        spark_round(spark_max("basefcst"), 2).alias("max_basefcst"),
        spark_round(avg("basehist"), 2).alias("avg_basehist"),
        spark_round(stddev("basehist"), 2).alias("std_basehist"),
        spark_round(spark_min("basehist"), 2).alias("min_basehist"),
        spark_round(spark_max("basehist"), 2).alias("max_basehist"),
        spark_round(spark_sum("basehist"), 0).alias("total_volume"),
    )
    .orderBy("fcstdate")
)
display(summary)

# Rec 6: automated volume-drift assertion (replaces manual inspection)
print("\n=== Data quality: actuals volume drift ===")
_profile_pdf = summary.toPandas()
_current = _profile_pdf.loc[_profile_pdf["fcstdate"].astype(str) == FCST_DATE]
_prior = _profile_pdf.loc[_profile_pdf["fcstdate"].astype(str) != FCST_DATE]

if _current.empty:
    dq_check("fcstdate present", False, f"{FCST_DATE} returned no rows from {VENDOR_TABLE}")
elif _prior.empty:
    print(f"  [SKIP] no prior weeks available to compare against {FCST_DATE}")
else:
    _cur_vol = float(_current["total_volume"].iloc[0])
    _baseline = float(_prior["total_volume"].median())
    _ratio = _cur_vol / _baseline if _baseline else 0.0
    print(f"  {FCST_DATE} volume = {_cur_vol:,.0f} vs prior-week median {_baseline:,.0f}")
    dq_range("actuals volume vs prior median", _ratio,
             DQ_MIN_VOLUME_RATIO, DQ_MAX_VOLUME_RATIO, unit="x")

    _cur_rows = float(_current["row_count"].iloc[0])
    _base_rows = float(_prior["row_count"].median())
    dq_range("row count vs prior median", (_cur_rows / _base_rows) if _base_rows else 0.0,
             DQ_MIN_VOLUME_RATIO, DQ_MAX_VOLUME_RATIO, unit="x")

# Vendor WMAPE trend across the profiled weeks
print("\n=== BY WMAPE by fcstdate ===")
wmape_compare = (
    vendor_table
    .filter(col("fcstdate").isin(profile_dates))
    .filter(col("basehist") > 0)
    .groupBy("fcstdate")
    .agg(
        (spark_sum(spark_abs(col("basehist") - col("basefcst"))) / spark_sum("basehist") * 100).alias("VENDOR_WMAPE"),
        spark_sum("basehist").alias("total_volume"),
        count("*").alias("dfu_count"),
    )
    .orderBy("fcstdate")
)
display(wmape_compare)

# COMMAND ----------

# DBTITLE 1,Vendor Filter
# ─────────────────────────────────────────────────────────────────────────────
# REFACTOR NOTE (Rec 6)
# Logic unchanged. Added a row-count guard so an empty/undersized BY pull fails
# here with a clear message rather than surfacing as an empty chart ten cells later.
# The intermediate `df` reassignments were chained into one expression for readability.
# ─────────────────────────────────────────────────────────────────────────────
from pyspark.sql.functions import concat_ws

VENDOR_FULL = (
    spark.read.table(VENDOR_TABLE)
    .filter((col("__is_current") == True) & (col("__is_deleted") == False))
    .filter(length(col("loc")) == 3)
    .select("dmdunit", "dmdgroup", "loc", "fcstdate", "basefcst", "lag", "basehist", "startdate")
    .filter((col("fcstdate") == FCST_DATE) & (col("lag") == LAG_VALUE))
    .withColumns({
        "dmdunit":     col("dmdunit").cast("string"),
        "dmdgroup":    col("dmdgroup").cast("string"),
        "loc":         col("loc").cast("string"),
        "fcstdate":    col("fcstdate").cast("date"),
        "startdate":   col("startdate").cast("date"),
        "VENDOR_Basefcst": col("basefcst").cast("double"),
        "basehist":    col("basehist").cast("double"),
    })
    .dropDuplicates()
    .drop("basefcst")
    .withColumn("dfu", concat_ws("_", col("dmdunit"), col("loc"), col("dmdgroup")))
    .drop("dmdunit", "dmdgroup", "loc")
    .cache()   # reused by the join and the row-count checks below
)

VENDOR_ROWS = VENDOR_FULL.count()
print(VENDOR_FULL.dtypes)
print("Rows:", f"{VENDOR_ROWS:,}")

# Rec 6: fail fast on an empty BY pull
dq_check("VENDOR_FULL not empty", VENDOR_ROWS > 0,
         f"{VENDOR_ROWS:,} rows for fcstdate={FCST_DATE}, lag={LAG_VALUE}")

display(VENDOR_FULL)

# COMMAND ----------

# DBTITLE 1,INTERNAL Filter
# ─────────────────────────────────────────────────────────────────────────────
# REFACTOR NOTE (Rec 1, 6)
# Rec 1 — reads through read_internal(), which resolves either the dated table
#         (name built from FCST_DATE) or the single compiled table filtered on
#         its source_date partition column. No hardcoded _MM_DD anywhere.
# Rec 6 — row-count guard added, plus a duplicate-DFU check: duplicates in the
#         INTERNAL input fan out the join downstream and silently inflate volume.
# ─────────────────────────────────────────────────────────────────────────────
from pyspark.sql.functions import to_date, countDistinct


def read_internal():
    """Resolve and read the INTERNAL forecast output for this run (Rec 1)."""
    df = spark.read.table(INTERNAL_TABLE)
    if USE_COMPILED_INTERNAL:
        if INTERNAL_DATE_COLUMN not in df.columns:
            raise ValueError(
                f"{INTERNAL_TABLE} has no '{INTERNAL_DATE_COLUMN}' column; "
                "set USE_COMPILED_INTERNAL = False or fix INTERNAL_DATE_COLUMN."
            )
        df = df.filter(col(INTERNAL_DATE_COLUMN) == FCST_DATE)
        print(f"  compiled mode: filtered {INTERNAL_DATE_COLUMN} = {FCST_DATE}")
    return df


INTERNAL_FULL_SCOPE = (
    read_internal()
    .filter(to_date(col("Forecast Date (must be Sunday)")) == to_date(lit(START_DATE)))
    .withColumnsRenamed({
        "Forecast Date (must be Sunday)": "startdate",
        "Lock Value (-1 to delete lock)": "Our Base Forecast",
        "Demand Group": "dmdgroup",
        "Location": "loc",
        "Item": "dmdunit",
    })
    .withColumns({
        "startdate": col("startdate").cast("date"),
        "loc":       col("loc").cast("string"),
        "dmdgroup":  col("dmdgroup").cast("string"),
        "dmdunit":   col("dmdunit").cast("string"),
    })
    .withColumn("dfu", concat_ws("_", col("dmdunit"), col("loc"), col("dmdgroup")))
    .select("dfu", "startdate", "Our Base Forecast", "Brand Code")
    .cache()
)

INTERNAL_ROWS = INTERNAL_FULL_SCOPE.count()
print(INTERNAL_FULL_SCOPE.dtypes)
print("Rows:", f"{INTERNAL_ROWS:,}")

# Rec 6: guards on the INTERNAL input
dq_check("INTERNAL not empty", INTERNAL_ROWS > 0,
         f"{INTERNAL_ROWS:,} rows from {INTERNAL_TABLE} for startdate={START_DATE}")

_distinct_dfu = INTERNAL_FULL_SCOPE.select(countDistinct("dfu")).first()[0]
dq_check(
    "INTERNAL DFU uniqueness",
    _distinct_dfu == INTERNAL_ROWS,
    f"{_distinct_dfu:,} distinct DFUs across {INTERNAL_ROWS:,} rows "
    f"({INTERNAL_ROWS - _distinct_dfu:,} duplicates would fan out the join)",
    fail_on_error=False,   # warn only: known duplicate e.g. A69W00351-20_440_100
)

display(INTERNAL_FULL_SCOPE)

# COMMAND ----------

# DBTITLE 1,Table Join + Scope Flag
# ─────────────────────────────────────────────────────────────────────────────
# REFACTOR NOTE (Rec 2, 6)
#
# Rec 2 — WAS: scope.collect() pulled every DFU into the driver as a Python list
#         and used col("dfu").isin(scope_list). That serializes the whole list
#         into the query plan on every task and breaks down as scope grows.
#         NOW: a broadcast left join on a de-duplicated scope DataFrame. Stays
#         distributed, and .distinct() prevents the known duplicate scope DFU
#         (A69W00351-20_440_100) from multiplying rows.
#
# Rec 6 — join guarded: the joined row count must retain >= DQ_MIN_JOIN_RETENTION
#         of the INTERNAL input, and a reasonable share of the scope list must
#         actually match. Key drift now fails loudly instead of silently
#         dropping rows.
# ─────────────────────────────────────────────────────────────────────────────
print(f"INTERNAL_FULL_SCOPE rows: {INTERNAL_ROWS:,}")
print(f"VENDOR_FULL rows: {VENDOR_ROWS:,}")

Full_table = (
    INTERNAL_FULL_SCOPE
    .join(VENDOR_FULL, ["dfu", "startdate"], how="inner")
    .select("dfu", "Brand Code", "fcstdate", "startdate", "lag",
            "Our Base Forecast", "VENDOR_Basefcst", "basehist")
)

JOINED_ROWS = Full_table.cache().count()
print(f"Full_table (after join) rows: {JOINED_ROWS:,}")

# Rec 6: did the join silently drop rows?
print("\n=== Data quality: join retention ===")
dq_row_retention("INTERNAL -> joined", INTERNAL_ROWS, JOINED_ROWS)

# Rec 2: broadcast join instead of collect() + isin()
scope = (
    spark.read.table(SCOPE_TABLE)
    .select("dfu")
    .distinct()                       # guards against duplicate scope entries
    .withColumn("_in_scope", lit(True))
)
SCOPE_ROWS = scope.count()

Full_table = (
    Full_table
    .join(broadcast(scope), on="dfu", how="left")
    .withColumn("Scope", when(col("_in_scope").isNotNull(), lit("In")).otherwise(lit("Out")))
    .drop("_in_scope")
)

POST_SCOPE_ROWS = Full_table.cache().count()
IN_SCOPE_ROWS = Full_table.filter(col("Scope") == "In").count()

# Rec 6: a left join must not change the row count, and scope should mostly match
print("\n=== Data quality: scope join ===")
dq_check("scope join did not fan out", POST_SCOPE_ROWS == JOINED_ROWS,
         f"{POST_SCOPE_ROWS:,} rows after scope join vs {JOINED_ROWS:,} before")
dq_check(
    "scope DFUs matched",
    SCOPE_ROWS == 0 or (IN_SCOPE_ROWS / SCOPE_ROWS) >= DQ_MIN_SCOPE_MATCH,
    f"{IN_SCOPE_ROWS:,} in-scope rows against a scope list of {SCOPE_ROWS:,} DFUs",
    fail_on_error=False,
)

print(Full_table.dtypes)
print("Rows:", f"{POST_SCOPE_ROWS:,}")
display(Full_table.filter(col("Scope") == "In"))

# COMMAND ----------

# DBTITLE 1,The Math — Absolute Errors
# ─────────────────────────────────────────────────────────────────────────────
# REFACTOR NOTE
# Logic unchanged; column names now come from the constants defined in the
# helpers cell so a source-schema rename is a one-line change.
# ─────────────────────────────────────────────────────────────────────────────
Full_table = (
    Full_table
    .withColumn(
        OUR_ERR_COL,
        when(col(ACTUALS_COL) != 0, spark_abs(col(ACTUALS_COL) - col(OUR_FCST_COL)))
        .otherwise(lit(0.0)),
    )
    .withColumn(
        VENDOR_ERR_COL,
        when(col(ACTUALS_COL) != 0, spark_abs(col(ACTUALS_COL) - col(VENDOR_FCST_COL)))
        .otherwise(lit(0.0)),
    )
)
display(Full_table)

# COMMAND ----------

# DBTITLE 1,Add Demand Classification
# ─────────────────────────────────────────────────────────────────────────────
# REFACTOR NOTE (Rec 2, 6)
# Rec 2 — the classification lookup is small and keyed on dfu, so it is
#         broadcast as well; this avoids a shuffle on the fact-side DataFrame.
# Rec 6 — a left join must preserve the row count; asserted here because a
#         duplicated dfu in demand_classification_by_dfu would double-count volume.
# ─────────────────────────────────────────────────────────────────────────────
Demand_spark = (
    spark.read.table(DEMAND_TABLE)
    .select("dfu", "demand_class")
    .dropDuplicates(["dfu"])          # one class per DFU; protects against fan-out
)

Full_table_classified = Full_table.join(broadcast(Demand_spark), on="dfu", how="left").cache()

CLASSIFIED_ROWS = Full_table_classified.count()
print("Rows:", f"{CLASSIFIED_ROWS:,}")

print("\n=== Data quality: demand-class join ===")
dq_check("demand-class join did not fan out", CLASSIFIED_ROWS == POST_SCOPE_ROWS,
         f"{CLASSIFIED_ROWS:,} rows after join vs {POST_SCOPE_ROWS:,} before")

_unclassified = Full_table_classified.filter(col("demand_class").isNull()).count()
dq_check("demand_class coverage",
         CLASSIFIED_ROWS == 0 or (_unclassified / CLASSIFIED_ROWS) <= 0.10,
         f"{_unclassified:,} of {CLASSIFIED_ROWS:,} rows have no demand_class",
         fail_on_error=False)

display(Full_table_classified)

# COMMAND ----------

# DBTITLE 1,WMAPE & Bias Breakdowns
# ─────────────────────────────────────────────────────────────────────────────
# REFACTOR NOTE (Rec 3)
# WAS: the same 9-metric .agg(...) block copy-pasted five times (~110 lines).
#      Any metric fix had to be applied in five places, and the wmape_total copy
#      had already drifted (a hardcoded lit(100.0) for Pct_of_Total_Volume).
# NOW: one call to compute_wmape_metrics() per breakdown. The five results are
#      built from a small spec list so adding a sixth breakdown is one line.
#      total_volume is computed once and shared, so every Pct_of_Total_Volume
#      uses the same denominator.
# ─────────────────────────────────────────────────────────────────────────────
from pyspark.sql.functions import split as spark_split

# Extract dmdgroup back out of the dfu key (format: dmdunit_loc_dmdgroup)
Full_table_classified = Full_table_classified.withColumn(
    "dmdgroup", spark_split(col("dfu"), "_", 3)[2]
)

# Rows with real actuals only — WMAPE is undefined against zero volume
base_df = Full_table_classified.filter(col(ACTUALS_COL) > 0).cache()
total_volume = base_df.select(spark_sum(col(ACTUALS_COL))).first()[0]
print(f"Base rows: {base_df.count():,}   Total volume: {total_volume:,.0f}")

# (variable_name, group_cols, heading, drop_null_groups)
BREAKDOWNS = [
    ("wmape_by_scope",      "Scope",        "WMAPE & Bias by Scope",                False),
    ("wmape_by_demand",     "demand_class", "WMAPE & Bias by Demand Classification", True),
    ("wmape_total",         None,           "WMAPE & Bias Total",                    False),
    ("wmape_by_customer",   "dmdgroup",     "WMAPE & Bias by Customer",              True),
    ("wmape_by_brand_code", "Brand Code",   "WMAPE & Bias by Brand Code",            True),
]

wmape_results = {}
for var_name, group_cols, heading, drop_nulls in BREAKDOWNS:
    print("\n" + "=" * 50)
    print(heading)
    print("=" * 50)
    result = compute_wmape_metrics(
        base_df,
        group_cols=group_cols,
        total_volume=total_volume,
        drop_null_groups=drop_nulls,
    )
    wmape_results[var_name] = result
    display(result)

# Keep the original variable names bound for downstream cells / ad-hoc use
wmape_by_scope      = wmape_results["wmape_by_scope"]
wmape_by_demand     = wmape_results["wmape_by_demand"]
wmape_total         = wmape_results["wmape_total"]
wmape_by_customer   = wmape_results["wmape_by_customer"]
wmape_by_brand_code = wmape_results["wmape_by_brand_code"]

# COMMAND ----------

# DBTITLE 1,Forecast Winner by DFU and Fcstdate
# ─────────────────────────────────────────────────────────────────────────────
# REFACTOR NOTE
# Unchanged apart from using the error-column constants. Kept as its own
# function because it is genuinely reused (row level here, DFU level in the
# forecast_wins cell below).
# ─────────────────────────────────────────────────────────────────────────────
def add_forecast_winner(df: DataFrame) -> DataFrame:
    """Flag, per row, whether our model beat Vendor.

    Win  = our absolute error < BY absolute error
    Loss = our absolute error > BY absolute error
    Tie  = equal absolute errors (includes rows where basehist == 0)
    """
    return df.withColumn(
        "Forecast_Winner",
        when(col(OUR_ERR_COL) < col(VENDOR_ERR_COL), lit("Our Model"))
        .when(col(OUR_ERR_COL) > col(VENDOR_ERR_COL), lit("Vendor"))
        .otherwise(lit("Tie")),
    )


Full_table_results = add_forecast_winner(Full_table_classified).cache()

print("=== Win / Loss / Tie Summary ===")
display(
    Full_table_results
    .groupBy("Forecast_Winner")
    .count()
    .orderBy("Forecast_Winner")
)

print("\n=== Detail by DFU & Fcstdate ===")
display(
    Full_table_results
    .select("dfu", "fcstdate", "startdate", ACTUALS_COL,
            OUR_FCST_COL, VENDOR_FCST_COL, OUR_ERR_COL, VENDOR_ERR_COL, "Forecast_Winner")
    .orderBy("dfu", "fcstdate")
)

# COMMAND ----------

# DBTITLE 1,Write forecast_wins to lookup table (no source mutation)
# ─────────────────────────────────────────────────────────────────────────────
# REFACTOR NOTE (Rec 7)
#
# WAS: ALTER TABLE + UPDATE ... SET forecast_wins = NULL + MERGE INTO the INTERNAL
#      input table. That made the notebook's input depend on its own output
#      (circular), made reruns non-idempotent (the column was reset to NULL
#      mid-run, so a failure left the source table in a half-written state), and
#      required write access to a table this notebook only ought to read.
#
# NOW: DFU-level wins are written to a separate lookup table, keyed by
#      (dfu, start_date, fcst_date). Delete-then-append on the run key makes it
#      idempotent — rerun as many times as you like, the source table is never
#      touched. Downstream consumers join to this table instead of reading a
#      column off the INTERNAL output.
#
# MOVED: this cell now runs BEFORE the ≥75% win-rate analysis, which reads the
#      lookup table (previously it read the flag the old MERGE had stamped).
# ─────────────────────────────────────────────────────────────────────────────
from pyspark.sql.functions import current_timestamp

# 1. Overall winner per DFU (lower total absolute error wins)
dfu_wins = (
    Full_table_results
    .groupBy("dfu")
    .agg(
        spark_sum(OUR_ERR_COL).alias("our_abs_error"),
        spark_sum(VENDOR_ERR_COL).alias("vendor_abs_error"),
    )
    .withColumn(
        "forecast_wins",
        when(col("our_abs_error") < col("vendor_abs_error"), lit("W"))
        .when(col("our_abs_error") > col("vendor_abs_error"), lit("L"))
        .otherwise(lit("T")),
    )
    .withColumns({
        "start_date":    lit(START_DATE),
        "fcst_date":     lit(FCST_DATE),
        "lag_value":     lit(LAG_VALUE),
        "run_timestamp": current_timestamp(),
    })
    .select("dfu", "forecast_wins", "our_abs_error", "vendor_abs_error",
            "start_date", "fcst_date", "lag_value", "run_timestamp")
)

# 2. Idempotent write: clear this run key, then append
if spark.catalog.tableExists(WINS_LOOKUP_TABLE):
    spark.sql(f"""
        DELETE FROM {WINS_LOOKUP_TABLE}
        WHERE fcst_date = '{FCST_DATE}'
          AND start_date = '{START_DATE}'
          AND lag_value = {LAG_VALUE}
    """)

(
    dfu_wins
    .write
    .format("delta")
    .mode("append")
    .option("mergeSchema", "true")
    .saveAsTable(WINS_LOOKUP_TABLE)
)

# 3. Verify
_wins_this_run = spark.read.table(WINS_LOOKUP_TABLE).filter(
    (col("fcst_date") == FCST_DATE) & (col("start_date") == START_DATE)
)
WINS_COUNT = _wins_this_run.filter(col("forecast_wins") == "W").count()
DFU_COUNT = _wins_this_run.count()
print(f"✓ {WINS_LOOKUP_TABLE}")
print(f"  {WINS_COUNT:,} of {DFU_COUNT:,} DFUs won by our model for start_date={START_DATE}")
display(_wins_this_run.orderBy(col("our_abs_error").asc()).limit(20))

# COMMAND ----------

# DBTITLE 1,DFUs where Our Model wins 75%+ of evaluated weeks
# ─────────────────────────────────────────────────────────────────────────────
# REFACTOR NOTE (Rec 2, 7)
#
# Rec 7 — WAS: read the forecast_wins column that the old MERGE had stamped onto
#         the compiled INTERNAL table. NOW: reads the wins lookup table written in
#         the previous cell. Set WINS_BACKFILL_FROM_COMPILED = True during the
#         cutover to union in weeks that were stamped the old way, so the
#         win-rate history is not lost; set it to False once backfilled.
#
# Rec 2 — WAS: eval_week_list = [r.source_date for r in ...collect()] then
#         .isin(eval_week_list). NOW: a distinct semi-join, so the week list is
#         never materialized in the driver.
# ─────────────────────────────────────────────────────────────────────────────
weekly_outcomes = (
    spark.read.table(WINS_LOOKUP_TABLE)
    .select(
        col("dfu"),
        col("start_date").cast("string").alias("source_date"),
        when(col("forecast_wins") == "W", 1).otherwise(0).alias("won_this_week"),
    )
)

if WINS_BACKFILL_FROM_COMPILED and spark.catalog.tableExists(COMPILED_TABLE):
    compiled = spark.read.table(COMPILED_TABLE)
    if "forecast_wins" in compiled.columns:
        legacy = (
            compiled
            .filter(col("forecast_wins").isNotNull())
            .withColumn("dfu", concat_ws("_", col("Item"),
                                         col("Location").cast("string"),
                                         col("`Demand Group`")))
            .groupBy("dfu", col("source_date").cast("string").alias("source_date"))
            .agg(spark_max(when(col("forecast_wins") == "W", 1).otherwise(0)).alias("won_this_week"))
        )
        # Lookup table wins on conflict: drop legacy rows for weeks already written
        already = weekly_outcomes.select("source_date").distinct()
        legacy = legacy.join(already, on="source_date", how="left_anti")
        weekly_outcomes = weekly_outcomes.unionByName(legacy)
        print("  backfilled historical weeks from the compiled table")

# Collapse to one row per DFU per week (a DFU won a week if any row won it)
dfu_week_outcome = (
    weekly_outcomes
    .groupBy("dfu", "source_date")
    .agg(spark_max("won_this_week").alias("won_this_week"))
)

EVAL_WEEKS = dfu_week_outcome.select("source_date").distinct().count()
print(f"Evaluated weeks: {EVAL_WEEKS}")

dfu_win_rates = (
    dfu_week_outcome
    .groupBy("dfu")
    .agg(
        count("*").alias("total_weeks"),
        spark_sum(col("won_this_week")).alias("our_wins"),
    )
    .withColumn("by_wins", col("total_weeks") - col("our_wins"))
    .withColumn("our_win_rate_pct", spark_round(col("our_wins") / col("total_weeks") * 100, 1))
    .orderBy(col("our_win_rate_pct").desc())
    .cache()
)

dfu_75pct_winners = dfu_win_rates.filter(col("our_win_rate_pct") >= 75)
TOTAL_DFUS = dfu_win_rates.count()
WINNING_DFUS = dfu_75pct_winners.count()
print(f"{WINNING_DFUS:,} of {TOTAL_DFUS:,} DFUs have Our Model winning >= 75% of evaluated weeks")
display(dfu_75pct_winners)

# COMMAND ----------

# DBTITLE 1,Persist 75% winners to Delta table
# ─────────────────────────────────────────────────────────────────────────────
# REFACTOR NOTE
# Logic unchanged (still a full overwrite so the table always reflects the
# latest evaluation). Added the evaluated-week count and run parameters as
# columns so a reader can tell which run produced the table without opening
# the notebook.
# ─────────────────────────────────────────────────────────────────────────────
prepped_winners = dfu_75pct_winners.withColumns({
    "evaluated_weeks": lit(EVAL_WEEKS),
    "fcst_date":       lit(FCST_DATE),
    "run_timestamp":   current_timestamp(),
})

(
    prepped_winners
    .write
    .format("delta")
    .mode("overwrite")
    .option("overwriteSchema", "true")
    .saveAsTable(WINNERS_TABLE)
)

_row_count = spark.read.table(WINNERS_TABLE).count()
print(f"✓ {WINNERS_TABLE} — {_row_count:,} DFUs with >=75% win rate across {EVAL_WEEKS} evaluated weeks")

# COMMAND ----------

# DBTITLE 1,Persist DataFrames to forecasting tables
# ─────────────────────────────────────────────────────────────────────────────
# REFACTOR NOTE (Rec 3, 6)
# Rec 3 — the write list is now built from the wmape_results dict produced by
#         compute_wmape_metrics(), so a new breakdown added to BREAKDOWNS is
#         persisted automatically instead of needing a second edit here.
# Rec 6 — each write is verified: the row count read back must match what was
#         written, which catches a partial write or a schema-merge surprise.
# ─────────────────────────────────────────────────────────────────────────────
def _sanitize_cols(df: DataFrame) -> DataFrame:
    """Delta-safe column names (spaces -> underscores)."""
    for c in df.columns:
        if " " in c:
            df = df.withColumnRenamed(c, c.replace(" ", "_"))
    return df


def _add_meta(df: DataFrame) -> DataFrame:
    """Stamp every persisted table with the run parameters."""
    return _sanitize_cols(df).withColumns({
        "run_timestamp": current_timestamp(),
        "fcst_date":     lit(FCST_DATE),
        "start_date":    lit(START_DATE),
        "lag_value":     lit(LAG_VALUE),
    })


# Suffix per breakdown, driven by the same spec used to compute them
TABLE_SUFFIXES = {
    "wmape_by_scope":      "wmape_by_scope",
    "wmape_by_demand":     "wmape_by_demand_class",
    "wmape_total":         "wmape_total",
    "wmape_by_customer":   "wmape_by_customer",
    "wmape_by_brand_code": "wmape_by_brand_code",
}

tables_to_write = [(wmape_results[k], suffix) for k, suffix in TABLE_SUFFIXES.items()]
tables_to_write.append((Full_table_results, "forecast_comparison_detail"))

written_tables = []
for df, suffix in tables_to_write:
    table_name = f"{SCHEMA}.comparison_{suffix}"
    prepped = _add_meta(df)
    expected = prepped.count()

    # Idempotent: clear this run key before appending
    if spark.catalog.tableExists(table_name):
        spark.sql(f"""
            DELETE FROM {table_name}
            WHERE fcst_date = '{FCST_DATE}'
              AND start_date = '{START_DATE}'
              AND lag_value = {LAG_VALUE}
        """)

    (
        prepped
        .write
        .format("delta")
        .mode("append")
        .option("mergeSchema", "true")
        .saveAsTable(table_name)
    )

    actual = spark.read.table(table_name).filter(
        (col("fcst_date") == FCST_DATE) & (col("start_date") == START_DATE)
    ).count()

    # Rec 6: verify the write landed
    dq_check(f"write {suffix}", actual == expected,
             f"{actual:,} rows read back vs {expected:,} written")
    written_tables.append(table_name)
    print(f"✓ {table_name} — {actual:,} rows for fcst_date={FCST_DATE}")

print("\nAll tables written successfully.")

# COMMAND ----------

# DBTITLE 1,Chart — In Scope
# ─────────────────────────────────────────────────────────────────────────────
# REFACTOR NOTE (Rec 4, 5)
# WAS: ~200 lines of chart code inline in this cell, duplicated almost verbatim
#      in the next one. Any styling fix had to be made twice and they had
#      already begun to drift.
# NOW: one call to plot_wmape_chart() (defined in the helpers cell). The PNG is
#      written to the UC Volume (Rec 5) rather than the driver filesystem, so it
#      survives cluster termination and can be linked to from a deck or email.
# ─────────────────────────────────────────────────────────────────────────────
fig_in_scope, df_in_scope = plot_wmape_chart(
    table_name=f"{SCHEMA}.comparison_wmape_by_scope",
    filter_expr="Scope = 'In'",
    title="Lag 3 - In-Scope Total",
    png_name=f"WMAPE_Chart_InScope_{FCST_DATE}.png",
)

# COMMAND ----------

# DBTITLE 1,Chart — Total
# ─────────────────────────────────────────────────────────────────────────────
# REFACTOR NOTE (Rec 4, 5)
# Same consolidation as the previous cell: this chart differs from the in-scope
# one only in source table, filter and title, so it is the same function call
# with different arguments. PNG also goes to the Volume.
# ─────────────────────────────────────────────────────────────────────────────
fig_total, df_total = plot_wmape_chart(
    table_name=f"{SCHEMA}.comparison_wmape_total",
    filter_expr=None,
    title="Lag 3 - Total",
    png_name=f"WMAPE_Chart_Total_{FCST_DATE}.png",
)

# COMMAND ----------

# DBTITLE 1,Export Summary Excel to Volume
# ─────────────────────────────────────────────────────────────────────────────
# REFACTOR NOTE (Rec 5, 6)
#
# Rec 5 — WAS: pd.ExcelWriter wrote to a bare relative filename, which lands on
#         the driver's local disk and disappears when the cluster terminates —
#         nobody but the notebook author could ever retrieve it.
#         NOW: written to /tmp (openpyxl needs a seekable handle) and copied to
#         the UC Volume, where it persists and inherits Unity Catalog
#         permissions for sharing. The final path is printed for handoff.
#
# Rec 6 — every sheet is checked for rows; an empty summary workbook is a
#         silent failure that stakeholders would only notice after opening it.
# ─────────────────────────────────────────────────────────────────────────────
timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
output_file = f"Forecast_Comparison_Summary_{FCST_DATE}_{timestamp_str}.xlsx"
local_path = f"/tmp/{output_file}"

latest_filter = f"fcst_date = '{FCST_DATE}' AND start_date = '{START_DATE}'"

sheets = {
    "WMAPE by Scope":        spark.read.table(f"{SCHEMA}.comparison_wmape_by_scope").filter(latest_filter).toPandas(),
    "WMAPE by Demand Class": spark.read.table(f"{SCHEMA}.comparison_wmape_by_demand_class").filter(latest_filter).toPandas(),
    "WMAPE Total":           spark.read.table(f"{SCHEMA}.comparison_wmape_total").filter(latest_filter).toPandas(),
    "WMAPE by Customer":     spark.read.table(f"{SCHEMA}.comparison_wmape_by_customer").filter(latest_filter).toPandas(),
    "WMAPE by Brand Code":   spark.read.table(f"{SCHEMA}.comparison_wmape_by_brand_code").filter(latest_filter).toPandas(),
}

# Rec 6: no silently-empty workbook
_empty = [name for name, pdf in sheets.items() if pdf.empty]
dq_check("all export sheets populated", not _empty,
         f"empty sheets: {_empty}" if _empty else "every sheet has rows",
         fail_on_error=False)

meta_cols = ["run_timestamp", "fcst_date", "start_date", "lag_value"]
with pd.ExcelWriter(local_path, engine="openpyxl") as writer:
    for sheet_name, pdf in sheets.items():
        drop = [c for c in meta_cols if c in pdf.columns]
        pdf.drop(columns=drop, errors="ignore").to_excel(writer, sheet_name=sheet_name, index=False)

# Rec 5: persist to the Volume
EXCEL_PATH = save_to_volume(local_path, output_file, subdir="summaries")
print(f"✓ Summary Excel: {EXCEL_PATH}")
print(f"  Sheets: {list(sheets.keys())}")

# COMMAND ----------

# DBTITLE 1,Run Summary
# ─────────────────────────────────────────────────────────────────────────────
# REFACTOR NOTE (Rec 8) — NEW CELL
# Auto-generates a plain-language summary of the run so the notebook can be read
# top-to-bottom without interpreting the display() grids. Rendered as formatted
# markdown/HTML and also printed as raw markdown so it can be pasted straight
# into Teams, email or a ticket.
# ─────────────────────────────────────────────────────────────────────────────
_total_row = wmape_total.first().asDict()
_scope_rows = {r["Scope"]: r.asDict() for r in wmape_by_scope.collect()}
_in_scope = _scope_rows.get("In", {})


def _fmt(value, suffix="", decimals=2):
    return f"{value:,.{decimals}f}{suffix}" if value is not None else "n/a"


_overall_delta = _total_row.get("WMAPE_Improvement_Percentage_Points")
_direction = "ahead of" if (_overall_delta or 0) > 0 else "behind"

summary_md = f"""
## Forecast Performance Summary — {START_DATE} (lag {LAG_VALUE})

**Run parameters** — INTERNAL table `{INTERNAL_TABLE}` · vendor fcstdate `{FCST_DATE}` · forecast week `{START_DATE}`

### Headline
Our model is **{_fmt(abs(_overall_delta or 0), ' pp')} {_direction}** Vendor on overall WMAPE.

| Metric | Our Model | Vendor | Delta |
|---|---|---|---|
| WMAPE — all DFUs | {_fmt(_total_row.get('Our_WMAPE'), '%')} | {_fmt(_total_row.get('VENDOR_WMAPE'), '%')} | {_fmt(_overall_delta, ' pp')} |
| WMAPE — in scope | {_fmt(_in_scope.get('Our_WMAPE'), '%')} | {_fmt(_in_scope.get('VENDOR_WMAPE'), '%')} | {_fmt(_in_scope.get('WMAPE_Improvement_Percentage_Points'), ' pp')} |
| Bias | {_fmt(_total_row.get('Our_Bias'), '%')} | {_fmt(_total_row.get('VENDOR_Bias'), '%')} | — |
| Row-level win rate | {_fmt(_total_row.get('Our_Win_Rate_Pct'), '%', 1)} | {_fmt(_total_row.get('VENDOR_Win_Rate_Pct'), '%', 1)} | — |

### Coverage
- **{DFU_COUNT:,}** DFUs evaluated this week; **{WINS_COUNT:,}** ({(WINS_COUNT / DFU_COUNT * 100) if DFU_COUNT else 0:.1f}%) won by our model.
- **{WINNING_DFUS:,} of {TOTAL_DFUS:,}** DFUs win at least 75% of the **{EVAL_WEEKS}** evaluated weeks.
- Total actuals volume: **{_fmt(_total_row.get('Total_Volume'), '', 0)}** across **{POST_SCOPE_ROWS:,}** joined rows
  ({IN_SCOPE_ROWS:,} in scope).

### Outputs
- Delta tables: `{SCHEMA}.comparison_*` ({len(written_tables)} tables)
- Wins lookup: `{WINS_LOOKUP_TABLE}` · 75% winners: `{WINNERS_TABLE}`
- Excel: `{EXCEL_PATH}`
- Charts: `{EXPORT_VOLUME}/charts/`

_Weeks excluded from charts: {', '.join(EXCLUDE_WEEKS) or 'none'}._
"""

print(summary_md)

# Rendered view for stakeholders reading the notebook
_rows_html = "".join(
    f"<tr><td style='padding:6px 14px;border-bottom:1px solid #eee'>{label}</td>"
    f"<td style='padding:6px 14px;border-bottom:1px solid #eee;text-align:right'>{ours}</td>"
    f"<td style='padding:6px 14px;border-bottom:1px solid #eee;text-align:right'>{theirs}</td></tr>"
    for label, ours, theirs in [
        ("WMAPE — all DFUs", _fmt(_total_row.get("Our_WMAPE"), "%"), _fmt(_total_row.get("VENDOR_WMAPE"), "%")),
        ("WMAPE — in scope", _fmt(_in_scope.get("Our_WMAPE"), "%"), _fmt(_in_scope.get("VENDOR_WMAPE"), "%")),
        ("Bias", _fmt(_total_row.get("Our_Bias"), "%"), _fmt(_total_row.get("VENDOR_Bias"), "%")),
        ("Win rate", _fmt(_total_row.get("Our_Win_Rate_Pct"), "%", 1), _fmt(_total_row.get("VENDOR_Win_Rate_Pct"), "%", 1)),
    ]
)

displayHTML(f"""
<div style="font-family:-apple-system,Segoe UI,Helvetica,sans-serif;max-width:760px">
  <h2 style="margin-bottom:4px">Forecast Performance — {START_DATE} (lag {LAG_VALUE})</h2>
  <p style="color:#555;margin-top:0">INTERNAL <code>{INTERNAL_TABLE}</code> · vendor fcstdate <code>{FCST_DATE}</code></p>
  <p style="font-size:18px">Our model is <b>{_fmt(abs(_overall_delta or 0), ' pp')} {_direction}</b> Vendor on overall WMAPE.</p>
  <table style="border-collapse:collapse;width:100%;font-size:14px">
    <tr style="background:#f4f6f8">
      <th style="padding:8px 14px;text-align:left">Metric</th>
      <th style="padding:8px 14px;text-align:right">Our Model</th>
      <th style="padding:8px 14px;text-align:right">Vendor</th>
    </tr>
    {_rows_html}
  </table>
  <ul style="font-size:14px;color:#333;line-height:1.7">
    <li><b>{WINS_COUNT:,}</b> of <b>{DFU_COUNT:,}</b> DFUs won this week</li>
    <li><b>{WINNING_DFUS:,}</b> of <b>{TOTAL_DFUS:,}</b> DFUs win ≥75% of {EVAL_WEEKS} evaluated weeks</li>
    <li>Excel: <code>{EXCEL_PATH}</code></li>
  </ul>
</div>
""")

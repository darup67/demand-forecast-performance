# Demand forecast performance analysis

A Databricks notebook that compares two weekly demand forecasts, an **internal model** and a **vendor** forecast,
against actuals, and reports where each one wins.

## What it does

* Reads the internal model's forecast output and the vendor's forecast-performance table for a chosen forecast date and lag.
* Joins them at the DFU (demand forecasting unit) level and flags whether each DFU is in scope.
* Computes **WMAPE** and bias by scope, customer, brand code and demand class, plus win rates for each forecast
  (`compute_wmape_metrics()` is the single aggregation used everywhere).
* Plots a weekly WMAPE comparison (`plot_wmape_chart()`), exports Excel and PNG files to a Unity Catalog volume,
  and writes a short markdown summary.
* Writes DFU-level wins to a separate lookup table instead of modifying the source table.
* Data-quality guards on every join and on volume drift, with automated profiling thresholds.

## Using it

All table, schema and volume names are placeholders (`my_catalog.*`). Edit the configuration cell:

| Setting | Meaning |
|---|---|
| `FCST_DATE`, `START_DATE`, `LAG_VALUE` | which forecast run and week to analyse |
| `INTERNAL_SCHEMA`, `INTERNAL_TABLE_PREFIX` | where the internal model's output lives |
| `USE_COMPILED_INTERNAL` | read one compiled table, or derive a dated table name from `FCST_DATE` |
| `VENDOR_TABLE`, `SCOPE_TABLE`, `DEMAND_TABLE` | vendor performance table, scope list, demand classes |
| `EXPORT_VOLUME` | Unity Catalog volume for exports |

Column names in the vendor and internal tables follow the code (`basefcst`, `basehist`, `dfu`, `startdate`, ...); adjust the
reader cells if your schema differs. Runs on Databricks (DBR 13+); `%pip install openpyxl` is the first cell.

## Notes

* This is a cleaned-up, generic version: organization-specific table names and labels were replaced with placeholders,
  and it ships with no data.
* The notebook is stored as a Databricks source file (`.py` with `# COMMAND ----------` separators); import it into a workspace to run it.

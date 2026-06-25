# OCI Migration Estimator

A Streamlit app that analyses a VMware RVTools export and generates a like-for-like OCI migration cost estimate using live Oracle pricing API.

## What it does

- Parses RVTools `.xlsx` export (vHost + vInfo sheets)
- Shows infrastructure summary: hosts, cores, memory, CPU/memory utilisation
- Builds a Migration BOM for workload VMs (VMware infra VMs auto-excluded)
- Prices compute, block storage, and Windows OS licensing using the live OCI pricing API
- Supports per-VM toggles: exclude from migration, custom run hours, BYOL Windows licensing
- Calculates RackWare migration cost (OCI Marketplace PAYGO)
- Estimates object storage cost for backups
- Exports an Excel file in OCI Investment Proposal format

## Requirements

- Python 3.8+
- Internet access (fetches live OCI pricing at runtime)

## Install and run

```bash
git clone https://github.com/kmittal8/-rvtools-estimator.git
cd -rvtools-estimator
pip install -r requirements.txt
streamlit run rvtools_app.py
```

Browser opens at `http://localhost:8501`. Upload your RVTools `.xlsx` export to begin.

## Input

Export from RVTools (File → Export → Export All to xlsx). The app reads the `vHost` and `vInfo` sheets.

## Disclaimer

This tool is not an official Oracle product. Always validate pricing with the [OCI Cost Estimator](https://www.oracle.com/anz/cloud/costestimator.html) before sharing with customers.

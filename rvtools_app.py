import streamlit as st
import pandas as pd
import requests
from io import BytesIO

st.set_page_config(page_title="RVTools Analyzer", page_icon="🖥️", layout="wide")
st.title("🖥️ RVTools Analyzer")
st.caption("Upload an RVTools Excel export to get infrastructure summary.")

st.warning(
    "⚠️ **This tool is not an official Oracle product.** "
    "Always validate pricing with the official Oracle Cost Estimator: "
    "[cloud.oracle.com/cost-estimator](https://www.oracle.com/anz/cloud/costestimator.html)",
    icon=None
)

uploaded_file = st.file_uploader("Upload RVTools Excel file", type=["xlsx"])

VMWARE_NAME_PREFIXES = ('vcls-', 'z-vrah-', 'z-vra-')
HOURS_PER_MONTH = 744
BALANCED_VPU = 10  # OCI Balanced performance tier = 10 VPU

# OCI SKUs
SKU_E4_OCPU   = 'B93113'
SKU_E4_MEM    = 'B93114'
SKU_BLOCK_STG = 'B91961'
SKU_BLOCK_VPU = 'B91962'
SKU_OBJ_STG   = 'B96484'

@st.cache_data(ttl=3600)
def fetch_oci_prices(currency):
    url = f"https://apexapps.oracle.com/pls/apex/cetools/api/v1/products/?currencyCode={currency}"
    r = requests.get(url, timeout=10)
    r.raise_for_status()
    items = r.json()['items']
    return {i['partNumber']: i['currencyCodeLocalizations'][0]['prices'][0]['value'] for i in items}

def os_family(name):
    n = str(name).lower()
    if 'windows' in n:
        return 'Windows'
    elif any(x in n for x in ['linux', 'debian', 'ubuntu', 'suse', 'red hat', 'centos', 'photon']):
        return 'Linux'
    else:
        return 'Other'

DISCLAIMER = (
    "Disclaimer:  This sample quote is provided solely for evaluation purposes and is intended to further "
    "discussions between you and Oracle.  This sample quote is not eligible for acceptance by you and is not "
    "a binding contract between you and Oracle for the services specified.  If you would like to purchase the "
    "services specified in this sample quote, please request that Oracle issue you a formal quote (which may "
    "include an OMA or a CSA if you do not already have an appropriate agreement in place with Oracle) for "
    "your acceptance and execution.  Your formal quote will be effective only upon Oracle's acceptance of the "
    "formal quote (and the OMA or CSA, if required)."
)

SKU_DESCRIPTIONS = {
    SKU_E4_OCPU:   ("Compute - Standard - E4 - OCPU (OCPU Per Hour)\nCapacity Type: On-Demand",   "OCPU Per Hour"),
    SKU_E4_MEM:    ("Compute - Standard - E4 - Memory (Gigabyte Per Hour)\nCapacity Type: On-Demand", "Gigabyte Per Hour"),
    SKU_BLOCK_STG: ("Storage - Block Volume - Storage (Gigabyte Storage Capacity Per Month)",      "Gigabyte Storage Capacity Per Month"),
    SKU_BLOCK_VPU: ("Storage - Block Volume - Performance Units (Performance Units Per Gigabyte Per Month)", "Performance Units Per Gigabyte Per Month"),
}

def to_excel(df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Migration BOM')
    return buf.getvalue()

def to_oci_excel(bom_df, prices, currency, shape):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
    from datetime import date

    wb = Workbook()
    ws = wb.active
    ws.title = "Migration BOM"

    # Column widths
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 60
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 13
    ws.column_dimensions['E'].width = 11
    ws.column_dimensions['F'].width = 13
    ws.column_dimensions['G'].width = 14
    ws.column_dimensions['H'].width = 15
    ws.column_dimensions['I'].width = 15

    header_font   = Font(bold=True)
    red_font      = Font(bold=True, color="FF0000")
    group_font    = Font(bold=True)
    light_blue    = PatternFill("solid", fgColor="DCE6F1")
    wrap          = Alignment(wrap_text=True, vertical="top")
    thin_border   = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'),  bottom=Side(style='thin')
    )

    row = 1
    today = date.today().strftime("%m/%d/%Y")
    ws.cell(row=row, column=1, value=f"Oracle Investment Proposal (as of {today})").font = Font(bold=True, size=12)
    row += 1
    ws.cell(row=row, column=1, value="Reference label: Migration Estimate")
    row += 1
    ws.cell(row=row, column=1, value=f"Currency: {currency}")
    row += 1
    ws.cell(row=row, column=1, value="Realm: PUBLIC")
    row += 1
    ws.cell(row=row, column=1, value="Service Type: IAAS")
    row += 1

    # Header row
    headers = ["Part", "Description", "Part Qty", "Instance Qty", "Usage Qty", "Unit Price", "Monthly Cost", "Custom Label", "Custom Note"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.font = header_font
        cell.fill = light_blue
    row += 1

    ocpu_rate  = prices[SKU_E4_OCPU]
    mem_rate   = prices[SKU_E4_MEM]
    stg_rate   = prices[SKU_BLOCK_STG]
    vpu_rate   = prices[SKU_BLOCK_VPU]
    grand_total = 0.0

    for _, vm in bom_df.iterrows():
        ocpus   = int(vm['OCPUs'])
        mem_gb  = float(vm['Memory GB'])
        stg_gb  = float(vm['Provisioned GB'])
        vpu_qty = int(stg_gb * BALANCED_VPU)

        vm_compute = (ocpus * ocpu_rate + mem_gb * mem_rate) * HOURS_PER_MONTH
        win_os_rate = prices.get('B88318', 0)
        is_windows  = str(vm.get('OS Family', '')).lower() == 'windows'
        byol        = bool(vm.get('BYOL', False))
        win_os_cost = (ocpus * win_os_rate * HOURS_PER_MONTH) if (is_windows and not byol) else 0

        vm_storage = stg_gb * stg_rate + vpu_qty * vpu_rate
        vm_total   = vm_compute + vm_storage + win_os_cost
        grand_total += vm_total

        # VM name group header
        os_label = f"{vm['VM']}  ({'Windows BYOL' if byol else vm.get('OS Family','')}{' — License Included' if is_windows and not byol else ''})"
        ws.cell(row=row, column=2, value=os_label).font = group_font
        row += 1

        # Sub-group: Virtual Machine
        ws.cell(row=row, column=2, value="Virtual Machine").font = group_font
        row += 1

        # OCPU row
        ocpu_cost = ocpus * ocpu_rate * HOURS_PER_MONTH
        for col, val in enumerate([SKU_E4_OCPU, SKU_DESCRIPTIONS[SKU_E4_OCPU][0], ocpus, 1, HOURS_PER_MONTH, ocpu_rate, round(ocpu_cost, 7)], 1):
            c = ws.cell(row=row, column=col, value=val)
            c.alignment = wrap
        ws.row_dimensions[row].height = 30
        row += 1

        # Memory row
        mem_cost = mem_gb * mem_rate * HOURS_PER_MONTH
        for col, val in enumerate([SKU_E4_MEM, SKU_DESCRIPTIONS[SKU_E4_MEM][0], mem_gb, 1, HOURS_PER_MONTH, mem_rate, round(mem_cost, 7)], 1):
            c = ws.cell(row=row, column=col, value=val)
            c.alignment = wrap
        ws.row_dimensions[row].height = 30
        row += 1

        # Sub-group: Boot Volume
        ws.cell(row=row, column=2, value="Boot Volume").font = group_font
        row += 1

        # Storage row
        stg_cost = stg_gb * stg_rate
        for col, val in enumerate([SKU_BLOCK_STG, SKU_DESCRIPTIONS[SKU_BLOCK_STG][0], stg_gb, 1, 1, stg_rate, round(stg_cost, 7)], 1):
            ws.cell(row=row, column=col, value=val)
        row += 1

        # VPU row
        vpu_cost = vpu_qty * vpu_rate
        for col, val in enumerate([SKU_BLOCK_VPU, SKU_DESCRIPTIONS[SKU_BLOCK_VPU][0], vpu_qty, 1, 1, vpu_rate, round(vpu_cost, 7)], 1):
            ws.cell(row=row, column=col, value=val)
        row += 1

        # Windows OS row (License Included only)
        if is_windows and not byol:
            ws.cell(row=row, column=2, value="Compute - OS Images").font = group_font
            row += 1
            for col, val in enumerate(['B88318', 'Compute - Windows OS (OCPU Per Hour)', ocpus, 1, HOURS_PER_MONTH, win_os_rate, round(win_os_cost, 7)], 1):
                ws.cell(row=row, column=col, value=val)
            row += 1

        # VM subtotal
        ws.cell(row=row, column=2, value="VM Monthly Total").font = group_font
        ws.cell(row=row, column=7, value=round(vm_total, 7)).font = group_font
        row += 2  # blank line between VMs

    # Grand total
    ws.cell(row=row, column=2, value="GRAND TOTAL — All VMs").font = Font(bold=True)
    ws.cell(row=row, column=7, value=round(grand_total, 2)).font = Font(bold=True)
    row += 2

    # "Quote is for investment proposal only."
    cell = ws.cell(row=row, column=1, value="Quote is for investment proposal only.")
    cell.font = red_font
    row += 2

    # Disclaimer box
    disc_cell = ws.cell(row=row, column=1, value=DISCLAIMER)
    disc_cell.alignment = Alignment(wrap_text=True, vertical="top")
    disc_cell.border = thin_border
    ws.merge_cells(start_row=row, start_column=1, end_row=row+4, end_column=9)
    ws.row_dimensions[row].height = 80

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()

if uploaded_file:
    try:
        df_vHost = pd.read_excel(uploaded_file, sheet_name='vHost')
        uploaded_file.seek(0)
        df_vInfo = pd.read_excel(uploaded_file, sheet_name='vInfo')
    except Exception as e:
        st.error(f"Failed to read file: {e}")
        st.stop()

    os_col = 'OS according to the configuration file'

    # --- vHost calculations ---
    vHost_filtered = df_vHost[df_vHost['CPU Model'] != 'VMware Virtual Processor'].copy()
    total_cores = int(vHost_filtered['# Cores'].sum())
    vHost_filtered['cores_used'] = (vHost_filtered['CPU usage %'] / 100) * vHost_filtered['# Cores']
    cores_based_on_avg_cpu = round(vHost_filtered['cores_used'].sum(), 2)
    avg_cpu_usage = round((cores_based_on_avg_cpu / total_cores) * 100, 2)
    total_memory_gb = round(vHost_filtered['# Memory'].sum() / 1024, 2)
    vHost_filtered['mem_used_gb'] = (vHost_filtered['Memory usage %'] / 100) * (vHost_filtered['# Memory'] / 1024)
    mem_based_on_avg_usage = round(vHost_filtered['mem_used_gb'].sum(), 2)
    avg_mem_usage = round((mem_based_on_avg_usage / total_memory_gb) * 100, 2)

    # --- All powered-on VMs ---
    powered_on_vms = df_vInfo[df_vInfo['Powerstate'] == 'poweredOn'].copy()
    powered_on_vms['Memory GB'] = round(powered_on_vms['Memory'] / 1024, 1)
    powered_on_vms['Provisioned GB'] = round(powered_on_vms['Provisioned MiB'] / 1024, 2)
    powered_on_vms['Provisioned TB'] = round(powered_on_vms['Provisioned MiB'] / 1024 / 1024, 3)
    powered_on_vms['OCPUs'] = (powered_on_vms['CPUs'] / 2).apply(lambda x: max(1, round(x)))
    powered_on_vms['OS Family'] = powered_on_vms[os_col].fillna('Unknown').apply(os_family)
    powered_on_vms['VMware Infra'] = (
        powered_on_vms['VM'].str.lower().str.startswith(VMWARE_NAME_PREFIXES) |
        powered_on_vms['VM'].str.lower().str.contains('vcsa', na=False) |
        powered_on_vms[os_col].str.contains('Photon CRX', case=False, na=False)
    )

    real_vms = powered_on_vms[~powered_on_vms['VMware Infra']].copy()
    vmware_vms = powered_on_vms[powered_on_vms['VMware Infra']].copy()

    tab1, tab2, tab3 = st.tabs(["📋 Current State", "🔍 Workload Analysis", "📦 Migration BOM"])

    # ── TAB 1: Current State ───────────────────────────────
    with tab1:
        st.subheader("Host Infrastructure")
        c1, c2, c3 = st.columns(3)
        c1.metric("Physical Hosts", len(vHost_filtered))
        c2.metric("Total Cores", total_cores)
        c3.metric("Total Memory", f"{total_memory_gb} GB")
        st.divider()

        st.subheader("CPU Utilisation")
        cpu1, cpu2 = st.columns(2)
        cpu1.metric("Avg CPU Usage", f"{avg_cpu_usage}%")
        cpu2.metric("Actual Core Usage (Average)", cores_based_on_avg_cpu)
        st.divider()

        st.subheader("Memory Utilisation")
        mem1, mem2 = st.columns(2)
        mem1.metric("Avg Memory Usage", f"{avg_mem_usage}%")
        mem2.metric("Actual Memory Usage (Average)", f"{mem_based_on_avg_usage} GB")
        st.divider()

        with st.expander("Raw vHost Data"):
            st.dataframe(df_vHost, use_container_width=True)

    # ── TAB 2: Workload Analysis ───────────────────────────
    with tab2:
        st.subheader("All Powered-On VMs")
        w1, w2, w3 = st.columns(3)
        w1.metric("Total VMs", len(powered_on_vms))
        w2.metric("Real Workload VMs", len(real_vms))
        w3.metric("VMware Infra VMs (excluded)", len(vmware_vms))
        st.divider()

        st.subheader("OS Breakdown (Real Workload Only)")
        real_os_counts = (
            real_vms[os_col].fillna('Unknown')
            .value_counts().reset_index()
        )
        real_os_counts.columns = ['OS', 'Count']
        real_os_counts['Family'] = real_os_counts['OS'].apply(os_family)

        col_win, col_lin = st.columns(2)
        with col_win:
            st.markdown("**Windows**")
            win_df = real_os_counts[real_os_counts['Family'] == 'Windows'][['OS', 'Count']].reset_index(drop=True)
            st.dataframe(win_df, use_container_width=True, hide_index=True)
        with col_lin:
            st.markdown("**Linux / Other**")
            other_df = real_os_counts[real_os_counts['Family'] != 'Windows'][['OS', 'Count']].reset_index(drop=True)
            st.dataframe(other_df, use_container_width=True, hide_index=True)
        st.divider()

        st.subheader("VMware Infrastructure VMs (Not Migrating)")
        st.caption("vCLS agents, vRealize agents, Photon CRX, vCenter — VMware platform VMs, not customer workloads.")
        vmware_display = vmware_vms[['VM', os_col, 'CPUs', 'Memory GB']].rename(columns={os_col: 'OS', 'CPUs': 'vCPUs'}).reset_index(drop=True)
        st.dataframe(vmware_display, use_container_width=True, hide_index=True)

    # ── TAB 3: Migration BOM ───────────────────────────────
    with tab3:
        st.subheader("Migration BOM — Real Workload")
        st.caption("Like-for-like migration based on assigned (provisioned) values in VMware. VMware infra VMs excluded. 2 vCPU = 1 OCPU.")

        # Currency + shape selector
        col_cur, col_shape, _ = st.columns([1, 1, 2])
        currency = col_cur.selectbox("Currency", ["NZD", "USD", "AUD", "GBP", "EUR"], index=0)
        shape = col_shape.selectbox("OCI Shape", ["VM.Standard.E4.Flex", "VM.Standard.E5.Flex"], index=0)

        # Fetch prices
        pricing_ok = False
        prices = {}
        with st.spinner("Fetching OCI pricing..."):
            try:
                prices = fetch_oci_prices(currency)
                ocpu_hr  = prices[SKU_E4_OCPU]
                mem_hr   = prices[SKU_E4_MEM]
                stg_mo   = prices[SKU_BLOCK_STG]
                vpu_mo   = prices[SKU_BLOCK_VPU]
                pricing_ok = True
            except Exception as e:
                st.error(f"Could not fetch OCI pricing: {e}")
                pricing_ok = False

        st.divider()

        bom_df = real_vms[['VM', 'OS Family', os_col, 'CPUs', 'OCPUs', 'Memory GB', 'Provisioned GB', 'Provisioned TB']].copy()
        bom_df = bom_df.rename(columns={os_col: 'OS', 'CPUs': 'vCPUs'})
        bom_df = bom_df.sort_values(['OS Family', 'VM']).reset_index(drop=True)

        # BYOL toggle — only editable for Windows VMs
        bom_df['BYOL'] = False
        bom_df['Windows License'] = bom_df['OS Family'].apply(
            lambda f: 'License Included' if f == 'Windows' else 'N/A'
        )

        if pricing_ok:
            win_os_hr = prices.get('B88318', 0)

            st.caption(f"Shape: {shape} | On-Demand | Storage: Balanced (10 VPU) | Prices as of OCI API ({currency})")
            st.info("⚠️ Quote is for investment proposal only.")
            st.caption("For Windows VMs, check **BYOL** if customer brings their own Windows license (removes OS cost).")

            # Editable table — only BYOL column is editable
            edited = st.data_editor(
                bom_df[['VM', 'OS Family', 'OS', 'vCPUs', 'OCPUs', 'Memory GB', 'Provisioned TB', 'BYOL']],
                column_config={
                    "BYOL": st.column_config.CheckboxColumn(
                        "BYOL",
                        help="Check if customer brings their own Windows license",
                        default=False,
                    ),
                    "VM": st.column_config.TextColumn("VM", disabled=True),
                    "OS Family": st.column_config.TextColumn("OS Family", disabled=True),
                    "OS": st.column_config.TextColumn("OS", disabled=True),
                    "vCPUs": st.column_config.NumberColumn("vCPUs", disabled=True),
                    "OCPUs": st.column_config.NumberColumn("OCPUs", disabled=True),
                    "Memory GB": st.column_config.NumberColumn("Memory GB", disabled=True),
                    "Provisioned TB": st.column_config.NumberColumn("Provisioned TB", disabled=True),
                },
                use_container_width=True,
                hide_index=True,
                key="bom_editor"
            )

            # Merge BYOL selections back
            bom_df['BYOL'] = edited['BYOL'].values

            # Recalculate with BYOL
            def vm_cost(row):
                compute = (row['OCPUs'] * ocpu_hr + row['Memory GB'] * mem_hr) * HOURS_PER_MONTH
                storage = row['Provisioned GB'] * stg_mo + BALANCED_VPU * row['Provisioned GB'] * vpu_mo
                win_os  = 0
                if row['OS Family'] == 'Windows' and not row['BYOL']:
                    win_os = row['OCPUs'] * win_os_hr * HOURS_PER_MONTH
                return round(compute, 2), round(storage, 2), round(win_os, 2)

            bom_df[['Compute/mo', 'Storage/mo', 'WinOS/mo']] = bom_df.apply(
                lambda r: pd.Series(vm_cost(r)), axis=1
            )
            bom_df['Total/mo'] = round(bom_df['Compute/mo'] + bom_df['Storage/mo'] + bom_df['WinOS/mo'], 2)

            total_monthly = round(bom_df['Total/mo'].sum(), 2)

            st.divider()

            win = bom_df[bom_df['OS Family'] == 'Windows']
            lin = bom_df[bom_df['OS Family'] != 'Windows']
            byol_label = "excl. Win License" if win['BYOL'].all() else ("incl. Win License" if not win['BYOL'].any() else "mixed BYOL")

            summary_html = f"""
<style>
.bs-table {{ width: 100%; border-collapse: collapse; font-family: sans-serif; }}
.bs-table td {{ padding: 6px 12px 6px 0; vertical-align: top; width: 20%; }}
.bs-label {{ font-size: 0.78rem; color: #888; margin-bottom: 2px; }}
.bs-value {{ font-size: 1.6rem; font-weight: 600; color: #111; }}
.bs-child .bs-value {{ font-size: 1.15rem; font-weight: 500; color: #333; }}
.bs-child .bs-label {{ padding-left: 18px; }}
.bs-child .bs-value {{ padding-left: 18px; }}
.bs-indent {{ color: #aaa; margin-right: 4px; }}
.bs-divider td {{ border-top: 1px solid #e5e5e5; padding-top: 10px; }}
</style>
<table class="bs-table">
  <tr>
    <td><div class="bs-label">Total VMs to Migrate</div><div class="bs-value">{len(bom_df)}</div></td>
    <td><div class="bs-label">Total OCPUs</div><div class="bs-value">{int(bom_df['OCPUs'].sum())}</div></td>
    <td><div class="bs-label">Total RAM</div><div class="bs-value">{round(bom_df['Memory GB'].sum(), 1)} GB</div></td>
    <td><div class="bs-label">Total Storage</div><div class="bs-value">{round(bom_df['Provisioned TB'].sum(), 2)} TB</div></td>
    <td><div class="bs-label">Est. Monthly Cost ({currency})</div><div class="bs-value">{currency} {total_monthly:,.2f}</div></td>
  </tr>
  <tr class="bs-divider bs-child">
    <td><div class="bs-label"><span class="bs-indent">└</span> Windows VMs ({byol_label})</div><div class="bs-value">{len(win)}</div></td>
    <td><div class="bs-label"><span class="bs-indent">└</span> Windows OCPUs</div><div class="bs-value">{int(win['OCPUs'].sum())}</div></td>
    <td><div class="bs-label"><span class="bs-indent">└</span> Windows RAM</div><div class="bs-value">{round(win['Memory GB'].sum(), 1)} GB</div></td>
    <td><div class="bs-label"><span class="bs-indent">└</span> Windows Storage</div><div class="bs-value">{round(win['Provisioned TB'].sum(), 2)} TB</div></td>
    <td><div class="bs-label"><span class="bs-indent">└</span> Windows Cost ({currency})</div><div class="bs-value">{currency} {round(win['Total/mo'].sum(), 2):,.2f}</div></td>
  </tr>
  <tr class="bs-child">
    <td><div class="bs-label"><span class="bs-indent">└</span> Linux / Other VMs</div><div class="bs-value">{len(lin)}</div></td>
    <td><div class="bs-label"><span class="bs-indent">└</span> Linux OCPUs</div><div class="bs-value">{int(lin['OCPUs'].sum())}</div></td>
    <td><div class="bs-label"><span class="bs-indent">└</span> Linux RAM</div><div class="bs-value">{round(lin['Memory GB'].sum(), 1)} GB</div></td>
    <td><div class="bs-label"><span class="bs-indent">└</span> Linux Storage</div><div class="bs-value">{round(lin['Provisioned TB'].sum(), 2)} TB</div></td>
    <td><div class="bs-label"><span class="bs-indent">└</span> Linux Cost ({currency})</div><div class="bs-value">{currency} {round(lin['Total/mo'].sum(), 2):,.2f}</div></td>
  </tr>
</table>
"""
            st.html(summary_html)

            st.divider()

            # Cost summary table with totals
            cost_cols = ['VM', 'OS Family', 'OCPUs', 'Memory GB', 'Provisioned TB', 'BYOL', 'Compute/mo', 'Storage/mo', 'WinOS/mo', 'Total/mo']
            cost_display = bom_df[cost_cols].copy()
            totals = pd.DataFrame([{
                'VM': 'TOTAL', 'OS Family': '', 'OCPUs': int(bom_df['OCPUs'].sum()),
                'Memory GB': round(bom_df['Memory GB'].sum(), 1),
                'Provisioned TB': round(bom_df['Provisioned TB'].sum(), 2),
                'BYOL': '',
                'Compute/mo': round(bom_df['Compute/mo'].sum(), 2),
                'Storage/mo': round(bom_df['Storage/mo'].sum(), 2),
                'WinOS/mo': round(bom_df['WinOS/mo'].sum(), 2),
                'Total/mo': round(bom_df['Total/mo'].sum(), 2),
            }])
            st.dataframe(pd.concat([cost_display, totals], ignore_index=True), use_container_width=True, hide_index=True)

            st.divider()

            # --- RackWare Migration Cost ---
            st.subheader("RackWare Migration Cost (One-Time)")
            st.caption("RackWare on OCI Marketplace — PAYGO. Billed per OCPU per hour during migration window only.")

            rw1, rw2, rw3 = st.columns(3)
            rw_rate_usd = rw1.number_input("RackWare Rate (USD/OCPU/hr)", value=2.53, step=0.01, format="%.2f")
            rw_days     = rw2.number_input("Migration Duration (days)", value=7, min_value=1, step=1)
            rw_hrs_day  = rw3.number_input("Hours/day RackWare runs", value=8, min_value=1, max_value=24, step=1)
            st.caption("💡 RackWare supports pause/resume and scheduled replication windows — it does not need to run 24/7. Typical migrations use 8 hrs/day (business hours). Adjust based on your agreed migration window with the customer.")

            total_ocpus    = int(bom_df['OCPUs'].sum())
            rw_hours       = rw_days * rw_hrs_day
            rw_cost_usd    = round(total_ocpus * rw_rate_usd * rw_hours, 2)

            st.caption(f"Formula: {total_ocpus} OCPUs × USD {rw_rate_usd}/hr × {rw_hours} hrs ({rw_days} days × {rw_hrs_day} hrs/day)")

            rc1, rc2, rc3, rc4 = st.columns(4)
            rc1.metric("Total OCPUs", total_ocpus)
            rc2.metric("Migration Hours", rw_hours)
            rc3.metric("RackWare Cost (USD)", f"USD {rw_cost_usd:,.2f}")
            rc4.metric("Total BOM incl. RackWare (first month)", f"USD {rw_cost_usd:,.2f} + {currency} {total_monthly:,.2f}")

            st.info("💡 RackWare is priced in USD on OCI Marketplace. Convert to local currency using current exchange rate.")

            st.divider()

            # --- Object Storage (Backups) Calculator ---
            st.subheader("Object Storage — Backups (Optional)")
            st.caption("OCI Standard Object Storage for VM backups / snapshots post-migration. Customer estimate — adjust to actual backup retention needs.")

            obj_rate = prices.get(SKU_OBJ_STG, 0.0436101)
            os_col1, os_col2, os_col3 = st.columns(3)
            obj_gb = os_col1.number_input(
                "Estimated Backup Storage (GB)",
                value=1024,
                min_value=0,
                step=128,
                help="First 20 GB free. Enter total backup storage required across all VMs."
            )
            free_gb = 20
            billable_gb = max(0, obj_gb - free_gb)
            obj_cost_mo = round(billable_gb * obj_rate, 2)

            os_col2.metric("Billable GB (after 20 GB free)", f"{billable_gb:,} GB")
            os_col3.metric(f"Est. Monthly Cost ({currency})", f"{currency} {obj_cost_mo:,.2f}")

            st.caption(f"Formula: ({obj_gb:,} GB − {free_gb} GB free) × {currency} {obj_rate}/GB/mo = {currency} {obj_cost_mo:,.2f}/month  |  SKU: {SKU_OBJ_STG}")

            total_with_obj = round(total_monthly + obj_cost_mo, 2)
            if obj_cost_mo > 0:
                st.info(f"💡 Including object storage backups: **{currency} {total_with_obj:,.2f}/month** (compute + block storage + backups)")

        else:
            b1, b2, b3, b4 = st.columns(4)
            b1.metric("VMs to Migrate", len(bom_df))
            b2.metric("Total OCPUs", int(bom_df['OCPUs'].sum()))
            b3.metric("Total RAM", f"{round(bom_df['Memory GB'].sum(), 1)} GB")
            b4.metric("Total Storage", f"{round(bom_df['Provisioned TB'].sum(), 2)} TB")

            totals = pd.DataFrame([{
                'VM': 'TOTAL', 'OS Family': '', 'OS': '',
                'vCPUs': int(bom_df['vCPUs'].sum()),
                'OCPUs': int(bom_df['OCPUs'].sum()),
                'Memory GB': round(bom_df['Memory GB'].sum(), 1),
                'Provisioned GB': round(bom_df['Provisioned GB'].sum(), 1),
                'Provisioned TB': round(bom_df['Provisioned TB'].sum(), 2),
            }])
            bom_display = pd.concat([bom_df, totals], ignore_index=True)
            st.dataframe(bom_display, use_container_width=True, hide_index=True)

        try:
            if pricing_ok and prices:
                excel_data = to_oci_excel(bom_df, prices, currency, shape)
            else:
                excel_data = to_excel(bom_df)
        except Exception:
            excel_data = to_excel(bom_df)
        st.download_button(
            label="⬇️ Export BOM to Excel",
            data=excel_data,
            file_name="migration_bom.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

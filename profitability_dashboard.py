import streamlit as st
import pandas as pd
import plotly.graph_objects as go

def highlight_key_rows(df):
    def row_style(row):
        if 'Particulars' in row.index:
            val = str(row['Particulars']).strip().lower()
            if val == 'sales':
                return ['font-weight: bold; color: #174ea6; background-color: #ffe066' for _ in row]
            elif val == 'gross profit':
                return ['font-weight: bold; color: #0b8043; background-color: #b7e4c7' for _ in row]
            elif val == 'net profit':
                return ['font-weight: bold; color: #b31412; background-color: #f4978e' for _ in row]
        return ['' for _ in row]
    styler = df.style.apply(row_style, axis=1).set_table_styles(
        [
            {'selector': 'th', 'props': [('font-size', '16px'), ('text-align', 'center')]},
            {'selector': 'td', 'props': [('padding', '5px')]},
            {'selector': 'table', 'props': [('border-collapse', 'collapse'), ('border', '1px solid #ccc'), ('border-radius', '5px')]},
            {'selector': 'tr:hover', 'props': [('background-color', '#f0f0f0')]},
            {'selector': 'tr:nth-child(even)', 'props': [('background-color', '#f9f9f9')]}
        ]
    )
    return styler.to_html()

st.set_page_config(page_title="Company Profitability Comparison", layout="wide")
st.title("Comparative Profitability Dashboard")

EXCEL_FILE = "Profitability_CEOITBOX.xlsx"

st.write("Starting dashboard...")

# Only load necessary columns for performance
try:
    df_sales = pd.read_excel(EXCEL_FILE, sheet_name="Sales")

except Exception as e:
    st.error(f"Error loading Sales sheet: {e}")
    st.stop()

# Clean column names (strip spaces)
df_sales.columns = df_sales.columns.str.strip()

# Rename month column
if "Month" in df_sales.columns:
    df_sales = df_sales.rename(columns={"Month": "Month_Year"})

# Format months to 'Apr-25' style
import pandas as pd
def format_month(month):
    try:
        return pd.to_datetime(str(month), format="%d-%m-%Y").strftime("%b-%y")
    except Exception:
        try:
            return pd.to_datetime(str(month)).strftime("%b-%y")
        except Exception:
            return str(month)

def format_indian_number(x):
    try:
        x = float(x)
        if pd.isna(x):
            return ""
        x = int(x)
        s = str(x)
        if len(s) <= 3:
            return s
        else:
            last3 = s[-3:]
            rest = s[:-3]
            rest_pairs = ''
            while len(rest) > 2:
                rest_pairs = ',' + rest[-2:] + rest_pairs
                rest = rest[:-2]
            if rest:
                return rest + rest_pairs + ',' + last3
            else:
                return rest_pairs[1:] + ',' + last3 if rest_pairs else last3
    except Exception:
        return x

if "Month_Year" in df_sales.columns:
    df_sales["Month_Year"] = df_sales["Month_Year"].apply(format_month)





# Sidebar: Month selector
def format_month(month):
    try:
        # Parse as DD-MM-YYYY, then format as 'Apr-25'
        return pd.to_datetime(str(month), format="%d-%m-%Y").strftime("%b-%y")
    except Exception:
        try:
            # Fallback: try parsing as date, then format as 'Apr-25'
            return pd.to_datetime(str(month)).strftime("%b-%y")
        except Exception:
            return str(month)

def parse_num(x):
    try:
        return int(round(float(str(x).replace(',', ''))))
    except Exception:
        return 0

def insert_total_column(df):
    cols = list(df.columns)
    if 'Other Services' in cols:
        idx = cols.index('Other Services')
        if 'Total' in cols:
            return df  # Prevent duplicate Total column
        domain_cols = [col for col in cols if col not in ['Particulars']]
        # Exclude Total if already present
        domain_cols = [col for col in domain_cols if col != 'Total']
        total_vals = df[domain_cols].apply(lambda row: sum([parse_num(x) for x in row]), axis=1)
        df.insert(idx+1, 'Total', total_vals)
    return df


def parse_num(x):
    try:
        return int(round(float(str(x).replace(',', ''))))
    except Exception:
        return 0

# Only one month selector: 'All', 'April', ..., 'March'
month_order = ["April", "May", "June", "July", "August", "September", "October", "November", "December", "January", "February", "March"]
month_abbr_map = {m: pd.to_datetime(m, format="%B").strftime("%b") for m in month_order}
month_options = ["All"] + month_order
selected_month_full = st.sidebar.selectbox("Select Month", month_options, key="month_selectbox")

if selected_month_full == "All":
    # Apr-25 to Mar-26 for FY 2025-26
    months_25_26 = [f"{abbr}-25" for abbr in list(month_abbr_map.values())[:9]] + [f"{abbr}-26" for abbr in list(month_abbr_map.values())[9:]]
    fy_25_26 = df_sales[(df_sales.get("FY", "2025-26") == "2025-26") & (df_sales["Month_Year"].isin(months_25_26))] if "FY" in df_sales.columns else df_sales[df_sales["Month_Year"].isin(months_25_26)]
    drop_cols_25_26 = [col for col in ["Month_Year", "FY"] if col in fy_25_26.columns]
    fy_25_26_tbl = fy_25_26.drop(columns=drop_cols_25_26, errors='ignore')
    fy_25_26_tbl = pd.DataFrame([fy_25_26_tbl.sum(numeric_only=True)])
    fy_25_26_tbl.insert(0, "Particulars", ["Sales"])

    # --- Deferred Revenue FY 2025-26 ---
    df_def_25_26 = pd.read_excel(EXCEL_FILE, sheet_name="Deferred Revenue 25-26")
    df_def_25_26["Month_Year"] = df_def_25_26["Month"].apply(format_month)
    months_25_26 = [f"{abbr}-25" for abbr in list(month_abbr_map.values())[:9]] + [f"{abbr}-26" for abbr in list(month_abbr_map.values())[9:]]
    def_rev_25_26 = df_def_25_26[df_def_25_26["Month_Year"].isin(months_25_26)]
    defrev_val_25_26 = def_rev_25_26["Def. Rev."].sum() if "Def. Rev." in def_rev_25_26.columns else 0
    deferred_row_25_26 = ["Deferred Revenue"] + [defrev_val_25_26 if col=="G-Suite Business" else 0 for col in fy_25_26_tbl.columns[1:]]
    deferred_row_25_26_df = pd.DataFrame([deferred_row_25_26], columns=fy_25_26_tbl.columns)
    fy_25_26_tbl = pd.concat([fy_25_26_tbl, deferred_row_25_26_df], ignore_index=True)

    # --- Purchase FY 2025-26 ---
    df_pur_25_26 = pd.read_excel(EXCEL_FILE, sheet_name="Purchases 25-26", header=None)
    # Row 7 (index 7) and 8 (index 8) are the domain rows, columns 2-13 are Apr-25 to Mar-26
    domain_row_map_25_26 = dict(zip(df_pur_25_26.iloc[7:9,1], df_pur_25_26.iloc[7:9,2:14].values))
    purchase_row_25_26 = ["Purchase"]
    for col in fy_25_26_tbl.columns[1:]:
        if col in domain_row_map_25_26:
            purchase_row_25_26.append(domain_row_map_25_26[col].sum())
        else:
            purchase_row_25_26.append(0)
    purchase_row_25_26_df = pd.DataFrame([purchase_row_25_26], columns=fy_25_26_tbl.columns)
    fy_25_26_tbl = pd.concat([fy_25_26_tbl, purchase_row_25_26_df], ignore_index=True)

    # --- Salary & Incentives FY 2025-26 ---
    df_salary_25_26 = pd.read_excel(EXCEL_FILE, sheet_name="Monthly Salary 25-26", header=1)
    df_salary_25_26.columns = df_salary_25_26.columns.map(lambda x: x.strip() if isinstance(x, str) else x)
    salary_month_cols_25_26 = df_salary_25_26.columns[3:15]  # D to O: 12 months
    # Use actual column names from the DataFrame (after stripping)
    dashboard_domains = ["Training Business", "Tech Assist Recruitment", "Consulting Services & Project work", "WhatsApp API Business", "G-Suite Business"]
    domain_cols = []
    for dom in dashboard_domains:
        found = None
        for col in df_salary_25_26.columns:
            if isinstance(col, str) and col.strip() == dom:
                found = col
                break
        if found:
            domain_cols.append(found)
    salary_row_25_26 = ["Salary & Incentives"]
    if selected_month_full == "All":
        # Salary columns: 3:15, Allocation columns: 18:23
        salary_month_cols = df_salary_25_26.columns[3:15]
        # Set domain order as per user sheet, exclude 'Consulting Services & Project work'
        domain_order = ['Training Business', 'Tech Assist Recruitment', 'WhatsApp API Business', 'G-Suite Business', 'Other Services']
        allocation_cols = [col for col in df_salary_25_26.columns if isinstance(col, str) and col.strip() in domain_order]
        allocation_cols = sorted(allocation_cols, key=lambda x: domain_order.index(x.strip()))
        for dom in allocation_cols:
            total = 0.0
            for m in salary_month_cols:
                month_sum = (df_salary_25_26[m].fillna(0) * df_salary_25_26[dom].fillna(0)).sum()
                total += month_sum
            salary_row_25_26.append(total)
    else:
        # Salary columns: 3:15, Allocation columns: 18:23
        salary_month_cols = df_salary_25_26.columns[3:15]
        # Set domain order as per user sheet, exclude 'Consulting Services & Project work'
        domain_order = ['Training Business', 'Tech Assist Recruitment', 'WhatsApp API Business', 'G-Suite Business', 'Other Services']
        allocation_cols = [col for col in df_salary_25_26.columns if isinstance(col, str) and col.strip() in domain_order]
        allocation_cols = sorted(allocation_cols, key=lambda x: domain_order.index(x.strip()))
        # Find the correct month_col
        month_map = {v: k for k, v in month_abbr_map.items()}
        month_col = None
        for col in salary_month_cols:
            if month_map.get(selected_month_full) in str(col):
                month_col = col
                break
        for dom in allocation_cols:
            if month_col:
                total = (df_salary_25_26[month_col].fillna(0) * df_salary_25_26[dom].fillna(0)).sum()
            else:
                total = 0.0
            salary_row_25_26.append(total)

    # --- Gross Profit FY 2025-26 ---
    def parse_num(x):
        try:
            return int(round(float(str(x).replace(',', ''))))
        except Exception:
            return 0
    sales_vals_25_26 = fy_25_26_tbl.iloc[0, 1:].apply(parse_num)
    defrev_vals_25_26 = fy_25_26_tbl.iloc[1, 1:].apply(parse_num)
    purchase_vals_25_26 = fy_25_26_tbl.iloc[2, 1:].apply(parse_num)
    gross_profit_25_26 = sales_vals_25_26 - defrev_vals_25_26 - purchase_vals_25_26
    gross_profit_row_25_26 = ["Gross Profit"] + gross_profit_25_26.tolist()
    gross_profit_row_25_26_df = pd.DataFrame([gross_profit_row_25_26], columns=fy_25_26_tbl.columns)
    fy_25_26_tbl = pd.concat([fy_25_26_tbl, gross_profit_row_25_26_df], ignore_index=True)
    # Insert Salary & Incentives row after Gross Profit
    salary_row_25_26_df = pd.DataFrame([salary_row_25_26], columns=fy_25_26_tbl.columns)
    fy_25_26_tbl = pd.concat([fy_25_26_tbl, salary_row_25_26_df], ignore_index=True)

    # --- Expenses FY 2025-26 (All months) ---
    df_exp_25_26 = pd.read_excel(EXCEL_FILE, sheet_name="Expenses 25-26", header=0)
    month_cols = df_exp_25_26.columns[2:14]  # C to N
    unique_expenses_25_26 = df_exp_25_26["Expenses"].dropna().unique()
    sales_vals = fy_25_26_tbl.iloc[0, 1:]
    total_sales = sales_vals.sum()
    sales_ratios = sales_vals / total_sales if total_sales != 0 else [0]*len(sales_vals)
    for expense in unique_expenses_25_26:
        total_expense = df_exp_25_26.loc[df_exp_25_26["Expenses"] == expense, month_cols].sum().sum()
        row = [expense]
        for ratio in sales_ratios:
            row.append(total_expense * ratio)
        expense_row_df = pd.DataFrame([row], columns=fy_25_26_tbl.columns)
        fy_25_26_tbl = pd.concat([fy_25_26_tbl, expense_row_df], ignore_index=True)

    # --- TNS Expenses FY 2025-26 (All months, domain allocation) ---
    try:
        df_tns_25_26 = pd.read_excel(EXCEL_FILE, sheet_name='Expense - TNS 25-26')
        df_tns_25_26.columns = df_tns_25_26.columns.str.strip()
        df_tns_25_26 = df_tns_25_26.loc[:, ~df_tns_25_26.columns.duplicated()]
        domain_cols = df_tns_25_26.columns[8:13]  # I to M
        fy_domains = fy_25_26_tbl.columns[1:]
        tns_domain_sums = dict.fromkeys(fy_domains, 0.0)
        for idx, row in df_tns_25_26.iterrows():
            try:
                amt = float(row['Amount']) if pd.notnull(row['Amount']) else 0.0
            except Exception:
                amt = 0.0
            for dcol, domain in zip(domain_cols, fy_domains):
                try:
                    percent = float(row[dcol]) if pd.notnull(row[dcol]) else 0.0
                except Exception:
                    percent = 0.0
                tns_domain_sums[domain] += amt * percent
                
        tns_row_25_26 = ['TNS Expenses'] + [tns_domain_sums[domain] for domain in fy_domains]
        tns_row_25_26_df = pd.DataFrame([tns_row_25_26], columns=fy_25_26_tbl.columns)
        fy_25_26_tbl = pd.concat([fy_25_26_tbl, tns_row_25_26_df], ignore_index=True)
    except Exception as e:
        st.warning(f"Could not load TNS Expenses 25-26: {e}")
    # --- Net Profit FY 2025-26 (All months) ---
    gross_profit_idx = fy_25_26_tbl[fy_25_26_tbl['Particulars'] == 'Gross Profit'].index[0]
    expense_start_idx = fy_25_26_tbl[fy_25_26_tbl['Particulars'] == 'Salary & Incentives'].index[0]
    expense_end_idx = len(fy_25_26_tbl)
    total_expenses = fy_25_26_tbl.iloc[expense_start_idx:expense_end_idx, 1:].applymap(lambda x: float(str(x).replace(',','')) if pd.notnull(x) and x != '' else 0).sum()
    gross_profit = fy_25_26_tbl.iloc[gross_profit_idx, 1:].apply(lambda x: float(str(x).replace(',','')) if pd.notnull(x) and x != '' else 0)
    net_profit = gross_profit - total_expenses
    net_profit_row = ['Net Profit'] + net_profit.tolist()
    net_profit_row_df = pd.DataFrame([net_profit_row], columns=fy_25_26_tbl.columns)
    fy_25_26_tbl = pd.concat([fy_25_26_tbl, net_profit_row_df], ignore_index=True)
    # --- Insert Total column after 'Other Services' ---
    fy_25_26_tbl = insert_total_column(fy_25_26_tbl)
    # --- Net Profit % row (after Total column is present) ---
    try:
        # Remove any existing Net Profit % row before appending
        fy_25_26_tbl = fy_25_26_tbl[fy_25_26_tbl['Particulars'] != 'Net Profit %']
        sales_row = fy_25_26_tbl[fy_25_26_tbl['Particulars'] == 'Sales'].iloc[0, 1:]
        net_profit_row_vals = fy_25_26_tbl[fy_25_26_tbl['Particulars'] == 'Net Profit'].iloc[0, 1:]
        net_profit_pct = []
        for np, s in zip(net_profit_row_vals, sales_row):
            try:
                pct = float(str(np).replace(',','')) * 100 / float(str(s).replace(',','')) if float(str(s).replace(',','')) != 0 else 0
                net_profit_pct.append(f"{pct:.2f}%")
            except:
                net_profit_pct.append('0.00%')
        net_profit_pct_row = ['Net Profit %'] + net_profit_pct
        while len(net_profit_pct_row) < len(fy_25_26_tbl.columns):
            net_profit_pct_row.append('')
        if len(net_profit_pct_row) > len(fy_25_26_tbl.columns):
            net_profit_pct_row = net_profit_pct_row[:len(fy_25_26_tbl.columns)]
        fy_25_26_tbl = pd.concat([fy_25_26_tbl, pd.DataFrame([net_profit_pct_row], columns=fy_25_26_tbl.columns)], ignore_index=True)
    except Exception as e:
        st.warning(f"Could not calculate Net Profit %: {e}")
    # Remove nested insert_total_column definition and use only the module-level function
    # Always check for 'Total' before inserting in the module-level insert_total_column
    fy_25_26_tbl = insert_total_column(fy_25_26_tbl)
    for col in fy_25_26_tbl.columns[1:]:
        fy_25_26_tbl[col] = fy_25_26_tbl[col].apply(lambda x: format_indian_number(x) if pd.notnull(x) else x)
    st.subheader("FY 2025-26: Domain-wise Sales (Apr-25 to Mar-26)")
    st.markdown(highlight_key_rows(fy_25_26_tbl), unsafe_allow_html=True)

    # --- Grouped Bar Chart: Sales, Gross Profit, Net Profit by Domain for FY 2025-26 ---
    import plotly.graph_objects as go
    domain_cols = [col for col in fy_25_26_tbl.columns if col not in ['Particulars', 'Total']]
    def safe_parse(row_name, tbl):
        row = tbl[tbl['Particulars'] == row_name]
        if not row.empty:
            return row.iloc[0][domain_cols].apply(lambda x: float(str(x).replace(',', '')) if pd.notnull(x) and str(x).replace(',', '').replace('.', '').lstrip('-').isdigit() else 0).values
        return [0]*len(domain_cols)
    sales_vals = safe_parse('Sales', fy_25_26_tbl)
    gross_profit_vals = safe_parse('Gross Profit', fy_25_26_tbl)
    net_profit_vals = safe_parse('Net Profit', fy_25_26_tbl)
    def lakhs_labels(values):
        return [f"{v/1e5:.2f}L" if v != 0 else "" for v in values]
    fig = go.Figure(data=[
        go.Bar(name='Sales', x=domain_cols, y=sales_vals, marker_color='#174ea6', text=lakhs_labels(sales_vals), textposition='outside'),
        go.Bar(name='Gross Profit', x=domain_cols, y=gross_profit_vals, marker_color='#0b8043', text=lakhs_labels(gross_profit_vals), textposition='outside'),
        go.Bar(name='Net Profit', x=domain_cols, y=net_profit_vals, marker_color='#b31412', text=lakhs_labels(net_profit_vals), textposition='outside')
    ])
    fig.update_layout(
        barmode='group',
        title='FY 2025-26: Sales, Gross Profit, and Net Profit by Domain',
        xaxis_title='Domain',
        yaxis_title='Amount (INR)',
        legend_title='Metric',
        template='plotly_white',
        height=500
    )
    st.plotly_chart(fig, use_container_width=True)

    # --- Pie Chart: Sales by Domain ---
    fig_sales_pie = px.pie(
        names=domain_cols,
        values=sales_vals,
        title='Sales by Domain',
        color_discrete_sequence=px.colors.qualitative.Set3,
        hole=0.3
    )
    fig_sales_pie.update_traces(textposition='inside', textinfo='percent+label', textfont_size=18)
    fig_sales_pie.update_layout(title_font_size=22)
    st.plotly_chart(fig_sales_pie, use_container_width=True)

    # --- Pie Chart: Net Profit by Domain ---
    fig_netprofit_pie = px.pie(
        names=domain_cols,
        values=net_profit_vals,
        title='Net Profit by Domain',
        color_discrete_sequence=px.colors.qualitative.Set1,
        hole=0.3
    )
    fig_netprofit_pie.update_traces(textposition='inside', textinfo='percent+label', textfont_size=18)
    fig_netprofit_pie.update_layout(title_font_size=22)
    st.plotly_chart(fig_netprofit_pie, use_container_width=True)



    # Apr-24 to Mar-25 for FY 2024-25
    months_24_25 = [f"{abbr}-24" for abbr in list(month_abbr_map.values())[:9]] + [f"{abbr}-25" for abbr in list(month_abbr_map.values())[9:]]
    fy_24_25 = df_sales[(df_sales.get("FY", "2024-25") == "2024-25") & (df_sales["Month_Year"].isin(months_24_25))] if "FY" in df_sales.columns else df_sales[df_sales["Month_Year"].isin(months_24_25)]
    drop_cols_24_25 = [col for col in ["Month_Year", "FY"] if col in fy_24_25.columns]
    fy_24_25_tbl = fy_24_25.drop(columns=drop_cols_24_25, errors='ignore')
    fy_24_25_tbl = pd.DataFrame([fy_24_25_tbl.sum(numeric_only=True)])
    fy_24_25_tbl.insert(0, "Particulars", ["Sales"])

    # --- Deferred Revenue FY 2024-25 ---
    df_def_24_25 = pd.read_excel(EXCEL_FILE, sheet_name="Deferred Revenue 24-25")
    df_def_24_25["Month_Year"] = df_def_24_25["Month"].apply(format_month)
    months_24_25 = [f"{abbr}-24" for abbr in list(month_abbr_map.values())[:9]] + [f"{abbr}-25" for abbr in list(month_abbr_map.values())[9:]]
    def_rev_24_25 = df_def_24_25[df_def_24_25["Month_Year"].isin(months_24_25)]
    defrev_val_24_25 = def_rev_24_25["Def. Rev."].sum() if "Def. Rev." in def_rev_24_25.columns else 0
    deferred_row_24_25 = ["Deferred Revenue"] + [defrev_val_24_25 if col=="G-Suite Business" else 0 for col in fy_24_25_tbl.columns[1:]]
    deferred_row_24_25_df = pd.DataFrame([deferred_row_24_25], columns=fy_24_25_tbl.columns)
    fy_24_25_tbl = pd.concat([fy_24_25_tbl, deferred_row_24_25_df], ignore_index=True)

    # --- Purchase FY 2024-25 ---
    df_pur_24_25 = pd.read_excel(EXCEL_FILE, sheet_name="Purchases 24-25", header=None)
    domain_row_map_24_25 = dict(zip(df_pur_24_25.iloc[7:9,1], df_pur_24_25.iloc[7:9,2:14].values))
    purchase_row_24_25 = ["Purchase"]
    for col in fy_24_25_tbl.columns[1:]:
        if col in domain_row_map_24_25:
            purchase_row_24_25.append(domain_row_map_24_25[col].sum())
        else:
            purchase_row_24_25.append(0)
    purchase_row_24_25_df = pd.DataFrame([purchase_row_24_25], columns=fy_24_25_tbl.columns)
    fy_24_25_tbl = pd.concat([fy_24_25_tbl, purchase_row_24_25_df], ignore_index=True)

    # --- Gross Profit FY 2024-25 ---
    def parse_num(x):
        try:
            return int(round(float(str(x).replace(',', ''))))
        except Exception:
            return 0
    sales_vals_24_25 = fy_24_25_tbl.iloc[0, 1:].apply(parse_num)
    defrev_vals_24_25 = fy_24_25_tbl.iloc[1, 1:].apply(parse_num)
    purchase_vals_24_25 = fy_24_25_tbl.iloc[2, 1:].apply(parse_num)
    gross_profit_24_25 = sales_vals_24_25 - defrev_vals_24_25 - purchase_vals_24_25
    # --- Salary & Incentives FY 2024-25 ---
    df_salary_24_25 = pd.read_excel(EXCEL_FILE, sheet_name="Monthly Salary 24-25", header=1)
    df_salary_24_25.columns = df_salary_24_25.columns.map(lambda x: x.strip() if isinstance(x, str) else x)
    salary_month_cols_24_25 = df_salary_24_25.columns[3:15]  # D to O: 12 months
    # Use actual column names from the DataFrame (after stripping)
    dashboard_domains = ["Training Business", "Tech Assist Recruitment", "Consulting Services & Project work", "WhatsApp API Business", "G-Suite Business"]
    domain_cols = []
    for dom in dashboard_domains:
        found = None
        for col in df_salary_24_25.columns:
            if isinstance(col, str) and col.strip() == dom:
                found = col
                break
        if found:
            domain_cols.append(found)
    salary_row_24_25 = ["Salary & Incentives"]
    if selected_month_full == "All":
        # Salary columns: 3:15, Allocation columns: 18:23
        salary_month_cols = df_salary_24_25.columns[3:15]
        # Set domain order as per user sheet, exclude 'Consulting Services & Project work'
        domain_order = ['Training Business', 'Tech Assist Recruitment', 'WhatsApp API Business', 'G-Suite Business', 'Other Services']
        allocation_cols = [col for col in df_salary_24_25.columns if isinstance(col, str) and col.strip() in domain_order]
        allocation_cols = sorted(allocation_cols, key=lambda x: domain_order.index(x.strip()))
        for dom in allocation_cols:
            total = 0.0
            for m in salary_month_cols:
                month_sum = (df_salary_24_25[m].fillna(0) * df_salary_24_25[dom].fillna(0)).sum()
                total += month_sum
            salary_row_24_25.append(total)
        # Calculate and insert Gross Profit row for 'All' months
        sales_vals = fy_24_25_tbl.iloc[0, 1:].apply(parse_num)
        defrev_vals = fy_24_25_tbl.iloc[1, 1:].apply(parse_num)
        purchase_vals = fy_24_25_tbl.iloc[2, 1:].apply(parse_num)
        gross_profit = sales_vals - defrev_vals - purchase_vals
        gross_profit_row = ['Gross Profit'] + gross_profit.tolist()
        gross_profit_row_df = pd.DataFrame([gross_profit_row], columns=fy_24_25_tbl.columns)
        fy_24_25_tbl = pd.concat([fy_24_25_tbl, gross_profit_row_df], ignore_index=True)
        # Insert Salary & Incentives row after Gross Profit
        salary_row_24_25_df = pd.DataFrame([['Salary & Incentives'] + salary_row_24_25[1:]], columns=fy_24_25_tbl.columns)
        fy_24_25_tbl = pd.concat([fy_24_25_tbl, salary_row_24_25_df], ignore_index=True)
    else:
        # Salary columns: 3:15, Allocation columns: 18:23
        salary_month_cols = df_salary_24_25.columns[3:15]
        # Set domain order as per user sheet, exclude 'Consulting Services & Project work'
        domain_order = ['Training Business', 'Tech Assist Recruitment', 'WhatsApp API Business', 'G-Suite Business', 'Other Services']
        allocation_cols = [col for col in df_salary_24_25.columns if isinstance(col, str) and col.strip() in domain_order]
        allocation_cols = sorted(allocation_cols, key=lambda x: domain_order.index(x.strip()))
        # Find the correct month_col
        month_map = {v: k for k, v in month_abbr_map.items()}
        month_col = None
        for col in salary_month_cols:
            if month_map.get(selected_month_full) in str(col):
                month_col = col
                break
        
        for dom in allocation_cols:
            if month_col:
                total = (df_salary_24_25[month_col].fillna(0) * df_salary_24_25[dom].fillna(0)).sum()
            else:
                total = 0.0
            salary_row_24_25.append(total)
        # Calculate Gross Profit for single month (before Salary row)
        sales_vals = fy_24_25_tbl.iloc[0, 1:].apply(parse_num)
        defrev_vals = fy_24_25_tbl.iloc[1, 1:].apply(parse_num)
        purchase_vals = fy_24_25_tbl.iloc[2, 1:].apply(parse_num)
        gross_profit = sales_vals - defrev_vals - purchase_vals
        gross_profit_row = ['Gross Profit'] + gross_profit.tolist()
        gross_profit_row_df = pd.DataFrame([gross_profit_row], columns=fy_24_25_tbl.columns)
        fy_24_25_tbl = pd.concat([fy_24_25_tbl, gross_profit_row_df], ignore_index=True)
        # Insert Salary & Incentives row after Gross Profit
        salary_row_24_25_df = pd.DataFrame([['Salary & Incentives'] + salary_row_24_25[1:]], columns=fy_24_25_tbl.columns)
        fy_24_25_tbl = pd.concat([fy_24_25_tbl, salary_row_24_25_df], ignore_index=True)
    # --- Expenses FY 2024-25 (All months) ---
    df_exp_24_25 = pd.read_excel(EXCEL_FILE, sheet_name="Expenses 24-25", header=0)
    month_cols = df_exp_24_25.columns[2:14]  # C to N
    unique_expenses_24_25 = df_exp_24_25["Expenses"].dropna().unique()
    sales_vals = fy_24_25_tbl.iloc[0, 1:]
    total_sales = sales_vals.sum()
    sales_ratios = sales_vals / total_sales if total_sales != 0 else [0]*len(sales_vals)
    for expense in unique_expenses_24_25:
        total_expense = df_exp_24_25.loc[df_exp_24_25["Expenses"] == expense, month_cols].sum().sum()
        row = [expense]
        for ratio in sales_ratios:
            row.append(total_expense * ratio)
        expense_row_df = pd.DataFrame([row], columns=fy_24_25_tbl.columns)
        fy_24_25_tbl = pd.concat([fy_24_25_tbl, expense_row_df], ignore_index=True)

    # --- TNS Expenses FY 2024-25 (All months, domain allocation) ---
    try:
        df_tns_24_25 = pd.read_excel(EXCEL_FILE, sheet_name='Expense - TNS 24-25')
        df_tns_24_25.columns = df_tns_24_25.columns.str.strip()
        df_tns_24_25 = df_tns_24_25.loc[:, ~df_tns_24_25.columns.duplicated()]
        domain_cols = df_tns_24_25.columns[8:13]  # I to M
        fy_domains = fy_24_25_tbl.columns[1:]
        tns_domain_sums = dict.fromkeys(fy_domains, 0.0)
        for idx, row in df_tns_24_25.iterrows():
            try:
                amt = float(row['Amount']) if pd.notnull(row['Amount']) else 0.0
            except Exception:
                amt = 0.0
            for dcol, domain in zip(domain_cols, fy_domains):
                try:
                    percent = float(row[dcol]) if pd.notnull(row[dcol]) else 0.0
                except Exception:
                    percent = 0.0
                tns_domain_sums[domain] += amt * percent
                
        tns_row_24_25 = ['TNS Expenses'] + [tns_domain_sums[domain] for domain in fy_domains]
        tns_row_24_25_df = pd.DataFrame([tns_row_24_25], columns=fy_24_25_tbl.columns)
        fy_24_25_tbl = pd.concat([fy_24_25_tbl, tns_row_24_25_df], ignore_index=True)
    except Exception as e:
        st.warning(f"Could not load TNS Expenses 24-25: {e}")
    # --- Net Profit FY 2024-25 (All months) ---
    # Find Gross Profit row index
    gross_profit_idx = fy_24_25_tbl[fy_24_25_tbl['Particulars'] == 'Gross Profit'].index[0]
    # All expense rows are after Salary & Incentives and before Net Profit (to be appended)
    expense_start_idx = fy_24_25_tbl[fy_24_25_tbl['Particulars'] == 'Salary & Incentives'].index[0]
    expense_end_idx = len(fy_24_25_tbl)  # up to end, since Net Profit not yet appended
    # Sum all expense rows (skip 'Particulars' col)
    total_expenses = fy_24_25_tbl.iloc[expense_start_idx:expense_end_idx, 1:].applymap(lambda x: float(str(x).replace(',','')) if pd.notnull(x) and x != '' else 0).sum()
    gross_profit = fy_24_25_tbl.iloc[gross_profit_idx, 1:].apply(lambda x: float(str(x).replace(',','')) if pd.notnull(x) and x != '' else 0)
    net_profit = gross_profit - total_expenses
    net_profit_row = ['Net Profit'] + net_profit.tolist()
    net_profit_row_df = pd.DataFrame([net_profit_row], columns=fy_24_25_tbl.columns)
    fy_24_25_tbl = pd.concat([fy_24_25_tbl, net_profit_row_df], ignore_index=True)
    # Only add Net Profit % row after Total column is present
    fy_24_25_tbl = insert_total_column(fy_24_25_tbl)
    # Remove any existing Net Profit % row before appending
    fy_24_25_tbl = fy_24_25_tbl[fy_24_25_tbl['Particulars'] != 'Net Profit %']
    try:
        sales_row = fy_24_25_tbl[fy_24_25_tbl['Particulars'] == 'Sales'].iloc[0, 1:]
        net_profit_row_vals = fy_24_25_tbl[fy_24_25_tbl['Particulars'] == 'Net Profit'].iloc[0, 1:]
        net_profit_pct = []
        for np, s in zip(net_profit_row_vals, sales_row):
            try:
                pct = float(str(np).replace(',','')) * 100 / float(str(s).replace(',','')) if float(str(s).replace(',','')) != 0 else 0
                net_profit_pct.append(f"{pct:.2f}%")
            except:
                net_profit_pct.append('0.00%')
        # Ensure length matches columns
        net_profit_pct_row = ['Net Profit %'] + net_profit_pct
        while len(net_profit_pct_row) < len(fy_24_25_tbl.columns):
            net_profit_pct_row.append('')
        fy_24_25_tbl = pd.concat([fy_24_25_tbl, pd.DataFrame([net_profit_pct_row], columns=fy_24_25_tbl.columns)], ignore_index=True)
    except Exception as e:
        print("Error calculating Net Profit %:", e)
        # Optionally, you can set a default row or handle the error as needed
    for col in fy_24_25_tbl.columns[1:]:
        fy_24_25_tbl[col] = fy_24_25_tbl[col].apply(lambda x: format_indian_number(x) if pd.notnull(x) and not isinstance(x, str) else x)
    st.subheader("FY 2024-25: Domain-wise Sales (Apr-24 to Mar-25)")
    st.markdown(highlight_key_rows(fy_24_25_tbl), unsafe_allow_html=True)

    # --- Grouped Bar Chart: Sales, Gross Profit, Net Profit by Domain for FY 2024-25 ---
    import plotly.graph_objects as go
    domain_cols = [col for col in fy_24_25_tbl.columns if col not in ['Particulars', 'Total']]
    def safe_parse(row_name, tbl):
        row = tbl[tbl['Particulars'] == row_name]
        if not row.empty:
            return row.iloc[0][domain_cols].apply(lambda x: float(str(x).replace(',', '')) if pd.notnull(x) and str(x).replace(',', '').replace('.', '').lstrip('-').isdigit() else 0).values
        return [0]*len(domain_cols)
    sales_vals = safe_parse('Sales', fy_24_25_tbl)
    gross_profit_vals = safe_parse('Gross Profit', fy_24_25_tbl)
    net_profit_vals = safe_parse('Net Profit', fy_24_25_tbl)
    def lakhs_labels(values):
        return [f"{v/1e5:.2f}L" if v != 0 else "" for v in values]
    fig = go.Figure(data=[
        go.Bar(name='Sales', x=domain_cols, y=sales_vals, marker_color='#174ea6', text=lakhs_labels(sales_vals), textposition='outside'),
        go.Bar(name='Gross Profit', x=domain_cols, y=gross_profit_vals, marker_color='#0b8043', text=lakhs_labels(gross_profit_vals), textposition='outside'),
        go.Bar(name='Net Profit', x=domain_cols, y=net_profit_vals, marker_color='#b31412', text=lakhs_labels(net_profit_vals), textposition='outside')
    ])
    fig.update_layout(
        barmode='group',
        title='FY 2024-25: Sales, Gross Profit, and Net Profit by Domain',
        xaxis_title='Domain',
        yaxis_title='Amount (INR)',
        legend_title='Metric',
        template='plotly_white',
        height=500
    )
    st.plotly_chart(fig, use_container_width=True)

    # --- Pie Chart: Sales by Domain ---
    fig_sales_pie = px.pie(
        names=domain_cols,
        values=sales_vals,
        title='Sales by Domain',
        color_discrete_sequence=px.colors.qualitative.Set3,
        hole=0.3
    )
    fig_sales_pie.update_traces(textposition='inside', textinfo='percent+label', textfont_size=18)
    fig_sales_pie.update_layout(title_font_size=22)
    st.plotly_chart(fig_sales_pie, use_container_width=True)

    # --- Pie Chart: Net Profit by Domain ---
    fig_netprofit_pie = px.pie(
        names=domain_cols,
        values=net_profit_vals,
        title='Net Profit by Domain',
        color_discrete_sequence=px.colors.qualitative.Set1,
        hole=0.3
    )
    fig_netprofit_pie.update_traces(textposition='inside', textinfo='percent+label', textfont_size=18)
    fig_netprofit_pie.update_layout(title_font_size=22)
    st.plotly_chart(fig_netprofit_pie, use_container_width=True)

else:
    abbr = month_abbr_map[selected_month_full]
    fy_25_26_month = f"{abbr}-25"
    fy_24_25_month = f"{abbr}-24"
    fy_25_26 = df_sales[(df_sales.get("FY", "2025-26") == "2025-26") & (df_sales["Month_Year"] == fy_25_26_month)] if "FY" in df_sales.columns else df_sales[df_sales["Month_Year"] == fy_25_26_month]
    drop_cols_25_26 = [col for col in ["Month_Year", "FY"] if col in fy_25_26.columns]
    fy_25_26_tbl = fy_25_26.drop(columns=drop_cols_25_26, errors='ignore')
    fy_25_26_tbl = pd.DataFrame([fy_25_26_tbl.sum(numeric_only=True)])
    fy_25_26_tbl.insert(0, "Particulars", ["Sales"])

    # --- Deferred Revenue FY 2025-26 (single month) ---
    df_def_25_26 = pd.read_excel(EXCEL_FILE, sheet_name="Deferred Revenue 25-26")
    df_def_25_26["Month_Year"] = df_def_25_26["Month"].apply(format_month)
    def_rev_25_26 = df_def_25_26[df_def_25_26["Month_Year"] == fy_25_26_month]
    defrev_val_25_26 = def_rev_25_26["Def. Rev."].sum() if "Def. Rev." in def_rev_25_26.columns else 0
    deferred_row_25_26 = ["Deferred Revenue"] + [defrev_val_25_26 if col=="G-Suite Business" else 0 for col in fy_25_26_tbl.columns[1:]]
    deferred_row_25_26_df = pd.DataFrame([deferred_row_25_26], columns=fy_25_26_tbl.columns)
    fy_25_26_tbl = pd.concat([fy_25_26_tbl, deferred_row_25_26_df], ignore_index=True)

    # --- Purchase FY 2025-26 (single month) ---
    df_pur_25_26 = pd.read_excel(EXCEL_FILE, sheet_name="Purchases 25-26", header=None)
    domain_row_map_25_26 = dict(zip(df_pur_25_26.iloc[7:9,1], df_pur_25_26.iloc[7:9,2:14].values))
    # Find the index for the selected month
    month_idx_25_26 = list(month_abbr_map.values()).index(abbr)
    purchase_row_25_26 = ["Purchase"]
    for col in fy_25_26_tbl.columns[1:]:
        if col in domain_row_map_25_26:
            purchase_row_25_26.append(domain_row_map_25_26[col][month_idx_25_26])
        else:
            purchase_row_25_26.append(0)
    purchase_row_25_26_df = pd.DataFrame([purchase_row_25_26], columns=fy_25_26_tbl.columns)
    fy_25_26_tbl = pd.concat([fy_25_26_tbl, purchase_row_25_26_df], ignore_index=True)
    # --- Gross Profit FY 2025-26 (single month) ---
    sales_vals_25_26 = fy_25_26_tbl.iloc[0, 1:].apply(parse_num)
    defrev_vals_25_26 = fy_25_26_tbl.iloc[1, 1:].apply(parse_num)
    purchase_vals_25_26 = fy_25_26_tbl.iloc[2, 1:].apply(parse_num)
    gross_profit_25_26 = sales_vals_25_26 - defrev_vals_25_26 - purchase_vals_25_26
    gross_profit_row_25_26 = ["Gross Profit"] + gross_profit_25_26.tolist()
    gross_profit_row_25_26_df = pd.DataFrame([gross_profit_row_25_26], columns=fy_25_26_tbl.columns)
    fy_25_26_tbl = pd.concat([fy_25_26_tbl, gross_profit_row_25_26_df], ignore_index=True)

    # --- Salary & Incentives FY 2025-26 (single month) ---
    df_salary_25_26 = pd.read_excel(EXCEL_FILE, sheet_name="Monthly Salary 25-26", header=1)
    df_salary_25_26.columns = df_salary_25_26.columns.map(lambda x: x.strip() if isinstance(x, str) else x)
    salary_month_cols_25_26 = df_salary_25_26.columns[3:15]
    domain_order_25_26 = ['Training Business', 'Tech Assist Recruitment', 'WhatsApp API Business', 'G-Suite Business', 'Other Services']
    allocation_cols_25_26 = [col for col in df_salary_25_26.columns if isinstance(col, str) and col.strip() in domain_order_25_26]
    allocation_cols_25_26 = sorted(allocation_cols_25_26, key=lambda x: domain_order_25_26.index(x.strip()))
    month_map_25_26 = {v: k for k, v in month_abbr_map.items()}
    # Special handling for April: always use 'Apr-25' for FY 2025-26
    import datetime
    # Map month name to month number
    month_num_map = {"April": "04", "May": "05", "June": "06", "July": "07", "August": "08", "September": "09", "October": "10", "November": "11", "December": "12", "January": "01", "February": "02", "March": "03"}
    # For FY 2025-26: April-Dec are 2025, Jan-Mar are 2026
    if selected_month_full in ["April", "May", "June", "July", "August", "September", "October", "November", "December"]:
        year = "2025"
    else:
        year = "2026"
    month_col_25_26_str = f"01-{month_num_map[selected_month_full]}-{year}"
    month_col_25_26 = datetime.datetime.strptime(month_col_25_26_str, "%d-%m-%Y")

    if month_col_25_26 not in salary_month_cols_25_26:
        month_col_25_26 = None
    salary_row_25_26 = ["Salary & Incentives"]
    for dom in allocation_cols_25_26:
        if month_col_25_26:
            total = (df_salary_25_26[month_col_25_26].fillna(0) * df_salary_25_26[dom].fillna(0)).sum()
        else:
            total = 0.0
        salary_row_25_26.append(total)
    salary_row_25_26_df = pd.DataFrame([salary_row_25_26], columns=fy_25_26_tbl.columns)
    fy_25_26_tbl = pd.concat([fy_25_26_tbl, salary_row_25_26_df], ignore_index=True)

    # --- Expenses FY 2025-26 (single month) ---
    df_exp_25_26 = pd.read_excel(EXCEL_FILE, sheet_name="Expenses 25-26", header=0)
    expense_month_col = None
    for col in df_exp_25_26.columns[2:14]:  # C to N
        try:
            if pd.to_datetime(col).month == month_col_25_26.month and pd.to_datetime(col).year == month_col_25_26.year:
                expense_month_col = col
                break
        except Exception:
            continue
    unique_expenses_25_26 = df_exp_25_26["Expenses"].dropna().unique()
    sales_vals = fy_25_26_tbl.iloc[0, 1:]
    total_sales = sales_vals.sum()
    sales_ratios = sales_vals / total_sales if total_sales != 0 else [0]*len(sales_vals)
    for expense in unique_expenses_25_26:
        total_expense = df_exp_25_26.loc[df_exp_25_26["Expenses"] == expense, expense_month_col].sum() if expense_month_col else 0
        row = [expense]
        for ratio in sales_ratios:
            row.append(total_expense * ratio)
        expense_row_df = pd.DataFrame([row], columns=fy_25_26_tbl.columns)
        fy_25_26_tbl = pd.concat([fy_25_26_tbl, expense_row_df], ignore_index=True)

    # --- TNS Expenses FY 2025-26 (single month, domain allocation) ---
    try:
        df_tns_25_26 = pd.read_excel(EXCEL_FILE, sheet_name='Expense - TNS 25-26')
        df_tns_25_26.columns = df_tns_25_26.columns.str.strip()
        df_tns_25_26 = df_tns_25_26.loc[:, ~df_tns_25_26.columns.duplicated()]
        domain_cols = df_tns_25_26.columns[8:13]  # I to M
        fy_domains = fy_25_26_tbl.columns[1:]
        tns_domain_sums = dict.fromkeys(fy_domains, 0.0)
        # Filter by selected month
        if 'Month' in df_tns_25_26.columns:
            month_str = f"01-{month_num_map[selected_month_full]}-2025" if selected_month_full in month_num_map else None
            
            if month_str:
                month_dt = pd.to_datetime(month_str, format="%d-%m-%Y", errors='coerce')
                excel_months = pd.to_datetime(df_tns_25_26['Month'], errors='coerce')
                df_tns_25_26 = df_tns_25_26[excel_months == month_dt]
        for idx, row in df_tns_25_26.iterrows():
            try:
                amt = float(row['Amount']) if pd.notnull(row['Amount']) else 0.0
            except Exception:
                amt = 0.0
            for dcol, domain in zip(domain_cols, fy_domains):
                try:
                    percent = float(row[dcol]) if pd.notnull(row[dcol]) else 0.0
                except Exception:
                    percent = 0.0
                tns_domain_sums[domain] += amt * percent
                
        tns_row_25_26 = ['TNS Expenses'] + [tns_domain_sums[domain] for domain in fy_domains]
        tns_row_25_26_df = pd.DataFrame([tns_row_25_26], columns=fy_25_26_tbl.columns)
        fy_25_26_tbl = pd.concat([fy_25_26_tbl, tns_row_25_26_df], ignore_index=True)
    except Exception as e:
        st.warning(f"Could not load TNS Expenses 25-26: {e}")
    # --- Net Profit FY 2025-26 (single month) ---
    gross_profit_idx = fy_25_26_tbl[fy_25_26_tbl['Particulars'] == 'Gross Profit'].index[0]
    expense_start_idx = fy_25_26_tbl[fy_25_26_tbl['Particulars'] == 'Salary & Incentives'].index[0]
    expense_end_idx = len(fy_25_26_tbl)
    total_expenses = fy_25_26_tbl.iloc[expense_start_idx:expense_end_idx, 1:].applymap(lambda x: float(str(x).replace(',','')) if pd.notnull(x) and x != '' else 0).sum()
    gross_profit = fy_25_26_tbl.iloc[gross_profit_idx, 1:].apply(lambda x: float(str(x).replace(',','')) if pd.notnull(x) and x != '' else 0)
    net_profit = gross_profit - total_expenses
    net_profit_row = ['Net Profit'] + net_profit.tolist()
    net_profit_row_df = pd.DataFrame([net_profit_row], columns=fy_25_26_tbl.columns)
    fy_25_26_tbl = pd.concat([fy_25_26_tbl, net_profit_row_df], ignore_index=True)
    # Only add Net Profit % row after Total column is present
    fy_25_26_tbl = insert_total_column(fy_25_26_tbl)
    # Remove any existing Net Profit % row before appending
    fy_25_26_tbl = fy_25_26_tbl[fy_25_26_tbl['Particulars'] != 'Net Profit %']
    try:
        sales_row = fy_25_26_tbl[fy_25_26_tbl['Particulars'] == 'Sales'].iloc[0, 1:]
        net_profit_row_vals = fy_25_26_tbl[fy_25_26_tbl['Particulars'] == 'Net Profit'].iloc[0, 1:]
        net_profit_pct = []
        for np, s in zip(net_profit_row_vals, sales_row):
            try:
                pct = float(str(np).replace(',','')) * 100 / float(str(s).replace(',','')) if float(str(s).replace(',','')) != 0 else 0
                net_profit_pct.append(f"{pct:.2f}%")
            except:
                net_profit_pct.append('0.00%')
        net_profit_pct_row = ['Net Profit %'] + net_profit_pct
        fy_25_26_tbl = pd.concat([fy_25_26_tbl, pd.DataFrame([net_profit_pct_row], columns=fy_25_26_tbl.columns)], ignore_index=True)
    except Exception as e:
        st.warning(f"Could not calculate Net Profit %: {e}")
    for col in fy_25_26_tbl.columns[1:]:
        fy_25_26_tbl[col] = fy_25_26_tbl[col].apply(lambda x: format_indian_number(x) if pd.notnull(x) and not isinstance(x, str) else x)
    st.subheader(f"FY 2025-26: Domain-wise Sales ({abbr}-25)")
    for col in fy_25_26_tbl.columns[1:]:
        fy_25_26_tbl[col] = fy_25_26_tbl[col].apply(lambda x: format_indian_number(x) if pd.notnull(x) and not isinstance(x, str) else x)
    st.markdown(highlight_key_rows(fy_25_26_tbl), unsafe_allow_html=True)

    # --- Grouped Bar Chart: Sales, Gross Profit, Net Profit by Domain for FY 2025-26 (single month) ---
    import plotly.graph_objects as go
    domain_cols = [col for col in fy_25_26_tbl.columns if col not in ['Particulars', 'Total']]
    def safe_parse(row_name, tbl):
        row = tbl[tbl['Particulars'] == row_name]
        if not row.empty:
            return row.iloc[0][domain_cols].apply(lambda x: float(str(x).replace(',', '')) if pd.notnull(x) and str(x).replace(',', '').replace('.', '').lstrip('-').isdigit() else 0).values
        return [0]*len(domain_cols)
    sales_vals = safe_parse('Sales', fy_25_26_tbl)
    gross_profit_vals = safe_parse('Gross Profit', fy_25_26_tbl)
    net_profit_vals = safe_parse('Net Profit', fy_25_26_tbl)
    def lakhs_labels(values):
        return [f"{v/1e5:.2f}L" if v != 0 else "" for v in values]
    fig = go.Figure(data=[
        go.Bar(name='Sales', x=domain_cols, y=sales_vals, marker_color='#174ea6', text=lakhs_labels(sales_vals), textposition='outside'),
        go.Bar(name='Gross Profit', x=domain_cols, y=gross_profit_vals, marker_color='#0b8043', text=lakhs_labels(gross_profit_vals), textposition='outside'),
        go.Bar(name='Net Profit', x=domain_cols, y=net_profit_vals, marker_color='#b31412', text=lakhs_labels(net_profit_vals), textposition='outside')
    ])
    fig.update_layout(
        barmode='group',
        title=f'{selected_month_full}: Sales, Gross Profit, and Net Profit by Domain',
        xaxis_title='Domain',
        yaxis_title='Amount (INR)',
        legend_title='Metric',
        template='plotly_white',
        height=500
    )
    st.plotly_chart(fig, use_container_width=True)

    # --- Pie Chart: Sales by Domain ---
    fig_sales_pie = px.pie(
        names=domain_cols,
        values=sales_vals,
        title='Sales by Domain',
        color_discrete_sequence=px.colors.qualitative.Set3,
        hole=0.3
    )
    fig_sales_pie.update_traces(textposition='inside', textinfo='percent+label', textfont_size=18)
    fig_sales_pie.update_layout(title_font_size=22)
    st.plotly_chart(fig_sales_pie, use_container_width=True)

    # --- Pie Chart: Net Profit by Domain ---
    fig_netprofit_pie = px.pie(
        names=domain_cols,
        values=net_profit_vals,
        title='Net Profit by Domain',
        color_discrete_sequence=px.colors.qualitative.Set1,
        hole=0.3
    )
    fig_netprofit_pie.update_traces(textposition='inside', textinfo='percent+label', textfont_size=18)
    fig_netprofit_pie.update_layout(title_font_size=22)
    st.plotly_chart(fig_netprofit_pie, use_container_width=True)


    fy_24_25 = df_sales[(df_sales.get("FY", "2024-25") == "2024-25") & (df_sales["Month_Year"] == fy_24_25_month)] if "FY" in df_sales.columns else df_sales[df_sales["Month_Year"] == fy_24_25_month]
    drop_cols_24_25 = [col for col in ["Month_Year", "FY"] if col in fy_24_25.columns]
    fy_24_25_tbl = fy_24_25.drop(columns=drop_cols_24_25, errors='ignore')
    fy_24_25_tbl = pd.DataFrame([fy_24_25_tbl.sum(numeric_only=True)])
    fy_24_25_tbl.insert(0, "Particulars", ["Sales"])

    # --- Deferred Revenue FY 2024-25 (single month) ---
    df_def_24_25 = pd.read_excel(EXCEL_FILE, sheet_name="Deferred Revenue 24-25")
    df_def_24_25["Month_Year"] = df_def_24_25["Month"].apply(format_month)
    def_rev_24_25 = df_def_24_25[df_def_24_25["Month_Year"] == fy_24_25_month]
    defrev_val_24_25 = def_rev_24_25["Def. Rev."].sum() if "Def. Rev." in def_rev_24_25.columns else 0
    deferred_row_24_25 = ["Deferred Revenue"] + [defrev_val_24_25 if col=="G-Suite Business" else 0 for col in fy_24_25_tbl.columns[1:]]
    deferred_row_24_25_df = pd.DataFrame([deferred_row_24_25], columns=fy_24_25_tbl.columns)
    fy_24_25_tbl = pd.concat([fy_24_25_tbl, deferred_row_24_25_df], ignore_index=True)

    # --- Purchase FY 2024-25 (single month) ---
    df_pur_24_25 = pd.read_excel(EXCEL_FILE, sheet_name="Purchases 24-25", header=None)
    domain_row_map_24_25 = dict(zip(df_pur_24_25.iloc[7:9,1], df_pur_24_25.iloc[7:9,2:14].values))
    month_idx_24_25 = list(month_abbr_map.values()).index(abbr)
    purchase_row_24_25 = ["Purchase"]
    for col in fy_24_25_tbl.columns[1:]:
        if col in domain_row_map_24_25:
            purchase_row_24_25.append(domain_row_map_24_25[col][month_idx_24_25])
        else:
            purchase_row_24_25.append(0)
    purchase_row_24_25_df = pd.DataFrame([purchase_row_24_25], columns=fy_24_25_tbl.columns)
    fy_24_25_tbl = pd.concat([fy_24_25_tbl, purchase_row_24_25_df], ignore_index=True)
    # --- Gross Profit FY 2024-25 (single month) ---
    sales_vals_24_25 = fy_24_25_tbl.iloc[0, 1:].apply(parse_num)
    defrev_vals_24_25 = fy_24_25_tbl.iloc[1, 1:].apply(parse_num)
    purchase_vals_24_25 = fy_24_25_tbl.iloc[2, 1:].apply(parse_num)
    gross_profit_24_25 = sales_vals_24_25 - defrev_vals_24_25 - purchase_vals_24_25
    gross_profit_row_24_25 = ["Gross Profit"] + gross_profit_24_25.tolist()
    gross_profit_row_24_25_df = pd.DataFrame([gross_profit_row_24_25], columns=fy_24_25_tbl.columns)
    fy_24_25_tbl = pd.concat([fy_24_25_tbl, gross_profit_row_24_25_df], ignore_index=True)

    # --- Salary & Incentives FY 2024-25 (single month) ---
    df_salary_24_25 = pd.read_excel(EXCEL_FILE, sheet_name="Monthly Salary 24-25", header=1)
    df_salary_24_25.columns = df_salary_24_25.columns.map(lambda x: x.strip() if isinstance(x, str) else x)
    salary_month_cols_24_25 = df_salary_24_25.columns[3:15]
    domain_order_24_25 = ['Training Business', 'Tech Assist Recruitment', 'WhatsApp API Business', 'G-Suite Business', 'Other Services']
    allocation_cols_24_25 = [col for col in df_salary_24_25.columns if isinstance(col, str) and col.strip() in domain_order_24_25]
    allocation_cols_24_25 = sorted(allocation_cols_24_25, key=lambda x: domain_order_24_25.index(x.strip()))
    month_map_24_25 = {v: k for k, v in month_abbr_map.items()}
    # Special handling for April: always use 'Apr-24' for FY 2024-25
    import datetime
    # Map month name to month number
    month_num_map = {"April": "04", "May": "05", "June": "06", "July": "07", "August": "08", "September": "09", "October": "10", "November": "11", "December": "12", "January": "01", "February": "02", "March": "03"}
    # For FY 2024-25: April-Dec are 2024, Jan-Mar are 2025
    if selected_month_full in ["April", "May", "June", "July", "August", "September", "October", "November", "December"]:
        year = "2024"
    else:
        year = "2025"
    month_col_24_25_str = f"01-{month_num_map[selected_month_full]}-{year}"
    month_col_24_25 = datetime.datetime.strptime(month_col_24_25_str, "%d-%m-%Y")

    if month_col_24_25 not in salary_month_cols_24_25:
        month_col_24_25 = None
    salary_row_24_25 = ["Salary & Incentives"]
    for dom in allocation_cols_24_25:
        if month_col_24_25:
            total = (df_salary_24_25[month_col_24_25].fillna(0) * df_salary_24_25[dom].fillna(0)).sum()
        else:
            total = 0.0
        salary_row_24_25.append(total)
    salary_row_24_25_df = pd.DataFrame([salary_row_24_25], columns=fy_24_25_tbl.columns)
    fy_24_25_tbl = pd.concat([fy_24_25_tbl, salary_row_24_25_df], ignore_index=True)

    # Compute month_col_24_25 for expense allocation
    month_num_map = {"April": "04", "May": "05", "June": "06", "July": "07", "August": "08", "September": "09", "October": "10", "November": "11", "December": "12", "January": "01", "February": "02", "March": "03"}
    if selected_month_full in ["April", "May", "June", "July", "August", "September", "October", "November", "December"]:
        year = "2024"
    else:
        year = "2025"
    month_col_24_25_str = f"01-{month_num_map[selected_month_full]}-{year}"
    import datetime
    month_col_24_25 = datetime.datetime.strptime(month_col_24_25_str, "%d-%m-%Y")
    # --- Expenses FY 2024-25 (single month) ---
    df_exp_24_25 = pd.read_excel(EXCEL_FILE, sheet_name="Expenses 24-25", header=0)
    expense_month_col = None
    for col in df_exp_24_25.columns[2:14]:
        try:
            parsed = pd.to_datetime(col, errors='raise')
            if parsed.month == month_col_24_25.month and parsed.year == month_col_24_25.year:
                expense_month_col = col
                break
        except Exception:
            continue
    unique_expenses_24_25 = df_exp_24_25["Expenses"].dropna().unique()
    sales_vals = fy_24_25_tbl.iloc[0, 1:]
    total_sales = sales_vals.sum()
    sales_ratios = sales_vals / total_sales if total_sales != 0 else [0]*len(sales_vals)
    for expense in unique_expenses_24_25:
        total_expense = df_exp_24_25.loc[df_exp_24_25["Expenses"] == expense, expense_month_col].sum() if expense_month_col else 0
        row = [expense]
        for ratio in sales_ratios:
            row.append(total_expense * ratio)
        expense_row_df = pd.DataFrame([row], columns=fy_24_25_tbl.columns)
        fy_24_25_tbl = pd.concat([fy_24_25_tbl, expense_row_df], ignore_index=True)

    # --- TNS Expenses FY 2024-25 (single month, domain allocation) ---
    try:
        df_tns_24_25 = pd.read_excel(EXCEL_FILE, sheet_name='Expense - TNS 24-25')
        df_tns_24_25.columns = df_tns_24_25.columns.str.strip()
        df_tns_24_25 = df_tns_24_25.loc[:, ~df_tns_24_25.columns.duplicated()]
        domain_cols = df_tns_24_25.columns[8:13]  # I to M
        fy_domains = fy_24_25_tbl.columns[1:]
        tns_domain_sums = dict.fromkeys(fy_domains, 0.0)
        # Filter by selected month
        if 'Month' in df_tns_24_25.columns:
            month_str = f"01-{month_num_map[selected_month_full]}-2024" if selected_month_full in month_num_map else None
            
            if month_str:
                month_dt = pd.to_datetime(month_str, format="%d-%m-%Y", errors='coerce')
                excel_months = pd.to_datetime(df_tns_24_25['Month'], errors='coerce')
                filtered = df_tns_24_25[excel_months == month_dt]
                df_tns_24_25 = filtered
        for idx, row in df_tns_24_25.iterrows():
            try:
                amt = float(row['Amount']) if pd.notnull(row['Amount']) else 0.0
            except Exception:
                amt = 0.0
            for dcol, domain in zip(domain_cols, fy_domains):
                try:
                    percent = float(row[dcol]) if pd.notnull(row[dcol]) else 0.0
                except Exception:
                    percent = 0.0
                tns_domain_sums[domain] += amt * percent
                
        tns_row_24_25 = ['TNS Expenses'] + [tns_domain_sums[domain] for domain in fy_domains]
        tns_row_24_25_df = pd.DataFrame([tns_row_24_25], columns=fy_24_25_tbl.columns)
        fy_24_25_tbl = pd.concat([fy_24_25_tbl, tns_row_24_25_df], ignore_index=True)
    except Exception as e:
        st.warning(f"Could not load TNS Expenses 24-25: {e}")
    # --- Net Profit FY 2024-25 (single month) ---
    gross_profit_idx = fy_24_25_tbl[fy_24_25_tbl['Particulars'] == 'Gross Profit'].index[0]
    expense_start_idx = fy_24_25_tbl[fy_24_25_tbl['Particulars'] == 'Salary & Incentives'].index[0]
    expense_end_idx = len(fy_24_25_tbl)
    total_expenses = fy_24_25_tbl.iloc[expense_start_idx:expense_end_idx, 1:].applymap(lambda x: float(str(x).replace(',','')) if pd.notnull(x) and x != '' else 0).sum()
    gross_profit = fy_24_25_tbl.iloc[gross_profit_idx, 1:].apply(lambda x: float(str(x).replace(',','')) if pd.notnull(x) and x != '' else 0)
    net_profit = gross_profit - total_expenses
    net_profit_row = ['Net Profit'] + net_profit.tolist()
    net_profit_row_df = pd.DataFrame([net_profit_row], columns=fy_24_25_tbl.columns)
    fy_24_25_tbl = pd.concat([fy_24_25_tbl, net_profit_row_df], ignore_index=True)
    # Only add Net Profit % row after Total column is present
    fy_24_25_tbl = insert_total_column(fy_24_25_tbl)
    # Remove any existing Net Profit % row before appending
    fy_24_25_tbl = fy_24_25_tbl[fy_24_25_tbl['Particulars'] != 'Net Profit %']
    try:
        sales_row = fy_24_25_tbl[fy_24_25_tbl['Particulars'] == 'Sales'].iloc[0, 1:]
        net_profit_row_vals = fy_24_25_tbl[fy_24_25_tbl['Particulars'] == 'Net Profit'].iloc[0, 1:]
        net_profit_pct = []
        for np, s in zip(net_profit_row_vals, sales_row):
            try:
                pct = float(str(np).replace(',','')) * 100 / float(str(s).replace(',','')) if float(str(s).replace(',','')) != 0 else 0
                net_profit_pct.append(f"{pct:.2f}%")
            except:
                net_profit_pct.append('0.00%')
        net_profit_pct_row = ['Net Profit %'] + net_profit_pct
        fy_24_25_tbl = pd.concat([fy_24_25_tbl, pd.DataFrame([net_profit_pct_row], columns=fy_24_25_tbl.columns)], ignore_index=True)
    except Exception as e:
        st.warning(f"Could not calculate Net Profit %: {e}")
    for col in fy_24_25_tbl.columns[1:]:
        fy_24_25_tbl[col] = fy_24_25_tbl[col].apply(lambda x: format_indian_number(x) if pd.notnull(x) and not isinstance(x, str) else x)
    st.subheader(f"FY 2024-25: Domain-wise Sales ({abbr}-24)")
    st.markdown(highlight_key_rows(fy_24_25_tbl), unsafe_allow_html=True)

    # --- Grouped Bar Chart: Sales, Gross Profit, Net Profit by Domain for FY 2024-25 (single month) ---
    import plotly.graph_objects as go
    domain_cols = [col for col in fy_24_25_tbl.columns if col not in ['Particulars', 'Total']]
    def safe_parse(row_name, tbl):
        row = tbl[tbl['Particulars'] == row_name]
        if not row.empty:
            return row.iloc[0][domain_cols].apply(lambda x: float(str(x).replace(',', '')) if pd.notnull(x) and str(x).replace(',', '').replace('.', '').lstrip('-').isdigit() else 0).values
        return [0]*len(domain_cols)
    sales_vals = safe_parse('Sales', fy_24_25_tbl)
    gross_profit_vals = safe_parse('Gross Profit', fy_24_25_tbl)
    net_profit_vals = safe_parse('Net Profit', fy_24_25_tbl)
    def lakhs_labels(values):
        return [f"{v/1e5:.2f}L" if v != 0 else "" for v in values]
    fig = go.Figure(data=[
        go.Bar(name='Sales', x=domain_cols, y=sales_vals, marker_color='#174ea6', text=lakhs_labels(sales_vals), textposition='outside'),
        go.Bar(name='Gross Profit', x=domain_cols, y=gross_profit_vals, marker_color='#0b8043', text=lakhs_labels(gross_profit_vals), textposition='outside'),
        go.Bar(name='Net Profit', x=domain_cols, y=net_profit_vals, marker_color='#b31412', text=lakhs_labels(net_profit_vals), textposition='outside')
    ])
    fig.update_layout(
        barmode='group',
        title=f'{selected_month_full}: Sales, Gross Profit, and Net Profit by Domain',
        xaxis_title='Domain',
        yaxis_title='Amount (INR)',
        legend_title='Metric',
        template='plotly_white',
        height=500
    )
    st.plotly_chart(fig, use_container_width=True)

    # --- Pie Chart: Sales by Domain ---
    fig_sales_pie = px.pie(
        names=domain_cols,
        values=sales_vals,
        title='Sales by Domain',
        color_discrete_sequence=px.colors.qualitative.Set3,
        hole=0.3
    )
    fig_sales_pie.update_traces(textposition='inside', textinfo='percent+label', textfont_size=18)
    fig_sales_pie.update_layout(title_font_size=22)
    st.plotly_chart(fig_sales_pie, use_container_width=True)

    # --- Pie Chart: Net Profit by Domain ---
    fig_netprofit_pie = px.pie(
        names=domain_cols,
        values=net_profit_vals,
        title='Net Profit by Domain',
        color_discrete_sequence=px.colors.qualitative.Set1,
        hole=0.3
    )
    fig_netprofit_pie.update_traces(textposition='inside', textinfo='percent+label', textfont_size=18)
    fig_netprofit_pie.update_layout(title_font_size=22)
    st.plotly_chart(fig_netprofit_pie, use_container_width=True)


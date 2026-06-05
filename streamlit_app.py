import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import asyncio
import sys
import requests
import io
from playwright.async_api import async_playwright

# Windows asyncio fix
if sys.platform == "win32":
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

def stop_script():
    st.stop()
    if not st.runtime.exists():
        sys.exit(0)

st.set_page_config(
    page_title="MMS Sales Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .stMetric {
        background-color: #f0f2f6;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        text-align: center;
    }
    .stMetric label {
        font-size: 1.1rem !important;
        color: #1e3c72 !important;
    }
    .stMetric [data-testid="stMetricValue"] {
        font-size: 2.5rem !important;
        font-weight: 700 !important;
        color: #0e2b4f !important;
    }
    h1, h2, h3 {
        color: #1e3c72;
    }
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
</style>
""", unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.header("📁 Data Import")
    uploaded_file = st.file_uploader(
        "Upload MMS Sales Tracker Excel file",
        type=["xlsx", "xls"],
        help="Select the Excel file exported from MMS Sales Tracker."
    )

    st.markdown("---")
    st.header("🔗 KoboToolbox Import")
    if st.button("Fetch from KoboToolbox"):
        with st.spinner("Fetching data from KoboToolbox..."):
            try:
                url = "https://kf.kobotoolbox.org/api/v2/assets/ai5wFipp7kxWNA5RJd5jre/export-settings/esdqM5MrXnavHqJcgWRZwcc/data.xlsx"
                response = requests.get(url)
                response.raise_for_status()
                df_raw = pd.read_excel(io.BytesIO(response.content))
                st.session_state.api_data = df_raw
                st.success("Data successfully loaded from KoboToolbox!")
                st.session_state.manual_upload_used = False
            except Exception as e:
                st.error(f"Failed to fetch data from KoboToolbox: {e}")
                st.session_state.api_data = None

    st.markdown("---")
    st.header("📉 Deductions")
    deductions_file = st.file_uploader(
        "Upload Deductions Excel file",
        type=["xlsx", "xls"],
        help="File with columns: Date, Sales Rep Number, Points Deducted, Invoice Number, Amount Deducted"
    )

    st.markdown("---")
    st.markdown("### ℹ️ Instructions")
    st.markdown("""
    1. Upload your **MMS Sales Tracker.xlsx** file **OR** click the **Fetch from KoboToolbox** button.
    2. (Optional) Upload a **Deductions.xlsx** file to apply returns.
    3. Use the filters below to narrow down by date and branch.
    4. The dashboard will automatically update with net points and sales after deductions.
    5. Download the summary as CSV or the full dashboard as interactive HTML.
    6. **Export as PDF** – captures a static PDF of the current dashboard.
    """)
    st.caption("Built with Streamlit & Plotly")

st.title("📊 MMS Sales Dashboard")
st.markdown("---")

# ========== DATA PROCESSING FUNCTIONS ==========
def process_sales_file(uploaded_file):
    r"""Reads the sales Excel, cleans Sales Rep Number (replace / and \ with ;, then preserve as text)."""
    try:
        df = pd.read_excel(uploaded_file)
        df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce')
        df = df.dropna(subset=['Amount', 'Sales Rep Number'])
        df['Date'] = pd.to_datetime(df['Date'])
        df_original = df.copy()

        # Replace / and \ with ; in Sales Rep Number
        df['Sales Rep Number'] = df['Sales Rep Number'].astype(str).str.replace('/', ';', regex=False).str.replace('\\', ';', regex=False)

        # Explode for per‑rep analysis (split amounts equally)
        df_exploded = df.copy()
        df_exploded['rep_count'] = df_exploded['Sales Rep Number'].str.split(';').apply(len)
        df_exploded['Sales Rep Number'] = df_exploded['Sales Rep Number'].str.split(';')
        df_exploded = df_exploded.explode('Sales Rep Number')
        df_exploded['Sales Rep Number'] = df_exploded['Sales Rep Number'].astype(str).str.strip()
        df_exploded = df_exploded[df_exploded['Sales Rep Number'] != '']
        # Keep rep numbers as text so decimal-like values such as 20.3 are accepted.
        df_exploded['Split Amount'] = df_exploded['Amount'] / df_exploded['rep_count']
        return df_original, df_exploded
    except Exception as e:
        st.error(f"Error reading sales file: {e}")
        return None, None

def process_deductions_file(uploaded_file):
    r"""Reads the deductions Excel, cleans Sales Rep Number (replace / and \ with ;, then preserve as text)."""
    try:
        df = pd.read_excel(uploaded_file)
        required_cols = ['Date', 'Sales Rep Number', 'Points Deducted', 'Invoice Number', 'Amount Deducted']
        if not all(col in df.columns for col in required_cols):
            st.error("Deductions file missing required columns.")
            return None, None
        df['Points Deducted'] = pd.to_numeric(df['Points Deducted'], errors='coerce')
        df['Amount Deducted'] = pd.to_numeric(df['Amount Deducted'], errors='coerce')
        df = df.dropna(subset=['Sales Rep Number', 'Points Deducted', 'Invoice Number', 'Amount Deducted'])
        df['Date'] = pd.to_datetime(df['Date'])

        # Replace / and \ with ; in Sales Rep Number
        df['Sales Rep Number'] = df['Sales Rep Number'].astype(str).str.replace('/', ';', regex=False).str.replace('\\', ';', regex=False)

        # Explode for per‑rep deductions
        df_exploded = df.copy()
        df_exploded['rep_count'] = df_exploded['Sales Rep Number'].str.split(';').apply(len)
        df_exploded['Sales Rep Number'] = df_exploded['Sales Rep Number'].str.split(';')
        df_exploded = df_exploded.explode('Sales Rep Number')
        df_exploded['Sales Rep Number'] = df_exploded['Sales Rep Number'].astype(str).str.strip()
        df_exploded = df_exploded[df_exploded['Sales Rep Number'] != '']
        # Keep rep numbers as text so decimal-like values such as 20.3 are accepted.
        # Split points and amount equally among reps
        df_exploded['Points Deducted'] = df_exploded['Points Deducted'] / df_exploded['rep_count']
        df_exploded['Amount Deducted'] = df_exploded['Amount Deducted'] / df_exploded['rep_count']
        return df, df_exploded
    except Exception as e:
        st.error(f"Error reading deductions file: {e}")
        return None, None

# ========== DETERMINE DATA SOURCE ==========
df_sales_orig = None
df_sales_exp = None

if uploaded_file is not None:
    # Manual upload takes precedence
    st.session_state.api_data = None
    st.session_state.manual_upload_used = True
    try:
        df_raw = pd.read_excel(uploaded_file)
        df_sales_orig, df_sales_exp = process_sales_file(uploaded_file)
    except Exception as e:
        st.error(f"Error reading uploaded file: {e}")
        stop_script()
elif 'api_data' in st.session_state and st.session_state.api_data is not None:
    # Use the API data if it exists
    buffer = io.BytesIO()
    st.session_state.api_data.to_excel(buffer, index=False)
    buffer.seek(0)
    df_sales_orig, df_sales_exp = process_sales_file(buffer)
    if df_sales_orig is None:
        stop_script()
else:
    st.info("👈 Please upload an MMS Sales Tracker Excel file or click the Fetch from KoboToolbox button to begin.")
    stop_script()

# ========== DEDUCTIONS PROCESSING ==========
if deductions_file is not None:
    df_ded_orig, df_ded_exp = process_deductions_file(deductions_file)
    if df_ded_orig is None or df_ded_exp is None:
        stop_script()
    has_deductions = True
else:
    has_deductions = False
    df_ded_orig = pd.DataFrame()
    df_ded_exp = pd.DataFrame()

# Check if sales data is valid
if df_sales_orig is None or df_sales_orig.empty or df_sales_orig['Date'].isna().all():
    st.error("No valid sales data with dates found. Please check your uploaded file or API data.")
    stop_script()

# ========== FILTERS ==========
st.sidebar.markdown("### 🔍 Filters")

# Date range filter
min_date = df_sales_orig['Date'].min().date()
max_date = df_sales_orig['Date'].max().date()
date_range = st.sidebar.date_input(
    "Date range",
    value=(min_date, max_date),
    min_value=min_date,
    max_value=max_date
)
if len(date_range) == 2:
    start_date, end_date = date_range
else:
    start_date, end_date = min_date, max_date

# Branch filter
branches = st.sidebar.multiselect(
    "Branch",
    options=sorted(df_sales_orig['Branch'].unique()),
    default=sorted(df_sales_orig['Branch'].unique())
)

# Apply filters to sales
if branches:
    mask_sales = (
        (df_sales_orig['Date'].dt.date >= start_date) &
        (df_sales_orig['Date'].dt.date <= end_date) &
        (df_sales_orig['Branch'].isin(branches))
    )
    df_sales_orig_filtered = df_sales_orig.loc[mask_sales].copy()
    invoices_in_filter = df_sales_orig_filtered['Invoice Number'].unique()
    df_sales_exp_filtered = df_sales_exp[df_sales_exp['Invoice Number'].isin(invoices_in_filter)].copy()
else:
    st.sidebar.warning("Please select at least one branch.")
    df_sales_orig_filtered = pd.DataFrame()
    df_sales_exp_filtered = pd.DataFrame()

# Apply date filter to deductions if present
if has_deductions and not df_ded_orig.empty:
    mask_ded = (
        (df_ded_orig['Date'].dt.date >= start_date) &
        (df_ded_orig['Date'].dt.date <= end_date)
    )
    df_ded_orig_filtered = df_ded_orig.loc[mask_ded].copy()
    invoices_ded_filtered = df_ded_orig_filtered['Invoice Number'].unique()
    df_ded_exp_filtered = df_ded_exp[df_ded_exp['Invoice Number'].isin(invoices_ded_filtered)].copy()
else:
    df_ded_orig_filtered = pd.DataFrame()
    df_ded_exp_filtered = pd.DataFrame()

if df_sales_orig_filtered.empty:
    st.warning("No sales data matches the selected filters. Please adjust your filter criteria.")
    stop_script()

# ========== COMPUTE STATISTICS ==========
sales_rep_stats = df_sales_exp_filtered.groupby('Sales Rep Number').agg(
    invoice_count=('Invoice Number', 'count'),
    total_sales=('Split Amount', 'sum')
).reset_index()

if has_deductions and not df_ded_exp_filtered.empty:
    ded_rep_stats = df_ded_exp_filtered.groupby('Sales Rep Number').agg(
        returned_invoice_count=('Invoice Number', 'count'),
        total_points_deducted=('Points Deducted', 'sum'),
        total_amount_deducted=('Amount Deducted', 'sum')
    ).reset_index()
else:
    ded_rep_stats = pd.DataFrame(columns=['Sales Rep Number', 'returned_invoice_count', 'total_points_deducted', 'total_amount_deducted'])

rep_stats_merged = pd.merge(sales_rep_stats, ded_rep_stats, on='Sales Rep Number', how='left').fillna(0)
rep_stats_merged['net_points'] = rep_stats_merged['invoice_count'] - 2 * rep_stats_merged['returned_invoice_count']
rep_stats_merged['net_sales'] = rep_stats_merged['total_sales'] - rep_stats_merged['total_amount_deducted']

rep_stats_net_sales = rep_stats_merged.sort_values('net_sales', ascending=False)
rep_stats_net_points = rep_stats_merged.sort_values('net_points', ascending=False)

total_sales_net = rep_stats_merged['net_sales'].sum()
total_invoices_orig = len(df_sales_orig_filtered)
total_sales_reps = rep_stats_merged['Sales Rep Number'].nunique()

# ========== METRICS ROW ==========
col1, col2, col3 = st.columns(3)
with col1:
    st.metric(label="💰 Total Sales (net)", value=f"${total_sales_net:,.2f}")
with col2:
    st.metric(label="📄 Total Invoices (original)", value=f"{total_invoices_orig:,}")
with col3:
    st.metric(label="👥 Total Sales Reps", value=f"{total_sales_reps}")

st.markdown("---")

# ========== DAILY SALES CHART ==========
st.subheader("📅 Total Sales by Day")
if not df_sales_orig_filtered.empty:
    daily_sales = df_sales_orig_filtered.groupby(df_sales_orig_filtered['Date'].dt.date)['Amount'].sum().reset_index()
    daily_sales.columns = ['Date', 'Total Sales']
    daily_sales = daily_sales.sort_values('Date')

    max_sales = daily_sales['Total Sales'].max()
    y_max = max_sales * 1.20 if max_sales > 0 else 100

    fig_daily = px.line(
        daily_sales,
        x='Date',
        y='Total Sales',
        markers=True,
        labels={'Total Sales': 'Total Sales ($)', 'Date': 'Date'},
        color_discrete_sequence=['#2ca02c']
    )
    fig_daily.update_traces(
        mode='lines+markers+text',
        text=daily_sales['Total Sales'].round(2),
        texttemplate='$%{text}',
        textposition='top center',
        textfont=dict(size=11, color='black'),
        cliponaxis=False
    )
    fig_daily.update_layout(
        height=450,
        margin=dict(l=20, r=20, t=50, b=50),
        plot_bgcolor='white',
        hovermode='x unified',
        xaxis=dict(tickangle=-45, automargin=True, tickfont=dict(size=10)),
        yaxis=dict(range=[0, y_max], automargin=True, tickfont=dict(size=10))
    )
    st.plotly_chart(fig_daily, width='stretch')
else:
    st.info("No sales data for daily chart.")
st.markdown("---")

# ========== COMBO CHART: Net Points vs Net Sales ==========
st.subheader("📊 Net Points vs Net Sales by Sales Rep")
if not rep_stats_net_sales.empty:
    fig_combo = make_subplots(specs=[[{"secondary_y": True}]])
    fig_combo.add_trace(
        go.Bar(
            x=rep_stats_net_sales['Sales Rep Number'],
            y=rep_stats_net_sales['net_points'],
            name='Net Points',
            marker_color='#3182bd',
            text=rep_stats_net_sales['net_points'],
            textposition='outside',
            textfont=dict(size=10, color='#1e3c72'),
            hovertemplate='Rep: %{x}<br>Net Points: %{y}<extra></extra>'
        ),
        secondary_y=False
    )
    fig_combo.add_trace(
        go.Scatter(
            x=rep_stats_net_sales['Sales Rep Number'],
            y=rep_stats_net_sales['net_sales'],
            name='Net Sales ($)',
            mode='lines+markers+text',
            line=dict(color='#e6550d', width=3),
            marker=dict(size=8, color='#e6550d'),
            text=rep_stats_net_sales['net_sales'].round(2).astype(str) + '$',
            texttemplate='$%{text}',
            textposition='top center',
            textfont=dict(size=10, color='#8b2c0d'),
            cliponaxis=False,
            hovertemplate='Rep: %{x}<br>Net Sales: $%{y:,.2f}<extra></extra>'
        ),
        secondary_y=True
    )
    fig_combo.update_layout(
        height=500,
        margin=dict(l=20, r=80, t=30, b=120),
        plot_bgcolor='white',
        hovermode='x unified',
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
        xaxis=dict(type='category', tickangle=-45, automargin=True, tickfont=dict(size=10))
    )
    fig_combo.update_yaxes(title_text="Net Points", secondary_y=False)
    fig_combo.update_yaxes(title_text="Net Sales ($)", secondary_y=True, title_font=dict(size=12))
    st.plotly_chart(fig_combo, width='stretch')
else:
    st.info("No data for combo chart.")
st.markdown("---")

# ========== NET SALES BAR CHART ==========
st.subheader("💰 Net Sales by Sales Rep")
if not rep_stats_net_sales.empty:
    fig_sales = px.bar(
        rep_stats_net_sales,
        x='Sales Rep Number',
        y='net_sales',
        color_discrete_sequence=['#e6550d'],
        text=rep_stats_net_sales['net_sales'].round(2),
        labels={'net_sales': 'Net Sales ($)', 'Sales Rep Number': 'Sales Rep'}
    )
    fig_sales.update_traces(texttemplate='$%{text}', textposition='outside', textfont=dict(size=11, color='#8b2c0d'))
    fig_sales.update_layout(
        height=450, margin=dict(l=20, r=20, t=30, b=120),
        plot_bgcolor='white', hovermode='x unified',
        xaxis=dict(type='category', tickangle=-45, automargin=True, tickfont=dict(size=10))
    )
    st.plotly_chart(fig_sales, width='stretch')
else:
    st.info("No net sales data for the selected filters.")

# ========== NET POINTS BAR CHART ==========
st.subheader("⭐ Net Points by Sales Rep")
if not rep_stats_net_points.empty:
    fig_points = px.bar(
        rep_stats_net_points,
        x='Sales Rep Number',
        y='net_points',
        color_discrete_sequence=['#3182bd'],
        text='net_points',
        labels={'net_points': 'Net Points', 'Sales Rep Number': 'Sales Rep'}
    )
    fig_points.update_traces(textposition='outside', textfont=dict(size=11, color='#1e3c72'))
    fig_points.update_layout(
        height=450, margin=dict(l=20, r=20, t=30, b=120),
        plot_bgcolor='white', hovermode='x unified',
        xaxis=dict(type='category', tickangle=-45, automargin=True, tickfont=dict(size=10))
    )
    st.plotly_chart(fig_points, width='stretch')
else:
    st.info("No net points data for the selected filters.")
st.markdown("---")

# ========== SUMMARY TABLE ==========
st.subheader("📋 Sales Rep Summary (Sorted by Net Sales)")
if not rep_stats_net_sales.empty:
    display_df = rep_stats_net_sales[['Sales Rep Number', 'net_points', 'net_sales']].copy()
    display_df['net_sales'] = display_df['net_sales'].map('${:,.2f}'.format)
    display_df.columns = ['Sales Rep', 'Net Points', 'Net Sales']
    st.dataframe(
        display_df,
        width='stretch',
        hide_index=True,
        column_config={
            "Sales Rep": st.column_config.NumberColumn("Sales Rep", format="%d"),
            "Net Points": st.column_config.NumberColumn("Net Points", format="%d"),
            "Net Sales": st.column_config.TextColumn("Net Sales"),
        }
    )
else:
    st.info("No rep data for the selected filters.")

# ========== TOP 10 INVOICES ==========
st.subheader("🏆 Top 10 Invoices by Amount (Original)")
if not df_sales_orig_filtered.empty:
    top_invoices = df_sales_orig_filtered.nlargest(10, 'Amount')[['Invoice Number', 'Amount', 'Sales Rep Number', 'Branch']].reset_index(drop=True)
    top_invoices_display = top_invoices.copy()
    top_invoices_display['Amount'] = top_invoices_display['Amount'].map('${:,.2f}'.format)
    top_invoices_display.columns = ['Invoice Number', 'Amount', 'Sales Rep(s)', 'Branch']
    st.dataframe(
        top_invoices_display,
        width='stretch',
        hide_index=True,
        column_config={
            "Invoice Number": st.column_config.TextColumn("Invoice Number"),
            "Amount": st.column_config.TextColumn("Amount"),
            "Sales Rep(s)": st.column_config.TextColumn("Sales Rep(s)"),
            "Branch": st.column_config.TextColumn("Branch"),
        }
    )
else:
    st.info("No invoices for the selected filters.")

# ========== DUPLICATE INVOICES SECTION ==========
st.markdown("---")
st.subheader("⚠️ Duplicate Invoices")
if not df_sales_orig_filtered.empty:
    # Identify invoice numbers that appear more than once
    duplicate_mask = df_sales_orig_filtered.duplicated(subset=['Invoice Number'], keep=False)
    duplicates_df = df_sales_orig_filtered[duplicate_mask].copy()
    if not duplicates_df.empty:
        # Select and reorder columns for display
        dup_display = duplicates_df[['Date', 'Sales Rep Number', 'Invoice Number', 'Amount']].copy()
        dup_display['Date'] = dup_display['Date'].dt.date
        dup_display['Amount'] = dup_display['Amount'].map('${:,.2f}'.format)
        dup_display = dup_display.sort_values(['Invoice Number', 'Date'])
        st.dataframe(
            dup_display,
            width='stretch',
            hide_index=True,
            column_config={
                "Date": st.column_config.DateColumn("Date"),
                "Sales Rep Number": st.column_config.TextColumn("Sales Rep(s)"),
                "Invoice Number": st.column_config.TextColumn("Invoice Number"),
                "Amount": st.column_config.TextColumn("Amount"),
            }
        )
        st.caption(f"Found {len(duplicates_df)} rows belonging to duplicate invoices.")
    else:
        st.success("No duplicate invoices found in the filtered data.")
else:
    st.info("No sales data to check for duplicates.")

# ========== DEDUCTIONS SECTION (if deductions exist) ==========
if has_deductions and not df_ded_exp_filtered.empty:
    st.markdown("---")
    st.subheader("📉 Deductions Overview")
    col_ded1, col_ded2 = st.columns(2)
    with col_ded1:
        st.subheader("Points Deductions by Sales Rep")
        ded_display = ded_rep_stats[['Sales Rep Number', 'total_points_deducted', 'total_amount_deducted']].copy()
        ded_display['total_amount_deducted'] = ded_display['total_amount_deducted'].map('${:,.2f}'.format)
        ded_display.columns = ['Sales Rep', 'Points Deducted', 'Amount Deducted']
        st.dataframe(
            ded_display,
            width='stretch',
            hide_index=True,
            column_config={
                "Sales Rep": st.column_config.NumberColumn("Sales Rep", format="%d"),
                "Points Deducted": st.column_config.NumberColumn("Points Deducted", format="%d"),
                "Amount Deducted": st.column_config.TextColumn("Amount Deducted"),
            }
        )
    with col_ded2:
        st.subheader("Amount Deducted by Invoice")
        invoice_ded = df_ded_orig_filtered.groupby('Invoice Number')['Amount Deducted'].sum().reset_index()
        invoice_ded = invoice_ded.sort_values('Amount Deducted', ascending=False).head(20)
        fig_ded = px.bar(
            invoice_ded,
            x='Invoice Number',
            y='Amount Deducted',
            color_discrete_sequence=['#d62728'],
            text=invoice_ded['Amount Deducted'].round(2),
            labels={'Amount Deducted': 'Amount Deducted ($)', 'Invoice Number': 'Invoice'}
        )
        fig_ded.update_traces(texttemplate='$%{text}', textposition='outside', textfont=dict(size=10))
        fig_ded.update_layout(
            height=400,
            margin=dict(l=20, r=20, t=30, b=80),
            plot_bgcolor='white',
            xaxis=dict(tickangle=-45, automargin=True)
        )
        st.plotly_chart(fig_ded, width='stretch')

st.markdown("---")

# ========== CSV DOWNLOAD ==========
csv = rep_stats_net_sales[['Sales Rep Number', 'net_points', 'net_sales']].to_csv(index=False).encode('utf-8')
st.download_button(label="📥 Download Summary as CSV", data=csv, file_name="sales_rep_summary_net.csv", mime="text/csv")

# ========== HTML EXPORT ==========
def generate_export_html():
    # Duplicate invoices HTML
    if not df_sales_orig_filtered.empty:
        dup_mask = df_sales_orig_filtered.duplicated(subset=['Invoice Number'], keep=False)
        dup_df = df_sales_orig_filtered[dup_mask].copy()
        if not dup_df.empty:
            dup_display = dup_df[['Date', 'Sales Rep Number', 'Invoice Number', 'Amount']].copy()
            dup_display['Date'] = dup_display['Date'].dt.date
            dup_display['Amount'] = dup_display['Amount'].map('${:,.2f}'.format)
            dup_display = dup_display.sort_values(['Invoice Number', 'Date'])
            dup_html = dup_display.to_html(index=False, escape=False, classes='duplicate-table')
        else:
            dup_html = "<p>No duplicate invoices found.</p>"
    else:
        dup_html = "<p>No sales data to check for duplicates.</p>"

    # Combo chart HTML
    if not rep_stats_net_sales.empty:
        combo_html = fig_combo.to_html(include_plotlyjs=False, full_html=False)
    else:
        combo_html = "<p>No combo data</p>"

    # Summary table HTML
    table_html = rep_stats_net_sales[['Sales Rep Number', 'net_points', 'net_sales']].copy()
    table_html['net_sales'] = table_html['net_sales'].map('${:,.2f}'.format)
    table_html.columns = ['Sales Rep', 'Net Points', 'Net Sales']
    table_html = table_html.to_html(index=False, escape=False, classes='summary-table')

    # Top invoices HTML
    top_invoices_html = top_invoices.copy()
    top_invoices_html['Amount'] = top_invoices_html['Amount'].map('${:,.2f}'.format)
    top_invoices_html.columns = ['Invoice Number', 'Amount', 'Sales Rep(s)', 'Branch']
    top_invoices_html = top_invoices_html.to_html(index=False, escape=False, classes='top-invoices-table')

    # Deductions table and chart HTML
    ded_section = ""
    if has_deductions and not df_ded_exp_filtered.empty:
        ded_table_html = ded_rep_stats[['Sales Rep Number', 'total_points_deducted', 'total_amount_deducted']].copy()
        ded_table_html['total_amount_deducted'] = ded_table_html['total_amount_deducted'].map('${:,.2f}'.format)
        ded_table_html.columns = ['Sales Rep', 'Points Deducted', 'Amount Deducted']
        ded_table_html = ded_table_html.to_html(index=False, escape=False, classes='deductions-table')

        invoice_ded = df_ded_orig_filtered.groupby('Invoice Number')['Amount Deducted'].sum().reset_index()
        invoice_ded = invoice_ded.sort_values('Amount Deducted', ascending=False).head(20)
        fig_ded = px.bar(
            invoice_ded,
            x='Invoice Number',
            y='Amount Deducted',
            color_discrete_sequence=['#d62728'],
            text=invoice_ded['Amount Deducted'].round(2),
            labels={'Amount Deducted': 'Amount Deducted ($)', 'Invoice Number': 'Invoice'}
        )
        fig_ded.update_traces(texttemplate='$%{text}', textposition='outside', textfont=dict(size=10))
        fig_ded.update_layout(height=400, margin=dict(l=20, r=20, t=30, b=80), plot_bgcolor='white', xaxis=dict(tickangle=-45, automargin=True))
        ded_chart_html = fig_ded.to_html(include_plotlyjs=False, full_html=False)

        ded_section = f"""
        <div class="table-container">
            <h3>📉 Points Deductions by Sales Rep</h3>
            {ded_table_html}
        </div>
        <div class="chart-container">
            <h3>📉 Amount Deducted by Invoice</h3>
            {ded_chart_html}
        </div>
        """

    # Other charts
    sales_html = fig_sales.to_html(include_plotlyjs=False, full_html=False) if not rep_stats_net_sales.empty else "<p>No data</p>"
    points_html = fig_points.to_html(include_plotlyjs=False, full_html=False) if not rep_stats_net_points.empty else "<p>No data</p>"
    daily_html = fig_daily.to_html(include_plotlyjs=False, full_html=False) if not df_sales_orig_filtered.empty else "<p>No daily data</p>"

    html_template = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>MMS Sales Dashboard Export</title>
        <script src="https://cdn.plot.ly/plotly-2.32.0.min.js"></script>
        <style>
            body {{
                font-family: 'Arial', sans-serif;
                background-color: #f0f2f6;
                margin: 0;
                padding: 20px;
            }}
            .container {{
                max-width: 1200px;
                margin: 0 auto;
            }}
            h1 {{
                color: #1e3c72;
                text-align: center;
                font-weight: 600;
                margin-bottom: 30px;
            }}
            .metrics-container {{
                display: flex;
                justify-content: space-between;
                gap: 20px;
                margin-bottom: 30px;
            }}
            .metric-card {{
                background-color: white;
                border-radius: 10px;
                box-shadow: 0 4px 6px rgba(0,0,0,0.1);
                padding: 20px;
                flex: 1;
                text-align: center;
            }}
            .metric-card.blue {{ background-color: #e3f2fd; }}
            .metric-card.green {{ background-color: #e8f5e8; }}
            .metric-card.orange {{ background-color: #fff3e0; }}
            .metric-label {{
                font-size: 1.2rem;
                color: #1e3c72;
                margin-bottom: 10px;
            }}
            .metric-value {{
                font-size: 2.5rem;
                font-weight: 700;
                color: #0e2b4f;
            }}
            .chart-container {{
                background-color: white;
                border-radius: 10px;
                padding: 15px;
                margin-bottom: 20px;
                box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            }}
            .chart-container h3 {{
                margin-top: 0;
                color: #1e3c72;
                font-weight: 500;
            }}
            .table-container {{
                background-color: white;
                border-radius: 10px;
                padding: 20px;
                margin-bottom: 20px;
                box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            }}
            .summary-table, .top-invoices-table, .deductions-table, .duplicate-table {{
                width: 100%;
                border-collapse: collapse;
            }}
            .summary-table th, .top-invoices-table th, .deductions-table th, .duplicate-table th {{
                background-color: #1e3c72;
                color: white;
                padding: 12px;
                text-align: left;
            }}
            .summary-table td, .top-invoices-table td, .deductions-table td, .duplicate-table td {{
                padding: 10px 12px;
                border-bottom: 1px solid #ddd;
            }}
            .summary-table tr:nth-child(even), .top-invoices-table tr:nth-child(even), .deductions-table tr:nth-child(even), .duplicate-table tr:nth-child(even) {{
                background-color: #f5f7fa;
            }}
            .summary-table tr:hover, .top-invoices-table tr:hover, .deductions-table tr:hover, .duplicate-table tr:hover {{
                background-color: #e9ecef;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>📊 MMS Sales Dashboard</h1>
            <div class="metrics-container">
                <div class="metric-card blue">
                    <div class="metric-label">💰 Total Sales (net)</div>
                    <div class="metric-value">${total_sales_net:,.2f}</div>
                </div>
                <div class="metric-card green">
                    <div class="metric-label">📄 Total Invoices (original)</div>
                    <div class="metric-value">{total_invoices_orig:,}</div>
                </div>
                <div class="metric-card orange">
                    <div class="metric-label">👥 Total Sales Reps</div>
                    <div class="metric-value">{total_sales_reps}</div>
                </div>
            </div>
            <div class="chart-container">
                <h3>📅 Total Sales by Day</h3>
                {daily_html}
            </div>
            <div class="chart-container">
                <h3>📊 Net Points vs Net Sales by Sales Rep</h3>
                {combo_html}
            </div>
            <div class="chart-container">
                <h3>💰 Net Sales by Sales Rep </h3>
                {sales_html}
            </div>
            <div class="chart-container">
                <h3>⭐ Net Points by Sales Rep </h3>
                {points_html}
            </div>
            <div class="table-container">
                <h3>📋 Sales Rep Summary (Sorted by Net Sales)</h3>
                {table_html}
            </div>
            <div class="table-container">
                <h3>⚠️ Duplicate Invoices</h3>
                {dup_html}
            </div>
            <div class="table-container">
                <h3>🏆 Top 10 Invoices by Amount (Original)</h3>
                {top_invoices_html}
            </div>
            {ded_section}
        </div>
    </body>
    </html>
    """
    return html_template

with st.sidebar:
    st.markdown("---")
    st.download_button(
        label="📥 Export Dashboard as HTML",
        data=generate_export_html(),
        file_name="mms_dashboard_export.html",
        mime="text/html",
        help="Download an interactive HTML version of the dashboard (charts remain interactive)."
    )

# ========== PDF EXPORT ==========
async def generate_pdf_from_html(html_content: str) -> bytes:
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()
        await page.set_content(html_content, wait_until="networkidle")
        await page.wait_for_timeout(3000)
        pdf_bytes = await page.pdf(format="A4", print_background=True)
        await browser.close()
        return pdf_bytes

def export_as_pdf():
    html_string = generate_export_html()
    try:
        import threading
        res = []
        err = []
        def worker():
            try:
                loop = asyncio.new_event_loop()
                asyncio.set_event_loop(loop)
                res.append(loop.run_until_complete(generate_pdf_from_html(html_string)))
                loop.close()
            except Exception as e:
                err.append(e)
        t = threading.Thread(target=worker)
        t.start()
        t.join()
        if err:
            raise err[0]
        return res[0]
    except Exception as e:
        st.error(f"PDF generation failed: {e}\n\nMake sure Playwright is installed: `pip install playwright && playwright install chromium`")
        return None

with st.sidebar:
    st.markdown("---")
    if st.button("📄 Export Dashboard as PDF", help="Generate a static PDF of the current dashboard (requires Playwright)"):
        with st.spinner("Generating PDF... (this may take a few seconds)"):
            pdf_bytes = export_as_pdf()
            if pdf_bytes:
                st.download_button(
                    label="📥 Download PDF",
                    data=pdf_bytes,
                    file_name="mms_dashboard.pdf",
                    mime="application/pdf"
                )
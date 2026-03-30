# app.py
# ─────────────────────────────────────────────────────────────────────────────
# SALES & CUSTOMER ANALYSIS DASHBOARD
# Built with Streamlit + Pandas + Plotly
#
# HOW TO RUN:
#   streamlit run app.py
#
# WHAT THIS FILE DOES:
#   1. Lets you upload your Excel sales file
#   2. Cleans the data automatically
#   3. Shows KPI cards, charts, and insights
# ─────────────────────────────────────────────────────────────────────────────

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import warnings
warnings.filterwarnings("ignore")  # Suppress minor warnings

# ── PAGE CONFIGURATION ───────────────────────────────────────────────────────
st.set_page_config(
    page_title="Sales & Customer Analysis",
    page_icon="📊",
    layout="wide",                # Use full screen width
    initial_sidebar_state="expanded"
)

# ── CUSTOM CSS (makes it look cleaner) ───────────────────────────────────────
st.markdown("""
    <style>
        .metric-card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 20px;
            border-radius: 12px;
            color: white;
            text-align: center;
            margin: 5px 0;
        }
        .metric-value {
            font-size: 28px;
            font-weight: bold;
            margin: 5px 0;
        }
        .metric-label {
            font-size: 13px;
            opacity: 0.85;
        }
        .insight-box {
            background-color: #f0f7ff;
            border-left: 4px solid #2196F3;
            padding: 12px 16px;
            border-radius: 4px;
            margin: 8px 0;
            color: #1a1a2e;
        }
        .warning-box {
            background-color: #fff8e1;
            border-left: 4px solid #FF9800;
            padding: 12px 16px;
            border-radius: 4px;
            margin: 8px 0;
            color: #1a1a2e;
        }
        .success-box {
            background-color: #e8f5e9;
            border-left: 4px solid #4CAF50;
            padding: 12px 16px;
            border-radius: 4px;
            margin: 8px 0;
            color: #1a1a2e;
        }
        div[data-testid="stMetricValue"] {
            font-size: 24px;
        }
    </style>
""", unsafe_allow_html=True)


# ═════════════════════════════════════════════════════════════════════════════
# SECTION 1: DATA LOADING & CLEANING
# ═════════════════════════════════════════════════════════════════════════════

@st.cache_data  # Cache so it doesn't reload on every click
def load_and_clean_data(uploaded_file):
    """
    Loads Excel file and cleans it automatically.
    Handles: missing values, wrong date formats, duplicate rows.
    Returns a clean DataFrame ready for analysis.
    """
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as e:
        return None, f"❌ Could not read file: {str(e)}"

    original_rows = len(df)

    # ── Clean column names (remove spaces, fix caps) ─────────────────────
    df.columns = df.columns.str.strip()

    # ── Rename columns to standard internal names ─────────────────────────
    # This makes the rest of the code work even if user has slightly
    # different column names in their real Excel file
    column_map = {
        "Order ID":              "order_id",
        "Order Date":            "order_date",
        "Customer Name":         "customer_name",
        "Customer Phone":        "customer_phone",
        "Product Name":          "product_name",
        "Quantity":              "quantity",
        "Unit Price (BDT)":      "unit_price",
        "Total Revenue (BDT)":   "revenue",
        "Payment Method":        "payment_method",
        "Delivery Status":       "delivery_status",
    }
    df = df.rename(columns=column_map)

    # ── Fix date column ────────────────────────────────────────────────────
    try:
        df["order_date"] = pd.to_datetime(df["order_date"], errors="coerce")
    except Exception:
        return None, "❌ Could not parse 'Order Date' column. Make sure dates are in YYYY-MM-DD format."

    # ── Drop rows where date or revenue is missing (can't analyze without) ─
    df = df.dropna(subset=["order_date", "revenue"])

    # ── Fill other missing values with sensible defaults ──────────────────
    df["customer_name"]    = df["customer_name"].fillna("Unknown Customer")
    df["product_name"]     = df["product_name"].fillna("Unknown Product")
    df["delivery_status"]  = df["delivery_status"].fillna("Unknown")
    df["payment_method"]   = df["payment_method"].fillna("Unknown")
    df["quantity"]         = df["quantity"].fillna(1)

    # ── Make sure numeric columns are actually numbers ─────────────────────
    df["revenue"]    = pd.to_numeric(df["revenue"],    errors="coerce").fillna(0)
    df["quantity"]   = pd.to_numeric(df["quantity"],   errors="coerce").fillna(1)
    df["unit_price"] = pd.to_numeric(df["unit_price"], errors="coerce").fillna(0)

    # ── Remove duplicate order IDs (keep first) ───────────────────────────
    if "order_id" in df.columns:
        df = df.drop_duplicates(subset=["order_id"], keep="first")

    # ── Add helper columns for time analysis ──────────────────────────────
    df["month"]        = df["order_date"].dt.to_period("M").astype(str)
    df["week"]         = df["order_date"].dt.to_period("W").astype(str)
    df["day_of_week"]  = df["order_date"].dt.day_name()
    df["month_name"]   = df["order_date"].dt.strftime("%b %Y")

    # ── Only keep Delivered orders for revenue analysis ────────────────────
    df_delivered = df[df["delivery_status"].str.lower() == "delivered"].copy()

    cleaned_rows = len(df)
    dropped      = original_rows - cleaned_rows

    return df, df_delivered, dropped, None  # Return both full and delivered-only


# ═════════════════════════════════════════════════════════════════════════════
# SECTION 2: HELPER FUNCTIONS
# ═════════════════════════════════════════════════════════════════════════════

def format_bdt(amount):
    """Format a number as BDT currency. E.g. 1234567 → ৳ 12,34,567"""
    return f"৳ {amount:,.0f}"

def show_insight(text, type="info"):
    """Show a colored insight box. type = 'info', 'warning', or 'success'"""
    css_class = {
        "info":    "insight-box",
        "warning": "warning-box",
        "success": "success-box"
    }.get(type, "insight-box")
    st.markdown(f'<div class="{css_class}">💡 {text}</div>', unsafe_allow_html=True)


# ═════════════════════════════════════════════════════════════════════════════
# SECTION 3: DASHBOARD HEADER & FILE UPLOAD
# ═════════════════════════════════════════════════════════════════════════════

st.title("📊 Sales & Customer Analysis Dashboard")
st.markdown("*Upload your sales Excel file to get instant insights and recommendations.*")
st.divider()

# ── Sidebar ───────────────────────────────────────────────────────────────
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/bar-chart.png", width=60)
    st.title("⚙️ Controls")
    st.markdown("---")

    uploaded_file = st.file_uploader(
        "📁 Upload Your Sales Excel File",
        type=["xlsx", "xls"],
        help="Upload your sales data in Excel format (.xlsx)"
    )

    st.markdown("---")
    st.markdown("### 📌 Expected Columns")
    st.markdown("""
    - Order ID
    - Order Date
    - Customer Name
    - Customer Phone
    - Product Name
    - Quantity
    - Unit Price (BDT)
    - Total Revenue (BDT)
    - Payment Method
    - Delivery Status
    """)
    st.markdown("---")
    st.caption("Built by ExpertBuilder 🛠️")


# ═════════════════════════════════════════════════════════════════════════════
# SECTION 4: MAIN DASHBOARD (only shows after file is uploaded)
# ═════════════════════════════════════════════════════════════════════════════

if uploaded_file is None:
    # ── Welcome screen when no file is uploaded yet ───────────────────────
    col1, col2, col3 = st.columns(3)
    with col1:
        st.info("### 📁 Step 1\nUpload your Excel sales file from the left sidebar")
    with col2:
        st.info("### 🔍 Step 2\nThe system auto-cleans and analyzes your data")
    with col3:
        st.info("### 📈 Step 3\nGet charts, insights & actionable recommendations")

    st.markdown("---")
    st.markdown("#### 👈 Upload your file from the sidebar to get started!")

else:
    # ── Load and clean the data ───────────────────────────────────────────
    result = load_and_clean_data(uploaded_file)

    if result[3]:  # If there's an error message
        st.error(result[3])
    else:
        df, df_delivered, dropped_rows, _ = result

        # Show data quality warning if rows were dropped
        if dropped_rows > 0:
            st.warning(f"⚠️ {dropped_rows} rows were removed due to missing dates or revenue. All other data cleaned successfully.")

        st.success(f"✅ Data loaded! **{len(df):,} orders** found | Analyzing **{len(df_delivered):,} delivered orders** for revenue.")

        # ── DATE FILTER in sidebar ────────────────────────────────────────
        with st.sidebar:
            st.markdown("### 📅 Filter by Date")
            min_date = df["order_date"].min().date()
            max_date = df["order_date"].max().date()

            date_range = st.date_input(
                "Select date range",
                value=(min_date, max_date),
                min_value=min_date,
                max_value=max_date
            )

        # Apply date filter
        if len(date_range) == 2:
            start, end = date_range
            df_filtered   = df[(df["order_date"].dt.date >= start) &
                               (df["order_date"].dt.date <= end)]
            df_rev        = df_delivered[(df_delivered["order_date"].dt.date >= start) &
                                         (df_delivered["order_date"].dt.date <= end)]
        else:
            df_filtered = df
            df_rev      = df_delivered

        # ── TABS for different analysis sections ──────────────────────────
        tab1, tab2, tab3 = st.tabs([
            "📈 Sales Overview",
            "🏆 Product Analysis",
            "💳 Payment & Delivery"
        ])


        # ─────────────────────────────────────────────────────────────────
        # TAB 1: SALES OVERVIEW
        # ─────────────────────────────────────────────────────────────────
        with tab1:
            st.subheader("📈 Sales Overview")

            # ── KPI CARDS ─────────────────────────────────────────────────
            total_revenue   = df_rev["revenue"].sum()
            total_orders    = len(df_rev)
            avg_order_value = total_revenue / total_orders if total_orders > 0 else 0
            total_qty       = df_rev["quantity"].sum()
            unique_customers = df_rev["customer_name"].nunique()
            return_rate     = (
                len(df[df["delivery_status"].str.lower() == "returned"]) / len(df) * 100
                if len(df) > 0 else 0
            )

            col1, col2, col3, col4, col5, col6 = st.columns(6)

            with col1:
                st.metric("💰 Total Revenue",    format_bdt(total_revenue))
            with col2:
                st.metric("🛒 Total Orders",      f"{total_orders:,}")
            with col3:
                st.metric("📦 Items Sold",         f"{int(total_qty):,}")
            with col4:
                st.metric("🎯 Avg Order Value",   format_bdt(avg_order_value))
            with col5:
                st.metric("👥 Customers",          f"{unique_customers:,}")
            with col6:
                st.metric("↩️ Return Rate",        f"{return_rate:.1f}%")

            st.markdown("---")

            # ── SALES TREND CHART ─────────────────────────────────────────
            st.subheader("📅 Sales Trend")

            trend_option = st.radio(
                "View by:",
                ["Monthly", "Weekly", "Day of Week"],
                horizontal=True
            )

            if trend_option == "Monthly":
                trend_df = (df_rev.groupby("month_name")
                                  .agg(revenue=("revenue", "sum"),
                                       orders=("order_id", "count"))
                                  .reset_index()
                                  .sort_values("month_name"))

                fig = go.Figure()
                fig.add_bar(
                    x=trend_df["month_name"],
                    y=trend_df["revenue"],
                    name="Revenue (BDT)",
                    marker_color="#667eea",
                    yaxis="y"
                )
                fig.add_scatter(
                    x=trend_df["month_name"],
                    y=trend_df["orders"],
                    name="Number of Orders",
                    mode="lines+markers",
                    marker=dict(color="#FF6B6B", size=8),
                    yaxis="y2"
                )
                fig.update_layout(
                    title="Monthly Revenue & Orders",
                    yaxis=dict(title="Revenue (BDT)"),
                    yaxis2=dict(title="Orders", overlaying="y", side="right"),
                    legend=dict(x=0, y=1.1, orientation="h"),
                    height=400
                )
                st.plotly_chart(fig, use_container_width=True)

                # Auto insight
                if len(trend_df) >= 2:
                    best_month  = trend_df.loc[trend_df["revenue"].idxmax(), "month_name"]
                    worst_month = trend_df.loc[trend_df["revenue"].idxmin(), "month_name"]
                    show_insight(f"Best month: <b>{best_month}</b> | Lowest month: <b>{worst_month}</b>. Consider running promotions in low months.", "info")

            elif trend_option == "Weekly":
                trend_df = (df_rev.groupby("week")
                                  .agg(revenue=("revenue", "sum"))
                                  .reset_index()
                                  .sort_values("week"))
                fig = px.line(
                    trend_df, x="week", y="revenue",
                    title="Weekly Revenue Trend",
                    markers=True,
                    color_discrete_sequence=["#667eea"]
                )
                fig.update_layout(height=400, xaxis_title="Week", yaxis_title="Revenue (BDT)")
                st.plotly_chart(fig, use_container_width=True)

            else:  # Day of Week
                day_order = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
                trend_df  = (df_rev.groupby("day_of_week")
                                   .agg(revenue=("revenue","sum"),
                                        orders=("order_id","count"))
                                   .reindex(day_order)
                                   .reset_index())
                fig = px.bar(
                    trend_df, x="day_of_week", y="orders",
                    title="Orders by Day of Week",
                    color="orders",
                    color_continuous_scale="Blues",
                    text="orders"
                )
                fig.update_layout(height=380, xaxis_title="Day", yaxis_title="Number of Orders")
                st.plotly_chart(fig, use_container_width=True)

                best_day = trend_df.loc[trend_df["orders"].idxmax(), "day_of_week"]
                show_insight(f"Your busiest day is <b>{best_day}</b>. Make sure you have enough stock and staff ready on that day.", "success")


        # ─────────────────────────────────────────────────────────────────
        # TAB 2: PRODUCT ANALYSIS
        # ─────────────────────────────────────────────────────────────────
        with tab2:
            st.subheader("🏆 Product Analysis")

            # ── Product summary table ──────────────────────────────────────
            product_df = (df_rev.groupby("product_name")
                                .agg(
                                    total_revenue = ("revenue",    "sum"),
                                    total_orders  = ("order_id",   "count"),
                                    total_qty     = ("quantity",   "sum"),
                                    avg_price     = ("unit_price", "mean")
                                )
                                .reset_index()
                                .sort_values("total_revenue", ascending=False))

            product_df["revenue_share_%"] = (
                product_df["total_revenue"] / product_df["total_revenue"].sum() * 100
            ).round(1)

            col1, col2 = st.columns(2)

            with col1:
                # ── TOP 5 PRODUCTS ─────────────────────────────────────────
                st.markdown("#### 🥇 Top 5 Products by Revenue")
                top5 = product_df.head(5)
                fig  = px.bar(
                    top5,
                    x="total_revenue", y="product_name",
                    orientation="h",
                    text="revenue_share_%",
                    color="total_revenue",
                    color_continuous_scale="Greens",
                    title="Top 5 Products"
                )
                fig.update_traces(texttemplate="%{text}%", textposition="outside")
                fig.update_layout(height=350, showlegend=False,
                                  yaxis=dict(autorange="reversed"),
                                  xaxis_title="Revenue (BDT)",
                                  yaxis_title="")
                st.plotly_chart(fig, use_container_width=True)

                best_product = product_df.iloc[0]["product_name"]
                best_share   = product_df.iloc[0]["revenue_share_%"]
                show_insight(f"<b>{best_product}</b> alone brings {best_share}% of total revenue. Keep it always in stock!", "success")

            with col2:
                # ── SLOW MOVERS ────────────────────────────────────────────
                st.markdown("#### 🐢 Slow Moving Products (Bottom 5)")
                bottom5 = product_df.tail(5).sort_values("total_revenue")
                fig = px.bar(
                    bottom5,
                    x="total_revenue", y="product_name",
                    orientation="h",
                    color="total_revenue",
                    color_continuous_scale="Reds",
                    title="Lowest Revenue Products"
                )
                fig.update_layout(height=350, showlegend=False,
                                  yaxis=dict(autorange="reversed"),
                                  xaxis_title="Revenue (BDT)",
                                  yaxis_title="")
                st.plotly_chart(fig, use_container_width=True)

                worst_product = bottom5.iloc[0]["product_name"]
                show_insight(f"<b>{worst_product}</b> is your slowest seller. Consider discounting it or removing from catalog.", "warning")

            # ── FULL PRODUCT TABLE ─────────────────────────────────────────
            st.markdown("#### 📋 Full Product Performance Table")
            display_df = product_df.copy()
            display_df["total_revenue"] = display_df["total_revenue"].apply(format_bdt)
            display_df["avg_price"]     = display_df["avg_price"].apply(format_bdt)
            display_df.columns = ["Product", "Revenue", "Orders", "Qty Sold", "Avg Price", "Revenue Share %"]
            st.dataframe(display_df, use_container_width=True, hide_index=True)

            # ── REVENUE SHARE PIE CHART ────────────────────────────────────
            st.markdown("#### 🥧 Revenue Share by Product")
            fig = px.pie(
                product_df, values="total_revenue", names="product_name",
                title="Revenue Distribution Across Products",
                hole=0.4,
                color_discrete_sequence=px.colors.qualitative.Set3
            )
            fig.update_traces(textposition="inside", textinfo="percent+label")
            fig.update_layout(height=450)
            st.plotly_chart(fig, use_container_width=True)


        # ─────────────────────────────────────────────────────────────────
        # TAB 3: PAYMENT & DELIVERY
        # ─────────────────────────────────────────────────────────────────
        with tab3:
            st.subheader("💳 Payment & Delivery Analysis")

            col1, col2 = st.columns(2)

            with col1:
                # ── PAYMENT METHOD BREAKDOWN ───────────────────────────────
                st.markdown("#### 💳 Payment Methods")
                pay_df = (df_filtered.groupby("payment_method")
                                     .agg(orders=("order_id","count"),
                                          revenue=("revenue","sum"))
                                     .reset_index()
                                     .sort_values("orders", ascending=False))

                fig = px.pie(
                    pay_df, values="orders", names="payment_method",
                    title="Orders by Payment Method",
                    hole=0.4,
                    color_discrete_sequence=px.colors.qualitative.Pastel
                )
                st.plotly_chart(fig, use_container_width=True)

                top_payment = pay_df.iloc[0]["payment_method"]
                show_insight(f"<b>{top_payment}</b> is your most used payment method. Make sure it never goes down.", "info")

            with col2:
                # ── DELIVERY STATUS ────────────────────────────────────────
                st.markdown("#### 🚚 Delivery Status")
                del_df = (df_filtered.groupby("delivery_status")
                                     .agg(orders=("order_id","count"))
                                     .reset_index())

                color_map = {
                    "Delivered": "#4CAF50",
                    "Returned":  "#F44336",
                    "Pending":   "#FF9800",
                    "Unknown":   "#9E9E9E"
                }
                fig = px.bar(
                    del_df, x="delivery_status", y="orders",
                    title="Orders by Delivery Status",
                    color="delivery_status",
                    color_discrete_map=color_map,
                    text="orders"
                )
                fig.update_traces(textposition="outside")
                fig.update_layout(height=380, showlegend=False,
                                  xaxis_title="Status", yaxis_title="Orders")
                st.plotly_chart(fig, use_container_width=True)

                returned = del_df[del_df["delivery_status"]=="Returned"]["orders"].sum() if "Returned" in del_df["delivery_status"].values else 0
                if returned > 0:
                    ret_pct = returned / len(df_filtered) * 100
                    if ret_pct > 10:
                        show_insight(f"⚠️ Return rate is <b>{ret_pct:.1f}%</b> — higher than ideal. Investigate top returned products.", "warning")
                    else:
                        show_insight(f"Return rate is <b>{ret_pct:.1f}%</b> — within acceptable range. Keep monitoring.", "success")

            # ── ACTIONABLE RECOMMENDATIONS ─────────────────────────────────
            st.markdown("---")
            st.subheader("🎯 Actionable Recommendations")

            rec_col1, rec_col2 = st.columns(2)

            with rec_col1:
                st.markdown("""
                | # | Recommendation | Effort | Impact |
                |---|---------------|--------|--------|
                | 1 | Stock up your top 3 products before peak months | 🟢 Low | 🔴 High |
                | 2 | Run a discount campaign on slow-moving products | 🟡 Medium | 🟡 Medium |
                | 3 | Offer bKash/Nagad cashback to boost payment adoption | 🟢 Low | 🔴 High |
                """)

            with rec_col2:
                st.markdown("""
                | # | Recommendation | Effort | Impact |
                |---|---------------|--------|--------|
                | 4 | Follow up on Pending orders within 24 hours | 🟢 Low | 🔴 High |
                | 5 | Bundle slow movers with top sellers as a combo deal | 🟡 Medium | 🟡 Medium |
                | 6 | Post on Facebook on your busiest day of the week | 🟢 Low | 🟡 Medium |
                """)
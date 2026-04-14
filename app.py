import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import numpy as np

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Sales Intelligence Hub",
    page_icon="🛒",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .stApp { background-color: #080c14; }
    section[data-testid="stSidebar"] {
        background-color: #0d1117;
        border-right: 1px solid #1f2937;
    }
    div[data-testid="metric-container"] {
        background: #0d1117;
        border: 1px solid #1f2937;
        border-radius: 12px;
        padding: 16px 20px;
    }
    div[data-testid="metric-container"]:hover {
        border-color: #3b82f6;
        transition: border-color 0.2s;
    }
    h1 { color: #f0f6fc !important; letter-spacing: -1px; }
    h2, h3 { color: #c9d1d9 !important; }
    .section-title {
        font-size: 13px; font-weight: 700; letter-spacing: 2px;
        color: #3b82f6; text-transform: uppercase; margin-bottom: 4px;
    }
    .chat-container {
        background: #0d1117;
        border: 1px solid #1f2937;
        border-radius: 12px;
        padding: 20px;
    }
    .divider { border-top: 1px solid #1f2937; margin: 28px 0; }
    [data-testid="stChatMessage"] { background: #0d1117 !important; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# DATA LOADING & PREPROCESSING
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data
def load_data():
    df = pd.read_csv(
        "/Users/neenubonny/Downloads/online_sales_dataset.csv",
        parse_dates=["InvoiceDate"],
    )
    # Clean & engineer
    df["PaymentMethod"] = df["PaymentMethod"].replace("paypall", "PayPal")
    df["Revenue"]    = (df["Quantity"] * df["UnitPrice"] * (1 - df["Discount"])).round(2)
    df["Profit"]     = (df["Revenue"] - df["ShippingCost"].fillna(0)).round(2)
    df["Month"]      = df["InvoiceDate"].dt.to_period("M").astype(str)
    df["Year"]       = df["InvoiceDate"].dt.year
    df["Quarter"]    = df["InvoiceDate"].dt.to_period("Q").astype(str)
    df["IsReturned"] = df["ReturnStatus"] == "Returned"
    df["IsOnline"]   = df["SalesChannel"] == "Online"
    return df

df = load_data()

# ── Country ISO codes for choropleth ──
ISO = {
    "France":"FRA","Germany":"DEU","United Kingdom":"GBR",
    "Sweden":"SWE","Netherlands":"NLD","Belgium":"BEL",
    "Portugal":"PRT","Norway":"NOR","Australia":"AUS",
    "Spain":"ESP","Italy":"ITA","Denmark":"DNK",
}

# ── Color palettes ──
CAT_COLORS   = px.colors.qualitative.Vivid
LAYOUT_BASE  = dict(
    paper_bgcolor="#080c14", plot_bgcolor="#0d1117",
    font=dict(color="#c9d1d9", family="Inter, sans-serif"),
    legend=dict(bgcolor="#0d1117", bordercolor="#1f2937", borderwidth=1),
    margin=dict(l=10, r=10, t=40, b=10),
)
AXIS_STYLE   = dict(gridcolor="#1f2937", zerolinecolor="#1f2937", color="#6b7280")

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR — FILTERS + API KEY
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## 🛒 Sales Intelligence")
    st.caption("Interactive dashboard · AI-powered insights")
    st.markdown("---")

    st.markdown("### 🎛️ Filters")
    countries   = st.multiselect("Country", sorted(df["Country"].unique()),
                                 default=sorted(df["Country"].unique()))
    categories  = st.multiselect("Category", sorted(df["Category"].unique()),
                                 default=sorted(df["Category"].unique()))
    channels    = st.multiselect("Sales Channel", df["SalesChannel"].unique().tolist(),
                                 default=df["SalesChannel"].unique().tolist())
    year_min    = int(df["Year"].min())
    year_max    = int(df["Year"].max())
    year_range  = st.slider("Year range", year_min, year_max, (year_min, year_max))

    st.markdown("---")
    st.markdown("### 🤖 AI Assistant")
    api_key = st.text_input(
        "Anthropic API Key",
        type="password",
        placeholder="sk-ant-...",
        help="Enter your Anthropic API key to enable the AI chatbot",
    )
    if api_key:
        st.success("AI Assistant ready ✓")
    else:
        st.info("Add API key above to enable chat")

# ── Apply filters ──────────────────────────────────────────────────────────────
fdf = df[
    df["Country"].isin(countries) &
    df["Category"].isin(categories) &
    df["SalesChannel"].isin(channels) &
    df["Year"].between(year_range[0], year_range[1])
].copy()

# ══════════════════════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("# 🛒 Sales Intelligence Hub")
st.caption(
    f"Analysing **{len(fdf):,} transactions** · "
    f"**{fdf['Country'].nunique()} countries** · "
    f"**{fdf['Category'].nunique()} categories** · "
    f"{fdf['InvoiceDate'].min().strftime('%b %Y')} → "
    f"{fdf['InvoiceDate'].max().strftime('%b %Y')}"
)

# ══════════════════════════════════════════════════════════════════════════════
# KPI CARDS
# ══════════════════════════════════════════════════════════════════════════════
total_rev        = fdf["Revenue"].sum()
total_orders     = len(fdf)
return_rate      = fdf["IsReturned"].mean() * 100
avg_order        = fdf["Revenue"].mean()
unique_customers = fdf["CustomerID"].nunique()
online_pct       = fdf["IsOnline"].mean() * 100

k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("💰 Total Revenue",   f"${total_rev:,.0f}")
k2.metric("📦 Total Orders",    f"{total_orders:,}")
k3.metric("👥 Customers",       f"{unique_customers:,}")
k4.metric("🧾 Avg Order Value", f"${avg_order:,.0f}")
k5.metric("🔄 Return Rate",     f"{return_rate:.1f}%")
k6.metric("🌐 Online Share",    f"{online_pct:.1f}%")

st.markdown("<div class='divider'></div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# ROW 1: Revenue Timeline + Sales Channel Split
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("<p class='section-title'>Revenue Trends</p>", unsafe_allow_html=True)
col_t, col_ch = st.columns([3, 1])

with col_t:
    st.subheader("Monthly Revenue & Orders Over Time")
    monthly_rev = (
        fdf.groupby(["Month", "SalesChannel"])["Revenue"]
        .sum().reset_index().sort_values("Month")
    )
    monthly_orders = (
        fdf.groupby("Month")["InvoiceNo"].count().reset_index()
        .rename(columns={"InvoiceNo": "Orders"}).sort_values("Month")
    )
    FILL_COLORS = {"Online": "rgba(59,130,246,0.1)", "In-store": "rgba(16,185,129,0.1)"}
    fig_time = go.Figure()
    for channel, color in [("Online","#3b82f6"), ("In-store","#10b981")]:
        m = monthly_rev[monthly_rev["SalesChannel"] == channel]
        fig_time.add_trace(go.Scatter(
            x=m["Month"], y=m["Revenue"],
            name=f"{channel} Revenue", mode="lines",
            line=dict(color=color, width=2.5),
            fill="tozeroy",
            fillcolor=FILL_COLORS[channel],
            hovertemplate=f"<b>{channel}</b><br>Month: %{{x}}<br>Revenue: $%{{y:,.0f}}<extra></extra>",
            yaxis="y1",
        ))
    fig_time.add_trace(go.Scatter(
        x=monthly_orders["Month"], y=monthly_orders["Orders"],
        name="Total Orders", mode="lines+markers",
        line=dict(color="#f59e0b", width=1.5, dash="dot"),
        marker=dict(size=3),
        hovertemplate="<b>Orders</b><br>Month: %{x}<br>Count: %{y:,}<extra></extra>",
        yaxis="y2",
    ))
    fig_time.update_layout(
        **LAYOUT_BASE, height=320,
        xaxis=dict(title="Month", **AXIS_STYLE, tickangle=-45,
                   tickvals=monthly_rev["Month"].unique()[::6]),
        yaxis=dict(title="Revenue ($)", **AXIS_STYLE),
        yaxis2=dict(title="Order Count", overlaying="y", side="right",
                    **AXIS_STYLE, showgrid=False),
    )
    fig_time.update_layout(legend=dict(orientation="h", y=1.08, bgcolor="rgba(0,0,0,0)"))
    st.plotly_chart(fig_time, use_container_width=True)

with col_ch:
    st.subheader("Channel Mix")
    ch_rev = fdf.groupby("SalesChannel")["Revenue"].sum().reset_index()
    fig_ch = go.Figure(go.Pie(
        labels=ch_rev["SalesChannel"], values=ch_rev["Revenue"],
        hole=0.6,
        marker=dict(colors=["#3b82f6","#10b981"],
                    line=dict(color="#080c14", width=2)),
        hovertemplate="<b>%{label}</b><br>Revenue: $%{value:,.0f}<br>Share: %{percent}<extra></extra>",
        textinfo="percent+label", textfont=dict(size=11),
    ))
    fig_ch.update_layout(**LAYOUT_BASE, height=320, showlegend=False,
        annotations=[dict(text="Channel<br>Split", x=0.5, y=0.5,
                         font_size=12, showarrow=False, font_color="#c9d1d9")])
    st.plotly_chart(fig_ch, use_container_width=True)

st.markdown("<div class='divider'></div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# ROW 2: Category Treemap + Country Choropleth
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("<p class='section-title'>Geographic & Category Performance</p>", unsafe_allow_html=True)
col_tree, col_map = st.columns(2)

with col_tree:
    st.subheader("Category Revenue & Return Rate")
    cat_stats = fdf.groupby("Category").agg(
        Revenue=("Revenue","sum"),
        Orders=("InvoiceNo","count"),
        ReturnRate=("IsReturned","mean"),
    ).reset_index()
    cat_stats["ReturnRate"] = (cat_stats["ReturnRate"] * 100).round(1)

    fig_tree = go.Figure(go.Treemap(
        labels=cat_stats["Category"],
        values=cat_stats["Revenue"],
        parents=[""] * len(cat_stats),
        customdata=cat_stats[["Orders","ReturnRate"]].values,
        texttemplate="<b>%{label}</b><br>$%{value:,.0f}<br>Return: %{customdata[1]:.1f}%",
        hovertemplate=(
            "<b>%{label}</b><br>"
            "Revenue: $%{value:,.0f}<br>"
            "Orders: %{customdata[0]:,}<br>"
            "Return Rate: %{customdata[1]:.1f}%"
            "<extra></extra>"
        ),
        marker=dict(
            colors=cat_stats["ReturnRate"],
            colorscale=[[0,"#10b981"],[0.5,"#f59e0b"],[1,"#ef4444"]],
            colorbar=dict(
                title="Return %",
                tickfont=dict(color="#c9d1d9"),
                titlefont=dict(color="#c9d1d9"),
                x=1.02,
            ),
        ),
    ))
    fig_tree.update_layout(**LAYOUT_BASE, height=380)
    st.plotly_chart(fig_tree, use_container_width=True)
    st.caption("Color = return rate (🟢 low · 🟡 medium · 🔴 high) · Size = revenue")

with col_map:
    st.subheader("Revenue by Country")
    top_n = st.slider("Show top N countries", min_value=3, max_value=len(df["Country"].unique()), value=10, key="top_n_map")
    country_rev = fdf.groupby("Country").agg(
        Revenue=("Revenue","sum"),
        Orders=("InvoiceNo","count"),
        ReturnRate=("IsReturned","mean"),
    ).reset_index().nlargest(top_n, "Revenue")
    country_rev["ISO"]        = country_rev["Country"].map(ISO)
    country_rev["ReturnRate"] = (country_rev["ReturnRate"] * 100).round(1)

    fig_map = go.Figure(go.Choropleth(
        locations=country_rev["ISO"],
        z=country_rev["Revenue"],
        customdata=country_rev[["Country","Orders","ReturnRate"]].values,
        colorscale=[[0,"#0d1117"],[0.3,"#1d4ed8"],[0.7,"#3b82f6"],[1,"#93c5fd"]],
        marker_line_color="#1f2937",
        marker_line_width=0.5,
        colorbar=dict(
            title="Revenue ($)",
            tickfont=dict(color="#c9d1d9"),
            titlefont=dict(color="#c9d1d9"),
        ),
        hovertemplate=(
            "<b>%{customdata[0]}</b><br>"
            "Revenue: $%{z:,.0f}<br>"
            "Orders: %{customdata[1]:,}<br>"
            "Return Rate: %{customdata[2]:.1f}%"
            "<extra></extra>"
        ),
    ))
    fig_map.update_layout(
        **LAYOUT_BASE, height=380,
        geo=dict(
            bgcolor="#080c14",
            showframe=False,
            showcoastlines=True,
            coastlinecolor="#1f2937",
            showland=True, landcolor="#0d1117",
            showocean=True, oceancolor="#080c14",
            showlakes=False,
            projection_type="natural earth",
        ),
    )
    st.plotly_chart(fig_map, use_container_width=True)

st.markdown("<div class='divider'></div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# ROW 3: Payment Sunburst + Warehouse Heatmap
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("<p class='section-title'>Payment & Operations Analysis</p>", unsafe_allow_html=True)
col_pay, col_wh = st.columns(2)

with col_pay:
    st.subheader("Payment Method × Sales Channel")
    pay_df    = fdf.groupby(["PaymentMethod","SalesChannel"])["Revenue"].sum().reset_index()
    pay_total = pay_df.groupby("PaymentMethod")["Revenue"].sum().reset_index()

    # Inner ring — unique IDs = payment method name
    p_ids     = pay_total["PaymentMethod"].tolist()
    p_labels  = pay_total["PaymentMethod"].tolist()
    p_parents = [""] * len(pay_total)
    p_values  = pay_total["Revenue"].tolist()

    # Outer ring — unique IDs = "PaymentMethod_Channel"
    c_ids     = [f"{r['PaymentMethod']}_{r['SalesChannel']}" for _, r in pay_df.iterrows()]
    c_labels  = pay_df["SalesChannel"].tolist()
    c_parents = pay_df["PaymentMethod"].tolist()
    c_values  = pay_df["Revenue"].tolist()

    fig_sun = go.Figure(go.Sunburst(
        ids          = p_ids + c_ids,
        labels       = p_labels + c_labels,
        parents      = p_parents + c_parents,
        values       = p_values + c_values,
        branchvalues = "total",
        marker=dict(colors=["#3b82f6","#10b981","#f59e0b",
                             "#1d4ed8","#059669","#d97706",
                             "#60a5fa","#34d399","#fbbf24"]),
        hovertemplate="<b>%{label}</b><br>Revenue: $%{value:,.0f}<extra></extra>",
        insidetextfont=dict(size=11),
    ))
    fig_sun.update_layout(**LAYOUT_BASE, height=380)
    st.plotly_chart(fig_sun, use_container_width=True)
    st.caption("Inner ring = payment method · Outer ring = sales channel within that method")

with col_wh:
    st.subheader("Revenue by Category & Sales Channel")
    cat_ch = fdf.groupby(["Category","SalesChannel"])["Revenue"].sum().reset_index()
    fig_catbar = go.Figure()
    for ch, color in [("Online","#3b82f6"), ("In-store","#10b981")]:
        c = cat_ch[cat_ch["SalesChannel"] == ch]
        fig_catbar.add_trace(go.Bar(
            x=c["Category"], y=c["Revenue"],
            name=ch,
            marker_color=color,
            hovertemplate=f"<b>%{{x}} · {ch}</b><br>Revenue: $%{{y:,.0f}}<extra></extra>",
        ))
    fig_catbar.update_layout(
        **LAYOUT_BASE, height=380,
        barmode="group",
        xaxis=dict(title="Category", **AXIS_STYLE, tickangle=-20),
        yaxis=dict(title="Revenue ($)", **AXIS_STYLE),
    )
    fig_catbar.update_layout(legend=dict(orientation="h", y=1.08, bgcolor="rgba(0,0,0,0)"))
    st.plotly_chart(fig_catbar, use_container_width=True)
    st.caption("Grouped bars show Online vs In-store revenue split per category")

st.markdown("<div class='divider'></div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# ROW 4: Return Analysis + Order Priority
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("<p class='section-title'>Returns & Order Priority</p>", unsafe_allow_html=True)
col_ret, col_pri = st.columns(2)

with col_ret:
    st.subheader("Return Rate by Country & Category")
    ret_df = fdf.groupby(["Country","Category"])["IsReturned"].mean().reset_index()
    ret_df["ReturnRate"] = (ret_df["IsReturned"] * 100).round(1)
    ret_pivot = ret_df.pivot(index="Country", columns="Category", values="ReturnRate").fillna(0)

    fig_ret = go.Figure(go.Heatmap(
        z=ret_pivot.values,
        x=ret_pivot.columns.tolist(),
        y=ret_pivot.index.tolist(),
        colorscale=[[0,"#10b981"],[0.5,"#f59e0b"],[1,"#ef4444"]],
        hovertemplate="<b>%{y} · %{x}</b><br>Return Rate: %{z:.1f}%<extra></extra>",
        text=[[f"{v:.0f}%" for v in row] for row in ret_pivot.values],
        texttemplate="%{text}", textfont=dict(size=10, color="white"),
        colorbar=dict(
            title="Return %",
            tickfont=dict(color="#c9d1d9"),
            titlefont=dict(color="#c9d1d9"),
        ),
    ))
    fig_ret.update_layout(
        **LAYOUT_BASE, height=360,
        xaxis=dict(color="#c9d1d9", tickangle=-30),
        yaxis=dict(color="#c9d1d9"),
    )
    st.plotly_chart(fig_ret, use_container_width=True)

with col_pri:
    st.subheader("Order Priority Distribution by Category")
    pri_df = fdf.groupby(["Category","OrderPriority"])["Revenue"].sum().reset_index()
    pri_colors = {"High":"#ef4444","Medium":"#f59e0b","Low":"#10b981"}

    fig_pri = go.Figure()
    for priority in ["High","Medium","Low"]:
        p = pri_df[pri_df["OrderPriority"] == priority]
        fig_pri.add_trace(go.Bar(
            x=p["Category"], y=p["Revenue"],
            name=priority,
            marker_color=pri_colors[priority],
            hovertemplate=f"<b>%{{x}} · {priority}</b><br>Revenue: $%{{y:,.0f}}<extra></extra>",
        ))
    fig_pri.update_layout(
        **LAYOUT_BASE, height=360, barmode="stack",
        xaxis=dict(title="Category", **AXIS_STYLE),
        yaxis=dict(title="Revenue ($)", **AXIS_STYLE),
    )
    st.plotly_chart(fig_pri, use_container_width=True)

st.markdown("<div class='divider'></div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# UNIQUE CHART 1 — SANKEY: Full Sales Journey
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("<p class='section-title'>Sales Journey Flow</p>", unsafe_allow_html=True)
st.subheader("🌊 Where Every Order Flows — Country → Category → Channel → Outcome")
st.caption("Follow revenue as it flows from origin country through category and channel to final outcome")

# Build Sankey nodes & links
top_countries = fdf.groupby("Country")["Revenue"].sum().nlargest(6).index.tolist()
san_df = fdf[fdf["Country"].isin(top_countries)].copy()

categories_list = sorted(san_df["Category"].unique().tolist())
channels_list   = san_df["SalesChannel"].unique().tolist()
outcomes_list   = san_df["ReturnStatus"].unique().tolist()

all_nodes   = top_countries + categories_list + channels_list + outcomes_list
node_idx    = {n: i for i, n in enumerate(all_nodes)}
node_colors = (
    ["#3b82f6"] * len(top_countries) +
    ["#10b981"] * len(categories_list) +
    ["#f59e0b"] * len(channels_list) +
    ["#ef4444", "#22c55e"]
)

src, tgt, val, lnk_colors = [], [], [], []

# Layer 1: Country → Category
for (country, cat), g in san_df.groupby(["Country", "Category"]):
    if country in node_idx and cat in node_idx:
        src.append(node_idx[country]); tgt.append(node_idx[cat])
        val.append(g["Revenue"].sum())
        lnk_colors.append("rgba(59,130,246,0.25)")

# Layer 2: Category → Channel
for (cat, ch), g in san_df.groupby(["Category", "SalesChannel"]):
    src.append(node_idx[cat]); tgt.append(node_idx[ch])
    val.append(g["Revenue"].sum())
    lnk_colors.append("rgba(16,185,129,0.25)")

# Layer 3: Channel → Return Status
for (ch, outcome), g in san_df.groupby(["SalesChannel", "ReturnStatus"]):
    src.append(node_idx[ch]); tgt.append(node_idx[outcome])
    val.append(g["Revenue"].sum())
    color = "rgba(239,68,68,0.3)" if outcome == "Returned" else "rgba(34,197,94,0.3)"
    lnk_colors.append(color)

fig_sankey = go.Figure(go.Sankey(
    arrangement="snap",
    node=dict(
        pad=18, thickness=20,
        label=all_nodes,
        color=node_colors[:len(all_nodes)],
        hovertemplate="%{label}<br>Revenue: $%{value:,.0f}<extra></extra>",
    ),
    link=dict(
        source=src, target=tgt, value=val,
        color=lnk_colors,
        hovertemplate="$%{value:,.0f}<extra></extra>",
    ),
))
fig_sankey.update_layout(**LAYOUT_BASE, height=500)
st.plotly_chart(fig_sankey, use_container_width=True)

st.markdown("<div class='divider'></div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# UNIQUE CHART 2 — WATERFALL: Revenue Breakdown
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("<p class='section-title'>Profit Waterfall</p>", unsafe_allow_html=True)
st.subheader("💧 Where Does Revenue Go? — Gross to Net Breakdown")
st.caption("See exactly how gross revenue erodes to net profit through discounts and shipping")

gross_rev   = (fdf["Quantity"] * fdf["UnitPrice"]).sum()
discount_loss = (fdf["Quantity"] * fdf["UnitPrice"] * fdf["Discount"]).sum()
shipping_cost = fdf["ShippingCost"].fillna(0).sum()
net_profit    = gross_rev - discount_loss - shipping_cost

wf_labels  = ["Gross Revenue", "Discounts", "Shipping Costs", "Net Profit"]
wf_values  = [gross_rev, -discount_loss, -shipping_cost, net_profit]
wf_measure = ["absolute", "relative", "relative", "total"]
wf_colors  = ["#3b82f6", "#ef4444", "#f59e0b", "#10b981"]

fig_wf = go.Figure(go.Waterfall(
    name="Revenue",
    orientation="v",
    measure=wf_measure,
    x=wf_labels,
    y=wf_values,
    text=[f"${abs(v)/1e6:.1f}M" for v in wf_values],
    textposition="outside",
    textfont=dict(color="#c9d1d9", size=13),
    connector=dict(line=dict(color="#1f2937", width=1.5)),
    increasing=dict(marker=dict(color="#10b981")),
    decreasing=dict(marker=dict(color="#ef4444")),
    totals=dict(marker=dict(color="#3b82f6")),
    hovertemplate="<b>%{x}</b><br>$%{y:,.0f}<extra></extra>",
))
fig_wf.update_layout(
    **LAYOUT_BASE, height=420,
    xaxis=dict(color="#c9d1d9"),
    yaxis=dict(title="Amount ($)", **AXIS_STYLE),
    showlegend=False,
)
st.plotly_chart(fig_wf, use_container_width=True)

wf_margin = (net_profit / gross_rev * 100)
col_wf1, col_wf2, col_wf3 = st.columns(3)
col_wf1.metric("Gross Revenue",   f"${gross_rev/1e6:.1f}M")
col_wf2.metric("Total Discounts", f"${discount_loss/1e6:.1f}M", f"-{discount_loss/gross_rev*100:.1f}%", delta_color="inverse")
col_wf3.metric("Net Margin",      f"{wf_margin:.1f}%")

st.markdown("<div class='divider'></div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# UNIQUE CHART 3 — HOURLY SALES CLOCK (Polar)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("<p class='section-title'>Sales Clock</p>", unsafe_allow_html=True)
col_clock, col_dow = st.columns(2)

with col_clock:
    st.subheader("🕐 When Do Customers Buy? — 24-Hour Sales Clock")
    st.caption("Each bar = average hourly revenue · Rotate around the clock")

    fdf["Hour"]      = fdf["InvoiceDate"].dt.hour
    hourly           = fdf.groupby("Hour")["Revenue"].mean().reset_index()
    hourly["Theta"]  = hourly["Hour"] * 15   # 360° / 24 hours

    fig_clock = go.Figure(go.Barpolar(
        r=hourly["Revenue"],
        theta=hourly["Theta"],
        width=14,
        marker=dict(
            color=hourly["Revenue"],
            colorscale=[[0,"#1f2937"],[0.5,"#3b82f6"],[1,"#93c5fd"]],
            showscale=False,
            line=dict(color="#080c14", width=1),
        ),
        hovertemplate="<b>%{customdata:02d}:00</b><br>Avg Revenue: $%{r:,.0f}<extra></extra>",
        customdata=hourly["Hour"],
    ))
    fig_clock.update_layout(
        **LAYOUT_BASE, height=400,
        polar=dict(
            bgcolor="#0d1117",
            radialaxis=dict(visible=True, color="#374151", gridcolor="#1f2937",
                           tickfont=dict(size=8, color="#6b7280")),
            angularaxis=dict(
                tickmode="array",
                tickvals=list(range(0, 360, 30)),
                ticktext=["12am","2am","4am","6am","8am","10am",
                          "12pm","2pm","4pm","6pm","8pm","10pm"],
                direction="clockwise", rotation=90,
                color="#6b7280", gridcolor="#1f2937",
            ),
        ),
        showlegend=False,
    )
    st.plotly_chart(fig_clock, use_container_width=True)

with col_dow:
    st.subheader("📅 Day of Week × Category Heatmap")
    st.caption("Which category peaks on which day?")

    DOW_ORDER = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
    fdf["DayOfWeek"] = fdf["InvoiceDate"].dt.day_name()
    dow_cat = (
        fdf.groupby(["DayOfWeek","Category"])["Revenue"]
        .sum().reset_index()
    )
    dow_pivot = (
        dow_cat.pivot(index="DayOfWeek", columns="Category", values="Revenue")
        .reindex(DOW_ORDER).fillna(0)
    )

    fig_dow = go.Figure(go.Heatmap(
        z=dow_pivot.values,
        x=dow_pivot.columns.tolist(),
        y=dow_pivot.index.tolist(),
        colorscale=[[0,"#0d1117"],[0.5,"#1d4ed8"],[1,"#93c5fd"]],
        hovertemplate="<b>%{y} · %{x}</b><br>Revenue: $%{z:,.0f}<extra></extra>",
        text=[[f"${v/1e3:.0f}K" for v in row] for row in dow_pivot.values],
        texttemplate="%{text}",
        textfont=dict(size=10),
        colorbar=dict(tickfont=dict(color="#c9d1d9"), titlefont=dict(color="#c9d1d9")),
    ))
    fig_dow.update_layout(
        **LAYOUT_BASE, height=400,
        xaxis=dict(color="#c9d1d9", tickangle=-20),
        yaxis=dict(color="#c9d1d9"),
    )
    st.plotly_chart(fig_dow, use_container_width=True)

st.markdown("<div class='divider'></div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# UNIQUE CHART 4 — CUSTOMER LOYALTY BUBBLE
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("<p class='section-title'>Customer Intelligence</p>", unsafe_allow_html=True)
st.subheader("🫧 Customer Loyalty Map — Who Buys Again & Again?")
st.caption(
    "X = average revenue per order · Y = number of orders · "
    "Size = total lifetime revenue · Color = top category purchased"
)

cust_df = (
    fdf.dropna(subset=["CustomerID"])
    .groupby("CustomerID")
    .agg(
        Orders       = ("InvoiceNo",  "count"),
        TotalRev     = ("Revenue",    "sum"),
        AvgRev       = ("Revenue",    "mean"),
        TopCategory  = ("Category",   lambda x: x.mode()[0]),
        ReturnRate   = ("IsReturned", "mean"),
    )
    .reset_index()
)
# Focus on repeat buyers with >1 order
repeat = cust_df[cust_df["Orders"] > 1].copy()

CAT_COLOR_MAP = {
    "Electronics": "#3b82f6", "Apparel": "#10b981",
    "Furniture": "#f59e0b",   "Accessories": "#a855f7",
    "Stationery": "#ef4444",
}

fig_bubble = go.Figure()
for cat in fdf["Category"].unique():
    sub = repeat[repeat["TopCategory"] == cat]
    if sub.empty:
        continue
    fig_bubble.add_trace(go.Scatter(
        x=sub["AvgRev"],
        y=sub["Orders"],
        mode="markers",
        name=cat,
        marker=dict(
            size=np.clip(sub["TotalRev"] / 500, 5, 35),
            color=CAT_COLOR_MAP.get(cat, "#888"),
            opacity=0.65,
            line=dict(color="white", width=0.5),
        ),
        customdata=sub[["TotalRev","ReturnRate","CustomerID"]].values,
        hovertemplate=(
            "<b>Customer %{customdata[2]:.0f}</b><br>"
            "Avg Order: $%{x:,.0f}<br>"
            "Total Orders: %{y}<br>"
            "Lifetime Value: $%{customdata[0]:,.0f}<br>"
            "Return Rate: %{customdata[1]:.0%}"
            "<extra>" + cat + "</extra>"
        ),
    ))

fig_bubble.add_hline(
    y=repeat["Orders"].median(), line_dash="dash",
    line_color="#374151", line_width=1,
    annotation_text="Median orders",
    annotation_font=dict(color="#6b7280", size=10),
)
fig_bubble.add_vline(
    x=repeat["AvgRev"].median(), line_dash="dash",
    line_color="#374151", line_width=1,
    annotation_text="Median avg order",
    annotation_font=dict(color="#6b7280", size=10),
)
fig_bubble.update_layout(
    **LAYOUT_BASE, height=480,
    xaxis=dict(title="Avg Revenue per Order ($)", **AXIS_STYLE),
    yaxis=dict(title="Number of Orders", **AXIS_STYLE),
)
st.plotly_chart(fig_bubble, use_container_width=True)

st.markdown("<div class='divider'></div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# UNIQUE CHART 5 — PARALLEL COORDINATES: Multi-Dimensional Order View
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("<p class='section-title'>Multi-Dimensional Analysis</p>", unsafe_allow_html=True)
st.subheader("🔀 Parallel Coordinates — Every Order's DNA")
st.caption(
    "Each line = one order · Drag axes to reorder · "
    "Click & drag on any axis to filter · Color = category"
)

CAT_NUMS = {c: i for i, c in enumerate(sorted(fdf["Category"].unique()))}
sample   = fdf.sample(min(3000, len(fdf)), random_state=42).copy()
sample["CatNum"] = sample["Category"].map(CAT_NUMS)

fig_par = go.Figure(go.Parcoords(
    line=dict(
        color=sample["CatNum"],
        colorscale=[[0,"#3b82f6"],[0.25,"#10b981"],[0.5,"#f59e0b"],
                    [0.75,"#a855f7"],[1,"#ef4444"]],
        showscale=False,
    ),
    dimensions=[
        dict(label="Quantity",      values=sample["Quantity"],
             range=[sample["Quantity"].min(), sample["Quantity"].quantile(0.99)]),
        dict(label="Unit Price ($)",values=sample["UnitPrice"],
             range=[sample["UnitPrice"].min(), sample["UnitPrice"].quantile(0.99)]),
        dict(label="Discount",      values=sample["Discount"],      range=[0, 1]),
        dict(label="Shipping ($)",  values=sample["ShippingCost"].fillna(0),
             range=[0, sample["ShippingCost"].quantile(0.99)]),
        dict(label="Revenue ($)",   values=sample["Revenue"],
             range=[sample["Revenue"].min(), sample["Revenue"].quantile(0.99)]),
        dict(label="Category",      values=sample["CatNum"],
             tickvals=list(CAT_NUMS.values()),
             ticktext=list(CAT_NUMS.keys())),
    ],
    labelfont=dict(color="#c9d1d9", size=12),
    tickfont=dict(color="#6b7280",  size=10),
    rangefont=dict(color="#6b7280", size=9),
))
fig_par.update_layout(**LAYOUT_BASE, height=420)
st.plotly_chart(fig_par, use_container_width=True)

st.markdown("<div class='divider'></div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# AI CHATBOT
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("<p class='section-title'>AI Data Assistant</p>", unsafe_allow_html=True)
st.subheader("🤖 Ask Anything About Your Sales Data")
st.caption("Powered by Claude · Answers are based on the filtered dataset above")

# Build data context for Claude
@st.cache_data
def build_context(n_rows):
    """Pre-compute stats so Claude can answer accurately."""
    top_country  = df.groupby("Country")["Revenue"].sum().idxmax()
    top_cat      = df.groupby("Category")["Revenue"].sum().idxmax()
    worst_return = df.groupby("Category")["IsReturned"].mean().idxmax()
    best_channel = df.groupby("SalesChannel")["Revenue"].sum().idxmax()
    monthly_peak = df.groupby("Month")["Revenue"].sum().idxmax()

    return f"""You are an expert data analyst AI assistant for a sales intelligence dashboard.
You have deep knowledge of the following dataset and must answer questions accurately using it.

DATASET OVERVIEW:
- Total records: {len(df):,}
- Date range: {df['InvoiceDate'].min().strftime('%b %Y')} to {df['InvoiceDate'].max().strftime('%b %Y')}
- Countries: {', '.join(sorted(df['Country'].unique()))} ({df['Country'].nunique()} total)
- Categories: {', '.join(sorted(df['Category'].unique()))}
- Sales channels: Online, In-store
- Payment methods: Bank Transfer, Credit Card, PayPal
- Shipment providers: FedEx, UPS, DHL, Royal Mail
- Warehouses: Amsterdam, Berlin, London, Paris, Rome

KEY METRICS:
- Total revenue: ${df['Revenue'].sum():,.0f}
- Average order value: ${df['Revenue'].mean():,.0f}
- Overall return rate: {df['IsReturned'].mean()*100:.1f}%
- Online share: {df['IsOnline'].mean()*100:.1f}%

TOP PERFORMERS:
- Highest revenue country: {top_country}
- Highest revenue category: {top_cat}
- Highest return rate category: {worst_return}
- Best performing channel: {best_channel}
- Peak revenue month: {monthly_peak}

CATEGORY BREAKDOWN:
{df.groupby('Category').agg(Revenue=('Revenue','sum'), Orders=('InvoiceNo','count'), ReturnRate=('IsReturned','mean')).round(2).to_string()}

COUNTRY BREAKDOWN:
{df.groupby('Country').agg(Revenue=('Revenue','sum'), Orders=('InvoiceNo','count')).round(2).to_string()}

PAYMENT METHOD BREAKDOWN:
{df.groupby('PaymentMethod')['Revenue'].sum().round(2).to_string()}

SHIPMENT PROVIDER BREAKDOWN:
{df.groupby('ShipmentProvider')['Revenue'].sum().round(2).to_string()}

Answer questions clearly and concisely. Use specific numbers from the data above.
If asked about trends or predictions, base them on the data patterns.
Keep answers under 150 words unless a detailed breakdown is specifically requested.
Format numbers with commas and dollar signs where appropriate."""

data_context = build_context(len(df))

# Chat UI
if "messages" not in st.session_state:
    st.session_state.messages = []

# Display history
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# Suggested questions
if not st.session_state.messages:
    st.markdown("**💡 Try asking:**")
    cols = st.columns(3)
    suggestions = [
        "Which country has the highest revenue?",
        "What is the return rate for Electronics?",
        "Which payment method is most popular?",
        "Which category is most profitable?",
        "How does online compare to in-store sales?",
        "Which warehouse ships the most Furniture?",
    ]
    for i, col in enumerate(cols):
        with col:
            for s in suggestions[i*2:(i+1)*2]:
                if st.button(s, key=s):
                    st.session_state.messages.append({"role":"user","content":s})
                    st.rerun()

# Chat input
if prompt := st.chat_input("Ask about your sales data..."):
    st.session_state.messages.append({"role":"user","content":prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    if not api_key:
        with st.chat_message("assistant"):
            st.warning(
                "Please enter your Anthropic API key in the sidebar to enable the AI assistant. "
                "You can get one at console.anthropic.com"
            )
        st.session_state.messages.append({
            "role":"assistant",
            "content":"⚠️ API key required. Add it in the sidebar to enable AI responses."
        })
    else:
        with st.chat_message("assistant"):
            with st.spinner("Analysing..."):
                try:
                    from anthropic import Anthropic
                    client = Anthropic(api_key=api_key)

                    api_messages = [
                        {"role": m["role"], "content": m["content"]}
                        for m in st.session_state.messages
                        if m["role"] in ("user", "assistant")
                        and not m["content"].startswith("⚠️")
                    ]

                    response = client.messages.create(
                        model="claude-haiku-4-5-20251001",
                        max_tokens=512,
                        system=data_context,
                        messages=api_messages,
                    )
                    reply = response.content[0].text
                    st.markdown(reply)
                    st.session_state.messages.append({"role":"assistant","content":reply})

                except Exception as e:
                    err = f"Error: {str(e)}"
                    st.error(err)
                    st.session_state.messages.append({"role":"assistant","content":err})

if st.session_state.messages:
    if st.button("🗑️ Clear chat", key="clear"):
        st.session_state.messages = []
        st.rerun()

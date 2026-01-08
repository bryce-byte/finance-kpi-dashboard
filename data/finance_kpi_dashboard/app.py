import io
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Finance KPI Dashboard", layout="wide")
st.title("Finance KPI Dashboard")

# ---- light UI polish (safe) ----
st.markdown(
    """
<style>
.block-container { padding-top: 1.2rem; padding-bottom: 2rem; }
section[data-testid="stSidebar"] .block-container { padding-top: 1rem; }
[data-testid="stDataFrame"] { border-radius: 12px; overflow: hidden; }
.stDownloadButton button { border-radius: 10px; padding: 0.55rem 0.85rem; }
</style>
""",
    unsafe_allow_html=True,
)

FILE_PATH = "data/finance_kpi.xlsx"


@st.cache_data
def load_data(path: str):
    actuals = pd.read_excel(path, sheet_name="Actuals")
    budget = pd.read_excel(path, sheet_name="Budget")
    return actuals, budget


def build_pdf_bytes(
    title: str,
    dept_choice: str,
    start_date,
    end_date,
    kpis: list,
    insights_text: str,
    wf_month: str | None,
):
    try:
        from reportlab.lib.pagesizes import LETTER
        from reportlab.lib.units import inch
        from reportlab.pdfgen import canvas
    except Exception:
        return None

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=LETTER)
    _, height = LETTER

    x = 0.75 * inch
    y = height - 0.9 * inch

    c.setFont("Helvetica-Bold", 18)
    c.drawString(x, y, title)
    y -= 0.35 * inch

    c.setFont("Helvetica", 11)
    c.drawString(x, y, f"Department: {dept_choice}")
    y -= 0.2 * inch
    c.drawString(x, y, f"Date Range: {start_date} to {end_date}")
    y -= 0.2 * inch
    if wf_month:
        c.drawString(x, y, f"Waterfall Month: {wf_month}")
        y -= 0.25 * inch
    else:
        y -= 0.1 * inch

    c.setFont("Helvetica-Bold", 12)
    c.drawString(x, y, "Key KPIs (Actual vs Budget)")
    y -= 0.25 * inch

    c.setFont("Helvetica", 10)
    for name, actual_str, delta_str in kpis:
        if y < 1.2 * inch:
            c.showPage()
            y = height - 0.9 * inch
            c.setFont("Helvetica", 10)
        c.drawString(x, y, f"{name}: {actual_str}   (Δ {delta_str})")
        y -= 0.18 * inch

    y -= 0.2 * inch
    c.setFont("Helvetica-Bold", 12)
    c.drawString(x, y, "Key Insights")
    y -= 0.25 * inch

    c.setFont("Helvetica", 10)
    words = insights_text.split()
    line = ""
    max_chars = 95
    for w in words:
        test = (line + " " + w).strip()
        if len(test) > max_chars:
            if y < 1.2 * inch:
                c.showPage()
                y = height - 0.9 * inch
                c.setFont("Helvetica", 10)
            c.drawString(x, y, line)
            y -= 0.16 * inch
            line = w
        else:
            line = test
    if line:
        c.drawString(x, y, line)

    c.showPage()
    c.save()
    buf.seek(0)
    return buf.getvalue()


try:
    actuals_df, budget_df = load_data(FILE_PATH)
    st.success("Loaded finance_kpi.xlsx ✅")

    # -------------------------
    # Sidebar filters (polished)
    # -------------------------
    st.sidebar.header("Filters")
    st.sidebar.caption("Adjust scope and the dashboard updates instantly.")

    departments = ["All"] + sorted(actuals_df["Department"].dropna().unique().tolist())
    dept_choice = st.sidebar.selectbox("Department", departments)

    if dept_choice != "All":
        actuals_view = actuals_df[actuals_df["Department"] == dept_choice].copy()
        budget_view = budget_df[budget_df["Department"] == dept_choice].copy()
    else:
        actuals_view = actuals_df.copy()
        budget_view = budget_df.copy()

    actuals_view["Date"] = pd.to_datetime(actuals_view["Date"])
    budget_view["Date"] = pd.to_datetime(budget_view["Date"])

    min_date = actuals_view["Date"].min().date()
    max_date = actuals_view["Date"].max().date()

    date_value = st.sidebar.date_input(
        "Date Range",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date,
    )

    if isinstance(date_value, (tuple, list)) and len(date_value) == 2:
        start_date, end_date = date_value
    else:
        start_date = end_date = date_value

    actuals_view = actuals_view[
        (actuals_view["Date"].dt.date >= start_date)
        & (actuals_view["Date"].dt.date <= end_date)
    ].copy()

    budget_view = budget_view[
        (budget_view["Date"].dt.date >= start_date)
        & (budget_view["Date"].dt.date <= end_date)
    ].copy()

    # ---- helpful flags for empty-state behavior ----
    has_revenue = (actuals_view["Account"] == "Revenue").any()
    has_cogs = (actuals_view["Account"] == "COGS").any()

    # -------------------------
    # KPIs (Actual + Budget + Deltas)
    # -------------------------
    act_revenue = actuals_view.loc[actuals_view["Account"] == "Revenue", "Amount"].sum()
    act_cogs = actuals_view.loc[actuals_view["Account"] == "COGS", "Amount"].sum()
    act_opex = actuals_view.loc[actuals_view["Account"] == "OpEx", "Amount"].sum()

    act_gross_margin = act_revenue + act_cogs
    act_net_income = act_gross_margin + act_opex
    act_gm_pct = (act_gross_margin / act_revenue) if act_revenue != 0 else 0

    bud_revenue = budget_view.loc[budget_view["Account"] == "Revenue", "Amount"].sum()
    bud_cogs = budget_view.loc[budget_view["Account"] == "COGS", "Amount"].sum()
    bud_opex = budget_view.loc[budget_view["Account"] == "OpEx", "Amount"].sum()

    bud_gross_margin = bud_revenue + bud_cogs
    bud_net_income = bud_gross_margin + bud_opex
    bud_gm_pct = (bud_gross_margin / bud_revenue) if bud_revenue != 0 else 0

    d_revenue = act_revenue - bud_revenue
    d_gm = act_gross_margin - bud_gross_margin
    d_gm_pct_pp = (act_gm_pct - bud_gm_pct) * 100
    d_opex = act_opex - bud_opex
    d_ni = act_net_income - bud_net_income
    d_opex_good = -d_opex  # savings positive

    kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)

    # Revenue + GM KPIs only make sense if revenue exists
    if has_revenue:
        kpi1.metric("Revenue", f"${act_revenue:,.0f}", delta=f"${d_revenue:,.0f}")
        kpi2.metric("Gross Margin", f"${act_gross_margin:,.0f}", delta=f"${d_gm:,.0f}")
        kpi3.metric("Gross Margin %", f"{act_gm_pct:.1%}", delta=f"{d_gm_pct_pp:+.1f} pp")
    else:
        kpi1.metric("Revenue", "$0", delta="$0")
        kpi2.metric("Gross Margin", "$0", delta="$0")
        kpi3.metric("Gross Margin %", "N/A", delta="")

    kpi4.metric(
        "Operating Expenses",
        f"${act_opex:,.0f}",
        delta=f"${d_opex_good:,.0f}",
        delta_color="inverse",
    )
    kpi5.metric("Net Income", f"${act_net_income:,.0f}", delta=f"${d_ni:,.0f}")

    # -------------------------
    # Insights
    # -------------------------
    st.subheader("Key Insights")
    insights = []

    if d_ni > 0:
        insights.append(f"Net income finished ${d_ni:,.0f} above budget.")
    elif d_ni < 0:
        insights.append(f"Net income finished ${abs(d_ni):,.0f} below budget.")
    else:
        insights.append("Net income finished in line with budget.")

    if has_revenue:
        if d_revenue < 0:
            insights.append("Revenue came in below plan.")
        elif d_revenue > 0:
            insights.append("Revenue exceeded plan.")

        if d_gm_pct_pp > 0:
            insights.append(f"Gross margin rate improved by {d_gm_pct_pp:.1f} pp.")
        elif d_gm_pct_pp < 0:
            insights.append(f"Gross margin rate declined by {abs(d_gm_pct_pp):.1f} pp.")

    if d_opex_good < 0:
        insights.append(f"Operating expenses were ${abs(d_opex_good):,.0f} higher than budget.")
    elif d_opex_good > 0:
        insights.append(f"Operating expenses were ${d_opex_good:,.0f} below budget.")

    insights_text = " ".join(insights)
    st.write(insights_text)

    # -------------------------
    # Tabs (UI polish)
    # -------------------------
    tab1, tab2, tab3 = st.tabs(["📈 Trends", "💰 Variance", "📋 Data & Export"])

    with tab1:
        st.subheader("Trends")

        if not has_revenue:
            st.info("This selection is a cost center (no revenue). Trends for Revenue and Gross Margin are not applicable.")
        else:
            actuals_month = actuals_view.copy()
            actuals_month["Month"] = actuals_month["Date"].dt.to_period("M").dt.to_timestamp()

            rev_m = (
                actuals_month.loc[actuals_month["Account"] == "Revenue"]
                .groupby("Month")["Amount"]
                .sum()
            )
            cogs_m = (
                actuals_month.loc[actuals_month["Account"] == "COGS"]
                .groupby("Month")["Amount"]
                .sum()
            )

            trend_df = pd.DataFrame({"Revenue": rev_m, "COGS": cogs_m}).fillna(0).sort_index()
            trend_df["GrossMargin"] = trend_df["Revenue"] + trend_df["COGS"]
            trend_df["GrossMarginPct"] = trend_df.apply(
                lambda r: (r["GrossMargin"] / r["Revenue"]) if r["Revenue"] != 0 else 0, axis=1
            )
            trend_df = trend_df.reset_index()

            left, right = st.columns([2, 1])
            with left:
                fig_rev = px.line(trend_df, x="Month", y="Revenue", title="Revenue by Month")
                fig_rev.update_yaxes(rangemode="tozero")
                st.plotly_chart(fig_rev, use_container_width=True)

            with right:
                fig = px.line(trend_df, x="Month", y="GrossMarginPct", title="Gross Margin %")
                fig.update_yaxes(tickformat=".0%", rangemode="tozero")
                st.plotly_chart(fig, use_container_width=True)

    with tab2:
        st.subheader("Variance")

        if has_revenue:
            st.markdown("**Budget vs Actual (Revenue)**")
            budget_month = budget_view.copy()
            budget_month["Month"] = budget_month["Date"].dt.to_period("M").dt.to_timestamp()

            # Recompute actuals_month safely for this tab if needed
            actuals_month = actuals_view.copy()
            actuals_month["Month"] = actuals_month["Date"].dt.to_period("M").dt.to_timestamp()

            act_rev = (
                actuals_month.loc[actuals_month["Account"] == "Revenue"]
                .groupby("Month")["Amount"]
                .sum()
            )
            bud_rev = (
                budget_month.loc[budget_month["Account"] == "Revenue"]
                .groupby("Month")["Amount"]
                .sum()
            )

            bva = pd.DataFrame({"Actual": act_rev, "Budget": bud_rev}).fillna(0).reset_index()
            fig_bva = px.bar(bva, x="Month", y=["Actual", "Budget"], barmode="group", title="Revenue: Budget vs Actual")
            fig_bva.update_yaxes(rangemode="tozero")
            st.plotly_chart(fig_bva, use_container_width=True)
        else:
            st.info("Revenue variance is not applicable for this selection.")

        st.markdown("**Net Income Variance (Selected Month)**")
        month_options = actuals_view["Date"].dt.to_period("M").astype(str).sort_values().unique()
        wf_month = None

        if len(month_options) == 0:
            st.warning("No data available for the selected filters.")
        else:
            wf_month = st.selectbox("Waterfall Month", month_options, index=len(month_options) - 1)
            wf_period = pd.Period(wf_month)

            act_m = actuals_view[actuals_view["Date"].dt.to_period("M") == wf_period]
            bud_m = budget_view[budget_view["Date"].dt.to_period("M") == wf_period]

            act_rev_t = act_m.loc[act_m["Account"] == "Revenue", "Amount"].sum()
            act_cogs_t = act_m.loc[act_m["Account"] == "COGS", "Amount"].sum()
            act_opex_t = act_m.loc[act_m["Account"] == "OpEx", "Amount"].sum()
            act_ni_m = act_rev_t + act_cogs_t + act_opex_t

            bud_rev_t = bud_m.loc[bud_m["Account"] == "Revenue", "Amount"].sum()
            bud_cogs_t = bud_m.loc[bud_m["Account"] == "COGS", "Amount"].sum()
            bud_opex_t = bud_m.loc[bud_m["Account"] == "OpEx", "Amount"].sum()
            bud_ni_m = bud_rev_t + bud_cogs_t + bud_opex_t

            rev_var = act_rev_t - bud_rev_t
            cogs_impact = -(act_cogs_t - bud_cogs_t)
            opex_impact = -(act_opex_t - bud_opex_t)

            fig_wf = go.Figure(
                go.Waterfall(
                    x=["Budget NI", "Revenue", "COGS", "OpEx", "Actual NI"],
                    measure=["absolute", "relative", "relative", "relative", "total"],
                    y=[bud_ni_m, rev_var, cogs_impact, opex_impact, act_ni_m],
                    texttemplate="$%{y:,.0f}",
                    textposition="outside",
                )
            )
            st.plotly_chart(fig_wf, use_container_width=True)

            st.caption(
                f"Month: {wf_month} | Budget NI: ${bud_ni_m:,.0f} | Actual NI: ${act_ni_m:,.0f} | Variance: ${(act_ni_m - bud_ni_m):,.0f}"
            )

    with tab3:
        st.subheader("Export")

        buffer_xlsx = io.BytesIO()
        with pd.ExcelWriter(buffer_xlsx, engine="openpyxl") as writer:
            actuals_view.to_excel(writer, index=False, sheet_name="Actuals_Filtered")
            budget_view.to_excel(writer, index=False, sheet_name="Budget_Filtered")

        st.download_button(
            label="⬇️ Download Excel (Filtered)",
            data=buffer_xlsx.getvalue(),
            file_name="finance_kpi_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        kpi_rows_for_pdf = [
            ("Revenue", f"${act_revenue:,.0f}", f"${d_revenue:,.0f}"),
            ("Gross Margin", f"${act_gross_margin:,.0f}", f"${d_gm:,.0f}"),
            ("Gross Margin %", f"{act_gm_pct:.1%}", f"{d_gm_pct_pp:+.1f} pp"),
            ("Operating Expenses", f"${act_opex:,.0f}", f"${d_opex_good:,.0f}"),
            ("Net Income", f"${act_net_income:,.0f}", f"${d_ni:,.0f}"),
        ]

        pdf_bytes = build_pdf_bytes(
            title="Finance KPI Dashboard - Executive Summary",
            dept_choice=dept_choice,
            start_date=start_date,
            end_date=end_date,
            kpis=kpi_rows_for_pdf,
            insights_text=insights_text,
            wf_month=wf_month,
        )

        if pdf_bytes is None:
            st.warning("PDF export is unavailable (reportlab not detected). Use the Excel export, or Print-to-PDF from your browser.")
        else:
            st.download_button(
                label="⬇️ Download PDF (Executive Summary)",
                data=pdf_bytes,
                file_name="finance_kpi_export.pdf",
                mime="application/pdf",
            )

        st.divider()
        st.subheader("Data Preview")

        st.write("Actuals (filtered)")
        st.dataframe(actuals_view, use_container_width=True)

        st.write("Budget (filtered)")
        st.dataframe(budget_view, use_container_width=True)

except Exception as e:
    st.error("Something went wrong.")
    st.write(e)

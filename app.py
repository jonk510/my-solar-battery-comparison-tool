"""
Solar & Battery Comparison App — Synergy WA
Upload your half-hourly meter data, configure two options, compare side-by-side.
Run with:  streamlit run app.py
"""
import sys, os
import io
import warnings
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import streamlit as st

warnings.filterwarnings("ignore")

# ── Import simulation engine from the analyser ────────────────────────────────
sys.path.insert(0, os.path.dirname(__file__))
from solar_battery_analyser import (
    validate_consumption_data,
    add_solar, simulate, payback, baseline,
    ANALYSIS_YEARS, OPPORTUNITY_RATE,
    A1_RATE, A1_SUPPLY, MS_SUPPLY,
    TARIFF_ESC, DEBS_DECL,
    BAT_DOD, BAT_RTE, BAT_DEG, SOL_DEG, SYS_EFF,
    ms_rate, debs_rate, _monthly_net,
    CS,
)

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Solar + Battery Analyser",
    page_icon="☀️",
    layout="wide",
)

SEASON_MONTHS = {
    "Summer (Dec–Feb)": [12, 1, 2],
    "Autumn (Mar–May)": [3, 4, 5],
    "Winter (Jun–Aug)": [6, 7, 8],
    "Spring (Sep–Nov)": [9, 10, 11],
}
OPTION_COLOURS = ["#e8463a", "#0f9d58"]   # red for A, green for B

# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def cached_load(file_bytes: bytes, filename: str) -> pd.DataFrame:
    buf    = io.BytesIO(file_bytes)
    suffix = os.path.splitext(filename)[1].lower()

    if suffix in (".xlsx", ".xls"):
        raw = pd.read_excel(buf)
    elif suffix == ".tsv":
        raw = pd.read_csv(buf, sep="\t")
    else:
        raw = pd.read_csv(buf)

    raw.columns = raw.columns.str.strip()

    # Mirror load_data column detection
    if "Date" in raw.columns and "Time" in raw.columns:
        raw["datetime"] = pd.to_datetime(
            raw["Date"].astype(str) + " " + raw["Time"].astype(str),
            dayfirst=True, errors="coerce",
        )
        kwh_col = next(
            (c for c in raw.columns if c.lower() in ("kwh", "consumption", "usage", "energy")),
            None,
        )
        if not kwh_col:
            raise ValueError(
                f"Found Date+Time columns but no kWh/consumption column. "
                f"Available: {list(raw.columns)}"
            )
        raw["consumption_kwh"] = pd.to_numeric(raw[kwh_col], errors="coerce")
    else:
        raw.columns = raw.columns.str.lower().str.replace(" ", "_")
        dc = next((c for c in raw.columns if "date" in c or "time" in c), None)
        kc = next(
            (c for c in raw.columns if any(k in c for k in ["kwh", "consumption", "usage", "energy"])),
            None,
        )
        if not dc:
            raise ValueError(f"Could not detect a datetime column. Found: {list(raw.columns)}")
        if not kc:
            raise ValueError(f"Could not detect a consumption column. Found: {list(raw.columns)}")
        raw["datetime"] = pd.to_datetime(raw[dc], dayfirst=True, errors="coerce")
        raw["consumption_kwh"] = pd.to_numeric(raw[kc], errors="coerce")

    raw = raw[raw["datetime"].notna()].reset_index(drop=True)
    df  = raw[["datetime", "consumption_kwh"]].sort_values("datetime").reset_index(drop=True)
    df  = validate_consumption_data(df)

    df["slot"] = (df["datetime"].dt.hour * 2 + df["datetime"].dt.minute // 30).astype(int)
    df["doy"]  = df["datetime"].dt.dayofyear.astype(int)
    df["date"] = df["datetime"].dt.date
    return df


@st.cache_data(show_spinner=False)
def cached_simulate(raw_df_hash, solar_kw, bat_kwh, inv_kw, tariff, _raw_df):
    df_solar = add_solar(_raw_df, solar_kw)
    return simulate(df_solar, solar_kw, bat_kwh, inv_kw, tariff)


@st.cache_data(show_spinner=False)
def cached_payback(raw_df_hash, solar_kw, bat_kwh, inv_kw, cost, tariff, label, _raw_df):
    return payback(_raw_df, solar_kw, bat_kwh, inv_kw, cost, tariff, label)


def seasonal_daily_cost(res: dict) -> dict:
    """Return average daily net cost per season from a simulate() result."""
    df  = res["df"].copy()
    df["month"] = df["datetime"].dt.month
    df["date"]  = df["datetime"].dt.date
    slots = df["slot"].values.astype(int)

    if res["tariff"] == "A1 Flat":
        df["ic"] = df["g_imp"] * A1_RATE
        sc = A1_SUPPLY
    else:
        df["ic"] = df["g_imp"] * np.vectorize(ms_rate)(slots)
        sc = MS_SUPPLY
    df["ec"] = df["g_exp"] * np.vectorize(debs_rate)(slots)

    costs = {}
    for season, months in SEASON_MONTHS.items():
        sub = df[df["month"].isin(months)]
        if sub.empty:
            costs[season] = None
            continue
        total_ic   = sub["ic"].sum()
        total_ec   = sub["ec"].sum()
        n_days     = sub["date"].nunique()
        total_sc   = sc * n_days
        costs[season] = (total_ic + total_sc - total_ec) / n_days if n_days else None
    return costs


def baseline_seasonal_cost(raw_df: pd.DataFrame, tariff: str) -> dict:
    """Average daily baseline cost per season (no solar/battery)."""
    df = raw_df.copy()
    df["month"] = df["datetime"].dt.month
    df["date"]  = df["datetime"].dt.date
    slots = df["slot"].values.astype(int)

    if tariff == "A1 Flat":
        df["ic"] = df["consumption_kwh"] * A1_RATE
        sc = A1_SUPPLY
    else:
        df["ic"] = df["consumption_kwh"] * np.vectorize(ms_rate)(slots)
        sc = MS_SUPPLY

    costs = {}
    for season, months in SEASON_MONTHS.items():
        sub = df[df["month"].isin(months)]
        if sub.empty:
            costs[season] = None
            continue
        n_days = sub["date"].nunique()
        costs[season] = (sub["ic"].sum() + sc * n_days) / n_days if n_days else None
    return costs


# ─────────────────────────────────────────────────────────────────────────────
# Plot functions
# ─────────────────────────────────────────────────────────────────────────────

def make_seasonal_fig(res_a, res_b, cfg_a: dict, cfg_b: dict, raw_df: pd.DataFrame) -> plt.Figure:
    """2×4 grid: rows = options A/B, cols = seasons."""
    costs_a   = seasonal_daily_cost(res_a)
    costs_b   = seasonal_daily_cost(res_b)
    base_a    = baseline_seasonal_cost(raw_df, cfg_a["tariff"])
    base_b    = baseline_seasonal_cost(raw_df, cfg_b["tariff"])

    fig, axes = plt.subplots(2, 4, figsize=(18, 7), sharey=False)
    fig.patch.set_facecolor("white")
    seasons   = list(SEASON_MONTHS.keys())
    row_labels = [
        f"Option A: {cfg_a['label']}",
        f"Option B: {cfg_b['label']}",
    ]
    row_res    = [res_a, res_b]
    row_costs  = [costs_a, costs_b]
    row_base   = [base_a, base_b]
    row_cols   = OPTION_COLOURS

    for row, (res, costs, base, col, row_lbl) in enumerate(
            zip(row_res, row_costs, row_base, row_cols, row_labels)):
        df = res["df"].copy()
        df["month"] = df["datetime"].dt.month

        for c, (season, months) in enumerate(SEASON_MONTHS.items()):
            ax  = axes[row][c]
            sub = df[df["month"].isin(months)]
            if sub.empty:
                ax.set_visible(False)
                continue

            avg = sub.groupby("slot")[
                ["solar_kwh", "s_slf", "b_dis", "g_imp", "consumption_kwh"]
            ].mean()
            h = avg.index / 2

            ax.stackplot(
                h,
                avg["s_slf"], avg["b_dis"], avg["g_imp"],
                labels=["Solar self-use", "Battery", "Grid import"],
                colors=[CS["solar"], CS["bdis"], CS["gimp"]],
                alpha=0.85,
            )
            ax.plot(h, avg["consumption_kwh"], "k-", lw=1.2, label="Load")
            ax.plot(h, avg["solar_kwh"], "--", color=CS["solar"], lw=0.8,
                    alpha=0.65, label="Solar gen")
            ax.axvspan(15, 21, alpha=0.07, color="red")
            ax.axvspan(9,  15, alpha=0.06, color="green")
            ax.set_xlim(0, 24)
            ax.set_xticks(range(0, 25, 6))
            ax.set_xticklabels([f"{h:02d}h" for h in range(0, 25, 6)], fontsize=7)
            ax.set_ylabel("Avg kWh / 30 min", fontsize=7)
            ax.tick_params(axis="y", labelsize=7)

            # Cost annotation
            daily  = costs.get(season)
            b_cost = base.get(season)
            if daily is not None and b_cost is not None:
                saving = b_cost - daily
                sign   = "+" if saving >= 0 else ""
                ax.text(
                    0.98, 0.97,
                    f"Avg daily cost: ${daily:.2f}\n"
                    f"Saving vs baseline: {sign}${saving:.2f}/day",
                    transform=ax.transAxes,
                    ha="right", va="top", fontsize=7,
                    bbox=dict(boxstyle="round,pad=0.3", fc="white",
                              ec=col, alpha=0.85, lw=1.0),
                )

            # Column header on top row only
            if row == 0:
                ax.set_title(season, fontsize=9, fontweight="bold")
            else:
                ax.set_title("")

        # Row label as y-axis label on first column
        axes[row][0].set_ylabel(
            f"{row_lbl}\nAvg kWh / 30 min", fontsize=7.5, fontweight="bold"
        )

    # Shared legend
    handles, labels_ = axes[0][0].get_legend_handles_labels()
    fig.legend(
        handles, labels_,
        loc="lower center", ncol=5, fontsize=8,
        bbox_to_anchor=(0.5, -0.02),
    )
    fig.tight_layout(rect=[0, 0.04, 1, 1])
    return fig


def make_payback_fig(pb_a: dict, pb_b: dict, cfg_a: dict, cfg_b: dict) -> plt.Figure:
    """Side-by-side cumulative cashflow for the two options."""
    fig, ax = plt.subplots(figsize=(11, 4.5))
    fig.patch.set_facecolor("white")
    yrs = list(range(ANALYSIS_YEARS + 1))

    for pb, cfg, col in [(pb_a, cfg_a, OPTION_COLOURS[0]),
                          (pb_b, cfg_b, OPTION_COLOURS[1])]:
        lbl = f"Option {'A' if col == OPTION_COLOURS[0] else 'B'}: {cfg['label']} ({cfg['tariff']})"
        if pb["pb_yr"]:
            lbl += f"  ← payback yr {pb['pb_yr']}"
        ax.plot(yrs, pb["cum"], color=col, lw=2.0, label=lbl)

    ax.axhline(0, color="black", lw=0.8, ls=":")
    ax.fill_between(yrs, 0, color="green", alpha=0.04)
    ax.set_xlabel("Year")
    ax.set_ylabel("Cumulative Cash-Flow (AUD)")
    ax.set_title(
        f"Cumulative Cash-Flow — {ANALYSIS_YEARS}-Year Horizon  ·  "
        f"Solid = A1 Flat, Dashed = Midday Saver"
    )
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v:,.0f}"))
    ax.xaxis.set_major_locator(mticker.MultipleLocator(5))
    ax.legend(fontsize=8)
    ax.grid(axis="y", alpha=0.3)
    fig.tight_layout()
    return fig


def make_monthly_fig(res_a, res_b, raw_df, cfg_a, cfg_b) -> plt.Figure:
    """Monthly net electricity cost comparison."""
    m_a    = _monthly_net(res_a)
    m_b    = _monthly_net(res_b)
    months = range(1, 13)
    mnames = ["Jan","Feb","Mar","Apr","May","Jun",
              "Jul","Aug","Sep","Oct","Nov","Dec"]

    # Baseline — use tariff of option A for a single baseline line
    tariff_bl = cfg_a["tariff"]
    df_bl     = raw_df.copy()
    slots     = df_bl["slot"].values.astype(int)
    if tariff_bl == "A1 Flat":
        df_bl["ic"] = df_bl["consumption_kwh"] * A1_RATE
        sc_bl = A1_SUPPLY
    else:
        df_bl["ic"] = df_bl["consumption_kwh"] * np.vectorize(ms_rate)(slots)
        sc_bl = MS_SUPPLY
    df_bl["month"] = df_bl["datetime"].dt.month
    df_bl["date"]  = df_bl["datetime"].dt.date
    grp    = df_bl.groupby("month").agg(ic=("ic","sum"), days=("date","nunique"))
    bl_net = (grp["ic"] + grp["days"] * sc_bl).reindex(range(1,13), fill_value=0)

    fig, ax = plt.subplots(figsize=(11, 4))
    fig.patch.set_facecolor("white")
    x = np.arange(12)
    w = 0.35
    ax.bar(x - w/2, [m_a["net"].get(m, 0) for m in months], w,
           label=f"A: {cfg_a['label']}", color=OPTION_COLOURS[0], alpha=0.8)
    ax.bar(x + w/2, [m_b["net"].get(m, 0) for m in months], w,
           label=f"B: {cfg_b['label']}", color=OPTION_COLOURS[1], alpha=0.8)
    ax.step(x, [bl_net.get(m, 0) for m in months], where="mid",
            color="black", lw=1.3, ls="--", label="Baseline (no solar)")
    ax.set_xticks(x); ax.set_xticklabels(mnames)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v:,.0f}"))
    ax.set_ylabel("Monthly Net Cost (AUD)")
    ax.set_title("Monthly Electricity Cost — Option A vs B vs Baseline")
    ax.axhline(0, color="black", lw=0.5, ls=":")
    ax.legend(fontsize=8)
    ax.grid(axis="y", alpha=0.3)
    fig.tight_layout()
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# Sidebar — upload + configuration
# ─────────────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.title("☀️ Solar + Battery Analyser")
    st.caption("Synergy WA — Perth / SWIS")

    st.divider()
    st.subheader("1. Upload meter data")
    uploaded = st.file_uploader(
        "Half-hourly export from Synergy (.xlsx or .csv)",
        type=["xlsx", "xls", "csv", "tsv"],
    )

    raw_df = None
    if uploaded:
        with st.spinner("Loading data…"):
            try:
                raw_df = cached_load(uploaded.getvalue(), uploaded.name)
                n_days = raw_df["date"].nunique()
                date_min = pd.to_datetime(raw_df["datetime"].min()).strftime("%d %b %Y")
                date_max = pd.to_datetime(raw_df["datetime"].max()).strftime("%d %b %Y")
                st.success(f"{n_days} days loaded  ·  {date_min} → {date_max}")
            except Exception as e:
                st.error(f"Could not load file: {e}")
                raw_df = None

    st.divider()
    st.subheader("2. Configure options")

    TARIFFS = ["A1 Flat", "Midday Saver"]

    def option_form(tag: str, default_solar, default_bat, default_inv,
                    default_cost, default_tariff, default_label):
        with st.expander(f"Option {tag}", expanded=True):
            label  = st.text_input(f"Label {tag}", value=default_label, key=f"lbl_{tag}")
            solar  = st.number_input(f"Solar array (kW)", min_value=0.0, max_value=30.0,
                                     value=float(default_solar), step=0.5, key=f"sol_{tag}")
            bat    = st.number_input(f"Battery (kWh)", min_value=0.0, max_value=60.0,
                                     value=float(default_bat), step=0.5, key=f"bat_{tag}")
            inv    = st.number_input(f"Inverter (kW)", min_value=0.0, max_value=15.0,
                                     value=float(default_inv), step=0.5, key=f"inv_{tag}")
            cost   = st.number_input(f"Gross cost ($)", min_value=0, max_value=100_000,
                                     value=int(default_cost), step=500, key=f"cost_{tag}")
            tariff = st.selectbox(f"Tariff", TARIFFS,
                                  index=TARIFFS.index(default_tariff), key=f"tariff_{tag}")
        return dict(label=label, solar=solar, bat=bat, inv=inv, cost=cost, tariff=tariff)

    cfg_a = option_form("A",
        default_solar=12.3, default_bat=16.0, default_inv=10.0,
        default_cost=25_180, default_tariff="Midday Saver",
        default_label="12.3 kW + 16 kWh battery")
    cfg_b = option_form("B",
        default_solar=6.3, default_bat=26.3, default_inv=10.0,
        default_cost=25_280, default_tariff="Midday Saver",
        default_label="6.3 kW + 26.3 kWh battery")

    run = st.button("Run analysis", type="primary", disabled=(raw_df is None))


# ─────────────────────────────────────────────────────────────────────────────
# Main area
# ─────────────────────────────────────────────────────────────────────────────

st.title("☀️ Solar + Battery Comparison")

if raw_df is None:
    st.info("Upload your Synergy half-hourly meter data in the sidebar to get started.")
    st.stop()

if not run:
    st.info("Configure your two options in the sidebar, then click **Run analysis**.")
    st.stop()

# ── Run simulations ──────────────────────────────────────────────────────────
raw_hash = hash(uploaded.getvalue())

with st.spinner("Simulating Option A…"):
    res_a = cached_simulate(
        raw_hash, cfg_a["solar"], cfg_a["bat"], cfg_a["inv"], cfg_a["tariff"], raw_df
    )
with st.spinner("Simulating Option B…"):
    res_b = cached_simulate(
        raw_hash, cfg_b["solar"], cfg_b["bat"], cfg_b["inv"], cfg_b["tariff"], raw_df
    )
with st.spinner("Computing payback…"):
    pb_a = cached_payback(
        raw_hash, cfg_a["solar"], cfg_a["bat"], cfg_a["inv"],
        cfg_a["cost"], cfg_a["tariff"], cfg_a["label"], raw_df
    )
    pb_b = cached_payback(
        raw_hash, cfg_b["solar"], cfg_b["bat"], cfg_b["inv"],
        cfg_b["cost"], cfg_b["tariff"], cfg_b["label"], raw_df
    )

bl_a = baseline(add_solar(raw_df, cfg_a["solar"]), cfg_a["tariff"])
bl_b = baseline(add_solar(raw_df, cfg_b["solar"]), cfg_b["tariff"])

# ── Metric cards ─────────────────────────────────────────────────────────────
st.subheader("Key Metrics")
col_a, col_b = st.columns(2)

def metric_col(col, res, pb, cfg, bl, colour):
    annual_saving = bl - res["net_cost"]
    with col:
        st.markdown(
            f"<div style='border-left:4px solid {colour}; padding-left:10px'>"
            f"<b>Option {('A' if colour == OPTION_COLOURS[0] else 'B')}: {cfg['label']}</b>"
            f"</div>",
            unsafe_allow_html=True,
        )
        r1, r2, r3 = st.columns(3)
        r1.metric("Annual saving", f"${annual_saving:,.0f}")
        r2.metric("Nominal payback", f"{pb['pb_yr']} yr" if pb["pb_yr"] else ">25 yr")
        r3.metric("NPV @ 8%", f"${pb['npv']:,.0f}")

        r4, r5, r6 = st.columns(3)
        r4.metric("25-yr saving", f"${pb['total_save']:,.0f}")
        r5.metric("ROI", f"{pb['roi']:.0f}%")
        r6.metric("Self-sufficiency", f"{res['self_suf_pct']:.0f}%")

        r7, r8 = st.columns(2)
        r7.metric("Annual export", f"{res['annual_export_kwh']:,.0f} kWh")
        r8.metric("Net system cost", f"${pb['net']:,.0f}")

metric_col(col_a, res_a, pb_a, cfg_a, bl_a, OPTION_COLOURS[0])
metric_col(col_b, res_b, pb_b, cfg_b, bl_b, OPTION_COLOURS[1])

# ── Typical seasonal days ─────────────────────────────────────────────────────
st.divider()
st.subheader("Typical Day by Season")
st.caption(
    "Average half-hourly energy flows for each Australian season. "
    "Cost box shows average daily net electricity cost and saving vs the no-solar baseline. "
    "Green shading = super off-peak window (9am–3pm), red = peak (3–9pm)."
)

with st.spinner("Generating seasonal charts…"):
    fig_seasonal = make_seasonal_fig(res_a, res_b, cfg_a, cfg_b, raw_df)
st.pyplot(fig_seasonal, use_container_width=True)
plt.close(fig_seasonal)

# ── Payback / cashflow ────────────────────────────────────────────────────────
st.divider()
st.subheader(f"{ANALYSIS_YEARS}-Year Cumulative Cash-Flow")

with st.spinner("Generating payback chart…"):
    fig_pb = make_payback_fig(pb_a, pb_b, cfg_a, cfg_b)
st.pyplot(fig_pb, use_container_width=True)
plt.close(fig_pb)

# Payback detail table
pb_rows = []
for pb, cfg, tag in [(pb_a, cfg_a, "A"), (pb_b, cfg_b, "B")]:
    pb_rows.append({
        "Option": f"{tag}: {cfg['label']}",
        "Tariff": cfg["tariff"],
        "Gross cost": f"${cfg['cost']:,.0f}",
        "Net cost (after rebates)": f"${pb['net']:,.0f}",
        "Nominal payback": f"{pb['pb_yr']} yr" if pb["pb_yr"] else ">25 yr",
        f"Opp-cost payback ({OPPORTUNITY_RATE*100:.0f}%)":
            f"{pb['pb_yr_disc']} yr" if pb.get("pb_yr_disc") else ">25 yr",
        f"NPV @ {OPPORTUNITY_RATE*100:.0f}%": f"${pb['npv']:,.0f}",
        "25-yr saving": f"${pb['total_save']:,.0f}",
        "ROI": f"{pb['roi']:.0f}%",
        "Year-1 saving": f"${pb['yr1_save']:,.0f}",
    })
st.dataframe(pd.DataFrame(pb_rows).set_index("Option"), use_container_width=True)

# ── Monthly cost ──────────────────────────────────────────────────────────────
st.divider()
st.subheader("Monthly Electricity Cost")

with st.spinner("Generating monthly chart…"):
    fig_monthly = make_monthly_fig(res_a, res_b, raw_df, cfg_a, cfg_b)
st.pyplot(fig_monthly, use_container_width=True)
plt.close(fig_monthly)

st.caption(
    f"Tariff escalation {TARIFF_ESC*100:.1f}%/yr and DEBS decline {DEBS_DECL*100:.0f}%/yr "
    f"not reflected in the monthly chart (year-0 rates shown). "
    f"Payback figures above do account for these trends over {ANALYSIS_YEARS} years."
)

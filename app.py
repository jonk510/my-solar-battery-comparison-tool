"""
Solar & Battery Comparison App — Synergy WA
Upload your half-hourly meter data, configure two options, compare side-by-side.
Run with:  streamlit run app.py
"""
import sys, os
import io
import re
import warnings
import urllib.request
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
    CS, _QUOTES_FILE, SOLAR_K, SYS_EFF as _SYS_EFF, _lat,
    solar_stc_rebate, STC_PRICE,
)

_QUOTES_PATH = _QUOTES_FILE


@st.cache_data(ttl=43200, show_spinner=False)   # refresh every 12 hours
def fetch_stc_price() -> tuple[float, str]:
    """Try to scrape the current STC spot price from Ecovantage.
    Returns (price_inc_gst, source_label).
    Falls back to the hardcoded STC_PRICE constant on any failure.
    """
    try:
        url = "https://www.ecovantage.com.au/energy-certificate-market-update/"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=5) as resp:
            html = resp.read().decode("utf-8", errors="ignore")
        # Look for a dollar amount near "STC" within a short window of text
        # Pattern: STC ... $XX.XX  or  $XX.XX ... STC
        m = re.search(
            r'STC[^$]{0,120}\$\s*(\d{1,2}(?:\.\d{1,2})?)',
            html, re.IGNORECASE
        )
        if not m:
            m = re.search(
                r'\$\s*(\d{1,2}(?:\.\d{1,2})?)[^$]{0,120}STC',
                html, re.IGNORECASE
            )
        if m:
            price = float(m.group(1))
            if 15.0 <= price <= 45.0:   # sanity bounds
                return price, "Ecovantage (live)"
    except Exception:
        pass
    return float(STC_PRICE), "default (no live data)"


def add_solar_shaded(raw_df: pd.DataFrame, solar_kw: float,
                     shade_summer: float, shade_autumn: float,
                     shade_winter: float) -> pd.DataFrame:
    """Like add_solar() but uses caller-supplied seasonal shading factors."""
    def _profile(kw, doy):
        decl  = np.radians(-23.45 * np.cos(2 * np.pi * (doy + 10) / 365))
        slots = np.arange(48)
        hours = (slots + 0.5) / 2.0
        ha    = np.radians(15.0 * (hours - 12.0))
        sin_e = np.sin(_lat) * np.sin(decl) + np.cos(_lat) * np.cos(decl) * np.cos(ha)
        irrad = np.maximum(0.0, sin_e)
        m = (pd.Timestamp("2024-01-01") + pd.Timedelta(days=int(doy) - 1)).month
        if m in (12, 1, 2):    sf = shade_summer
        elif m in (6, 7, 8):   sf = shade_winter
        else:                  sf = shade_autumn   # Mar–May and Sep–Nov
        return kw * irrad * SOLAR_K * _SYS_EFF * 0.5 * sf

    cache = {}
    def get(row):
        doy = int(row["doy"])
        if doy not in cache:
            cache[doy] = _profile(solar_kw, doy)
        return cache[doy][int(row["slot"])]

    df = raw_df.copy()
    df["solar_kwh"] = df.apply(get, axis=1)
    return df

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
def cached_simulate(raw_df_hash, solar_kw, bat_kwh, inv_kw, tariff,
                    shade_summer, shade_autumn, shade_winter, _raw_df):
    df_solar = add_solar_shaded(_raw_df, solar_kw, shade_summer, shade_autumn, shade_winter)
    return simulate(df_solar, solar_kw, bat_kwh, inv_kw, tariff)


@st.cache_data(show_spinner=False)
def cached_payback(raw_df_hash, solar_kw, bat_kwh, inv_kw, cost, tariff, label,
                   shade_summer, shade_autumn, shade_winter, stc_price, _raw_df):
    df = add_solar_shaded(_raw_df, solar_kw, shade_summer, shade_autumn, shade_winter)
    return payback(df, solar_kw, bat_kwh, inv_kw, cost, tariff, label, stc_price=stc_price)


@st.cache_data(show_spinner=False)
def cached_simulate_no_solar(raw_df_hash, tariff, _raw_df):
    df = _raw_df.copy()
    df["solar_kwh"] = 0.0
    return simulate(df, 0.0, 0.0, 5.0, tariff)


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

def make_seasonal_fig(res_base, res_a, res_b, cfg_base: dict, cfg_a: dict, cfg_b: dict) -> plt.Figure:
    """4×3 grid: rows = seasons, cols = [Status Quo | Option A | Option B]."""
    costs_base = seasonal_daily_cost(res_base)
    costs_a    = seasonal_daily_cost(res_a)
    costs_b    = seasonal_daily_cost(res_b)

    col_configs = [
        (res_base, costs_base, "gray",            "Status Quo\n(no solar/battery)"),
        (res_a,    costs_a,    OPTION_COLOURS[0], f"Option A: {cfg_a['label']}"),
        (res_b,    costs_b,    OPTION_COLOURS[1], f"Option B: {cfg_b['label']}"),
    ]

    fig, axes = plt.subplots(4, 3, figsize=(18, 18), sharey="row")
    fig.patch.set_facecolor("white")

    for col_idx, (res, costs, col, col_lbl) in enumerate(col_configs):
        df = res["df"].copy()
        df["month"] = df["datetime"].dt.month

        for row_idx, (season, months) in enumerate(SEASON_MONTHS.items()):
            ax  = axes[row_idx][col_idx]
            sub = df[df["month"].isin(months)]
            if sub.empty:
                ax.set_visible(False)
                continue

            avg = sub.groupby("slot")[
                ["solar_kwh", "s_slf", "b_dis", "g_imp", "g_exp", "consumption_kwh"]
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
            # Grid export — fill above load line when solar exceeds consumption
            export_top = avg["consumption_kwh"] + avg["g_exp"]
            ax.fill_between(h, avg["consumption_kwh"], export_top,
                            color=CS.get("gexp", "#4fc3f7"), alpha=0.5,
                            label="Grid export")
            ax.axvspan(15, 21, alpha=0.07, color="red")
            ax.axvspan(9,  15, alpha=0.06, color="green")
            ax.set_xlim(0, 24)
            ax.set_xticks(range(0, 25, 6))
            ax.set_xticklabels([f"{h:02d}h" for h in range(0, 25, 6)], fontsize=7)
            ax.tick_params(axis="y", labelsize=7)

            # Cost annotation
            daily = costs.get(season)
            if daily is not None:
                if col_idx == 0:
                    annotation = f"Avg daily cost: ${daily:.2f}"
                else:
                    base_d = costs_base.get(season)
                    if base_d is not None:
                        saving = base_d - daily
                        sign   = "+" if saving >= 0 else ""
                        annotation = (f"Avg daily cost: ${daily:.2f}\n"
                                      f"vs status quo: {sign}${saving:.2f}/day")
                    else:
                        annotation = f"Avg daily cost: ${daily:.2f}"
                ax.text(
                    0.98, 0.97, annotation,
                    transform=ax.transAxes,
                    ha="right", va="top", fontsize=7,
                    bbox=dict(boxstyle="round,pad=0.3", fc="white",
                              ec=col, alpha=0.85, lw=1.0),
                )

            # Season name on left column; short y-label on others
            if col_idx == 0:
                ax.set_ylabel(f"{season}\nAvg kWh / 30 min", fontsize=7.5, fontweight="bold")
            else:
                ax.set_ylabel("Avg kWh / 30 min", fontsize=7)

            # Column header on top row only
            if row_idx == 0:
                ax.set_title(col_lbl, fontsize=9, fontweight="bold", color=col)

    # Shared legend
    handles, labels_ = axes[0][0].get_legend_handles_labels()
    fig.legend(
        handles, labels_,
        loc="lower center", ncol=5, fontsize=8,
        bbox_to_anchor=(0.5, -0.01),
    )
    fig.tight_layout(rect=[0, 0.03, 1, 1])
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


def make_monthly_fig(res_base, res_a, res_b, cfg_base: dict, cfg_a: dict, cfg_b: dict) -> plt.Figure:
    """Monthly net electricity cost — Status Quo vs Option A vs Option B."""
    m_base = _monthly_net(res_base)
    m_a    = _monthly_net(res_a)
    m_b    = _monthly_net(res_b)
    months = range(1, 13)
    mnames = ["Jan","Feb","Mar","Apr","May","Jun",
              "Jul","Aug","Sep","Oct","Nov","Dec"]

    fig, ax = plt.subplots(figsize=(13, 4))
    fig.patch.set_facecolor("white")
    x = np.arange(12)
    w = 0.25
    ax.bar(x - w, [m_base["net"].get(m, 0) for m in months], w,
           label="Status Quo (no solar/battery)", color="gray", alpha=0.75)
    ax.bar(x,     [m_a["net"].get(m, 0) for m in months], w,
           label=f"Option A: {cfg_a['label']}", color=OPTION_COLOURS[0], alpha=0.8)
    ax.bar(x + w, [m_b["net"].get(m, 0) for m in months], w,
           label=f"Option B: {cfg_b['label']}", color=OPTION_COLOURS[1], alpha=0.8)
    ax.set_xticks(x); ax.set_xticklabels(mnames)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v:,.0f}"))
    ax.set_ylabel("Monthly Net Cost (AUD)")
    ax.set_title("Monthly Electricity Cost — Status Quo vs Option A vs Option B")
    ax.axhline(0, color="black", lw=0.5, ls=":")
    ax.legend(fontsize=8)
    ax.grid(axis="y", alpha=0.3)
    fig.tight_layout()
    return fig


def make_solar_profile_fig(cfg_a: dict, cfg_b: dict,
                            shade_summer: float, shade_autumn: float,
                            shade_winter: float) -> plt.Figure:
    """2×2 grid — one subplot per season, each with 3 curves:
      · Option A unshaded  (shade = 1.0)
      · Option A shaded    (user shading factors)
      · Option B shaded    (user shading factors)
    x-axis = hour of day, y-axis = kWh per 30-min slot.
    """
    # Representative mid-season day-of-year (Southern Hemisphere)
    SEASONS = [
        ("Summer (Dec–Feb)", 15,  shade_summer),   # ~Jan 15
        ("Autumn (Mar–May)", 105, shade_autumn),   # ~Apr 15
        ("Winter (Jun–Aug)", 196, shade_winter),   # ~Jul 15
        ("Spring (Sep–Nov)", 288, shade_autumn),   # ~Oct 15
    ]

    def _day_profile(solar_kw: float, doy: int, shade: float) -> np.ndarray:
        decl  = np.radians(-23.45 * np.cos(2 * np.pi * (doy + 10) / 365))
        slots = np.arange(48)
        hours = (slots + 0.5) / 2.0
        ha    = np.radians(15.0 * (hours - 12.0))
        sin_e = (np.sin(_lat) * np.sin(decl)
                 + np.cos(_lat) * np.cos(decl) * np.cos(ha))
        irrad = np.maximum(0.0, sin_e)
        return solar_kw * irrad * SOLAR_K * _SYS_EFF * 0.5 * shade

    fig, axes = plt.subplots(2, 2, figsize=(12, 7), sharey=False)
    fig.patch.set_facecolor("white")
    hours = (np.arange(48) + 0.5) / 2.0

    for ax, (season, doy, shade) in zip(axes.flat, SEASONS):
        kw_a = cfg_a["solar"]
        kw_b = cfg_b["solar"]

        unshaded_a = _day_profile(kw_a, doy, 1.0)
        shaded_a   = _day_profile(kw_a, doy, shade)
        shaded_b   = _day_profile(kw_b, doy, shade)

        ax.plot(hours, unshaded_a, color=OPTION_COLOURS[0], ls="--", lw=1.4, alpha=0.6,
                label=f"A unshaded ({kw_a} kW)")
        ax.plot(hours, shaded_a,   color=OPTION_COLOURS[0], ls="-",  lw=2.0,
                label=f"A shaded  ({kw_a} kW, ×{shade:.2f})")
        ax.plot(hours, shaded_b,   color=OPTION_COLOURS[1], ls="-",  lw=2.0,
                label=f"B shaded  ({kw_b} kW, ×{shade:.2f})")

        ax.fill_between(hours, shaded_a, unshaded_a,
                        color=OPTION_COLOURS[0], alpha=0.08, label="_")

        ax.set_title(season, fontsize=10, fontweight="bold")
        ax.set_xlim(5, 20)
        ax.set_xticks(range(6, 21, 2))
        ax.set_xticklabels([f"{h:02d}:00" for h in range(6, 21, 2)], fontsize=8)
        ax.set_xlabel("Hour of day", fontsize=8)
        ax.set_ylabel("kWh / 30 min", fontsize=8)
        ax.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.2f"))
        ax.tick_params(labelsize=8)
        ax.legend(fontsize=7, loc="upper left")
        ax.grid(axis="y", alpha=0.25)

        # Shading loss annotation
        peak_loss_pct = (1 - shade) * 100
        ax.text(0.98, 0.97, f"Shading loss: {peak_loss_pct:.0f}%",
                transform=ax.transAxes, ha="right", va="top", fontsize=8,
                bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="#aaaaaa",
                          alpha=0.85, lw=0.8))

    fig.suptitle("Solar Generation Profiles by Season — Shaded vs Unshaded",
                 fontsize=11, fontweight="bold", y=1.01)
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
    st.subheader("2. Select quotes")

    # ── Load quotes from xlsx (flexible column matching) ─────────────────────
    quotes_df = None
    if _QUOTES_PATH.exists():
        try:
            quotes_df = pd.read_excel(_QUOTES_PATH)
            quotes_df.columns = quotes_df.columns.str.strip()
            cols = {c.lower(): c for c in quotes_df.columns}

            def _fc(keywords):
                for kw in keywords:
                    for lc, orig in cols.items():
                        if kw in lc:
                            return orig
                raise ValueError(f"No column matching {keywords} in {list(quotes_df.columns)}")

            col_vendor   = _fc(["vendor", "name", "label"])
            col_solar    = _fc(["solar"])
            col_battery  = _fc(["battery", "bat"])
            col_inverter = _fc(["inverter", "inv"])
            col_cost     = _fc(["price", "cost", "aud"])
            col_order    = next(
                (cols[lc] for lc in cols if any(k in lc for k in ["quote", "number", "order"])),
                None,
            )

            if col_order:
                quotes_df = quotes_df.sort_values(col_order)

            # Normalise to consistent internal column names
            quotes_df = quotes_df.rename(columns={
                col_vendor:   "Vendor",
                col_solar:    "Solar_kW",
                col_battery:  "Battery_kWh",
                col_inverter: "Inverter_kW",
                col_cost:     "Cost_AUD",
            }).reset_index(drop=True)

            # Filter out base-case rows (solar=0 and battery=0)
            solar_quotes = quotes_df[
                ~((quotes_df["Solar_kW"] == 0) & (quotes_df["Battery_kWh"] == 0))
            ].reset_index(drop=True)
        except Exception as e:
            st.warning(f"Could not load quotes file: {e}")
            solar_quotes = None
    else:
        st.warning(
            "`solar_battery_quotes.xlsx` not found in the app folder. "
            "Add it to the repository to enable quote selection."
        )
        solar_quotes = None

    TARIFFS = ["Midday Saver", "A1 Flat"]

    def option_selector(tag: str, default_idx: int):
        colour = "#e8463a" if tag == "A" else "#0f9d58"
        st.markdown(
            f"<span style='border-left:4px solid {colour};"
            f"padding-left:6px;font-weight:bold'>Option {tag}</span>",
            unsafe_allow_html=True,
        )
        if solar_quotes is not None and len(solar_quotes) > 0:
            def _fmt(i):
                r = solar_quotes.iloc[i]
                return (f"{r['Vendor']}  —  {r['Solar_kW']:.2g} kW solar / "
                        f"{r['Battery_kWh']:.4g} kWh battery  (${int(r['Cost_AUD']):,})")
            sel_idx = st.selectbox(
                "Quote", range(len(solar_quotes)),
                format_func=_fmt,
                index=min(default_idx, len(solar_quotes) - 1),
                key=f"quote_{tag}",
            )
            row   = solar_quotes.iloc[sel_idx]
            solar = float(row["Solar_kW"])
            bat   = float(row["Battery_kWh"])
            inv   = float(row["Inverter_kW"])
            cost  = int(row["Cost_AUD"])
            label = _fmt(sel_idx)
        else:
            # Fallback to manual entry if no quotes file
            label = st.text_input("Label", key=f"lbl_{tag}",
                                  value="Option A" if tag == "A" else "Option B")
            solar = st.number_input("Solar (kW)", 0.0, 30.0, 6.6, 0.5, key=f"sol_{tag}")
            bat   = st.number_input("Battery (kWh)", 0.0, 60.0, 16.0, 0.5, key=f"bat_{tag}")
            inv   = st.number_input("Inverter (kW)", 0.0, 15.0, 10.0, 0.5, key=f"inv_{tag}")
            cost  = st.number_input("Gross cost ($)", 0, 100_000, 20_000, 500, key=f"cost_{tag}")

        tariff = st.selectbox("Tariff", TARIFFS, key=f"tariff_{tag}")
        st.markdown("")
        return dict(label=label, solar=solar, bat=bat, inv=inv, cost=cost, tariff=tariff)

    cfg_a = option_selector("A", default_idx=0)
    cfg_b = option_selector("B", default_idx=1)

    st.divider()
    st.subheader("3. Site shading")
    st.caption("Fraction of unshaded solar reaching the panels each season.")
    shade_summer  = st.slider("Summer shading (Dec–Feb)",  0.0, 1.0, 0.60, 0.05)
    shade_autumn  = st.slider("Autumn/Spring shading",     0.0, 1.0, 0.50, 0.05)
    shade_winter  = st.slider("Winter shading (Jun–Aug)",  0.0, 1.0, 0.40, 0.05)

    st.divider()
    st.subheader("4. STC price")
    _live_stc, _stc_source = fetch_stc_price()
    st.caption(
        f"Small-scale Technology Certificate spot price (inc GST). "
        f"Default loaded from: **{_stc_source}**. "
        f"Clearing-house max is $40 ex-GST (~$44 inc GST)."
    )
    stc_price = st.slider("STC spot price ($/STC)", 20.0, 45.0, _live_stc, 0.50)

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

shading = (shade_summer, shade_autumn, shade_winter)

with st.spinner("Simulating Option A…"):
    res_a = cached_simulate(
        raw_hash, cfg_a["solar"], cfg_a["bat"], cfg_a["inv"], cfg_a["tariff"],
        *shading, raw_df,
    )
with st.spinner("Simulating Option B…"):
    res_b = cached_simulate(
        raw_hash, cfg_b["solar"], cfg_b["bat"], cfg_b["inv"], cfg_b["tariff"],
        *shading, raw_df,
    )
with st.spinner("Computing payback…"):
    pb_a = cached_payback(
        raw_hash, cfg_a["solar"], cfg_a["bat"], cfg_a["inv"],
        cfg_a["cost"], cfg_a["tariff"], cfg_a["label"],
        *shading, stc_price, raw_df,
    )
    pb_b = cached_payback(
        raw_hash, cfg_b["solar"], cfg_b["bat"], cfg_b["inv"],
        cfg_b["cost"], cfg_b["tariff"], cfg_b["label"],
        *shading, stc_price, raw_df,
    )

with st.spinner("Computing status quo…"):
    res_base = cached_simulate_no_solar(raw_hash, cfg_a["tariff"], raw_df)
cfg_base = {"label": "Status Quo", "tariff": cfg_a["tariff"]}

bl_a = baseline(add_solar_shaded(raw_df, cfg_a["solar"], *shading), cfg_a["tariff"])
bl_b = baseline(add_solar_shaded(raw_df, cfg_b["solar"], *shading), cfg_b["tariff"])

# ── Metric cards ─────────────────────────────────────────────────────────────
st.subheader("Key Metrics")
col_base_m, col_a_m, col_b_m = st.columns(3)

with col_base_m:
    st.markdown(
        "<div style='border-left:4px solid gray; padding-left:10px'>"
        "<b>Status Quo (no solar/battery)</b>"
        "</div>",
        unsafe_allow_html=True,
    )
    base_annual = res_base["net_cost"]
    r1, r2 = st.columns(2)
    r1.metric("Annual cost", f"${base_annual:,.0f}")
    r2.metric("Monthly avg", f"${base_annual/12:,.0f}")
    r3, r4 = st.columns(2)
    r3.metric("Self-sufficiency", "0%")
    r4.metric("Grid reliance", "100%")

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
        r2.metric("Nominal payback", f"{pb['pb_yr']} yr" if pb["pb_yr"] else ">20 yr")
        r3.metric("NPV @ 8%", f"${pb['npv']:,.0f}")

        r4, r5, r6 = st.columns(3)
        r4.metric("20-yr saving", f"${pb['total_save']:,.0f}")
        r5.metric("ROI", f"{pb['roi']:.0f}%")
        r6.metric("Self-sufficiency", f"{res['self_suf_pct']:.0f}%")

        r7, r8 = st.columns(2)
        r7.metric("Annual export", f"{res['annual_export_kwh']:,.0f} kWh")
        r8.metric("Net system cost", f"${pb['net']:,.0f}")

metric_col(col_a_m, res_a, pb_a, cfg_a, bl_a, OPTION_COLOURS[0])
metric_col(col_b_m, res_b, pb_b, cfg_b, bl_b, OPTION_COLOURS[1])

# ── Solar generation profiles ────────────────────────────────────────────────
st.divider()
st.subheader("Solar Generation Profiles by Season")
st.caption(
    "Theoretical mid-season output for each option (kWh per 30-min slot). "
    "Dashed = Option A unshaded (shade factor 1.0). "
    "Solid = with your shading factors applied. "
    "Shaded area shows the generation lost to shading."
)
fig_profiles = make_solar_profile_fig(cfg_a, cfg_b, shade_summer, shade_autumn, shade_winter)
st.pyplot(fig_profiles, use_container_width=True)
plt.close(fig_profiles)

# ── Typical seasonal days ─────────────────────────────────────────────────────
st.divider()
st.subheader("Typical Day by Season")
st.caption(
    "Average half-hourly energy flows for each Australian season. "
    "Cost box shows average daily net electricity cost and saving vs the no-solar baseline. "
    "Green shading = super off-peak window (9am–3pm), red = peak (3–9pm)."
)

with st.spinner("Generating seasonal charts…"):
    fig_seasonal = make_seasonal_fig(res_base, res_a, res_b, cfg_base, cfg_a, cfg_b)
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
        "Retail price (gross)": f"${cfg['cost']:,.0f}",
        "  – STC solar rebate": f"-${pb.get('stc', 0):,.0f}",
        "  – WA battery rebate": f"-${pb.get('state', 0):,.0f}",
        "  – Federal battery rebate": f"-${pb.get('fed', 0):,.0f}",
        "Purchase price (net)": f"${pb['net']:,.0f}",
        "Nominal payback": f"{pb['pb_yr']} yr" if pb["pb_yr"] else ">20 yr",
        f"Opp-cost payback ({OPPORTUNITY_RATE*100:.0f}%)":
            f"{pb['pb_yr_disc']} yr" if pb.get("pb_yr_disc") else ">20 yr",
        f"NPV @ {OPPORTUNITY_RATE*100:.0f}%": f"${pb['npv']:,.0f}",
        "20-yr saving": f"${pb['total_save']:,.0f}",
        "ROI": f"{pb['roi']:.0f}%",
        "Year-1 saving": f"${pb['yr1_save']:,.0f}",
    })
st.dataframe(pd.DataFrame(pb_rows).set_index("Option"), use_container_width=True)

# ── Monthly cost ──────────────────────────────────────────────────────────────
st.divider()
st.subheader("Monthly Electricity Cost")

with st.spinner("Generating monthly chart…"):
    fig_monthly = make_monthly_fig(res_base, res_a, res_b, cfg_base, cfg_a, cfg_b)
st.pyplot(fig_monthly, use_container_width=True)
plt.close(fig_monthly)

st.caption(
    f"Tariff escalation {TARIFF_ESC*100:.1f}%/yr and DEBS decline {DEBS_DECL*100:.0f}%/yr "
    f"not reflected in the monthly chart (year-0 rates shown). "
    f"Payback figures above do account for these trends over {ANALYSIS_YEARS} years."
)

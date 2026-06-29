"""
Solar & Battery Comparison App — Synergy WA
Upload your half-hourly meter data, configure two options, compare side-by-side.
Run with:  streamlit run app.py
"""
import sys, os
import io
from pathlib import Path
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
    compute_irr,
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

# Make the shared library importable when running locally (not pip-installed).
try:
    import shared as _shared_pkg  # noqa: F401
except ModuleNotFoundError:
    import os as _os, sys as _sys
    _sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))
from shared.style import apply_theme, page_header
apply_theme()
page_header("Solar + Battery Analyser", "Financial & energy comparison across tariff options and battery sizes")

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
                    shade_summer, shade_autumn, shade_winter,
                    grid_charge, gc_start, gc_end, tariff_esc, _raw_df):
    df_solar = add_solar_shaded(_raw_df, solar_kw, shade_summer, shade_autumn, shade_winter)
    return simulate(df_solar, solar_kw, bat_kwh, inv_kw, tariff,
                    grid_charge=grid_charge, grid_charge_start=gc_start,
                    grid_charge_end=gc_end, tariff_esc=tariff_esc)


@st.cache_data(show_spinner=False)
def cached_payback(raw_df_hash, solar_kw, bat_kwh, inv_kw, cost, tariff, label,
                   shade_summer, shade_autumn, shade_winter, stc_price, rebates_included,
                   grid_charge, gc_start, gc_end, tariff_esc, discount_rate,
                   _raw_df):
    df = add_solar_shaded(_raw_df, solar_kw, shade_summer, shade_autumn, shade_winter)
    return payback(df, solar_kw, bat_kwh, inv_kw, cost, tariff, label,
                   stc_price=stc_price, rebates_included=rebates_included,
                   grid_charge=grid_charge, grid_charge_start=gc_start,
                   grid_charge_end=gc_end, tariff_esc=tariff_esc,
                   discount_rate=discount_rate)


@st.cache_data(show_spinner=False)
def cached_build_pdf(
    raw_df_hash,
    solar_a, bat_a, inv_a, tariff_a,
    solar_b, bat_b, inv_b, tariff_b,
    shade_s_a, shade_au_a, shade_w_a,
    shade_s_b, shade_au_b, shade_w_b,
    tariff_esc, discount_rate,
    _raw_df, _res_a, _res_b, _res_base, _pb_a, _pb_b, _cfg_a, _cfg_b,
):
    return _build_pdf(
        _raw_df, _res_a, _res_b, _res_base, _pb_a, _pb_b,
        _cfg_a, _cfg_b,
        (shade_s_a, shade_au_a, shade_w_a), (shade_s_b, shade_au_b, shade_w_b),
        tariff_esc, discount_rate,
    )


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

    fig, axes = plt.subplots(4, 3, figsize=(18, 18), sharey=True)
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


def make_total_spend_fig(pb_a: dict, pb_b: dict, cfg_a: dict, cfg_b: dict) -> plt.Figure:
    """Cumulative total spend: upfront cost + ongoing bills for each option vs base case."""
    fig, ax = plt.subplots(figsize=(11, 4.5))
    fig.patch.set_facecolor("white")
    yrs = list(range(ANALYSIS_YEARS + 1))

    # Base case: no solar — accumulate annual bills
    base_spend = [0]
    for yr in range(1, ANALYSIS_YEARS + 1):
        base_yr = pb_a["savings"][yr - 1] + pb_a["net_costs"][yr - 1]
        base_spend.append(base_spend[-1] + base_yr)
    ax.plot(yrs, base_spend, color="gray", lw=1.5, ls="--", label="No solar/battery (bills only)")

    for pb, cfg, col in [(pb_a, cfg_a, OPTION_COLOURS[0]),
                          (pb_b, cfg_b, OPTION_COLOURS[1])]:
        spend = [pb["net"]]
        for yr in range(1, ANALYSIS_YEARS + 1):
            spend.append(spend[-1] + pb["net_costs"][yr - 1])
        lbl = f"Option {'A' if col == OPTION_COLOURS[0] else 'B'}: {cfg['label']} ({cfg['tariff']})"
        ax.plot(yrs, spend, color=col, lw=2.0, label=lbl)

    ax.set_xlabel("Year")
    ax.set_ylabel("Cumulative Total Spend (AUD)")
    ax.set_title(f"Cumulative Total Spend — {ANALYSIS_YEARS}-Year Horizon")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v:,.0f}"))
    ax.xaxis.set_major_locator(mticker.MultipleLocator(5))
    ax.legend(fontsize=8)
    ax.grid(axis="y", alpha=0.3)
    fig.tight_layout()
    return fig


def make_pv_spend_fig(pb_a: dict, pb_b: dict, cfg_a: dict, cfg_b: dict) -> plt.Figure:
    """Total spend discounted to present value at the user's chosen discount rate."""
    dr = pb_a.get("discount_rate", OPPORTUNITY_RATE)
    fig, ax = plt.subplots(figsize=(11, 4.5))
    fig.patch.set_facecolor("white")
    yrs = list(range(ANALYSIS_YEARS + 1))

    # Base case: PV of cumulative bills (bills already escalate via savings+net_costs)
    base_pv = [0]
    for yr in range(1, ANALYSIS_YEARS + 1):
        base_yr = pb_a["savings"][yr - 1] + pb_a["net_costs"][yr - 1]
        base_pv.append(base_pv[-1] + base_yr / (1 + dr) ** yr)
    ax.plot(yrs, base_pv, color="gray", lw=1.5, ls="--", label="No solar/battery (bills only)")

    for pb, cfg, col in [(pb_a, cfg_a, OPTION_COLOURS[0]),
                          (pb_b, cfg_b, OPTION_COLOURS[1])]:
        pv_spend = [pb["net"]]  # upfront cost is today — no discounting
        for yr in range(1, ANALYSIS_YEARS + 1):
            pv_spend.append(pv_spend[-1] + pb["net_costs"][yr - 1] / (1 + dr) ** yr)
        lbl = f"Option {'A' if col == OPTION_COLOURS[0] else 'B'}: {cfg['label']} ({cfg['tariff']})"
        ax.plot(yrs, pv_spend, color=col, lw=2.0, label=lbl)

    ax.set_xlabel("Year")
    ax.set_ylabel("Cumulative Spend — Present Value (AUD)")
    ax.set_title(
        f"Cumulative Total Spend (PV-Adjusted @ {dr*100:.1f}% discount) "
        f"— {ANALYSIS_YEARS}-Year Horizon"
    )
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v:,.0f}"))
    ax.xaxis.set_major_locator(mticker.MultipleLocator(5))
    ax.legend(fontsize=8)
    ax.grid(axis="y", alpha=0.3)
    fig.tight_layout()
    return fig


def make_payback_fig(pb_a: dict, pb_b: dict, cfg_a: dict, cfg_b: dict) -> plt.Figure:
    """Side-by-side cumulative cashflow for the two options."""
    fig, ax = plt.subplots(figsize=(11, 4.5))
    fig.patch.set_facecolor("white")
    yrs = list(range(ANALYSIS_YEARS + 1))

    # Base case: no solar/battery — cumulative bills starting from $0
    base_cum = [0]
    for yr in range(1, ANALYSIS_YEARS + 1):
        base_yr = pb_a["savings"][yr - 1] + pb_a["net_costs"][yr - 1]
        base_cum.append(base_cum[-1] - base_yr)
    ax.plot(yrs, base_cum, color="gray", lw=1.5, ls="--", label="No solar/battery (cumulative bills)")

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


def make_npv_fig(pb_a: dict, pb_b: dict, cfg_a: dict, cfg_b: dict) -> plt.Figure:
    """Side-by-side DCF waterfall: year-0 investment bar then PV of annual savings."""
    dr  = pb_a.get("discount_rate", OPPORTUNITY_RATE)
    fig, axes = plt.subplots(1, 2, figsize=(14, 5))
    fig.patch.set_facecolor("white")

    for ax, pb, cfg, col, tag in [
        (axes[0], pb_a, cfg_a, OPTION_COLOURS[0], "A"),
        (axes[1], pb_b, cfg_b, OPTION_COLOURS[1], "B"),
    ]:
        yrs = list(range(ANALYSIS_YEARS + 1))
        pv_flows = [-pb["net"]] + [
            sav / (1.0 + dr) ** yr
            for yr, sav in enumerate(pb["savings"], 1)
        ]
        cum_npv, running = [], 0.0
        for v in pv_flows:
            running += v
            cum_npv.append(running)

        bar_cols = ["#b71c1c" if v < 0 else col for v in pv_flows]
        ax.bar(yrs, pv_flows, color=bar_cols, alpha=0.72, width=0.7, zorder=2)
        ax.plot(yrs, cum_npv, color=col, lw=2.0, marker=".", markersize=3, zorder=5,
                label="Cumulative NPV")
        ax.axhline(0, color="black", lw=0.8, ls=":")
        if pb.get("pb_yr_disc"):
            ax.axvline(pb["pb_yr_disc"], color="gray", lw=1.0, ls="--",
                       label=f"Disc. payback yr {pb['pb_yr_disc']}")

        irr = compute_irr(pb)
        irr_str = f"{irr*100:.1f}%" if irr is not None else "N/A"
        ax.text(0.97, 0.04,
                f"NPV: ${pb['npv']:,.0f}\nIRR:  {irr_str}",
                transform=ax.transAxes, ha="right", va="bottom", fontsize=9,
                bbox=dict(boxstyle="round,pad=0.4", fc="white", ec=col, alpha=0.9, lw=1.2))

        ax.set_xlabel("Year")
        ax.set_ylabel("Present Value of Cash Flow (AUD)")
        ax.set_title(f"Option {tag}: {cfg['label']}", fontsize=9, fontweight="bold", color=col)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v:,.0f}"))
        ax.xaxis.set_major_locator(mticker.MultipleLocator(5))
        ax.legend(fontsize=7, loc="upper left")
        ax.grid(axis="y", alpha=0.3)

    fig.suptitle(
        f"Discounted Cash Flow — {ANALYSIS_YEARS}-Year Horizon  ·  "
        f"Discount rate {dr*100:.1f}%  ·  "
        f"Red bar = upfront cost, coloured bars = PV of annual savings",
        fontsize=9, fontweight="bold",
    )
    fig.tight_layout()
    return fig


def make_npv_sensitivity_fig(pb_a: dict, pb_b: dict, cfg_a: dict, cfg_b: dict) -> plt.Figure:
    """NPV as a function of discount rate — x-intercept is the IRR."""
    current_dr = pb_a.get("discount_rate", OPPORTUNITY_RATE)
    rates = np.arange(0.0, 0.205, 0.005)
    fig, ax = plt.subplots(figsize=(11, 4.5))
    fig.patch.set_facecolor("white")

    for pb, cfg, col, tag in [
        (pb_a, cfg_a, OPTION_COLOURS[0], "A"),
        (pb_b, cfg_b, OPTION_COLOURS[1], "B"),
    ]:
        cashflows = [-pb["net"]] + pb["savings"]
        npvs = [
            sum(cf / (1.0 + r) ** i for i, cf in enumerate(cashflows))
            for r in rates
        ]
        ax.plot(rates * 100, npvs, color=col, lw=2.0,
                label=f"Option {tag}: {cfg['label']}")
        irr = compute_irr(pb)
        if irr is not None and irr <= 0.20:
            ax.plot(irr * 100, 0, "o", color=col, ms=9, zorder=5,
                    label=f"IRR ({tag}) = {irr*100:.1f}%")

    ax.axhline(0, color="black", lw=0.8, ls=":")
    ax.axvline(current_dr * 100, color="dimgray", lw=1.2, ls="--",
               label=f"Current rate ({current_dr*100:.1f}%)")
    ax.set_xlabel("Discount / Hurdle Rate (%/yr)")
    ax.set_ylabel("NPV (AUD)")
    ax.set_title(
        f"NPV Sensitivity to Discount Rate — {ANALYSIS_YEARS}-Year Horizon\n"
        f"Where a line crosses $0 is the IRR (your effective annual return on the investment)"
    )
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v:,.0f}"))
    ax.legend(fontsize=8)
    ax.grid(axis="y", alpha=0.3)
    fig.tight_layout()
    return fig


def make_efficiency_scatter_fig(pbs: list, pb_a: dict, pb_b: dict, discount_rate: float):
    """NPV vs net capex investment efficiency scatter.

    All quotes × tariffs shown as small markers; selected options A and B
    are highlighted with stars.  Horizontal dashed line at NPV = $0.
    """
    active = [pb for pb in pbs if pb.get("net", 0) > 0]
    if not active and pb_a.get("net", 0) <= 0 and pb_b.get("net", 0) <= 0:
        return None

    tariff_col = {"A1 Flat": "#1a73e8", "Midday Saver": "#0f9d58"}
    tariff_mrk = {"A1 Flat": "o",       "Midday Saver": "s"}

    fig, ax = plt.subplots(figsize=(11, 6))
    fig.patch.set_facecolor("white")
    seen_tariffs: set = set()

    for pb in active:
        col = tariff_col.get(pb["tariff"], "#888")
        mrk = tariff_mrk.get(pb["tariff"], "o")
        leg = pb["tariff"] if pb["tariff"] not in seen_tariffs else "_"
        seen_tariffs.add(pb["tariff"])
        ax.scatter(pb["net"], pb["npv"], color=col, marker=mrk,
                   s=65, zorder=4, label=leg, alpha=0.8)
        short = str(pb["label"])[:22]
        ax.annotate(short, (pb["net"], pb["npv"]),
                    textcoords="offset points", xytext=(6, 3),
                    fontsize=7, color="#444444")

    # Highlight selected options A and B
    for pb, tag, col in [(pb_a, "A", "#e8463a"), (pb_b, "B", "#0f9d58")]:
        if pb.get("net", 0) > 0:
            ax.scatter(pb["net"], pb["npv"], color=col, marker="*",
                       s=220, zorder=7, edgecolors="white", linewidths=0.6,
                       label=f"Option {tag} (selected)")
            ax.annotate(f"  ★ {tag}: {str(pb['label'])[:22]}",
                        (pb["net"], pb["npv"]),
                        textcoords="offset points", xytext=(7, 4),
                        fontsize=7.5, fontweight="bold", color=col)

    ax.axhline(0, color="black", lw=1.3, ls="--",
               label=f"Hurdle: NPV = $0  ({discount_rate*100:.1f}%/yr)", zorder=3)

    ylo, yhi = ax.get_ylim()
    xlo, xhi = ax.get_xlim()
    ax.fill_between([xlo, xhi], ylo, 0, color="#e74c3c", alpha=0.04)
    ax.fill_between([xlo, xhi], 0, max(yhi, 1), color="#2ecc71", alpha=0.04)
    ax.set_xlim(xlo, xhi)
    ax.set_ylim(ylo, max(yhi, 1))

    ax.set_xlabel("Net System Cost after Rebates (AUD)", fontsize=10)
    ax.set_ylabel(f"NPV — {ANALYSIS_YEARS}-yr @ {discount_rate*100:.1f}%/yr (AUD)", fontsize=10)
    ax.set_title(
        "Investment Efficiency — NPV vs Net Capex\n"
        "★ = selected options  ·  Above dashed line: positive NPV (beats hurdle rate)  ·"
        "  Vertical gap below line = value shortfall",
        fontsize=9,
    )
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v:,.0f}"))
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v:,.0f}"))
    ax.legend(fontsize=8, loc="best")
    ax.grid(alpha=0.25)
    fig.tight_layout()
    return fig


def make_tornado_fig(pb_base: dict, raw_df_hash: int, cfg: dict,
                     shading: tuple, stc_price: float, rebates_inc: bool,
                     grid_charge: bool, gc_start: float, gc_end: float,
                     tariff_esc: float, discount_rate: float, _raw_df) -> plt.Figure:
    """Tornado sensitivity diagram — NPV swing for ±range on each key input."""
    base_npv = pb_base["npv"]
    solar, bat, inv, cost = cfg["solar"], cfg["bat"], cfg["inv"], cfg["cost"]
    tariff = cfg["tariff"]
    shade_s, shade_au, shade_w = shading

    def _npv(ts=tariff_esc, dr=discount_rate, c=cost,
             ss=shade_s, sau=shade_au, sw=shade_w):
        return cached_payback(
            raw_df_hash, solar, bat, inv, c, tariff, "_tornado_",
            ss, sau, sw, stc_price, rebates_inc,
            grid_charge, gc_start, gc_end, ts, dr, _raw_df,
        )["npv"]

    rows = [
        ("Tariff escalation",
         f"{max(0, tariff_esc - 0.02)*100:.1f}%/yr",
         f"{(tariff_esc + 0.03)*100:.1f}%/yr",
         _npv(ts=max(0, tariff_esc - 0.02)),
         _npv(ts=tariff_esc + 0.03)),

        ("Discount rate",
         f"{max(0, discount_rate - 0.04)*100:.1f}%/yr",
         f"{(discount_rate + 0.04)*100:.1f}%/yr",
         _npv(dr=max(0, discount_rate - 0.04)),
         _npv(dr=discount_rate + 0.04)),

        ("System cost",
         "−15%", "+15%",
         _npv(c=cost * 0.85),
         _npv(c=cost * 1.15)),

        ("Summer shading",
         f"×{max(0.10, shade_s * 0.75):.2f}",
         f"×{min(1.0,  shade_s * 1.25):.2f}",
         _npv(ss=max(0.10, shade_s * 0.75)),
         _npv(ss=min(1.0,  shade_s * 1.25))),

        ("Winter shading",
         f"×{max(0.05, shade_w * 0.70):.2f}",
         f"×{min(1.0,  shade_w * 1.30):.2f}",
         _npv(sw=max(0.05, shade_w * 0.70)),
         _npv(sw=min(1.0,  shade_w * 1.30))),
    ]

    rows.sort(key=lambda r: abs(r[3] - r[4]), reverse=True)

    all_vals = [r[3] for r in rows] + [r[4] for r in rows] + [base_npv, 0]
    x_range = max(all_vals) - min(all_vals) if max(all_vals) != min(all_vals) else 1
    margin = x_range * 0.015

    n = len(rows)
    fig, ax = plt.subplots(figsize=(11, max(4, n * 1.0 + 1.5)))
    fig.patch.set_facecolor("white")

    for i, (name, lo_lbl, hi_lbl, npv_lo, npv_hi) in enumerate(reversed(rows)):
        better = max(npv_lo, npv_hi)
        worse  = min(npv_lo, npv_hi)
        ax.barh(i, better - base_npv, left=base_npv,
                color="#2ecc71", alpha=0.80, height=0.55, zorder=3)
        ax.barh(i, worse  - base_npv, left=base_npv,
                color="#e74c3c", alpha=0.80, height=0.55, zorder=3)
        left_lbl  = lo_lbl if npv_lo <= npv_hi else hi_lbl
        right_lbl = hi_lbl if npv_lo <= npv_hi else lo_lbl
        ax.text(worse  - margin, i, left_lbl,  ha="right", va="center",
                fontsize=7.5, color="#c0392b")
        ax.text(better + margin, i, right_lbl, ha="left",  va="center",
                fontsize=7.5, color="#27ae60")

    ax.set_yticks(range(n))
    ax.set_yticklabels([r[0] for r in reversed(rows)], fontsize=9)
    ax.axvline(base_npv, color="black", lw=1.5, zorder=5,
               label=f"Base NPV: ${base_npv:,.0f}")
    ax.axvline(0, color="gray", lw=0.9, ls=":", alpha=0.6, label="NPV = $0")
    ax.set_xlabel(f"NPV — {ANALYSIS_YEARS}-yr (AUD)", fontsize=10)
    ax.set_title(
        f"Sensitivity (Tornado) — {str(cfg.get('label', ''))[:45]}  ·  {cfg['tariff']}\n"
        "Green = upside, Red = downside.  Widest bars = most impactful variables.",
        fontsize=10,
    )
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v:,.0f}"))
    ax.legend(fontsize=8, loc="lower right")
    ax.grid(axis="x", alpha=0.3)
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
                            shade_a: tuple, shade_b: tuple) -> plt.Figure:
    """2×2 grid — one subplot per season, each with 3 curves:
      · Option A unshaded  (shade = 1.0)
      · Option A shaded    (per-option shading factors)
      · Option B shaded    (per-option shading factors)
    shade_a / shade_b are (shade_summer, shade_autumn, shade_winter) tuples.
    """
    # Representative mid-season day-of-year (Southern Hemisphere); third element = shade tuple index
    SEASONS = [
        ("Summer (Dec–Feb)", 15,  0),   # index 0 = summer
        ("Autumn (Mar–May)", 105, 1),   # index 1 = autumn
        ("Winter (Jun–Aug)", 196, 2),   # index 2 = winter
        ("Spring (Sep–Nov)", 288, 1),   # spring uses autumn factor
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

    for ax, (season, doy, shade_idx) in zip(axes.flat, SEASONS):
        kw_a = cfg_a["solar"]
        kw_b = cfg_b["solar"]
        sv_a = shade_a[shade_idx]
        sv_b = shade_b[shade_idx]

        unshaded_a = _day_profile(kw_a, doy, 1.0)
        shaded_a   = _day_profile(kw_a, doy, sv_a)
        shaded_b   = _day_profile(kw_b, doy, sv_b)

        ax.plot(hours, unshaded_a, color=OPTION_COLOURS[0], ls="--", lw=1.4, alpha=0.6,
                label=f"A unshaded ({kw_a} kW)")
        ax.plot(hours, shaded_a,   color=OPTION_COLOURS[0], ls="-",  lw=2.0,
                label=f"A shaded  ({kw_a} kW, ×{sv_a:.2f})")
        ax.plot(hours, shaded_b,   color=OPTION_COLOURS[1], ls="-",  lw=2.0,
                label=f"B shaded  ({kw_b} kW, ×{sv_b:.2f})")

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
        ax.text(0.98, 0.97, f"A loss: {(1-sv_a)*100:.0f}%  B loss: {(1-sv_b)*100:.0f}%",
                transform=ax.transAxes, ha="right", va="top", fontsize=8,
                bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="#aaaaaa",
                          alpha=0.85, lw=0.8))

    fig.suptitle("Solar Generation Profiles by Season — Shaded vs Unshaded",
                 fontsize=11, fontweight="bold", y=1.01)
    fig.tight_layout()
    return fig


def make_load_heatmap_fig(raw_df: pd.DataFrame) -> plt.Figure:
    """Heatmap: months on y-axis, hours of day on x-axis, cell = avg kWh per 30-min slot.
    Right column = average daily total (kWh/day) per month.
    Bottom row = average kWh/slot per hour across all months."""
    from matplotlib.gridspec import GridSpec

    df = raw_df.copy()
    df["month"] = df["datetime"].dt.month
    df["hour"]  = df["datetime"].dt.hour

    pivot = (
        df.groupby(["month", "hour"])["consumption_kwh"]
          .mean()
          .unstack("hour")
          .reindex(index=range(1, 13), columns=range(24))
    ) * 2   # ×2: two 30-min slots per hour → convert to kWh/hour

    # Row totals: true average daily kWh per month (sum all 48 half-hour slots each day)
    row_totals = (
        df.groupby(["date", "month"])["consumption_kwh"]
          .sum()
          .groupby(level="month")
          .mean()
          .reindex(range(1, 13))
    )
    # Column totals: average kWh/hour for each hour, averaged across months
    col_totals = pivot.mean(axis=0)

    month_labels = ["Jan","Feb","Mar","Apr","May","Jun",
                    "Jul","Aug","Sep","Oct","Nov","Dec"]

    fig = plt.figure(figsize=(21, 8))
    fig.patch.set_facecolor("white")
    gs = GridSpec(2, 2, figure=fig,
                  width_ratios=[22, 2], height_ratios=[11, 1],
                  wspace=0.02, hspace=0.03)
    ax_main   = fig.add_subplot(gs[0, 0])
    ax_right  = fig.add_subplot(gs[0, 1])
    ax_bottom = fig.add_subplot(gs[1, 0])
    ax_corner = fig.add_subplot(gs[1, 1])
    ax_corner.axis("off")

    # ── Main heatmap ──────────────────────────────────────────────────────────
    data = pivot.values
    vmax = np.nanmax(data)
    im = ax_main.imshow(data, aspect="auto", cmap="YlOrRd", interpolation="nearest")
    ax_main.set_xticks(np.arange(-0.5, 24, 1), minor=True)
    ax_main.set_yticks(np.arange(-0.5, 12, 1), minor=True)
    ax_main.grid(which="minor", color="white", linewidth=0.5)
    ax_main.tick_params(which="minor", bottom=False, left=False)
    threshold = vmax * 0.55
    for r in range(12):
        for c in range(24):
            val = data[r, c]
            if np.isnan(val):
                continue
            ax_main.text(c, r, f"{val:.2f}", ha="center", va="center",
                         fontsize=5.5, fontweight="bold",
                         color="white" if val > threshold else "#333333")
    cbar = fig.colorbar(im, ax=ax_main, pad=0.01, fraction=0.015)
    cbar.set_label("Avg kWh / hour", fontsize=8)
    ax_main.set_yticks(range(12))
    ax_main.set_yticklabels(month_labels, fontsize=9)
    ax_main.set_xticks(range(0, 24, 2))
    ax_main.set_xticklabels([f"{h:02d}:00" for h in range(0, 24, 2)],
                             fontsize=8, rotation=45, ha="right")
    ax_main.set_title("Average Electricity Consumption — Month × Hour of Day",
                      fontsize=11, fontweight="bold")

    # ── Right column: daily average total per month ───────────────────────────
    rt = row_totals.values.reshape(12, 1)
    rt_max = np.nanmax(rt)
    ax_right.imshow(rt, aspect="auto", cmap="Blues", interpolation="nearest",
                    vmin=0, vmax=rt_max * 1.05)
    ax_right.set_xticks([0])
    ax_right.set_xticklabels(["Daily avg\n(kWh/day)"], fontsize=7)
    ax_right.set_yticks([])
    ax_right.tick_params(which="both", left=False, bottom=False)
    rt_thresh = rt_max * 0.55
    for r in range(12):
        val = rt[r, 0]
        if not np.isnan(val):
            ax_right.text(0, r, f"{val:.1f}", ha="center", va="center",
                          fontsize=7.5, fontweight="bold",
                          color="white" if val > rt_thresh else "#333333")

    # ── Bottom row: hourly average across all months ──────────────────────────
    ct = col_totals.values.reshape(1, 24)
    ct_max = np.nanmax(ct)
    ax_bottom.imshow(ct, aspect="auto", cmap="Blues", interpolation="nearest",
                     vmin=0, vmax=ct_max * 1.05)
    ax_bottom.set_yticks([0])
    ax_bottom.set_yticklabels(["All-month\navg"], fontsize=7)
    ax_bottom.set_xticks([])
    ax_bottom.tick_params(which="both", left=False, bottom=False)
    ct_thresh = ct_max * 0.55
    for c in range(24):
        val = ct[0, c]
        if not np.isnan(val):
            ax_bottom.text(c, 0, f"{val:.2f}", ha="center", va="center",
                           fontsize=5.5, fontweight="bold",
                           color="white" if val > ct_thresh else "#333333")

    fig.tight_layout()
    return fig


def make_soc_seasonal_fig(res_a, res_b, cfg_a, cfg_b) -> plt.Figure:
    """4×2 grid: rows = seasons, cols = [Option A | Option B].
    Shows the typical average battery state of charge over the day per season."""
    options = [
        (res_a, cfg_a, OPTION_COLOURS[0], "A"),
        (res_b, cfg_b, OPTION_COLOURS[1], "B"),
    ]

    global_soc_ylim = max(cfg_a["bat"], cfg_b["bat"], 0.1) * BAT_DOD * 1.15

    fig, axes = plt.subplots(4, 2, figsize=(12, 16), sharey=True)
    fig.patch.set_facecolor("white")

    for col_idx, (res, cfg, col, tag) in enumerate(options):
        df = res["df"].copy()
        df["month"] = df["datetime"].dt.month
        bat_usable = cfg["bat"] * BAT_DOD

        for row_idx, (season, months) in enumerate(SEASON_MONTHS.items()):
            ax = axes[row_idx][col_idx]
            sub = df[df["month"].isin(months)]

            if sub.empty:
                ax.set_visible(False)
                continue

            avg_soc = sub.groupby("slot")["soc"].mean()
            h = avg_soc.index / 2

            ax.fill_between(h, avg_soc.values, alpha=0.35, color=col)
            ax.plot(h, avg_soc.values, color=col, lw=2.0, label="Avg SoC")

            if bat_usable > 0:
                ax.axhline(bat_usable, color="gray", lw=0.9, ls="--",
                           label=f"Max usable ({bat_usable:.1f} kWh)")
                ax.set_ylim(0, global_soc_ylim)
                avg_val = float(avg_soc.mean())
                avg_pct = avg_val / bat_usable * 100
                ax.text(0.98, 0.97, f"Avg SoC: {avg_val:.1f} kWh ({avg_pct:.0f}%)",
                        transform=ax.transAxes, ha="right", va="top", fontsize=7,
                        bbox=dict(boxstyle="round,pad=0.3", fc="white",
                                  ec=col, alpha=0.85, lw=1.0))
            else:
                ax.set_ylim(0, 1)
                ax.text(0.5, 0.5, "No battery configured", transform=ax.transAxes,
                        ha="center", va="center", fontsize=9, color="gray")

            ax.axvspan(15, 21, alpha=0.07, color="red")
            ax.axvspan(9, 15, alpha=0.06, color="green")
            ax.set_xlim(0, 24)
            ax.set_xticks(range(0, 25, 6))
            ax.set_xticklabels([f"{hh:02d}h" for hh in range(0, 25, 6)], fontsize=7)
            ax.tick_params(axis="y", labelsize=7)

            if row_idx == 0:
                ax.set_title(f"Option {tag}: {cfg['label']}", fontsize=9,
                             fontweight="bold", color=col)
            if col_idx == 0:
                ax.set_ylabel(f"{season}\nSoC (kWh)", fontsize=7.5, fontweight="bold")
            else:
                ax.set_ylabel("SoC (kWh)", fontsize=7)

    handles, labels_ = axes[0][0].get_legend_handles_labels()
    fig.legend(handles, labels_, loc="lower center", ncol=2, fontsize=8,
               bbox_to_anchor=(0.5, -0.005))
    fig.suptitle("Typical Battery State of Charge by Season",
                 fontsize=11, fontweight="bold")
    fig.tight_layout(rect=[0, 0.03, 1, 0.98])
    return fig


def make_seasonal_energy_table(res_a, res_b) -> pd.DataFrame:
    """Avg daily consumption and solar kWh per season for Option A and Option B."""
    rows = []
    for season, months in SEASON_MONTHS.items():
        row: dict = {"Season": season}
        for res, tag in [(res_a, "A"), (res_b, "B")]:
            df = res["df"].copy()
            df["month"] = df["datetime"].dt.month
            sub = df[df["month"].isin(months)]
            if sub.empty:
                if tag == "A":
                    row["Avg Daily Consumption (kWh)"] = None
                row[f"Option {tag} Avg Daily Solar (kWh)"] = None
            else:
                daily = sub.groupby("date")[["consumption_kwh", "solar_kwh"]].sum()
                if tag == "A":
                    row["Avg Daily Consumption (kWh)"] = round(daily["consumption_kwh"].mean(), 1)
                row[f"Option {tag} Avg Daily Solar (kWh)"] = round(daily["solar_kwh"].mean(), 1)
        rows.append(row)
    return pd.DataFrame(rows).set_index("Season")


def _make_summary_table_fig(pb_a, pb_b, cfg_a, cfg_b, tariff_esc, discount_rate) -> plt.Figure:
    """One-page financial summary table for the PDF report."""
    import matplotlib.colors as mcolors
    dr     = pb_a.get("discount_rate", OPPORTUNITY_RATE)
    dr_pct = dr * 100
    rows = []
    for pb, cfg, tag in [(pb_a, cfg_a, "A"), (pb_b, cfg_b, "B")]:
        npv10 = -pb["net"] + sum(pb["savings"][yr-1] / (1+dr)**yr for yr in range(1, 11))
        irr   = compute_irr(pb)
        rows.append([
            f"Option {tag}  ·  {cfg['tariff']}",
            f"${pb['net']:,.0f}",
            f"{pb['pb_yr']} yr" if pb["pb_yr"] else ">20 yr",
            f"{pb['pb_yr_disc']} yr" if pb.get("pb_yr_disc") else ">20 yr",
            f"${npv10:,.0f}",
            f"${pb['npv']:,.0f}",
            f"{irr*100:.1f}%" if irr else "N/A",
            f"${pb['total_save']:,.0f}",
            f"{pb['roi']:.0f}%",
            f"${pb['yr1_save']:,.0f}",
        ])
    col_labels = [
        "Option / Tariff", "Net Cost",
        "Nominal\nPayback", f"Disc. Payback\n({dr_pct:.1f}%)",
        f"NPV 10yr\n@ {dr_pct:.1f}%", f"NPV 20yr\n@ {dr_pct:.1f}%",
        "IRR", "20yr Saving", "20yr Return", "Yr-1 Saving",
    ]
    fig, ax = plt.subplots(figsize=(14, 3.5))
    fig.patch.set_facecolor("white")
    ax.axis("off")
    tbl = ax.table(cellText=rows, colLabels=col_labels, cellLoc="center", loc="center")
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(9)
    tbl.scale(1.0, 2.5)
    for c in range(len(col_labels)):
        tbl[0, c].set_facecolor("#1c2b3a")
        tbl[0, c].set_text_props(color="white", fontweight="bold")
    light_a = (*mcolors.to_rgb(OPTION_COLOURS[0]), 0.15)
    light_b = (*mcolors.to_rgb(OPTION_COLOURS[1]), 0.15)
    for c in range(len(col_labels)):
        tbl[1, c].set_facecolor(light_a)
        tbl[2, c].set_facecolor(light_b)
    ax.set_title("Financial Summary", fontsize=13, fontweight="bold", pad=20)
    fig.tight_layout()
    return fig


def _build_pdf(raw_df, res_a, res_b, res_base, pb_a, pb_b,
               cfg_a, cfg_b, shade_a: tuple, shade_b: tuple,
               tariff_esc, discount_rate) -> bytes:
    """Compile all charts + explanatory text into a multi-page PDF."""
    import io as _io
    import textwrap
    from datetime import date
    from matplotlib.backends.backend_pdf import PdfPages

    dr     = pb_a.get("discount_rate", OPPORTUNITY_RATE)
    dr_pct = dr * 100
    today  = date.today().strftime("%d %B %Y")
    cfg_base_local = {"label": "Status Quo", "tariff": cfg_a["tariff"]}
    PAGE   = (11.69, 8.27)
    DARK   = "#1c2b3a"
    LIGHT  = "#f4f6f8"

    def _footer(fig):
        fig.text(0.97, 0.012, f"Solar & Battery Investment Report  ·  {today}",
                 fontsize=7, ha="right", va="bottom", color="#999999",
                 transform=fig.transFigure)

    def _save_chart(pdf, fig, caption):
        """Resize to A4 landscape, add caption strip at bottom, save."""
        fig.set_size_inches(*PAGE)
        fig.patch.set_facecolor("white")
        fig.subplots_adjust(bottom=0.16, top=0.93)
        wrapped = "\n".join(textwrap.wrap(caption, 148))
        fig.text(0.03, 0.005, wrapped, fontsize=8, va="bottom", color="#444444",
                 style="italic", transform=fig.transFigure)
        _footer(fig)
        pdf.savefig(fig, bbox_inches="tight", dpi=72)
        plt.close(fig)

    def _text_page(pdf, section_num, title, paras):
        """A4 landscape text page with a dark header bar and body paragraphs."""
        fig = plt.figure(figsize=PAGE)
        fig.patch.set_facecolor("white")
        # Header bar
        ax_h = fig.add_axes([0.0, 0.88, 1.0, 0.12])
        ax_h.set_facecolor(DARK); ax_h.axis("off")
        ax_h.text(0.03, 0.65, f"Section {section_num}", color="#aabccc",
                  fontsize=9, va="center", ha="left", transform=ax_h.transAxes)
        ax_h.text(0.03, 0.25, title, color="white", fontsize=15, fontweight="bold",
                  va="center", ha="left", transform=ax_h.transAxes)
        ax_h.text(0.97, 0.45, f"Solar & Battery Investment Report  ·  {today}",
                  color="#aabccc", fontsize=8, va="center", ha="right",
                  transform=ax_h.transAxes)
        # Body
        y = 0.82
        for para in paras:
            if para.startswith("  •"):           # bullet
                wrapped = textwrap.fill(para, 105)
                fig.text(0.06, y, wrapped, fontsize=10, va="top", color="#333333",
                         linespacing=1.4, transform=fig.transFigure)
                y -= 0.05 + wrapped.count("\n") * 0.033
            else:
                wrapped = textwrap.fill(para, 105)
                fig.text(0.06, y, wrapped, fontsize=10.5, va="top", color="#222222",
                         linespacing=1.55, transform=fig.transFigure)
                y -= 0.065 + wrapped.count("\n") * 0.038
        _footer(fig)
        pdf.savefig(fig, bbox_inches="tight", dpi=72)
        plt.close(fig)

    buf = _io.BytesIO()
    with PdfPages(buf) as pdf:

        # ── Cover page ──────────────────────────────────────────────────────
        fig = plt.figure(figsize=PAGE)
        fig.patch.set_facecolor(DARK)
        fig.text(0.5, 0.91, "Solar & Battery Investment Report",
                 ha="center", fontsize=28, fontweight="bold", color="white")
        fig.text(0.5, 0.84, "Perth, Western Australia  ·  Synergy Network",
                 ha="center", fontsize=13, color="#aabccc")
        fig.text(0.5, 0.79, f"Prepared: {today}",
                 ha="center", fontsize=10, color="#7a8fa0")

        for xi, pb, cfg, col in [
            (0.25, pb_a, cfg_a, OPTION_COLOURS[0]),
            (0.75, pb_b, cfg_b, OPTION_COLOURS[1]),
        ]:
            tag   = "A" if xi < 0.5 else "B"
            irr   = compute_irr(pb)
            irr_s = f"{irr*100:.1f}%" if irr else "N/A"
            npv10 = -pb["net"] + sum(pb["savings"][yr-1]/(1+dr)**yr for yr in range(1, 11))
            ax = fig.add_axes([xi - 0.21, 0.04, 0.42, 0.67])
            ax.set_facecolor(col); ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")
            # Option header
            ax.add_patch(plt.Rectangle((0, 0.88), 1, 0.12, color="black", alpha=0.2))
            ax.text(0.5, 0.94, f"Option {tag}", ha="center", va="center",
                    fontsize=14, fontweight="bold", color="white")
            # Label
            ax.text(0.5, 0.81, cfg["label"][:52], ha="center", va="center",
                    fontsize=8, color="white", alpha=0.9)
            # Divider
            ax.axhline(0.75, color="white", alpha=0.3, lw=0.8)
            # Metrics
            metrics = [
                ("Net system cost",            f"${pb['net']:,.0f}"),
                ("Nominal payback",             f"{pb['pb_yr']} yr" if pb["pb_yr"] else ">20 yr"),
                (f"NPV 10yr @ {dr_pct:.0f}%",  f"${npv10:,.0f}"),
                (f"NPV 20yr @ {dr_pct:.0f}%",  f"${pb['npv']:,.0f}"),
                ("IRR",                         irr_s),
                ("20-yr total return",          f"{pb['roi']:.0f}%"),
                ("Year-1 saving",               f"${pb['yr1_save']:,.0f}"),
            ]
            for i, (label, value) in enumerate(metrics):
                y_pos = 0.68 - i * 0.095
                ax.text(0.08, y_pos, label, va="center", fontsize=9.5, color="white", alpha=0.85)
                ax.text(0.92, y_pos, value, va="center", fontsize=9.5, color="white",
                        fontweight="bold", ha="right")
        pdf.savefig(fig, bbox_inches="tight", dpi=72); plt.close(fig)

        # ── Section 1: Consumption ───────────────────────────────────────────
        _text_page(pdf, 1, "Your Electricity Consumption", [
            "Before evaluating solar and battery options, it helps to understand when and how much "
            "electricity your household uses. The chart on the next page shows your average consumption "
            "for every hour of the day, broken down by month — based on your actual Synergy meter data.",
            "What to look for: the darkest orange and red cells show when your demand is highest. For "
            "most Perth homes this is the evening (5–9 pm) when people are home cooking, watching TV, "
            "and running appliances, and to a lesser extent the morning (6–8 am).",
            "Solar panels generate electricity between roughly 7 am and 5 pm. Any demand outside those "
            "hours must be met by a battery (if you have one) or the grid. A large evening peak relative "
            "to daytime demand makes a battery particularly valuable — it stores cheap midday solar and "
            "releases it into the expensive evening peak window.",
        ])
        _save_chart(pdf, make_load_heatmap_fig(raw_df),
                    "Each cell shows average electricity use (kWh per hour) in a given month and hour "
                    "of day. Darker orange/red = higher consumption. Solar generation window is roughly 7 am – 5 pm.")

        # ── Section 2: Solar Generation ──────────────────────────────────────
        _text_page(pdf, 2, "Solar Generation by Season", [
            "Solar panels generate electricity from sunlight, with output varying throughout the day and "
            "across seasons. Perth enjoys some of Australia's best solar resources, but performance still "
            "drops significantly in winter due to shorter days and a lower sun angle in the sky.",
            "The chart on the next page shows the expected generation profile for each option across the "
            "four seasons. Two curves are shown for Option A: the dashed line is the theoretical maximum "
            "with no shading, and the solid line is the estimated actual output after applying your "
            "seasonal shading factors. Option B's shaded output is also shown.",
            "Shading from trees, neighbouring buildings, or roof obstructions can meaningfully reduce "
            "annual yield — even partial shading can cut output disproportionately. The shaded area "
            "between the two Option A lines shows how much generation is being lost.",
        ])
        _save_chart(pdf, make_solar_profile_fig(cfg_a, cfg_b, shade_a, shade_b),
                    "Dashed line = theoretical maximum (no shading). Solid line = estimated actual output "
                    "with your seasonal shading factors applied. Shaded area shows generation lost to "
                    "obstructions. Winter output is lower due to Perth's shorter days at 32° S latitude.")

        # ── Section 3: Energy Flows ───────────────────────────────────────────
        _text_page(pdf, 3, "How Your Energy Is Used — Typical Day by Season", [
            "This section shows how your electricity is supplied on a typical day in each season, "
            "comparing three scenarios: no solar (Status Quo), Option A, and Option B.",
            "Reading the charts: the stacked coloured areas show where your electricity comes from at "
            "each point in the day. Yellow is solar power used directly in your home (free electricity). "
            "Teal is energy supplied by the battery. Blue is electricity imported from the grid (costs "
            "money). The light blue area above the black load line shows surplus solar exported back to "
            "the grid, earning DEBS (Distributed Energy Buyback Scheme) credits.",
            "The cost box in each chart shows the average daily electricity cost and how much is saved "
            "compared to the no-solar baseline. Green shading highlights the Synergy Midday Saver "
            "super off-peak window (9 am–3 pm, 8.6 ¢/kWh) and red marks the peak window "
            "(3–9 pm, 53.8 ¢/kWh). Batteries charged during the green window and discharged in the "
            "red window deliver the greatest savings.",
        ])
        _save_chart(pdf, make_seasonal_fig(res_base, res_a, res_b, cfg_base_local, cfg_a, cfg_b),
                    "Stacked areas: solar self-use (yellow), battery discharge (teal), grid import (blue). "
                    "Light blue above the load line = solar exported to grid (earns DEBS credits). "
                    "Green band = cheap super off-peak (9 am–3 pm); red band = expensive peak (3–9 pm).")

        # ── Section 4: Battery ────────────────────────────────────────────────
        _text_page(pdf, 4, "Battery State of Charge", [
            "A battery stores surplus solar energy generated during the day and releases it later — "
            "typically during the evening peak when grid electricity is most expensive.",
            "The chart shows the average battery charge level (in kWh) at each hour of the day across "
            "all four seasons. The dashed horizontal line marks the maximum usable capacity, which is "
            f"the nameplate size reduced by the depth-of-discharge limit ({BAT_DOD*100:.0f}%) to protect "
            "battery lifespan.",
            "Ideally, the battery should be well charged by 3 pm to cover the evening peak. If the "
            "battery is frequently exhausted before 9 pm, it may be undersized relative to your "
            "evening demand, or the charging window may need adjustment. A battery that is consistently "
            "full by midday and still has charge at 9 pm is well-matched to your usage pattern.",
        ])
        _save_chart(pdf, make_soc_seasonal_fig(res_a, res_b, cfg_a, cfg_b),
                    f"Average battery charge level (kWh) through the day. Dashed line = maximum usable "
                    f"capacity (nameplate × {BAT_DOD*100:.0f}% DoD limit). Green band = cheap charging "
                    f"window; red band = peak discharge window.")

        # ── Section 5: Monthly Bills ──────────────────────────────────────────
        _text_page(pdf, 5, "Monthly Electricity Bills", [
            "This section compares estimated monthly electricity bills across the three scenarios, "
            "calculated at today's tariff rates. It shows the seasonal pattern of savings — solar is "
            "generally most valuable in summer when days are long and panels produce the most energy.",
            "Months with bars below $0 indicate that solar export earnings exceed all electricity "
            "charges for that month — meaning the system effectively earns money rather than costing "
            "money in those months. This is most common in summer.",
            f"Important note: this chart uses today's electricity rates. The 20-year financial "
            f"projections in the following sections account for electricity price escalation "
            f"({tariff_esc*100:.1f}%/yr) and the gradual decline in DEBS export credits (5%/yr), "
            f"which significantly affect the long-run economics.",
        ])
        _save_chart(pdf, make_monthly_fig(res_base, res_a, res_b, cfg_base_local, cfg_a, cfg_b),
                    f"Monthly electricity bills at year-0 tariff rates. Negative values = export "
                    f"credits exceed all charges for that month. Seasonal patterns reflect Perth's "
                    f"strong summer solar resource.")

        # ── Section 6: Payback & Cash Flow ───────────────────────────────────
        _text_page(pdf, 6, "Payback Period & Cumulative Cash Flow", [
            "The most common question about solar and batteries is: when does it pay for itself?",
            "The payback period is the number of years until your cumulative electricity bill savings "
            "equal the net purchase price of the system. For example, a 7-year payback on a $15,000 "
            "net-cost system means you recoup the full investment through reduced bills in 7 years. "
            "Everything after that is pure profit.",
            "The three charts on the following pages show this from different angles: (1) raw cumulative "
            "cashflow showing savings building up from the initial investment; (2) total money spent — "
            "solar system plus ongoing bills compared to just paying bills indefinitely with no solar; "
            "and (3) the same total spend adjusted for the time value of money (present value). The "
            "present-value chart is the most rigorous — it asks whether the investment beats putting "
            "the money in a bank or investment account instead.",
        ])
        _save_chart(pdf, make_payback_fig(pb_a, pb_b, cfg_a, cfg_b),
                    f"Starts at −[net system cost] in Year 0. The line rises each year as bill savings "
                    f"accumulate. Where the line crosses $0 is the nominal payback year. Electricity "
                    f"prices assumed to escalate at {tariff_esc*100:.1f}%/yr; DEBS credits decline 5%/yr.")
        _save_chart(pdf, make_total_spend_fig(pb_a, pb_b, cfg_a, cfg_b),
                    "Total money spent: upfront system cost plus all ongoing electricity bills over 20 years. "
                    "Where a solar/battery line dips below the grey no-solar baseline, the system has paid "
                    "for itself in total expenditure terms — you are spending less in total than if you had "
                    "done nothing.")
        _save_chart(pdf, make_pv_spend_fig(pb_a, pb_b, cfg_a, cfg_b),
                    f"Same as above, but future bills are discounted to today's dollars at {dr_pct:.1f}%/yr "
                    f"— representing what you could have earned by investing elsewhere. Crossing below the "
                    f"grey baseline indicates a positive net present value at this discount rate.")

        # ── Section 7: NPV & IRR ──────────────────────────────────────────────
        _text_page(pdf, 7, "Net Present Value (NPV) & Internal Rate of Return (IRR)", [
            "NPV and IRR are the gold-standard metrics for comparing the true value of any investment "
            "after accounting for the fact that money today is worth more than money in the future "
            "(because today's money can be invested and grow).",
            "Net Present Value (NPV): add up all future bill savings, shrink each one back to today's "
            f"value by discounting at {dr_pct:.1f}%/yr, then subtract the upfront cost. A positive NPV "
            "means the solar system delivers more value than you'd get by putting the same money in the "
            "bank at that rate. Larger positive NPV is better.",
            "Internal Rate of Return (IRR): the effective annual percentage return on your investment — "
            "analogous to the interest rate a bank account would need to pay to match the solar system. "
            "For example, an IRR of 12% means solar earns the equivalent of a 12%/yr return. "
            "Compare to alternatives: term deposits ~4–5%/yr; diversified share portfolio ~8–10%/yr "
            "long-term. If the IRR exceeds your next best alternative, solar is the better investment.",
            f"The analysis uses a {ANALYSIS_YEARS}-year horizon and nominal cashflows (future dollar "
            "values, not adjusted for general inflation).",
        ])
        _save_chart(pdf, make_npv_fig(pb_a, pb_b, cfg_a, cfg_b),
                    f"Each bar = present value of that year's bill savings, discounted at {dr_pct:.1f}%/yr. "
                    f"The line = running cumulative NPV. Red bar = upfront investment. Final NPV and IRR "
                    f"are shown in the annotation box. NPV > $0 means the investment beats a "
                    f"{dr_pct:.1f}%/yr alternative.")
        _save_chart(pdf, make_npv_sensitivity_fig(pb_a, pb_b, cfg_a, cfg_b),
                    "Shows how NPV changes if a different discount rate is applied. The dot marks the IRR "
                    "(where NPV crosses $0). If the IRR dot is to the right of your required return rate, "
                    "the investment exceeds that hurdle. A steeper downward curve means results are more "
                    "sensitive to the discount rate assumption.")

        # ── Section 8: Financial Summary ──────────────────────────────────────
        _text_page(pdf, 8, "Financial Summary & Modelling Assumptions", [
            "The table on the next page summarises all key financial metrics side by side for both options.",
            "Modelling assumptions used in all 20-year projections:",
            f"  •  Electricity price escalation: {tariff_esc*100:.1f}%/yr",
            f"  •  DEBS export credit decline: 5%/yr (Synergy policy — credits reduce over time)",
            f"  •  Solar panel output degradation: 0.5%/yr (manufacturer typical warranty)",
            f"  •  Battery capacity degradation: 2%/yr (lithium battery typical)",
            f"  •  Battery depth of discharge (DoD): {BAT_DOD*100:.0f}% (protects battery lifespan)",
            f"  •  System efficiency (inverter + wiring losses): {SYS_EFF*100:.0f}%",
            f"  •  Discount / opportunity-cost rate: {dr_pct:.1f}%/yr",
            f"  •  Analysis horizon: {ANALYSIS_YEARS} years",
        ])
        _save_chart(pdf, _make_summary_table_fig(pb_a, pb_b, cfg_a, cfg_b, tariff_esc, discount_rate),
                    "All figures in nominal AUD. Disc. payback = discounted payback period at the chosen "
                    "opportunity-cost rate. NPV = net present value. IRR = internal rate of return. "
                    "20yr return = total 20-year savings divided by net system cost (not annualised).")

    buf.seek(0)
    return buf.read()


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
    st.subheader("2. Manage quotes")

    # ── Optional: upload a different quotes file ──────────────────────────────
    _quotes_upload = st.file_uploader(
        "Upload quotes file (.xlsx or .csv)",
        type=["xlsx", "xls", "csv"],
        help="Optional — uses the bundled solar_battery_quotes.xlsx if not uploaded.",
        key="quotes_upload",
    )

    # ── Load + normalise quotes from file ────────────────────────────────────
    def _load_quotes_df(source) -> "pd.DataFrame | None":
        try:
            if isinstance(source, Path):
                df = pd.read_excel(source)
            else:
                buf = io.BytesIO(source.getvalue())
                suffix = os.path.splitext(source.name)[1].lower()
                df = pd.read_excel(buf) if suffix in (".xlsx", ".xls") else pd.read_csv(buf)
        except Exception as exc:
            st.warning(f"Could not load quotes file: {exc}")
            return None

        df.columns = df.columns.str.strip()
        cols = {c.lower(): c for c in df.columns}

        def _fc(keywords):
            for kw in keywords:
                for lc, orig in cols.items():
                    if kw in lc:
                        return orig
            raise ValueError(f"No column matching {keywords} in {list(df.columns)}")

        def _fc_opt(keywords):
            for kw in keywords:
                for lc, orig in cols.items():
                    if kw in lc:
                        return orig
            return None

        try:
            col_vendor   = _fc(["vendor", "name", "label"])
            col_solar    = _fc(["solar"])
            col_battery  = _fc(["battery", "bat"])
            col_inverter = _fc(["inverter", "inv"])
            col_cost     = _fc(["price", "cost", "aud"])
        except ValueError as exc:
            st.warning(str(exc))
            return None

        col_order = next(
            (cols[lc] for lc in cols if any(k in lc for k in ["quote", "number", "order"])),
            None,
        )
        col_rebates_inc = _fc_opt(["rebate", "include", "already"])
        col_shade_s  = _fc_opt(["shade_summer"])
        col_shade_au = _fc_opt(["shade_autumn"])
        col_shade_w  = _fc_opt(["shade_winter"])

        if col_order:
            df = df.sort_values(col_order)

        df = df.rename(columns={
            col_vendor:   "Vendor",
            col_solar:    "Solar_kW",
            col_battery:  "Battery_kWh",
            col_inverter: "Inverter_kW",
            col_cost:     "Cost_AUD",
        }).reset_index(drop=True)

        for col_nm, key, default in [
            ("Shade_Summer", col_shade_s,  0.60),
            ("Shade_Autumn", col_shade_au, 0.30),
            ("Shade_Winter", col_shade_w,  0.15),
        ]:
            df[col_nm] = pd.to_numeric(df[key], errors="coerce").fillna(default) if key else default

        if col_rebates_inc:
            df["Rebates_Included"] = (
                df[col_rebates_inc].astype(str).str.strip().str.lower()
                .isin(["yes", "y", "true", "1"])
            )
        else:
            df["Rebates_Included"] = False

        for c in ["Solar_kW", "Battery_kWh", "Inverter_kW", "Cost_AUD"]:
            df[c] = pd.to_numeric(df[c], errors="coerce")

        df = df.dropna(subset=["Vendor", "Solar_kW", "Battery_kWh", "Inverter_kW", "Cost_AUD"])
        return df[df["Vendor"].astype(str).str.strip() != ""].reset_index(drop=True)

    _src = _quotes_upload if _quotes_upload is not None else (
        _QUOTES_PATH if _QUOTES_PATH.exists() else None
    )
    _raw_quotes_df = _load_quotes_df(_src) if _src is not None else None

    # ── Editable quotes table ────────────────────────────────────────────────
    _QCOLS = ["Vendor", "Solar_kW", "Battery_kWh", "Inverter_kW",
              "Cost_AUD", "Rebates_Included", "Shade_Summer", "Shade_Autumn", "Shade_Winter"]
    _QDEFAULTS = {
        "Vendor": "", "Solar_kW": 6.6, "Battery_kWh": 13.5, "Inverter_kW": 5.0,
        "Cost_AUD": 15000, "Rebates_Included": False,
        "Shade_Summer": 0.60, "Shade_Autumn": 0.30, "Shade_Winter": 0.15,
    }
    _edit_base = (
        _raw_quotes_df[_QCOLS].copy()
        if _raw_quotes_df is not None
        else pd.DataFrame([_QDEFAULTS])
    )

    _valid_edit: pd.DataFrame = pd.DataFrame(columns=_QCOLS)

    with st.expander("Edit / add quotes", expanded=(_raw_quotes_df is None)):
        st.caption(
            "Edit cells directly, add rows with the **+** button, or delete rows by selecting them. "
            "Changes apply for this session only — download the file to save permanently."
        )
        _edited = st.data_editor(
            _edit_base,
            num_rows="dynamic",
            column_config={
                "Vendor":           st.column_config.TextColumn("Vendor / Label", width="large"),
                "Solar_kW":         st.column_config.NumberColumn("Solar (kW)",    min_value=0.0, max_value=30.0, step=0.5,  format="%.2g"),
                "Battery_kWh":      st.column_config.NumberColumn("Battery (kWh)", min_value=0.0, max_value=60.0, step=0.5,  format="%.4g"),
                "Inverter_kW":      st.column_config.NumberColumn("Inverter (kW)", min_value=0.0, max_value=15.0, step=0.5,  format="%.4g"),
                "Cost_AUD":         st.column_config.NumberColumn("Cost ($AUD)",   min_value=0,                  step=500),
                "Rebates_Included": st.column_config.CheckboxColumn("Rebates incl.?"),
                "Shade_Summer":     st.column_config.NumberColumn("Shade Summer",  min_value=0.0, max_value=1.0, step=0.05, format="%.2f"),
                "Shade_Autumn":     st.column_config.NumberColumn("Shade Au/Sp",   min_value=0.0, max_value=1.0, step=0.05, format="%.2f"),
                "Shade_Winter":     st.column_config.NumberColumn("Shade Winter",  min_value=0.0, max_value=1.0, step=0.05, format="%.2f"),
            },
            hide_index=True,
            use_container_width=True,
            key="quotes_editor",
        )

        _valid_edit = _edited.dropna(
            subset=["Vendor", "Solar_kW", "Battery_kWh", "Inverter_kW", "Cost_AUD"]
        )
        _valid_edit = _valid_edit[
            _valid_edit["Vendor"].astype(str).str.strip() != ""
        ].reset_index(drop=True)

        if len(_valid_edit) > 0:
            _dl_buf = io.BytesIO()
            with pd.ExcelWriter(_dl_buf, engine="openpyxl") as _w:
                _valid_edit.to_excel(_w, sheet_name="quotes", index=False)
            _dl_buf.seek(0)
            st.download_button(
                "⬇️ Download quotes as Excel",
                data=_dl_buf.read(),
                file_name="solar_battery_quotes.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # Use the (possibly edited) quotes for the rest of the session
    solar_quotes = (
        _valid_edit[~((_valid_edit["Solar_kW"] == 0) & (_valid_edit["Battery_kWh"] == 0))]
        .reset_index(drop=True)
        if len(_valid_edit) > 0 else None
    )

    TARIFFS = ["Midday Saver", "A1 Flat"]

    def option_selector(tag: str, default_idx: int):
        colour = "#e8463a" if tag == "A" else "#0f9d58"
        _def_s, _def_au, _def_w = 0.60, 0.30, 0.15
        st.markdown(f"**Option {tag}**")
        if solar_quotes is not None and len(solar_quotes) > 0:
            CUSTOM_IDX = len(solar_quotes)

            def _fmt(i):
                if i == CUSTOM_IDX:
                    return "— Custom —"
                r = solar_quotes.iloc[i]
                price_tag = "net, rebates incl." if r.get("Rebates_Included", False) else "gross"
                return (f"{r['Vendor']}  —  {r['Solar_kW']:.2g} kW solar / "
                        f"{r['Battery_kWh']:.4g} kWh battery  "
                        f"(${int(r['Cost_AUD']):,} {price_tag})")

            sel_idx = st.selectbox(
                "Quote", range(CUSTOM_IDX + 1),
                format_func=_fmt,
                index=min(default_idx, CUSTOM_IDX - 1),
                key=f"quote_{tag}",
            )

            if sel_idx == CUSTOM_IDX:
                label       = st.text_input("Label", value=f"Custom {tag}", key=f"lbl_{tag}")
                solar       = st.number_input("Solar (kW)",    0.0, 30.0,    6.6,    0.5, key=f"sol_{tag}")
                bat         = st.number_input("Battery (kWh)", 0.0, 60.0,   13.5,    0.5, key=f"bat_{tag}")
                inv         = st.number_input("Inverter (kW)", 0.0, 15.0,    5.0,    0.5, key=f"inv_{tag}")
                cost        = st.number_input("Gross cost ($)",  0, 100_000, 15_000, 500, key=f"cost_{tag}")
                rebates_inc = False
            else:
                row          = solar_quotes.iloc[sel_idx]
                solar        = float(row["Solar_kW"])
                bat          = float(row["Battery_kWh"])
                inv          = float(row["Inverter_kW"])
                cost         = int(row["Cost_AUD"])
                rebates_inc  = bool(row.get("Rebates_Included", False))
                label        = _fmt(sel_idx)
                _def_s  = float(row.get("Shade_Summer", 0.60))
                _def_au = float(row.get("Shade_Autumn", 0.30))
                _def_w  = float(row.get("Shade_Winter", 0.15))
        else:
            # Fallback to manual entry if no quotes file
            label = st.text_input("Label", key=f"lbl_{tag}",
                                  value="Option A" if tag == "A" else "Option B")
            solar = st.number_input("Solar (kW)", 0.0, 30.0, 6.6, 0.5, key=f"sol_{tag}")
            bat   = st.number_input("Battery (kWh)", 0.0, 60.0, 16.0, 0.5, key=f"bat_{tag}")
            inv   = st.number_input("Inverter (kW)", 0.0, 15.0, 10.0, 0.5, key=f"inv_{tag}")
            cost  = st.number_input("Gross cost ($)", 0, 100_000, 20_000, 500, key=f"cost_{tag}")
            rebates_inc = False

        tariff = st.selectbox("Tariff", TARIFFS, key=f"tariff_{tag}")
        with st.expander("Shading factors", expanded=False):
            st.caption("Fraction of unshaded generation reaching the panels (1.0 = fully unshaded).")
            shade_s_q  = st.slider("Summer (Dec–Feb)",  0.0, 1.0, _def_s,  0.05, key=f"shade_s_{tag}")
            shade_au_q = st.slider("Autumn/Spring",     0.0, 1.0, _def_au, 0.05, key=f"shade_au_{tag}")
            shade_w_q  = st.slider("Winter (Jun–Aug)",  0.0, 1.0, _def_w,  0.05, key=f"shade_w_{tag}")
        with st.expander("Battery grid charging", expanded=False):
            st.caption("Midday Saver tariff only. When enabled, the battery tops up from the grid during the chosen window.")
            gc_enabled = st.toggle(
                "Charge battery from grid",
                value=True,
                key=f"gc_enabled_{tag}",
            )
            if gc_enabled:
                _gc_hours = [h / 2 for h in range(0, 49)]
                def _fmt_h(h):
                    return f"{int(h):02d}:{int(round(h % 1 * 60)):02d}"
                gc_window = st.select_slider(
                    "Grid charge window",
                    options=_gc_hours,
                    value=(9.0, 15.0),
                    format_func=_fmt_h,
                    key=f"gc_window_{tag}",
                )
                gc_start_q, gc_end_q = float(gc_window[0]), float(gc_window[1])
            else:
                gc_start_q, gc_end_q = 9.0, 15.0
        st.markdown("")
        return dict(label=label, solar=solar, bat=bat, inv=inv, cost=cost,
                    tariff=tariff, rebates_inc=rebates_inc,
                    shade=(shade_s_q, shade_au_q, shade_w_q),
                    grid_charge=gc_enabled, gc_start=gc_start_q, gc_end=gc_end_q)

    cfg_a = option_selector("A", default_idx=0)
    cfg_b = option_selector("B", default_idx=1)

    st.divider()
    st.subheader("3. STC price")
    _live_stc, _stc_source = fetch_stc_price()
    st.caption(
        f"Small-scale Technology Certificate spot price (inc GST). "
        f"Default loaded from: **{_stc_source}**. "
        f"Clearing-house max is $40 ex-GST (~$44 inc GST)."
    )
    stc_price = st.slider("STC spot price ($/STC)", 20.0, 45.0, _live_stc, 0.50)

    st.divider()
    st.subheader("4. Future cost assumptions")
    apply_inflation = st.toggle(
        "Apply electricity price inflation",
        value=True,
        help="When on, grid electricity prices escalate each year in the 20-year projections. "
             "Turn off to model flat (today's) prices for all 20 years.",
    )
    if apply_inflation:
        tariff_esc_pct = st.slider(
            "Electricity price inflation (%/yr)",
            min_value=0.0, max_value=10.0, value=4.0, step=0.5,
            help="Synergy A1 rate rose from ~12.5 ¢/kWh (2009) to ~31.7 ¢/kWh (2024) — "
                 "a 15-year compound average of ~6.4%/yr. "
                 "Use 2–3% for a conservative view, 6–7% for the historical WA average.",
        )
        tariff_esc = tariff_esc_pct / 100.0
    else:
        tariff_esc = 0.0

    discount_pct = st.slider(
        "Discount / opportunity-cost rate (%/yr)",
        min_value=0.0, max_value=20.0,
        value=round(OPPORTUNITY_RATE * 100, 1),
        step=0.5,
        help=(
            "The annual return you could earn by investing the upfront cost elsewhere "
            "(e.g. ~4–5% term deposit, ~8–10% diversified ETF). "
            "Use a nominal rate when inflation is ON, or a real rate (~5.5%) when OFF."
        ),
    )
    discount_rate = discount_pct / 100.0

    run = st.button("Run analysis", type="primary", disabled=(raw_df is None))

    # Warn when settings have changed since the last run
    _run_key = (
        getattr(uploaded, "name", ""), getattr(uploaded, "size", 0),
        cfg_a["solar"], cfg_a["bat"], cfg_a["inv"], cfg_a["cost"],
        cfg_a["tariff"], cfg_a["shade"],
        cfg_b["solar"], cfg_b["bat"], cfg_b["inv"], cfg_b["cost"],
        cfg_b["tariff"], cfg_b["shade"],
        stc_price,
        cfg_a["grid_charge"], cfg_a["gc_start"], cfg_a["gc_end"],
        cfg_b["grid_charge"], cfg_b["gc_start"], cfg_b["gc_end"],
        tariff_esc, discount_rate,
    )
    if run:
        st.session_state["_run_key"] = _run_key
    if (not run and "_run_key" in st.session_state
            and st.session_state["_run_key"] != _run_key):
        st.warning("⚠️ Settings changed — click **Run analysis** to update.")


# ─────────────────────────────────────────────────────────────────────────────
# Main area
# ─────────────────────────────────────────────────────────────────────────────

st.title("☀️ Solar + Battery Comparison")
st.markdown(
    "Compare two solar and battery quotes side-by-side using your **actual Synergy half-hourly consumption data**. "
    "The tool simulates how each system would perform against your real usage — including seasonal variation, "
    "battery cycling, DEBS export credits, and 20-year financial projections with panel and battery degradation."
)

if raw_df is None:
    st.info(
        "**To get started:** download your half-hourly meter data from "
        "[MyAccount](https://myaccount.synergy.net.au) (Manage Account → Usage & Payments → "
        "Download consumption data), then upload the file in the sidebar."
    )
    st.stop()

if not run:
    st.info("Configure your two options in the sidebar, then click **Run analysis**.")
    st.stop()

# ── Run simulations ──────────────────────────────────────────────────────────
raw_hash = hash(uploaded.getvalue())

with st.spinner("Simulating Option A…"):
    res_a = cached_simulate(
        raw_hash, cfg_a["solar"], cfg_a["bat"], cfg_a["inv"], cfg_a["tariff"],
        *cfg_a["shade"], cfg_a["grid_charge"], cfg_a["gc_start"], cfg_a["gc_end"],
        tariff_esc, raw_df,
    )
with st.spinner("Simulating Option B…"):
    res_b = cached_simulate(
        raw_hash, cfg_b["solar"], cfg_b["bat"], cfg_b["inv"], cfg_b["tariff"],
        *cfg_b["shade"], cfg_b["grid_charge"], cfg_b["gc_start"], cfg_b["gc_end"],
        tariff_esc, raw_df,
    )
with st.spinner("Computing payback…"):
    pb_a = cached_payback(
        raw_hash, cfg_a["solar"], cfg_a["bat"], cfg_a["inv"],
        cfg_a["cost"], cfg_a["tariff"], cfg_a["label"],
        *cfg_a["shade"], stc_price, cfg_a["rebates_inc"],
        cfg_a["grid_charge"], cfg_a["gc_start"], cfg_a["gc_end"],
        tariff_esc, discount_rate, raw_df,
    )
    pb_b = cached_payback(
        raw_hash, cfg_b["solar"], cfg_b["bat"], cfg_b["inv"],
        cfg_b["cost"], cfg_b["tariff"], cfg_b["label"],
        *cfg_b["shade"], stc_price, cfg_b["rebates_inc"],
        cfg_b["grid_charge"], cfg_b["gc_start"], cfg_b["gc_end"],
        tariff_esc, discount_rate, raw_df,
    )

with st.spinner("Computing status quo…"):
    res_base = cached_simulate_no_solar(raw_hash, cfg_a["tariff"], raw_df)
cfg_base = {"label": "Status Quo", "tariff": cfg_a["tariff"]}

bl_a = baseline(add_solar_shaded(raw_df, cfg_a["solar"], *cfg_a["shade"]), cfg_a["tariff"])
bl_b = baseline(add_solar_shaded(raw_df, cfg_b["solar"], *cfg_b["shade"]), cfg_b["tariff"])

# ── Metric cards ─────────────────────────────────────────────────────────────
st.subheader("Key Metrics")
st.caption(
    "Year-1 snapshot based on your uploaded meter data and current tariff rates. "
    "**Annual saving** = reduction in electricity bills vs no solar. "
    "**Self-sufficiency** = share of your load met by solar + battery (not from the grid). "
    "**NPV** = net value of the investment in today's dollars over 20 years (positive = better than the bank). "
    "**20-yr total return** = total savings ÷ net system cost (not annualised)."
)
col_base_m, col_a_m, col_b_m = st.columns(3)

with col_base_m:
    st.markdown("**Status Quo (no solar/battery)**")
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
        _tag = 'A' if colour == OPTION_COLOURS[0] else 'B'
        st.markdown(f"**Option {_tag}:** {cfg['label']}")
        _dr_pct = pb.get("discount_rate", OPPORTUNITY_RATE) * 100
        _dr     = pb.get("discount_rate", OPPORTUNITY_RATE)
        _npv10  = -pb["net"] + sum(pb["savings"][yr-1] / (1+_dr)**yr for yr in range(1, 11))

        r1, r2 = st.columns(2)
        r1.metric("Annual saving", f"${annual_saving:,.0f}")
        r2.metric("Nominal payback", f"{pb['pb_yr']} yr" if pb["pb_yr"] else ">20 yr")

        rn1, rn2 = st.columns(2)
        rn1.metric(f"NPV 10yr @ {_dr_pct:.0f}%", f"${_npv10:,.0f}")
        rn2.metric(f"NPV 20yr @ {_dr_pct:.0f}%", f"${pb['npv']:,.0f}")

        r4, r5, r6 = st.columns(3)
        r4.metric("20-yr saving", f"${pb['total_save']:,.0f}")
        r5.metric("20-yr total return", f"{pb['roi']:.0f}%")
        r6.metric("Self-sufficiency", f"{res['self_suf_pct']:.0f}%")

        r7, r8 = st.columns(2)
        r7.metric("Annual export", f"{res['annual_export_kwh']:,.0f} kWh")
        r8.metric("Net system cost", f"${pb['net']:,.0f}")

metric_col(col_a_m, res_a, pb_a, cfg_a, bl_a, OPTION_COLOURS[0])
metric_col(col_b_m, res_b, pb_b, cfg_b, bl_b, OPTION_COLOURS[1])

with st.expander("🔌 Inverter utilisation detail"):
    st.caption(
        "Shows how often the inverter is running at its rated limit during discharge (evening) "
        "and charge (midday). High saturation % means a larger inverter could improve performance. "
        "Low saturation means the current size is adequate — upsizing would make little difference."
    )
    inv_cols = st.columns(2)
    for col, res, cfg, colour, tag in [
        (inv_cols[0], res_a, cfg_a, OPTION_COLOURS[0], "A"),
        (inv_cols[1], res_b, cfg_b, OPTION_COLOURS[1], "B"),
    ]:
        inv_kw  = cfg["inv"]
        sol_kw  = cfg["solar"]
        df_inv  = res["df"]
        dis_lim = inv_kw * 0.5          # kWh per 30-min slot at rated power
        chg_lim = min(inv_kw, sol_kw) * 0.5

        # Slots where battery was actually doing something
        dis_active = df_inv["b_dis"] > 0.001
        chg_active = df_inv["b_chg"] > 0.001
        dis_sat = (df_inv.loc[dis_active, "b_dis"] >= dis_lim * 0.97).mean() * 100 if dis_active.any() else 0.0
        chg_sat = (df_inv.loc[chg_active, "b_chg"] >= chg_lim * 0.97).mean() * 100 if chg_active.any() else 0.0
        peak_dis_kw = df_inv["b_dis"].max() * 2   # kWh/slot → kW
        peak_chg_kw = df_inv["b_chg"].max() * 2

        with col:
            st.markdown(f"**Option {tag} — {inv_kw:.4g} kW inverter**")
            c1, c2 = st.columns(2)
            c1.metric("Discharge limit hit", f"{dis_sat:.0f}% of discharge slots",
                      help=f"Rated discharge: {inv_kw:.4g} kW · Peak recorded: {peak_dis_kw:.1f} kW")
            c2.metric("Charge limit hit", f"{chg_sat:.0f}% of charge slots",
                      help=f"Rated charge: {min(inv_kw, sol_kw):.4g} kW · Peak recorded: {peak_chg_kw:.1f} kW")
            st.caption(
                f"Rated discharge: **{inv_kw:.4g} kW** · Peak seen: **{peak_dis_kw:.1f} kW**  |  "
                f"Rated charge: **{min(inv_kw, sol_kw):.4g} kW** · Peak seen: **{peak_chg_kw:.1f} kW**"
            )

# ── Consumption heatmap ───────────────────────────────────────────────────────
st.divider()
st.subheader("Consumption Heatmap — Month × Hour of Day")
st.caption(
    "Average electricity consumption (kWh per hour) across all days in your data, "
    "grouped by month and hour of day. Darker cells indicate higher demand. "
    "Comparing this against the solar generation window (~7 am – 5 pm) shows how much "
    "of your load can be met directly by solar vs battery storage or grid import."
)
fig_hm = make_load_heatmap_fig(raw_df)
st.pyplot(fig_hm, use_container_width=True)
plt.close(fig_hm)

# ── Solar generation profiles ────────────────────────────────────────────────
st.divider()
st.subheader("Solar Generation Profiles by Season")
st.caption(
    "Theoretical mid-season output for each option (kWh per 30-min slot). "
    "Dashed = Option A unshaded (shade factor 1.0). "
    "Solid = with your shading factors applied. "
    "Shaded area shows the generation lost to shading."
)
fig_profiles = make_solar_profile_fig(cfg_a, cfg_b, cfg_a["shade"], cfg_b["shade"])
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

# ── Battery state of charge by season ────────────────────────────────────────
st.divider()
st.subheader("Battery State of Charge — Typical Day by Season")
st.caption(
    "Average battery SoC (kWh) over the day for each option and season. "
    "Dashed line = maximum usable capacity (nameplate × DoD). "
    "Green shading = super off-peak charge window (9am–3pm), red = peak discharge window (3–9pm)."
)
with st.spinner("Generating SoC charts…"):
    fig_soc = make_soc_seasonal_fig(res_a, res_b, cfg_a, cfg_b)
st.pyplot(fig_soc, use_container_width=True)
plt.close(fig_soc)

# ── Seasonal energy summary table ─────────────────────────────────────────────
st.divider()
st.subheader("Seasonal Energy Summary — Average Daily kWh")
st.caption(
    "Average daily totals across all days in each season from your meter data. "
    "Consumption is the same for both options (from your meter). "
    "Solar generation differs because the two options have different panel sizes."
)
tbl = make_seasonal_energy_table(res_a, res_b)
st.dataframe(
    tbl.style.format("{:.1f}", na_rep="—"),
    use_container_width=True,
)

# ── Payback / cashflow ────────────────────────────────────────────────────────
st.divider()
st.subheader(f"{ANALYSIS_YEARS}-Year Cumulative Cash-Flow")
st.caption(
    "Starts at −[net system cost] in Year 0 (the upfront investment). "
    "Each year the line rises by that year's bill savings. "
    "The point where a line crosses **$0** is the **nominal payback year** — when cumulative savings have recovered the purchase price. "
    "Future savings include electricity price escalation and export credit changes, but are shown in nominal (not inflation-adjusted) dollars."
)
with st.spinner("Generating payback chart…"):
    fig_pb = make_payback_fig(pb_a, pb_b, cfg_a, cfg_b)
st.pyplot(fig_pb, use_container_width=True)
plt.close(fig_pb)

st.subheader(f"{ANALYSIS_YEARS}-Year Cumulative Total Spend")
st.caption(
    "Total money out of pocket over 20 years: upfront system cost plus all ongoing electricity bills. "
    "The grey dashed line is the no-solar baseline — bills only, no upfront cost. "
    "Where a solar/battery line dips **below** the grey line, the system has paid for itself in total spend terms."
)
with st.spinner("Generating total spend chart…"):
    fig_ts = make_total_spend_fig(pb_a, pb_b, cfg_a, cfg_b)
st.pyplot(fig_ts, use_container_width=True)
plt.close(fig_ts)

st.subheader(f"{ANALYSIS_YEARS}-Year Cumulative Total Spend (PV-Adjusted)")
st.caption(
    f"Same as above but future bills and ongoing costs are discounted to today's dollars "
    f"at {discount_rate*100:.1f}%/yr — reflects the time value of money."
)
with st.spinner("Generating inflation-adjusted spend chart…"):
    fig_pv = make_pv_spend_fig(pb_a, pb_b, cfg_a, cfg_b)
st.pyplot(fig_pv, use_container_width=True)
plt.close(fig_pv)

# ── NPV Analysis ──────────────────────────────────────────────────────────────
st.divider()
_dr_pct_main = pb_a.get("discount_rate", OPPORTUNITY_RATE) * 100
st.subheader("NPV Analysis")
st.caption(
    f"**Net Present Value (NPV)** = total value of the investment in today's dollars **over {ANALYSIS_YEARS} years**, "
    f"discounting future savings at {_dr_pct_main:.1f}%/yr (your opportunity-cost rate). "
    f"A positive NPV means the system earns more than you'd get by investing the same money elsewhere at that rate. "
    f"**IRR** (Internal Rate of Return) is your effective annual return — the rate at which NPV = $0. "
    f"Compare it to alternatives: term deposit ~4–5%, diversified ETF ~8–10%. "
    f"The waterfall bars show the present value of each year's savings; the line is the running total."
)
with st.spinner("Generating DCF chart…"):
    fig_npv = make_npv_fig(pb_a, pb_b, cfg_a, cfg_b)
st.pyplot(fig_npv, use_container_width=True)
plt.close(fig_npv)

st.subheader("NPV Sensitivity to Discount Rate")
st.caption(
    f"Shows how NPV changes if you used a different discount rate. "
    f"The **dot** marks the IRR — where NPV = $0 and the line crosses the x-axis. "
    f"The **dashed vertical line** is your current rate ({_dr_pct_main:.1f}%). "
    f"If the dot (IRR) is **to the right** of the dashed line, the investment outperforms your hurdle rate — it's worth doing."
)
with st.spinner("Generating sensitivity chart…"):
    fig_npv_sens = make_npv_sensitivity_fig(pb_a, pb_b, cfg_a, cfg_b)
st.pyplot(fig_npv_sens, use_container_width=True)
plt.close(fig_npv_sens)

st.subheader("Payback & Financial Summary")
st.caption(
    f"Key modelling assumptions: "
    f"electricity price escalation {tariff_esc*100:.1f}%/yr · "
    f"DEBS export credits decline 5%/yr · "
    f"solar panel output degrades 0.5%/yr · "
    f"battery capacity degrades 2%/yr · "
    f"battery depth of discharge (DoD) {BAT_DOD*100:.0f}% · "
    f"system efficiency (inverter + wiring losses) {SYS_EFF*100:.0f}%."
)
pb_rows = []
for pb, cfg, tag in [(pb_a, cfg_a, "A"), (pb_b, cfg_b, "B")]:
    inc    = cfg.get("rebates_inc", False)
    _dr    = pb.get("discount_rate", OPPORTUNITY_RATE)
    _drpct = _dr * 100
    _npv10 = -pb["net"] + sum(pb["savings"][yr-1] / (1+_dr)**yr for yr in range(1, 11))
    pb_rows.append({
        "Option": f"{tag}: {cfg['label']}",
        "Tariff": cfg["tariff"],
        "Quoted price": f"${cfg['cost']:,.0f} ({'net — rebates already included' if inc else 'gross — rebates extra'})",
        "  – STC solar rebate": "included in quote" if inc else f"-${pb.get('stc', 0):,.0f}",
        "  – WA battery rebate": "included in quote" if inc else f"-${pb.get('state', 0):,.0f}",
        "  – Federal battery rebate": "included in quote" if inc else f"-${pb.get('fed', 0):,.0f}",
        "Purchase price (net)": f"${pb['net']:,.0f}",
        "Nominal payback": f"{pb['pb_yr']} yr" if pb["pb_yr"] else ">20 yr",
        f"Disc. payback ({_drpct:.1f}%)":
            f"{pb['pb_yr_disc']} yr" if pb.get("pb_yr_disc") else ">20 yr",
        f"NPV 10yr @ {_drpct:.1f}%": f"${_npv10:,.0f}",
        f"NPV 20yr @ {_drpct:.1f}%": f"${pb['npv']:,.0f}",
        "IRR": (lambda r: f"{r*100:.1f}%" if r is not None else "N/A")(compute_irr(pb)),
        "20-yr saving": f"${pb['total_save']:,.0f}",
        "20-yr total return": f"{pb['roi']:.0f}%",
        "Year-1 saving": f"${pb['yr1_save']:,.0f}",
        "Electricity price inflation": f"{tariff_esc*100:.1f}%/yr",
    })
st.dataframe(pd.DataFrame(pb_rows).set_index("Option"), use_container_width=True)

# ── Investment Efficiency Scatter ─────────────────────────────────────────────
st.divider()
st.subheader("Investment Efficiency")
st.caption(
    "NPV vs net system cost for all quotes in your quotes file, on both tariffs. "
    "★ = currently selected options A and B. "
    f"Above the dashed line: investment beats the {discount_rate*100:.1f}%/yr hurdle rate. "
    "Option A shading is used for all quotes to enable like-for-like comparison."
)

_eff_pbs: list = []
if solar_quotes is not None and len(solar_quotes) > 0:
    with st.spinner("Computing investment efficiency…"):
        for _, _row in solar_quotes.iterrows():
            for _tariff in ["A1 Flat", "Midday Saver"]:
                try:
                    _eff_pbs.append(cached_payback(
                        raw_hash,
                        float(_row["Solar_kW"]), float(_row["Battery_kWh"]),
                        float(_row["Inverter_kW"]), int(_row["Cost_AUD"]),
                        _tariff, str(_row["Vendor"]),
                        *cfg_a["shade"],
                        stc_price, bool(_row.get("Rebates_Included", False)),
                        cfg_a["grid_charge"], cfg_a["gc_start"], cfg_a["gc_end"],
                        tariff_esc, discount_rate,
                        raw_df,
                    ))
                except Exception:
                    pass

fig_eff = make_efficiency_scatter_fig(_eff_pbs, pb_a, pb_b, discount_rate)
if fig_eff:
    st.pyplot(fig_eff, use_container_width=True)
    plt.close(fig_eff)
else:
    st.info("No valid quotes to display. Check that your quotes file has cost data.")

# ── Sensitivity Analysis (Tornado) ────────────────────────────────────────────
st.divider()
st.subheader("Sensitivity Analysis")
st.caption(
    "NPV impact when each input is varied independently from its base case value. "
    "Green = upside, Red = downside. Widest bars = most impactful variables. "
    "Inputs varied: tariff escalation (−2 / +3 %/yr), discount rate (±4%), "
    "system cost (±15%), summer shading (±25%), winter shading (±30%)."
)
with st.spinner("Computing sensitivity analysis (first run only — results cached after)…"):
    for _pb, _cfg, _tag in [(pb_a, cfg_a, "A"), (pb_b, cfg_b, "B")]:
        st.markdown(
            f"**Option {_tag}: {str(_cfg.get('label',''))[:50]}  ·  {_cfg['tariff']}**"
        )
        fig_tornado = make_tornado_fig(
            _pb, raw_hash, _cfg, _cfg["shade"],
            stc_price, bool(_cfg.get("rebates_inc", False)),
            _cfg["grid_charge"], _cfg["gc_start"], _cfg["gc_end"],
            tariff_esc, discount_rate, raw_df,
        )
        st.pyplot(fig_tornado, use_container_width=True)
        plt.close(fig_tornado)

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

# ── PDF Report ────────────────────────────────────────────────────────────────
st.divider()
st.subheader("📄 Download Report")
st.caption(
    "Full report: cover page, all charts, financial summary table, and explanatory captions. "
    "Generated once per unique set of inputs — subsequent downloads are instant."
)
with st.spinner("Preparing PDF report…"):
    try:
        _pdf_bytes = cached_build_pdf(
            raw_hash,
            cfg_a["solar"], cfg_a["bat"], cfg_a["inv"], cfg_a["tariff"],
            cfg_b["solar"], cfg_b["bat"], cfg_b["inv"], cfg_b["tariff"],
            *cfg_a["shade"], *cfg_b["shade"],
            tariff_esc, discount_rate,
            raw_df, res_a, res_b, res_base, pb_a, pb_b, cfg_a, cfg_b,
        )
        st.download_button(
            "⬇️ Download PDF Report",
            data=_pdf_bytes,
            file_name="solar_battery_report.pdf",
            mime="application/pdf",
            type="primary",
        )
    except Exception as _e:
        st.error(f"PDF generation failed — {_e}")

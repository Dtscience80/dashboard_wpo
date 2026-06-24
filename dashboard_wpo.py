"""
Dashboard Analisis Mesin Diesel — Campuran Waste Plastic Oil (WPO)
Dataset: 23Juni_Kompilasi_Dataset_Perbaikan.xlsx (Sheet: Perbaikan)
EDA Komprehensif — Performa & Emisi Mesin Diesel dengan Variasi Campuran WPO
"""

import warnings
warnings.filterwarnings("ignore")

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import seaborn as sns
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# =========================================================================
# KONFIGURASI HALAMAN — WAJIB PERTAMA
# =========================================================================
st.set_page_config(
    page_title="Dashboard WPO — Analisis Mesin Diesel",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================================================================
# KONSTANTA VISUAL: Cartoon Brutalist
# =========================================================================
CARTOON_COLORS = [
    "#FF6B6B", "#4ECDC4", "#45B7D1", "#96CEB4", "#FFEAA7",
    "#DDA0DD", "#98D8C8", "#F7DC6F", "#BB8FCE", "#85C1E9",
    "#F0B27A", "#82E0AA", "#F1948A", "#AED6F1", "#D5F5E3",
]
plt.rcParams.update({
    "figure.facecolor": "white",
    "axes.edgecolor": "black",
    "axes.linewidth": 2,
    "axes.grid": False,
    "font.size": 12,
    "font.weight": "bold",
})
sns.set_style("white")

# =========================================================================
# LOAD DATA
# =========================================================================
@st.cache_data
def load_data():
    xls_path = "23Juni_Kompilasi_Dataset_Perbaikan.xlsx"
    df = pd.read_excel(xls_path, sheet_name="Perbaikan", engine="openpyxl")
    # Bersihkan kolom — hapus spasi ekstra
    df.columns = df.columns.str.strip()

    # Konversi kolom object yg mungkin masih pakai koma
    # Catatan: errors='ignore' dihapus di pandas 2.0+; gunakan 'coerce' lalu revert NaN
    for col in df.columns:
        if df[col].dtype == object:
            try:
                cleaned = df[col].astype(str).str.replace(",", "", regex=False)
                converted = pd.to_numeric(cleaned, errors="coerce")
                # Hanya ganti kolom jika konversi berhasil untuk mayoritas nilai (>50%)
                if converted.notna().mean() > 0.5:
                    df[col] = converted
            except Exception:
                pass

    # Pastikan numerik benar2 numerik
    numeric_cols = [
        "wpo_blend_pct", "engine_load_pct", "engine_speed_rpm",
        "engine_displacement_cc", "engine_torque_nm", "engine_power_kw",
        "brake_specific_fuel_consumption_kg_per_kwh", "brake_thermal_efficiency_pct",
        "nox_emission_ppm", "hc_emission_ppm", "co_emission_ppm", "co2_emission_ppm",
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Kolom kategorikal
    cat_cols = [
        "Merek/Model mesin", "Tipe mesin", "Jumlah silinder",
        "Sistem pendingin", "Bahan bakar uji", "Source paper", "DOI",
        "Keterangan beban uji",
    ]
    for col in cat_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    return df

df = load_data()

# =========================================================================
# KOLOM NUMERIK UTAMA
# =========================================================================
NUM_COLS = [
    "wpo_blend_pct", "engine_load_pct", "engine_speed_rpm",
    "engine_displacement_cc", "engine_torque_nm", "engine_power_kw",
    "brake_specific_fuel_consumption_kg_per_kwh", "brake_thermal_efficiency_pct",
    "nox_emission_ppm", "hc_emission_ppm", "co_emission_ppm", "co2_emission_ppm",
]

NUM_LABELS = {
    "wpo_blend_pct": "Campuran WPO (%)",
    "engine_load_pct": "Beban Mesin (%)",
    "engine_speed_rpm": "Kecepatan Mesin (RPM)",
    "engine_displacement_cc": "Volume Silinder (cc)",
    "engine_torque_nm": "Torsi (Nm)",
    "engine_power_kw": "Daya (kW)",
    "brake_specific_fuel_consumption_kg_per_kwh": "BSFC (kg/kWh)",
    "brake_thermal_efficiency_pct": "BTE (%)",
    "nox_emission_ppm": "NOx (ppm)",
    "hc_emission_ppm": "HC (ppm)",
    "co_emission_ppm": "CO (ppm)",
    "co2_emission_ppm": "CO₂ (ppm)",
}

CAT_COLS = [
    "Merek/Model mesin", "Tipe mesin", "Jumlah silinder",
    "Sistem pendingin", "Bahan bakar uji", "Source paper",
]

# =========================================================================
# SIDEBAR — FILTER
# =========================================================================
st.sidebar.title("🏭 Dashboard WPO")
st.sidebar.markdown("**Analisis Mesin Diesel — Campuran Waste Plastic Oil**")
st.sidebar.markdown("---")

# Filter WPO blend
wpo_options = sorted(df["wpo_blend_pct"].dropna().unique())
selected_wpo = st.sidebar.multiselect(
    "Campuran WPO (%)", wpo_options, default=wpo_options,
)
# Filter engine speed
rpm_options = sorted(df["engine_speed_rpm"].dropna().unique(), key=int)
selected_rpm = st.sidebar.multiselect(
    "Kecepatan Mesin (RPM)", rpm_options, default=rpm_options,
)
# Filter source paper
paper_col = "Source paper"
paper_options = df[paper_col].dropna().unique().tolist()
selected_papers = st.sidebar.multiselect(
    "Sumber Paper", paper_options, default=paper_options,
)

st.sidebar.markdown("---")
st.sidebar.markdown("**Tentang Dataset**")
st.sidebar.info(
    f"Dataset kompilasi performa & emisi mesin diesel\n"
    f"dengan campuran Waste Plastic Oil (WPO).\n\n"
    f"- **Total sampel**: {df.shape[0]}\n"
    f"- **Variabel numerik**: 12\n"
    f"- **Sumber paper**: {df[paper_col].nunique()}\n"
    f"- **Rentang WPO**: {int(df['wpo_blend_pct'].min())}–{int(df['wpo_blend_pct'].max())}%"
)

# Apply filters
mask = (
    df["wpo_blend_pct"].isin(selected_wpo)
    & df["engine_speed_rpm"].isin(selected_rpm)
    & df[paper_col].isin(selected_papers)
)
df_f = df[mask].copy()

# =========================================================================
# MAIN TITLE
# =========================================================================
st.title("🏭 Dashboard Analisis Mesin Diesel — Campuran WPO")
st.markdown(
    "**Waste Plastic Oil (WPO)** sebagai bahan bakar alternatif mesin diesel. "
    "Dashboard ini menyajikan eksplorasi data komprehensif tentang performa "
    "dan emisi mesin diesel dengan berbagai variasi campuran WPO."
)

# Summary metrics
tot = len(df_f)
st.subheader("📊 Ringkasan Data")
m1, m2, m3, m4, m5 = st.columns(5)
with m1:
    st.metric("Total Sampel", tot)
with m2:
    st.metric("Variasi WPO", df_f["wpo_blend_pct"].nunique())
with m3:
    st.metric("Rentang RPM", f"{int(df_f['engine_speed_rpm'].min())}–{int(df_f['engine_speed_rpm'].max())}")
with m4:
    st.metric("Sumber Paper", df_f[paper_col].nunique())
with m5:
    st.metric("Rerata BTE", f"{df_f['brake_thermal_efficiency_pct'].mean():.1f}%" if "brake_thermal_efficiency_pct" in df_f.columns else "—")

# =========================================================================
# TABS
# =========================================================================
tabs = st.tabs([
    "📋 Ringkasan Dataset",
    "📊 Distribusi & Univariat",
    "🔥 Korelasi",
    "🛢️ Pengaruh Campuran WPO",
    "🏭 Emisi vs Performa",
    "📄 Perbandingan Paper",
    "🔍 Eksplorasi Mandiri",
])

# =====================================================================
# TAB 1: RINGKASAN DATASET
# =====================================================================
with tabs[0]:
    c1, c2 = st.columns([1, 1])

    with c1:
        st.subheader("📋 Overview Dataset")
        # Info dataset
        info_data = {
            "Jumlah sampel": [df_f.shape[0]],
            "Jumlah variabel": [df_f.shape[1]],
            "Variabel numerik": [len(df_f.select_dtypes(include=[np.number]).columns)],
            "Variabel kategorikal": [len(df_f.select_dtypes(include=["object", "category"]).columns)],
            "Sumber paper": [df_f[paper_col].nunique()],
        }
        st.table(pd.DataFrame(info_data, index=["Nilai"]))

        st.subheader("📄 5 Sampel Data")
        display_cols = [c for c in NUM_COLS if c in df_f.columns] + [paper_col]
        display_cols = [c for c in display_cols if c in df_f.columns]
        st.dataframe(
            df_f[display_cols].head(5).style.format(
                {c: "{:.2f}" for c in NUM_COLS if c in df_f.columns}
            ),
            use_container_width=True,
            height=200,
        )

    with c2:
        st.subheader("❌ Missing Values")
        missing = df_f.isna().sum()
        missing = missing[missing > 0]
        if len(missing) > 0:
            fig, ax = plt.subplots(figsize=(6, 4))
            colors_miss = ["#FF6B6B", "#FFEAA7", "#4ECDC4"][:len(missing)]
            bars = ax.barh(
                missing.index, missing.values,
                color=colors_miss, edgecolor="black", linewidth=2,
            )
            for bar in bars:
                w = bar.get_width()
                ax.text(w + 0.1, bar.get_y() + bar.get_height() / 2,
                         f"{int(w)}", va="center", fontweight="bold")
            ax.set_xlabel("Jumlah Missing")
            ax.set_title("Missing Values per Kolom")
            ax.margins(x=0.3)
            st.pyplot(fig)
            plt.close()
        else:
            st.success("✅ Tidak ada missing values dalam dataset terfilter!")
            # Tampilkan missing utuh
            fig, ax = plt.subplots(figsize=(6, 2))
            ax.text(0.5, 0.5, "✅ Tidak ada missing values", ha="center", va="center",
                    fontsize=16, fontweight="bold")
            ax.axis("off")
            st.pyplot(fig)
            plt.close()

    st.subheader("📊 Statistik Deskriptif")
    desc = df_f[NUM_COLS].describe().T
    desc.columns = ["Count", "Mean", "Std", "Min", "25%", "50%", "75%", "Max"]
    desc.index = [NUM_LABELS.get(c, c) for c in desc.index]
    st.dataframe(desc.style.format("{:.2f}"), use_container_width=True)

# =====================================================================
# TAB 2: DISTRIBUSI & UNIVARIAT
# =====================================================================
with tabs[1]:
    st.subheader("📊 Distribusi Setiap Variabel Numerik")

    # Row 1: Engine parameters
    row1_cols = ["wpo_blend_pct", "engine_load_pct", "engine_speed_rpm", "engine_displacement_cc"]
    act_cols1 = [c for c in row1_cols if c in df_f.columns]
    cols = st.columns(len(act_cols1))
    for i, col_name in enumerate(act_cols1):
        with cols[i]:
            fig, ax = plt.subplots(figsize=(4, 3))
            data = df_f[col_name].dropna()
            ax.hist(data, bins=12, color=CARTOON_COLORS[i % len(CARTOON_COLORS)],
                    edgecolor="black", linewidth=2, density=True, alpha=0.85)
            # KDE
            from scipy.stats import gaussian_kde
            try:
                kde = gaussian_kde(data)
                x_kde = np.linspace(data.min(), data.max(), 200)
                ax.plot(x_kde, kde(x_kde), color="black", linewidth=2.5, linestyle="--")
            except Exception:
                pass
            ax.set_xlabel(NUM_LABELS.get(col_name, col_name), fontsize=10)
            ax.set_ylabel("Density", fontsize=9)
            ax.set_title(NUM_LABELS.get(col_name, col_name), fontsize=11, fontweight="bold")
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            st.pyplot(fig)
            plt.close()

    # Row 2: Performance
    row2_cols = ["engine_torque_nm", "engine_power_kw", "brake_specific_fuel_consumption_kg_per_kwh", "brake_thermal_efficiency_pct"]
    act_cols2 = [c for c in row2_cols if c in df_f.columns]
    cols2 = st.columns(len(act_cols2))
    for i, col_name in enumerate(act_cols2):
        with cols2[i]:
            fig, ax = plt.subplots(figsize=(4, 3))
            data = df_f[col_name].dropna()
            ax.hist(data, bins=12, color=CARTOON_COLORS[(i + 4) % len(CARTOON_COLORS)],
                    edgecolor="black", linewidth=2, density=True, alpha=0.85)
            try:
                kde = gaussian_kde(data)
                x_kde = np.linspace(data.min(), data.max(), 200)
                ax.plot(x_kde, kde(x_kde), color="black", linewidth=2.5, linestyle="--")
            except Exception:
                pass
            ax.set_xlabel(NUM_LABELS.get(col_name, col_name), fontsize=10)
            ax.set_ylabel("Density", fontsize=9)
            ax.set_title(NUM_LABELS.get(col_name, col_name), fontsize=11, fontweight="bold")
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            st.pyplot(fig)
            plt.close()

    # Row 3: Emissions
    row3_cols = ["nox_emission_ppm", "hc_emission_ppm", "co_emission_ppm", "co2_emission_ppm"]
    act_cols3 = [c for c in row3_cols if c in df_f.columns]
    cols3 = st.columns(len(act_cols3))
    for i, col_name in enumerate(act_cols3):
        with cols3[i]:
            fig, ax = plt.subplots(figsize=(4, 3))
            data = df_f[col_name].dropna()
            ax.hist(data, bins=12, color=CARTOON_COLORS[(i + 8) % len(CARTOON_COLORS)],
                    edgecolor="black", linewidth=2, density=True, alpha=0.85)
            try:
                kde = gaussian_kde(data)
                x_kde = np.linspace(data.min(), data.max(), 200)
                ax.plot(x_kde, kde(x_kde), color="black", linewidth=2.5, linestyle="--")
            except Exception:
                pass
            ax.set_xlabel(NUM_LABELS.get(col_name, col_name), fontsize=10)
            ax.set_ylabel("Density", fontsize=9)
            ax.set_title(NUM_LABELS.get(col_name, col_name), fontsize=11, fontweight="bold")
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            st.pyplot(fig)
            plt.close()

    # Boxplot per WPO blend
    st.subheader("📦 Boxplot per Variabel — Dikelompokkan berdasarkan Campuran WPO")
    box_var = st.selectbox(
        "Pilih variabel untuk boxplot:",
        [c for c in NUM_COLS if c in df_f.columns and c != "wpo_blend_pct"],
        format_func=lambda x: NUM_LABELS.get(x, x),
    )
    if box_var:
        fig, ax = plt.subplots(figsize=(10, 5))
        wpo_groups = sorted(df_f["wpo_blend_pct"].dropna().unique())
        data_groups = [df_f[df_f["wpo_blend_pct"] == w][box_var].dropna().values for w in wpo_groups]
        bp = ax.boxplot(
            data_groups,
            patch_artist=True, widths=0.5,
            boxprops=dict(edgecolor="black", linewidth=2),
            whiskerprops=dict(color="black", linewidth=2),
            capprops=dict(color="black", linewidth=2),
            medianprops=dict(color="black", linewidth=2, linestyle="-"),
            flierprops=dict(marker="o", markerfacecolor="#FF6B6B", markersize=6, markeredgecolor="black"),
        )
        ax.set_xticklabels([f"{int(w)}%" for w in wpo_groups])
        for patch, color in zip(bp["boxes"], CARTOON_COLORS):
            patch.set_facecolor(color)
        ax.set_xlabel("Campuran WPO (%)", fontsize=12, fontweight="bold")
        ax.set_ylabel(NUM_LABELS.get(box_var, box_var), fontsize=12, fontweight="bold")
        ax.set_title(f"Distribusi {NUM_LABELS.get(box_var, box_var)} per Campuran WPO",
                     fontsize=14, fontweight="bold")
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        st.pyplot(fig)
        plt.close()

# =====================================================================
# TAB 3: KORELASI
# =====================================================================
with tabs[2]:
    st.subheader("🔥 Matriks Korelasi Antar Variabel Numerik")

    corr_cols = [c for c in NUM_COLS if c in df_f.columns]
    corr_data = df_f[corr_cols].dropna()
    corr = corr_data.corr()

    fig, ax = plt.subplots(figsize=(12, 9))
    mask = np.triu(np.ones_like(corr, dtype=bool), k=1)
    cmap = sns.diverging_palette(250, 15, s=75, l=40, n=12, center="light", as_cmap=True)
    sns.heatmap(
        corr, mask=mask, annot=True, fmt=".2f", cmap=cmap,
        square=True, linewidths=2, linecolor="black",
        cbar_kws={"shrink": 0.75, "label": "Korelasi"},
        ax=ax,
        annot_kws={"fontsize": 8, "fontweight": "bold"},
    )
    labels = [NUM_LABELS.get(c, c) for c in corr.columns]
    ax.set_xticklabels(labels, rotation=45, ha="right", fontsize=9)
    ax.set_yticklabels(labels, rotation=0, fontsize=9)
    ax.set_title("Matriks Korelasi Pearson", fontsize=16, fontweight="bold", pad=20)
    st.pyplot(fig)
    plt.close()

    st.subheader("📈 Pairplot — Variabel Kunci")
    pair_cols = [
        "wpo_blend_pct", "engine_load_pct", "engine_torque_nm",
        "engine_power_kw", "brake_specific_fuel_consumption_kg_per_kwh",
        "brake_thermal_efficiency_pct", "nox_emission_ppm",
        "hc_emission_ppm", "co_emission_ppm",
    ]
    pair_cols_act = [c for c in pair_cols if c in df_f.columns]
    pair_data = df_f[pair_cols_act].dropna().copy()
    # Ekstrak nilai WPO sebagai array (hindari index-alignment issue di pandas 2+/3+)
    wpo_array = pair_data["wpo_blend_pct"].values
    pair_data = pair_data.drop(columns=["wpo_blend_pct"])
    pair_data["WPO"] = [f"{int(x)}%" for x in wpo_array]
    pair_vars = [c for c in pair_cols_act if c != "wpo_blend_pct"]

    fig = sns.pairplot(
        pair_data, vars=pair_vars, hue="WPO",
        palette=CARTOON_COLORS[:pair_data["WPO"].nunique()],
        diag_kind="kde", plot_kws={"alpha": 0.6, "s": 30, "edgecolor": "black", "linewidth": 0.5},
        diag_kws={"fill": True},
    )
    fig.fig.suptitle("Pairplot — Hubungan Antar Variabel Kunci", fontsize=16, fontweight="bold", y=1.02)
    st.pyplot(fig)
    plt.close()

# =====================================================================
# TAB 4: PENGARUH CAMPURAN WPO
# =====================================================================
with tabs[3]:
    st.subheader("🛢️ Pengaruh Campuran WPO terhadap Performa & Emisi")

    # WPO vs Emissions — plotly line
    st.markdown("### Emisi vs Campuran WPO")

    emit_cols = [c for c in ["nox_emission_ppm", "hc_emission_ppm", "co_emission_ppm", "co2_emission_ppm"]
                 if c in df_f.columns]

    # Agregasi rata2 per WPO blend
    agg_emit = df_f.groupby("wpo_blend_pct")[emit_cols].mean().reset_index()

    fig = make_subplots(
        rows=2, cols=2,
        subplot_titles=[NUM_LABELS.get(c, c) for c in emit_cols],
        vertical_spacing=0.12, horizontal_spacing=0.08,
    )
    colors_line = ["#FF6B6B", "#4ECDC4", "#45B7D1", "#FFEAA7"]
    for i, col in enumerate(emit_cols):
        r = i // 2 + 1
        c = i % 2 + 1
        fig.add_trace(
            go.Scatter(
                x=agg_emit["wpo_blend_pct"], y=agg_emit[col],
                mode="lines+markers",
                name=NUM_LABELS.get(col, col),
                line=dict(color=colors_line[i], width=3),
                marker=dict(size=10, color=colors_line[i], line=dict(color="black", width=2)),
            ),
            row=r, col=c,
        )
        fig.update_xaxes(title_text="WPO (%)", row=r, col=c)
        fig.update_yaxes(title_text=NUM_LABELS.get(col, col).split("(")[0].strip(), row=r, col=c)

    fig.update_layout(
        height=500,
        title_text="Rata-rata Emisi vs Campuran WPO",
        font=dict(size=12),
        showlegend=False,
    )
    st.plotly_chart(fig, use_container_width=True)

    # WPO vs Performance — plotly line
    st.markdown("### Performa vs Campuran WPO")

    perf_cols = [c for c in ["engine_torque_nm", "engine_power_kw",
                              "brake_specific_fuel_consumption_kg_per_kwh", "brake_thermal_efficiency_pct"]
                 if c in df_f.columns]

    agg_perf = df_f.groupby("wpo_blend_pct")[perf_cols].mean().reset_index()

    fig2 = make_subplots(
        rows=2, cols=2,
        subplot_titles=[NUM_LABELS.get(c, c) for c in perf_cols],
        vertical_spacing=0.12, horizontal_spacing=0.08,
    )
    colors_line2 = ["#DDA0DD", "#98D8C8", "#F7DC6F", "#BB8FCE"]
    for i, col in enumerate(perf_cols):
        r = i // 2 + 1
        c = i % 2 + 1
        fig2.add_trace(
            go.Scatter(
                x=agg_perf["wpo_blend_pct"], y=agg_perf[col],
                mode="lines+markers",
                name=NUM_LABELS.get(col, col),
                line=dict(color=colors_line2[i], width=3),
                marker=dict(size=10, color=colors_line2[i], line=dict(color="black", width=2)),
            ),
            row=r, col=c,
        )
        fig2.update_xaxes(title_text="WPO (%)", row=r, col=c)
        fig2.update_yaxes(title_text=NUM_LABELS.get(col, col).split("(")[0].strip(), row=r, col=c)

    fig2.update_layout(
        height=500,
        title_text="Rata-rata Performa vs Campuran WPO",
        font=dict(size=12),
        showlegend=False,
    )
    st.plotly_chart(fig2, use_container_width=True)

    # WPO split by engine speed
    st.markdown("### Pengaruh WPO — Dibedakan berdasarkan Kecepatan Mesin")
    rpm_choices = sorted(df_f["engine_speed_rpm"].dropna().unique(), key=int)
    split_var = st.selectbox(
        "Pilih variabel dependen:",
        [c for c in perf_cols + emit_cols],
        format_func=lambda x: NUM_LABELS.get(x, x),
        key="split_var",
    )
    if split_var and len(rpm_choices) > 0:
        fig3, ax3 = plt.subplots(figsize=(10, 5))
        for i, rpm in enumerate(rpm_choices):
            sub = df_f[df_f["engine_speed_rpm"] == rpm].groupby("wpo_blend_pct")[split_var].mean().reset_index()
            ax3.plot(
                sub["wpo_blend_pct"], sub[split_var],
                marker="o", markersize=8, linewidth=2.5,
                color=CARTOON_COLORS[i % len(CARTOON_COLORS)],
                label=f"{int(rpm)} RPM",
                markerfacecolor=CARTOON_COLORS[i % len(CARTOON_COLORS)],
                markeredgecolor="black", markeredgewidth=2,
            )
        ax3.set_xlabel("Campuran WPO (%)", fontweight="bold")
        ax3.set_ylabel(NUM_LABELS.get(split_var, split_var), fontweight="bold")
        ax3.set_title(f"{NUM_LABELS.get(split_var, split_var)} vs WPO — per Kecepatan Mesin",
                      fontsize=14, fontweight="bold")
        ax3.legend(frameon=True, edgecolor="black", facecolor="white", fontsize=11)
        ax3.spines["top"].set_visible(False)
        ax3.spines["right"].set_visible(False)
        st.pyplot(fig3)
        plt.close()

    # Bar chart — WPO vs key metrics
    st.markdown("### Perbandingan Rerata — BTE dan BSFC per Campuran WPO")
    bar_cols = [c for c in ["brake_thermal_efficiency_pct", "brake_specific_fuel_consumption_kg_per_kwh"]
                if c in df_f.columns]
    if len(bar_cols) == 2:
        agg_bar = df_f.groupby("wpo_blend_pct")[bar_cols].mean().reset_index()
        fig4, axes = plt.subplots(1, 2, figsize=(12, 5))
        for idx, col in enumerate(bar_cols):
            ax = axes[idx]
            x_pos = range(len(agg_bar))
            ax.bar(
                x_pos, agg_bar[col],
                color=[CARTOON_COLORS[i % len(CARTOON_COLORS)] for i in range(len(agg_bar))],
                edgecolor="black", linewidth=2, width=0.6,
            )
            ax.set_xticks(x_pos)
            ax.set_xticklabels([f"{int(w)}%" for w in agg_bar["wpo_blend_pct"]], fontweight="bold")
            ax.set_xlabel("Campuran WPO (%)", fontweight="bold")
            ax.set_ylabel(NUM_LABELS.get(col, col).split("(")[0].strip(), fontweight="bold")
            ax.set_title(NUM_LABELS.get(col, col), fontsize=13, fontweight="bold")
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            # Nilai di atas bar
            for j, v in enumerate(agg_bar[col]):
                ax.text(j, v + 0.005 * max(agg_bar[col]), f"{v:.2f}",
                        ha="center", va="bottom", fontweight="bold", fontsize=10)
        fig4.suptitle("Perbandingan Rerata per Campuran WPO", fontsize=15, fontweight="bold")
        fig4.tight_layout()
        st.pyplot(fig4)
        plt.close()

# =====================================================================
# TAB 5: EMISI VS PERFORMA
# =====================================================================
with tabs[4]:
    st.subheader("🏭 Hubungan antara Emisi dan Performa Mesin")

    col_emit = [c for c in ["nox_emission_ppm", "hc_emission_ppm", "co_emission_ppm", "co2_emission_ppm"]
                if c in df_f.columns]
    col_perf = [c for c in ["engine_torque_nm", "engine_power_kw",
                             "brake_specific_fuel_consumption_kg_per_kwh", "brake_thermal_efficiency_pct"]
                if c in df_f.columns]

    sc_x = st.selectbox("Sumbu X:", col_emit + col_perf, index=0,
                        format_func=lambda x: NUM_LABELS.get(x, x), key="sc_x")
    sc_y = st.selectbox("Sumbu Y:", col_emit + col_perf, index=1,
                        format_func=lambda x: NUM_LABELS.get(x, x), key="sc_y")
    sc_color = st.selectbox("Warna berdasarkan:", ["wpo_blend_pct"] + CAT_COLS,
                            format_func=lambda x: NUM_LABELS.get(x, x) if x in NUM_LABELS else x,
                            key="sc_color")

    sc_data = df_f[[sc_x, sc_y, sc_color]].dropna().copy()
    if sc_color == "wpo_blend_pct":
        fig5 = px.scatter(
            sc_data, x=sc_x, y=sc_y, color=sc_color,
            color_continuous_scale="viridis",
            labels={sc_x: NUM_LABELS.get(sc_x, sc_x), sc_y: NUM_LABELS.get(sc_y, sc_y)},
            title=f"{NUM_LABELS.get(sc_y, sc_y)} vs {NUM_LABELS.get(sc_x, sc_x)}",
        )
        fig5.update_traces(
            marker=dict(size=10, line=dict(width=2, color="black")),
        )
    else:
        fig5 = px.scatter(
            sc_data, x=sc_x, y=sc_y, color=sc_color,
            color_discrete_sequence=CARTOON_COLORS,
            labels={sc_x: NUM_LABELS.get(sc_x, sc_x), sc_y: NUM_LABELS.get(sc_y, sc_y)},
            title=f"{NUM_LABELS.get(sc_y, sc_y)} vs {NUM_LABELS.get(sc_x, sc_x)}",
        )
        fig5.update_traces(
            marker=dict(size=10, line=dict(width=2, color="black")),
        )

    fig5.update_layout(
        font=dict(size=13),
        xaxis=dict(title_font=dict(size=13, color="black")),
        yaxis=dict(title_font=dict(size=13, color="black")),
    )
    st.plotly_chart(fig5, use_container_width=True)

    # Scatter matrix (plotly)
    st.subheader("🔷 Scatter Matrix — Emisi vs Performa")
    sm_cols = st.multiselect(
        "Pilih variabel untuk scatter matrix:",
        col_emit + col_perf + ["engine_load_pct"],
        default=(col_emit + col_perf[:2])[:5],
        format_func=lambda x: NUM_LABELS.get(x, x),
    )
    if len(sm_cols) >= 2:
        sm_data = df_f[sm_cols].dropna()
        fig6 = px.scatter_matrix(
            sm_data,
            labels={c: NUM_LABELS.get(c, c).split("(")[0].strip() for c in sm_cols},
            dimensions=sm_cols,
            opacity=0.6,
        )
        fig6.update_traces(
            marker=dict(size=6, color="#4ECDC4", line=dict(width=1, color="black")),
            diagonal_visible=False,
        )
        fig6.update_layout(
            height=700,
            title="Scatter Matrix — Hubungan Antar Variabel",
            font=dict(size=11),
        )
        st.plotly_chart(fig6, use_container_width=True)

    # Correlation bar chart
    st.subheader("📊 Korelasi Emisi vs Performa")
    if len(col_emit) > 0 and len(col_perf) > 0:
        corr_ep = df_f[col_emit + col_perf].dropna().corr().loc[col_emit, col_perf]
        fig7, ax7 = plt.subplots(figsize=(10, 6))
        sns.heatmap(
            corr_ep, annot=True, fmt=".2f", cmap="RdBu_r",
            center=0, vmin=-1, vmax=1,
            linewidths=2, linecolor="black",
            cbar_kws={"shrink": 0.8, "label": "Korelasi"},
            ax=ax7,
            annot_kws={"fontsize": 10, "fontweight": "bold"},
        )
        emit_labels = [NUM_LABELS.get(c, c).split("(")[0].strip() for c in col_emit]
        perf_labels = [NUM_LABELS.get(c, c).split("(")[0].strip() for c in col_perf]
        ax7.set_xticklabels(perf_labels, rotation=30, ha="right", fontsize=10)
        ax7.set_yticklabels(emit_labels, rotation=0, fontsize=10)
        ax7.set_title("Korelasi: Emisi vs Parameter Performa", fontsize=14, fontweight="bold")
        st.pyplot(fig7)
        plt.close()

# =====================================================================
# TAB 6: PERBANDINGAN PAPER
# =====================================================================
with tabs[5]:
    st.subheader("📄 Perbandingan Antar Sumber Paper")

    paper_counts = df_f[paper_col].value_counts()
    st.markdown(f"**Jumlah paper unik**: {len(paper_counts)}")

    # Data per paper
    paper_agg = df_f.groupby(paper_col)[
        [c for c in NUM_COLS if c in df_f.columns and c != "wpo_blend_pct"]
    ].mean()

    # BTE comparison
    if "brake_thermal_efficiency_pct" in paper_agg.columns:
        st.markdown("### ⚡ Rata-rata BTE per Paper")
        bte_data = paper_agg["brake_thermal_efficiency_pct"].reset_index().dropna()
        fig8, ax8 = plt.subplots(figsize=(12, 5))
        short_labels = [s[:60] + "..." if len(s) > 60 else s for s in bte_data[paper_col]]
        bars = ax8.barh(
            range(len(bte_data)), bte_data["brake_thermal_efficiency_pct"],
            color=[CARTOON_COLORS[i % len(CARTOON_COLORS)] for i in range(len(bte_data))],
            edgecolor="black", linewidth=2,
        )
        ax8.set_yticks(range(len(bte_data)))
        ax8.set_yticklabels(short_labels, fontsize=8)
        ax8.set_xlabel("BTE Rata-rata (%)", fontweight="bold")
        ax8.set_title("Perbandingan Brake Thermal Efficiency (BTE) antar Paper", fontsize=14, fontweight="bold")
        ax8.invert_yaxis()
        ax8.spines["top"].set_visible(False)
        ax8.spines["right"].set_visible(False)
        for bar, val in zip(bars, bte_data["brake_thermal_efficiency_pct"]):
            ax8.text(val + 0.2, bar.get_y() + bar.get_height() / 2,
                     f"{val:.1f}%", va="center", fontweight="bold", fontsize=9)
        st.pyplot(fig8)
        plt.close()

    # Multi-metric comparison — plotly
    st.markdown("### 📊 Perbandingan Multi-Metrik antar Paper")
    compare_metrics = st.multiselect(
        "Pilih metrik yang dibandingkan:",
        [c for c in NUM_COLS if c in df_f.columns and c != "wpo_blend_pct"],
        default=["brake_thermal_efficiency_pct", "nox_emission_ppm", "brake_specific_fuel_consumption_kg_per_kwh"],
        format_func=lambda x: NUM_LABELS.get(x, x),
        key="compare_metrics",
    )
    if compare_metrics:
        # Normalize for radar chart
        radar_data = df_f.groupby(paper_col)[compare_metrics].mean().dropna()
        if len(radar_data) > 0:
            fig9 = go.Figure()
            for i, (paper_name, row) in enumerate(radar_data.iterrows()):
                short_name = paper_name[:40] + "..." if len(paper_name) > 40 else paper_name
                fig9.add_trace(go.Scatterpolar(
                    r=row.values,
                    theta=[NUM_LABELS.get(c, c).split("(")[0].strip() for c in compare_metrics],
                    fill="toself",
                    name=short_name,
                    line=dict(color=CARTOON_COLORS[i % len(CARTOON_COLORS)], width=2),
                ))
            fig9.update_layout(
                polar=dict(
                    radialaxis=dict(visible=True),
                ),
                title="Perbandingan Multi-Metrik antar Paper (Radar Chart)",
                font=dict(size=11),
                height=600,
            )
            st.plotly_chart(fig9, use_container_width=True)

    # Table per paper
    st.markdown("### 📋 Tabel Rata-rata per Paper")
    display_metrics = st.multiselect(
        "Tampilkan metrik:",
        [c for c in NUM_COLS if c in df_f.columns],
        default=NUM_COLS[:7],
        format_func=lambda x: NUM_LABELS.get(x, x),
        key="display_metrics",
    )
    if display_metrics:
        paper_table = df_f.groupby(paper_col)[display_metrics].mean().round(2)
        paper_table.index = [s[:60] + "..." if len(s) > 60 else s for s in paper_table.index]
        st.dataframe(paper_table.style.background_gradient(cmap="YlOrRd"), use_container_width=True)

# =====================================================================
# TAB 7: EKSPLORASI MANDIRI
# =====================================================================
with tabs[6]:
    st.subheader("🔍 Eksplorasi Mandiri — Scatter Plot Interaktif")

    all_numeric = [c for c in NUM_COLS if c in df_f.columns]
    all_cats = [c for c in CAT_COLS if c in df_f.columns]

    x_axis = st.selectbox("Sumbu X:", all_numeric, index=0,
                          format_func=lambda x: NUM_LABELS.get(x, x), key="expl_x")
    y_axis = st.selectbox("Sumbu Y:", all_numeric, index=1,
                          format_func=lambda x: NUM_LABELS.get(x, x), key="expl_y")
    color_by = st.selectbox("Warna:", all_cats + all_numeric,
                            format_func=lambda x: NUM_LABELS.get(x, x) if x in NUM_LABELS else x,
                            key="expl_color")
    size_by = st.selectbox("Ukuran titik:", ["(sama)"] + all_numeric,
                           format_func=lambda x: NUM_LABELS.get(x, x) if x in NUM_LABELS else "(sama)",
                           key="expl_size")

    hover_extra = [c for c in NUM_COLS[:4] if c in df_f.columns] + [paper_col]

    # Kumpulkan semua kolom yang dibutuhkan plot_data (hindari duplikat)
    base_cols = [x_axis, y_axis, color_by] + ([size_by] if size_by != "(sama)" else [])
    extra_cols = [c for c in hover_extra if c in df_f.columns and c not in base_cols]
    all_plot_cols = list(dict.fromkeys(base_cols + extra_cols))  # deduplicate, jaga urutan
    plot_data = df_f[all_plot_cols].dropna()

    # hover_data hanya boleh referensikan kolom yang ADA di plot_data
    hover_data = {c: True for c in hover_extra if c in plot_data.columns}

    if color_by in all_numeric:
        fig10 = px.scatter(
            plot_data, x=x_axis, y=y_axis, color=color_by,
            size=size_by if size_by != "(sama)" else None,
            hover_data=hover_data,
            color_continuous_scale="viridis",
            labels={x_axis: NUM_LABELS.get(x_axis, x_axis),
                     y_axis: NUM_LABELS.get(y_axis, y_axis),
                     color_by: NUM_LABELS.get(color_by, color_by)},
            title=f"{NUM_LABELS.get(y_axis, y_axis)} vs {NUM_LABELS.get(x_axis, x_axis)}",
        )
    else:
        fig10 = px.scatter(
            plot_data, x=x_axis, y=y_axis, color=color_by,
            size=size_by if size_by != "(sama)" else None,
            hover_data=hover_data,
            color_discrete_sequence=CARTOON_COLORS,
            labels={x_axis: NUM_LABELS.get(x_axis, x_axis),
                     y_axis: NUM_LABELS.get(y_axis, y_axis),
                     color_by: color_by},
            title=f"{NUM_LABELS.get(y_axis, y_axis)} vs {NUM_LABELS.get(x_axis, x_axis)}",
        )

    fig10.update_traces(
        marker=dict(line=dict(width=1.5, color="black")),
    )
    fig10.update_layout(
        height=600,
        font=dict(size=13),
        xaxis=dict(title_font=dict(size=14, color="black")),
        yaxis=dict(title_font=dict(size=14, color="black")),
    )
    st.plotly_chart(fig10, use_container_width=True)

    # Data table
    st.markdown("### 📋 Data Tersaring")
    show_cols = st.multiselect(
        "Pilih kolom untuk ditampilkan:",
        df_f.columns.tolist(),
        default=[c for c in NUM_COLS if c in df_f.columns] + [paper_col],
        format_func=lambda x: NUM_LABELS.get(x, x) if x in NUM_LABELS else x,
        key="show_cols",
    )
    if show_cols:
        st.dataframe(
            df_f[show_cols].style.format(
                {c: "{:.2f}" for c in all_numeric if c in show_cols}
            ),
            use_container_width=True,
            height=400,
        )

# =========================================================================
# FOOTER
# =========================================================================
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: #666; padding: 20px;'>"
    "<b>🏭 Dashboard Analisis Mesin Diesel — Campuran Waste Plastic Oil (WPO)</b><br>"
    "Dataset kompilasi dari berbagai paper | Dibuat dengan Streamlit + Python"
    "</div>",
    unsafe_allow_html=True,
)

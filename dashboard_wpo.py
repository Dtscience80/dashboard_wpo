"""
Dashboard Analisis Mesin Diesel — Campuran Waste/Plastic Pyrolysis Oil (PPO)
Dataset: 10_Juli_Dataset_Performance_dan_emisi_WPO_diesel_engine_gabung_paper.xlsx
          (Sheet: "Table 1")
EDA Komprehensif — Performa, Emisi, Properti Bahan Bakar & Spesifikasi Mesin
"""

import warnings
warnings.filterwarnings("ignore")

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from scipy.stats import gaussian_kde

# =========================================================================
# KONFIGURASI HALAMAN — WAJIB PERTAMA
# =========================================================================
st.set_page_config(
    page_title="Dashboard PPO/WPO — Analisis Mesin Diesel",
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
RAW_TO_SNAKE = {
    "PPO (%)": "ppo_pct",
    "Load (%)": "load_pct",
    "Speed (rpm)": "speed_rpm",
    "Power (kW)": "power_kw",
    "Torsi (N.m)": "torque_nm",
    "BTE": "bte_pct",
    "BSFC (kg/kW-hr)": "bsfc_kg_kwh",
    "NOx (ppm)": "nox_ppm",
    "HC": "hc_ppm",
    "CO": "co_ppm",
    "CO2": "co2_ppm",
    "Engine model": "engine_model",
    "Cooling system": "cooling_system",
    "Bore x stroke": "bore_stroke",
    "Displacement volume (cc)": "displacement_cc",
    "Rated power (kW)": "rated_power_kw",
    "Maximum torque (N.m)": "max_torque_nm",
    "Compression rasio": "compression_ratio",
    "Komponen \nBahan Bakar": "fuel_composition",
    "Higher heating value (MJ/kg)": "hhv_mj_kg",
    "Kinematics viscosity (mm2/s)": "viscosity_mm2s",
    "Density (g/cm3)": "density_g_cm3",
    "Flash point (°C)": "flash_point_c",
    "Cetane Index": "cetane_index",
    "Source paper": "source_paper",
}

NUM_COLS = [
    "ppo_pct", "load_pct", "speed_rpm", "power_kw", "torque_nm",
    "bte_pct", "bsfc_kg_kwh", "nox_ppm", "hc_ppm", "co_ppm", "co2_ppm",
]
FUEL_PROP_COLS = ["hhv_mj_kg", "viscosity_mm2s", "density_g_cm3", "flash_point_c", "cetane_index"]
ENGINE_SPEC_COLS = ["displacement_cc", "rated_power_kw", "max_torque_nm"]
CAT_COLS = ["engine_model", "cooling_system", "compression_ratio", "fuel_composition", "source_paper"]

NUM_LABELS = {
    "ppo_pct": "Campuran PPO (%)",
    "load_pct": "Beban Mesin (%)",
    "speed_rpm": "Kecepatan Mesin (RPM)",
    "power_kw": "Daya (kW)",
    "torque_nm": "Torsi (Nm)",
    "bte_pct": "BTE — Brake Thermal Efficiency (%)",
    "bsfc_kg_kwh": "BSFC (kg/kWh)",
    "nox_ppm": "NOx (ppm)",
    "hc_ppm": "HC (ppm)",
    "co_ppm": "CO (ppm)",
    "co2_ppm": "CO₂ (ppm)",
    "hhv_mj_kg": "Nilai Kalor / HHV (MJ/kg)",
    "viscosity_mm2s": "Viskositas Kinematik (mm²/s)",
    "density_g_cm3": "Densitas (g/cm³)",
    "flash_point_c": "Flash Point (°C)",
    "cetane_index": "Cetane Index",
    "displacement_cc": "Volume Silinder (cc)",
    "rated_power_kw": "Daya Terpasang (kW)",
    "max_torque_nm": "Torsi Maksimum (Nm)",
    "engine_model": "Model Mesin",
    "cooling_system": "Sistem Pendingin",
    "compression_ratio": "Rasio Kompresi",
    "fuel_composition": "Komposisi Bahan Bakar",
    "source_paper": "Sumber Paper",
}


@st.cache_data
def load_data():
    xls_path = "10_Juli_Dataset_Performance_dan_emisi_WPO_diesel_engine_gabung_paper.xlsx"
    df = pd.read_excel(xls_path, sheet_name="Table 1", engine="openpyxl")
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]

    # Cocokkan nama kolom mentah -> snake_case (toleran terhadap variasi spasi/newline)
    rename_map = {}
    for raw_col in df.columns:
        key = raw_col.replace("\n", " ").strip() if isinstance(raw_col, str) else raw_col
        matched = None
        for raw_ref, snake in RAW_TO_SNAKE.items():
            ref_key = raw_ref.replace("\n", " ").strip()
            if key == ref_key:
                matched = snake
                break
        rename_map[raw_col] = matched if matched else raw_col
    df = df.rename(columns=rename_map)

    # Pastikan kolom numerik benar-benar numerik
    for col in NUM_COLS + FUEL_PROP_COLS + ENGINE_SPEC_COLS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Kolom kategorikal — bersihkan spasi
    for col in CAT_COLS + ["bore_stroke"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace({"nan": np.nan, "None": np.nan})

    return df


df = load_data()
paper_col = "source_paper"

# =========================================================================
# SIDEBAR — FILTER
# =========================================================================
st.sidebar.title("🏭 Dashboard PPO/WPO")
st.sidebar.markdown("**Analisis Mesin Diesel — Campuran Plastic Pyrolysis Oil (PPO/WPO)**")
st.sidebar.markdown("---")

ppo_options = sorted(df["ppo_pct"].dropna().unique())
selected_ppo = st.sidebar.multiselect("Campuran PPO (%)", ppo_options, default=ppo_options)

rpm_options = sorted(df["speed_rpm"].dropna().unique())
selected_rpm = st.sidebar.multiselect("Kecepatan Mesin (RPM)", rpm_options, default=rpm_options)

engine_options = df["engine_model"].dropna().unique().tolist()
selected_engines = st.sidebar.multiselect("Model Mesin", engine_options, default=engine_options)

paper_options = df[paper_col].dropna().unique().tolist()
selected_papers = st.sidebar.multiselect(
    "Sumber Paper", paper_options, default=paper_options,
    format_func=lambda x: (x[:45] + "...") if len(x) > 45 else x,
)

st.sidebar.markdown("---")
st.sidebar.markdown("**Tentang Dataset**")
st.sidebar.info(
    f"Dataset kompilasi performa, emisi, properti bahan bakar,\n"
    f"dan spesifikasi mesin diesel dengan campuran Plastic\n"
    f"Pyrolysis Oil / Waste Plastic Oil (PPO/WPO).\n\n"
    f"- **Total sampel**: {df.shape[0]}\n"
    f"- **Variasi PPO**: {df['ppo_pct'].nunique()} ({int(df['ppo_pct'].min())}–{int(df['ppo_pct'].max())}%)\n"
    f"- **Model mesin**: {df['engine_model'].nunique()}\n"
    f"- **Sumber paper**: {df[paper_col].nunique()}"
)

mask = (
    df["ppo_pct"].isin(selected_ppo)
    & df["speed_rpm"].isin(selected_rpm)
    & df["engine_model"].isin(selected_engines)
    & df[paper_col].isin(selected_papers)
)
df_f = df[mask].copy()

# =========================================================================
# MAIN TITLE
# =========================================================================
st.title("🏭 Dashboard Analisis Mesin Diesel — Campuran PPO/WPO")
st.markdown(
    "**Plastic Pyrolysis Oil (PPO)**, juga dikenal sebagai *Waste Plastic Oil (WPO)*, "
    "adalah bahan bakar alternatif hasil pirolisis sampah plastik untuk mesin diesel. "
    "Dashboard ini menyajikan eksplorasi data komprehensif — mulai dari performa dan emisi mesin, "
    "hingga properti fisiko-kimia bahan bakar dan spesifikasi mesin uji — dari kompilasi 3 paper penelitian."
)

if df_f.empty:
    st.warning("⚠️ Tidak ada data pada kombinasi filter ini. Silakan ubah filter di sidebar.")
    st.stop()

tot = len(df_f)
st.subheader("📊 Ringkasan Data")
m1, m2, m3, m4, m5 = st.columns(5)
with m1:
    st.metric("Total Sampel", tot)
with m2:
    st.metric("Variasi PPO", df_f["ppo_pct"].nunique())
with m3:
    rpm_min, rpm_max = int(df_f["speed_rpm"].min()), int(df_f["speed_rpm"].max())
    st.metric("Rentang RPM", f"{rpm_min}–{rpm_max}" if rpm_min != rpm_max else f"{rpm_min}")
with m4:
    st.metric("Model Mesin", df_f["engine_model"].nunique())
with m5:
    bte_mean = df_f["bte_pct"].mean()
    st.metric("Rerata BTE", f"{bte_mean:.1f}%" if pd.notna(bte_mean) else "—")

# =========================================================================
# TABS
# =========================================================================
tabs = st.tabs([
    "📋 Ringkasan Dataset",
    "📊 Distribusi & Univariat",
    "🔥 Korelasi",
    "🛢️ Pengaruh Campuran PPO",
    "🏭 Emisi vs Performa",
    "🧪 Properti Bahan Bakar & Mesin",
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
        info_data = {
            "Jumlah sampel": [df_f.shape[0]],
            "Jumlah variabel": [df_f.shape[1]],
            "Variabel numerik": [len(df_f.select_dtypes(include=[np.number]).columns)],
            "Variabel kategorikal": [len(df_f.select_dtypes(include=["object"]).columns)],
            "Model mesin": [df_f["engine_model"].nunique()],
            "Sumber paper": [df_f[paper_col].nunique()],
        }
        st.table(pd.DataFrame(info_data, index=["Nilai"]))

        st.subheader("📄 5 Sampel Data")
        display_cols = [c for c in NUM_COLS if c in df_f.columns] + ["engine_model", paper_col]
        st.dataframe(
            df_f[display_cols].head(5).style.format({c: "{:.2f}" for c in NUM_COLS if c in df_f.columns}),
            use_container_width=True, height=200,
        )

    with c2:
        st.subheader("❌ Missing Values")
        missing = df_f.isna().sum()
        missing = missing[missing > 0].sort_values(ascending=True)
        if len(missing) > 0:
            fig, ax = plt.subplots(figsize=(6, 4))
            bars = ax.barh(
                [NUM_LABELS.get(c, c) for c in missing.index], missing.values,
                color=[CARTOON_COLORS[i % len(CARTOON_COLORS)] for i in range(len(missing))],
                edgecolor="black", linewidth=2,
            )
            for bar in bars:
                w = bar.get_width()
                ax.text(w + 0.2, bar.get_y() + bar.get_height() / 2,
                        f"{int(w)}", va="center", fontweight="bold")
            ax.set_xlabel("Jumlah Missing")
            ax.set_title("Missing Values per Kolom")
            ax.margins(x=0.25)
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            st.pyplot(fig)
            plt.close()
            st.caption(
                "Missing values terjadi karena tidak semua paper melaporkan variabel yang sama "
                "(mis. CO₂, torsi maksimum, atau sistem pendingin tidak dilaporkan di sebagian paper)."
            )
        else:
            st.success("✅ Tidak ada missing values dalam data terfilter!")

    st.subheader("📊 Statistik Deskriptif — Performa & Emisi")
    desc = df_f[NUM_COLS].describe().T
    desc.columns = ["Count", "Mean", "Std", "Min", "25%", "50%", "75%", "Max"]
    desc.index = [NUM_LABELS.get(c, c) for c in desc.index]
    st.dataframe(desc.style.format("{:.2f}"), use_container_width=True)

    st.subheader("🧪 Statistik Deskriptif — Properti Bahan Bakar")
    desc_fuel = df_f[FUEL_PROP_COLS].describe().T
    desc_fuel.columns = ["Count", "Mean", "Std", "Min", "25%", "50%", "75%", "Max"]
    desc_fuel.index = [NUM_LABELS.get(c, c) for c in desc_fuel.index]
    st.dataframe(desc_fuel.style.format("{:.2f}"), use_container_width=True)

# =====================================================================
# TAB 2: DISTRIBUSI & UNIVARIAT
# =====================================================================
with tabs[1]:
    st.subheader("📊 Distribusi Setiap Variabel Numerik")

    def plot_hist_row(cols_list, color_offset):
        act_cols = [c for c in cols_list if c in df_f.columns]
        st_cols = st.columns(len(act_cols))
        for i, col_name in enumerate(act_cols):
            with st_cols[i]:
                fig, ax = plt.subplots(figsize=(4, 3))
                data = df_f[col_name].dropna()
                if len(data) == 0:
                    ax.text(0.5, 0.5, "Tidak ada data", ha="center", va="center")
                    ax.axis("off")
                else:
                    ax.hist(data, bins=min(12, max(3, data.nunique())),
                            color=CARTOON_COLORS[(i + color_offset) % len(CARTOON_COLORS)],
                            edgecolor="black", linewidth=2, density=True, alpha=0.85)
                    if data.nunique() > 1:
                        try:
                            kde = gaussian_kde(data)
                            x_kde = np.linspace(data.min(), data.max(), 200)
                            ax.plot(x_kde, kde(x_kde), color="black", linewidth=2.5, linestyle="--")
                        except Exception:
                            pass
                    ax.set_xlabel(NUM_LABELS.get(col_name, col_name), fontsize=9)
                    ax.set_ylabel("Density", fontsize=9)
                    ax.set_title(NUM_LABELS.get(col_name, col_name).split("(")[0].strip(), fontsize=10, fontweight="bold")
                    ax.spines["top"].set_visible(False)
                    ax.spines["right"].set_visible(False)
                st.pyplot(fig)
                plt.close()

    st.markdown("**Parameter Uji Mesin**")
    plot_hist_row(["ppo_pct", "load_pct", "speed_rpm"], 0)

    st.markdown("**Performa Mesin**")
    plot_hist_row(["power_kw", "torque_nm", "bte_pct", "bsfc_kg_kwh"], 4)

    st.markdown("**Emisi**")
    plot_hist_row(["nox_ppm", "hc_ppm", "co_ppm", "co2_ppm"], 8)

    st.subheader("📦 Boxplot per Variabel — Dikelompokkan berdasarkan Campuran PPO")
    box_var = st.selectbox(
        "Pilih variabel untuk boxplot:",
        [c for c in NUM_COLS if c in df_f.columns and c != "ppo_pct"],
        format_func=lambda x: NUM_LABELS.get(x, x),
    )
    if box_var:
        ppo_groups = sorted(df_f["ppo_pct"].dropna().unique())
        data_groups = [df_f[df_f["ppo_pct"] == w][box_var].dropna().values for w in ppo_groups]
        # Buang grup kosong agar boxplot tidak error
        valid = [(w, d) for w, d in zip(ppo_groups, data_groups) if len(d) > 0]
        if valid:
            ppo_groups, data_groups = zip(*valid)
            fig, ax = plt.subplots(figsize=(10, 5))
            bp = ax.boxplot(
                data_groups, patch_artist=True, widths=0.5,
                boxprops=dict(edgecolor="black", linewidth=2),
                whiskerprops=dict(color="black", linewidth=2),
                capprops=dict(color="black", linewidth=2),
                medianprops=dict(color="black", linewidth=2, linestyle="-"),
                flierprops=dict(marker="o", markerfacecolor="#FF6B6B", markersize=6, markeredgecolor="black"),
            )
            for patch, color in zip(bp["boxes"], CARTOON_COLORS):
                patch.set_facecolor(color)
            ax.set_xticklabels([f"{int(w)}%" for w in ppo_groups])
            ax.set_xlabel("Campuran PPO (%)", fontweight="bold")
            ax.set_ylabel(NUM_LABELS.get(box_var, box_var), fontweight="bold")
            ax.set_title(f"Distribusi {NUM_LABELS.get(box_var, box_var)} per Campuran PPO", fontsize=14, fontweight="bold")
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            st.pyplot(fig)
            plt.close()

# =====================================================================
# TAB 3: KORELASI
# =====================================================================
with tabs[2]:
    st.subheader("🔥 Matriks Korelasi — Semua Variabel Numerik")
    corr_cols = [c for c in NUM_COLS + FUEL_PROP_COLS if c in df_f.columns]
    corr_data = df_f[corr_cols].dropna(how="all")
    corr = corr_data.corr()
    fig, ax = plt.subplots(figsize=(11, 9))
    labels = [NUM_LABELS.get(c, c).split("(")[0].strip() for c in corr.columns]
    sns.heatmap(
        corr, annot=True, fmt=".2f", cmap="RdBu_r", center=0, vmin=-1, vmax=1,
        linewidths=2, linecolor="black", cbar_kws={"shrink": 0.8, "label": "Korelasi"},
        xticklabels=labels, yticklabels=labels, ax=ax,
        annot_kws={"fontsize": 8, "fontweight": "bold"},
    )
    ax.set_title("Korelasi Antar Variabel Performa, Emisi & Properti Bahan Bakar", fontsize=14, fontweight="bold")
    plt.xticks(rotation=45, ha="right", fontsize=9)
    plt.yticks(rotation=0, fontsize=9)
    st.pyplot(fig)
    plt.close()

    st.markdown(
        "💡 **Insight**: perhatikan korelasi `Campuran PPO (%)` dengan `BSFC` dan `NOx` — "
        "umumnya kenaikan campuran PPO berasosiasi dengan perubahan pada efisiensi bahan bakar "
        "dan emisi karena perbedaan nilai kalor dan viskositas bahan bakar plastik dibanding solar murni."
    )

# =====================================================================
# TAB 4: PENGARUH CAMPURAN PPO
# =====================================================================
with tabs[3]:
    st.subheader("🛢️ Pengaruh Campuran PPO terhadap Performa & Emisi")

    perf_cols = [c for c in ["power_kw", "torque_nm", "bsfc_kg_kwh", "bte_pct"] if c in df_f.columns]
    emit_cols = [c for c in ["nox_ppm", "hc_ppm", "co_ppm", "co2_ppm"] if c in df_f.columns]

    agg_perf = df_f.groupby("ppo_pct")[perf_cols].mean().reset_index()
    fig2 = make_subplots(
        rows=2, cols=2, subplot_titles=[NUM_LABELS.get(c, c) for c in perf_cols],
        vertical_spacing=0.15, horizontal_spacing=0.08,
    )
    colors_line2 = ["#DDA0DD", "#98D8C8", "#F7DC6F", "#BB8FCE"]
    for i, col in enumerate(perf_cols):
        r, c = i // 2 + 1, i % 2 + 1
        fig2.add_trace(
            go.Scatter(
                x=agg_perf["ppo_pct"], y=agg_perf[col], mode="lines+markers",
                name=NUM_LABELS.get(col, col), line=dict(color=colors_line2[i], width=3),
                marker=dict(size=10, color=colors_line2[i], line=dict(color="black", width=2)),
            ), row=r, col=c,
        )
        fig2.update_xaxes(title_text="PPO (%)", row=r, col=c)
        fig2.update_yaxes(title_text=NUM_LABELS.get(col, col).split("(")[0].strip(), row=r, col=c)
    fig2.update_layout(height=550, title_text="Rata-rata Performa vs Campuran PPO", font=dict(size=12), showlegend=False)
    st.plotly_chart(fig2, use_container_width=True)

    st.markdown("### Pengaruh PPO — Dibedakan berdasarkan Kecepatan Mesin")
    rpm_choices = sorted(df_f["speed_rpm"].dropna().unique())
    split_var = st.selectbox(
        "Pilih variabel dependen:", perf_cols + emit_cols,
        format_func=lambda x: NUM_LABELS.get(x, x), key="split_var",
    )
    if split_var and len(rpm_choices) > 0:
        fig3, ax3 = plt.subplots(figsize=(10, 5))
        has_data = False
        for i, rpm in enumerate(rpm_choices):
            sub = df_f[df_f["speed_rpm"] == rpm].groupby("ppo_pct")[split_var].mean().reset_index().dropna()
            if len(sub) == 0:
                continue
            has_data = True
            ax3.plot(
                sub["ppo_pct"], sub[split_var], marker="o", markersize=8, linewidth=2.5,
                color=CARTOON_COLORS[i % len(CARTOON_COLORS)], label=f"{int(rpm)} RPM",
                markerfacecolor=CARTOON_COLORS[i % len(CARTOON_COLORS)],
                markeredgecolor="black", markeredgewidth=2,
            )
        if has_data:
            ax3.set_xlabel("Campuran PPO (%)", fontweight="bold")
            ax3.set_ylabel(NUM_LABELS.get(split_var, split_var), fontweight="bold")
            ax3.set_title(f"{NUM_LABELS.get(split_var, split_var)} vs PPO — per Kecepatan Mesin", fontsize=14, fontweight="bold")
            ax3.legend(frameon=True, edgecolor="black", facecolor="white", fontsize=11)
            ax3.spines["top"].set_visible(False)
            ax3.spines["right"].set_visible(False)
            st.pyplot(fig3)
        plt.close()

    st.markdown("### Perbandingan Rerata — BTE dan BSFC per Campuran PPO")
    bar_cols = [c for c in ["bte_pct", "bsfc_kg_kwh"] if c in df_f.columns]
    if len(bar_cols) == 2:
        agg_bar = df_f.groupby("ppo_pct")[bar_cols].mean().reset_index().dropna()
        if len(agg_bar) > 0:
            fig4, axes = plt.subplots(1, 2, figsize=(12, 5))
            for idx, col in enumerate(bar_cols):
                ax = axes[idx]
                x_pos = range(len(agg_bar))
                ax.bar(x_pos, agg_bar[col],
                       color=[CARTOON_COLORS[i % len(CARTOON_COLORS)] for i in range(len(agg_bar))],
                       edgecolor="black", linewidth=2, width=0.6)
                ax.set_xticks(x_pos)
                ax.set_xticklabels([f"{int(w)}%" for w in agg_bar["ppo_pct"]], fontweight="bold")
                ax.set_xlabel("Campuran PPO (%)", fontweight="bold")
                ax.set_ylabel(NUM_LABELS.get(col, col).split("(")[0].strip(), fontweight="bold")
                ax.set_title(NUM_LABELS.get(col, col), fontsize=13, fontweight="bold")
                ax.spines["top"].set_visible(False)
                ax.spines["right"].set_visible(False)
                for j, v in enumerate(agg_bar[col]):
                    ax.text(j, v + 0.01 * max(agg_bar[col]), f"{v:.2f}", ha="center", va="bottom", fontweight="bold", fontsize=10)
            fig4.suptitle("Perbandingan Rerata per Campuran PPO", fontsize=15, fontweight="bold")
            fig4.tight_layout()
            st.pyplot(fig4)
            plt.close()

# =====================================================================
# TAB 5: EMISI VS PERFORMA
# =====================================================================
with tabs[4]:
    st.subheader("🏭 Hubungan antara Emisi dan Performa Mesin")

    col_emit = [c for c in ["nox_ppm", "hc_ppm", "co_ppm", "co2_ppm"] if c in df_f.columns]
    col_perf = [c for c in ["power_kw", "torque_nm", "bsfc_kg_kwh", "bte_pct"] if c in df_f.columns]

    sc_x = st.selectbox("Sumbu X:", col_emit + col_perf, index=0, format_func=lambda x: NUM_LABELS.get(x, x), key="sc_x")
    sc_y = st.selectbox("Sumbu Y:", col_emit + col_perf, index=1, format_func=lambda x: NUM_LABELS.get(x, x), key="sc_y")
    sc_color = st.selectbox("Warna berdasarkan:", ["ppo_pct"] + CAT_COLS, format_func=lambda x: NUM_LABELS.get(x, x), key="sc_color")

    sc_data = df_f[[sc_x, sc_y, sc_color]].dropna().copy()
    if len(sc_data) > 0:
        if sc_color == "ppo_pct":
            fig5 = px.scatter(sc_data, x=sc_x, y=sc_y, color=sc_color, color_continuous_scale="viridis",
                               labels={sc_x: NUM_LABELS.get(sc_x, sc_x), sc_y: NUM_LABELS.get(sc_y, sc_y)},
                               title=f"{NUM_LABELS.get(sc_y, sc_y)} vs {NUM_LABELS.get(sc_x, sc_x)}")
        else:
            fig5 = px.scatter(sc_data, x=sc_x, y=sc_y, color=sc_color, color_discrete_sequence=CARTOON_COLORS,
                               labels={sc_x: NUM_LABELS.get(sc_x, sc_x), sc_y: NUM_LABELS.get(sc_y, sc_y)},
                               title=f"{NUM_LABELS.get(sc_y, sc_y)} vs {NUM_LABELS.get(sc_x, sc_x)}")
        fig5.update_traces(marker=dict(size=11, line=dict(width=2, color="black")))
        fig5.update_layout(font=dict(size=13))
        st.plotly_chart(fig5, use_container_width=True)
    else:
        st.info("Tidak ada data lengkap untuk kombinasi variabel ini.")

    st.subheader("🔷 Scatter Matrix — Emisi vs Performa")
    sm_cols = st.multiselect(
        "Pilih variabel untuk scatter matrix:", col_emit + col_perf + ["load_pct"],
        default=(col_emit + col_perf[:2])[:5], format_func=lambda x: NUM_LABELS.get(x, x),
    )
    if len(sm_cols) >= 2:
        sm_data = df_f[sm_cols].dropna()
        if len(sm_data) > 1:
            fig6 = px.scatter_matrix(
                sm_data, labels={c: NUM_LABELS.get(c, c).split("(")[0].strip() for c in sm_cols},
                dimensions=sm_cols, opacity=0.6,
            )
            fig6.update_traces(marker=dict(size=6, color="#4ECDC4", line=dict(width=1, color="black")), diagonal_visible=False)
            fig6.update_layout(height=700, title="Scatter Matrix — Hubungan Antar Variabel", font=dict(size=11))
            st.plotly_chart(fig6, use_container_width=True)

    st.subheader("📊 Korelasi Emisi vs Performa")
    if len(col_emit) > 0 and len(col_perf) > 0:
        corr_ep = df_f[col_emit + col_perf].dropna().corr().loc[col_emit, col_perf]
        fig7, ax7 = plt.subplots(figsize=(9, 5))
        sns.heatmap(corr_ep, annot=True, fmt=".2f", cmap="RdBu_r", center=0, vmin=-1, vmax=1,
                    linewidths=2, linecolor="black", cbar_kws={"shrink": 0.8, "label": "Korelasi"}, ax=ax7,
                    annot_kws={"fontsize": 10, "fontweight": "bold"})
        emit_labels = [NUM_LABELS.get(c, c).split("(")[0].strip() for c in col_emit]
        perf_labels = [NUM_LABELS.get(c, c).split("(")[0].strip() for c in col_perf]
        ax7.set_xticklabels(perf_labels, rotation=30, ha="right", fontsize=10)
        ax7.set_yticklabels(emit_labels, rotation=0, fontsize=10)
        ax7.set_title("Korelasi: Emisi vs Parameter Performa", fontsize=14, fontweight="bold")
        st.pyplot(fig7)
        plt.close()

# =====================================================================
# TAB 6: PROPERTI BAHAN BAKAR & SPESIFIKASI MESIN (INSIDE VIEW BARU)
# =====================================================================
with tabs[5]:
    st.subheader("🧪 Properti Fisiko-Kimia Bahan Bakar")
    st.markdown(
        "Dataset ini mencatat properti bahan bakar hasil pengujian (nilai kalor, viskositas, "
        "densitas, flash point, cetane index) untuk tiap komposisi campuran — sesuatu yang jarang "
        "ditampilkan di dashboard EDA biasa, namun penting untuk memahami *mengapa* performa berubah."
    )

    fuel_tbl = df_f.groupby("fuel_composition")[FUEL_PROP_COLS].mean().round(2)
    fuel_tbl = fuel_tbl.dropna(how="all")
    st.markdown("### 📋 Rata-rata Properti per Komposisi Bahan Bakar")
    st.dataframe(fuel_tbl.style.background_gradient(cmap="YlGnBu"), use_container_width=True)

    prop_choice = st.selectbox(
        "Pilih properti bahan bakar untuk divisualisasikan:", FUEL_PROP_COLS,
        format_func=lambda x: NUM_LABELS.get(x, x), key="prop_choice",
    )
    prop_data = df_f[["fuel_composition", prop_choice]].dropna().drop_duplicates()
    if len(prop_data) > 0:
        prop_data = prop_data.sort_values(prop_choice)
        fig11, ax11 = plt.subplots(figsize=(11, 5))
        short_labels = [s[:35] + "..." if len(s) > 35 else s for s in prop_data["fuel_composition"]]
        bars = ax11.barh(
            range(len(prop_data)), prop_data[prop_choice],
            color=[CARTOON_COLORS[i % len(CARTOON_COLORS)] for i in range(len(prop_data))],
            edgecolor="black", linewidth=2,
        )
        ax11.set_yticks(range(len(prop_data)))
        ax11.set_yticklabels(short_labels, fontsize=9)
        ax11.set_xlabel(NUM_LABELS.get(prop_choice, prop_choice), fontweight="bold")
        ax11.set_title(f"{NUM_LABELS.get(prop_choice, prop_choice)} per Komposisi Bahan Bakar", fontsize=13, fontweight="bold")
        ax11.spines["top"].set_visible(False)
        ax11.spines["right"].set_visible(False)
        for bar, val in zip(bars, prop_data[prop_choice]):
            ax11.text(val, bar.get_y() + bar.get_height() / 2, f" {val:.2f}", va="center", fontweight="bold", fontsize=9)
        st.pyplot(fig11)
        plt.close()

    st.markdown("### 🔗 Hubungan Properti Bahan Bakar dengan Performa Mesin")
    fp_x = st.selectbox("Properti bahan bakar (X):", FUEL_PROP_COLS, format_func=lambda x: NUM_LABELS.get(x, x), key="fp_x")
    fp_y = st.selectbox(
        "Variabel performa/emisi (Y):", [c for c in NUM_COLS if c not in ("ppo_pct", "load_pct", "speed_rpm")],
        format_func=lambda x: NUM_LABELS.get(x, x), key="fp_y",
    )
    fp_data = df_f[[fp_x, fp_y, "ppo_pct", "fuel_composition"]].dropna()
    if len(fp_data) > 1:
        fig12 = px.scatter(
            fp_data, x=fp_x, y=fp_y, color="ppo_pct", color_continuous_scale="viridis",
            hover_data=["fuel_composition"],
            labels={fp_x: NUM_LABELS.get(fp_x, fp_x), fp_y: NUM_LABELS.get(fp_y, fp_y), "ppo_pct": "PPO (%)"},
            title=f"{NUM_LABELS.get(fp_y, fp_y)} vs {NUM_LABELS.get(fp_x, fp_x)}",
        )
        fig12.update_traces(marker=dict(size=11, line=dict(width=2, color="black")))
        st.plotly_chart(fig12, use_container_width=True)
    else:
        st.info("Data properti bahan bakar tidak cukup lengkap untuk kombinasi ini.")

    st.markdown("---")
    st.subheader("⚙️ Spesifikasi Mesin Uji")
    spec_cols_show = ["engine_model", "cooling_system", "bore_stroke", "compression_ratio"] + ENGINE_SPEC_COLS
    spec_tbl = df_f[[c for c in spec_cols_show if c in df_f.columns]].drop_duplicates(subset=["engine_model"])
    spec_tbl = spec_tbl.rename(columns=NUM_LABELS)
    st.dataframe(spec_tbl.reset_index(drop=True), use_container_width=True)

    n_engines = df_f["engine_model"].nunique()
    if n_engines > 1:
        st.markdown("### 📊 Perbandingan BTE Rata-rata antar Model Mesin")
        eng_bte = df_f.groupby("engine_model")["bte_pct"].mean().reset_index().dropna().sort_values("bte_pct")
        fig13, ax13 = plt.subplots(figsize=(9, 4.5))
        bars = ax13.barh(eng_bte["engine_model"], eng_bte["bte_pct"],
                          color=CARTOON_COLORS[:len(eng_bte)], edgecolor="black", linewidth=2)
        ax13.set_xlabel("BTE Rata-rata (%)", fontweight="bold")
        ax13.set_title("BTE Rata-rata per Model Mesin", fontsize=13, fontweight="bold")
        ax13.spines["top"].set_visible(False)
        ax13.spines["right"].set_visible(False)
        for bar, val in zip(bars, eng_bte["bte_pct"]):
            ax13.text(val + 0.2, bar.get_y() + bar.get_height() / 2, f"{val:.1f}%", va="center", fontweight="bold")
        st.pyplot(fig13)
        plt.close()

# =====================================================================
# TAB 7: PERBANDINGAN PAPER
# =====================================================================
with tabs[6]:
    st.subheader("📄 Perbandingan Antar Sumber Paper")

    paper_counts = df_f[paper_col].value_counts()
    st.markdown(f"**Jumlah paper unik**: {len(paper_counts)}")

    paper_agg = df_f.groupby(paper_col)[[c for c in NUM_COLS if c in df_f.columns and c != "ppo_pct"]].mean()

    if "bte_pct" in paper_agg.columns:
        st.markdown("### ⚡ Rata-rata BTE per Paper")
        bte_data = paper_agg["bte_pct"].reset_index().dropna()
        fig8, ax8 = plt.subplots(figsize=(12, 4))
        short_labels = [s[:60] + "..." if len(s) > 60 else s for s in bte_data[paper_col]]
        bars = ax8.barh(range(len(bte_data)), bte_data["bte_pct"],
                         color=[CARTOON_COLORS[i % len(CARTOON_COLORS)] for i in range(len(bte_data))],
                         edgecolor="black", linewidth=2)
        ax8.set_yticks(range(len(bte_data)))
        ax8.set_yticklabels(short_labels, fontsize=8)
        ax8.set_xlabel("BTE Rata-rata (%)", fontweight="bold")
        ax8.set_title("Perbandingan Brake Thermal Efficiency (BTE) antar Paper", fontsize=14, fontweight="bold")
        ax8.invert_yaxis()
        ax8.spines["top"].set_visible(False)
        ax8.spines["right"].set_visible(False)
        for bar, val in zip(bars, bte_data["bte_pct"]):
            ax8.text(val + 0.2, bar.get_y() + bar.get_height() / 2, f"{val:.1f}%", va="center", fontweight="bold", fontsize=9)
        st.pyplot(fig8)
        plt.close()

    st.markdown("### 📊 Perbandingan Multi-Metrik antar Paper")
    compare_metrics = st.multiselect(
        "Pilih metrik yang dibandingkan:",
        [c for c in NUM_COLS if c in df_f.columns and c != "ppo_pct"],
        default=["bte_pct", "nox_ppm", "bsfc_kg_kwh"],
        format_func=lambda x: NUM_LABELS.get(x, x), key="compare_metrics",
    )
    if compare_metrics:
        radar_data = df_f.groupby(paper_col)[compare_metrics].mean().dropna()
        if len(radar_data) > 0:
            fig9 = go.Figure()
            for i, (paper_name, row) in enumerate(radar_data.iterrows()):
                short_name = paper_name[:40] + "..." if len(paper_name) > 40 else paper_name
                fig9.add_trace(go.Scatterpolar(
                    r=row.values, theta=[NUM_LABELS.get(c, c).split("(")[0].strip() for c in compare_metrics],
                    fill="toself", name=short_name,
                    line=dict(color=CARTOON_COLORS[i % len(CARTOON_COLORS)], width=2),
                ))
            fig9.update_layout(polar=dict(radialaxis=dict(visible=True)),
                                title="Perbandingan Multi-Metrik antar Paper (Radar Chart)",
                                font=dict(size=11), height=600)
            st.plotly_chart(fig9, use_container_width=True)

    st.markdown("### 📋 Tabel Rata-rata per Paper")
    display_metrics = st.multiselect(
        "Tampilkan metrik:", [c for c in NUM_COLS if c in df_f.columns],
        default=NUM_COLS[:7], format_func=lambda x: NUM_LABELS.get(x, x), key="display_metrics",
    )
    if display_metrics:
        paper_table = df_f.groupby(paper_col)[display_metrics].mean().round(2)
        paper_table.index = [s[:60] + "..." if len(s) > 60 else s for s in paper_table.index]
        st.dataframe(paper_table.style.background_gradient(cmap="YlOrRd"), use_container_width=True)

# =====================================================================
# TAB 8: EKSPLORASI MANDIRI
# =====================================================================
with tabs[7]:
    st.subheader("🔍 Eksplorasi Mandiri — Scatter Plot Interaktif")

    all_numeric = [c for c in NUM_COLS + FUEL_PROP_COLS + ENGINE_SPEC_COLS if c in df_f.columns]
    all_cats = [c for c in CAT_COLS if c in df_f.columns]

    x_axis = st.selectbox("Sumbu X:", all_numeric, index=0, format_func=lambda x: NUM_LABELS.get(x, x), key="expl_x")
    y_axis = st.selectbox("Sumbu Y:", all_numeric, index=1, format_func=lambda x: NUM_LABELS.get(x, x), key="expl_y")
    color_by = st.selectbox("Warna:", all_cats + all_numeric,
                             format_func=lambda x: NUM_LABELS.get(x, x), key="expl_color")
    size_by = st.selectbox("Ukuran titik:", ["(sama)"] + all_numeric,
                            format_func=lambda x: NUM_LABELS.get(x, x) if x in NUM_LABELS else "(sama)", key="expl_size")

    hover_extra = [c for c in NUM_COLS[:4] if c in df_f.columns] + [paper_col]
    base_cols = [x_axis, y_axis, color_by] + ([size_by] if size_by != "(sama)" else [])
    extra_cols = [c for c in hover_extra if c in df_f.columns and c not in base_cols]
    all_plot_cols = list(dict.fromkeys(base_cols + extra_cols))
    plot_data = df_f[all_plot_cols].dropna()
    hover_data = {c: True for c in hover_extra if c in plot_data.columns}

    if len(plot_data) == 0:
        st.info("Tidak ada data lengkap untuk kombinasi variabel ini.")
    else:
        if color_by in all_numeric:
            fig10 = px.scatter(
                plot_data, x=x_axis, y=y_axis, color=color_by,
                size=size_by if size_by != "(sama)" else None, hover_data=hover_data,
                color_continuous_scale="viridis",
                labels={x_axis: NUM_LABELS.get(x_axis, x_axis), y_axis: NUM_LABELS.get(y_axis, y_axis),
                        color_by: NUM_LABELS.get(color_by, color_by)},
                title=f"{NUM_LABELS.get(y_axis, y_axis)} vs {NUM_LABELS.get(x_axis, x_axis)}",
            )
        else:
            fig10 = px.scatter(
                plot_data, x=x_axis, y=y_axis, color=color_by,
                size=size_by if size_by != "(sama)" else None, hover_data=hover_data,
                color_discrete_sequence=CARTOON_COLORS,
                labels={x_axis: NUM_LABELS.get(x_axis, x_axis), y_axis: NUM_LABELS.get(y_axis, y_axis), color_by: NUM_LABELS.get(color_by, color_by)},
                title=f"{NUM_LABELS.get(y_axis, y_axis)} vs {NUM_LABELS.get(x_axis, x_axis)}",
            )
        fig10.update_traces(marker=dict(line=dict(width=1.5, color="black")))
        fig10.update_layout(height=600, font=dict(size=13))
        st.plotly_chart(fig10, use_container_width=True)

    st.markdown("### 📋 Data Tersaring")
    show_cols = st.multiselect(
        "Pilih kolom untuk ditampilkan:", df_f.columns.tolist(),
        default=[c for c in NUM_COLS if c in df_f.columns] + ["engine_model", paper_col],
        format_func=lambda x: NUM_LABELS.get(x, x) if x in NUM_LABELS else x, key="show_cols",
    )
    if show_cols:
        num_show = [c for c in all_numeric if c in show_cols]
        st.dataframe(
            df_f[show_cols].style.format({c: "{:.2f}" for c in num_show}),
            use_container_width=True, height=400,
        )

# =========================================================================
# FOOTER
# =========================================================================
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: #666; padding: 20px;'>"
    "<b>🏭 Dashboard Analisis Mesin Diesel — Campuran Plastic Pyrolysis Oil / Waste Plastic Oil (PPO/WPO)</b><br>"
    "Dataset kompilasi dari 3 paper penelitian | Dibuat oleh xynsu dengan Streamlit + Python"
    "</div>",
    unsafe_allow_html=True,
)

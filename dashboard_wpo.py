"""
Dashboard Analisis Mesin Diesel — Campuran Waste/Plastic Pyrolysis Oil (PPO)
Dataset: 23_Juli_Dataset_Gabungan_Performance_Emisi_PPO_v01.xlsx
"""

import warnings
warnings.filterwarnings("ignore")

import os
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
import plotly.graph_objects as go
from scipy.stats import gaussian_kde

# =========================================================================
# KONFIGURASI HALAMAN
# =========================================================================
st.set_page_config(
    page_title="Corporate Dashboard PPO/WPO",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================================================================
# KONSTANTA VISUAL: Corporate Aesthetic & Professional Clean Layout
# =========================================================================
CORPORATE_COLORS = [
    "#003f5c", "#2f4b7c", "#665191", "#a05195", "#d45087",
    "#f95d6a", "#ff7c43", "#ffa600", "#1f77b4", "#aec7e8"
]

plt.rcParams.update({
    "figure.facecolor": "white",
    "axes.edgecolor": "#D3D3D3",
    "axes.linewidth": 1.2,
    "axes.grid": True,
    "grid.alpha": 0.3,
    "grid.color": "#CCCCCC",
    "font.size": 11,
    "font.family": "sans-serif",
    "axes.titlesize": 13,
    "axes.titleweight": "bold",
})
sns.set_style("whitegrid")

# =========================================================================
# LOAD DATA
# =========================================================================
# Nama file dataset baru (letakkan file ini di folder yang sama dengan script)
DATA_FILE = "23_Juli_Dataset_Gabungan_Performance_Emisi_PPO_v01.xlsx"

# Mapping dari nama kolom mentah (persis seperti di file Excel) ke nama snake_case
# yang dipakai di seluruh dashboard. Disesuaikan dengan header pada dataset baru.
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
    "Smoke (%)": "smoke_pct",
    "BMEP (kPa)": "bmep_kpa",
    "Engine model": "engine_model",
    "Cooling system": "cooling_system",
    "Bore x stroke": "bore_stroke",
    "Displacement volume (cc)": "displacement_cc",
    "Rated power (kW)": "rated_power_kw",
    "Maximum torque (N.m)": "max_torque_nm",
    "Compression ratio": "compression_ratio",
    "Komponen Bahan Bakar": "fuel_composition",
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
    "smoke_pct", "bmep_kpa",
]
FUEL_PROP_COLS = ["hhv_mj_kg", "viscosity_mm2s", "density_g_cm3", "flash_point_c", "cetane_index"]
ENGINE_SPEC_COLS = ["displacement_cc", "rated_power_kw", "max_torque_nm"]
CAT_COLS = ["engine_model", "cooling_system", "compression_ratio", "fuel_composition", "source_paper"]

NUM_LABELS = {
    "ppo_pct": "Campuran PPO (%)",
    "load_pct": "Beban Mesin / Load (%)",
    "speed_rpm": "Kecepatan Mesin (RPM)",
    "power_kw": "Daya (kW)",
    "torque_nm": "Torsi (Nm)",
    "bte_pct": "Brake Thermal Efficiency / BTE (%)",
    "bsfc_kg_kwh": "BSFC (kg/kWh)",
    "nox_ppm": "NOx (ppm)",
    "hc_ppm": "HC (ppm)",
    "co_ppm": "CO (ppm)",
    "co2_ppm": "CO₂ (ppm)",
    "smoke_pct": "Smoke / Opasitas Asap (%)",
    "bmep_kpa": "BMEP (kPa)",
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
def load_data(xls_path: str):
    if not os.path.exists(xls_path):
        st.error(
            f"❌ File dataset '{xls_path}' tidak ditemukan. "
            "Pastikan file .xlsx berada di folder yang sama dengan script ini."
        )
        st.stop()

    try:
        df = pd.read_excel(xls_path, engine="openpyxl")
    except Exception as e:
        st.error(
            "❌ Gagal membaca file Excel. Kemungkinan file tersimpan dalam format "
            "**Strict Open XML Spreadsheet**, yang tidak didukung oleh openpyxl/pandas.\n\n"
            "Cara memperbaiki: buka file di Microsoft Excel, lalu **Save As** kembali dengan "
            "format standar **'Excel Workbook (*.xlsx)'** (bukan 'Strict Open XML Spreadsheet'), "
            "lalu coba lagi.\n\n"
            f"Detail error teknis: {e}"
        )
        st.stop()

    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]

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

    for col in NUM_COLS + FUEL_PROP_COLS + ENGINE_SPEC_COLS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    for col in CAT_COLS + ["bore_stroke"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace({"nan": np.nan, "None": np.nan})

    return df


df = load_data(DATA_FILE)
paper_col = "source_paper"

# =========================================================================
# SIDEBAR — FILTER (Classic Widget Layout)
# =========================================================================
st.sidebar.title("🏢 Executive Dashboard")
st.sidebar.markdown("Pusat Kendali Analisis Mesin Diesel PPO/WPO")
st.sidebar.markdown("---")

ppo_options = sorted(df["ppo_pct"].dropna().unique())
selected_ppo = st.sidebar.multiselect("Campuran PPO (%)", ppo_options, default=ppo_options)

rpm_options = sorted(df["speed_rpm"].dropna().unique())
selected_rpm = st.sidebar.multiselect("Kecepatan Mesin (RPM)", rpm_options, default=rpm_options)

engine_options = df["engine_model"].dropna().unique().tolist()
selected_engines = st.sidebar.multiselect("Model Mesin", engine_options, default=engine_options)

paper_options = df[paper_col].dropna().unique().tolist()
selected_papers = st.sidebar.multiselect(
    "Sumber Paper (Jurnal)", paper_options, default=paper_options,
    format_func=lambda x: (x[:45] + "...") if len(x) > 45 else x,
)

st.sidebar.markdown("---")
st.sidebar.markdown("**Informasi Dataset Gabungan**")
st.sidebar.info(
    f"- **Total Data**: {df.shape[0]} observasi\n"
    f"- **Jumlah Sumber Paper**: {df[paper_col].nunique()} jurnal\n"
    f"- **Rentang Load**: {df['load_pct'].min():.1f}% – {df['load_pct'].max():.1f}%\n"
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
st.title("Executive Analysis: Plastic Pyrolysis Oil (PPO/WPO) on Diesel Engine")
st.markdown(
    "Eksplorasi data hasil kompilasi (gabungan) beberapa studi performa dan emisi mesin diesel "
    "berbahan bakar campuran Plastic/Waste Pyrolysis Oil (PPO/WPO). "
    "Platform ini menyediakan analisis termodinamika mendalam dengan antarmuka yang terstruktur."
)

if df_f.empty:
    st.warning("⚠️ Tidak ada data pada kombinasi filter ini. Silakan ubah filter di navigasi samping.")
    st.stop()

# =========================================================================
# TABS NAVIGATION
# =========================================================================
tabs = st.tabs([
    "📋 Overview Dataset",
    "📈 Kurva Termodinamika (Load)",
    "🧊 3D Surface Analysis",
    "📊 Distribusi Univariat",
    "🔥 Matriks Korelasi",
    "🛢️ Pengaruh PPO",
    "🧪 Properti Fisikokimia",
    "🔍 Eksplorasi Interaktif",
])

# =====================================================================
# TAB 1: OVERVIEW DATASET
# =====================================================================
with tabs[0]:
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Sampel Aktif", len(df_f))
    c2.metric("Titik Load Minimum", f"{df_f['load_pct'].min():.1f}%")
    c3.metric("Rerata Efisiensi (BTE)", f"{df_f['bte_pct'].mean():.2f}%")
    c4.metric("Sumber Jurnal", df_f[paper_col].nunique())

    st.markdown("---")
    st.subheader("Data Snapshot")
    display_cols = [c for c in NUM_COLS if c in df_f.columns] + ["engine_model"]
    st.dataframe(
        df_f[display_cols].head(10).style.format({c: "{:.2f}" for c in NUM_COLS if c in df_f.columns}),
        use_container_width=True
    )

    st.subheader("Statistik Deskriptif Utama")
    desc = df_f[NUM_COLS].describe().T
    desc = desc[["mean", "std", "min", "50%", "max"]]
    desc.index = [NUM_LABELS.get(c, c) for c in desc.index]
    st.dataframe(desc.style.format("{:.2f}").background_gradient(cmap='Blues'), use_container_width=True)

# =====================================================================
# TAB 2: KURVA TERMODINAMIKA (LOAD VS SEMUA PARAMETER)
# =====================================================================
with tabs[1]:
    st.subheader("📈 Analisis Termodinamika: Load (%) vs Parameter Kinerja & Emisi")
    st.markdown(
        "Halaman khusus ini memvisualisasikan bagaimana perubahan beban mesin memengaruhi "
        "parameter utama. Garis-garis dikelompokkan berdasarkan persentase campuran PPO."
    )

    # Pilih 1 jurnal dan 1 RPM agar garis tidak berantakan
    available_papers = df_f[paper_col].unique()
    selected_paper = st.selectbox("Pilih Jurnal untuk isolasi kurva fisik:", available_papers)
    available_rpms = df_f[df_f[paper_col] == selected_paper]["speed_rpm"].unique()
    selected_rpm_curve = st.selectbox("Pilih RPM:", available_rpms)

    df_curve = df_f[(df_f[paper_col] == selected_paper) & (df_f["speed_rpm"] == selected_rpm_curve)].copy()
    df_curve["PPO (%)"] = df_curve["ppo_pct"].astype(str) + "%"

    target_kurva = [
        ("power_kw", "Power (Daya Mesin)"),
        ("torque_nm", "Torsi"),
        ("bte_pct", "Brake Thermal Efficiency (BTE)"),
        ("bsfc_kg_kwh", "Brake Specific Fuel Consumption (BSFC)"),
        ("nox_ppm", "Emisi NOx"),
        ("hc_ppm", "Emisi HC"),
        ("co_ppm", "Emisi CO"),
        ("smoke_pct", "Smoke / Opasitas Asap"),
    ]

    # Render semua grafik dalam 2 kolom yang rapi
    cols = st.columns(2)
    plot_idx = 0
    for col_id, title_name in target_kurva:
        if col_id in df_curve.columns and df_curve[col_id].notna().any():
            fig = px.line(
                df_curve.sort_values(by=["load_pct"]),
                x="load_pct", y=col_id, color="PPO (%)",
                markers=True, title=f"Load (%) vs {title_name}",
                color_discrete_sequence=CORPORATE_COLORS
            )
            fig.update_layout(xaxis_title="Load (%)", yaxis_title=title_name, hovermode="x unified")
            min_load = df_curve["load_pct"].min()
            fig.add_vline(
                x=min_load, line_dash="dash", line_color="red",
                annotation_text=f"Load Minimum Tercatat ({min_load:.1f}%)"
            )
            cols[plot_idx % 2].plotly_chart(fig, use_container_width=True)
            plot_idx += 1

# =====================================================================
# TAB 3: 3D SURFACE PLOT
# =====================================================================
with tabs[2]:
    st.subheader("🧊 Pemetaan Spasial 3D (Surface Plot)")
    st.markdown(
        "Pemetaan inovatif ini memungkinkan Anda melihat elevasi performa mesin dari dua dimensi input sekaligus: "
        "**Load (%)** dan **Campuran PPO (%)**. Berguna untuk mendeteksi 'Sweet Spot' mesin."
    )

    z_options_3d = [c for c in ["power_kw", "bte_pct", "bsfc_kg_kwh", "nox_ppm", "co_ppm", "smoke_pct"] if c in df_f.columns]
    z_var = st.selectbox("Pilih Parameter Ketinggian (Sumbu Z):", z_options_3d,
                          format_func=lambda x: NUM_LABELS.get(x, x))

    if z_var in df_f.columns:
        # Group by PPO and Load to create a meshgrid
        agg_3d = df_f.groupby(["load_pct", "ppo_pct"])[z_var].mean().reset_index()

        fig_3d = go.Figure(data=[go.Mesh3d(
            x=agg_3d['load_pct'],
            y=agg_3d['ppo_pct'],
            z=agg_3d[z_var],
            opacity=0.8,
            colorscale='Blues',
            intensity=agg_3d[z_var]
        )])

        fig_3d.update_layout(
            scene=dict(
                xaxis_title='Load (%)',
                yaxis_title='Campuran PPO (%)',
                zaxis_title=NUM_LABELS.get(z_var, z_var)
            ),
            margin=dict(l=0, r=0, b=0, t=40),
            title=f"3D Surface: Interaksi Load dan PPO terhadap {NUM_LABELS.get(z_var, z_var)}"
        )
        st.plotly_chart(fig_3d, use_container_width=True)

# =====================================================================
# TAB 4: DISTRIBUSI UNIVARIAT
# =====================================================================
with tabs[3]:
    st.subheader("📊 Analisis Distribusi Data")
    dist_var = st.selectbox("Pilih variabel untuk dianalisis:", NUM_COLS, format_func=lambda x: NUM_LABELS.get(x, x))

    if dist_var in df_f.columns:
        col1, col2 = st.columns(2)
        with col1:
            fig_hist = px.histogram(
                df_f, x=dist_var, marginal="box", color_discrete_sequence=[CORPORATE_COLORS[0]],
                title=f"Histogram & Boxplot: {NUM_LABELS.get(dist_var, dist_var)}"
            )
            st.plotly_chart(fig_hist, use_container_width=True)

        with col2:
            fig_box = px.box(
                df_f, x="ppo_pct", y=dist_var, color_discrete_sequence=[CORPORATE_COLORS[1]],
                title=f"Distribusi {NUM_LABELS.get(dist_var, dist_var)} per Campuran PPO"
            )
            st.plotly_chart(fig_box, use_container_width=True)

# =====================================================================
# TAB 5: MATRIKS KORELASI
# =====================================================================
with tabs[4]:
    st.subheader("🔥 Peta Panas Korelasi Pearson")

    candidate_cols = [c for c in NUM_COLS + FUEL_PROP_COLS if c in df_f.columns]

    # Hanya ikutkan kolom yang punya cukup data valid (>=2 baris) dan tidak konstan
    # pada subset data yang sedang difilter — kolom yang seluruhnya NaN/konstan
    # akan membuat korelasinya NaN dan terlihat "kosong" di heatmap.
    valid_cols = []
    for c in candidate_cols:
        series = df_f[c].dropna()
        if len(series) >= 2 and series.std(skipna=True) > 0:
            valid_cols.append(c)

    dropped_cols = [c for c in candidate_cols if c not in valid_cols]

    if len(valid_cols) < 2:
        st.warning(
            "⚠️ Data pada filter saat ini tidak cukup (terlalu sedikit baris, atau nilainya "
            "konstan) untuk menghitung matriks korelasi. Silakan longgarkan filter di sidebar."
        )
    else:
        corr_data = df_f[valid_cols].corr(min_periods=2)
        labels = [NUM_LABELS.get(c, c).split("(")[0].strip() for c in corr_data.columns]

        fig_corr = px.imshow(
            corr_data.values,
            x=labels,
            y=labels,
            color_continuous_scale="RdBu",
            zmin=-1, zmax=1,
            text_auto=".2f",
            aspect="auto",
        )
        fig_corr.update_layout(
            title="Korelasi Antar Parameter Numerik",
            height=700,
            coloraxis_colorbar=dict(title="r"),
        )
        fig_corr.update_traces(
            hovertemplate="%{x} vs %{y}<br>r = %{z:.2f}<extra></extra>",
            textfont_size=10,
        )
        st.plotly_chart(fig_corr, use_container_width=True)

        if dropped_cols:
            dropped_labels = ", ".join(NUM_LABELS.get(c, c) for c in dropped_cols)
            st.caption(
                f"ℹ️ Kolom berikut tidak diikutkan karena datanya kosong/konstan pada filter "
                f"saat ini: {dropped_labels}."
            )


# =====================================================================
# TAB 6: PENGARUH PPO
# =====================================================================
with tabs[5]:
    st.subheader("🛢️ Analisis Komposisi Bahan Bakar")

    bar_cols = [c for c in ["bte_pct", "bsfc_kg_kwh"] if c in df_f.columns]
    if len(bar_cols) == 2:
        agg_bar = df_f.groupby("ppo_pct")[bar_cols].mean().reset_index().dropna()

        c_bar1, c_bar2 = st.columns(2)
        with c_bar1:
            fig_bte = px.bar(agg_bar, x='ppo_pct', y='bte_pct', title='Rerata BTE per PPO (%)',
                              color_discrete_sequence=[CORPORATE_COLORS[2]])
            st.plotly_chart(fig_bte, use_container_width=True)
        with c_bar2:
            fig_bsfc = px.bar(agg_bar, x='ppo_pct', y='bsfc_kg_kwh', title='Rerata BSFC per PPO (%)',
                               color_discrete_sequence=[CORPORATE_COLORS[3]])
            st.plotly_chart(fig_bsfc, use_container_width=True)

# =====================================================================
# TAB 7: PROPERTI FISIKOKIMIA & MESIN
# =====================================================================
with tabs[6]:
    st.subheader("🧪 Karakteristik Dasar Bahan Bakar dan Spesifikasi")

    c_prop1, c_prop2 = st.columns([1.5, 1])
    with c_prop1:
        st.markdown("**Rata-rata Properti per Komposisi Bahan Bakar**")
        fuel_tbl = df_f.groupby("fuel_composition")[FUEL_PROP_COLS].mean().round(2).dropna(how="all")
        st.dataframe(fuel_tbl.style.background_gradient(cmap="Blues"), use_container_width=True)

    with c_prop2:
        st.markdown("**Spesifikasi Mesin Uji**")
        spec_cols_show = ["engine_model", "cooling_system", "compression_ratio"]
        spec_tbl = df_f[[c for c in spec_cols_show if c in df_f.columns]].drop_duplicates()
        st.dataframe(spec_tbl, use_container_width=True)

# =====================================================================
# TAB 8: EKSPLORASI INTERAKTIF
# =====================================================================
with tabs[7]:
    st.subheader("🔍 Pembangun Plot Kustom")

    all_numeric = [c for c in NUM_COLS + FUEL_PROP_COLS + ENGINE_SPEC_COLS if c in df_f.columns]

    e1, e2, e3 = st.columns(3)
    with e1:
        x_axis = st.selectbox("Sumbu X:", all_numeric, index=all_numeric.index("power_kw") if "power_kw" in all_numeric else 0, format_func=lambda x: NUM_LABELS.get(x, x))
    with e2:
        y_axis = st.selectbox("Sumbu Y:", all_numeric, index=all_numeric.index("nox_ppm") if "nox_ppm" in all_numeric else 1, format_func=lambda x: NUM_LABELS.get(x, x))
    with e3:
        color_by = st.selectbox("Warna (Grup):", ["ppo_pct", "engine_model", "source_paper"], format_func=lambda x: NUM_LABELS.get(x, x))

    fig_explore = px.scatter(
        df_f.dropna(subset=[x_axis, y_axis, color_by]),
        x=x_axis, y=y_axis, color=color_by,
        color_discrete_sequence=CORPORATE_COLORS,
        title=f"Analisis Khusus: {NUM_LABELS.get(y_axis, y_axis)} berdasarkan {NUM_LABELS.get(x_axis, x_axis)}"
    )
    fig_explore.update_traces(marker=dict(size=10, line=dict(width=1, color="DarkSlateGrey")))
    fig_explore.update_layout(height=500)
    st.plotly_chart(fig_explore, use_container_width=True)

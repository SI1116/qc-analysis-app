# --- ライブラリ読み込み ---
import streamlit as st
import pandas as pd
import numpy as np
import re
import matplotlib.pyplot as plt
import matplotlib
import matplotlib.font_manager as fm
import plotly.express as px
import os
import platform

# --- 日本語フォント設定（Windows用） ---
def configure_japanese_font():
    global font_prop
    if platform.system() == "Windows":
        font_path = "C:/Windows/Fonts/msgothic.ttc"
        if os.path.exists(font_path):
            font_prop = fm.FontProperties(fname=font_path)
            font_name = font_prop.get_name()
            fm.fontManager.addfont(font_path)
            matplotlib.rcParams['font.family'] = font_name
            plt.rcParams['font.family'] = font_name
        else:
            font_prop = fm.FontProperties()
    else:
        font_prop = fm.FontProperties()

configure_japanese_font()
st.set_page_config(page_title="QC分析ツール", layout="wide")

# --- ユーティリティ関数 ---
def parse_time(t):
    if pd.isna(t):
        return np.nan
    m = re.match(r"(\d+)\s*['’′](\d+)\s*[\"”″]?$", str(t).strip())
    if m:
        minutes, seconds = map(int, m.groups())
        return minutes * 60 + seconds
    return np.nan

def to_minsec(sec):
    if pd.isna(sec):
        return "―"
    m = int(sec // 60)
    s = int(sec % 60)
    return f"{m}'{s:02d}\""

def qc_judge(row, specs):
    for col, (lo, hi) in specs.items():
        v = row.get(col, None)
        if pd.isna(v):
            continue
        if (lo is not None and v < lo) or (hi is not None and v > hi):
            return "NG"
    return "OK"

def add_highlight_col(df):
    df = df.copy()
    if "QC結果" in df.columns:
        df["QC結果"] = df["QC結果"].apply(lambda x: "🟥 NG" if x == "NG" else "OK")
    return df

@st.cache_data(show_spinner="読み込み中…")
def load_excel(uploaded_file, ext):
    engine = "openpyxl" if ext in [".xlsx", ".xlsm"] else "xlrd"
    with pd.ExcelFile(uploaded_file, engine=engine) as xls:
        df = pd.read_excel(xls, sheet_name=0)
        spec_df = pd.read_excel(xls, sheet_name="規格マスタ")
    df.columns = df.columns.str.strip()
    spec_df.columns = spec_df.columns.str.strip()
    return df, spec_df

# --- UI開始 ---
st.sidebar.header("① ファイル選択")
uploaded = st.sidebar.file_uploader("Excelファイル (.xlsx, .xlsm, .xls)", type=["xls", "xlsx", "xlsm"])

if uploaded:
    ext = os.path.splitext(uploaded.name)[1].lower()
    df, spec_df = load_excel(uploaded, ext)

    if "水2L濾過時間" in df.columns:
        df["水2L濾過時間_sec"] = df["水2L濾過時間"].apply(parse_time)
        df["水2L濾過時間_mmss"] = df["水2L濾過時間_sec"].apply(to_minsec)

    if "ロットNo" in df.columns:
        df["ロットNo"] = df["ロットNo"].astype(str).str.zfill(3)
        df["ロットNo_3桁"] = df["ロットNo"].str.slice(0, 3)

    # 製品フィルタ
    st.sidebar.header("② 製品フィルタ")
    products = sorted(df["製品名称"].dropna().unique())
    if "select_all_products" not in st.session_state:
        st.session_state.select_all_products = False

    col1, col2 = st.sidebar.columns([1, 1])
    with col1:
        if st.button("全選択", key="select_all_btn"):
            st.session_state.select_all_products = True
    with col2:
        if st.button("全解除", key="clear_all_btn"):
            st.session_state.select_all_products = False

    default_selection = products if st.session_state.select_all_products else []
    sel_products = st.sidebar.multiselect("製品を選択", products, default=default_selection)

    if not sel_products:
        st.warning("製品名が選択されていません。左のサイドバーから製品名を選択してください。")
        st.stop()

    df_filtered = df[df["製品名称"].isin(sel_products)]

    # ロットフィルタ
    lot_mode = st.sidebar.checkbox("ロットNo別で集計・表示する（先頭3桁）")
    if lot_mode:
        all_lots = sorted(df_filtered["ロットNo_3桁"].dropna().unique())
        if "select_all_lots" not in st.session_state:
            st.session_state.select_all_lots = False
        col1, col2 = st.sidebar.columns([1, 1])
        with col1:
            if st.button("ロット全選択", key="select_all_lots_btn"):
                st.session_state.select_all_lots = True
        with col2:
            if st.button("ロット全解除", key="clear_all_lots_btn"):
                st.session_state.select_all_lots = False

        default_lot_selection = all_lots if st.session_state.select_all_lots else []
        selected_lots = st.sidebar.multiselect("対象ロットNo（3桁）を選択", all_lots, default=default_lot_selection)
        df_filtered = df_filtered[df_filtered["ロットNo_3桁"].isin(selected_lots)]
    # 規格値入力
    st.sidebar.header("③ 規格値設定")
    specs = {}
    dynamic_targets = []

    if len(sel_products) == 1:
        prod = sel_products[0]
        row = spec_df[spec_df["製品名称"] == prod]
        if not row.empty:
            row = row.iloc[0]
            limit_cols = [col for col in row.index if col.endswith("_下限") or col.endswith("_上限")]
            base_items = sorted(set(col.rsplit("_", 1)[0] for col in limit_cols))

            with st.sidebar.expander("📏 規格入力（編集可能）", expanded=True):
                for base in base_items:
                    lo_val = row.get(f"{base}_下限", None)
                    hi_val = row.get(f"{base}_上限", None)

                    if base == "水2L濾過時間(秒)":
                        display_label = "水2L濾過時間"
                        col_actual = "水2L濾過時間_sec"
                    else:
                        display_label = base
                        col_actual = base

                    c1, c2 = st.columns(2)
                    with c1:
                        lo = st.number_input(f"{display_label} 下限", value=lo_val if pd.notna(lo_val) else None,
                                             step=0.01, format="%.2f", key=f"{prod}_{col_actual}_lo")
                    with c2:
                        hi = st.number_input(f"{display_label} 上限", value=hi_val if pd.notna(hi_val) else None,
                                             step=0.01, format="%.2f", key=f"{prod}_{col_actual}_hi")

                    if col_actual in df_filtered.columns:
                        specs[col_actual] = (lo, hi)
                        dynamic_targets.append((display_label, col_actual))
        else:
            st.sidebar.warning("規格マスタに製品が見つかりません。")
    else:
        st.sidebar.warning("複数製品が選択されています。規格値は手動で入力してください。")
        all_numeric_cols = df_filtered.select_dtypes(include=np.number).columns.tolist()
        with st.sidebar.expander("📏 手動で規格値入力（複数製品用）", expanded=True):
            for col_actual in all_numeric_cols:
                display_label = col_actual if not col_actual.endswith("_sec") else "水2L濾過時間"
                c1, c2 = st.columns(2)
                with c1:
                    lo = st.number_input(f"{display_label} 下限", value=None, step=0.01, format="%.2f", key=f"{col_actual}_multi_lo")
                with c2:
                    hi = st.number_input(f"{display_label} 上限", value=None, step=0.01, format="%.2f", key=f"{col_actual}_multi_hi")

                specs[col_actual] = (lo, hi)
                dynamic_targets.append((display_label, col_actual))

    # QC判定とNG集計
    df_filtered["QC結果"] = df_filtered.apply(qc_judge, axis=1, args=(specs,))
    st.header("QC分析結果")

    if lot_mode:
        st.subheader("ロットNo別集計（3桁）")
        summary = df_filtered.groupby("ロットNo_3桁")["QC結果"].value_counts().unstack().fillna(0).astype(int)
        st.dataframe(summary)
    else:
        st.write(f"**NG件数：{(df_filtered['QC結果'] == 'NG').sum()} / {len(df_filtered)}**")

    st.subheader("詳細データ")
    show_only_ng = st.checkbox("⚠ NG（規格外）行のみ表示する", value=False)
    df_view = df_filtered[df_filtered["QC結果"] == "NG"] if show_only_ng else df_filtered.copy()
    filtered_count = len(df_view)
    ng_count = (df_view["QC結果"] == "NG").sum()
    ng_rate = (ng_count / filtered_count * 100) if filtered_count > 0 else 0.0
    st.markdown(f"**不良率：{ng_rate:.2f}% ({ng_count} / {filtered_count})**")
    df_view = add_highlight_col(df_view)
    st.dataframe(df_view, height=600, use_container_width=True)

    csv = df_filtered.to_csv(index=False).encode("utf-8-sig")
    st.download_button("分析結果をCSVでダウンロード", csv, file_name="qc_result.csv", mime="text/csv")

    # 統計・分布分析
    st.subheader("📊 統計・分布分析")
    for label, actual_col in dynamic_targets:
        values = df_filtered[actual_col].dropna()
        if values.empty:
            continue

        mean = values.mean()
        std = values.std()
        lower_3σ, upper_3σ = mean - 3 * std, mean + 3 * std
        lo, hi = specs.get(actual_col, (None, None))

        st.markdown(f"### {label}")
        if "濾過時間" in label:
            st.markdown(f"- Ave: **{to_minsec(mean)}**, ±σ: ±{int(std)} 秒")
            st.markdown(f"- 3σ range: {to_minsec(lower_3σ)} ～ {to_minsec(upper_3σ)}")
        else:
            st.markdown(f"- Ave: **{mean:.2f}**, ±σ: **{std:.2f}**")
            st.markdown(f"- 3σ range: {lower_3σ:.2f} ～ {upper_3σ:.2f}")

        with st.expander(f"📐 {label} x-axis range", expanded=False):
            default_min = min(values.min(), lower_3σ, lo if lo is not None else values.min())
            default_max = max(values.max(), upper_3σ, hi if hi is not None else values.max())
            x_min = st.number_input(f"{label} x-min", value=round(default_min - 0.5, 2), step=0.1, format="%.2f", key=f"{actual_col}_xmin")
            x_max = st.number_input(f"{label} x-max", value=round(default_max + 0.5, 2), step=0.1, format="%.2f", key=f"{actual_col}_xmax")

        bin_width = max((x_max - x_min) / 40, 0.01)
        bins = np.arange(x_min, x_max + bin_width, bin_width)

        fig, ax = plt.subplots(figsize=(7, 3.5))
        ax.hist(values, bins=bins, alpha=0.7, edgecolor='black')
        ax.axvline(mean, color='blue', linestyle='--', label='Ave')
        ax.axvline(lower_3σ, color='orange', linestyle=':', label='-3σ')
        ax.axvline(upper_3σ, color='orange', linestyle=':', label='+3σ')
        if lo is not None:
            ax.axvline(lo, color='red', linestyle='-', label='Spec Low')
        if hi is not None:
            ax.axvline(hi, color='red', linestyle='-', label='Spec High')
        ax.set_xlim(x_min, x_max)
        ax.set_title(f"{label} Distribution", fontproperties=font_prop)
        ax.legend(prop=font_prop)
        st.pyplot(fig)
    # Plotlyロット推移グラフ（製品別色分け対応）
    if lot_mode:
        st.subheader("📈 Lot Trend (3桁)")
        for label, actual_col in dynamic_targets:
            df_plot = df_filtered[["ロットNo_3桁", "製品名称", actual_col]].dropna()
            if df_plot.empty:
                continue
            df_plot = df_plot.groupby(["ロットNo_3桁", "製品名称"])[actual_col].mean().reset_index()
            fig = px.line(df_plot, x="ロットNo_3桁", y=actual_col, color="製品名称", markers=True,
                          title=f"{label} Lot Trend（製品別）")
            fig.update_layout(xaxis_title="Lot No（3桁）", yaxis_title=label)
            st.plotly_chart(fig, use_container_width=True)

    # 散布図＋Hexbinヒートマップ
    with st.expander("📉 任意項目で散布図・密度ヒートマップを表示", expanded=False):
        numeric_cols = df_filtered.select_dtypes(include=np.number).columns.tolist()
        if len(numeric_cols) < 2:
            st.info("数値列が2つ以上必要です。")
        else:
            col1, col2 = st.columns(2)
            with col1:
                x_axis = st.selectbox("X軸項目", numeric_cols, key="scatter_x")
            with col2:
                y_axis = st.selectbox("Y軸項目", numeric_cols, key="scatter_y")

            if x_axis == y_axis:
                st.warning("X軸とY軸は異なる項目を選択してください。")
            else:
                common = df_filtered[[x_axis, y_axis]].dropna()
                c1, c2 = st.columns(2)
                with c1:
                    x_min = st.number_input("X軸最小", value=round(common[x_axis].min(), 2), key="x_min")
                    x_max = st.number_input("X軸最大", value=round(common[x_axis].max(), 2), key="x_max")
                with c2:
                    y_min = st.number_input("Y軸最小", value=round(common[y_axis].min(), 2), key="y_min")
                    y_max = st.number_input("Y軸最大", value=round(common[y_axis].max(), 2), key="y_max")

                common_clipped = common[(common[x_axis] >= x_min) & (common[x_axis] <= x_max) &
                                        (common[y_axis] >= y_min) & (common[y_axis] <= y_max)]

                col1, col2 = st.columns(2)
                with col1:
                    fig1, ax1 = plt.subplots(figsize=(6, 4))
                    ax1.scatter(common_clipped[x_axis], common_clipped[y_axis], alpha=0.2,
                                edgecolors='black', linewidths=0.5)
                    ax1.set_title("散布図", fontproperties=font_prop)
                    ax1.set_xlabel(x_axis, fontproperties=font_prop)
                    ax1.set_ylabel(y_axis, fontproperties=font_prop)
                    ax1.set_xlim(x_min, x_max)
                    ax1.set_ylim(y_min, y_max)
                    ax1.grid(True)
                    st.pyplot(fig1)

                with col2:
                    fig2, ax2 = plt.subplots(figsize=(6, 4))
                    hb = ax2.hexbin(common_clipped[x_axis], common_clipped[y_axis],
                                    gridsize=50, cmap='Blues', mincnt=1)
                    cb = fig2.colorbar(hb, ax=ax2)
                    cb.set_label("Count")
                    ax2.set_title("Hexbin ヒートマップ", fontproperties=font_prop)
                    ax2.set_xlabel(x_axis, fontproperties=font_prop)
                    ax2.set_ylabel(y_axis, fontproperties=font_prop)
                    ax2.set_xlim(x_min, x_max)
                    ax2.set_ylim(y_min, y_max)
                    ax2.grid(True)
                    st.pyplot(fig2)
else:
    st.info("左サイドバーから .xls, .xlsx, .xlsm ファイルをアップロードしてください。")

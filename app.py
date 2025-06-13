import streamlit as st
import pandas as pd
import numpy as np
import re
import matplotlib.pyplot as plt
import matplotlib
import os

try:
    matplotlib.rcParams['font.family'] = 'Yu Gothic'
except:
    pass

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

# --- UI ---
st.sidebar.header("① ファイル選択")
uploaded = st.sidebar.file_uploader("Excelファイル (.xlsx, .xlsm, .xls)", type=["xls", "xlsx", "xlsm"])

if uploaded:
    with st.spinner("読み込み中…"):
        ext = os.path.splitext(uploaded.name)[1].lower()
        engine = "openpyxl" if ext in [".xlsx", ".xlsm"] else "xlrd"
        with pd.ExcelFile(uploaded, engine=engine) as xls:
            df = pd.read_excel(xls, sheet_name=0)
            spec_df = pd.read_excel(xls, sheet_name="規格マスタ")
        df.columns = df.columns.str.strip()
        spec_df.columns = spec_df.columns.str.strip()

    if "水2L濾過時間" in df.columns:
        df["水2L濾過時間_sec"] = df["水2L濾過時間"].apply(parse_time)
        df["水2L濾過時間_mmss"] = df["水2L濾過時間_sec"].apply(to_minsec)

    st.sidebar.header("② 製品フィルタ")
    products = sorted(df["製品名称"].dropna().unique())
    sel_products = st.sidebar.multiselect("製品を選択", products, default=products[:1])
    df = df[df["製品名称"].isin(sel_products)]

    if not sel_products:
        st.warning("製品名が選択されていません。左のサイドバーから製品名を選択してください。")
    else:
        lot_mode = st.sidebar.checkbox("ロットNo別で集計・表示する")
        if lot_mode:
            all_lots = sorted(df["ロットNo"].dropna().unique())
            selected_lots = st.sidebar.multiselect("対象ロットNoを選択", all_lots, default=[])
            df = df[df["ロットNo"].isin(selected_lots)]

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

                        if col_actual in df.columns:
                            specs[col_actual] = (lo, hi)
                            dynamic_targets.append((display_label, col_actual))
            else:
                st.sidebar.warning("規格マスタに製品が見つかりません。")

        # ✅ QC結果を再計算
        df["QC結果"] = df.apply(qc_judge, axis=1, args=(specs,))

        st.header("QC分析結果")
        if lot_mode:
            st.subheader("ロットNo別集計")
            summary = df.groupby("ロットNo")["QC結果"].value_counts().unstack().fillna(0).astype(int)
            st.dataframe(summary)
        else:
            st.write(f"**NG件数：{(df['QC結果'] == 'NG').sum()} / {len(df)}**")

        # ✅ 詳細データ（手動表示）
        st.subheader("詳細データ")
        show_only_ng = st.checkbox("⚠ NG（規格外）行のみ表示する", value=False)

        df_view = df.copy()
        if show_only_ng:
            df_view = df_view[df_view["QC結果"] == "NG"]

        filtered_count = len(df_view)
        ng_count = (df_view["QC結果"] == "NG").sum()
        ng_rate = (ng_count / filtered_count * 100) if filtered_count > 0 else 0.0
        st.markdown(f"**不良率：{ng_rate:.2f}% ({ng_count} / {filtered_count})**")

        if st.button("📋 詳細データを表示"):
            df_view = add_highlight_col(df_view)
            st.dataframe(df_view, height=600, use_container_width=True)

            csv = df.to_csv(index=False).encode("utf-8-sig")
            st.download_button("分析結果をCSVでダウンロード", csv, file_name="qc_result.csv", mime="text/csv")

        st.subheader("📊 統計・分布分析")
        for label, actual_col in dynamic_targets:
            if actual_col not in df.columns or df[actual_col].dropna().empty:
                continue

            values = df[actual_col].dropna()
            mean = values.mean()
            std = values.std()
            lower_3σ, upper_3σ = mean - 3 * std, mean + 3 * std
            lo, hi = specs.get(actual_col, (None, None))

            st.markdown(f"### {label}")
            if "濾過時間" in label:
                st.markdown(f"- 平均: **{to_minsec(mean)}**, 標準偏差: ±{int(std)} 秒")
                st.markdown(f"- 3σ範囲: {to_minsec(lower_3σ)} ～ {to_minsec(upper_3σ)}")
            else:
                st.markdown(f"- 平均: **{mean:.2f}**, 標準偏差: **{std:.2f}**")
                st.markdown(f"- 3σ範囲: {lower_3σ:.2f} ～ {upper_3σ:.2f}")

            key_show = f"show_plot_{actual_col}"
            if key_show not in st.session_state:
                st.session_state[key_show] = False

            if st.button(f"{label} の分布グラフを描画", key=f"btn_{actual_col}"):
                st.session_state[key_show] = True

            if st.session_state[key_show]:
                with st.expander(f"📐 {label} のx軸表示範囲を調整", expanded=False):
                    default_min = min(values.min(), lower_3σ, lo if lo is not None else values.min())
                    default_max = max(values.max(), upper_3σ, hi if hi is not None else values.max())
                    x_min = st.number_input(f"{label} x軸最小", value=float(round(default_min - 0.5, 2)),
                                            step=0.1, format="%.2f", key=f"{actual_col}_xmin")
                    x_max = st.number_input(f"{label} x軸最大", value=float(round(default_max + 0.5, 2)),
                                            step=0.1, format="%.2f", key=f"{actual_col}_xmax")

                fig, ax = plt.subplots(figsize=(7, 3.5))
                ax.hist(values, bins="auto", alpha=0.7, edgecolor='black')
                ax.axvline(mean, color='blue', linestyle='--', label='平均')
                ax.axvline(lower_3σ, color='orange', linestyle=':', label='-3σ')
                ax.axvline(upper_3σ, color='orange', linestyle=':', label='+3σ')
                if lo is not None:
                    ax.axvline(lo, color='red', linestyle='-', label='規格下限')
                if hi is not None:
                    ax.axvline(hi, color='red', linestyle='-', label='規格上限')
                ax.set_xlim(x_min, x_max)
                ax.set_title(f"{label} 分布（選択ロット）")
                ax.legend()
                st.pyplot(fig)

        if lot_mode:
            st.subheader("📈 ロット順推移（折れ線グラフ）")
            for label, actual_col in dynamic_targets:
                if actual_col not in df.columns or df[actual_col].dropna().empty:
                    continue

                key_line = f"show_line_{actual_col}"
                if key_line not in st.session_state:
                    st.session_state[key_line] = False

                if st.button(f"{label} のロット推移グラフを描画", key=f"btn_line_{actual_col}"):
                    st.session_state[key_line] = True

                if st.session_state[key_line]:
                    df_plot = df[["ロットNo", actual_col]].dropna().copy()
                    df_plot = df_plot.groupby("ロットNo")[actual_col].mean().reset_index()
                    df_plot = df_plot.sort_values("ロットNo")

                    fig, ax = plt.subplots(figsize=(7, 3.5))
                    ax.plot(df_plot["ロットNo"], df_plot[actual_col], marker='o')
                    ax.set_xlabel("ロットNo")
                    ax.set_ylabel(label)
                    ax.set_title(f"{label} のロット推移")
                    ax.grid(True)
                    st.pyplot(fig)
else:
    st.info("左サイドバーから .xls, .xlsx, .xlsm ファイルをアップロードしてください。")

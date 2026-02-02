"""
Keepa Excel結合ツール

実行方法:
    streamlit run app.py

機能:
    - keepa-*.xlsxファイルのアップロード（複数選択可）
    - ASIN列の追加（全カラムが右にシフト）
    - 複数ファイルの縦結合
    - CSVエクスポート
"""

import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime
import io

st.set_page_config(page_title="Keepa Excel結合ツール", page_icon="📊", layout="wide")

st.title("📊 Keepa Excel結合ツール")
st.markdown("複数の `keepa-*.xlsx` ファイルをアップロードし、ASIN列を追加して1つのCSVに結合します。")

# ===== セッション状態の初期化 =====
if 'merged_df' not in st.session_state:
    st.session_state.merged_df = None
if 'file_list' not in st.session_state:
    st.session_state.file_list = []

# ===== ファイルアップロード =====
st.subheader("📤 ファイルをアップロード")

uploaded_files = st.file_uploader(
    "keepa-*.xlsx ファイルを選択してください（複数選択可）",
    type=["xlsx"],
    accept_multiple_files=True,
    help="Keepa形式のExcelファイルをアップロード"
)

if uploaded_files:
    file_info = []
    for f in uploaded_files:
        # シート名からASINを取得（Note以外の最初のシート）
        try:
            sheet_names = pd.ExcelFile(f).sheet_names
            asin = next((s for s in sheet_names if s.lower() != "note"), "不明")
        except Exception:
            asin = "不明"

        size_mb = len(f.getvalue()) / (1024 * 1024)

        file_info.append({
            "ファイル名": f.name,
            "ASIN": asin,
            "サイズ (MB)": f"{size_mb:.2f}",
            "ファイルオブジェクト": f
        })

    st.session_state.file_list = file_info
    st.success(f"✅ {len(uploaded_files)} 件のファイルをアップロードしました")

# ===== ファイルリスト表示 =====
if st.session_state.file_list:
    st.subheader(f"📋 検出ファイル一覧 ({len(st.session_state.file_list)} 件)")

    # 表示用にパスを除外
    display_df = pd.DataFrame(st.session_state.file_list)
    display_columns = ["ファイル名", "ASIN", "サイズ (MB)"]
    st.dataframe(display_df[display_columns], use_container_width=True)

    # ===== 結合処理 =====
    st.divider()

    if st.button("🔗 結合実行", type="primary", use_container_width=True):
        all_data = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, file_info in enumerate(st.session_state.file_list):
            try:
                # シート名から正確なASINを取得
                status_text.text(f"処理中: {file_info['ファイル名']}")

                # ファイル読み込み（アップロードモードのみ）
                wb = openpyxl.load_workbook(
                    io.BytesIO(file_info["ファイルオブジェクト"].getvalue()),
                    data_only=True
                )

                # ASINシートを探す（Noteシート以外の最初のシート = ASIN名）
                asin = next((name for name in wb.sheetnames if name.lower() != "note"), None)

                if not asin:
                    st.warning(f"⚠️ {file_info['ファイル名']}: データシートが見つかりません（Noteシート以外）")
                    continue

                # シート読み込み
                file_info["ファイルオブジェクト"].seek(0)
                df = pd.read_excel(file_info["ファイルオブジェクト"], sheet_name=asin)

                # A列にASINを追加（既存カラムを右にシフト）
                df.insert(0, "ASIN", asin)

                all_data.append(df)

            except Exception as e:
                st.error(f"❌ {file_info['ファイル名']}: エラーが発生しました - {str(e)}")
                continue

            # 進捗更新
            progress_bar.progress((idx + 1) / len(st.session_state.file_list))

        if all_data:
            # 縦結合（sort=FalseでBSRカラム等の差異も全て保持）
            st.session_state.merged_df = pd.concat(all_data, ignore_index=True, sort=False)
            status_text.text("✅ 結合完了!")
            progress_bar.empty()
            st.success(f"🎉 結合完了: {len(all_data)} ファイル → {len(st.session_state.merged_df)} 行")
        else:
            status_text.text("❌ 結合できるデータがありませんでした")
            progress_bar.empty()

# ===== 結合結果表示 =====
if st.session_state.merged_df is not None:
    st.divider()
    st.subheader("📊 結合結果")

    # 日付カラムの存在確認と型変換
    date_column = None
    if "日付" in st.session_state.merged_df.columns:
        date_column = "日付"
        st.session_state.merged_df[date_column] = pd.to_datetime(
            st.session_state.merged_df[date_column], errors='coerce'
        )
    elif "Date" in st.session_state.merged_df.columns:
        date_column = "Date"
        st.session_state.merged_df[date_column] = pd.to_datetime(
            st.session_state.merged_df[date_column], errors='coerce'
        )

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("総行数", f"{len(st.session_state.merged_df):,}")
    with col2:
        unique_asins = st.session_state.merged_df["ASIN"].nunique()
        st.metric("ASIN数", unique_asins)
    with col3:
        st.metric("カラム数", len(st.session_state.merged_df.columns))

    # 日付範囲フィルター
    filtered_df = st.session_state.merged_df.copy()

    if date_column and st.session_state.merged_df[date_column].notna().any():
        st.divider()
        st.subheader("📅 日付範囲フィルター")

        min_date = st.session_state.merged_df[date_column].min().date()
        max_date = st.session_state.merged_df[date_column].max().date()

        col_date1, col_date2 = st.columns(2)
        with col_date1:
            start_date = st.date_input(
                "開始日",
                value=min_date,
                min_value=min_date,
                max_value=max_date,
                help="この日付以降のデータを抽出"
            )
        with col_date2:
            end_date = st.date_input(
                "終了日",
                value=max_date,
                min_value=min_date,
                max_value=max_date,
                help="この日付以前のデータを抽出"
            )

        # フィルタリング実行
        if start_date <= end_date:
            mask = (
                (st.session_state.merged_df[date_column].dt.date >= start_date) &
                (st.session_state.merged_df[date_column].dt.date <= end_date)
            )
            filtered_df = st.session_state.merged_df[mask].copy()

            if len(filtered_df) < len(st.session_state.merged_df):
                st.info(f"📊 フィルター結果: {len(filtered_df):,} 行 / {len(st.session_state.merged_df):,} 行")
        else:
            st.error("⚠️ 開始日は終了日より前に設定してください")

    # プレビュー
    st.divider()
    st.markdown("**プレビュー（先頭10行）**")
    st.dataframe(filtered_df.head(10), use_container_width=True)

    # ダウンロード
    st.divider()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_filename = f"keepa_merged_{timestamp}.csv"

    csv_buffer = io.StringIO()
    filtered_df.to_csv(csv_buffer, index=False, encoding="utf-8-sig")
    csv_data = csv_buffer.getvalue()

    st.download_button(
        label="💾 CSVダウンロード",
        data=csv_data,
        file_name=csv_filename,
        mime="text/csv",
        type="primary",
        use_container_width=True
    )

    st.info(f"📥 ダウンロードファイル名: `{csv_filename}` ({len(filtered_df):,} 行)")

# ===== フッター =====
st.divider()
st.caption("📝 Tips: 複数ファイルを一度に選択できます（Ctrl/Cmd + クリック）。")

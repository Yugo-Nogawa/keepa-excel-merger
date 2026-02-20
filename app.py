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
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
from pathlib import Path
import io

# ===== セール情報読み込み =====
def load_sale_periods():
    """CSVからセール情報を読み込み"""
    csv_path = Path(__file__).parent / "sale_periods.csv"
    try:
        df = pd.read_csv(csv_path, encoding='utf-8-sig')
        # タプルのリストに変換
        return [(row['開始日'], row['終了日'], row['セール分類']) for _, row in df.iterrows()]
    except FileNotFoundError:
        st.error(f"⚠️ セール情報ファイルが見つかりません: {csv_path}")
        return []
    except Exception as e:
        st.error(f"⚠️ セール情報の読み込みエラー: {str(e)}")
        return []

SALE_PERIODS = load_sale_periods()

def classify_sale(target_date):
    """日付からセール分類を判定"""
    if pd.isna(target_date):
        return None

    if isinstance(target_date, str):
        target_date = pd.to_datetime(target_date).date()
    elif hasattr(target_date, 'date'):
        target_date = target_date.date()

    for start_str, end_str, sale_type in SALE_PERIODS:
        start = datetime.strptime(start_str, "%Y-%m-%d").date()
        end = datetime.strptime(end_str, "%Y-%m-%d").date()
        if start <= target_date <= end:
            return sale_type

    return None

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

                # B列（日付カラム）の存在確認とセール分類追加
                date_col = None
                if "日付" in df.columns:
                    date_col = "日付"
                elif "Date" in df.columns:
                    date_col = "Date"

                if date_col:
                    # 日付型に変換
                    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
                    # C列にセール分類を追加
                    df.insert(2, "セール分類", df[date_col].apply(classify_sale))

                # カラムの整理と追加
                # 定価: FBA価格とList価格の最大値
                fba_col = "FBA 価格(￥)" if "FBA 価格(￥)" in df.columns else "FBA価格(￥)"
                list_col = "List 価格(￥)" if "List 価格(￥)" in df.columns else "List価格(￥)"

                if fba_col in df.columns and list_col in df.columns:
                    df["定価"] = df[[fba_col, list_col]].max(axis=1)
                elif fba_col in df.columns:
                    df["定価"] = df[fba_col]
                elif list_col in df.columns:
                    df["定価"] = df[list_col]

                # 販売価格: Buybox価格
                buybox_col = "Buybox 価格(￥)" if "Buybox 価格(￥)" in df.columns else "Buybox価格(￥)"
                if buybox_col in df.columns:
                    df["販売価格"] = df[buybox_col]

                # サブカテゴリーBSR: BSR[****]系カラムの最小値
                bsr_columns = [col for col in df.columns if col.startswith("BSR[") and col.endswith("]")]
                if bsr_columns:
                    df["サブカテゴリーBSR"] = df[bsr_columns].min(axis=1)

                # 不要カラムの削除
                cols_to_drop = [
                    "Buybox 価格(￥)", "Buybox価格(￥)",
                    "価格(￥)",
                    "Prime 価格(￥)", "Prime価格(￥)",
                    "Coupon 価格(￥)", "Coupon価格(￥)",
                    "Coupon 割引", "Coupon割引",
                    "Deal 価格(￥)", "Deal価格(￥)",
                    "Deal 価格情報", "Deal価格情報",
                    "FBA 価格(￥)", "FBA価格(￥)",
                    "FBM 価格(￥)", "FBM価格(￥)",
                    "List 価格(￥)", "List価格(￥)",
                    "販売数(子)",
                    "評価", "評価数", "セラー数"
                ]
                # 存在するカラムのみ削除
                cols_to_drop_existing = [col for col in cols_to_drop if col in df.columns]
                df.drop(columns=cols_to_drop_existing, inplace=True)

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

        # デフォルト開始日: 12ヶ月前の月初
        today = datetime.now().date()
        default_start = (today.replace(day=1) - relativedelta(months=12))
        # データの範囲内に収める
        default_start = max(default_start, min_date)

        col_date1, col_date2 = st.columns(2)
        with col_date1:
            start_date = st.date_input(
                "開始日",
                value=default_start,
                min_value=min_date,
                max_value=max_date,
                help="この日付以降のデータを抽出（デフォルト: 12ヶ月前の月初）"
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

    # サマリデータの生成
    st.divider()
    st.subheader("📈 サマリデータ")

    summary_data = []

    if "セール分類" in filtered_df.columns and "定価" in filtered_df.columns and "販売価格" in filtered_df.columns:
        # ASIN × セール種別でグループ化
        for asin in filtered_df["ASIN"].unique():
            asin_df = filtered_df[filtered_df["ASIN"] == asin]

            # 直近のサブカテゴリーBSR（セール関係なく最新日付）
            latest_subcategory_bsr = None
            if "サブカテゴリーBSR" in asin_df.columns:
                latest_row = asin_df.sort_values(date_column, ascending=False).iloc[0]
                latest_subcategory_bsr = latest_row["サブカテゴリーBSR"]

            # セール種別ごとの集計
            for sale_type in ["MDE", "ビッグセール", "ビッグセールのアーリー"]:
                sale_df = asin_df[asin_df["セール分類"] == sale_type].copy()

                if len(sale_df) > 0:
                    # 参加判定: 定価から5%以上値下げした日
                    sale_df["値下げ率"] = (sale_df["定価"] - sale_df["販売価格"]) / sale_df["定価"]
                    participated_df = sale_df[sale_df["値下げ率"] >= 0.05]

                    # セール期間内の総日数（フィルター範囲内）
                    total_days = len(sale_df)

                    # 実際に参加した日数
                    participated_days = len(participated_df)

                    # 参加頻度（%）
                    participation_rate = (participated_days / total_days * 100) if total_days > 0 else 0

                    # 定価（参加日の最頻値または平均）
                    list_price = None
                    if len(participated_df) > 0:
                        list_price = participated_df["定価"].mode()[0] if not participated_df["定価"].mode().empty else participated_df["定価"].mean()
                    else:
                        list_price = sale_df["定価"].mode()[0] if not sale_df["定価"].mode().empty else sale_df["定価"].mean()

                    # 最安値・最高値セール売価（参加日のみ）
                    min_price = participated_df["販売価格"].min() if len(participated_df) > 0 else None
                    max_price = participated_df["販売価格"].max() if len(participated_df) > 0 else None

                    summary_data.append({
                        "ASIN": asin,
                        "参加セール種別": sale_type,
                        "カテゴリランク（直近）": latest_subcategory_bsr,
                        "参加頻度（%）": round(participation_rate, 1),
                        "定価": list_price,
                        "最安値セール売価": min_price,
                        "最高値セール売価": max_price
                    })

    summary_df = pd.DataFrame(summary_data)

    if not summary_df.empty:
        st.dataframe(summary_df, use_container_width=True)

        # サマリCSVダウンロード
        summary_csv_buffer = io.StringIO()
        summary_df.to_csv(summary_csv_buffer, index=False, encoding="utf-8-sig")
        summary_csv_data = summary_csv_buffer.getvalue()

        st.download_button(
            label="📊 サマリCSVダウンロード",
            data=summary_csv_data,
            file_name=f"keepa_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
            use_container_width=True
        )
    else:
        st.info("⚠️ サマリデータがありません（セール分類が設定されていない可能性があります）")

    # 詳細データフィルター
    st.divider()
    st.subheader("🔍 詳細データフィルター")

    download_df = filtered_df.copy()

    # BSR[***]形式のカラムを検出
    bsr_columns = [col for col in download_df.columns if col.startswith("BSR[") and col.endswith("]")]

    # 各レコードの所属カテゴリーを判定（最小BSR値を持つカテゴリー）
    if bsr_columns:
        def get_primary_category(row):
            """各行について、最小BSR値を持つカテゴリー名を返す"""
            min_val = None
            min_category = None
            for col in bsr_columns:
                val = row[col]
                if pd.notna(val) and (min_val is None or val < min_val):
                    min_val = val
                    min_category = col[4:-1]  # "BSR[カテゴリー名]" から カテゴリー名 を抽出
            return min_category

        download_df["サブカテゴリー"] = download_df.apply(get_primary_category, axis=1)

        # カテゴリー一覧を取得（NaN除外）
        available_categories = sorted(download_df["サブカテゴリー"].dropna().unique())

        if available_categories:
            col_cat, col_bsr = st.columns(2)

            with col_cat:
                selected_categories = st.multiselect(
                    "属するサブカテゴリー",
                    options=available_categories,
                    default=None,
                    help="複数選択可能（OR条件）。選択したカテゴリーのいずれかに所属するレコードを抽出"
                )

            with col_bsr:
                # サブカテゴリーBSR範囲フィルター
                if "サブカテゴリーBSR" in download_df.columns:
                    bsr_values = download_df["サブカテゴリーBSR"].dropna()
                    if len(bsr_values) > 0:
                        min_bsr = int(bsr_values.min())
                        max_bsr = int(bsr_values.max())

                        bsr_range = st.slider(
                            "サブカテゴリーBSR範囲",
                            min_value=min_bsr,
                            max_value=max_bsr,
                            value=(min_bsr, max_bsr),
                            help="この範囲内のBSRを持つレコードを抽出"
                        )
                    else:
                        bsr_range = None
                else:
                    bsr_range = None

            # 大カテゴリBSR範囲フィルター
            main_bsr_range = None
            if "BSR" in download_df.columns:
                main_bsr_values = download_df["BSR"].dropna()
                if len(main_bsr_values) > 0:
                    main_bsr_min = int(main_bsr_values.min())
                    main_bsr_max = int(main_bsr_values.max())

                    main_bsr_range = st.slider(
                        "大カテゴリBSR範囲",
                        min_value=main_bsr_min,
                        max_value=main_bsr_max,
                        value=(main_bsr_min, main_bsr_max),
                        help="この範囲内の大カテゴリ（全体）BSRを持つレコードを抽出"
                    )

            # フィルタリング適用
            filter_applied = False

            # カテゴリーフィルター
            if selected_categories:
                download_df = download_df[download_df["サブカテゴリー"].isin(selected_categories)]
                filter_applied = True

            # サブカテゴリーBSR範囲フィルター
            if bsr_range and "サブカテゴリーBSR" in download_df.columns:
                download_df = download_df[
                    (download_df["サブカテゴリーBSR"] >= bsr_range[0]) &
                    (download_df["サブカテゴリーBSR"] <= bsr_range[1])
                ]
                filter_applied = True

            # 大カテゴリBSR範囲フィルター
            if main_bsr_range and "BSR" in download_df.columns:
                download_df = download_df[
                    (download_df["BSR"] >= main_bsr_range[0]) &
                    (download_df["BSR"] <= main_bsr_range[1])
                ]
                filter_applied = True

            if filter_applied:
                st.info(f"📊 フィルター結果: {len(download_df):,} 行 / {len(filtered_df):,} 行")

    # 詳細データのカラムを整理（必要なカラムのみ残す）
    detail_columns = []
    if "ASIN" in download_df.columns:
        detail_columns.append("ASIN")
    if date_column:
        detail_columns.append(date_column)
    if "セール分類" in download_df.columns:
        detail_columns.append("セール分類")
    if "BSR" in download_df.columns:
        detail_columns.append("BSR")
    if "サブカテゴリー" in download_df.columns:
        detail_columns.append("サブカテゴリー")
    if "サブカテゴリーBSR" in download_df.columns:
        detail_columns.append("サブカテゴリーBSR")
    if "定価" in download_df.columns:
        detail_columns.append("定価")
    if "販売価格" in download_df.columns:
        detail_columns.append("販売価格")

    # 存在するカラムのみでフィルタリング
    detail_columns = [col for col in detail_columns if col in download_df.columns]
    download_df = download_df[detail_columns]

    # プレビュー
    st.divider()
    st.markdown("**詳細データプレビュー（先頭10行）**")

    # 日付カラムを日付のみの表示に変換
    preview_df = download_df.head(10).copy()
    if date_column and date_column in preview_df.columns:
        preview_df[date_column] = preview_df[date_column].dt.strftime('%Y-%m-%d')

    st.dataframe(preview_df, use_container_width=True)

    # 詳細データダウンロード
    st.divider()
    st.subheader("💾 詳細データダウンロード")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_filename = f"keepa_merged_{timestamp}.csv"

    csv_buffer = io.StringIO()
    download_df.to_csv(csv_buffer, index=False, encoding="utf-8-sig")
    csv_data = csv_buffer.getvalue()

    st.download_button(
        label="💾 詳細CSVダウンロード",
        data=csv_data,
        file_name=csv_filename,
        mime="text/csv",
        type="primary",
        use_container_width=True
    )

    st.info(f"📥 ダウンロードファイル名: `{csv_filename}` ({len(download_df):,} 行)")

# ===== フッター =====
st.divider()
st.caption("📝 Tips: 複数ファイルを一度に選択できます（Ctrl/Cmd + クリック）。")

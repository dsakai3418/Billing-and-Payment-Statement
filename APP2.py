import streamlit as st
import pandas as pd
from datetime import datetime
import io

# --- Streamlit UI ---
st.title("請求書管理アプリ")
st.write("NP掛け払いとバクラク請求書のCSVを処理し、共通フォーマットで出力します。")

# --- NP掛け払いCSVアップロード ---
st.header("1. NP掛け払いCSVのアップロード")
uploaded_np_file = st.file_uploader("NP掛け払いCSVファイルをアップロードしてください", type=["csv"], key="np_uploader")

np_df_processed = pd.DataFrame()
if uploaded_np_file is not None:
    try:
        np_df_raw = pd.read_csv(uploaded_np_file, encoding='shift_jis') # 仮定: Shift-JIS
    except UnicodeDecodeError:
        try:
            np_df_raw = pd.read_csv(uploaded_np_file, encoding='utf-8')
        except Exception as e:
            st.error(f"NP掛け払いCSVの読み込み中にエラーが発生しました: {e}。エンコーディングを確認してください。")
            np_df_raw = pd.DataFrame()
            
    if not np_df_raw.empty:
        st.subheader("NP掛け払いCSV内容（先頭5行）")
        st.dataframe(np_df_raw.head())

        # NP掛け払いデータ整形
        try:
            # ここで仮定しているカラム名が実際のCSVと異なる場合、KeyErrorが発生します。
            # 実際のCSVのカラム名に合わせて修正してください。
            required_np_cols = ['請求番号', '顧客名', '請求金額', '入金ステータス', '請求日付', 'お支払期日'] # 仮の必須カラム
            if not all(col in np_df_raw.columns for col in required_np_cols):
                st.error(f"NP掛け払いCSVに必要なカラム({', '.join(required_np_cols)})が見つかりません。CSVファイルを確認してください。")
                np_df_raw = pd.DataFrame() # 必要なカラムがなければ処理を中断
            else:
                np_df_processed = np_df_raw.copy()
                np_df_processed['入金有無'] = np_df_processed['入金ステータス'].apply(lambda x: 'あり' if x == '入金完了' else 'なし')
                np_df_processed['未入金金額合計(税込)'] = np_df_processed.apply(
                    lambda row: row['請求金額'] if row['入金有無'] == 'なし' else 0, axis=1
                )
                np_df_processed['ご請求金額合計(税込)'] = np_df_processed['請求金額']
                np_df_processed['ご利用年月'] = pd.to_datetime(np_df_processed['請求日付']).dt.strftime('%Y年%m月') # 例
                np_df_processed['ご請求方法'] = 'NP掛け払い' # 固定値
                
                # 最終フォーマットに合わせるためのカラム選択とリネーム
                np_df_processed = np_df_processed.rename(columns={
                    '請求番号': '請求書番号',
                    '請求日付': '請求書発行日', # 仮定：請求日付を請求書発行日とする
                    'お支払期日': 'お支払期日', # 仮定：そのまま使用
                })
                # フォーマットにないカラムは削除
                np_df_processed = np_df_processed[[
                    'ご利用年月', 'ご請求方法', 'ご請求金額合計(税込)', '未入金金額合計(税込)', 
                    '請求書番号', '請求書発行日', 'お支払期日', '入金有無'
                ]]
                
                st.subheader("NP掛け払いデータ (整形後)")
                st.dataframe(np_df_processed.head())

        except KeyError as e:
            st.warning(f"NP掛け払いCSVの処理中にカラム'{e}'が見つかりません。CSVのカラム名を確認してください。")
            np_df_processed = pd.DataFrame() # エラー時は空に

# --- バクラク請求書CSVアップロード ---
st.header("2. バクラク請求書CSVのアップロード")
uploaded_bakuraku_file = st.file_uploader("バクラク請求書CSVファイルをアップロードしてください", type=["csv"], key="bakuraku_uploader")

bakuraku_df_processed = pd.DataFrame()

if uploaded_bakuraku_file is not None:
    try:
        bakuraku_df_raw = pd.read_csv(uploaded_bakuraku_file, encoding='utf-8-sig') # 仮定: UTF-8-sig
    except UnicodeDecodeError:
        try:
            bakuraku_df_raw = pd.read_csv(uploaded_bakuraku_file, encoding='utf-8')
        except Exception as e:
            st.error(f"バクラク請求書CSVの読み込み中にエラーが発生しました: {e}。エンコーディングを確認してください。")
            bakuraku_df_raw = pd.DataFrame()

    if not bakuraku_df_raw.empty:
        st.subheader("バクラク請求書CSV内容（先頭5行）")
        st.dataframe(bakuraku_df_raw.head())

        # ここで仮定しているカラム名が実際のCSVと異なる場合、KeyErrorが発生します。
        # 実際のCSVのカラム名に合わせて修正してください。
        required_bakuraku_cols = ['書類番号', '送付先名', '金額', '日付', '支払期日'] # 仮の必須カラム
        if not all(col in bakuraku_df_raw.columns for col in required_bakuraku_cols):
            st.error(f"バクラクCSVに必要なカラム({', '.join(required_bakuraku_cols)})が見つかりません。CSVファイルを確認してください。")
            bakuraku_df_raw = pd.DataFrame()
        else:
            bakuraku_temp_df = bakuraku_df_raw.copy()
            bakuraku_temp_df['入金状況'] = '未入金（ユーザー未選択）' # 初期値

            bakuraku_temp_df['請求識別子'] = bakuraku_temp_df['書類番号'].astype(str) + " - " + \
                                          bakuraku_temp_df['送付先名'].astype(str) + " - " + \
                                          bakuraku_temp_df['金額'].astype(str)

            # Step 2a: 入金済み請求の選択
            st.subheader("2a. 入金済みと認識するバクラク請求書を選択してください")
            selected_paid_invoices_ids = st.multiselect(
                "入金済みの請求を選択",
                options=bakuraku_temp_df['請求識別子'].tolist(),
                help="入金が完了している請求を選択してください。",
                key="bakuraku_paid_select"
            )
            bakuraku_temp_df.loc[bakuraku_temp_df['請求識別子'].isin(selected_paid_invoices_ids), '入金状況'] = '入金済み（ユーザー選択）'
            
            st.subheader("現在のバクラク請求書入金状況")
            st.dataframe(bakuraku_temp_df[['書類番号', '送付先名', '金額', '入金状況']].head())

            # Step 2b: 出力から除外する請求書を選択
            st.subheader("2b. 出力から除外するバクラク請求書を選択してください")
            excluded_from_output_ids = st.multiselect(
                "出力から除外する請求書を選択",
                options=bakuraku_temp_df['請求識別子'].tolist(),
                help="Excel/CSVファイルに出力したくない請求書を選択してください。",
                key="bakuraku_exclude_select"
            )
            bakuraku_temp_df = bakuraku_temp_df[~bakuraku_temp_df['請求識別子'].isin(excluded_from_output_ids)].copy()

            if not bakuraku_temp_df.empty:
                # バクラクデータ整形
                bakuraku_df_processed = bakuraku_temp_df.copy()
                bakuraku_df_processed['入金有無'] = bakuraku_df_processed['入金状況'].apply(lambda x: 'あり' if x.startswith('入金済み') else 'なし')
                bakuraku_df_processed['未入金金額合計(税込)'] = bakuraku_df_processed.apply(
                    lambda row: row['金額'] if row['入金有無'] == 'なし' else 0, axis=1
                )
                bakuraku_df_processed['ご請求金額合計(税込)'] = bakuraku_df_processed['金額']
                bakuraku_df_processed['ご利用年月'] = pd.to_datetime(bakuraku_df_processed['日付']).dt.strftime('%Y年%m月') # 日付からご利用年月
                bakuraku_df_processed['ご請求方法'] = '直接請求' # 仮定：バクラクは直接請求

                # 最終フォーマットに合わせるためのカラム選択とリネーム
                bakuraku_df_processed = bakuraku_df_processed.rename(columns={
                    '書類番号': '請求書番号',
                    '日付': '請求書発行日',
                    '支払期日': 'お支払期日',
                })
                bakuraku_df_processed = bakuraku_df_processed[[
                    'ご利用年月', 'ご請求方法', 'ご請求金額合計(税込)', '未入金金額合計(税込)', 
                    '請求書番号', '請求書発行日', 'お支払期日', '入金有無'
                ]]
                
                st.subheader("バクラク請求書データ (整形後)")
                st.dataframe(bakuraku_df_processed.head())
            else:
                st.info("すべてのバクラク請求書が出力対象から除外されました。")
    else:
        st.info("バクラク請求書がアップロードされていないか、データが空です。")

# --- 最終的な結合と出力 ---
st.header("3. 結合結果をファイルに出力")

# 最終的な出力データフレームを結合
final_output_df = pd.DataFrame()
if not np_df_processed.empty and not bakuraku_df_processed.empty:
    final_output_df = pd.concat([np_df_processed, bakuraku_df_processed], ignore_index=True)
elif not np_df_processed.empty:
    final_output_df = np_df_processed
elif not bakuraku_df_processed.empty:
    final_output_df = bakuraku_df_processed

# 合計行の追加
if not final_output_df.empty:
    total_row = pd.DataFrame([
        {
            'ご利用年月': '',
            'ご請求方法': '合計', # 合計行のご請求方法は「合計」とする
            'ご請求金額合計(税込)': final_output_df['ご請求金額合計(税込)'].sum(),
            '未入金金額合計(税込)': final_output_df['未入金金額合計(税込)'].sum(),
            '請求書番号': '',
            '請求書発行日': '',
            'お支払期日': '',
            '入金有無': ''
        }
    ])
    final_output_df = pd.concat([final_output_df, total_row], ignore_index=True)
    
    st.subheader("最終出力データ (合計行含む)")
    st.dataframe(final_output_df)

# 出力オプション
if not final_output_df.empty:
    output_format = st.radio("出力形式を選択してください:", ("Excel (.xlsx)", "CSV (.csv)"))

    timestamp_str = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_filename_base = f"ご請求およびご入金状況一覧_{timestamp_str}"

    if output_format == "Excel (.xlsx)":
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            final_output_df.to_excel(writer, sheet_name="ご入金状況一覧", index=False)
        
        st.download_button(
            label="Excelでダウンロード",
            data=excel_buffer.getvalue(),
            file_name=f"{output_filename_base}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_excel"
        )
    else: # CSV (.csv)
        # CSVエンコーディングの選択を変更
        csv_encoding_choice = st.radio(
            "CSVの文字エンコーディングを選択してください:",
            ("UTF-8 (BOMなし)", "Windows (CP932)"), # 選択肢名をより分かりやすく
            key="csv_encoding_select"
        )
        
        # 選択に基づいてエンコーディングを設定
        if csv_encoding_choice == "UTF-8 (BOMなし)":
            selected_encoding = 'utf-8' # BOMなしのUTF-8
        else: # Windows (CP932)
            selected_encoding = 'cp932' # Windows環境の日本語Shift-JIS

        st.download_button(
            label="CSVでダウンロード",
            data=final_output_df.to_csv(index=False, encoding=selected_encoding),
            file_name=f"{output_filename_base}.csv",
            mime="text/csv",
            key="download_csv"
        )
else:
    st.info("出力対象のデータがありません。")

st.info("注: 画像内のヘッダー・フッターなどの固定テキストは、データ部分の出力には含まれません。")

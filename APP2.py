import streamlit as st
import pandas as pd
from datetime import datetime
import io # Excel/CSV出力用
import zipfile # ZIP圧縮用

# --- Streamlit UI ---
st.title("請求書管理アプリ")
st.write("NP掛け払いとバクラク請求書のCSVを処理し、Excel/CSVファイルに出力します。")

# --- NP掛け払いCSVアップロード ---
st.header("1. NP掛け払いCSVのアップロード")
uploaded_np_file = st.file_uploader("NP掛け払いCSVファイルをアップロードしてください", type=["csv"], key="np_uploader")

np_df = pd.DataFrame()
if uploaded_np_file is not None:
    # 実際のNP掛け払いCSVのカラム名とエンコーディングに合わせてください
    # 例: encoding='shift_jis', 'utf-8', 'cp932'など
    try:
        np_df = pd.read_csv(uploaded_np_file, encoding='shift_jis')
    except UnicodeDecodeError:
        try:
            np_df = pd.read_csv(uploaded_np_file, encoding='utf-8')
        except Exception as e:
            st.error(f"NP掛け払いCSVの読み込み中にエラーが発生しました: {e}。エンコーディングを確認してください。")
            np_df = pd.DataFrame() # エラー時は空のDataFrameを設定
            
    if not np_df.empty:
        st.subheader("NP掛け払いCSV内容（先頭5行）")
        st.dataframe(np_df.head())

        # NP掛け払い処理（入金ステータスに基づく判別）
        # NP掛け払いCSVの実際のカラム名に合わせる
        if '入金ステータス' in np_df.columns: # 仮定: NP掛け払いCSVに「入金ステータス」カラムがある
            np_df['入金状況'] = np_df['入金ステータス'].apply(lambda x: '入金済み' if x == '入金完了' else '未入金')
            st.subheader("NP掛け払い入金状況")
            # NP掛け払いCSVのカラム名をここに合わせる
            # 例: np_df[['請求番号', '顧客名', '請求金額', '入金ステータス', '入金状況']]
            try:
                st.dataframe(np_df[['請求番号', '顧客名', '請求金額', '入金ステータス', '入金状況']].head())
            except KeyError as e:
                st.warning(f"NP掛け払いCSVの表示カラムが見つかりません。カラム名'{e}'を確認してください。")
        else:
            st.warning("NP掛け払いCSVに'入金ステータス'カラムが見つかりませんでした。NP掛け払いCSVのカラム名を確認してください。")

# --- バクラク請求書CSVアップロード ---
st.header("2. バクラク請求書CSVのアップロード")
uploaded_bakuraku_file = st.file_uploader("バクラク請求書CSVファイルをアップロードしてください", type=["csv"], key="bakuraku_uploader")

bakuraku_df = pd.DataFrame()
bakuraku_output_unpaid_df = pd.DataFrame() # 出力するバクラク未入金データ
bakuraku_paid_df = pd.DataFrame() # バクラク入金済みデータ

if uploaded_bakuraku_file is not None:
    # バクラクCSVは通常UTF-8-sigが多いですが、念のため
    try:
        bakuraku_df = pd.read_csv(uploaded_bakuraku_file, encoding='utf-8-sig')
    except UnicodeDecodeError:
        try:
            bakuraku_df = pd.read_csv(uploaded_bakuraku_file, encoding='utf-8')
        except Exception as e:
            st.error(f"バクラク請求書CSVの読み込み中にエラーが発生しました: {e}。エンコーディングを確認してください。")
            bakuraku_df = pd.DataFrame() # エラー時は空のDataFrameを設定

    if not bakuraku_df.empty:
        st.subheader("バクラク請求書CSV内容（先頭5行）")
        st.dataframe(bakuraku_df.head())

        # 念のため、必要なカラムが存在するかチェック
        required_bakuraku_cols = ['書類番号', '送付先名', '金額']
        if not all(col in bakuraku_df.columns for col in required_bakuraku_cols):
            st.error(f"バクラクCSVに必要なカラム({', '.join(required_bakuraku_cols)})が見つかりません。CSVファイルを確認してください。")
            bakuraku_df = pd.DataFrame() # 必要なカラムがなければ処理を中断
        else:
            bakuraku_df['請求識別子'] = bakuraku_df['書類番号'].astype(str) + " - " + bakuraku_df['送付先名'].astype(str) + " - " + bakuraku_df['金額'].astype(str)

            # Step 2a: 入金済み請求の選択
            st.subheader("2a. 入金済みと認識するバクラク請求書を選択してください")
            selected_paid_invoices = st.multiselect(
                "入金済みの請求を選択",
                options=bakuraku_df['請求識別子'].tolist(),
                help="入金が完了している請求を選択してください。",
                key="bakuraku_paid_select"
            )

            # 選択された請求を「入金済み」としてマーク
            bakuraku_paid_df = bakuraku_df[bakuraku_df['請求識別子'].isin(selected_paid_invoices)].copy()
            if not bakuraku_paid_df.empty:
                bakuraku_paid_df['入金状況'] = '入金済み（ユーザー選択）'
                st.subheader("選択された入金済みのバクラク請求書")
                st.dataframe(bakuraku_paid_df[['書類番号', '送付先名', '金額', '入金状況']].head())

            # 選択されなかった請求を「未入金」として仮マーク
            bakuraku_temp_unpaid_df = bakuraku_df[~bakuraku_df['請求識別子'].isin(selected_paid_invoices)].copy()
            if not bakuraku_temp_unpaid_df.empty:
                st.subheader("2b. 出力から除外する未入金請求書を選択してください")
                excluded_unpaid_from_output = st.multiselect(
                    "出力から除外する未入金請求書を選択",
                    options=bakuraku_temp_unpaid_df['請求識別子'].tolist(),
                    help="Excel/CSVファイルに未入金として出力したくない請求書を選択してください。",
                    key="bakuraku_exclude_select"
                )
                bakuraku_output_unpaid_df = bakuraku_temp_unpaid_df[~bakuraku_temp_unpaid_df['請求識別子'].isin(excluded_unpaid_from_output)].copy()
                
                if not bakuraku_output_unpaid_df.empty:
                    bakuraku_output_unpaid_df['入金状況'] = '未入金（出力対象）'
                    st.subheader("出力対象の未入金バクラク請求書")
                    st.dataframe(bakuraku_output_unpaid_df[['書類番号', '送付先名', '金額', '入金状況']].head())
                else:
                    st.info("すべての未入金バクラク請求書が出力対象から除外されました。")
            else:
                st.info("すべてのバクラク請求書が入金済みとして選択されました。出力する未入金請求書はありません。")
    else:
        st.info("バクラク請求書がアップロードされていないか、データが空です。")

# --- ファイルへの出力 ---
st.header("3. 結果をファイルに出力")

# 出力可能なデータフレームをリストアップ
output_dataframes = {}
np_output_df_final = pd.DataFrame()

if not np_df.empty:
    # ここもNP掛け払いCSVのカラム名に合わせて修正する
    # 例: ['請求番号', '顧客名', '請求金額', '入金ステータス', '入金状況']
    try:
        np_output_df_final = np_df[['請求番号', '顧客名', '請求金額', '入金ステータス', '入金状況']]
        output_dataframes["NP掛け払い"] = np_output_df_final
    except KeyError as e:
        st.warning(f"NP掛け払いデータ出力用のカラム'{e}'が見つかりません。CSVを確認してください。")

if not bakuraku_output_unpaid_df.empty:
    output_dataframes["バクラク未入金"] = bakuraku_output_unpaid_df[['書類番号', '送付先名', '金額', '入金状況']]

if not bakuraku_paid_df.empty:
    output_dataframes["バクラク入金済み"] = bakuraku_paid_df[['書類番号', '送付先名', '金額', '入金状況']]

# 少なくともどちらかのデータが存在する場合にのみ、出力オプションを表示
if output_dataframes:
    output_format = st.radio("出力形式を選択してください:", ("Excel (.xlsx)", "CSV (ZIP)"))

    if output_format == "Excel (.xlsx)":
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            for sheet_name, df in output_dataframes.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # データが何も書き込まれていない可能性があるため、最終的なバイト数を確認してからダウンロードボタンを表示
        if excel_buffer.getbuffer().nbytes > 0:
            st.download_button(
                label="結果をExcelでダウンロード",
                data=excel_buffer.getvalue(),
                file_name=f"請求処理結果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_excel_all"
            )
        else:
            st.warning("出力対象のデータがありません。")

    else: # CSV (ZIP)
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
            for file_name_prefix, df in output_dataframes.items():
                # CSVファイル名をシート名として使用
                csv_file_name = f"{file_name_prefix}.csv"
                # DataFrameをCSV形式でメモリに書き込む
                csv_data = df.to_csv(index=False, encoding='utf-8-sig')
                # ZIPファイルに追加
                zf.writestr(csv_file_name, csv_data)
        
        if zip_buffer.getbuffer().nbytes > 0:
            st.download_button(
                label="結果をCSV (ZIP) でダウンロード",
                data=zip_buffer.getvalue(),
                file_name=f"請求処理結果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip",
                key="download_csv_zip_all"
            )
        else:
            st.warning("出力対象のデータがありません。")
else:
    st.info("データがアップロードされていないか、処理されたデータがありません。")

st.info("添付のExcel形式については、出力するデータフレームのカラム名を調整することで対応できます。")

import streamlit as st
import pandas as pd
from datetime import datetime
import io # Excel/CSV出力用

# --- Streamlit UI ---
st.title("請求書管理アプリ")
st.write("NP掛け払いとバクラク請求書のCSVを処理し、Excel/CSVファイルに出力します。")

# --- NP掛け払いCSVアップロード ---
st.header("1. NP掛け払いCSVのアップロード")
uploaded_np_file = st.file_uploader("NP掛け払いCSVファイルをアップロードしてください", type=["csv"], key="np_uploader")

np_df = pd.DataFrame()
if uploaded_np_file is not None:
    np_df = pd.read_csv(uploaded_np_file, encoding='shift_jis') # 必要に応じてencodingを変更
    st.subheader("NP掛け払いCSV内容（先頭5行）")
    st.dataframe(np_df.head())

    if '入金ステータス' in np_df.columns:
        np_df['入金状況'] = np_df['入金ステータス'].apply(lambda x: '入金済み' if x == '入金完了' else '未入金')
        st.subheader("NP掛け払い入金状況")
        st.dataframe(np_df[['請求番号', '顧客名', '請求金額', '入金ステータス', '入金状況']].head())
    else:
        st.warning("NP掛け払いCSVに'入金ステータス'カラムが見つかりませんでした。")

# --- バクラク請求書CSVアップロード ---
st.header("2. バクラク請求書CSVのアップロード")
uploaded_bakuraku_file = st.file_uploader("バクラク請求書CSVファイルをアップロードしてください", type=["csv"], key="bakuraku_uploader")

bakuraku_df = pd.DataFrame()
bakuraku_output_unpaid_df = pd.DataFrame() # スプレッドシートに出力するバクラク未入金
bakuraku_paid_df = pd.DataFrame() # バクラク入金済み

if uploaded_bakuraku_file is not None:
    bakuraku_df = pd.read_csv(uploaded_bakuraku_file, encoding='utf-8') # 必要に応じてencodingを変更
    st.subheader("バクラク請求書CSV内容（先頭5行）")
    st.dataframe(bakuraku_df.head())

    if not bakuraku_df.empty:
        bakuraku_df['請求識別子'] = bakuraku_df['請求番号'].astype(str) + " - " + bakuraku_df['取引先'].astype(str) + " - " + bakuraku_df['請求額'].astype(str)

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
            st.dataframe(bakuraku_paid_df[['請求番号', '取引先', '請求額', '入金状況']].head())

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
                st.dataframe(bakuraku_output_unpaid_df[['請求番号', '取引先', '請求額', '入金状況']].head())
            else:
                st.info("すべての未入金バクラク請求書が出力対象から除外されました。")
        else:
            st.info("すべてのバクラク請求書が入金済みとして選択されました。出力する未入金請求書はありません。")
    else:
        st.info("バクラク請求書がアップロードされていないか、データが空です。")

# --- ファイルへの出力 ---
st.header("3. 結果をファイルに出力")

# 出力形式の選択
output_format = st.radio("出力形式を選択してください:", ("Excel (.xlsx)", "CSV (.csv)"))

# ダウンロードボタンの準備
# NP掛け払いデータ
if not np_df.empty:
    np_output_df_final = np_df[['請求番号', '顧客名', '請求金額', '入金ステータス', '入金状況']] # 出力するカラムを調整
    
    if output_format == "Excel (.xlsx)":
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            np_output_df_final.to_excel(writer, sheet_name="NP掛け払い", index=False)
            if not bakuraku_output_unpaid_df.empty:
                bakuraku_output_unpaid_df[['請求番号', '取引先', '請求額', '入金状況']].to_excel(writer, sheet_name="バクラク未入金", index=False)
            if not bakuraku_paid_df.empty:
                bakuraku_paid_df[['請求番号', '取引先', '請求額', '入金状況']].to_excel(writer, sheet_name="バクラク入金済み", index=False)
        st.download_button(
            label="結果をExcelでダウンロード",
            data=excel_buffer.getvalue(),
            file_name=f"請求処理結果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_excel_all"
        )
    else: # CSV (.csv)
        col1, col2, col3 = st.columns(3)
        if not np_output_df_final.empty:
            with col1:
                st.download_button(
                    label="NP掛け払いCSVダウンロード",
                    data=np_output_df_final.to_csv(index=False, encoding='utf-8-sig'),
                    file_name=f"NP掛け払い_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    key="download_np_csv"
                )
        if not bakuraku_output_unpaid_df.empty:
            with col2:
                st.download_button(
                    label="バクラク未入金CSVダウンロード",
                    data=bakuraku_output_unpaid_df[['請求番号', '取引先', '請求額', '入金状況']].to_csv(index=False, encoding='utf-8-sig'),
                    file_name=f"バクラク未入金_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    key="download_bakuraku_unpaid_csv"
                )
        if not bakuraku_paid_df.empty:
            with col3:
                st.download_button(
                    label="バクラク入金済みCSVダウンロード",
                    data=bakuraku_paid_df[['請求番号', '取引先', '請求額', '入金状況']].to_csv(index=False, encoding='utf-8-sig'),
                    file_name=f"バクラク入金済み_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    key="download_bakuraku_paid_csv"
                )
else:
    st.info("NP掛け払いデータがアップロードされていないため、ファイル出力オプションは表示されません。")

st.info("添付のExcel形式については、出力するデータフレームのカラム名を調整することで対応できます。")

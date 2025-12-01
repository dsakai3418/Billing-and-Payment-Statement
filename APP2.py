import streamlit as st
import pandas as pd
import gspread
from gspread_dataframe import set_with_dataframe
from datetime import datetime

# --- Streamlit UI ---
st.title("請求書管理アプリ")
st.write("NP掛け払いとバクラク請求書のCSVを処理し、スプレッドシートに出力します。")

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

    # 請求を特定するためのユニークなID
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
            st.subheader("2b. スプレッドシートから除外する未入金請求書を選択してください")
            # 出力対象から外す未入金請求書を選択
            excluded_unpaid_from_output = st.multiselect(
                "出力から除外する未入金請求書を選択",
                options=bakuraku_temp_unpaid_df['請求識別子'].tolist(),
                help="スプレッドシートに未入金として出力したくない請求書を選択してください。",
                key="bakuraku_exclude_select"
            )
            # 選択されなかったものが「出力対象の未入金請求書」となる
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

# --- スプレッドシートへの出力 ---
st.header("3. 結果をスプレッドシートに出力")

# ここでGoogle Sheets APIの認証を行います。
try:
    gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
except Exception as e:
    st.error(f"Google Sheets APIの認証に失敗しました。Streamlit Secretsの設定を確認してください。エラー: {e}")
    gc = None

if st.button("スプレッドシートに出力", disabled=(np_df.empty and bakuraku_output_unpaid_df.empty) or gc is None):
    if gc is not None:
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            spreadsheet_name = f"請求処理結果_{timestamp}"
            sh = gc.create(spreadsheet_name) # 新しいスプレッドシートを作成
            st.success(f"新しいスプレッドシート '{spreadsheet_name}' を作成しました。")
            st.write(f"[スプレッドシートへのリンク]({sh.url})")

            # NP掛け払いデータを出力
            if not np_df.empty:
                np_output_df = np_df[['請求番号', '顧客名', '請求金額', '入金ステータス', '入金状況']]
                worksheet_np = sh.add_worksheet(title="NP掛け払い", rows=len(np_output_df)+1, cols=len(np_output_df.columns))
                set_with_dataframe(worksheet_np, np_output_df)
                st.info("NP掛け払いデータをスプレッドシートに出力しました。")

            # バクラク出力対象未入金データを出力
            if not bakuraku_output_unpaid_df.empty:
                bakuraku_unpaid_output_df = bakuraku_output_unpaid_df[['請求番号', '取引先', '請求額', '入金状況']]
                worksheet_bakuraku_unpaid = sh.add_worksheet(title="バクラク未入金", rows=len(bakuraku_unpaid_output_df)+1, cols=len(bakuraku_unpaid_output_df.columns))
                set_with_dataframe(worksheet_bakuraku_unpaid, bakuraku_unpaid_output_df)
                st.info("バクラク未入金データをスプレッドシートに出力しました。")
            else:
                st.info("バクラク未入金データは存在しないため、出力されませんでした。")

            # バクラク入金済みデータを出力（必要であれば、シートを分けるか検討）
            if not bakuraku_paid_df.empty:
                bakuraku_paid_output_df = bakuraku_paid_df[['請求番号', '取引先', '請求額', '入金状況']]
                worksheet_bakuraku_paid = sh.add_worksheet(title="バクラク入金済み", rows=len(bakuraku_paid_output_df)+1, cols=len(bakuraku_paid_output_df.columns))
                set_with_dataframe(worksheet_bakuraku_paid, bakuraku_paid_output_df)
                st.info("バクラク入金済みデータもスプレッドシートに出力しました。")
            else:
                st.info("バクラク入金済みデータは存在しないため、出力されませんでした。")


            st.balloons()
        except Exception as e:
            st.error(f"スプレッドシートへの出力中にエラーが発生しました: {e}")
    else:
        st.warning("Google Sheets APIが認証されていません。")

st.info("添付のExcel形式については、出力するデータフレームのカラム名を調整することで対応できます。")

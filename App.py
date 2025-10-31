import streamlit as st
import pandas as pd
import io
import datetime
# from weasyprint import HTML, CSS # PDF生成用 - 削除

st.set_page_config(layout="wide") # レイアウトを広めに設定

st.title("請求・入金状況確認アプリ")

# --- NP掛け払いCSVのアップロード ---
st.header("NP掛け払いCSVアップロード")
uploaded_file_np = st.file_uploader("NP掛け払いCSVファイルをアップロードしてください", type="csv", key="np_uploader")

df_np = None
if uploaded_file_np is not None:
    df_np = pd.read_csv(uploaded_file_np)
    st.subheader("NP掛け払いデータプレビュー")
    st.dataframe(df_np.head())

    # NP掛け払いデータの処理
    df_np['請求書発行日'] = pd.to_datetime(df_np['請求書発行日'])
    df_np['支払期限日'] = pd.to_datetime(df_np['支払期限日'])
    df_np['入金有無'] = df_np['入金ステータス'].apply(lambda x: 'あり' if x == '入金済み' else 'なし')
    df_np['ご請求方法'] = 'NP掛け払い'
    df_np['未入金金額合計 (税込)'] = df_np.apply(lambda row: row['請求金額'] if row['入金有無'] == 'なし' else 0, axis=1)
    df_np_processed = df_np[['請求書発行日', '支払期限日', '請求番号', '企業名', 'ご請求方法', '請求金額', '未入金金額合計 (税込)', '入金有無']]
    df_np_processed = df_np_processed.rename(columns={'請求金額': 'ご請求金額合計 (税込)'})
    
    st.subheader("NP掛け払い処理結果")
    st.dataframe(df_np_processed)


# --- バクラク請求書CSVのアップロード ---
st.header("バクラク請求書CSVアップロード")
uploaded_file_bakuraku = st.file_uploader("バクラク請求書CSVファイルをアップロードしてください", type="csv", key="bakuraku_uploader")

df_bakuraku = None
if uploaded_file_bakuraku is not None:
    df_bakuraku = pd.read_csv(uploaded_file_bakuraku)
    st.subheader("バクラク請求書データプレビュー")
    st.dataframe(df_bakuraku.head())

    # バクラク請求書データの処理
    df_bakuraku['日付'] = pd.to_datetime(df_bakuraku['日付'])
    df_bakuraku['支払期日'] = pd.to_datetime(df_bakuraku['支払期日'])
    df_bakuraku['ご請求方法'] = '直接請求' # 仮に設定
    df_bakuraku['ご請求金額合計 (税込)'] = df_bakuraku['金額']

    st.subheader("バクラク請求書 未入金状況選択")
    st.write("未入金の請求書にチェックを入れてください。")
    
    selected_unpaid_bakuraku = {}
    if df_bakuraku is not None:
        with st.expander("バクラク請求書一覧を開く"):
            for index, row in df_bakuraku.iterrows():
                unique_key = f"bakuraku_unpaid_{row['書類番号']}_{row['日付']}_{row['金額']}"
                if st.checkbox(f"書類番号: {row['書類番号']}, 日付: {row['日付'].strftime('%Y-%m-%d')}, 金額: {row['金額']:,}円", key=unique_key):
                    selected_unpaid_bakuraku[index] = True
                else:
                    selected_unpaid_bakuraku[index] = False
    
    df_bakuraku['入金有無'] = df_bakuraku.index.map(lambda idx: 'なし' if selected_unpaid_bakuraku.get(idx, False) else 'あり')
    df_bakuraku['未入金金額合計 (税込)'] = df_bakuraku.apply(lambda row: row['金額'] if row['入金有無'] == 'なし' else 0, axis=1)
    
    df_bakuraku_processed = df_bakuraku[['日付', '支払期日', '書類番号', '送付先名', 'ご請求方法', 'ご請求金額合計 (税込)', '未入金金額合計 (税込)', '入金有無']]
    df_bakuraku_processed = df_bakuraku_processed.rename(columns={'日付': '請求書発行日', '支払期日': 'お支払期日', '書類番号': '請求書番号', '送付先名': '企業名'})

    st.subheader("バクラク請求書処理結果")
    st.dataframe(df_bakuraku_processed)

# --- 統合結果の表示とExcel出力 ---
if df_np is not None and df_bakuraku is not None:
    st.header("統合された請求および入金状況")
    
    # 統合データフレームの作成
    df_np_processed['ご利用年月'] = df_np_processed['請求書発行日'].dt.strftime('%Y年%m月')
    df_bakuraku_processed['ご利用年月'] = df_bakuraku_processed['請求書発行日'].dt.strftime('%Y年%m月')

    combined_df = pd.concat([
        df_np_processed[['ご利用年月', 'ご請求方法', 'ご請求金額合計 (税込)', '未入金金額合計 (税込)', '請求書番号', '請求書発行日', 'お支払期日', '入金有無']],
        df_bakuraku_processed[['ご利用年月', 'ご請求方法', 'ご請求金額合計 (税込)', '未入金金額合計 (税込)', '請求書番号', '請求書発行日', 'お支払期日', '入金有無']]
    ])
    
    combined_df = combined_df.sort_values(by=['ご利用年月', '請求書発行日']).reset_index(drop=True)

    total_請求金額 = combined_df['ご請求金額合計 (税込)'].sum()
    total_未入金金額 = combined_df['未入金金額合計 (税込)'].sum()
    
    total_row = pd.DataFrame([{
        'ご利用年月': '',
        'ご請求方法': '合計',
        'ご請求金額合計 (税込)': total_請求金額,
        '未入金金額合計 (税込)': total_未入金金額,
        '請求書番号': '',
        '請求書発行日': '',
        'お支払期日': '',
        '入金有無': ''
    }])
    
    combined_df_with_total = pd.concat([combined_df, total_row], ignore_index=True)

    # Streamlitでの表示 (PDFのような形式を意識)
    # Excel出力のみなので、HTML表示はシンプルにDataFrameをst.dataframeで表示するか、Markdownで体裁を整える
    
    st.markdown("### 株式会社BHUSAL ENTERPRISESさま")
    st.markdown("### ご請求およびご入金状況一覧")
    st.markdown(f"**作成日: {datetime.date.today().strftime('%Y/%m/%d')}**")

    # DataFrameをそのまま表示 (PDFのようなフォーマットはExcelで実現)
    st.dataframe(combined_df_with_total.style.format({
        'ご請求金額合計 (税込)': '{:,.0f}',
        '未入金金額合計 (税込)': '{:,.0f}'
    })) # 金額にカンマ区切りフォーマットを適用

    st.markdown(f"**※{datetime.date.today().strftime('%Y年%m月')}時点での未入金合計金額: {total_未入金金額:,}円**")


    # --- Excel出力ボタン ---
    current_date_str = datetime.date.today().strftime('%Y%m%d_%H%M%S')

    excel_buffer = io.BytesIO()
    
    # 日付列をExcel friendlyな形式に変換
    output_df = combined_df.copy()
    output_df['請求書発行日'] = output_df['請求書発行日'].dt.strftime('%Y/%m/%d')
    output_df['お支払期日'] = output_df['お支払期日'].dt.strftime('%Y/%m/%d')
    
    # 合計行を追加したDataFrameをExcelに書き込む
    output_df_with_total = pd.concat([output_df, total_row.astype(output_df.dtypes)], ignore_index=True)

    # Excelに書き出し
    output_df_with_total.to_excel(excel_buffer, index=False, engine='openpyxl')
    excel_buffer.seek(0) # バッファの先頭に戻す

    st.download_button(
        label="Excelファイルとしてダウンロード",
        data=excel_buffer,
        file_name=f"請求入金状況_{current_date_str}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.info("Excelファイルとしてダウンロード可能です。")

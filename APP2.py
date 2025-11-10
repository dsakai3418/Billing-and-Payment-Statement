import streamlit as st
import pandas as pd
import io
import datetime
import unicodedata # 文字列の正規化に使用
import chardet # エンコーディング検出用に追加

# Streamlitページの基本設定
st.set_page_config(
    page_title="請求・入金状況確認アプリ",
    layout="wide", # レイアウトを広めに設定
    initial_sidebar_state="expanded" # サイドバーをデフォルトで開く
)

st.title("請求・入金状況確認アプリ")

# --- ファイルのエンコーディングを検出して読み込む関数 ---
@st.cache_data
def load_csv_with_encoding_detection(uploaded_file):
    raw_data = uploaded_file.read()
    result = chardet.detect(raw_data)
    encoding = result['encoding']
    
    # 検出されたエンコーディングで読み込みを試みる
    try:
        df = pd.read_csv(io.BytesIO(raw_data), encoding=encoding)
        return df
    except Exception as e:
        # 検出されたエンコーディングで失敗した場合、一般的なエンコーディングを試す
        common_encodings = ['utf-8', 'shift_jis', 'cp932', 'euc-jp']
        for enc in common_encodings:
            try:
                df = pd.read_csv(io.BytesIO(raw_data), encoding=enc)
                return df
            except:
                continue
        st.error(f"ファイル '{uploaded_file.name}' の読み込み中にエラーが発生しました: {e}。様々なエンコーディングを試しましたが失敗しました。")
        return None

# --- NP掛け払いCSVのアップロード ---
st.header("1. NP掛け払いCSVのアップロード")
uploaded_files_np = st.file_uploader(
    "NP掛け払いCSVファイルを複数選択してアップロードしてください。",
    type="csv",
    accept_multiple_files=True,
    key="np_uploader"
)

df_np_processed = None # 初期化

if uploaded_files_np:
    all_df_np_raw = []
    for uploaded_file in uploaded_files_np:
        df_temp = load_csv_with_encoding_detection(uploaded_file)
        if df_temp is not None:
            all_df_np_raw.append(df_temp)
    
    if all_df_np_raw:
        df_np_raw_combined = pd.concat(all_df_np_raw, ignore_index=True)
        st.subheader("NP掛け払いデータプレビュー (結合後)")
        st.dataframe(df_np_raw_combined.head())

        df_np = df_np_raw_combined.copy()
        
        new_columns = []
        for col in df_np.columns:
            normalized_col = col.strip()
            normalized_col = unicodedata.normalize('NFKC', normalized_col)
            new_columns.append(normalized_col)
        df_np.columns = new_columns
        
        if '企業名' not in df_np.columns:
            st.warning("NP掛け払いCSVに '企業名' 列が見つかりませんでした。空文字列として処理を続行します。")
            df_np['企業名'] = ''

        if '請求番号' not in df_np.columns:
            st.warning("NP掛け払いCSVに '請求番号' 列が見つかりませんでした。空文字列として列を追加します。")
            df_np['請求番号'] = ''
        
        required_np_columns_for_processing = ['請求書発行日', '支払期限日', '請求金額', '入金ステータス'] 
        missing_np_cols = [col for col in required_np_columns_for_processing if col not in df_np.columns]
        
        if missing_np_cols:
            st.error(f"NP掛け払いCSVに以下の必須列が見つかりません: {', '.join(missing_np_cols)}")
            df_np_processed = None
        else:
            df_np['請求書発行日'] = pd.to_datetime(df_np['請求書発行日'], errors='coerce')
            df_np['支払期限日'] = pd.to_datetime(df_np['支払期限日'], errors='coerce') 

            if df_np['請求書発行日'].isnull().any() or df_np['支払期限日'].isnull().any():
                st.warning("NP掛け払いCSVの日付列に無効な値がありました。該当行はNaNとして処理されます。")

            df_np['入金有無'] = df_np['入金ステータス'].apply(lambda x: 'あり' if x == '入金済み' else 'なし')
            df_np['ご請求方法'] = 'NP掛け払い'
            df_np['請求金額'] = pd.to_numeric(df_np['請求金額'], errors='coerce').fillna(0)
            df_np['未入金金額合計 (税込)'] = df_np.apply(lambda row: row['請求金額'] if row['入金有無'] == 'なし' else 0, axis=1)
            
            cols_for_np_processed = ['請求書発行日', '支払期限日', '請求番号', '企業名', 'ご請求方法', '請求金額', '未入金金額合計 (税込)', '入金有無']
            
            if all(col in df_np.columns for col in cols_for_np_processed):
                df_np_processed = df_np[cols_for_np_processed].copy()
                df_np_processed = df_np_processed.rename(columns={
                    '請求金額': 'ご請求金額合計 (税込)', 
                    '支払期限日': 'お支払期日',
                    '請求番号': '請求書番号'
                })
                
                st.subheader("NP掛け払い処理結果")
                st.dataframe(df_np_processed)
            else:
                st.error("NP掛け払い処理済みデータフレームの作成に必要な列が不足しています。予期せぬエラーが発生しました。")
                df_np_processed = None 
    else:
        st.info("NP掛け払いCSVファイルがアップロードされていません。")


# --- バクラク請求書CSVのアップロード ---
st.header("2. バクラク請求書CSVのアップロード")
uploaded_files_bakuraku = st.file_uploader(
    "バクラク請求書CSVファイルを複数選択してアップロードしてください。",
    type="csv",
    accept_multiple_files=True,
    key="bakuraku_uploader"
)

df_bakuraku_processed = None # 初期化

if uploaded_files_bakuraku:
    all_df_bakuraku_raw = []
    for uploaded_file in uploaded_files_bakuraku:
        df_temp = load_csv_with_encoding_detection(uploaded_file)
        if df_temp is not None:
            all_df_bakuraku_raw.append(df_temp)

    if all_df_bakuraku_raw:
        df_bakuraku_raw_combined = pd.concat(all_df_bakuraku_raw, ignore_index=True)
        st.subheader("バクラク請求書データプレビュー (結合後)")
        st.dataframe(df_bakuraku_raw_combined.head())

        df_bakuraku = df_bakuraku_raw_combined.copy()

        new_columns_bakuraku = []
        for col in df_bakuraku.columns:
            normalized_col = col.strip()
            normalized_col = unicodedata.normalize('NFKC', normalized_col)
            new_columns_bakuraku.append(normalized_col)
        df_bakuraku.columns = new_columns_bakuraku
        
        # '送付先名' 列の存在チェックと代替
        if '送付先名' not in df_bakuraku.columns:
            st.warning("バクラク請求書CSVに '送付先名' 列が見つかりませんでした。空文字列として処理を続行します。")
            df_bakuraku['送付先名'] = '' # 存在しない場合は空の列を追加

        required_bakuraku_columns = ['日付', '支払期日', '書類種別', '書類番号', '送付先名', '金額']
        missing_bakuraku_cols = [col for col in required_bakuraku_columns if col not in df_bakuraku.columns]
        if missing_bakuraku_cols:
            st.error(f"バクラク請求書CSVに以下の必須列が見つかりません: {', '.join(missing_bakuraku_cols)}")
            df_bakuraku_processed = None 
        else:
            df_bakuraku['日付'] = pd.to_datetime(df_bakuraku['日付'], errors='coerce')
            df_bakuraku['支払期日'] = pd.to_datetime(df_bakuraku['支払期日'], errors='coerce')

            if df_bakuraku['日付'].isnull().any() or df_bakuraku['支払期日'].isnull().any():
                st.warning("バクラク請求書CSVの日付列に無効な値がありました。該当行はNaNとして処理されます。")

            df_bakuraku['ご請求方法'] = df_bakuraku['書類種別'].fillna('不明')
            
            df_bakuraku['金額'] = pd.to_numeric(df_bakuraku['金額'], errors='coerce').fillna(0)
            df_bakuraku['ご請求金額合計 (税込)'] = df_bakuraku['金額']

            st.subheader("バクラク請求書 未入金状況選択")
            st.write("未入金の請求書にチェックを入れてください。")
            
            selected_unpaid_bakuraku = {}
            if df_bakuraku is not None and not df_bakuraku.empty: 
                with st.expander("バクラク請求書一覧を開く"):
                    df_bakuraku_temp = df_bakuraku.reset_index().rename(columns={'index': 'original_index'})
                    # 表示する際に、同じ「書類番号,日付,金額」の行を重複させないようにする
                    df_bakuraku_display = df_bakuraku_temp[['original_index', '書類番号', '日付', '金額']].drop_duplicates(subset=['書類番号', '日付', '金額']).set_index('original_index')

                    # 既存の選択状態をロード
                    if 'bakuraku_selection' in st.session_state:
                        selected_unpaid_bakuraku = st.session_state.bakuraku_selection
                        
                    # チェックボックスの表示と状態更新
                    for original_idx, row in df_bakuraku_display.iterrows():
                        date_str = row['日付'].strftime('%Y-%m-%d') if pd.notna(row['日付']) else '日付不明'
                        amount_display = f"{row['金額']:,}円" if pd.notna(row['金額']) else '金額不明'
                        
                        # original_idxに対応するチェック状態を取得。存在しなければFalse
                        is_checked = selected_unpaid_bakuraku.get(original_idx, False)

                        unique_key = f"bakuraku_unpaid_{row['書類番号']}_{date_str}_{row['金額']}_{original_idx}"
                        
                        if st.checkbox(f"書類番号: {row['書類番号']}, 日付: {date_str}, 金額: {amount_display}", key=unique_key, value=is_checked):
                            selected_unpaid_bakuraku[original_idx] = True
                        else:
                            selected_unpaid_bakuraku[original_idx] = False
                    
                    # セッションステートに選択状態を保存
                    st.session_state.bakuraku_selection = selected_unpaid_bakuraku


            if df_bakuraku is not None:
                # selected_unpaid_bakurakuはoriginal_indexをキーに持っているので、元のdf_bakurakuのindexを使う
                df_bakuraku['入金有無'] = df_bakuraku.index.map(lambda idx: 'なし' if selected_unpaid_bakuraku.get(idx, False) else 'あり')
                df_bakuraku['未入金金額合計 (税込)'] = df_bakuraku.apply(lambda row: row['金額'] if row['入金有無'] == 'なし' else 0, axis=1)
                
                df_bakuraku_processed = df_bakuraku.groupby(['書類番号', '日付', '支払期日', '送付先名', 'ご請求方法']).agg(
                    金額合計=('ご請求金額合計 (税込)', 'sum'),
                    未入金合計=('未入金金額合計 (税込)', 'sum'),
                    入金有無=('入金有無', lambda x: 'なし' if 'なし' in x.values else 'あり')
                ).reset_index()
                
                df_bakuraku_processed = df_bakuraku_processed.rename(columns={
                    '日付': '請求書発行日', 
                    '支払期日': 'お支払期日', 
                    '書類番号': '請求書番号', 
                    '送付先名': '企業名',
                    '金額合計': 'ご請求金額合計 (税込)', 
                    '未入金合計': '未入金金額合計 (税込)'
                })
            
            st.subheader("バクラク請求書処理結果")
            if df_bakuraku_processed is not None:
                st.dataframe(df_bakuraku_processed)
            else:
                st.write("バクラク請求書データが処理されていません。")
    else:
        st.info("バクラク請求書CSVファイルがアップロードされていません。")


# --- 統合結果の表示とExcel出力 ---
st.header("3. 統合された請求および入金状況")

combined_df_with_total = None
common_cols = ['ご利用年月', 'ご請求方法', 'ご請求金額合計 (税込)', '未入金金額合計 (税込)', '請求書番号', '請求書発行日', 'お支払期日', '入金有無', '企業名']

# 両方のDFが存在する場合
if df_np_processed is not None and df_bakuraku_processed is not None:
    df_np_processed['ご利用年月'] = df_np_processed['請求書発行日'].apply(lambda x: x.strftime('%Y年%m月') if pd.notna(x) else '')
    df_bakuraku_processed['ご利用年月'] = df_bakuraku_processed['請求書発行日'].apply(lambda x: x.strftime('%Y年%m月') if pd.notna(x) else '')

    missing_cols_np = [col for col in common_cols if col not in df_np_processed.columns]
    missing_cols_bakuraku = [col for col in common_cols if col not in df_bakuraku_processed.columns]

    if missing_cols_np:
        st.error(f"NP掛け払い処理結果データに結合に必要な列が不足しています: {', '.join(missing_cols_np)}")
    elif missing_cols_bakuraku:
        st.error(f"バクラク請求書処理結果データに結合に必要な列が不足しています: {', '.join(missing_cols_bakuraku)}")
    else:
        combined_df = pd.concat([
            df_np_processed[common_cols],
            df_bakuraku_processed[common_cols]
        ])
        
        combined_df['請求書発行日'] = pd.to_datetime(combined_df['請求書発行日'], errors='coerce')
        combined_df['お支払期日'] = pd.to_datetime(combined_df['お支払期日'], errors='coerce')

        combined_df = combined_df.sort_values(by=['ご利用年月', '請求書発行日'], na_position='last').reset_index(drop=True)

        total_請求金額 = combined_df['ご請求金額合計 (税込)'].sum()
        total_未入金金額 = combined_df['未入金金額合計 (税込)'].sum()
        
        total_row = pd.DataFrame([{
            'ご利用年月': '',
            'ご請求方法': '合計',
            'ご請求金額合計 (税込)': total_請求金額,
            '未入金金額合計 (税込)': total_未入金金額,
            '請求書番号': '',
            '請求書発行日': pd.NaT, 
            'お支払期日': pd.NaT, 
            '入金有無': '',
            '企業名': '' 
        }])
        
        for col in ['ご請求金額合計 (税込)', '未入金金額合計 (税込)']:
            if col in combined_df.columns and total_row[col].dtype != combined_df[col].dtype:
                total_row[col] = total_row[col].astype(combined_df[col].dtype)
        
        combined_df_with_total = pd.concat([combined_df, total_row], ignore_index=True)

# NP掛け払いデータのみ存在する場合
elif df_np_processed is not None:
    df_np_processed['ご利用年月'] = df_np_processed['請求書発行日'].apply(lambda x: x.strftime('%Y年%m月') if pd.notna(x) else '')
    missing_cols_np = [col for col in common_cols if col not in df_np_processed.columns]
    if missing_cols_np:
        st.error(f"NP掛け払い処理結果データに表示に必要な列が不足しています: {', '.join(missing_cols_np)}")
    else:
        combined_df = df_np_processed[common_cols].copy()
        combined_df['請求書発行日'] = pd.to_datetime(combined_df['請求書発行日'], errors='coerce')
        combined_df['お支払期日'] = pd.to_datetime(combined_df['お支払期日'], errors='coerce')
        combined_df = combined_df.sort_values(by=['ご利用年月', '請求書発行日'], na_position='last').reset_index(drop=True)

        total_請求金額 = combined_df['ご請求金額合計 (税込)'].sum()
        total_未入金金額 = combined_df['未入金金額合計 (税込)'].sum()
        total_row = pd.DataFrame([{
            'ご利用年月': '', 'ご請求方法': '合計', 'ご請求金額合計 (税込)': total_請求金額,
            '未入金金額合計 (税込)': total_未入金金額, '請求書番号': '', '請求書発行日': pd.NaT,
            'お支払期日': pd.NaT, '入金有無': '', '企業名': ''
        }])
        for col in ['ご請求金額合計 (税込)', '未入金金額合計 (税込)']:
            if col in combined_df.columns and total_row[col].dtype != combined_df[col].dtype:
                total_row[col] = total_row[col].astype(combined_df[col].dtype)
        combined_df_with_total = pd.concat([combined_df, total_row], ignore_index=True)

# バクラク請求書データのみ存在する場合
elif df_bakuraku_processed is not None:
    df_bakuraku_processed['ご利用年月'] = df_bakuraku_processed['請求書発行日'].apply(lambda x: x.strftime('%Y年%m月') if pd.notna(x) else '')
    missing_cols_bakuraku = [col for col in common_cols if col not in df_bakuraku_processed.columns]
    if missing_cols_bakuraku:
        st.error(f"バクラク請求書処理結果データに表示に必要な列が不足しています: {', '.join(missing_cols_bakuraku)}")
    else:
        combined_df = df_bakuraku_processed[common_cols].copy()
        combined_df['請求書発行日'] = pd.to_datetime(combined_df['請求書発行日'], errors='coerce')
        combined_df['お支払期日'] = pd.to_datetime(combined_df['お支払期日'], errors='coerce')
        combined_df = combined_df.sort_values(by=['ご利用年月', '請求書発行日'], na_position='last').reset_index(drop=True)

        total_請求金額 = combined_df['ご請求金額合計 (税込)'].sum()
        total_未入金金額 = combined_df['未入金金額合計 (税込)'].sum()
        total_row = pd.DataFrame([{
            'ご利用年月': '', 'ご請求方法': '合計', 'ご請求金額合計 (税込)': total_請求金額,
            '未入金金額合計 (税込)': total_未入金金額, '請求書番号': '', '請求書発行日': pd.NaT,
            'お支払期日': pd.NaT, '入金有無': '', '企業名': ''
        }])
        for col in ['ご請求金額合計 (税込)', '未入金金額合計 (税込)']:
            if col in combined_df.columns and total_row[col].dtype != combined_df[col].dtype:
                total_row[col] = total_row[col].astype(combined_df[col].dtype)
        combined_df_with_total = pd.concat([combined_df, total_row], ignore_index=True)

# どちらかのデータが存在すれば、以下の表示ロジックを実行
if combined_df_with_total is not None:
    company_name = "取引先" 
    all_unique_companies = []

    if df_np_processed is not None and not df_np_processed.empty and '企業名' in df_np_processed.columns:
        all_unique_companies.extend(df_np_processed['企業名'].dropna().unique().tolist())
        
    if df_bakuraku_processed is not None and not df_bakuraku_processed.empty and '企業名' in df_bakuraku_processed.columns:
        all_unique_companies.extend(df_bakuraku_processed['企業名'].dropna().unique().tolist())
    
    unique_companies_set = sorted(list(set(all_unique_companies)))

    if len(unique_companies_set) == 1:
        company_name = unique_companies_set[0]
    elif len(unique_companies_set) > 1:
        company_name = ", ".join(unique_companies_set[:2]) + " 他"
    
    if not company_name or company_name.strip() == "": 
        company_name = "取引先"


    st.markdown(f"### {company_name}さま")
    st.markdown("### ご請求およびご入金状況一覧")
    st.markdown(f"**作成日: {datetime.date.today().strftime('%Y/%m/%d')}**")

    st.dataframe(combined_df_with_total.style.format({
        'ご請求金額合計 (税込)': '{:,.0f}',
        '未入金金額合計 (税込)': '{:,.0f}',
        '請求書発行日': lambda x: x.strftime('%Y/%m/%d') if pd.notna(x) else '',
        'お支払期日': lambda x: x.strftime('%Y/%m/%d') if pd.notna(x) else ''
    }))

    total_未入金金額 = combined_df_with_total['未入金金額合計 (税込)'].iloc[-1] if not combined_df_with_total.empty else 0
    st.markdown(f"**※{datetime.date.today().strftime('%Y年%m月')}時点での未入金合計金額: {total_未入金金額:,}円**")


    # --- Excel出力ボタン ---
    current_date_str = datetime.date.today().strftime('%Y%m%d_%H%M%S')

    excel_buffer = io.BytesIO()
    
    output_df = combined_df_with_total.copy() 
    output_df['請求書発行日'] = output_df['請求書発行日'].dt.strftime('%Y/%m/%d').fillna('')
    output_df['お支払期日'] = output_df['お支払期日'].dt.strftime('%Y/%m/%d').fillna('')
    
    output_df.to_excel(excel_buffer, index=False, sheet_name='請求入金状況', engine='openpyxl')
    excel_buffer.seek(0) 

    st.download_button(
        label="Excelファイルとしてダウンロード",
        data=excel_buffer,
        file_name=f"請求入金状況_{current_date_str}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.info("Excelファイルとしてダウンロード可能です。")
else: 
    st.info("NP掛け払いCSVファイル、またはバクラク請求書CSVファイルをアップロードしてください。")

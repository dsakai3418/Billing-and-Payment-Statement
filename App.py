import streamlit as st
import pandas as pd
import io
import datetime
import unicodedata # 文字列の正規化に使用

# Streamlitページの基本設定
st.set_page_config(
    page_title="請求・入金状況確認アプリ",
    layout="wide", # レイアウトを広めに設定
    initial_sidebar_state="expanded" # サイドバーをデフォルトで開く
)

st.title("請求・入金状況確認アプリ")

# --- NP掛け払いCSVのアップロード ---
st.header("1. NP掛け払いCSVのアップロード")
# 複数ファイルのアップロードを許可
uploaded_files_np = st.file_uploader(
    "NP掛け払いCSVファイルを複数選択してアップロードしてください。",
    type="csv",
    accept_multiple_files=True,
    key="np_uploader"
)

df_np_processed = None # 初期化

if uploaded_files_np: # ファイルがアップロードされた場合のみ処理
    all_df_np_raw = []
    for uploaded_file in uploaded_files_np:
        try:
            # エンコーディングの自動判別を試みる (utf-8, shift_jis)
            df_temp = pd.read_csv(uploaded_file, encoding='utf-8')
            all_df_np_raw.append(df_temp)
        except UnicodeDecodeError:
            try:
                df_temp = pd.read_csv(uploaded_file, encoding='shift_jis')
                all_df_np_raw.append(df_temp)
            except Exception as e:
                st.error(f"ファイル '{uploaded_file.name}' の読み込み中にエラーが発生しました: {e}。エンコーディングを確認してください。")
                continue # 次のファイルへスキップ
        except Exception as e:
            st.error(f"ファイル '{uploaded_file.name}' の読み込み中にエラーが発生しました: {e}")
            continue # 次のファイルへスキップ
    
    if all_df_np_raw:
        # 全てのNP掛け払いCSVを結合
        df_np_raw_combined = pd.concat(all_df_np_raw, ignore_index=True)
        st.subheader("NP掛け払いデータプレビュー (結合後)")
        st.dataframe(df_np_raw_combined.head())

        # NP掛け払いデータの処理
        df_np = df_np_raw_combined.copy() # 処理用コピー
        
        # --- 列名の正規化 ---
        # 列名の前後の空白を除去し、英数字を半角に、カタカナを全角に変換
        new_columns = []
        for col in df_np.columns:
            normalized_col = col.strip() # 前後の空白を除去
            normalized_col = unicodedata.normalize('NFKC', normalized_col) # 全角と半角の統一（英数字は半角、カタカナは全角になりやすい）
            new_columns.append(normalized_col)
        df_np.columns = new_columns
        st.info(f"デバッグ: NPデータフレーム列名正規化後の列: {df_np.columns.tolist()}") # デバッグ表示

        # '企業名' 列の存在チェックと代替
        if '企業名' not in df_np.columns:
            st.warning("NP掛け払いCSVに '企業名' 列が見つかりませんでした。空文字列として処理を続行します。")
            df_np['企業名'] = '' # 存在しない場合は空の列を追加

        # '請求番号' 列の存在チェックと代替
        # 正規化後も '請求番号' がなければ、強制的に追加
        if '請求番号' not in df_np.columns:
            st.warning("NP掛け払いCSVに '請求番号' 列が見つかりませんでした。空文字列として列を追加します。")
            df_np['請求番号'] = '' # 存在しない場合は空の列を追加
            st.info(f"デバッグ: '請求番号' 列追加後のNPデータフレームの列: {df_np.columns.tolist()}") # デバッグ表示
        else:
            st.info(f"デバッグ: NPデータフレームに'請求番号'列は元から存在していました。現在の列: {df_np.columns.tolist()}") # デバッグ表示


        # 必須列の存在チェック (上で追加した'企業名', '請求番号'はここでチェックしない)
        required_np_columns_for_processing = ['請求書発行日', '支払期限日', '請求金額', '入金ステータス'] 
        missing_np_cols = [col for col in required_np_columns_for_processing if col not in df_np.columns]
        
        if missing_np_cols:
            st.error(f"NP掛け払いCSVに以下の必須列が見つかりません: {', '.join(missing_np_cols)}")
            df_np = None # 処理を中断
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
            
            st.info(f"デバッグ: df_np_processed作成直前のdf_npの列: {df_np.columns.tolist()}") # デバッグ表示
            if all(col in df_np.columns for col in cols_for_np_processed):
                df_np_processed = df_np[cols_for_np_processed].copy()
                df_np_processed = df_np_processed.rename(columns={
                    '請求金額': 'ご請求金額合計 (税込)', 
                    '支払期限日': 'お支払期日',
                    '請求番号': '請求書番号' # ★この行を追加★
                })
                
                st.subheader("NP掛け払い処理結果")
                st.dataframe(df_np_processed)
            else:
                st.error("NP掛け払い処理済みデータフレームの作成に必要な列が不足しています。予期せぬエラーが発生しました。不足している列は上記デバッグ情報をご確認ください。")
                df_np_processed = None 
    else:
        st.info("NP掛け払いCSVファイルがアップロードされていません。")


# --- バクラク請求書CSVのアップロード ---
st.header("2. バクラク請求書CSVのアップロード")
# 複数ファイルのアップロードを許可
uploaded_files_bakuraku = st.file_uploader(
    "バクラク請求書CSVファイルを複数選択してアップロードしてください。",
    type="csv",
    accept_multiple_files=True,
    key="bakuraku_uploader"
)

df_bakuraku_processed = None # 初期化

if uploaded_files_bakuraku: # ファイルがアップロードされた場合のみ処理
    all_df_bakuraku_raw = []
    for uploaded_file in uploaded_files_bakuraku:
        try:
            df_temp = pd.read_csv(uploaded_file, encoding='utf-8')
            all_df_bakuraku_raw.append(df_temp)
        except UnicodeDecodeError:
            try:
                df_temp = pd.read_csv(uploaded_file, encoding='shift_jis')
                all_df_bakuraku_raw.append(df_temp)
            except Exception as e:
                st.error(f"ファイル '{uploaded_file.name}' の読み込み中にエラーが発生しました: {e}。エンコーディングを確認してください。")
                continue 
        except Exception as e:
            st.error(f"ファイル '{uploaded_file.name}' の読み込み中にエラーが発生しました: {e}")
            continue 

    if all_df_bakuraku_raw:
        # 全てのバクラク請求書CSVを結合
        df_bakuraku_raw_combined = pd.concat(all_df_bakuraku_raw, ignore_index=True)
        st.subheader("バクラク請求書データプレビュー (結合後)")
        st.dataframe(df_bakuraku_raw_combined.head())

        # バクラク請求書データの処理
        df_bakuraku = df_bakuraku_raw_combined.copy() # 処理用コピー

        # --- 列名の正規化 ---
        new_columns_bakuraku = []
        for col in df_bakuraku.columns:
            normalized_col = col.strip()
            normalized_col = unicodedata.normalize('NFKC', normalized_col)
            new_columns_bakuraku.append(normalized_col)
        df_bakuraku.columns = new_columns_bakuraku
        st.info(f"デバッグ: バクラクデータフレーム列名正規化後の列: {df_bakuraku.columns.tolist()}") # デバッグ表示


        # 必須列の存在チェック
        required_bakuraku_columns = ['日付', '支払期日', '書類種別', '書類番号', '送付先名', '金額'] # '書類種別'も確認
        missing_bakuraku_cols = [col for col in required_bakuraku_columns if col not in df_bakuraku.columns]
        if missing_bakuraku_cols:
            st.error(f"バクラク請求書CSVに以下の必須列が見つかりません: {', '.join(missing_bakuraku_cols)}")
            df_bakuraku = None 
        else:
            df_bakuraku['日付'] = pd.to_datetime(df_bakuraku['日付'], errors='coerce')
            df_bakuraku['支払期日'] = pd.to_datetime(df_bakuraku['支払期日'], errors='coerce')

            if df_bakuraku['日付'].isnull().any() or df_bakuraku['支払期日'].isnull().any():
                st.warning("バクラク請求書CSVの日付列に無効な値がありました。該当行はNaNとして処理されます。")

            # 'ご請求方法'を'書類種別'から取得するように修正
            df_bakuraku['ご請求方法'] = df_bakuraku['書類種別'].fillna('不明') # '書類種別'が存在しない場合やNaNの場合の対応
            
            df_bakuraku['金額'] = pd.to_numeric(df_bakuraku['金額'], errors='coerce').fillna(0)
            df_bakuraku['ご請求金額合計 (税込)'] = df_bakuraku['金額']

            st.subheader("バクラク請求書 未入金状況選択")
            st.write("未入金の請求書にチェックを入れてください。")
            
            selected_unpaid_bakuraku = {}
            if df_bakuraku is not None and not df_bakuraku.empty: 
                with st.expander("バクラク請求書一覧を開く"):
                    # 'original_index' を一時的に追加して、後で元の行にマッピングできるようにする
                    df_bakuraku_temp = df_bakuraku.reset_index().rename(columns={'index': 'original_index'})
                    df_bakuraku_display = df_bakuraku_temp[['original_index', '書類番号', '日付', '金額']].drop_duplicates(subset=['書類番号', '日付', '金額']).set_index('original_index')

                    for original_idx, row in df_bakuraku_display.iterrows():
                        date_str = row['日付'].strftime('%Y-%m-%d') if pd.notna(row['日付']) else '日付不明'
                        amount_display = f"{row['金額']:,}円" if pd.notna(row['金額']) else '金額不明'
                        
                        # チェックボックスの初期状態は、この表示行に対応する元のレコードの選択状態を反映
                        # ここでは、特定のoriginal_idxのselected_unpaid_bakurakuの状態を利用
                        # ただし、同じ「書類番号,日付,金額」を持つ複数のoriginal_idxがある場合、最初のoriginal_idxの状態を代表として使用
                        is_checked = selected_unpaid_bakuraku.get(original_idx, False)

                        unique_key = f"bakuraku_unpaid_{row['書類番号']}_{date_str}_{row['金額']}_{original_idx}" # original_idxを含めてキーをユニークに
                        
                        if st.checkbox(f"書類番号: {row['書類番号']}, 日付: {date_str}, 金額: {amount_display}", key=unique_key, value=is_checked):
                            # この表示行に対応する全ての元のデータフレームのレコードの入金状態を更新
                            # 条件に一致する全てのインデックスを取得
                            matching_indices = df_bakuraku[(df_bakuraku['書類番号'] == row['書類番号']) & 
                                                           (df_bakuraku['日付'] == row['日付']) & 
                                                           (df_bakuraku['金額'] == row['金額'])].index
                            for idx in matching_indices:
                                selected_unpaid_bakuraku[idx] = True
                        else:
                            matching_indices = df_bakuraku[(df_bakuraku['書類番号'] == row['書類番号']) & 
                                                           (df_bakuraku['日付'] == row['日付']) & 
                                                           (df_bakuraku['金額'] == row['金額'])].index
                            for idx in matching_indices:
                                selected_unpaid_bakuraku[idx] = False

            if df_bakuraku is not None:
                df_bakuraku['入金有無'] = df_bakuraku.index.map(lambda idx: 'なし' if selected_unpaid_bakuraku.get(idx, False) else 'あり')
                df_bakuraku['未入金金額合計 (税込)'] = df_bakuraku.apply(lambda row: row['金額'] if row['入金有無'] == 'なし' else 0, axis=1)
                
                # 最終的な処理済みデータフレームを作成する際に、同じ請求書をグループ化
                # 'ご請求方法' もグループ化キーに含める
                df_bakuraku_processed = df_bakuraku.groupby(['書類番号', '日付', '支払期日', '送付先名', 'ご請求方法']).agg(
                    金額合計=('ご請求金額合計 (税込)', 'sum'),
                    未入金合計=('未入金金額合計 (税込)', 'sum'),
                    入金有無=('入金有無', lambda x: 'なし' if 'なし' in x.values else 'あり') # 一つでも「なし」があれば「なし」
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
if df_np_processed is not None and df_bakuraku_processed is not None: 
    st.header("3. 統合された請求および入金状況")
    
    df_np_processed['ご利用年月'] = df_np_processed['請求書発行日'].apply(lambda x: x.strftime('%Y年%m月') if pd.notna(x) else '')
    df_bakuraku_processed['ご利用年月'] = df_bakuraku_processed['請求書発行日'].apply(lambda x: x.strftime('%Y年%m月') if pd.notna(x) else '')

    common_cols = ['ご利用年月', 'ご請求方法', 'ご請求金額合計 (税込)', '未入金金額合計 (税込)', '請求書番号', '請求書発行日', 'お支払期日', '入金有無', '企業名']
    
    # 結合する前に、各DFが共通の列を持っているか最終確認
    missing_cols_np = [col for col in common_cols if col not in df_np_processed.columns]
    missing_cols_bakuraku = [col for col in common_cols if col not in df_bakuraku_processed.columns]

    if missing_cols_np:
        st.error(f"NP掛け払い処理結果データに結合に必要な列が不足しています: {', '.join(missing_cols_np)}")
        combined_df_with_total = None
    elif missing_cols_bakuraku:
        st.error(f"バクラク請求書処理結果データに結合に必要な列が不足しています: {', '.join(missing_cols_bakuraku)}")
        combined_df_with_total = None
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
            if col in combined_df.columns and combined_df[col].dtype != total_row[col].dtype:
                total_row[col] = total_row[col].astype(combined_df[col].dtype)
        
        combined_df_with_total = pd.concat([combined_df, total_row], ignore_index=True)


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
        st.error("データの結合または処理に問題が発生したため、統合された結果は表示できません。上記のエラーメッセージを確認してください。")

else: 
    if uploaded_files_np is None or not uploaded_files_np: 
        st.info("NP掛け払いCSVファイルをアップロードしてください。")
    if uploaded_files_bakuraku is None or not uploaded_files_bakuraku: 
        st.info("バクラク請求書CSVファイルをアップロードしてください。")

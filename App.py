import streamlit as st
import pandas as pd
import io
import datetime

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
        
        # '企業名' 列の存在チェックと代替
        if '企業名' not in df_np.columns:
            st.warning("NP掛け払いCSVに '企業名' 列が見つかりませんでした。空文字列として処理を続行します。")
            df_np['企業名'] = '' # 存在しない場合は空の列を追加

        # '請求番号' 列の存在チェックと代替 ★ここを強化します★
        # どの時点でのdf_npに列が存在しないか確認するため、デバッグ表示を追加
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

            # 日付変換エラーのチェック
            if df_np['請求書発行日'].isnull().any() or df_np['支払期日'].isnull().any():
                st.warning("NP掛け払いCSVの日付列に無効な値がありました。該当行はNaNとして処理されます。")

            df_np['入金有無'] = df_np['入金ステータス'].apply(lambda x: 'あり' if x == '入金済み' else 'なし')
            df_np['ご請求方法'] = 'NP掛け払い'
            # 請求金額が数値でない場合に備えてエラー処理
            df_np['請求金額'] = pd.to_numeric(df_np['請求金額'], errors='coerce').fillna(0)
            df_np['未入金金額合計 (税込)'] = df_np.apply(lambda row: row['請求金額'] if row['入金有無'] == 'なし' else 0, axis=1)
            
            # df_np_processed を作成するために必要な列リスト
            cols_for_np_processed = ['請求書発行日', '支払期限日', '請求番号', '企業名', 'ご請求方法', '請求金額', '未入金金額合計 (税込)', '入金有無']
            
            # df_np にすべての必要な列があることを最終確認
            st.info(f"デバッグ: df_np_processed作成直前のdf_npの列: {df_np.columns.tolist()}") # デバッグ表示
            if all(col in df_np.columns for col in cols_for_np_processed):
                df_np_processed = df_np[cols_for_np_processed].copy()
                df_np_processed = df_np_processed.rename(columns={'請求金額': 'ご請求金額合計 (税込)'})
                df_np_processed = df_np_processed.rename(columns={'支払期限日': 'お支払期日'}) 
                
                st.subheader("NP掛け払い処理結果")
                st.dataframe(df_np_processed)
            else:
                st.error("NP掛け払い処理済みデータフレームの作成に必要な列が不足しています。予期せぬエラーが発生しました。不足している列は上記デバッグ情報をご確認ください。")
                df_np_processed = None # エラーのため処理済みデータフレームをNoneにする
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
                continue # 次のファイルへスキップ
        except Exception as e:
            st.error(f"ファイル '{uploaded_file.name}' の読み込み中にエラーが発生しました: {e}")
            continue # 次のファイルへスキップ

    if all_df_bakuraku_raw:
        # 全てのバクラク請求書CSVを結合
        df_bakuraku_raw_combined = pd.concat(all_df_bakuraku_raw, ignore_index=True)
        st.subheader("バクラク請求書データプレビュー (結合後)")
        st.dataframe(df_bakuraku_raw_combined.head())

        # バクラク請求書データの処理
        df_bakuraku = df_bakuraku_raw_combined.copy() # 処理用コピー

        # 必須列の存在チェック
        required_bakuraku_columns = ['日付', '支払期日', '書類番号', '送付先名', '金額']
        missing_bakuraku_cols = [col for col in required_bakuraku_columns if col not in df_bakuraku.columns]
        if missing_bakuraku_cols:
            st.error(f"バクラク請求書CSVに以下の必須列が見つかりません: {', '.join(missing_bakuraku_cols)}")
            df_bakuraku = None # 処理を中断
        else:
            df_bakuraku['日付'] = pd.to_datetime(df_bakuraku['日付'], errors='coerce')
            df_bakuraku['支払期日'] = pd.to_datetime(df_bakuraku['支払期日'], errors='coerce')

            # 日付変換エラーのチェック
            if df_bakuraku['日付'].isnull().any() or df_bakuraku['支払期日'].isnull().any():
                st.warning("バクラク請求書CSVの日付列に無効な値がありました。該当行はNaNとして処理されます。")

            df_bakuraku['ご請求方法'] = '直接請求' # 仮に設定
            # 金額が数値でない場合に備えてエラー処理
            df_bakuraku['金額'] = pd.to_numeric(df_bakuraku['金額'], errors='coerce').fillna(0)
            df_bakuraku['ご請求金額合計 (税込)'] = df_bakuraku['金額']

            st.subheader("バクラク請求書 未入金状況選択")
            st.write("未入金の請求書にチェックを入れてください。")
            
            selected_unpaid_bakuraku = {}
            if df_bakuraku is not None and not df_bakuraku.empty: # データフレームが空でないことを確認
                with st.expander("バクラク請求書一覧を開く"):
                    # ユニークなキーを生成するために、書類番号と日付、金額を組み合わせる
                    # 表示用には重複を除去したものを利用 (同一請求を1つとして表示)
                    # 元のdf_bakurakuのインデックスを保存するため、merge後に設定
                    df_bakuraku_display = df_bakuraku[['書類番号', '日付', '金額']].drop_duplicates().reset_index()

                    for display_idx, row in df_bakuraku_display.iterrows():
                        original_idx_for_display = row['index'] # この行に対応する元のdf_bakurakuのインデックス
                        
                        # NaNの日付は表示しないか、適切に処理
                        date_str = row['日付'].strftime('%Y-%m-%d') if pd.notna(row['日付']) else '日付不明'
                        # 金額がNaT/Noneの場合に備えて0を表示
                        amount_display = f"{row['金額']:,}円" if pd.notna(row['金額']) else '金額不明'
                        
                        unique_key = f"bakuraku_unpaid_{row['書類番号']}_{date_str}_{row['金額']}_{display_idx}"
                        
                        # チェックボックスの初期状態は、この表示行に対応する最初の元のレコードの選択状態を反映
                        is_checked = selected_unpaid_bakuraku.get(original_idx_for_display, False)

                        if st.checkbox(f"書類番号: {row['書類番号']}, 日付: {date_str}, 金額: {amount_display}", key=unique_key, value=is_checked):
                            # この表示行に対応する全ての元のデータフレームのレコードの入金状態を更新
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

            # 選択結果をdf_bakurakuに適用
            if df_bakuraku is not None:
                # selected_unpaid_bakuraku辞書を基にdf_bakurakuの'入金有無'を更新
                # 辞書に存在しないインデックスはデフォルトで'あり'とする
                df_bakuraku['入金有無'] = df_bakuraku.index.map(lambda idx: 'なし' if selected_unpaid_bakuraku.get(idx, False) else 'あり')
                df_bakuraku['未入金金額合計 (税込)'] = df_bakuraku.apply(lambda row: row['金額'] if row['入金有無'] == 'なし' else 0, axis=1)
                
                # 最終的な処理済みデータフレームを作成する際に、同じ請求書をグループ化
                df_bakuraku_processed = df_bakuraku.groupby(['書類番号', '日付', '支払期日', '送付先名', 'ご請求方法']).agg(
                    金額合計=('ご請求金額合計 (税込)', 'sum'),
                    未入金合計=('未入金金額合計 (税込)', 'sum'),
                    入金有無=('入金有無', lambda x: 'なし' if 'なし' in x.values else 'あり') # 一つでも「なし」があれば「なし」
                ).reset_index()
                
                # df_bakuraku_processed の作成時に、結合用に列名を統一
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
# 両方のデータがアップロードされ、処理された場合のみ表示
if df_np_processed is not None and df_bakuraku_processed is not None: 
    st.header("3. 統合された請求および入金状況")
    
    # 統合データフレームの作成
    # 日付がNaTでないことを確認してからフォーマット
    # df_np_processed['請求書発行日']がpd.Timestamp型であることを保証
    # df_np_processed['ご利用年月'] = df_np_processed['請求書発行日'].dt.strftime('%Y年%m月') # 既にTimestampなのでエラーにならないはず
    df_np_processed['ご利用年月'] = df_np_processed['請求書発行日'].apply(lambda x: x.strftime('%Y年%m月') if pd.notna(x) else '')
    df_bakuraku_processed['ご利用年月'] = df_bakuraku_processed['請求書発行日'].apply(lambda x: x.strftime('%Y年%m月') if pd.notna(x) else '')

    # 結合する際に使用する共通列の定義
    common_cols = ['ご利用年月', 'ご請求方法', 'ご請求金額合計 (税込)', '未入金金額合計 (税込)', '請求書番号', '請求書発行日', 'お支払期日', '入金有無', '企業名']
    
    # 結合する前に、各DFが共通の列を持っているか最終確認
    if not all(col in df_np_processed.columns for col in common_cols):
        st.error(f"NP掛け払い処理結果データに結合に必要な列が不足しています: {', '.join(set(common_cols) - set(df_np_processed.columns))}")
        combined_df_with_total = None
    elif not all(col in df_bakuraku_processed.columns for col in common_cols):
        st.error(f"バクラク請求書処理結果データに結合に必要な列が不足しています: {', '.join(set(common_cols) - set(df_bakuraku_processed.columns))}")
        combined_df_with_total = None
    else:
        # pd.Timestamp と NaT が混在する可能性があるため、concat後に適切な型に変換
        combined_df = pd.concat([
            df_np_processed[common_cols],
            df_bakuraku_processed[common_cols]
        ])
        
        # 請求書発行日と支払期日のdtypeをdatetimeに統一（NaT含む）
        combined_df['請求書発行日'] = pd.to_datetime(combined_df['請求書発行日'], errors='coerce')
        combined_df['お支払期日'] = pd.to_datetime(combined_df['お支払期日'], errors='coerce')

        # ソート
        # 日付がNaTの場合でもエラーにならないようna_position='last'を指定
        combined_df = combined_df.sort_values(by=['ご利用年月', '請求書発行日'], na_position='last').reset_index(drop=True)

        # 合計行の追加
        total_請求金額 = combined_df['ご請求金額合計 (税込)'].sum()
        total_未入金金額 = combined_df['未入金金額合計 (税込)'].sum()
        
        total_row = pd.DataFrame([{
            'ご利用年月': '',
            'ご請求方法': '合計',
            'ご請求金額合計 (税込)': total_請求金額,
            '未入金金額合計 (税込)': total_未入金金額,
            '請求書番号': '',
            '請求書発行日': pd.NaT, # 日付型としてpd.NaTを設定
            'お支払期日': pd.NaT, # 日付型としてpd.NaTを設定
            '入金有無': '',
            '企業名': '' # 合計行にも'企業名'列を追加
        }])
        
        # combined_df_with_totalを作成する前に、combined_dfとtotal_rowの列順と型をできるだけ合わせる
        # pd.concatが失敗しないようにtotal_rowの型を調整
        # (以前のコードでこれを自動的に行うようにしていましたが、明示的にキャストします)
        # combined_dfの数値列はfloat64になりうるため、total_rowの数値も合わせる
        for col in ['ご請求金額合計 (税込)', '未入金金額合計 (税込)']:
            if col in combined_df.columns and combined_df[col].dtype != total_row[col].dtype:
                total_row[col] = total_row[col].astype(combined_df[col].dtype)
        
        combined_df_with_total = pd.concat([combined_df, total_row], ignore_index=True)


    if combined_df_with_total is not None:
        # Streamlitでの表示
        # 企業名を動的に取得（最初のNP掛け払いデータまたはバクラク請求書の送付先名から）
        company_name = "不明な企業"
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
        
        # もし企業名が空の文字列だったらデフォルト値を設定
        if not company_name:
            company_name = "取引先"


        st.markdown(f"### {company_name}さま")
        st.markdown("### ご請求およびご入金状況一覧")
        st.markdown(f"**作成日: {datetime.date.today().strftime('%Y/%m/%d')}**")

        # DataFrameをそのまま表示 (金額にカンマ区切りフォーマットを適用)
        st.dataframe(combined_df_with_total.style.format({
            'ご請求金額合計 (税込)': '{:,.0f}',
            '未入金金額合計 (税込)': '{:,.0f}',
            # 日付列のフォーマット (NaTは空文字列)
            '請求書発行日': lambda x: x.strftime('%Y/%m/%d') if pd.notna(x) else '',
            'お支払期日': lambda x: x.strftime('%Y/%m/%d') if pd.notna(x) else ''
        }))

        st.markdown(f"**※{datetime.date.today().strftime('%Y年%m月')}時点での未入金合計金額: {total_未入金金額:,}円**")


        # --- Excel出力ボタン ---
        current_date_str = datetime.date.today().strftime('%Y%m%d_%H%M%S')

        excel_buffer = io.BytesIO()
        
        # Excel出力用に日付列を文字列に変換（NaNは空文字列に）
        output_df = combined_df_with_total.copy() # 合計行も含んだデータフレームをコピー
        output_df['請求書発行日'] = output_df['請求書発行日'].dt.strftime('%Y/%m/%d').fillna('')
        output_df['お支払期日'] = output_df['お支払期日'].dt.strftime('%Y/%m/%d').fillna('')
        
        # Excelに書き出し
        output_df.to_excel(excel_buffer, index=False, sheet_name='請求入金状況', engine='openpyxl')
        excel_buffer.seek(0) # バッファの先頭に戻す

        st.download_button(
            label="Excelファイルとしてダウンロード",
            data=excel_buffer,
            file_name=f"請求入金状況_{current_date_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.info("Excelファイルとしてダウンロード可能です。")
    else:
        st.error("データの結合または処理に問題が発生したため、統合された結果は表示できません。上記のエラーメッセージを確認してください。")

else: # どちらかのCSVがアップロードされていない場合、または処理に失敗した場合
    if uploaded_files_np is None:
        st.info("NP掛け払いCSVファイルをアップロードしてください。")
    if uploaded_files_bakuraku is None:
        st.info("バクラク請求書CSVファイルをアップロードしてください。")

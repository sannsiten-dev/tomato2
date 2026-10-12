import streamlit as st
import pandas as pd
import io
import re

# ページ基本設定
st.set_page_config(page_title="トマトらーめん出数＆構成比集計システム", layout="wide")

st.title("🍅 トマトらーめん グループ別出数・客数構成比集計アプリ")
st.write("「日別メニュー出数.xlsx」および「カレンダー売上一覧.xlsx」をアップロードしてください。")

col1, col2 = st.columns(2)
with col1:
    menu_file = st.file_uploader("1. 日別メニュー出数 (.xlsx)", type=["xlsx"], key="menu")
with col2:
    cal_file = st.file_uploader("2. カレンダー売上一覧 (.xlsx)", type=["xlsx"], key="cal")

if menu_file is not None:
    try:
        # --- 1. メニュー出数ファイルの解析 ---
        excel_data = pd.ExcelFile(menu_file)
        sheet_name = excel_data.sheet_names[0]
        df_menu = pd.read_excel(menu_file, sheet_name=sheet_name)
        
        header_days = df_menu.iloc[3, 6:].values
        header_weekdays = df_menu.iloc[4, 6:].values
        
        menu_rows = df_menu.iloc[5:30].copy()
        day_cols = list(header_days)
        menu_rows.columns = ['商品CD', '商品名', '消費税率', '売上金額', '販売数量', '平均販売数'] + day_cols
        
        group1_names = ['野菜トマトらーめん', '小・トマトらーめん', '☆野菜トマトらーめん', '☆小・トマトらーめん']
        group2_names = ['チーズまみれトマト', '☆チーズまみれトマト', 'チーズまみれらーめん']
        group3_names = ['トマト唐麺', '☆トマト唐麺']
        group4_names = ['スパイシートマト', '小・スパイシートマト', '☆スパイシートマト', '☆小・スパイシトマト']
        
        def get_group_sum(group_names):
            sub_df = menu_rows[menu_rows['商品名'].isin(group_names)]
            return sub_df[day_cols].fillna(0).astype(int).sum(axis=0)
        
        g1_sum = get_group_sum(group1_names)
        g2_sum = get_group_sum(group2_names)
        g3_sum = get_group_sum(group3_names)
        g4_sum = get_group_sum(group4_names)
        
        # --- 2. カレンダーファイルの解析（客数の自動抽出） ---
        date_customer_map = {}
        if cal_file is not None:
            df_cal = pd.read_excel(cal_file, sheet_name=0)
            for r in range(4, len(df_cal), 4):
                if r >= len(df_cal):
                    break
                days_row = df_cal.iloc[r].values
                cust_row = df_cal.iloc[r+2].values if (r+2) < len(df_cal) else [None]*7
                
                for day_val, cust_val in zip(days_row, cust_row):
                    if pd.notna(day_val):
                        day_match = re.search(r'(\d{1,2})', str(day_val))
                        if day_match:
                            day_str = day_match.group(1).zfill(2)
                            if pd.notna(cust_val):
                                cust_match = re.search(r'/\s*(\d+)\s*人', str(cust_val))
                                if cust_match:
                                    date_customer_map[day_str] = int(cust_match.group(1))

        # --- 3. 日別データの構築 ---
        day_labels = [f"{str(day).zfill(2)}({wd})" for day, wd in zip(header_days, header_weekdays)]
        customers = [date_customer_map.get(str(day).zfill(2), 0) for day in header_days]
        
        daily_records = []
        for idx in range(len(header_days)):
            lbl = day_labels[idx]
            wd = str(header_weekdays[idx]).strip()
            cust = customers[idx]
            
            v1 = g1_sum.values[idx]
            v4 = g4_sum.values[idx]
            v3 = g3_sum.values[idx]
            v2 = g2_sum.values[idx]
            
            # 出数が全て0の日（販売開始前等）は除外オプション用に判定できるようにしておく
            is_zero_day = (v1 + v4 + v3 + v2) == 0
            
            daily_records.append({
                '日付(曜日)': lbl,
                '曜日': wd,
                '客数': cust,
                '野菜トマト 出数': v1,
                'スパイシートマト 出数': v4,
                'トマト唐麺 出数': v3,
                '夜チーズ 出数': v2,
                'is_zero_day': is_zero_day
            })
            
        # --- 4. 週計行を含むテーブルの組み立て ---
        rows_with_weekly = []
        current_week = []
        
        for rec in daily_records:
            # ゼロの日を自動非表示にする設定の準備
            current_week.append(rec)
            
            # 日曜日（またはデータ最終日）で週計を計算・挿入
            if rec['曜日'] in ['日', 'Sun'] or rec == daily_records[-1]:
                # 有効な日（出数が0でない日）がある場合のみ処理
                non_zero_days = [r for r in current_week if not r['is_zero_day']]
                
                if non_zero_days:
                    for r in current_week:
                        if not r['is_zero_day']:
                            c = r['客数']
                            rows_with_weekly.append({
                                '日付(曜日)': r['日付(曜日)'],
                                '客数': c,
                                '野菜トマト 出数': r['野菜トマト 出数'],
                                '野菜トマト 構成比': r['野菜トマト 出数'] / c if c > 0 else 0.0,
                                'スパイシートマト 出数': r['スパイシートマト 出数'],
                                'スパイシートマト 構成比': r['スパイシートマト 出数'] / c if c > 0 else 0.0,
                                'トマト唐麺 出数': r['トマト唐麺 出数'],
                                'トマト唐麺 構成比': r['トマト唐麺 出数'] / c if c > 0 else 0.0,
                                '夜チーズ 出数': r['夜チーズ 出数'],
                                'is_summary': False
                            })
                    
                    # 週計行の追加
                    w_cust = sum(r['客数'] for r in current_week if not r['is_zero_day'])
                    w_v1 = sum(r['野菜トマト 出数'] for r in current_week if not r['is_zero_day'])
                    w_v4 = sum(r['スパイシートマト 出数'] for r in current_week if not r['is_zero_day'])
                    w_v3 = sum(r['トマト唐麺 出数'] for r in current_week if not r['is_zero_day'])
                    w_v2 = sum(r['夜チーズ 出数'] for r in current_week if not r['is_zero_day'])
                    
                    rows_with_weekly.append({
                        '日付(曜日)': '★ 週計',
                        '客数': w_cust,
                        '野菜トマト 出数': w_v1,
                        '野菜トマト 構成比': w_v1 / w_cust if w_cust > 0 else 0.0,
                        'スパイシートマト 出数': w_v4,
                        'スパイシートマト 構成比': w_v4 / w_cust if w_cust > 0 else 0.0,
                        'トマト唐麺 出数': w_v3,
                        'トマト唐麺 構成比': w_v3 / w_cust if w_cust > 0 else 0.0,
                        '夜チーズ 出数': w_v2,
                        'is_summary': True
                    })
                
                current_week = []
                
        res_df = pd.DataFrame(rows_with_weekly)
        display_df = res_df.drop(columns=['is_summary'])
        
        # --- 5. 表示・フォーマット ---
        st.subheader("📊 集計結果（日別 ＆ 週計対客数構成比）")
        
        format_dict = {}
        if cal_file is not None:
            format_dict['野菜トマト 構成比'] = "{:.2%}"
            format_dict['スパイシートマト 構成比'] = "{:.2%}"
            format_dict['トマト唐麺 構成比'] = "{:.2%}"
            
        # 強調表示用スタイリング関数（週計行を背景強調）
        def highlight_weekly(row):
            if row['日付(曜日)'] == '★ 週計':
                return ['background-color: #f0f2f6; font-weight: bold'] * len(row)
            return [''] * len(row)
            
        st.dataframe(
            display_df.style.format(format_dict).apply(highlight_weekly, axis=1),
            use_container_width=True
        )
        
        # --- 6. Excelダウンロード機能 ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            display_df.to_excel(writer, sheet_name='週計付き出数表', index=False)
        processed_data = output.getvalue()
        
        st.download_button(
            label="📥 週計入り集計表をExcelでダウンロード",
            data=processed_data,
            file_name="トマト出数_週計付き集計表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("週計を含む集計が完了しました！")

    except Exception as e:
        st.error(f"ファイル解析中にエラーが発生しました: {e}")

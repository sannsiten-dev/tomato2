import streamlit as st
import pandas as pd
import io
import re

# ページ基本設定
st.set_page_config(page_title="トマトらーめん出数＆構成比集計システム", layout="wide")

st.title("🍅 トマトらーめん グループ別出数・構成比・チーズ総量集計アプリ")
st.write("必要なファイル（1.限定麺、2.カレンダー、3.トッピング）をアップロードしてください。")

col1, col2, col3 = st.columns(3)
with col1:
    menu_file = st.file_uploader("1. 日別メニュー出数 (限定麺)", type=["xlsx"], key="menu")
with col2:
    cal_file = st.file_uploader("2. カレンダー売上一覧", type=["xlsx"], key="cal")
with col3:
    top_file = st.file_uploader("3. 日別メニュー出数 (トッピング)", type=["xlsx"], key="top")

if menu_file is not None:
    try:
        # --- 1. 限定麺メニュー出数ファイルの解析 ---
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
        
        def get_group_sum(sub_df):
            return sub_df[day_cols].fillna(0).astype(int).sum(axis=0)
        
        g1_sum = get_group_sum(menu_rows[menu_rows['商品名'].isin(group1_names)])
        g2_sum = get_group_sum(menu_rows[menu_rows['商品名'].isin(group2_names)])
        g3_sum = get_group_sum(menu_rows[menu_rows['商品名'].isin(group3_names)])
        g4_sum = get_group_sum(menu_rows[menu_rows['商品名'].isin(group4_names)])
        
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

        # --- 3. トッピングファイルの解析（チーズ増の自動抽出） ---
        cheese_add_map = {str(d).zfill(2): 0 for d in header_days}
        if top_file is not None:
            df_top = pd.read_excel(top_file, sheet_name=0)
            top_rows = df_top.iloc[5:31].copy()
            top_days = list(df_top.iloc[3, 6:].values)
            top_rows.columns = ['商品CD', '商品名', '消費税率', '売上金額', '販売数量', '平均販売数'] + top_days
            
            cheese_names = ['☆とろけるチーズ', '◇とろけるチーズ', 'とろけるチーズ']
            cheese_sub = top_rows[top_rows['商品名'].isin(cheese_names)]
            c_sum = cheese_sub[top_days].fillna(0).astype(int).sum(axis=0)
            
            for d, val in zip(top_days, c_sum.values):
                cheese_add_map[str(d).zfill(2)] = int(val)

        # --- 4. 日別データの構築 ---
        day_labels = [f"{str(day).zfill(2)}({wd})" for day, wd in zip(header_days, header_weekdays)]
        customers = [date_customer_map.get(str(day).zfill(2), 0) for day in header_days]
        
        daily_records = []
        for idx in range(len(header_days)):
            d_str = str(header_days[idx]).zfill(2)
            lbl = day_labels[idx]
            wd = str(header_weekdays[idx]).strip()
            cust = customers[idx]
            
            v1 = g1_sum.values[idx]
            v4 = g4_sum.values[idx]
            v3 = g3_sum.values[idx]
            v2 = g2_sum.values[idx]
            c_add = cheese_add_map.get(d_str, 0)
            
            # チーズ総量 = 夜チーズ * 3 + チーズ増
            c_total = (v2 * 3) + c_add
            
            is_zero_day = (v1 + v4 + v3 + v2 + c_add) == 0
            
            daily_records.append({
                '日付(曜日)': lbl,
                '曜日': wd,
                '客数': cust,
                '野菜トマト 出数': v1,
                'スパイシートマト 出数': v4,
                'トマト唐麺 出数': v3,
                '夜チーズ 出数': v2,
                'チーズ増 出数': c_add,
                'チーズ総量': c_total,
                'is_zero_day': is_zero_day
            })
            
        # --- 5. 週計行を含むテーブルの組み立て ---
        rows_with_weekly = []
        current_week = []
        
        for rec in daily_records:
            current_week.append(rec)
            
            if rec['曜日'] in ['日', 'Sun'] or rec == daily_records[-1]:
                non_zero_days = [r for r in current_week if not r['is_zero_day']]
                
                if non_zero_days:
                    for r in current_week:
                        if not r['is_zero_day']:
                            c = r['客数']
                            row_dict = {'日付(曜日)': r['日付(曜日)']}
                            if cal_file is not None:
                                row_dict['客数'] = c
                            
                            row_dict['野菜トマト 出数'] = r['野菜トマト 出数']
                            if cal_file is not None:
                                row_dict['野菜トマト 構成比'] = r['野菜トマト 出数'] / c if c > 0 else 0.0
                                
                            row_dict['スパイシートマト 出数'] = r['スパイシートマト 出数']
                            if cal_file is not None:
                                row_dict['スパイシートマト 構成比'] = r['スパイシートマト 出数'] / c if c > 0 else 0.0
                                
                            row_dict['トマト唐麺 出数'] = r['トマト唐麺 出数']
                            if cal_file is not None:
                                row_dict['トマト唐麺 構成比'] = r['トマト唐麺 出数'] / c if c > 0 else 0.0
                                
                            row_dict['夜チーズ 出数'] = r['夜チーズ 出数']
                            if top_file is not None:
                                row_dict['チーズ増 出数'] = r['チーズ増 出数']
                                row_dict['チーズ総量(食)'] = r['チーズ総量']
                                
                            rows_with_weekly.append(row_dict)
                    
                    # 週計行の計算
                    w_cust = sum(r['客数'] for r in current_week if not r['is_zero_day'])
                    w_v1 = sum(r['野菜トマト 出数'] for r in current_week if not r['is_zero_day'])
                    w_v4 = sum(r['スパイシートマト 出数'] for r in current_week if not r['is_zero_day'])
                    w_v3 = sum(r['トマト唐麺 出数'] for r in current_week if not r['is_zero_day'])
                    w_v2 = sum(r['夜チーズ 出数'] for r in current_week if not r['is_zero_day'])
                    w_cadd = sum(r['チーズ増 出数'] for r in current_week if not r['is_zero_day'])
                    w_ctotal = sum(r['チーズ総量'] for r in current_week if not r['is_zero_day'])
                    
                    w_row = {'日付(曜日)': '★ 週計'}
                    if cal_file is not None:
                        w_row['客数'] = w_cust
                        
                    w_row['野菜トマト 出数'] = w_v1
                    if cal_file is not None:
                        w_row['野菜トマト 構成比'] = w_v1 / w_cust if w_cust > 0 else 0.0
                        
                    w_row['スパイシートマト 出数'] = w_v4
                    if cal_file is not None:
                        w_row['スパイシートマト 構成比'] = w_v4 / w_cust if w_cust > 0 else 0.0
                        
                    w_row['トマト唐麺 出数'] = w_v3
                    if cal_file is not None:
                        w_row['トマト唐麺 構成比'] = w_v3 / w_cust if w_cust > 0 else 0.0
                        
                    w_row['夜チーズ 出数'] = w_v2
                    if top_file is not None:
                        w_row['チーズ増 出数'] = w_cadd
                        w_row['チーズ総量(食)'] = w_ctotal
                        
                    rows_with_weekly.append(w_row)
                
                current_week = []
                
        res_df = pd.DataFrame(rows_with_weekly)
        
        # --- 6. 表示・フォーマット ---
        st.subheader("📊 集計結果（出数・構成比・チーズ総量）")
        
        format_dict = {}
        if cal_file is not None:
            format_dict['野菜トマト 構成比'] = "{:.2%}"
            format_dict['スパイシートマト 構成比'] = "{:.2%}"
            format_dict['トマト唐麺 構成比'] = "{:.2%}"
            
        def highlight_weekly(row):
            if row['日付(曜日)'] == '★ 週計':
                return ['background-color: #e6f0fa; font-weight: bold'] * len(row)
            return [''] * len(row)
            
        st.dataframe(
            res_df.style.format(format_dict).apply(highlight_weekly, axis=1),
            use_container_width=True
        )
        
        # --- 7. Excelダウンロード機能 ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            res_df.to_excel(writer, sheet_name='出数・チーズ総量集計表', index=False)
        processed_data = output.getvalue()
        
        st.download_button(
            label="📥 集計結果（出数・チーズ総量表）をExcelでダウンロード",
            data=processed_data,
            file_name="トマト出数_チーズ総量集計表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("チーズ増およびチーズ総量を含む集計が完了しました！")

    except Exception as e:
        st.error(f"ファイル解析中にエラーが発生しました: {e}")

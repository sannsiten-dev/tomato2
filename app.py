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

        # --- 3. 集計テーブルの作成 ---
        day_labels = [f"{str(day).zfill(2)}({wd})" for day, wd in zip(header_days, header_weekdays)]
        customers = [date_customer_map.get(str(day).zfill(2), 0) for day in header_days]
        
        res_data = {'日付(曜日)': day_labels}
        
        if cal_file is not None:
            res_data['客数'] = customers
            
        res_data['野菜トマト 出数'] = g1_sum.values
        if cal_file is not None:
            res_data['野菜トマト 構成比'] = [out / cust if cust > 0 else 0.0 for out, cust in zip(g1_sum.values, customers)]
            
        res_data['スパイシートマト 出数'] = g4_sum.values
        if cal_file is not None:
            res_data['スパイシートマト 構成比'] = [out / cust if cust > 0 else 0.0 for out, cust in zip(g4_sum.values, customers)]
            
        res_data['トマト唐麺 出数'] = g3_sum.values
        if cal_file is not None:
            res_data['トマト唐麺 構成比'] = [out / cust if cust > 0 else 0.0 for out, cust in zip(g3_sum.values, customers)]
            
        res_data['夜チーズ 出数'] = g2_sum.values
        
        res_df = pd.DataFrame(res_data)
        
        # 表示コントロール
        st.subheader("📊 集計結果（出数＆対客数構成比）")
        
        hide_zero = st.checkbox("販売期間外（出数がすべて0の日）を隠す", value=True)
        display_df = res_df.copy()
        
        if hide_zero:
            total_out = res_df['野菜トマト 出数'] + res_df['スパイシートマト 出数'] + res_df['トマト唐麺 出数'] + res_df['夜チーズ 出数']
            display_df = display_df[total_out > 0]
            
        # フォーマット設定（構成比をパーセント表示に調整）
        format_dict = {}
        if cal_file is not None:
            format_dict['野菜トマト 構成比'] = "{:.2%}"
            format_dict['スパイシートマト 構成比'] = "{:.2%}"
            format_dict['トマト唐麺 構成比'] = "{:.2%}"
            
        st.dataframe(display_df.style.format(format_dict), use_container_width=True)
        
        # --- 4. Excelダウンロード機能 ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            res_df.to_excel(writer, sheet_name='出数・構成比集計', index=False)
        processed_data = output.getvalue()
        
        st.download_button(
            label="📥 集計結果（出数表）をExcelでダウンロード",
            data=processed_data,
            file_name="トマト出数_構成比集計表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("集計が完了しました！")

    except Exception as e:
        st.error(f"ファイル解析中にエラーが発生しました: {e}")

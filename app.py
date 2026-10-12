import streamlit as st
import pandas as pd
import io

# ページ基本設定
st.set_page_config(page_title="トマトらーめん出数集計システム", layout="wide")

st.title("🍅 トマトらーめん グループ別出数集計アプリ")
st.write("店舗から出力した「日別メニュー出数.xlsx」をアップロードしてください。")

# ファイルアップローダー
uploaded_file = st.file_uploader("日別メニュー出数ファイル (.xlsx) を選択", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Excelファイルの読み込み
        excel_data = pd.ExcelFile(uploaded_file)
        sheet_name = excel_data.sheet_names[0]  # 1つ目のシートを取得
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        
        # 3行目が日付、4行目が曜日ヘッダー
        header_days = df.iloc[3, 6:].values
        header_weekdays = df.iloc[4, 6:].values
        
        # メニューデータ部分（5行目以降）
        menu_df = df.iloc[5:30].copy()
        day_cols = list(header_days)
        menu_df.columns = ['商品CD', '商品名', '消費税率', '売上金額', '販売数量', '平均販売数'] + day_cols
        
        # 集計グループの定義
        group1_names = ['野菜トマトらーめん', '小・トマトらーめん', '☆野菜トマトらーめん', '☆小・トマトらーめん']
        group2_names = ['チーズまみれトマト', '☆チーズまみれトマト', 'チーズまみれらーめん']
        group3_names = ['トマト唐麺', '☆トマト唐麺']
        group4_names = ['スパイシートマト', '小・スパイシートマト', '☆スパイシートマト', '☆小・スパイシトマト']
        
        def get_group_sum(group_names):
            sub_df = menu_df[menu_df['商品名'].isin(group_names)]
            return sub_df[day_cols].fillna(0).astype(int).sum(axis=0)
        
        g1_sum = get_group_sum(group1_names)
        g2_sum = get_group_sum(group2_names)
        g3_sum = get_group_sum(group3_names)
        g4_sum = get_group_sum(group4_names)
        
        # 日付ラベル（例: "21(金)"）
        day_labels = [f"{str(day).zfill(2)}({wd})" for day, wd in zip(header_days, header_weekdays)]
        
        # 集計テーブル作成
        res_df = pd.DataFrame({
            '日付(曜日)': day_labels,
            '野菜トマト ＋ 小・トマト': g1_sum.values,
            '夜チーズ (チーズまみれ)': g2_sum.values,
            'トマト唐麺': g3_sum.values,
            'スパイシートマト ＋ 小・スパイシートマト': g4_sum.values
        })
        res_df['日別合計'] = res_df.iloc[:, 1:].sum(axis=1)
        
        # 販売数量が全ゼロの日（販売期間外）を自動判定して非表示にするか選べるオプション
        st.subheader("📊 集計結果")
        
        # 表示コントロール
        hide_zero_days = st.checkbox("販売期間外（すべての出数が0の日）を隠す", value=True)
        
        display_df = res_df.copy()
        if hide_zero_days:
            display_df = display_df[display_df['日別合計'] > 0]
            
        # 画面に表を表示
        st.dataframe(display_df, use_container_width=True)
        
        # 合計値の表示（メトリクス表示）
        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("野菜トマト合計", f"{g1_sum.sum()} 食")
        col2.metric("夜チーズ合計", f"{g2_sum.sum()} 食")
        col3.metric("トマト唐麺合計", f"{g3_sum.sum()} 食")
        col4.metric("スパイシートマト合計", f"{g4_sum.sum()} 食")
        col5.metric("全グループ総計", f"{res_df['日別合計'].sum()} 食")
        
        # 集計済みExcelファイルのダウンロードボタンを作成
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            res_df.to_excel(writer, sheet_name='集計結果', index=False)
        processed_data = output.getvalue()
        
        st.download_button(
            label="📥 集計結果をExcelでダウンロード",
            data=processed_data,
            file_name="トマトらーめん集計結果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("集計が完了しました！")
        
    except Exception as e:
        st.error(f"ファイルの解析中にエラーが発生しました: {e}")
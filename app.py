import streamlit as st
import pandas as pd

st.set_page_config(layout="wide")

# 修改標題顏色為更亮的藍色
# 使用 #4169E1 (RoyalBlue) 作為範例，你可以替換成其他十六進制顏色碼
st.markdown("<h1 style='text-align: center; color: #0000A0;'>互動式資料樞紐分析工具</h1>", unsafe_allow_html=True)
# 上傳檔案
uploaded_file = st.file_uploader("請上傳 Excel 檔案", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # 讀取所有 Sheet 名稱
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names

        if not sheet_names:
            st.error("此 Excel 檔案中沒有任何 Sheet。")
            st.stop()  # 如果沒有 Sheet，停止執行

        # 讓用戶選擇 Sheet
        selected_sheet = st.selectbox("請選擇要分析的 Sheet", sheet_names)

        # 讀取選定的 Sheet
        # 關鍵修改：使用 dtype 參數強制 '收益中心' 和 '管理科目' 列為字符串 (如果仍有此問題)
        # 如果沒有此問題，可以移除 dtype 參數
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet,
                           dtype={'收益中心': str, '管理科目': str, '年月': str})  # 再次檢查：如果'年月'也有問題，也加進去
        st.success(f"檔案上傳成功，並已選取 Sheet: **{selected_sheet}**！")

        st.subheader("原始資料預覽")
        st.dataframe(df)

        st.subheader("設定樞紐分析")

        # 獲取所有欄位名稱
        all_columns = df.columns.tolist()

        # ------------------- 篩選器部分開始 -------------------
        st.sidebar.subheader("資料篩選")

        # 讓用戶選擇用於篩選的欄位
        filter_column = st.sidebar.selectbox(
            "選擇要篩選的欄位 (可選)",
            ['無'] + all_columns,  # 添加 '無' 選項表示不篩選
            index=0
        )

        filtered_df = df.copy()  # 建立一個副本進行篩選

        if filter_column != '無':
            # 獲取該欄位的所有唯一值作為篩選選項
            filter_options = sorted(df[filter_column].astype(str).unique().tolist())

            # 讓用戶選擇具體的篩選值
            selected_filter_values = st.sidebar.multiselect(
                f"選擇 {filter_column} 的具體值",
                filter_options,
                default=[]
            )

            if selected_filter_values:
                # 應用篩選
                filtered_df = df[df[filter_column].astype(str).isin(selected_filter_values)].copy()
                st.sidebar.success(f"已應用 {len(selected_filter_values)} 個篩選條件於 {filter_column}。")
            else:
                st.sidebar.info(f"請選擇 {filter_column} 的值來應用篩選。")

        # ------------------- 篩選器部分結束 -------------------

        # 選擇列 (Columns) - 現在對 filtered_df 操作
        st.sidebar.subheader("選擇列 (Columns)")
        selected_columns = st.sidebar.multiselect(
            "選擇要作為列的欄位 (可以選擇多個)",
            all_columns,  # 仍然使用 all_columns 作為選項，但樞紐分析將使用 filtered_df
            default=[]
        )

        # 選擇索引 (Index/Rows) - 現在對 filtered_df 操作
        st.sidebar.subheader("選擇索引 (Rows)")
        selected_index = st.sidebar.multiselect(
            "選擇要作為索引 (列) 的欄位 (可以選擇多個)",
            all_columns,  # 仍然使用 all_columns 作為選項
            default=[]
        )

        # 選擇值 (Values) - 現在對 filtered_df 操作
        st.sidebar.subheader("選擇值 (Values)")
        numeric_columns = filtered_df.select_dtypes(include=['number']).columns.tolist()  # 注意這裡使用 filtered_df
        default_values = numeric_columns[0] if numeric_columns else None
        selected_values = st.sidebar.multiselect(
            "選擇要作為值的欄位 (可以選擇多個)",
            numeric_columns,
            default=[default_values] if default_values else []
        )

        # 選擇聚合函數
        st.sidebar.subheader("選擇聚合函數")
        aggregation_function = st.sidebar.selectbox(
            "選擇聚合函數",
            ['sum', 'mean', 'count', 'min', 'max', 'median', 'std'],
            index=0
        )

        # 執行樞紐分析
        if not selected_index:
            st.info("請至少選擇一個索引 (Rows) 來生成樞紐分析表。")
        elif not selected_values:
            st.info("請至少選擇一個值 (Values) 來生成樞紐分析表。")
        else:
            try:
                agg_func_dict = {}
                for value_col in selected_values:
                    # 這裡檢查的 numeric_columns 應該是基於原始 df 的，因為篩選不改變類型
                    if value_col in df.select_dtypes(include=['number']).columns.tolist():
                        agg_func_dict[value_col] = aggregation_function
                    else:
                        st.warning(f"欄位 '{value_col}' 不是數值型，將不會用於聚合。")

                if not agg_func_dict:
                    st.error("沒有可供聚合的數值型欄位。請檢查 '值' 的選擇。")
                else:
                    pivot_table_df = pd.pivot_table(
                        filtered_df,  # *** 這裡改為使用經過篩選的 DataFrame ***
                        values=selected_values,
                        index=selected_index,
                        columns=selected_columns,
                        aggfunc=agg_func_dict
                    )
                    st.subheader("樞紐分析結果")
                    st.dataframe(pivot_table_df)


                    # 下載樞紐分析結果
                    @st.cache_data
                    def convert_df_to_csv(df_to_convert):
                        # 關鍵修改：指定 encoding 為 'utf-8-sig'
                        return df_to_convert.to_csv(encoding='utf-8-sig').encode('utf-8-sig')


                    csv = convert_df_to_csv(pivot_table_df)

                    st.download_button(
                        label="下載樞紐分析結果為 CSV",
                        data=csv,
                        file_name='pivot_table_result.csv',
                        mime='text/csv',
                    )

            except KeyError as ke:
                st.error(f"欄位錯誤：您選擇的某些欄位可能不存在或名稱有誤。錯誤詳情: {ke}")
                st.info("請檢查您選擇的欄位名稱是否與原始資料表頭完全一致。")
            except ValueError as ve:
                st.error(f"值錯誤：聚合函數可能無法應用於您選擇的資料類型。錯誤詳情: {ve}")
                st.info("請確認作為 '值' 的欄位是數值型資料。")
            except Exception as e:
                st.error(f"執行樞紐分析時發生未預期的錯誤：{e}")
                st.info("請檢查您選擇的欄位和聚合函數是否正確。")

    except Exception as e:
        st.error(f"讀取 Excel 檔案或 Sheet 時發生錯誤：{e}")
        st.info("請確認您上傳的是有效的 Excel 檔案，且選擇的 Sheet 存在。")
else:
    st.info("請上傳一個 Excel 檔案來開始分析。")
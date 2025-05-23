import pandas as pd
import streamlit as st
from datetime import datetime
from pathlib import Path

st.set_page_config(layout="wide")

st.title('RPA reports')
uploaded_file = st.file_uploader('Upload an Excel file', type=['xlsx','xls'])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    # st.write(df)
else:
    st.write('Please upload an Excel file.')


if uploaded_file is not None:
    # ==================== 读取与清洗数据 ====================
    try:
        df = pd.read_excel(uploaded_file)
        
        # 检查必要列
        required_columns = ['分公司', '合同号', '险种名称', '保费']
        if not set(required_columns).issubset(df.columns):
            raise ValueError(f"文件缺少必要列，需包含：{required_columns}")
        
        # 过滤万能型险种
        filtered_df = df[~df['险种名称'].str.contains('万能型', case=False, na=False)]
        
        # ==================== 统计计算（仅按险种） ====================
        # 分组统计（跨分公司去重合同号）
        result = filtered_df.groupby(['险种名称'], as_index=False).agg(
            保单件数=('合同号', 'nunique'),
            保费合计=('保费', 'sum')
        )
        
        # ==================== 排序与汇总 ====================
        # 按保单件数降序排序
        result = result.sort_values(by='保单件数', ascending=False)
        
        # 计算汇总值
        total_policies = result['保单件数'].sum()  # 总保单件数（所有险种去重合同号总和）
        total_premium = result['保费合计'].sum()   # 总保费（所有险种保费总和）
        
        # 创建汇总行
        total_row = pd.DataFrame({
            '险种名称': ['合计'],
            '保单件数': [total_policies],
            '保费合计': [total_premium]
        })
        
        # 合并汇总行（保留原排序，汇总行在最后）
        result = pd.concat([result, total_row], ignore_index=True)
        
        # ==================== 输出文件 ====================
        today = datetime.today().strftime('%Y%m%d')
        output_name = f'分公司险种分布{today}.xlsx'
        desktop_path = Path.home() / 'Desktop'
        output_path = desktop_path / output_name
        
        result.to_excel(output_path, index=False, sheet_name='险种统计')
        
        print(f"处理完成！结果文件已保存至：{output_path}")
        
    except Exception as e:
        print(f"处理失败，错误原因：{str(e)}")
           
     # 展示结果数据
    st.dataframe(data=result.reset_index(drop=True),use_container_width=True,hide_index=True)


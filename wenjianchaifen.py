import streamlit as st
import pandas as pd
import os
import zipfile
from io import BytesIO
import base64

# 初始化session state
if 'split_result' not in st.session_state:
    st.session_state.split_result = None
if 'merge_result' not in st.session_state:
    st.session_state.merge_result = None
if 'original_filename' not in st.session_state:
    st.session_state.original_filename = None

def read_excel_columns(file):
    """读取Excel文件的列名而不加载数据"""
    df = pd.read_excel(file, nrows=0)
    return df.columns.tolist()

def split_excel(file, num_splits, selected_columns=None):
    """将Excel文件拆分为多个子文件"""
    df = pd.read_excel(file)
    
    if selected_columns:
        df = df[selected_columns]
    
    total_rows = len(df)
    rows_per_split = total_rows // num_splits
    remainder = total_rows % num_splits
    
    split_dfs = []
    start_idx = 0
    
    for i in range(num_splits):
        current_rows = rows_per_split + (1 if i < remainder else 0)
        end_idx = start_idx + current_rows
        split_dfs.append(df.iloc[start_idx:end_idx])
        start_idx = end_idx
    
    return split_dfs

def merge_excel(files, selected_columns=None):
    """合并多个Excel文件为一个"""
    dfs = []
    
    for file in files:
        df = pd.read_excel(file)
        if selected_columns:
            df = df[selected_columns]
        dfs.append(df)
    
    merged_df = pd.concat(dfs, ignore_index=True)
    return merged_df

def get_zip_download_link(split_dfs, original_filename):
    """生成包含所有拆分文件的ZIP下载链接"""
    zip_buffer = BytesIO()
    base_name, ext = os.path.splitext(original_filename)
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for i, df in enumerate(split_dfs):
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
            excel_buffer.seek(0)
            zipf.writestr(f"{base_name}——拆分{i+1}of{len(split_dfs)}{ext}", excel_buffer.getvalue())
    
    zip_buffer.seek(0)
    b64 = base64.b64encode(zip_buffer.read()).decode()
    href = f'<a href="data:application/zip;base64,{b64}" download="{base_name}_拆分文件.zip">下载所有拆分文件 (ZIP)</a>'
    return href

def get_excel_download_link(df, original_filename):
    """生成Excel下载链接"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='合并数据')
    output.seek(0)
    
    b64 = base64.b64encode(output.read()).decode()
    base_name, _ = os.path.splitext(original_filename if original_filename else "合并文件")
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{base_name}_合并.xlsx">下载合并后的Excel文件</a>'
    return href

def main():
    st.title("Excel文件拆分与合并工具")
    
    # 选择操作类型
    operation = st.radio("选择操作类型", ["拆分文件", "合并文件"])
    
    if operation == "拆分文件":
        st.header("Excel文件拆分")
        
        # 上传文件
        uploaded_file = st.file_uploader("上传Excel文件", type=["xlsx", "xls"])
        
        if uploaded_file:
            st.session_state.original_filename = uploaded_file.name
            
            # 只读取列名
            all_columns = read_excel_columns(uploaded_file)
            
            # 设置拆分数量
            num_splits = st.number_input("拆分为几个文件", min_value=1, max_value=100, value=2)
            
            # 选择输出列
            selected_columns = st.multiselect(
                "选择要保留的列（未选择的列将被删除）",
                all_columns,
                default=all_columns
            )
            
            if st.button("执行拆分"):
                if not selected_columns:
                    st.error("请至少选择一列")
                else:
                    # 完整读取数据并执行拆分
                    with st.spinner("正在读取数据并拆分文件..."):
                        st.session_state.split_result = split_excel(uploaded_file, num_splits, selected_columns)
                    
                    # 成功提示将在下面显示，确保不会被重新运行清除
            
            # 显示结果（如果有）
            if st.session_state.split_result is not None:
                st.success(f"已成功将文件拆分为 {num_splits} 个部分")
                
                # 显示前几个拆分文件的预览
                for i, split_df in enumerate(st.session_state.split_result[:3]):
                    st.subheader(f"拆分文件 {i+1}/{len(st.session_state.split_result)} 预览")
                    st.dataframe(split_df.head(10))
                
                # 生成ZIP下载链接
                st.markdown(get_zip_download_link(
                    st.session_state.split_result, 
                    st.session_state.original_filename
                ), unsafe_allow_html=True)
    
    else:  # 合并文件
        st.header("Excel文件合并")
        
        # 上传文件
        uploaded_files = st.file_uploader("上传多个Excel文件", type=["xlsx", "xls"], accept_multiple_files=True)
        
        if uploaded_files:
            if uploaded_files:
                all_columns = read_excel_columns(uploaded_files[0])
                st.session_state.original_filename = uploaded_files[0].name if uploaded_files else "合并文件"
            else:
                all_columns = []
            
            # 选择输出列
            selected_columns = st.multiselect(
                "选择要保留的列（未选择的列将被删除）",
                all_columns,
                default=all_columns
            )
            
            if st.button("执行合并"):
                if not selected_columns:
                    st.error("请至少选择一列")
                elif not uploaded_files:
                    st.error("请上传至少一个文件")
                else:
                    # 完整读取数据并执行合并
                    with st.spinner("正在读取数据并合并文件..."):
                        st.session_state.merge_result = merge_excel(uploaded_files, selected_columns)
            
            # 显示结果（如果有）
            if st.session_state.merge_result is not None:
                st.success(f"已成功合并 {len(uploaded_files)} 个文件")
                st.dataframe(st.session_state.merge_result.head(20))
                
                # 生成Excel下载链接
                st.markdown(get_excel_download_link(
                    st.session_state.merge_result, 
                    st.session_state.original_filename
                ), unsafe_allow_html=True)

if __name__ == "__main__":
    main()    

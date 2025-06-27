import streamlit as st
import pandas as pd
import os
import zipfile
from io import BytesIO
import base64

def split_excel(file, num_splits, selected_columns=None):
    """将Excel文件拆分为多个子文件"""
    # 读取文件
    df = pd.read_excel(file)
    
    # 如果指定了输出列，则筛选数据
    if selected_columns:
        df = df[selected_columns]
    
    # 计算每个拆分文件的行数
    total_rows = len(df)
    rows_per_split = total_rows // num_splits
    remainder = total_rows % num_splits
    
    # 准备拆分文件
    split_dfs = []
    start_idx = 0
    
    for i in range(num_splits):
        # 确定当前拆分的行数
        current_rows = rows_per_split + (1 if i < remainder else 0)
        end_idx = start_idx + current_rows
        
        # 提取数据并添加到列表
        split_dfs.append(df.iloc[start_idx:end_idx])
        start_idx = end_idx
    
    return split_dfs

def merge_excel(files, selected_columns=None):
    """合并多个Excel文件为一个"""
    dfs = []
    
    # 读取所有文件
    for file in files:
        df = pd.read_excel(file)
        
        # 如果指定了输出列，则筛选数据
        if selected_columns:
            df = df[selected_columns]
            
        dfs.append(df)
    
    # 合并所有DataFrame
    merged_df = pd.concat(dfs, ignore_index=True)
    return merged_df

def get_table_download_link(df, original_filename, part_num, total_parts):
    """生成Excel下载链接"""
    # 确定新文件名
    base_name, ext = os.path.splitext(original_filename)
    new_filename = f"{base_name}——拆分{part_num+1}of{total_parts}{ext}"
    
    # 创建Excel文件
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    output.seek(0)
    
    # 生成下载链接
    b64 = base64.b64encode(output.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{new_filename}">下载拆分文件 {part_num+1}/{total_parts}</a>'
    return href

def get_zip_download_link(files_dict, original_filename):
    """生成ZIP文件下载链接"""
    # 创建ZIP文件
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for filename, content in files_dict.items():
            zipf.writestr(filename, content)
    
    # 生成下载链接
    zip_buffer.seek(0)
    b64 = base64.b64encode(zip_buffer.read()).decode()
    href = f'<a href="data:application/zip;base64,{b64}" download="{original_filename}_合并.zip">下载合并文件</a>'
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
            # 读取文件以获取列名
            df = pd.read_excel(uploaded_file)
            all_columns = df.columns.tolist()
            
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
                    # 执行拆分
                    split_dfs = split_excel(uploaded_file, num_splits, selected_columns)
                    
                    # 显示结果并提供下载链接
                    st.success(f"已成功将文件拆分为 {num_splits} 个部分")
                    
                    # 生成并保存所有拆分文件
                    for i, split_df in enumerate(split_dfs):
                        st.subheader(f"拆分文件 {i+1}/{num_splits}")
                        st.dataframe(split_df.head(10))  # 显示前10行
                        st.markdown(get_table_download_link(split_df, uploaded_file.name, i, num_splits), unsafe_allow_html=True)
    
    else:  # 合并文件
        st.header("Excel文件合并")
        
        # 上传文件
        uploaded_files = st.file_uploader("上传多个Excel文件", type=["xlsx", "xls"], accept_multiple_files=True)
        
        if uploaded_files:
            # 读取第一个文件以获取列名
            df = pd.read_excel(uploaded_files[0])
            all_columns = df.columns.tolist()
            
            # 选择输出列
            selected_columns = st.multiselect(
                "选择要保留的列（未选择的列将被删除）",
                all_columns,
                default=all_columns
            )
            
            if st.button("执行合并"):
                if not selected_columns:
                    st.error("请至少选择一列")
                else:
                    # 执行合并
                    merged_df = merge_excel(uploaded_files, selected_columns)
                    
                    # 显示结果并提供下载链接
                    st.success(f"已成功合并 {len(uploaded_files)} 个文件")
                    st.dataframe(merged_df.head(20))  # 显示前20行
                    
                    # 创建单个Excel文件
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        merged_df.to_excel(writer, index=False, sheet_name='合并数据')
                    output.seek(0)
                    
                    # 生成下载链接
                    b64 = base64.b64encode(output.read()).decode()
                    original_filename = "合并文件" if len(uploaded_files) == 0 else os.path.splitext(uploaded_files[0].name)[0]
                    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{original_filename}_合并.xlsx">下载合并后的Excel文件</a>'
                    st.markdown(href, unsafe_allow_html=True)
                    
                    # 可选：如果需要ZIP格式（例如包含多个工作表）
                    # files_dict = {f"合并数据.xlsx": output.getvalue()}
                    # st.markdown(get_zip_download_link(files_dict, original_filename), unsafe_allow_html=True)

if __name__ == "__main__":
    main()    

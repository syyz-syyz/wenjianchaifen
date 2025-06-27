import streamlit as st
import pandas as pd
import os
from datetime import datetime
import zipfile
from io import BytesIO

def split_excel(file, num_splits):
    """将 Excel 文件拆分为指定数量的子文件"""
    try:
        # 读取 Excel 文件
        df = pd.read_excel(file)
        
        # 计算每个子文件的行数
        total_rows = len(df)
        rows_per_file = total_rows // num_splits
        
        # 处理余数
        remainder = total_rows % num_splits
        
        # 获取原始文件名（不带扩展名）
        original_filename = os.path.splitext(file.name)[0]
        
        # 创建保存拆分文件的目录
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = f"split_files_{timestamp}"
        os.makedirs(output_dir, exist_ok=True)
        
        # 存储拆分后的文件路径
        output_files = []
        
        start_idx = 0
        for i in range(num_splits):
            # 确定当前文件的行数
            current_rows = rows_per_file + (1 if i < remainder else 0)
            
            # 提取数据
            end_idx = start_idx + current_rows
            sub_df = df.iloc[start_idx:end_idx]
            
            # 保存子文件
            output_filename = f"{output_dir}/{original_filename}——拆分{i+1}.xlsx"
            sub_df.to_excel(output_filename, index=False)
            output_files.append(output_filename)
            
            # 更新起始索引
            start_idx = end_idx
        
        return output_files
    
    except Exception as e:
        st.error(f"处理文件时出错: {str(e)}")
        return []

def merge_excel(files):
    """合并多个 Excel 文件为一个"""
    try:
        if not files:
            st.error("请上传至少一个 Excel 文件")
            return None
        
        # 创建保存合并文件的目录
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = f"merged_files_{timestamp}"
        os.makedirs(output_dir, exist_ok=True)
        
        # 读取并合并所有文件
        dfs = []
        for file in files:
            df = pd.read_excel(file)
            dfs.append(df)
        
        # 合并数据框
        merged_df = pd.concat(dfs, ignore_index=True)
        
        # 获取原始文件名（不带扩展名）
        if len(files) == 1:
            original_filename = os.path.splitext(files[0].name)[0]
            output_filename = f"{output_dir}/{original_filename}——合并.xlsx"
        else:
            output_filename = f"{output_dir}/合并文件_{timestamp}.xlsx"
        
        # 保存合并后的文件
        merged_df.to_excel(output_filename, index=False)
        
        return output_filename
    
    except Exception as e:
        st.error(f"合并文件时出错: {str(e)}")
        return None

def main():
    st.title("Excel 文件处理工具")
    
    # 选择操作类型
    operation = st.radio("选择操作类型", ["文件拆分", "文件合并"])
    
    if operation == "文件拆分":
        st.subheader("Excel 文件拆分")
        
        # 上传文件
        uploaded_file = st.file_uploader("选择一个 Excel 文件", type=["xlsx", "xls"])
        
        if uploaded_file is not None:
            # 显示文件信息
            file_details = {"文件名": uploaded_file.name, "文件大小": uploaded_file.size}
            st.write(file_details)
            
            # 获取拆分数量
            num_splits = st.number_input("请输入要拆分的文件数量", min_value=1, max_value=100, value=2)
            
            # 拆分按钮
            if st.button("开始拆分"):
                with st.spinner("正在处理文件..."):
                    output_files = split_excel(uploaded_file, num_splits)
                    
                    if output_files:
                        st.success(f"成功将文件拆分为 {len(output_files)} 个子文件！")
                        
                        # 创建 ZIP 文件以便批量下载
                        zip_buffer = BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w') as zipf:
                            for file_path in output_files:
                                zipf.write(file_path, os.path.basename(file_path))
                        
                        # 定位到 ZIP 文件的开始
                        zip_buffer.seek(0)
                        
                        # 显示下载链接
                        st.download_button(
                            label="下载所有拆分文件 (ZIP)",
                            data=zip_buffer,
                            file_name="拆分文件合集.zip",
                            mime="application/zip"
                        )
                        
                        # 单独文件下载选项
                        st.subheader("或单独下载拆分后的文件")
                        for file_path in output_files:
                            with open(file_path, "rb") as f:
                                file_bytes = f.read()
                                file_name = os.path.basename(file_path)
                                st.download_button(
                                    label=f"下载 {file_name}",
                                    data=file_bytes,
                                    file_name=file_name,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                        
                        # 清理临时文件
                        if st.button("清理临时文件"):
                            for file_path in output_files:
                                os.remove(file_path)
                            os.rmdir(os.path.dirname(output_files[0]))
                            st.info("临时文件已清理完毕。")
    
    else:  # 文件合并
        st.subheader("Excel 文件合并")
        
        # 上传多个文件
        uploaded_files = st.file_uploader("选择多个 Excel 文件", type=["xlsx", "xls"], accept_multiple_files=True)
        
        if uploaded_files:
            # 显示上传的文件列表
            st.write(f"已上传 {len(uploaded_files)} 个文件:")
            for file in uploaded_files:
                st.write(f"- {file.name}")
            
            # 合并按钮
            if st.button("开始合并"):
                with st.spinner("正在合并文件..."):
                    output_file = merge_excel(uploaded_files)
                    
                    if output_file:
                        st.success("文件合并成功！")
                        
                        # 显示下载链接
                        with open(output_file, "rb") as f:
                            file_bytes = f.read()
                            file_name = os.path.basename(output_file)
                            st.download_button(
                                label=f"下载合并后的文件",
                                data=file_bytes,
                                file_name=file_name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        
                        # 清理临时文件
                        if st.button("清理临时文件"):
                            os.remove(output_file)
                            os.rmdir(os.path.dirname(output_file))
                            st.info("临时文件已清理完毕。")

if __name__ == "__main__":
    main()   

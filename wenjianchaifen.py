import streamlit as st
import pandas as pd
import os
from datetime import datetime

def split_excel(file, num_splits):
    """将 Excel 文件拆分为指定数量的子文件（支持大文件）"""
    try:
        # 读取 Excel 文件
        xls = pd.ExcelFile(file)
        
        # 获取所有表名
        sheet_names = xls.sheet_names
        
        # 获取原始文件名（不带扩展名）
        original_filename = os.path.splitext(file.name)[0]
        
        # 用于存储每个拆分文件的缓冲区和文件名
        output_files = []
        
        # 处理每个表
        for sheet_name in sheet_names:
            # 获取当前表的数据
            df = xls.parse(sheet_name)
            
            # 计算每个子文件的行数
            total_rows = len(df)
            rows_per_file = total_rows // num_splits
            remainder = total_rows % num_splits
            
            start_idx = 0
            for i in range(num_splits):
                # 确定当前文件的行数
                current_rows = rows_per_file + (1 if i < remainder else 0)
                
                # 提取数据
                end_idx = start_idx + current_rows
                sub_df = df.iloc[start_idx:end_idx]
                
                # 创建内存中的 Excel 文件
                excel_buffer = st.spinner()
                with excel_buffer:
                    excel_buffer = st.spinner()
                    with excel_buffer:
                        excel_buffer = BytesIO()
                        if len(sheet_names) == 1:
                            sub_df.to_excel(excel_buffer, index=False)
                            file_name = f"{original_filename}——拆分{i+1}.xlsx"
                        else:
                            sub_df.to_excel(excel_buffer, sheet_name=sheet_name, index=False)
                            file_name = f"{original_filename}——拆分{i+1}_{sheet_name}.xlsx"
                        
                        # 定位到 Excel 文件的开始
                        excel_buffer.seek(0)
                        
                        # 添加到输出文件列表
                        output_files.append((excel_buffer, file_name))
                
                # 更新起始索引
                start_idx = end_idx
        
        return output_files
    
    except Exception as e:
        st.error(f"处理文件时出错: {str(e)}")
        return []

def merge_excel(files):
    """合并多个 Excel 文件为一个（支持大文件）"""
    try:
        if not files:
            st.error("请上传至少一个 Excel 文件")
            return None
        
        # 创建内存中的 Excel 文件
        output_buffer = BytesIO()
        
        # 创建 ExcelWriter 对象
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            # 处理每个上传的文件
            for file in files:
                xls = pd.ExcelFile(file)
                sheet_names = xls.sheet_names
                
                # 获取原始文件名（不带扩展名）
                original_filename = os.path.splitext(file.name)[0]
                
                # 处理每个表
                for sheet_name in sheet_names:
                    # 读取数据
                    df = xls.parse(sheet_name)
                    
                    # 确定工作表名称
                    if len(sheet_names) == 1:
                        # 如果文件中只有一个表，使用文件名作为表名
                        sheet_name_out = original_filename
                    else:
                        # 如果文件中有多个表，使用原表名
                        sheet_name_out = f"{original_filename}_{sheet_name}"
                    
                    # 将数据写入 ExcelWriter
                    df.to_excel(writer, sheet_name=sheet_name_out, index=False)
        
        # 定位到 Excel 文件的开始
        output_buffer.seek(0)
        
        return output_buffer
    
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
                        st.success(f"成功将文件拆分为 {num_splits} 个子文件！")
                        
                        # 显示下载链接
                        for buffer, file_name in output_files:
                            st.download_button(
                                label=f"下载 {file_name}",
                                data=buffer,
                                file_name=file_name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
    
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
                    output_excel = merge_excel(uploaded_files)
                    
                    if output_excel:
                        st.success("文件合并成功！")
                        
                        # 获取原始文件名（不带扩展名）
                        if len(uploaded_files) == 1:
                            original_filename = os.path.splitext(uploaded_files[0].name)[0]
                            file_name = f"{original_filename}——合并.xlsx"
                        else:
                            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                            file_name = f"合并文件_{timestamp}.xlsx"
                        
                        # 显示下载链接
                        st.download_button(
                            label=f"下载合并后的文件",
                            data=output_excel,
                            file_name=file_name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

if __name__ == "__main__":
    main()    

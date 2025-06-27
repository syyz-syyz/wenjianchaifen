import streamlit as st
import pandas as pd
from datetime import datetime
import zipfile
from io import BytesIO

def split_excel(file, num_splits):
    """将 Excel 文件拆分为指定数量的子文件（纯内存处理）"""
    try:
        df = pd.read_excel(file)
        total_rows = len(df)
        rows_per_file = total_rows // num_splits
        remainder = total_rows % num_splits
        original_filename = file.name.split('.')[0]
        
        output_files = []
        start_idx = 0
        for i in range(num_splits):
            current_rows = rows_per_file + (1 if i < remainder else 0)
            end_idx = start_idx + current_rows
            sub_df = df.iloc[start_idx:end_idx]
            
            buffer = BytesIO()
            sub_df.to_excel(buffer, index=False)
            buffer.seek(0)
            
            file_name = f"{original_filename}——拆分{i+1}.xlsx"
            output_files.append((file_name, buffer))
            start_idx = end_idx
        
        return output_files
    
    except Exception as e:
        st.error(f"拆分错误: {str(e)}")
        return []

def merge_excel(files):
    """合并多个 Excel 文件（纯内存处理）"""
    try:
        if not files:
            st.error("请上传至少一个文件")
            return None
        
        dfs = []
        for file in files:
            df = pd.read_excel(file)
            dfs.append(df)
        
        merged_df = pd.concat(dfs, ignore_index=True)
        
        if len(files) == 1:
            original_filename = files[0].name.split('.')[0]
            output_name = f"{original_filename}——合并.xlsx"
        else:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_name = f"合并文件_{timestamp}.xlsx"
        
        buffer = BytesIO()
        merged_df.to_excel(buffer, index=False)
        buffer.seek(0)
        
        return (output_name, buffer)
    
    except Exception as e:
        st.error(f"合并错误: {str(e)}")
        return None

def main():
    st.title("Excel 文件处理工具")
    operation = st.radio("操作类型", ["拆分", "合并"])
    
    # ------------------- 拆分功能 -------------------
    if operation == "拆分":
        st.subheader("Excel 文件拆分")
        file = st.file_uploader("上传文件", type=["xlsx", "xls"])
        
        if file:
            st.info(f"已上传: {file.name}")
            splits = st.number_input("拆分数量", 1, 100, 2)
            
            if st.button("开始拆分"):
                with st.spinner("处理中..."):
                    result = split_excel(file, splits)
                    if result:
                        st.success(f"成功拆分出 {len(result)} 个文件！")
                        
                        # 打包ZIP下载
                        zip_buffer = BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w') as zipf:
                            for name, buf in result:
                                zipf.writestr(name, buf.getvalue())
                        zip_buffer.seek(0)
                        st.download_button(
                            "下载全部拆分文件", 
                            zip_buffer, 
                            "拆分文件合集.zip", 
                            "application/zip"
                        )
                        
                        # 单独下载
                        st.subheader("单独下载")
                        for name, buf in result:
                            st.download_button(
                                f"下载 {name}", 
                                buf, 
                                name, 
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
    
    # ------------------- 合并功能 -------------------
    else:
        st.subheader("Excel 文件合并")
        files = st.file_uploader("上传多个文件", type=["xlsx", "xls"], accept_multiple_files=True)
        
        if files:
            st.info(f"已上传 {len(files)} 个文件")
            if st.button("开始合并"):
                with st.spinner("合并中..."):
                    result = merge_excel(files)
                    if result:
                        name, buf = result
                        st.success("合并完成！")
                        st.download_button(
                            "下载合并文件", 
                            buf, 
                            name, 
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

if __name__ == "__main__":
    main()

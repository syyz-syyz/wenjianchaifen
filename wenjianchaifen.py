import streamlit as st
import pandas as pd
import os
from io import BytesIO

def split_excel(file, num_splits):
    """将Excel文件拆分为指定数量的子文件"""
    try:
        # 读取Excel文件
        df = pd.read_excel(file)
        file_name = os.path.splitext(file.name)[0]
        
        # 计算每个子文件的行数
        total_rows = len(df)
        rows_per_split = total_rows // num_splits
        remainder = total_rows % num_splits
        
        # 准备输出文件列表
        output_files = []
        
        # 拆分文件
        start_idx = 0
        for i in range(num_splits):
            # 确定当前子文件的行数
            current_rows = rows_per_split + (1 if i < remainder else 0)
            end_idx = start_idx + current_rows
            
            # 获取子数据框
            sub_df = df.iloc[start_idx:end_idx]
            
            # 创建新的Excel文件
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                sub_df.to_excel(writer, index=False)
            output.seek(0)
            
            # 生成输出文件名
            output_name = f"{file_name}——拆分{i+1}.xlsx"
            output_files.append((output_name, output))
            
            # 更新起始索引
            start_idx = end_idx
        
        return output_files
    
    except Exception as e:
        st.error(f"拆分过程中出错: {str(e)}")
        return []

def merge_excel(files):
    """合并多个Excel文件为一个"""
    try:
        if not files:
            st.warning("请上传至少一个Excel文件")
            return None
        
        # 读取所有文件
        dfs = []
        original_file_name = None
        
        for file in files:
            try:
                df = pd.read_excel(file)
                dfs.append(df)
                
                # 尝试提取原始文件名（如果是拆分文件）
                file_name = os.path.splitext(file.name)[0]
                if "——拆分" in file_name:
                    original_name = file_name.split("——拆分")[0]
                    if original_file_name is None:
                        original_file_name = original_name
                    elif original_name != original_file_name:
                        st.warning("检测到不同来源的文件，将使用第一个文件名作为基础")
            except Exception as e:
                st.error(f"读取文件 {file.name} 时出错: {str(e)}")
        
        if not dfs:
            st.error("无法读取任何文件")
            return None
        
        # 合并数据框
        merged_df = pd.concat(dfs, ignore_index=True)
        
        # 创建合并后的Excel文件
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            merged_df.to_excel(writer, index=False)
        output.seek(0)
        
        # 确定输出文件名
        if original_file_name:
            output_name = f"{original_file_name}——合并.xlsx"
        else:
            output_name = "合并文件.xlsx"
        
        return output_name, output
    
    except Exception as e:
        st.error(f"合并过程中出错: {str(e)}")
        return None

def main():
    """主函数：Streamlit应用入口"""
    st.title("Excel文件拆分与合并工具")
    
    # 创建选项卡
    tab1, tab2 = st.tabs(["拆分文件", "合并文件"])
    
    with tab1:
        st.header("Excel文件拆分")
        file = st.file_uploader("上传Excel文件", type=["xlsx", "xls"])
        
        if file:
            num_splits = st.number_input("拆分为几个文件", min_value=1, max_value=100, value=2, step=1)
            
            if st.button("开始拆分"):
                with st.spinner("正在处理..."):
                    output_files = split_excel(file, num_splits)
                    
                    if output_files:
                        st.success(f"成功将文件拆分为 {len(output_files)} 个子文件")
                        
                        # 显示下载链接
                        for name, data in output_files:
                            st.download_button(
                                label=f"下载 {name}",
                                data=data,
                                file_name=name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
    
    with tab2:
        st.header("Excel文件合并")
        files = st.file_uploader("上传多个Excel文件", type=["xlsx", "xls"], accept_multiple_files=True)
        
        if files:
            if st.button("开始合并"):
                with st.spinner("正在处理..."):
                    merged_file = merge_excel(files)
                    
                    if merged_file:
                        st.success("文件合并成功")
                        name, data = merged_file
                        st.download_button(
                            label=f"下载 {name}",
                            data=data,
                            file_name=name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

if __name__ == "__main__":
    main()    

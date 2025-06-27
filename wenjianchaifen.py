import streamlit as st
import pandas as pd
import os
import zipfile
from io import BytesIO
import base64
import gc

# 初始化session state
if 'split_result' not in st.session_state:
    st.session_state.split_result = None
if 'merge_result' not in st.session_state:
    st.session_state.merge_result = None
if 'original_filename' not in st.session_state:
    st.session_state.original_filename = None
if 'uploaded_file_content' not in st.session_state:
    st.session_state.uploaded_file_content = None
if 'uploaded_files_content' not in st.session_state:
    st.session_state.uploaded_files_content = None
if 'data_chunks' not in st.session_state:
    st.session_state.data_chunks = {}
if 'total_rows' not in st.session_state:
    st.session_state.total_rows = 0
if 'file_size_warning' not in st.session_state:
    st.session_state.file_size_warning = False

def read_excel_columns(file_content):
    """读取Excel文件的列名而不加载数据"""
    df = pd.read_excel(file_content, nrows=0)
    return df.columns.tolist()

def split_excel_by_dict(file_content, num_splits, selected_columns=None, chunk_size=10000):
    """将Excel文件按块读取到字典中，并按指定规则拆分"""
    # 先获取总行列数判断文件大小
    df_info = pd.ExcelFile(file_content).parse(nrows=0)
    total_cols = df_info.shape[1]
    
    # 检查文件大小
    file_size = len(file_content.getvalue()) / (1024 * 1024)  # MB
    if file_size > 500:
        st.session_state.file_size_warning = True
        st.warning(f"检测到大型文件（{file_size:.2f}MB），正在分块处理以降低内存消耗...")
    
    # 清空现有数据块缓存
    st.session_state.data_chunks = {}
    st.session_state.total_rows = 0
    
    # 分块读取数据到字典
    excel_file = pd.ExcelFile(file_content)
    chunk_idx = 0
    
    for chunk in excel_file.parse(chunksize=chunk_size):
        if selected_columns:
            chunk = chunk[selected_columns]
        
        st.session_state.data_chunks[chunk_idx] = chunk
        st.session_state.total_rows += len(chunk)
        chunk_idx += 1
        
        # 释放内存
        del chunk
        gc.collect()
    
    # 计算每个拆分文件的行数
    rows_per_split = st.session_state.total_rows // num_splits
    remainder = st.session_state.total_rows % num_splits
    
    # 准备拆分计划
    split_plans = []
    current_row = 0
    
    for i in range(num_splits):
        target_rows = rows_per_split + (1 if i < remainder else 0)
        split_plans.append({
            'start_row': current_row,
            'end_row': current_row + target_rows,
            'chunks_needed': []  # 记录需要的块索引和范围
        })
        current_row += target_rows
    
    # 确定每个拆分需要哪些数据块
    for plan in split_plans:
        start_row = plan['start_row']
        end_row = plan['end_row']
        current_pos = 0
        
        for chunk_idx, chunk in st.session_state.data_chunks.items():
            chunk_rows = len(chunk)
            chunk_start = current_pos
            chunk_end = current_pos + chunk_rows
            
            # 检查当前块是否与拆分范围有交集
            if chunk_end > start_row and chunk_start < end_row:
                # 计算交集范围
                overlap_start = max(start_row, chunk_start) - chunk_start
                overlap_end = min(end_row, chunk_end) - chunk_start
                
                plan['chunks_needed'].append({
                    'chunk_idx': chunk_idx,
                    'start': overlap_start,
                    'end': overlap_end
                })
            
            current_pos += chunk_rows
    
    return split_plans

def merge_excel(files_content, selected_columns=None, chunk_size=10000):
    """分块合并多个Excel文件，降低内存占用"""
    merged_chunks = []
    
    # 遍历每个文件
    for file_idx, file in enumerate(files_content):
        st.text(f"正在处理文件 {file_idx+1}/{len(files_content)}...")
        excel_file = pd.ExcelFile(file)
        
        # 分块读取当前文件
        for chunk in excel_file.parse(chunksize=chunk_size):
            if selected_columns:
                chunk = chunk[selected_columns]
            merged_chunks.append(chunk)
            
            # 定期合并块以释放内存
            if len(merged_chunks) >= 5:
                merged_chunks = [pd.concat(merged_chunks, ignore_index=True)]
        
        # 释放内存
        del excel_file, chunk
        gc.collect()
    
    # 合并所有块
    if merged_chunks:
        merged_df = pd.concat(merged_chunks, ignore_index=True)
        return merged_df
    return pd.DataFrame()

def get_zip_download_link(split_plans, original_filename):
    """根据拆分计划生成ZIP下载链接"""
    zip_buffer = BytesIO()
    base_name, ext = os.path.splitext(original_filename)
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for i, plan in enumerate(split_plans):
            # 创建新的DataFrame收集数据
            split_df_parts = []
            
            for chunk_info in plan['chunks_needed']:
                chunk = st.session_state.data_chunks[chunk_info['chunk_idx']]
                split_df_parts.append(chunk.iloc[chunk_info['start']:chunk_info['end']])
            
            # 合并所有部分
            if split_df_parts:
                split_df = pd.concat(split_df_parts, ignore_index=True)
                
                # 写入Excel
                excel_buffer = BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    split_df.to_excel(writer, index=False, sheet_name='Sheet1')
                excel_buffer.seek(0)
                
                # 添加到ZIP
                zipf.writestr(f"{base_name}——拆分{i+1}of{len(split_plans)}{ext}", excel_buffer.getvalue())
    
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
    st.title("Excel文件拆分与合并工具（字典缓存版）")
    
    # 显示内存优化提示
    if st.session_state.file_size_warning:
        st.info("当前文件较大，系统已启用分块处理模式，可能需要更长时间，请耐心等待...")
    
    # 选择操作类型
    operation = st.radio("选择操作类型", ["拆分文件", "合并文件"])
    
    if operation == "拆分文件":
        st.header("Excel文件拆分")
        
        # 上传文件并保存到session state
        uploaded_file = st.file_uploader("上传Excel文件", type=["xlsx", "xls"])
        if uploaded_file is not None:
            st.session_state.uploaded_file_content = uploaded_file
            st.session_state.original_filename = uploaded_file.name
            st.session_state.file_size_warning = False  # 重置警告状态
        
        # 检查是否有上传的文件内容
        if st.session_state.uploaded_file_content is not None:
            # 只读取列名
            all_columns = read_excel_columns(st.session_state.uploaded_file_content)
            
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
                    # 重置结果以避免旧数据干扰
                    st.session_state.split_result = None
                    
                    # 按块读取数据到字典并生成拆分计划
                    with st.spinner("正在分块读取数据并生成拆分计划..."):
                        try:
                            split_plans = split_excel_by_dict(
                                st.session_state.uploaded_file_content, 
                                num_splits, 
                                selected_columns
                            )
                            st.session_state.split_result = split_plans
                            
                            # 显示基本信息
                            st.success(f"已成功规划拆分方案: {num_splits} 个文件，共 {st.session_state.total_rows} 行数据")
                        except Exception as e:
                            st.error(f"处理过程中出错: {str(e)}")
                            st.error("请尝试减少拆分数量或选择更少的列，或使用更小的文件。")
            
            # 显示结果（如果有）
            if st.session_state.split_result is not None:
                st.success(f"已成功规划拆分方案: {len(st.session_state.split_result)} 个文件")
                
                # 显示拆分概要
                st.subheader("拆分概要")
                for i, plan in enumerate(st.session_state.split_result):
                    st.text(f"文件 {i+1}: 行 {plan['start_row']+1}-{plan['end_row']}")
                
                # 生成ZIP下载链接
                st.markdown(get_zip_download_link(
                    st.session_state.split_result, 
                    st.session_state.original_filename
                ), unsafe_allow_html=True)
    
    else:  # 合并文件
        st.header("Excel文件合并")
        
        # 上传文件并保存到session state
        uploaded_files = st.file_uploader("上传多个Excel文件", type=["xlsx", "xls"], accept_multiple_files=True)
        if uploaded_files is not None and len(uploaded_files) > 0:
            st.session_state.uploaded_files_content = uploaded_files
            st.session_state.original_filename = uploaded_files[0].name if uploaded_files else "合并文件"
            st.session_state.file_size_warning = False  # 重置警告状态
        
        # 检查是否有上传的文件内容
        if st.session_state.uploaded_files_content is not None and len(st.session_state.uploaded_files_content) > 0:
            # 读取第一个文件的列名
            all_columns = read_excel_columns(st.session_state.uploaded_files_content[0])
            
            # 选择输出列
            selected_columns = st.multiselect(
                "选择要保留的列（未选择的列将被删除）",
                all_columns,
                default=all_columns
            )
            
            if st.button("执行合并"):
                if not selected_columns:
                    st.error("请至少选择一列")
                elif not st.session_state.uploaded_files_content:
                    st.error("请上传至少一个文件")
                else:
                    # 重置结果以避免旧数据干扰
                    st.session_state.merge_result = None
                    
                    # 完整读取数据并执行合并（分块处理）
                    with st.spinner("正在分块合并文件...这可能需要一些时间..."):
                        try:
                            st.session_state.merge_result = merge_excel(
                                st.session_state.uploaded_files_content, 
                                selected_columns
                            )
                            if st.session_state.merge_result is not None and not st.session_state.merge_result.empty:
                                st.success(f"已成功合并 {len(st.session_state.uploaded_files_content)} 个文件")
                            else:
                                st.error("合并结果为空，请检查文件内容")
                        except Exception as e:
                            st.error(f"合并过程中出错: {str(e)}")
                            st.error("请尝试减少合并文件数量或选择更少的列，或使用更小的文件。")
            
            # 显示结果（如果有）
            if st.session_state.merge_result is not None and not st.session_state.merge_result.empty:
                st.success(f"已成功合并 {len(st.session_state.uploaded_files_content)} 个文件")
                
                # 显示合并结果预览
                st.dataframe(st.session_state.merge_result.head(20))
                
                # 生成Excel下载链接
                st.markdown(get_excel_download_link(
                    st.session_state.merge_result, 
                    st.session_state.original_filename
                ), unsafe_allow_html=True)
    
    # 手动触发垃圾回收
    gc.collect()

if __name__ == "__main__":
    main()

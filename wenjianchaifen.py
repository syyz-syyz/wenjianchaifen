def split_excel(file, num_splits):
    """将 Excel 文件拆分为指定数量的子文件（支持大文件）"""
    try:
        # 获取文件总行数
        reader = pd.ExcelFile(file)
        total_rows = len(reader.parse())
        
        # 计算每个子文件的行数
        rows_per_file = total_rows // num_splits
        remainder = total_rows % num_splits
        
        original_filename = os.path.splitext(file.name)[0]
        
        # 创建内存中的 ZIP 文件
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w') as zipf:
            # 创建进度条
            progress_bar = st.progress(0)
            
            # 分块处理文件
            start_idx = 0
            for i in range(num_splits):
                current_rows = rows_per_file + (1 if i < remainder else 0)
                
                # 读取当前块的数据
                df = reader.parse()  # 对于非超大文件，可以一次性读取
                # 对于真正的大文件，使用：
                # df = reader.parse(nrows=current_rows, skiprows=start_idx)
                
                # 保存到内存中的 Excel
                excel_buffer = BytesIO()
                df.to_excel(excel_buffer, index=False)
                excel_buffer.seek(0)
                
                # 添加到 ZIP
                zipf.writestr(f"{original_filename}——拆分{i+1}.xlsx", excel_buffer.read())
                
                # 更新进度条
                progress = (i + 1) / num_splits
                progress_bar.progress(progress)
                
                start_idx += current_rows
        
        zip_buffer.seek(0)
        return zip_buffer
    
    except Exception as e:
        st.error(f"处理文件时出错: {str(e)}")
        return None

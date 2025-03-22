import os
import pandas as pd
import re

def clean_title(title):
    """清洗标题，去除非英文字符和方括号中的内容"""
    if not isinstance(title, str):
        return title
    # 去除不可见字符
    title = ''.join(ch for ch in title if ch.isprintable())
    # 去除方括号及其中内容
    title = re.sub(r'\\[.*?\\]', '', title)
    # 去除非英文字符
    title = re.sub(r'[^a-zA-Z\\s]', '', title)
    # 标准化大小写（转换为小写）
    title = title.lower()
    # 去除多余空格
    title = re.sub(r'\\s+', ' ', title).strip()
    return title

def merge_and_clean_excel_files(folder_path, output_file):
    """合并文件夹中的 Excel 文件并去重"""
    all_data = []

    # 遍历文件夹中的所有 Excel 文件
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
            file_path = os.path.join(folder_path, file_name)
            try:
                # 读取 Excel 文件
                df = pd.read_excel(file_path)
                # 检查是否为空 DataFrame
                if df.empty:
                    print(f"Skipping empty file: {file_name}")
                    continue
                # 规范列名
                df.columns = [col.lower() for col in df.columns]
                # 仅保留指定列
                df = df[[col for col in df.columns if col in ['title', 'authors', 'author', 'year']]]
                # 统一列名
                df.rename(columns={
                    'authors': 'author',
                    'author': 'author',
                    'year': 'year',
                    'title': 'title'
                }, inplace=True)
                all_data.append(df)
            except Exception as e:
                print(f"Error processing file {file_name}: {e}")

    # 合并所有数据
    if not all_data:
        print("No valid Excel files found in the folder.")
        return

    merged_data = pd.concat(all_data, ignore_index=True)

    # 创建清洗后的标题列
    merged_data['cleaned_title'] = merged_data['title'].apply(clean_title)

    # 去重，保留原始标题
    merged_data = merged_data.drop_duplicates(subset=['cleaned_title'])

    # 删除清洗后的标题列
    merged_data.drop(columns=['cleaned_title'], inplace=True)

    # 保存到新的 Excel 文件
    merged_data.to_excel(output_file, index=False)
    print(f"Merged and cleaned data has been saved to {output_file}")

folder_path = r"."
output_file = r".\Merged.xlsx"  # 输出

merge_and_clean_excel_files(folder_path, output_file)

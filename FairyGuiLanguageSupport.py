import pandas as pd
import xml.etree.ElementTree as ET
import os

# 使用相对路径
current_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(current_dir, 'UILanguage.xlsx')

# Read the Excel file into a DataFrame
df = pd.read_excel(file_path)

# 获取列名作为语言标识
languages = df.columns[1:]  # 跳过第一列（键名列）
key_column = df.iloc[:, 0]  # 获取第一列作为键名

# 为每种语言生成对应的XML文件
for col_idx, language in enumerate(languages, start=1):
    # 创建输出文件名
    output_file = os.path.join(current_dir, f'UILanguage_{language}.xml')
    
    # 获取当前语言的值
    value_column_data = df.iloc[:, col_idx]
    
    # 创建XML内容
    xml_content = '<?xml version="1.0" encoding="utf-8"?>\n<resources>\n'
    
    # 遍历数据并添加到XML中
    for index, key in key_column.items():
        value = value_column_data[index]
        # 确保值不是NaN且键不为空
        if pd.notna(value) and pd.notna(key) and str(key).strip():
            # 将键名作为属性名
            xml_content += f'  <string {str(key)}>{str(value)}</string>\n'
    
    xml_content += '</resources>'
    
    # 保存XML文件
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(xml_content)
    
    print(f"已生成语言文件: {output_file}")

print("所有语言文件生成完成！")
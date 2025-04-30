from openpyxl import load_workbook, Workbook
import lxml.etree as etree
import pandas as pd
from openpyxl.styles import PatternFill
import os

"""
核心功能模块: 包含所有XML和Excel处理功能
"""

class XMLProcessor:
    """XML处理相关的功能"""
    
    @staticmethod
    def create_entry(key, value):
        """创建一个单独的entry元素"""
        entry = etree.Element("entry")
        key_elem = etree.SubElement(entry, "KEY")
        key_elem.text = str(key)
        value_elem = etree.SubElement(entry, "VALUE1")
        value_elem.text = str(value)
        return entry
    
    @staticmethod
    def save_xml_file(root, output_path):
        """保存XML文件并添加XML声明"""
        etree.ElementTree(root).write(output_path, pretty_print=True, encoding="utf-8")
        
        # 添加XML声明
        with open(output_path, 'r', encoding='utf-8') as file:
            xml_content = file.read()

        if not xml_content.startswith('<?xml'):
            xml_content = '<?xml version="1.0" encoding="UTF-8" standalone="no" ?>\n' + xml_content
            with open(output_path, 'w', encoding='utf-8') as file:
                file.write(xml_content)
    
    @staticmethod
    def excel_to_xml(input_path, output_path):
        """将Excel文件转换为XML格式"""
        try:
            # 读取Excel文件
            wb = load_workbook(input_path)
            ws = wb.active
            
            # 创建XML结构
            root = etree.Element("root")
            data = etree.SubElement(root, "data")
            lang_conv = etree.SubElement(data, "LanguageStringConvertor")
            
            # 遍历Excel行并创建entry元素
            for row in ws.iter_rows(min_row=2):  # 跳过标题行
                if row[0].value is not None:  # 只处理有KEY的行
                    key = row[0].value
                    value = row[1].value if len(row) > 1 and row[1].value is not None else ""
                    entry = XMLProcessor.create_entry(key, value)
                    lang_conv.append(entry)
            
            # 保存XML文件
            XMLProcessor.save_xml_file(root, output_path)
            return True
        except Exception as e:
            print(f"Excel转XML失败: {str(e)}")
            raise
    
    @staticmethod
    def xml_to_excel(input_file, output_file):
        """将XML文件转换为Excel格式"""
        try:
            # 解析XML文件
            tree = etree.parse(input_file)
            
            # 查找所有entry元素
            xpath_expression = "/root/data/LanguageStringConvertor/entry"
            items = tree.xpath(xpath_expression)

            # 创建新的Excel文件
            wb = Workbook()
            ws = wb.active
            ws.append(["KEY", "VALUE1"])  # 添加标题行

            # 提取XML数据并添加到Excel
            for item in items:
                id_value = item.find("KEY").text if item.find("KEY") is not None else ""
                name_value = item.find("VALUE1").text if item.find("VALUE1") is not None else ""
                ws.append([id_value, name_value])
            
            # 保存Excel文件
            wb.save(output_file)
            return True
        except Exception as e:
            print(f"XML转Excel失败: {str(e)}")
            raise


class ExcelProcessor:
    """Excel处理相关的功能"""
    
    @staticmethod
    def clean_sheet(sheet):
        """清理Excel表格中的空行和空列"""
        # 删除空行
        rows_to_delete = list(sheet.iter_rows(min_row=1, max_row=sheet.max_row))
        for row in reversed(rows_to_delete):
            if all(cell.value is None for cell in row):
                sheet.delete_rows(row[0].row, amount=1)

        # 删除空列
        cols_to_delete = list(sheet.iter_cols(min_col=1, max_col=sheet.max_column))
        for col in reversed(cols_to_delete):
            if all(cell.value is None for cell in col):
                sheet.delete_cols(col[0].col_idx, amount=1)
        
        return sheet
    
    @staticmethod
    def compare_excel(input_file, dist_file):
        """
        比较两个Excel表格的KEY和VALUE1字段，直接更新dist_file文件，
        并用颜色标记新增和修改的内容。
        """
        try:
            # 获取文件
            source = pd.read_excel(input_file)
            dist = pd.read_excel(dist_file)
            
            # 确保两个表格都有KEY和VALUE1列
            if 'KEY' not in source.columns or 'VALUE1' not in source.columns:
                raise ValueError("源文件缺少KEY或VALUE1列")
            if 'KEY' not in dist.columns or 'VALUE1' not in dist.columns:
                raise ValueError("目标文件缺少KEY或VALUE1列")
            
            # 创建源表的键值对字典
            source_dict = dict(zip(source['KEY'], source['VALUE1']))
            
            # 得到目标表原始列
            original_columns = dist.columns.tolist()
            
            # 创建结果表的副本
            result = dist.copy()
            
            # 记录需要标记的单元格
            cells_to_highlight = []  # [(行索引, "新增"或"修改")]
            
            # 检查修改的内容
            for idx, row in result.iterrows():
                key = row['KEY']
                if key in source_dict and row['VALUE1'] != source_dict[key]:
                    # 值不同 - 更新值
                    result.at[idx, 'VALUE1'] = source_dict[key]
                    cells_to_highlight.append((idx, "修改"))
            
            # 检查新增的内容
            existing_keys = set(result['KEY'])
            new_rows = []
            for key, value in source_dict.items():
                if key not in existing_keys:
                    # 准备新行数据
                    new_row = {col: None for col in original_columns}
                    new_row['KEY'] = key
                    new_row['VALUE1'] = value
                    new_rows.append(new_row)
            
            # 添加新行到结果
            if new_rows:
                new_df = pd.DataFrame(new_rows)
                result = pd.concat([result, new_df], ignore_index=True)
                # 记录新增行的索引
                for i in range(len(new_rows)):
                    cells_to_highlight.append((len(dist) + i, "新增"))
            
            # 导出为Excel (覆盖原文件)
            result.to_excel(dist_file, index=False)
            
            # 打开并设置颜色
            wb = load_workbook(dist_file)
            ws = wb.active
            
            # 定义填充颜色
            new_fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")  # 绿色
            changed_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # 黄色
            
            # 为标记的单元格设置颜色
            for idx, status in cells_to_highlight:
                # 调整行索引：Excel行从1开始，第1行是标题，所以+2
                excel_row = idx + 2
                
                # 为整行设置颜色
                for col in range(1, ws.max_column + 1):
                    cell = ws.cell(row=excel_row, column=col)
                    if status == "新增":
                        cell.fill = new_fill
                    elif status == "修改":
                        cell.fill = changed_fill
            
            # 保存结果 (覆盖原文件)
            wb.save(dist_file)
            
            # 返回修改的统计信息
            modifications = sum(1 for _, status in cells_to_highlight if status == "修改")
            new_entries = sum(1 for _, status in cells_to_highlight if status == "新增")
            
            stats = {
                "modifications": modifications,
                "new_entries": new_entries
            }
            
            return stats
        except Exception as e:
            print(f"Excel对比失败: {str(e)}")
            raise


# 公共API函数，供其他模块调用
def convert_xml_to_excel(input_file, output_file):
    """将XML文件转换为Excel文件"""
    return XMLProcessor.xml_to_excel(input_file, output_file)

def convert_excel_to_xml(input_file, output_file):
    """将Excel文件转换为XML文件"""
    return XMLProcessor.excel_to_xml(input_file, output_file)

def compare_language_excel(input_file, dist_file):
    """比较和更新Excel文件"""
    stats = ExcelProcessor.compare_excel(input_file, dist_file)
    print(f"文件已更新: {dist_file}")
    print(f"已修改 {stats['modifications']} 个条目，新增 {stats['new_entries']} 个条目")
    return stats


from openpyxl import load_workbook
import lxml.etree as etree


"""
XML Functions
"""
class XMLFunctions :
    # child_data : Type :  TreeNode
    @staticmethod
    def create_Item(child_data) :
        root_item = etree.Element("entry")
        element_id = etree.SubElement(root_item, "KEY")
        element_id.text = str(child_data.data.id)
        element_name = etree.SubElement(root_item, "VALUE1")
        element_name.text = child_data.data.name
        return root_item
    
    @staticmethod
    def createXMLTree(root, root_xml, elements) :
        elements_xml = XMLFunctions.create_Items_List(elements)
        stack = [root, elements[0]]
        stack_xml = [root_xml, elements_xml[0]]
        root_xml.append(elements_xml[0])
        
        root.add_child(elements[0])
        for i in range(1, len(elements_xml)) :
            if elements[i - 1].data.level < elements[i].data.level :
                stack[len(stack) - 1].add_child(elements[i])
                if (len(stack_xml) == 1) :
                    stack_xml[0].append(elements_xml[i])
                else :
                    stack_xml[len(stack_xml) - 1].find("Children").append(elements_xml[i])
            else :
                while (elements[i].data.level <= stack[len(stack) - 1].data.level) :
                    stack.pop()
                    stack_xml.pop()
                stack[len(stack) - 1].add_child(elements[i])
                if (len(stack_xml) == 1) :
                    stack_xml[0].append(elements_xml[i])
                else :
                    stack_xml[len(stack_xml) - 1].find("Children").append(elements_xml[i])
            stack.append(elements[i])
            stack_xml.append(elements_xml[i])
        return root_xml
    
    @staticmethod
    def create_Items_List(elements) :
        elements_xml = []
        for element in elements :
            elements_xml.append(XMLFunctions.create_Item(element))
        return elements_xml
    
    @staticmethod
    def SaveXMLFile(root, output_path) :
        etree.ElementTree(root).write(output_path, pretty_print=True, encoding="utf-8")
        XMLFunctions.add_xml_declaration(output_path)
    
    @staticmethod
    def add_xml_declaration(xml_file):
    
        with open(xml_file, 'r', encoding='utf-8') as file:
            xml_content = file.read()

        # Add the XML declaration if it doesn't exist
        if not xml_content.startswith('<?xml'):
            xml_content = '<?xml version="1.0" encoding="UTF-8" standalone="no" ?>\n' + xml_content

            # Write the modified content back to the file
            with open(xml_file, 'w', encoding='utf-8') as file:
                file.write(xml_content)

    @staticmethod
    def createNewXMLFile(root_TreeNode, elements, output_path, name, editionVersion, year, month, day, source):
        root = etree.Element("BuildingInformation")
        root_xml = etree.SubElement(root, "Classification")
        system = etree.SubElement(root_xml, "System")
        
        name_element = etree.SubElement(system, "Name")
        name_element.text = name

        editionVersion_element = etree.SubElement(system, "EditionVersion")
        editionVersion_element.text = editionVersion

        editionDate = etree.SubElement(system, "EditionDate")
        
        year_element = etree.SubElement(editionDate, "Year")
        year_element.text = year

        month_element = etree.SubElement(editionDate, "Month")
        month_element.text = month

        day_element = etree.SubElement(editionDate, "Day")
        day_element.text = day

        description = etree.SubElement(system, "Description")

        source_element = etree.SubElement(system, "Source")
        source_element.text = source

        items = etree.SubElement(system, "Items")

        XMLFunctions.createXMLTree(root_TreeNode, items, elements)

        XMLFunctions.SaveXMLFile(root, output_path)

def excelSheet_modulation(sheet) :

     # Convert the generator to a list for reversing
    rows_to_delete = list(sheet.iter_rows(min_row=1, max_row=sheet.max_row))
    
    
    # Delete empty rows
    for row in reversed(rows_to_delete):
        if all(cell.value is None for cell in row):
            sheet.delete_rows(row[0].row, amount=1)

    # Convert the generator to a list for reversing
    cols_to_delete = list(sheet.iter_cols(min_col=1, max_col=sheet.max_column))
    
    # Delete empty columns
    for col in reversed(cols_to_delete):
        if all(cell.value is None for cell in col):
            sheet.delete_cols(col[0].col_idx, amount=1)
    
    return sheet


class ExcelElementsClass :
    def __init__(self, id, name, description, level) :
        self.id = id
        self.name = name
        self.description = description
        self.level = level
    @staticmethod
    def getAllRowsFromExcel(sheet):
        elements = []
        allRowsList = list(sheet.iter_rows(min_row=2, max_row=sheet.max_row))
        
        # add rows value to the elements, distinct: id, name, description, level
        for row in allRowsList:
            elements.append(TreeNode(ExcelElementsClass(*[cell.value for cell in row][:4])))

        return elements

class TreeNode:
    def __init__(self, data):
        self.data = data
        self.children = []


    def add_child(self, child_node):
        self.children.append(child_node)

def print_tree(node, level=0, prefix="Root"):
    if level == 0:
        print(f"{prefix} - {node.data.name}")
    else:
        indent = " " * (level * 4)
        print(f"{indent}└── {node.data.level} - {node.data.name}")

    for child in node.children:
        print_tree(child, level + 1, f"{prefix}.{child.data.level}")


def createTree(root, elements) :
    # LIFO : Last In First Out ==> Stack
    stack = [root, elements[0]]
    root.add_child(elements[0])
    for i in range(1, len(elements)) :
        if elements[i - 1].data.level < elements[i].data.level :
            stack[len(stack) - 1].add_child(elements[i])
        else :
            while (elements[i].data.level <= stack[len(stack) - 1].data.level) :
                stack.pop()
            stack[len(stack) - 1].add_child(elements[i])
        stack.append(elements[i])
    return root


def createXMLFile(input_path, output_path, name, editionVersion, year, month, day, source) :
    # Create a Tree
    root = TreeNode(ExcelElementsClass('0', "Persons",None, 0))

    excelFile = load_workbook(input_path)
    workSheet = excelFile.active
    # OR :  workSheet = excelFile["Sheet1"]

    # To delete the extra empty rows and cols
    workSheet = excelSheet_modulation(workSheet)

    # contain all the excel rows with distinct value according to ('id', 'name', 'description', 'Level')
    elements = ExcelElementsClass.getAllRowsFromExcel(workSheet)
    createTree(root, elements)
    
    # call createNewXMLFile, using the Tree Root, as an XML element, and then save it as an xml file .
    XMLFunctions.createNewXMLFile(root, elements, output_path, name, editionVersion, year, month, day, source)
    
    
    # Print the tree starting from the root
    ## print_tree(root)

    # excelFile.save('FamilyExcel.xlsx')
    # open_file('FamilyExcel.xlsx')




from openpyxl import Workbook
import lxml.etree as etree
import pandas as pd

def return_row(item) :
    
    id_value = item.find("KEY").text if item.find("KEY") is not None else ""

    name_value = item.find("VALUE1").text if item.find("VALUE1") is not None else ""

    try :
        return [id_value, name_value]
    except :
        return []
def convert_xml_to_excel(input_file, output_file) :
    # XML File
    tree = etree.parse(input_file)
    
    xpath_expression = "/root/data/LanguageStringConvertor/entry"
    items = tree.xpath(xpath_expression)

    # Create new Excel File
    wb = Workbook()
    ws = wb.active
    ws.append(["KEY", "VALUE1"])

    for item in items :
        ws.append(return_row(item))
    wb.save(output_file)

def compare_language_excel(input_file, dist_file):
    """
    比较两个Excel表格的KEY和VALUE1字段，直接更新dist_file文件，
    并用颜色标记新增和修改的内容。
    
    Args:
        input_file: 源文件路径
        dist_file: 目标文件路径（将被直接修改）
    """
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
            new_row = {col: None for col in original_columns}  # 先用None填充所有列
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
    
    # 直接更新目标文件
    from openpyxl.styles import PatternFill
    
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
    print(f"文件已更新: {dist_file}")
    
    # 返回修改的统计信息
    modifications = sum(1 for _, status in cells_to_highlight if status == "修改")
    new_entries = sum(1 for _, status in cells_to_highlight if status == "新增")
    
    print(f"已修改 {modifications} 个条目，新增 {new_entries} 个条目")
    
    return result


import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os

from scripts import convert_xml_to_excel, convert_excel_to_xml, compare_language_excel

class ExcelToXmlConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("游戏多语言工具")
        self.root.geometry("600x800")

        # 设置颜色主题
        self.setup_theme()
        
        # 创建主菜单
        self.create_main_menu()
    
    def setup_theme(self):
        """设置应用程序的主题和颜色"""
        # 设置柔和的颜色方案
        self.root.option_add("*TButton*background", "#AED6F1")  # 浅蓝色
        self.root.option_add("*TButton*foreground", "#2C3E50")  # 深蓝色
        self.root.option_add("*TEntry*background", "#D5F5E3")   # 浅绿色
        self.root.option_add("*TEntry*foreground", "#2C3E50")   # 深蓝色
        self.root.option_add("*TLabel*background", "#FDEBD0")   # 奶油色
        self.root.option_add("*TLabel*foreground", "#2C3E50")   # 深蓝色
    
    def create_main_menu(self):
        """创建主菜单界面"""
        self.button1 = ttk.Button(self.root, text="转换XML到Excel", command=self.convert_xml_to_excel_interface)
        self.button1.pack(pady=(50, 10))

        self.button2 = ttk.Button(self.root, text="多语言表格比对 Excel2Excel", command=self.compare_excel_to_excel_interface)
        self.button2.pack(pady=10)

        self.button3 = ttk.Button(self.root, text="Excel转XML", command=self.excel_to_xml_interface)
        self.button3.pack(pady=10)
    
    def convert_xml_to_excel_interface(self):
        """XML转Excel界面"""
        self.clear_root()

        # 输入XML文件
        self.input_file_label = ttk.Label(self.root, text="输入XML文件:")
        self.input_file_label.pack(pady=(10, 5))

        self.input_file_entry = ttk.Entry(self.root, state="disabled", width=40)
        self.input_file_entry.pack(pady=5)

        self.choose_file_button = ttk.Button(self.root, text="选择文件", command=lambda: self.choose_file("xml"))
        self.choose_file_button.pack(pady=(5, 10))

        # 输出文件夹和Excel文件
        self.output_folder_label = ttk.Label(self.root, text="输出文件夹:")
        self.output_folder_label.pack(pady=(10, 5))

        self.output_folder_entry = ttk.Entry(self.root, state="disabled", width=40)
        self.output_folder_entry.pack(pady=5)

        self.choose_folder_button = ttk.Button(self.root, text="选择文件夹", command=self.choose_folder)
        self.choose_folder_button.pack(pady=(5, 10))

        self.output_file_label = ttk.Label(self.root, text="Excel文件名:")
        self.output_file_label.pack(pady=(10, 5))

        self.output_file_entry = ttk.Entry(self.root, width=40)
        self.output_file_entry.pack(pady=5)

        # 运行转换按钮
        self.run_button = ttk.Button(self.root, text="开始转换", command=self.run_xml_to_excel_conversion)
        self.run_button.pack(pady=(20, 10))

        # 返回按钮
        self.return_button = ttk.Button(self.root, text="返回主菜单", command=self.return_to_menu)
        self.return_button.pack(pady=(20, 10))
    
    def excel_to_xml_interface(self):
        """Excel转XML界面"""
        self.clear_root()

        # 输入Excel文件
        self.input_file_label = ttk.Label(self.root, text="输入Excel文件:")
        self.input_file_label.pack(pady=(10, 5))

        self.input_file_entry = ttk.Entry(self.root, state="disabled", width=40)
        self.input_file_entry.pack(pady=5)

        self.choose_file_button = ttk.Button(self.root, text="选择文件", command=lambda: self.choose_file("excel"))
        self.choose_file_button.pack(pady=(5, 10))

        # 输出文件夹和XML文件
        self.output_folder_label = ttk.Label(self.root, text="输出文件夹:")
        self.output_folder_label.pack(pady=(10, 5))

        self.output_folder_entry = ttk.Entry(self.root, state="disabled", width=40)
        self.output_folder_entry.pack(pady=5)

        self.choose_folder_button = ttk.Button(self.root, text="选择文件夹", command=self.choose_folder)
        self.choose_folder_button.pack(pady=(5, 10))

        self.output_file_label = ttk.Label(self.root, text="XML文件名:")
        self.output_file_label.pack(pady=(10, 5))

        self.output_file_entry = ttk.Entry(self.root, width=40)
        self.output_file_entry.pack(pady=5)

        # 运行转换按钮
        self.run_button = ttk.Button(self.root, text="开始转换", command=self.run_excel_to_xml_conversion)
        self.run_button.pack(pady=(20, 10))

        # 返回按钮
        self.return_button = ttk.Button(self.root, text="返回主菜单", command=self.return_to_menu)
        self.return_button.pack(pady=(20, 10))
    
    def compare_excel_to_excel_interface(self):
        """Excel文件对比界面"""
        self.clear_root()

        # 导入源目标表格
        self.input_file_label = ttk.Label(self.root, text="翻译源目标:")
        self.input_file_label.pack(pady=(10, 5))

        self.input_file_entry = ttk.Entry(self.root, state="disabled", width=40)
        self.input_file_entry.pack(pady=5)

        self.choose_file_button = ttk.Button(self.root, text="选择文件", command=lambda: self.choose_file("excel"))
        self.choose_file_button.pack(pady=(5, 10))

        # 导入翻译母本
        self.output_folder_label = ttk.Label(self.root, text="翻译母本:")
        self.output_folder_label.pack(pady=(10, 5))

        self.output_folder_entry = ttk.Entry(self.root, state="disabled", width=40)
        self.output_folder_entry.pack(pady=5)

        self.choose_folder_button = ttk.Button(self.root, text="选择文件", command=lambda: self.choose_file_dist("excel"))
        self.choose_folder_button.pack(pady=(5, 10))

        # 执行对比按钮
        self.run_button = ttk.Button(self.root, text="执行对比", command=self.run_compare_conversion)
        self.run_button.pack(pady=(20, 10))

        # 返回按钮
        self.return_button = ttk.Button(self.root, text="返回主菜单", command=self.return_to_menu)
        self.return_button.pack(pady=(20, 10))
    
    def choose_file(self, fileType):
        """选择文件对话框"""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls") if fileType == "excel" else ("XML files", "*.xml")])
        if file_path:
            self.input_file_entry.config(state="normal")
            self.input_file_entry.delete(0, tk.END)
            self.input_file_entry.insert(0, file_path)
            self.input_file_entry.config(state="disabled")

    def choose_file_dist(self, fileType):
        """选择目标文件对话框"""
        file_path_dist = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls") if fileType == "excel" else ("XML files", "*.xml")])
        if file_path_dist:
            self.output_folder_entry.config(state="normal")
            self.output_folder_entry.delete(0, tk.END)
            self.output_folder_entry.insert(0, file_path_dist)
            self.output_folder_entry.config(state="disabled")

    def choose_folder(self):
        """选择文件夹对话框"""
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.output_folder_entry.config(state="normal")
            self.output_folder_entry.delete(0, tk.END)
            self.output_folder_entry.insert(0, folder_path)
            self.output_folder_entry.config(state="disabled")
    
    def run_excel_to_xml_conversion(self):
        """执行Excel转XML转换"""
        input_file = self.input_file_entry.get()
        output_folder = self.output_folder_entry.get()
        output_file_name = self.output_file_entry.get()

        if not input_file or not output_folder or not output_file_name:
            messagebox.showerror("错误", "请选择输入文件、输出文件夹并提供XML文件名。")
            return

        # 处理文件扩展名
        index = output_file_name.find('.')
        if index != -1:
            output_file_name = output_file_name[:index]
        output_file_name += '.xml'
        output_file_path = os.path.join(output_folder, output_file_name)

        try:
            convert_excel_to_xml(input_file, output_file_path)
            messagebox.showinfo("成功", "转换成功！")
        except Exception as e:
            messagebox.showerror("错误", f"发生错误: {str(e)}")
    
    def run_compare_conversion(self):
        """执行Excel对比和更新"""
        input_file = self.input_file_entry.get()
        dist_file = self.output_folder_entry.get()

        if not input_file or not dist_file:
            messagebox.showerror("错误", "路径不能为空。")
            return
        
        try:
            stats = compare_language_excel(input_file, dist_file)
            messagebox.showinfo("成功", f"对比成功！已修改 {stats['modifications']} 个条目，新增 {stats['new_entries']} 个条目")
        except Exception as e:
            messagebox.showerror("错误", f"发生错误: {str(e)}")

    def run_xml_to_excel_conversion(self):
        """执行XML转Excel转换"""
        input_file = self.input_file_entry.get()
        output_folder = self.output_folder_entry.get()
        output_file_name = self.output_file_entry.get()

        if not input_file or not output_folder or not output_file_name:
            messagebox.showerror("错误", "请选择输入文件、输出文件夹并提供Excel文件名。")
            return

        # 处理文件扩展名
        index = output_file_name.find('.')
        if index != -1:
            output_file_name = output_file_name[:index]
        output_file_name += '.xlsx'
        output_file_path = os.path.join(output_folder, output_file_name)

        try:
            convert_xml_to_excel(input_file, output_file_path)
            messagebox.showinfo("成功", "转换成功！")
        except Exception as e:
            messagebox.showerror("错误", f"发生错误: {str(e)}")

    def clear_root(self):
        """清空窗口内容"""
        for widget in self.root.winfo_children():
            widget.destroy()

    def return_to_menu(self):
        """返回主菜单"""
        self.clear_root()
        self.create_main_menu()


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToXmlConverterApp(root)
    root.mainloop()

# EXCEL2XML

Fork From https://github.com/KhaledKammoun/EXCEL2XML

## 简介
主要用于fairyGUI多语言对比，转出excel表格导出

## 功能
- XML转Excel：支持从XML导出为Excel格式
- Excel对比：对比两个Excel表格，自动标记新增和修改的内容
- 自动更新：直接更新目标文件并用颜色标记变更

## 使用方法
1. 运行可执行文件EXCEL2XML.exe
2. 选择源Excel文件和目标Excel文件
3. 点击"比较并更新"按钮
4. 查看更新后的目标文件，新增条目将以绿色背景显示，修改条目将以黄色背景显示

## 下载
从[Releases](链接到你的GitHub发布页面)页面下载最新版本

## 从源码构建
```bash
# 安装依赖
pip install -r requirements.txt

# 运行
python scripts.py

# 构建可执行文件
pip install pyinstaller
pyinstaller --onefile scripts.py
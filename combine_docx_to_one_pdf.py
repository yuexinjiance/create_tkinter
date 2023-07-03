from tkinter import messagebox
# 导入PDF操作工具
from PyPDF2 import PdfMerger
# 导入word转换pdf工具
from docx2pdf import convert
import os


def combine_process(factory_name):
    # 报告的word文件夹路径以及PDF文件夹路径
    path_to_word_files = fr'{factory_name}'
    new_pdf_dir_name = fr'{factory_name}-PDF版报告'

    # 创建公司PDF文件夹存放所有转换的单个PDF文件
    os.makedirs(new_pdf_dir_name)
    # 将word文件夹中的所有文件转换为PDF，并保存到PDF的文件夹中
    convert(path_to_word_files, new_pdf_dir_name)

    # 合并PDF文件夹中的所有文件并输出
    merger = PdfMerger()
    for root, dirs, file_names in os.walk(new_pdf_dir_name):
        for file_name in file_names:
            merger.append(new_pdf_dir_name + '/' + file_name)

    merger.write(fr"{factory_name}最终PDF合并版.pdf")
    merger.close()
    messagebox.showinfo(title='合并完成', message='文件合并已完成!')


if __name__ == "__main__":
    combine_process("绍兴鑫沃工程有限公司")

from docx2pdf import convert
import os


factory_name = "绍兴鑫沃工程有限公司"

path_to_word_files = fr'{factory_name}'
new_pdf_dir_name = fr'{factory_name}-PDF版报告'

os.makedirs(new_pdf_dir_name)
convert(path_to_word_files, new_pdf_dir_name)
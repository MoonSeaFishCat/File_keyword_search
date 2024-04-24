import os
import re
from collections import defaultdict
import pandas as pd
import fitz  # PyMuPDF


def normalize_school_name(school):
    return re.sub(r'\s+', r'\\s+', school)


def find_schools_in_pdf(pdf_path, schools):
    school_count = defaultdict(int)
    # 打开PDF文件
    with fitz.open(pdf_path) as doc:
        text = ""
        for page in doc:
            text += page.get_text().lower()  # 获取页面文本并转换为小写

    # 对于每一个学校，生成一个模糊匹配的正则表达式并计数
    for school in schools:
        normalized_school = normalize_school_name(school)
        pattern = re.compile(normalized_school, flags=re.I)
        school_count[school] += len(pattern.findall(text))

    return school_count


current_directory = os.getcwd()
data_file_path = os.path.join(current_directory, 'data.txt')

with open(data_file_path, 'r', encoding='utf-8') as file:
    schools = [line.strip() for line in file.readlines()]

results = []
for root, dirs, files in os.walk(os.path.join(current_directory, 'input')):
    print(root)
    for file in files:
        if file.lower().endswith('.pdf'):
            pdf_path = os.path.join(root, file)
            relative_folder = os.path.relpath(root, start=current_directory)
            school_count = find_schools_in_pdf(pdf_path, schools)
            for school, count in school_count.items():
                if count > 0:
                    results.append({
                        'File Path': pdf_path,
                        'Folder': relative_folder,
                        'School': school,
                        'Count': count
                    })

df = pd.DataFrame(results)
output_directory = os.path.join(current_directory, 'output')
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

excel_path = os.path.join(output_directory, 'schools_count_pdf.xlsx')
df.to_excel(excel_path, index=False)

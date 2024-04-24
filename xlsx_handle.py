import os
import re
from collections import defaultdict
import pandas as pd


def normalize_school_name(school):
    return re.sub(r'\s+', r'\\s+', school)


def find_schools_in_excel(excel_path, schools):
    school_count = defaultdict(int)
    # 读取Excel文件所有的sheet
    xls = pd.ExcelFile(excel_path)
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
        # 将所有字符串列联合为一个长文本
        text = " ".join(df.select_dtypes(include=['object']).astype(str).values.flatten()).lower()

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
        if file.lower().endswith(('.xlsx', '.xls')):
            excel_path = os.path.join(root, file)
            relative_folder = os.path.relpath(root, start=current_directory)
            school_count = find_schools_in_excel(excel_path, schools)
            for school, count in school_count.items():
                if count > 0:
                    results.append({
                        'File Path': excel_path,
                        'Folder': relative_folder,
                        'School': school,
                        'Count': count
                    })

df = pd.DataFrame(results)
output_directory = os.path.join(current_directory, 'output')
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

excel_path = os.path.join(output_directory, 'schools_count_excel.xlsx')
df.to_excel(excel_path, index=False)

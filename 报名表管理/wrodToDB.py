import os
import pandas as pd
from docx import Document

# 创建一个空的DataFrame，用于存储提取的数据
data_df = pd.DataFrame(columns=["姓名", "性别", "年龄"])

# 指定存储Word文档的目录
docx_directory = "data"

# 遍历目录下的所有Word文档
for filename in os.listdir(docx_directory):
    if filename.endswith(".docx"):
        doc = Document(os.path.join(docx_directory, filename))

        # 初始化一个字典来存储当前文档的数据
        data = {"姓名": None, "性别": None, "年龄": None}

        # 遍历文档段落并提取字段值
        for para in doc.paragraphs:
            for key in data:
                if key in para.text:
                    data[key] = para.text.split(":")[1].strip()

        # 添加当前文档的数据到DataFrame
        data_df = data_df._append(data, ignore_index=True)

# 保存数据到Excel表
data_df.to_excel("output_data.xlsx", index=False)

# 案例自动分析软件






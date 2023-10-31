import pandas as pd
import docx

# 打开Word文档

doc = docx.Document('测试案例1.docx')

# 获取第一个表格

table = doc.tables[0]
print(len(table.rows))
# 遍历表格中的行和列

for i, row in enumerate(table.rows):
    original_list = [cell.text for cell in row.cells]
    # 使用列表推导式过滤空白数据并保持顺序
    filtered_list = [x for x in original_list if x is not None and x != '']

    # 创建一个新列表来保持顺序，同时删除重复数据
    unique_ordered_list = []
    for item in filtered_list:
        if item not in unique_ordered_list:
            unique_ordered_list.append(item)
    print(unique_ordered_list)

import docx
import re

# 设置参数
input_file = 'input.docx'
table_width = "\\textwidth"


def read_table(docx_file):
    # 读取docx文件
    doc = docx.Document(docx_file)

    # 获取所有表格
    tables = doc.tables
    t = 0
    data = [[] for i in range(len(tables))]
    for table in tables:

        # 获取表格行数和列数
        rows = len(table.rows)
        cols = len(table.columns)

        # 创建一个空的二维列表
        data[t] = [[0] * cols for i in range(rows)]

        # 将表格内容存入二维列表中
        for i in range(rows):
            for j in range(cols):
                # 去除多余空格和自动编号
                cell_text = re.sub(r"\n", " ", table.cell(i, j).text)
                data[t][i][j] = re.sub(r"(^ +| +$)", "", cell_text)

        t += 1
        # 返回二维列表
    return data


def table2latex(indata, width="0.9\\textwidth"):
    # 表格数
    num_tables = len(indata)
    latex_code = ''
    for data in indata:
        # data = indata[0]
        # 获取表格列数
        num_cols = len(data[0])

        # 表格格式字符串
        table_format = "| " + " | ".join(["c"] * num_cols) + " |"

        # 生成表格代码
        latex_code += "\n"
        latex_code += "\\begin{table}\n\\centering\n"
        latex_code += "\\resizebox{" + width + "}{!}{"
        latex_code += "\\begin{tabular}{" + table_format + "}\n"
        latex_code += "\\hline\n"

        # 添加表头
        header = data[0]
        latex_code += " & ".join(header) + " \\\\\n"
        latex_code += "\\hline\n"

        # 添加表格内容
        for row in data[1:]:
            latex_code += " & ".join(row) + " \\\\\n"
            latex_code += "\\hline\n"

        latex_code += "\\end{tabular}\n}\\end{table}"

    return latex_code



def TableLaTex(in_file, table_width, out_file='output.tex'):
    # 读取word表格
    data = read_table(in_file)

    # 生成LaTeX代码（设置表格宽度为0.7倍文本宽度）
    latex_code = table2latex(data, width=table_width)

    # 将LaTeX代码保存到文本文档中
    with open(out_file, 'w', encoding='utf-8') as f:
        f.write(latex_code)

    #
    return latex_code


# 打印LaTeX代码
print(TableLaTex(input_file, table_width))

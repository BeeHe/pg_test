"""封装读写Microsoft Document文件包."""
# coding: utf-8


import os
import json

from docx import Document
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, _Row, Table
from docx.text.paragraph import Paragraph


def iter_block_items(parent):
    """
    遍历文档所有的段落和表格.

    Generate a reference to each paragraph and table child within *parent*,
    in document order. Each returned value is an instance of either Table or
    Paragraph. *parent* would most commonly be a reference to a main
    Document object, but also works for a _Cell object, which itself can
    contain paragraphs and tables.
    """
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    elif isinstance(parent, _Row):
        parent_elm = parent._tr
    else:
        raise ValueError("something's not right")
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def convert_tbl2dict(table):
    """将表格转化为字典类型."""
    if not isinstance(table, Table):
        raise TypeError('not Table type')
    # 中文名称	英文名称	类型	非空	注释
    column_list = ['cn_name', 'name', 'type', 'not_null', 'comment']
    table_structure = {}
    for row in table.rows[1:]:
        column_info = dict(zip(
            column_list, [cell.text.lower().strip() for cell in row.cells]))
        column_name = column_info['name']
        # 将oracle类型转化为postgres的类型
        column_info['type'] = (column_info['type'].lower()
                               .replace('number', 'numeric')
                               .replace('varchar2', 'varchar')
                               .replace('date', 'timestamp'))
        table_structure[column_name] = column_info

        # for cell in row.cells:
        #     for paragraph in cell.paragraphs:
        #         row_data.append(paragraph.text)
        # print("\t".join(row_data))
    return table_structure


def get_docx_table(path):
    """将数模文档中的关系型数据库表章节中的表格转换为字典."""
    result_info = {}
    table_par_list = []
    heading_par = None
    document = Document(path)
    for block in iter_block_items(document):
        # print(block.text if isinstance(block, Paragraph) else '<table>')
        if isinstance(block, Paragraph):
            style_name = block.style.name
            if style_name.startswith('Heading'):
                heading_par = block
                head_level = int(style_name.split(' ')[-1])
                # 重新生成从标题到head1的结果
                table_par_list = table_par_list[:head_level - 1] + [block]
                print([(par.text, par.style.name) for par in table_par_list])
            elif table_par_list:
                table_par_list.append(block)
        elif isinstance(block, Table):
            # 获取表英文名
            if not table_par_list or heading_par is None:
                continue
            table_name = heading_par.text.split('/')[-1].split('（')[0].lower()
            if not table_name:
                print(' '.join(par.text for par in table_par_list))

            table_info = {}
            table_structure = convert_tbl2dict(block)
            table_info['structure'] = table_structure

            # 获得主键字段
            normal_par_list = [par for par in table_par_list[head_level:-1]
                               if par.style.name == 'Normal' and '主键' in par.text]
            if not normal_par_list:
                continue
            pk_columns = [i.lower().split('/')[-1].strip()
                          for i in normal_par_list[0].text.split('+')[-1]]
            table_structure['pk_columns'] = pk_columns
            result_info[table_name] = table_info
    return result_info


def convert_table2sql(table_info: dict):
    """将docx中读取的表格的字典信息转化为SQL建表语句."""
    base_sql = """
    CREATE TABLE IF NOT EXISTS {table} (
     {column}
);
"""
    table_structure = table_info['structure']
    column_str = '\n\t'.join(
        f", {column_name} {column_info['type']} "
        f"{'NOT NULL' if column_info['not_null'] == 'y' else '' } "
        for column_name, column_info
        in table_structure.items())
    table_name = table_info['name']
    result_sql = base_sql.format(table_name, column_str)
    if table_info.get('pk_columns'):
        primary_constraint = ("ALTER TABLE ON {table_name} ADD CONSTRAINT "
                              f"pk_{table_name} ({table_info['pk_columns']});")
        result_sql += primary_constraint
    return result_sql


# 测试
if __name__ == '__main__':
    # tables
    '''
    以表格为单位存储到字典中->
    {}'''
    path = '/Users/HeBee/PersonalData/example/GW-MD02_994.docx'
    json_file = '/Users/HeBee/PersonalData/example/docx_tbl.json'
    if os.path.exists(json_file):
        with open(json_file, 'r') as f:
            raw_json = f.read()
        tables_info = json.loads(raw_json)
    else:
        tables_info = get_docx_table(path)
        with open(json_file, 'w') as f:
            f.write(json.dumps(tables_info, ensure_ascii=False, indent=4))

    for tbl_name, table_info in tables_info.items():
        sql = convert_table2sql(table_info)
        print(sql)
    # with open('/tmp/test.csv', 'r') as fileobj:
    #     csv_writer = csv.writer(fileobj)
    #     title = []
    #     csv_writer.writerow()
    # from pprint import pprint
    # pprint(table_info)

# -*- coding: utf-8 -*-
from export.type import *

# 第一列：Excel的列名，第二列：导出的字段名，第三列：字段类型
define = (
    ('编号', 'id', Int()),
    ('姓名', 'name', Str()),
    ('绩点', 'point', Float()),
    ('选修科目', 'subjects', List(Int())),
    ('成绩', 'score', Dict(Int(), Int())),
)

config = {
    # 源文件 excel 路径
    'source': '../excel/sample.xls',
    # 导出目标文件
    'targets': (
        ('../python/sample.py', 'py'),
        ('../lua/sample.lua', 'lua'),
        ('../json/sample.json', 'json'),
        ('../js/sample.js', 'js'),
    ),
    # 表格主键,在 define 中定义的值
    'key': 'id',
    # 要到处的sheet
    'sheet': '第一页'
}

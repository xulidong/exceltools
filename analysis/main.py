# -*- coding: utf-8 -*-
import xlwt


def get_table():
    """返回一个测试数据字典"""
    tab = {
        1: {"id": 1, "name": "赵钱", "point": 2.1, "score": {1: 100, 2: 90}, "subjects": [1, 2]},
        2: {"id": 2, "name": "孙李", "point": 3.5, "score": {1: 95}, "subjects": [1, ]},
        3: {"id": 3, "name": "周吴", "point": 2.02, "score": {2: 98}, "subjects": [2, ]},
        4: {"id": 4, "name": "郑王", "point": 3.0, "score": {2: 67, 4: 55}, "subjects": [2, 4]},
    }
    return tab


def save_to_excel(tab):
    """将字典保存到文件"""
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('Sheet1')
    sort_ids = sorted(tab)
    col = False
    for i, obj_id in enumerate(sort_ids):
        obj = tab[obj_id]
        sort_keys = sorted(obj)
        if not col:
            for j in range(len(sort_keys)):
                worksheet.write(0, j, sort_keys[j])
            col = True
        for k, key in enumerate(sort_keys):
            val = obj[key]
            if isinstance(val, dict) or isinstance(val, list) or isinstance(val, tuple):
                val = str(val)
            worksheet.write(i + 1, k, val)
    workbook.save('sample.xls')


def main():
    tab = get_table()
    save_to_excel(tab)


if __name__ == "__main__":
    main()

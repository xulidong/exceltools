# -*- coding: utf-8 -*-
import os
import glob
import platform
import sys
import time
import traceback
import xlrd


DEFINE_PATH = './define'


"""
公共方法
"""


def encode_filename(filename):
    """将文件名编码成utf-8的格式 """
    if platform.system() == 'Windows':
        return filename.decode("gbk").encode("utf-8")
    else:
        return filename


def get_file_basename(path):
    """取文件的基础名，不包括后缀"""
    filename = os.path.split(path)[1]
    return os.path.splitext(filename)[0]


def abspath(filename):
    """取绝对路径"""
    if platform.system() == 'Windows':
        return os.path.abspath(filename.decode("utf-8").encode("gbk")).decode("gbk").encode("utf-8")
    else:
        return os.path.abspath(filename)


def isabs(filename):
    """是否是绝对路径"""
    if platform.system() == 'Windows':
        return os.path.isabs(filename.decode("utf-8").encode("gbk"))
    else:
        return os.path.isabs(filename)


def get_abs_path(path, filename):
    """取文件的绝对路径"""
    if not isabs(filename):
        return abspath(os.path.join(path, filename))
    else:
        return filename


def import_module(name):
    """动态导入模块"""
    if platform.system() == 'Windows':
        return __import__(name.decode("utf-8").encode("gbk"))
    else:
        return __import__(name)


def decode_str(s):
    """解成unicode"""
    if not isinstance(s, str):
        return s

    if isinstance(s, unicode):
        return s

    try:
        return s.decode("utf-8")
    except UnicodeError:
        return s.decode("gbk")


def encode_str(s):
    """编码成utf-8"""
    if isinstance(s, unicode):
        return s.encode("utf-8")

    if not isinstance(s, str):
        return s

    try:
        return s.decode('gbk').encode("utf-8")
    except UnicodeError:
        return s


def getmtime(f):
    """取文件的访问时间"""
    if platform.system() == 'Windows':
        return os.path.getmtime(f.decode("utf-8").encode("gbk"))
    else:
        return os.path.getmtime(f)


def openfile(f, m):
    """打开文件"""
    if platform.system() == 'Windows':
        return open(f.decode("utf-8").encode("gbk"), m)
    else:
        return open(f, m)


def is_file_exists(f):
    """判断文件是否存在"""
    if platform.system() == 'Windows':
        return os.path.exists(f.decode("utf-8").encode("gbk"))
    else:
        return os.path.exists(f)


"""
逻辑相关
"""


def need_convert(define_path, define_file, define_module):
    source_time = 0
    source_file = get_abs_path(define_path, define_file)
    try:
        source_time = max(source_time, getmtime(source_file))
        source_file = get_abs_path(define_path, define_module.config["source"])
        source_time = max(source_time, getmtime(source_file))

        target_time = time.time()
        for file_path, file_type in define_module.config["target"]:
            file_path = get_abs_path(define_path, file_path)
            if is_file_exists(file_path):
                target_time = min(target_time, getmtime(file_path))
            else:
                return True
        return target_time < source_time
    except Exception:
        traceback.print_exc()
        print decode_str(source_file)
        raise


def scan_define_files(path):
    """扫描配置定义文件"""
    define_files = glob.glob1(path, '*.py')
    for i, filename in enumerate(define_files):
        define_files[i] = encode_filename(filename)
    return define_files


def get_source_table(define_path, source_file, sheet):
    try:
        source_file = get_abs_path(define_path, source_file)
        source_doc = xlrd.open_workbook(decode_str(source_file))
        if not sheet:
            source_table = source_doc.sheet_by_index(0)
        else:
            source_table = source_doc.sheet_by_name(decode_str(sheet))
    except IOError:
        print decode_str("  错误：打开源文件失败：%s" % source_file)
        traceback.print_exc()
        return None

    if source_table.nrows <= 0:
        print decode_str("  错误：源文件没有内容：%s" % source_file)
        return None
    return source_table


def get_target_table(source_file, source_table, define_module, sheet_name):
    target_table = {}
    col_list = set()

    for row in range(1, source_table.nrows):
        row_dict = {}
        key = None
        col_name = None
        try:
            for col in range(source_table.ncols):
                col_name = encode_str(source_table.cell(0, col).value)
                col_list.add(col_name)
                define_item = get_define_item(define_module, col_name)
                if define_item:
                    value = encode_str(source_table.cell(row, col).value)
                    if define_item[1] == define_module.config['key']:
                        if value == '':
                            key = None
                            break

                    value = define_item[2].convert(value)
                    row_dict[define_item[1]] = value
                    if define_item[1] == define_module.config['key']:
                        key = value
        except Exception:
            print decode_str("error source_file: %s  key: %s  sheet: %s  col: %s" % (
                source_file, key, (sheet_name or "Sheet1"), col_name))
            raise

        if key is not None:
            if key in target_table:
                print decode_str("  错误：存在相同的键值：%s : %s" % (source_file, key))
                return None
            else:
                target_table[key] = row_dict

    check_define_item(define_module, col_list, source_file, sheet_name)

    return target_table


def get_define_item(define_module, name):
    for define_item in define_module.define:
        if (define_item[0] == name) or (define_item[1] == name):
            return define_item
    return None


def check_define_item(define_module, excel_col_list, source_file, sheet_name):
    for define_item in define_module.define:
        exist = False
        if define_item[0] in excel_col_list:
            exist = True
        if define_item[1] in excel_col_list:
            exist = True
        if not exist:
            print decode_str("excel:{},sheet:{} 不存在导出列:{}".format(source_file,
                                                                        (sheet_name or "Sheet"),
                                                                        define_item[0]))


def write_to_targets(define_path, target_table, define_module):
    target_files = define_module.config.get("targets", ())
    for target_file in target_files:
        file_path = target_file[0]
        file_type = target_file[1]
        if file_type == 'py':
            write_to_python(define_path, target_table, file_path)
        elif file_type == 'json':
            write_to_json(define_path, target_table, file_path)
        elif file_type == 'lua':
            write_to_lua(define_path, target_table, file_path)
        elif file_type == 'js':
            write_to_js(define_path, target_table, file_path)
        else:
            print decode_str("  错误：不支持导出文件的类型: %s" % file_path)
            return


def write_to_python(define_path, target_table, target_file):
    def write_obj(obj, root=False):
        content = ''
        if isinstance(obj, int) or isinstance(obj, float) or isinstance(obj, long):
            content += str(obj)
        elif isinstance(obj, str):
            content += '"' + obj + '"'
        elif isinstance(obj, tuple):
            content += '('
            for i in xrange(len(obj)):
                content += write_obj(obj[i])
                if i == 0 or i < len(obj) - 1:
                    content += ', '
            content += ')'
        elif isinstance(obj, list):
            content += '['
            for i in xrange(len(obj)):
                content += write_obj(obj[i])
                if i == 0 or i < len(obj) - 1:
                    content += ', '
            content += ']'
        elif isinstance(obj, dict):
            if root:
                keys = obj.keys()
                keys.sort()
                content += '{\n'
                for k in keys:
                    v = obj[k]
                    content += '\t' + write_obj(k, False) + ': '
                    content += write_obj(v) + ', \n'
                content += '}'
            else:
                i = 0
                c = len(obj)
                keys = obj.keys()
                keys.sort()
                content += '{'
                for k in keys:
                    v = obj[k]
                    content += write_obj(k, False) + ': '
                    content += write_obj(v)
                    if i < c - 1:
                        content += ', '
                    i += 1
                content += '}'
        return content

    file_content = '# -*- coding: utf-8 -*-\n\n' + \
        get_file_basename(target_file) + ' = ' + write_obj(target_table, True)

    file_path = get_abs_path(define_path, target_file)
    with openfile(file_path, 'wb') as f:
        f.write(file_content)


def write_to_lua(define_path, target_table, target_file):
    def write_obj(obj, root=False, is_key=False):
        content = ''
        if isinstance(obj, bool):
            content += 'true' if obj else 'false'
        elif isinstance(obj, int) or isinstance(obj, float) or isinstance(obj, long):
            if is_key:
                content += '[' + str(obj) + ']'
            else:
                content += str(obj)
        elif isinstance(obj, str):
            if is_key:
                content += obj
            else:
                content += "'" + obj + "'"
        elif isinstance(obj, tuple) or isinstance(obj, list):
            content += '{'
            for i in xrange(len(obj)):
                content += write_obj(obj[i])
                if i == 0 or i < len(obj) - 1:
                    content += ', '
            content += '}'
        elif isinstance(obj, dict):
            if root:
                keys = obj.keys()
                keys.sort()
                content += '{\n'
                for k in keys:
                    v = obj[k]
                    content += '\t' + write_obj(k, False, True) + ' = '
                    content += write_obj(v) + ', \n'
                content += '}'
            else:
                i = 0
                c = len(obj)
                keys = obj.keys()
                keys.sort()
                content += '{'
                for k in keys:
                    v = obj[k]
                    content += write_obj(k, False, True) + ' = '
                    content += write_obj(v)
                    if i < c - 1:
                        content += ', '
                    i += 1
                content += '}'
        return content

    name = get_file_basename(target_file)
    file_content = 'local ' + name + ' = ' + write_obj(target_table, True)
    file_content = file_content + '\nreturn ' + name

    file_path = get_abs_path(define_path, target_file)
    with openfile(file_path, 'wb') as f:
        f.write(file_content)


def write_to_json(define_path, target_table, target_file):
    def write_obj(obj, root=False):
        content = ''
        if isinstance(obj, bool):
            content += 'true' if obj else 'false'
        elif isinstance(obj, int) or isinstance(obj, float) or isinstance(obj, long):
            content += str(obj)
        elif isinstance(obj, str):
            content += '"' + obj + '"'
        elif isinstance(obj, list) or isinstance(obj, tuple):
            content += '['
            for i in xrange(len(obj)):
                content += write_obj(obj[i])
                if i == 0 or i < len(obj) - 1:
                    content += ', '
            content += ']'
        elif isinstance(obj, dict):
            if root:
                i = 0
                c = len(obj)
                keys = obj.keys()
                keys.sort()
                content += '{\n'
                for k in keys:
                    v = obj[k]
                    content += '\t' + write_obj(str(k), False) + ': '
                    content += write_obj(v)
                    if i < c - 1:
                        content += ', '
                    content += '\n'
                    i += 1
                content += '}'
            else:
                i = 0
                c = len(obj)
                keys = obj.keys()
                keys.sort()
                content += '{'
                for k in keys:
                    v = obj[k]
                    content += write_obj(str(k), False) + ': '
                    content += write_obj(v)
                    if i < c - 1:
                        content += ', '
                    i += 1
                content += '}'
        return content

    file_content = write_obj(target_table, True)
    file_path = get_abs_path(define_path, target_file)
    with openfile(file_path, 'wb') as f:
        f.write(file_content)


def write_to_js(define_path, target_table, target_file):
    def write_obj(obj, root=False):
        content = ''
        if isinstance(obj, bool):
            content += 'true' if obj else 'false'
        elif isinstance(obj, int) or isinstance(obj, float) or isinstance(obj, long):
            content += str(obj)
        elif isinstance(obj, str):
            content += '"' + obj + '"'
        elif isinstance(obj, list) or isinstance(obj, tuple):
            content += '['
            for i in xrange(len(obj)):
                content += write_obj(obj[i])
                if i == 0 or i < len(obj) - 1:
                    content += ', '
            content += ']'
        elif isinstance(obj, dict):
            if root:
                i = 0
                c = len(obj)
                keys = obj.keys()
                keys.sort()
                content += '{\n'
                for k in keys:
                    v = obj[k]
                    content += '\t' + str(k) + ': '
                    content += write_obj(v)
                    if i < c - 1:
                        content += ', '
                    content += '\n'
                    i += 1
                content += '};'
            else:
                i = 0
                c = len(obj)
                keys = obj.keys()
                keys.sort()
                content += '{'
                for k in keys:
                    v = obj[k]
                    content += str(k) + ': '
                    content += write_obj(v)
                    if i < c - 1:
                        content += ', '
                    i += 1
                content += '}'
        return content

    name = get_file_basename(target_file)
    file_content = 'var ' + name + ' = ' + write_obj(target_table, True)
    file_content += '\n'
    file_content += 'module.exports = %s; ' % name

    file_path = get_abs_path(define_path, target_file)
    with openfile(file_path, 'wb') as f:
        f.write(file_content)


def convert_excels(define_path, define_files, force):
    for define_file in define_files:
        convert_excel(define_path, define_file, force)


def convert_excel(define_path, define_file, force):
    if define_path not in sys.path:
        sys.path.append(define_path)

    mod_name = get_file_basename(define_file)
    define_module = import_module(mod_name)

    if not hasattr(define_module, "config"):
        print decode_str("  错误：定义文件不存在Config：%s" % define_file)
        return

    source_file = define_module.config.get("source", None)
    if source_file is None:
        print decode_str("  错误：必须指定源文件, source - %s" % define_file)
        return

    target_files = define_module.config.get("targets", None)
    if target_files is None:
        print decode_str("  错误：必须指定导出文件, target - %s" % define_file)
        return

    if define_module.config.get("key", None) is None:
        print decode_str("  提示：未指定主键名")
        return

    if not force and not need_convert(define_path, define_file, define_module):
        # 非强制导出时，源文件无修改，无须导出
        return

    sheet_name = define_module.config.get('sheet')
    key_set = set()
    for index, item in enumerate(define_module.define):
        key = item[1]
        if key in key_set:
            print decode_str("{}.{} 存在相同的数据key:{}".format(source_file, sheet_name, key))
            return
        key_set.add(key)
        if len(item) < 3:
            print decode_str("{} {} define项数据不足".format(source_file, index))
            return

    source_table = get_source_table(define_path, source_file, sheet_name)
    if source_table is None:
        return

    try:
        target_table = get_target_table(source_file, source_table, define_module, sheet_name)
        if target_table is None:
            return

        write_to_targets(define_path, target_table, define_module)
        print decode_str("成功导出：%s" % define_file)
    except Exception:
        print decode_str("导出失败：%s" % define_file)
        print traceback.format_exc()
        raise


def main():
    define_path = os.path.abspath(DEFINE_PATH)
    define_files = scan_define_files(define_path)
    convert_excels(define_path, define_files, True)


if __name__ == "__main__":
    main()

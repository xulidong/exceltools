# -*- coding: utf-8 -*-
"""
支持的导出的数据类型定义：
Bool
Int/Int64
Float/Float64
List/Tuple
Dict
"""


class ConvertError(StandardError):
    pass


class Converter(object):
    def get_desc(self):
        pass

    def get_error_desc(self, data):
        return '%s 不是有效的 %s' % (data, self.get_desc())

    def convert(self, data):
        return data

    def get_default(self):
        return None


class Bool(Converter):
    def convert(self, data):
        try:
            if data == '':
                d = self.get_default()
            else:
                d = bool(int(float(data)))
        except ValueError:
            raise ConvertError(self.get_error_desc(data))
        return d

    def get_desc(self):
        return 'Bool'

    def get_default(self):
        return False


class Int(Converter):
    def convert(self, data):
        try:
            if data == '':
                d = self.get_default()
            else:
                if type(data) == str:
                    d = int(float(data))
                else:
                    d = int(data)
        except ValueError:
            raise ConvertError(self.get_error_desc(data))
        return d

    def get_desc(self):
        return 'Int'

    def get_default(self):
        return 0


class Int64(Converter):
    def convert(self, data):
        try:
            if data == '':
                d = self.get_default()
            else:
                if type(data) == str:
                    d = long(float(data))
                else:
                    d = long(data)
        except ValueError:
            raise ConvertError(self.get_error_desc(data))
        return d

    def get_desc(self):
        return 'Int64'

    def get_default(self):
        return 0


class Float(Converter):
    def convert(self, data):
        try:
            if data == '':
                d = self.get_default()
            else:
                d = float(data)
        except ValueError:
            raise ConvertError(self.get_error_desc(data))
        return d

    def get_desc(self):
        return 'Float'

    def get_default(self):
        return 0.0


class Float64(Converter):
    def convert(self, data):
        try:
            if data == '':
                d = self.get_default()
            else:
                d = float(data)
        except ValueError:
            raise ConvertError(self.get_error_desc(data))
        return d

    def get_desc(self):
        return 'Float64'

    def get_default(self):
        return 0.0


class Str(Converter):
    def convert(self, data):
        if data is None:
            data = self.get_default()
        else:
            try:
                if int(data) == float(data):
                    data = int(data)
            except ValueError:
                pass
            data = str(data)
        data = str(data).replace("\r\n", '\\n')
        data = str(data).replace("\n", '\\n')
        return data

    def get_desc(self):
        return 'Str'

    def get_default(self):
        return ''


class List(Converter):
    def __init__(self, date_type=Str(), separator=";"):
        super(List, self).__init__()
        self.data_type = date_type
        self.separator = separator

    def convert(self, data):
        data = str(data)
        if data == '':
            return self.get_default()
        try:
            values = data.split(self.separator)
            data_list = [self.data_type.convert(x) for x in values]
            if values[-1] == '':
                data_list.pop(-1)
        except Exception:
            raise ConvertError(self.get_error_desc(data))
        return data_list

    def get_desc(self):
        return 'List'

    def get_default(self):
        return []


class Tuple(List):
    def __init__(self, date_type=Str(), separator=";"):
        super(List, self).__init__()
        self.data_type = date_type
        self.separator = separator

    def convert(self, data):
        data = str(data)
        if data == '':
            return self.get_default()
        try:
            d = [self.data_type.convert(x) for x in data.split(self.separator)]
        except Exception:
            raise ConvertError(self.get_error_desc(data))
        return tuple(d)

    def get_desc(self):
        return 'Tuple'

    def get_default(self):
        return ()


class Dict(Converter):
    def __init__(self, key_type=Int(), date_type=Str(), separator=";"):
        super(Dict, self).__init__()
        self.separator = separator
        self.key_type = key_type
        self.data_type = date_type

    def convert(self, data):
        data = str(data)
        if data == '':
            return self.get_default()
        try:
            data_list = [key_val for key_val in data.split(self.separator)]
            if data_list[-1] == '':
                data_list.pop(-1)
            data_dict = {}
            for key_val in data_list:
                (key, val) = key_val.split(":")
                data_dict[self.key_type.convert(key)] = self.data_type.convert(val)
        except Exception:
            raise ConvertError(self.get_error_desc(data))
        return data_dict

    def get_desc(self):
        return 'Dict'

    def get_default(self):
        return {}

import xlrd
import re
import pymysql
from pymysql import ProgrammingError
import sys
import os
from configparser import ConfigParser


class Config:
    def __init__(self):
        if os.path.exists('conf.ini'):
            config = ConfigParser()
            config.read('conf.ini', encoding='utf-8')
            self.host = config.get('db', 'host')
            self.user = config.get('db', 'user')
            self.passwd = config.get('db', 'passwd')
            self.port = config.getint('db', 'port')
            self.database = config.get('db', 'database')
            self.filepath = config.get('db', 'filepath')
        else:
            print("找不到配置文件")
            os.system('pause')


class Table:

    def __init__(self, table):
        if is_table_name(table[0]):
            name, comment = get_table_name(table[0][0].value)
            self.table_name = name
            self.comment = comment
            self.columns = []
            for col in table[2:-1]:
                self.columns.append(Column.parse(col))
            # self.columns.append(Column)
            if table[-1][1].value:
                num = self.parse_l(table[-1][1].value)
                for i in range(int(num)):
                    self.columns.append(Column.create_l(i + 1))

    def __str__(self):
        return f'表名: {self.table_name.lower():8s} 说明: {self.comment:12s} 列：{self.columns}'

    @staticmethod
    def parse_l(col):
        return re.findall(r'(\d+)', col)[0]

    def sql(self):
        sql_str = f"create table {self.table_name.lower()}("
        for i, col in enumerate(self.columns):
            if i != len(self.columns) - 1:
                sql_str += f'{col.sql()},'
            else:
                sql_str += f'{col.sql()}'
        sql_str += f")comment= '{self.comment}' ENGINE=InnoDB DEFAULT CHARSET=utf8;"
        return self.table_name.lower(), sql_str


class Column:
    def __init__(self, comment, name, data_type, length, pk):
        self.comment = comment
        self.name = name
        self.data_type = data_type
        self.length = length
        self.pk = pk
        # self.default = default
        # self.isnull = isnull

    @classmethod
    def create_l(cls, num):
        return cls('', f'l{num}', 'varchar', '50', False)

    @classmethod
    def parse(cls, cols):
        comment = cols[1].value.strip()
        name = cols[2].value.strip()
        data_type = cols[3].value.strip()
        length = int(cols[4].value) if cols[4].value != '' else ''
        if cols[7].value.strip() == '':
            pk = False
        else:
            pk = True
        return cls(comment, name, data_type, length, pk)

    def sql(self):
        sql_str = f'`{self.name}` '
        if self.data_type == 'int':
            sql_str += ' int'
        elif self.data_type == 'varchar':
            sql_str += f' varchar({self.length})'
        elif self.data_type == 'datetime':
            sql_str += ' varchar(100)'

        if self.pk:
            sql_str += ' primary key'
        if self.comment:
            sql_str += f" comment '{self.comment}'"
        return sql_str

    def __str__(self):
        return f'列名: {self.name:8s} 说明：{self.comment:10s} 类型：{self.data_type:5s} 长度：{self.length} 主键：{self.pk} '

    def __repr__(self):
        return f'\n列名: {self.name:^8}说明：{self.comment:18}' \
               f'类型：{self.data_type:^20}长度:{self.length:^20}主键：{str(self.pk):^5}'


def is_table_name(row: list):
    """
    判断是否是数据库表名
    :param row:
    :return:
    """
    if row[0].value != '':
        for i in range(1, len(row)):
            if row[i].value != '':
                return False
    else:
        return False
    return True


def get_table_name(cell: str):
    """
    :return: 表名和说明
    """
    result = re.split(r'([a-zA-Z]+)\s*', cell)
    return result[1], result[2]


def slice_rows(sheet):
    nrows = sheet.nrows
    lists = []
    temp_list = []
    for rowx in range(nrows):
        row = sheet.row(rowx)
        if all([cell.value == '' for cell in row]):
            if len(temp_list) > 0:
                lists.append(temp_list)
                temp_list = []
        else:
            temp_list.append(row)
    if len(temp_list) > 0:
        lists.append(temp_list)
    return lists


# def read_excel(file):
#     data = xlrd.open_workbook(file)
#     for sheet_name in data.sheet_names():
#         if sheet_name != '物资管理系统后台菜单':
#             # print('\n' + "=" * 15 + sheet_name + "=" * 15)
#             sheet = data.sheet_by_name(sheet_name)
#             table_list = slice_rows(sheet)
#             # print(table_list[0])
#             create_table(table_list)


def create_table(table_list, config: Config):
    """
    将切分好的表格转化，在数据库建表
    :param config:
    :param table_list:
    :return:
    """
    db = pymysql.connect(config.host, config.user, config.passwd, config.database)
    cursor = db.cursor()
    for item in table_list:
        table = Table(item)
        table_name, sql = table.sql()
        try:
            cursor.execute(f"drop table if exists {table_name};")
            cursor.execute(sql)
            print(f'成功创建表 表名：{table_name}')
        except ProgrammingError:
            print(f"数据库执行失败！！  具体信息：{table}", file=sys.stderr)


def print_sql(table_list):
    with open('./output.sql', 'w', encoding='utf-8') as f:
        for item in table_list:
            table = Table(item)
            table_name, sql = table.sql()
            f.write(sql + '\n')


def cli(config: Config):
    if not config.filepath:
        print("请将excel拖入窗口或直接输入路径")
        path = input()
    else:
        path = config.filepath
    try:
        data = xlrd.open_workbook(path)
        print('请选择包含数据库的sheet，例如 0 1 2 ')

        for index, sheet_name in enumerate(data.sheet_names()):
            print(f'{index}: 【{sheet_name}】')
        li = input()
        print('请选择模式： 1：打印sql 2：直接操作数据库')
        mod = input()
        for i in li.split(" "):
            sheet = data.sheet_by_index(int(i))
            table_list = slice_rows(sheet)
            if mod.strip() == '1':
                print_sql(table_list)
            elif mod.strip() == '2':
                create_table(table_list, config)
            else:
                print("没有这个模式")
                sys.exit(0)
    except FileNotFoundError:
        print("请检查输入路径是否正确", file=sys.stderr)
    except (TypeError, IndexError):
        print("输入有误", file=sys.stderr)


if __name__ == '__main__':
    excel_path = r"C:\Users\zhao\Documents\Tencent Files\1847132713\FileRecv\【开发】物资管理系统数据库20200930.xlsx"
    # read_excel(excel_path)
    cfg = Config()
    cli(cfg)

    os.system('pause')

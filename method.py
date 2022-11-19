import xlrd as xl
from xlutils.copy import copy
import os



class method:
    def __int__(self):
        self.zhengshu()
        self.zifu()
        self.search()
        self.data_basics()



    #整数检查方法封装
    def zhengshu(self, num, num_column):
        sheet = xl.open_workbook('mainTask .xlsx').sheets()[1]
        num_line = range(num, int(sheet.nrows))
            #整数检查方法
        for line_int in num_line:
            data = sheet.cell_value(line_int, num_column)
                #检查data是否为整数
            if isinstance(data, float) and int(data) != None:
                print("第", line_int+1, "行是整数")


    #字符检查方法封装
    def zifu(self, num, num_column):
        sheet = xl.open_workbook('mainTask .xlsx').sheets()[1]
        num_line = range(num, int(sheet.nrows))
            #字符串检查方法
        for line_str in num_line:
            data1 = sheet.cell_value(line_str, num_column)
            #检查data是否为字符串
            if isinstance(data1, str) and len(data1) != 0:
                print('第', line_str+1, '行是字符串')

    #表格检查方法
    def jiancha(self, data, line, column, place, suffix, int_data):
        excel_data = self.search(self, place, suffix)
        #拿取表格基础数据：表格名称、分表名称、表格列数
        for excel in excel_data:
            excel_name = excel.get(self, 'excel')
            sheet_name = excel.get(self, 'sheet')
            maxcolumn = excel.get(self, 'column')
            #表格列循环
            for inspect_column in range(column, maxcolumn):
                inspect_sheet = xl.open_workbook(excel.get(self, 'excel'))
                inspect_sheet1 = inspect_sheet.sheet_by_name(excel.get(self, 'sheet'))
                #判断表头配置
                if inspect_sheet.cell_value(self, line, column) == int_data:
                    #单列检查循环
                    for data_line in range(line, inspect_sheet1.nrows):
                        inspect_data = inspect_sheet1.cell_value(self, line, column)
                        #整数检查方法
                        if not isinstance(inspect_data, int) and len(inspect_data) != 0:
                            print(inspect_sheet, '-', inspect_sheet1, '-', inspect_column, '列', data_line, '行与表头配置不符合')




    #表格名称、分表名称、表格列数获取
    def search(self, place, suffix):
        data = []
        file = os.listdir(place)
        for file_name in file:
            file = place + '/' + file_name
            if os.path.isfile(file) and suffix in file_name:
                sheet_name = xl.open_workbook(file)
                for sheet in sheet_name.sheet_names():
                    single_data = {"excel": "", "sheet": "", "column": ""}
                    single_data["excel"] = file_name
                    single_data["sheet"] = sheet
                    single_data["column"] = str(sheet_name.sheet_by_name(sheet).ncols)
                    data.append(single_data)
        return data


    #基础数据读取
    def data_basics(self, line, column):
        sheet = xl.open_workbook('data.xlsx').sheet_by_name("基础数据")
        data = sheet.cell_value(line, column)
        return data

    #基础数据写入
    def input_basics(self, line, column, data):
        basics = copy(xl.open_workbook("data.xlsx"))
        basics_data = basics.get_sheet("基础数据")
        basics_data.write( line, column, data)
        basics.save("data.xlsx")


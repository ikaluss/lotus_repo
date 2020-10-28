import xlrd
from xlrd import xldate_as_tuple
import xlwt
from xlutils.copy import copy
import datetime
'''
xlrd中单元格的数据类型
数字一律按浮点型输出，日期输出成一串小数，布尔型输出0或1，所以我们必须在程序中做判断处理转换
成我们想要的数据类型
0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
'''
class ExcelData():
    # 初始化方法
    def __init__(self, data_path, sheetname):
        
        #定义一个属性接收文件路径
        self.data_path = data_path
        
        # 定义一个属性接收工作表名称
        self.sheetname = sheetname
        
        # 使用xlrd模块打开excel表读取数据
        self.data = xlrd.open_workbook(self.data_path)
        
        # 根据工作表的名称获取工作表中的内容（方式①）
        self.table = self.data.sheet_by_name(self.sheetname)
        
        # 获取第一行所有内容,如果括号中1就是第二行，这点跟列表索引类似
        self.keys = self.table.row_values(0)
        
        # 获取工作表的有效行数
        self.rowNum = self.table.nrows
        
        # 获取工作表的有效列数
        self.colNum = self.table.ncols

    # 定义一个读取excel表的方法
    def readExcel(self):
        # 定义一个空列表
        datas = []
        for i in range(1, self.rowNum):
            # 定义一个空字典
            sheet_data = {}
            for j in range(self.colNum):
                # 获取单元格数据类型
                c_type = self.table.cell(i,j).ctype
                # 获取单元格数据
                c_cell = self.table.cell_value(i, j)
                if c_type == 2 and c_cell % 1 == 0:  # 如果是整形
                    c_cell = int(c_cell)
                elif c_type == 3:
                    # 转成datetime对象
                    date = datetime.datetime(*xldate_as_tuple(c_cell,0))
                    c_cell = date.strftime('%Y/%m/%d %H:%M:%S')
                elif c_type == 4:
                    c_cell = True if c_cell == 1 else False
                sheet_data[self.keys[j]] = c_cell
                # 循环每一个有效的单元格，将字段与值对应存储到字典中
                # 字典的key就是excel表中每列第一行的字段
                # sheet_data[self.keys[j]] = self.table.row_values(i)[j]
            # 再将字典追加到列表中
            datas.append(sheet_data)
        # 返回从excel中获取到的数据：以列表存字典的形式返回
        return datas

    def writeExcel(self,value,startCol):
        count = len(value)
        new_workbook = copy(self.data)  # 将xlrd对象拷贝转化为xlwt对象
        new_worksheet = new_workbook.get_sheet(0)

        for i in range(0,count):
            new_worksheet.write(self.rowNum, i+startCol, value[i])
        
        new_workbook.save(self.data_path)

    def overWriteExcel(self,value):
        count = len(value)
        new_workbook = copy(self.data)  # 将xlrd对象拷贝转化为xlwt对象
        new_worksheet = new_workbook.get_sheet(0)

        for i in range(0,count):
            new_worksheet.write(self.rowNum-1, i, value[i])
        
        new_workbook.save(self.data_path)


    
if __name__ == "__main__":
    # data_path = "a.xls"
    # sheetname = "专线日志"
    # get_data = ExcelData(data_path, sheetname)
    # datas = get_data.readExcel()
    # print(datas)

    # value = ["2020/07/07 10:10:52","2","电信FIS专线","2"]
    # get_data.writeExcel(value,7)

    filePath = "lineStatus.xls"
    sheetName = "status"
    getData = ExcelData(filePath, sheetName)
    value = [2,2,2,2,2,2]
    getData.overWriteExcel(value)



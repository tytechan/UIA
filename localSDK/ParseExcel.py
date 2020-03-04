from openpyxl.styles import Border,Side,Font
from config.ErrConfig import CNBMException, handleErr
import openpyxl
import win32com.client

class ParseExcel:
    def __init__(self):
        self.workbook = None
        self.excelFile = None

    # @CNBMException
    def loadWorkBook(self, filaPath):
        ''' 将excel文件加载到内存，并获取其workbook对象
        :param filaPath: 文件路径
        '''
        try:
            self.workbook = openpyxl.load_workbook(filaPath)
            self.excelFile = filaPath
            return self.workbook
        except Exception as e:
            handleErr(e)
            raise e

    # @CNBMException
    def createSheet(self, title=None, index=None):
        ''' 新建sheet
        :param title: sheet名，选填
        :param index: sheet序号
        :return:
        '''
        try:
            self.sheet = self.workbook.create_sheet(title=title, index=index)
            self.workbook.save(self.excelFile)
            return ParseSheet(self.workbook, self.excelFile, self.sheet)
        except Exception as e:
            handleErr(e)
            # raise e

    # @CNBMException
    def getSheet(self, sheetName=None, sheetIndex=0):
        ''' 获取sheet对象（若名称不为空，则优先通过名称查找；否则通过索引号查找，默认获取第一个sheet）
        :param sheetName: sheet名称
        :param sheetIndex: sheet序号
        '''
        try:
            if sheetName:
                self.sheet = self.workbook[sheetName]
                # sheet = self.workbook.get_sheet_by_name(sheetName)
            else:
                self.sheet = self.workbook.worksheets[sheetIndex]
            return ParseSheet(self.workbook, self.excelFile, self.sheet)
        except Exception as e:
            handleErr(e)
            raise e

    # @CNBMException
    def remove(self, sheetName=None, sheetIndex=None):
        ''' 删除选定sheet并保存
        （若名称不为空，则优先通过名称查找；否则通过索引号查找，默认获取第一个sheet）
        :param sheetName: sheet名称
        :param sheetIndex: sheet序号
        '''
        try:
            self.getSheet(sheetName, sheetIndex)
            self.workbook.remove(self.sheet)
            self.workbook.save(self.excelFile)
        except Exception as e:
            handleErr(e)
            raise e

    # @CNBMException
    def rename(self, newName, sheetName=None, sheetIndex=None):
        ''' 重命名选定sheet并保存
        （若名称不为空，则优先通过名称查找；否则通过索引号查找，默认获取第一个sheet）
        :param newName: 新定义的sheet名称
        :param sheetName: sheet名称
        :param sheetIndex: sheet序号
        '''
        try:
            self.getSheet(sheetName, sheetIndex)
            self.sheet.title = newName
            self.workbook.save(self.excelFile)
        except Exception as e:
            handleErr(e)
            raise e

    # @CNBMException
    def runVBA(self, func, visible=True):
        ''' 在 self.workbook 中运行VBA
        :param func: 函数名
        :param visible: 运行过程是否可见，默认可见
        '''
        try:
            xlApp = win32com.client.DispatchEx("Excel.Application")
            xlApp.Visible = visible
            xlApp.DisplayAlerts = 0
            xlBook = xlApp.Workbooks.Open(self.excelFile, False)
            xlBook.Application.Run(func)
            xlBook.Close(True)
            xlApp.quit()
        except Exception as e:
            handleErr(e)
            raise e

    # @CNBMException
    def sheets(self):
        '''
        返回 self.workbook 中所有sheet列表（list）
        '''
        try:
            return self.workbook.worksheets
        except Exception as e:
            handleErr(e)
            raise e

    # @CNBMException
    def create(self, filePath):
        ''' 创建新工作簿并保存
        :param filePath: 文件保存路径
        '''
        try:
            self.workbook = openpyxl.Workbook()
            self.excelFile = filePath
            self.workbook.save(self.excelFile)
            return self.workbook
        except Exception as e:
            handleErr(e)
            raise e


class ParseSheet:
    def __init__(self, workbook, excelFile, sheet):
        self.workbook = workbook
        self.excelFile = excelFile
        self.sheet = sheet
        #颜色对应的RGB值
        self.RGBDict = {'red':'FFFF3030', 'green':'FF008B00'}

    # @CNBMException
    def rowCount(self):
        '''
        获取sheet中有数据区域的结束行号
        '''
        try:
            return self.sheet.max_row
        except Exception as e:
            handleErr(e)
            raise e

    # @CNBMException
    def colCount(self):
        '''
        获取sheet中有数据区域的结束列号
        '''
        try:
            return self.sheet.max_column
        except Exception as e:
            handleErr(e)
            raise e

    # @CNBMException
    def minRowNum(self):
        '''
        获取sheet中有数据区域的开始行号
        '''
        try:
            return self.sheet.min_row
        except Exception as e:
            handleErr(e)
            raise e

    # @CNBMException
    def minColNum(self):
        '''
        获取sheet中有数据区域的开始列号
        '''
        try:
            return self.sheet.min_column
        except Exception as e:
            handleErr(e)
            raise e

    # @CNBMException
    def row(self, rowNo):
        ''' 获取sheet中某一行，返回该行所有数据内容组成的tuple
        :param rowNo: 行号（从1开始）
        '''
        try:
            myrows = self.sheet[rowNo]
            return myrows

            # colNo = ParseExcel().getColsNumber(sheet)
            # print("colNo:",colNo)
            # MyValue = []
            # for i in range(colNo):
            #     MyValue = MyValue.append(sheet.cell(rowNo,i).value)
            # return MyValue
        except Exception as e:
            handleErr(e)
            raise e

    # @CNBMException
    def column(self, colNo):
        ''' 获取sheet中某一列，返回该列所有数据内容组成的tuple
        :param colNo: 列号（从1开始）
        '''
        try:
            return self.sheet[colNo]
        except Exception as e:
            handleErr(e)
            raise e

    # @CNBMException
    def read(self, coordinate=None, rowNo=None, colsNo=None):
        '''
        根据单元格所在的位置索引获取该单元格中的值
        sheet.cell(row = 1,column = 1).value 表示excel中第一行第一列的值
        :param coordinate: 单元格坐标（“A3”、"B5"）
        :param rowNo: 行号（从1开始）
        :param colsNo: 列号（从1开始）
        '''
        try:
            if coordinate != None:
                # coordinate = coordinate.decode('utf-8')
                return self.sheet[coordinate].value
            elif coordinate is None and rowNo is not None and colsNo is not None:
                return self.sheet.cell(row=rowNo, column=colsNo).value
        except Exception as e:
            handleErr(e)
            raise e

    # @CNBMException
    def cell(self, coordinate=None, rowNo=None, colsNo=None):
        '''
        获取某个单元格对象，可以根据单元格所在位置的数字索引
        也可以直接根据excel中单元格的编码及坐标，
        如：getCell(sheet,coordinate = 'A1')
        或：getCell(sheet,rowNo = 1,colsNo = 2)
        :param coordinate: 单元格坐标（“A3”、"B5"）
        :param rowNo: 行号（从1开始）
        :param colsNo: 列号（从1开始）
        '''
        try:
            if coordinate != None:
                # coordinate = coordinate.decode('utf-8')
                return self.sheet[coordinate]
                # return sheet.cell(coordinate = coordinate)
            elif coordinate == None and rowNo is not None and colsNo is not None:
                return self.sheet.cell(row=rowNo,column=colsNo)
        except Exception as e:
            handleErr(e)
            raise e

    # @CNBMException
    def write(self, content, coordinate=None, rowNo=None, colsNo=None, style=None):
        '''
        根据单元格在excel中的编码坐标或者数字索引坐标向单元格中写入数据（写之前确认文件关闭），
        下表从1开始，参数style表示字体的颜色名称，如：red、green
        :param content: 输入内容
        :param coordinate: 单元格坐标（“A3”、"B5"）
        :param rowNo: 行号（从1开始）
        :param colsNo: 列号（从1开始）
        :param style: 字体的颜色名称（“red”、“green”，默认为黑色）
        '''
        try:
            if coordinate is not None:
                # coordinate = coordinate.decode('utf-8')
                self.sheet[coordinate].value = content
                # sheet.cell(coordinate = coordinate).value = content
                if style is not None:
                    self.sheet[coordinate].font = Font(color = self.RGBDict[style])
                    # sheet.cell(coordinate = coordinate).font = Font(color = self.RGBDict[style])
            elif coordinate == None and rowNo is not None and colsNo is not None:
                # sheet.cell(row = rowNo,column = colsNo).value = ""
                self.sheet.cell(row=rowNo, column=colsNo).value = content
                if style:
                    self.sheet.cell(row=rowNo, column=colsNo).font = Font(color=self.RGBDict[style])
            self.workbook.save(self.excelFile)
        except Exception as e:
            handleErr(e)
            raise e

    # @CNBMException
    def save(self):
        '''
        保存文件
        '''
        try:
            self.workbook.save(self.excelFile)
        except Exception as e:
            handleErr(e)
            raise e

    # @CNBMException
    def close(self):
        '''
        关闭文件，默认保存
        '''
        try:
            self.save()
            self.workbook.close()
        except Exception as e:
            handleErr(e)
            raise e


if __name__ == "__main__":
    # 调试前先注释 @CNBMException
    Excel = ParseExcel()
    # 创建新工作簿
    # Excel.create(r"E:\python相关\RPA\UIA\abc.xlsx")

    # 操作已有工作簿
    wb = Excel.loadWorkBook(r"E:\python相关\RPA\UIA\test.xlsm")
    print("所有sheet列表：", Excel.sheets())

    sht = Excel.getSheet("Sheet1")
    print("最大行数：%s" %sht.rowCount())
    print("最小行数：%s" %sht.minRowNum())
    print("最大列数：%s" %sht.colCount())
    print("最小列数：%s\n" %sht.minColNum())

    cell = sht.cell("B3")
    print("B3值：", cell.value)
    cell1 = sht.cell(rowNo=3, colsNo=3)
    print("C3值：", cell1.value)
    print("C2值：", sht.read("C2"))
    print("D2的值：", sht.read(rowNo=2, colsNo=4), "\n")

    col2 = sht.column(2)
    print("第二列为：", col2)
    row3 = sht.row(3)
    print("第三行为：", row3)
    print("C3的值：", row3[2].value, "\n")

    # 写值
    sht.write("test", "G4")
    sht.write("test1", rowNo=7, colsNo=5)

    # 创建sheet
    sht1 = Excel.createSheet("test")
    sht2 = Excel.createSheet("test1", 0)

    # 删除sheet
    Excel.remove("test")
    Excel.remove(sheetIndex=0)

    # 重命名sheet
    Excel.rename("重命名", "Sheet1")
    Excel.rename("Sheet1", sheetIndex=0)

    # 运行vba
    Excel.runVBA("test")
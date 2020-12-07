from PySide2.QtGui import QStandardItemModel, QStandardItem, QFont
from PySide2.QtWidgets import QApplication, QCheckBox
from PySide2.QtUiTools import QUiLoader
from PySide2.QtCore import QFile, Qt
import xlrd
import utils


class Fruit:
    def __init__(self):
        # 从文件中加载UI定义
        qfile_stats = QFile('fruit.ui')
        qfile_stats.open(QFile.ReadOnly)
        qfile_stats.close()

        # 从 UI 定义中动态 创建一个相应的窗口对象
        # 注意：里面的控件对象也成为窗口对象的属性了
        # 比如 self.ui.button , self.ui.textEdit
        self.ui = QUiLoader().load(qfile_stats)
        self.model = QStandardItemModel(4, 11)
        self.model2 = QStandardItemModel(8, 6)
        # 多选框
        self.check_box = QCheckBox(self.ui.tableView)
        # self.check_box.isChecked()
        self.model.setHorizontalHeaderLabels(['', 'name', 'id', 'Import', 'weight', 'price',
                                              'discount', 'shelfLife', 'sweetness', 'hardness', 'food'])
        self.model2.setHorizontalHeaderLabels(
            ['names', 'Euclidean', 'Manhattan', 'Cosine', 'disNode', 'PearsonCorrelation'])

        self.ui.tableView_2.setModel(self.model2)

        date = self.readData()
        values = date[0]
        nrows = date[1]
        ncols = date[2]

        # 增加多选框的字典，记录多选框和行号
        # 改动： 把多选框挪出一个for
        self.checks = {}
        c_NUM = 0
        for row in range(1, nrows):
            for column in range(ncols):
                item = QStandardItem('%s' % values[row][column])
                # 设置每个位置的文本值
                self.model.setItem(row - 1, column, item)

            # 添加多选框
            item_checked = QStandardItem()
            item_checked.setCheckState(Qt.Unchecked)
            item_checked.setCheckable(True)
            # 记录已经添加的多选框的位置
            self.checks[c_NUM] = [row - 1, 0]
            c_NUM += 1
            self.model.setItem(row - 1, 0, item_checked)

        self.ui.tableView.setModel(self.model)

        # 设置表格行高列宽
        self.ui.tableView.resizeRowsToContents()
        self.ui.tableView.setColumnWidth(0, 30)
        self.ui.tableView.setColumnWidth(8, 120)
        self.ui.tableView.setColumnWidth(9, 100)
        self.ui.tableView.setColumnWidth(10, 70)
        for column in range(1, ncols - 3):
            self.ui.tableView.setColumnWidth(column, 70)
        # tableView_2
        self.ui.tableView_2.resizeRowsToContents()
        self.ui.tableView_2.setColumnWidth(0, 190)
        self.ui.tableView_2.setColumnWidth(5, 120)
        for column in range(1, 5):
            self.ui.tableView_2.setColumnWidth(column, 80)

        self.ui.tableView.clicked.connect(self.m)

    def readData(self):
        # 读取表格数据
        book = xlrd.open_workbook('fruit.xlsx')
        sheet1 = book.sheets()[0]
        nrows = sheet1.nrows
        # print('表格总行数', nrows)
        ncols = sheet1.ncols
        # print('表格总列数', ncols)
        values = []
        for row in range(nrows):
            row_values = sheet1.row_values(row)
            values.append(row_values)
        return values, nrows, ncols

    def m(self):
        # 数组 记录被点击的行号
        num_list = []
        # 判断是否点击多选框
        for i in self.checks.values():
            # 一个个判断多选框选中状态
            x, y = i
            if self.model.item(x, y).checkState():
                # 记录行号
                num = int(self.model.item(x, y).row()) + 1
                num_list.append(num)
                # 如果点击了输出行号
                # print(num)
        data = self.readData()
        values = data[0]
        print(num_list)
        # count = self.model2.rowCount()

        for num in range(1, len(num_list)):
            self.model2.removeRow(num)

        for num in range(1, len(num_list)):
            name1 = values[num][1]
            name2 = values[num + 1][1]
            name = name1 + "-" + name2
            print(name)
            value1 = values[num]
            value2 = values[num + 1]
            fruit1 = utils.tra(value1, value2)
            # print(value1, value2)
            fruit2 = utils.tra2(value2)
            # print(f'{name1}={value1}={fruit1}')
            # print(f'{name2}={value2}={fruit2}')
            Euclidean = utils.Euclidean(fruit1, fruit2)
            Manhattan = utils.Manhattan(fruit1, fruit2)
            Cosine = utils.Cosine(fruit1, fruit2)
            disNode = utils.disNode(fruit1, fruit2)
            PearsonCorrelation = utils.PearsonCorrelation(fruit1, fruit2)
            # 距离值
            disValue = [name, Euclidean, Manhattan, Cosine, disNode, PearsonCorrelation]
            #  设置每个位置的文本值
            for col in range(6):
                item = QStandardItem('%s' % disValue[col])
                self.ui.tableView_2.resizeRowsToContents()
                self.model2.setItem(num - 1, col, item)


app = QApplication([])
fruit = Fruit()
fruit.ui.show()
app.exec_()

from PySide2.QtGui import QStandardItemModel, QStandardItem, QFont
from PySide2.QtWidgets import QApplication, QCheckBox
from PySide2.QtUiTools import QUiLoader
from PySide2.QtCore import QFile, Qt
import xlrd
import utils


class Fruit:
    def __init__(self):
        # Load UI definition from file
        qfile_stats = QFile('fruit.ui')
        qfile_stats.open(QFile.ReadOnly)
        qfile_stats.close()

        # Dynamically create a corresponding window object from the UI definition
        self.ui = QUiLoader().load(qfile_stats)
        self.model = QStandardItemModel(4, 11)
        self.model2 = QStandardItemModel(8, 6)
        # Checkbox
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

        # Add a dictionary for checkboxes, record checkboxes and line numbers
        self.checks = {}
        c_NUM = 0
        for row in range(1, nrows):
            for column in range(ncols):
                item = QStandardItem('%s' % values[row][column])
                # Set the text value of each position
                self.model.setItem(row - 1, column, item)

            # Add checkbox
            item_checked = QStandardItem()
            item_checked.setCheckState(Qt.Unchecked)
            item_checked.setCheckable(True)
            # Record the position of the added checkbox
            self.checks[c_NUM] = [row - 1, 0]
            c_NUM += 1
            self.model.setItem(row - 1, 0, item_checked)

        self.ui.tableView.setModel(self.model)

        # Set table row height and column width
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
        # Read table data
        book = xlrd.open_workbook('fruit.xlsx')
        sheet1 = book.sheets()[0]
        nrows = sheet1.nrows
        ncols = sheet1.ncols
        values = []
        for row in range(nrows):
            row_values = sheet1.row_values(row)
            values.append(row_values)
        return values, nrows, ncols

    def m(self):
        # rray record the row number that was clicked
        num_list = []
        # Determine whether to click the check box
        for i in self.checks.values():
            # Each one judges the checkbox selected state
            x, y = i
            if self.model.item(x, y).checkState():
                # Record line number
                num = int(self.model.item(x, y).row()) + 1
                num_list.append(num)
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
            # distance
            disValue = [name, Euclidean, Manhattan, Cosine, disNode, PearsonCorrelation]
            #  Set the text value of each position
            for col in range(6):
                item = QStandardItem('%s' % disValue[col])
                self.ui.tableView_2.resizeRowsToContents()
                self.model2.setItem(num - 1, col, item)


app = QApplication([])
fruit = Fruit()
fruit.ui.show()
app.exec_()

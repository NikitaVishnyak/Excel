from PyQt5 import QtCore, QtGui, QtWidgets
import excel


class Cell:
    def __init__(self, row, col, tableWidget):
        self.row = row
        self.col = col
        self.tableWidget = tableWidget
        self.msg = excel.MessageBox()

    def getExpressionFromCell(self, row, col):
        return self.tableWidget.item(row, col).text()

    def fillCell(self, expr):
        self.tableWidget.setItem(self.row, self.col, QtWidgets.QTableWidgetItem(expr))

    def parsing(self, expr):
        self.parser = excel.Parser(expr, self.tableWidget)
        if expr[0] == '=':
            self.parsingInCell(expr)
        elif expr[0] == '#':
            self.replacementParsing(expr)
        else:
            self.parsingInLine(expr)

    def parsingInCell(self, expr):
        self.cellCalculation()

    def parsingInLine(self, expr):
        self.lineCalculation()

    def replacementParsing(self, expr):
        res = self.parser.replacement()
        if res is not None:
            self.fillCell(str(res))
        elif res == '0':
            self.fillCell(str(expr))
        else:
            self.msg.incorrectExpression()

    def lineCalculation(self):
        res = self.parser.calculationFromLine()
        if res is not None:
            self.fillCell(str(res))
        else:
            self.msg.incorrectExpression()

    def cellCalculation(self):
        res = self.parser.calculationFromCell()
        if res is not None and res is not False:
            self.fillCell(str(res))
        elif res is False:
            self.fillCell('')
        else:
            self.msg.incorrectExpression()
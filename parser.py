from PyQt5 import QtCore, QtGui, QtWidgets
import excel
import re
import operator


class Parser(excel.Cell):
    def __init__(self, expr, tableWidget):
        self.expression = expr
        self.tableWidget = tableWidget
        self.msg = excel.MessageBox()
        self.op = {'+': lambda x, y: x + y,
                   '-': lambda x, y: x - y,
                   '*': lambda x, y: x * y,
                   '/': lambda x, y: x / y if (y != 0) else self.msg.zeroDeviding(),
                   '^': lambda x, y: x ** y,
                   '%': lambda x, y: x % y if (y != 0) else self.msg.zeroDeviding()}

    def calculationFromLine(self):
        pattern = re.compile('^(\-\d*\.?\d*|\d*\.?\d*)\s?(\+|\-|\*|\/|\^)\s?(\-\d*\.?\d*|\d*\.?\d*)$')
        selection = re.finditer(pattern, self.expression)
        operands = None
        if selection:
            for element in selection:
                operands = element.group(1, 2, 3)
                if operands is not None:
                    return self.op[operands[1]](float(operands[0]), float(operands[2]))
                else:
                    return None

    def calculationFromCell(self):
        pattern = re.compile('^\=(\-?)([A-Z]*)(\-\d*\.?\d*|\d*\.?\d*)\s?(\+|\-|\*|\/|\^)\s?([A-Z]*)(\-\d*\.?\d*|\d*\.?\d*)$')
        selection = re.finditer(pattern, self.expression)

        first = True
        second = True
        parts = [0, 0]

        try:
            if selection:
                for element in selection:
                    operands = element.group(1, 2, 3, 4, 5, 6)
                    if operands is not None:
                        for k in range(self.tableWidget.columnCount()):
                            if operands[1] == self.tableWidget.horizontalHeaderItem(k).text() and first:
                                if self.tableWidget.item(int(operands[2]) - 1, k).text() != '':
                                    parts[0] = self.tableWidget.item(int(operands[2]) - 1, k).text()
                                    first = False
                                else:
                                    parts[0] = 0
                                    first = False
                            elif operands[4] == self.tableWidget.horizontalHeaderItem(k).text() and second:
                                if self.tableWidget.item(int(operands[5]) - 1, k).text() != '':
                                    parts[1] = self.tableWidget.item(int(operands[5]) - 1, k).text()
                                    second = False
                                else:
                                    parts[1] = 0
                                    second = False

                            if operands[1] == '' and first:
                                parts[0] = operands[2]
                                first = False
                            if operands[4] == '' and second:
                                parts[1] = operands[5]
                                second = False

                        if operands[0] == '-':
                            parts[0] = float(parts[0]) * -1
                        return self.op[operands[3]](float(parts[0]), float(parts[1]))
                    else:
                        return None
        except:
            self.msg.wrongEntry()
            return False

    def replacement(self):
        pattern = re.compile('^\#([A-Z]+)(\d+)$')
        selection = re.finditer(pattern, self.expression)
        try:
            if selection:
                for element in selection:
                    operands = element.group(1, 2)
                    if operands is not None:
                        for k in range(self.tableWidget.columnCount()):
                            if operands[0] == self.tableWidget.horizontalHeaderItem(k).text():
                                if self.tableWidget.item(int(operands[1]) - 1, k).text() != '':
                                    return self.tableWidget.item(int(operands[1]) - 1, k).text()
                                else:
                                    return 0
                    else:
                        return None
        except:
            self.msg.wrongEntry()
            return False
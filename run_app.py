from PyQt5 import QtWidgets
from excel import mainWindow

import sys

if __name__ == "__main__":
    application = QtWidgets.QApplication(sys.argv)
    app = mainWindow()
    app.show()
    sys.exit(application.exec_())
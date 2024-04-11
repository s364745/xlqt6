import sys

from PyQt6 import QtWidgets as qtw
from PyQt6 import QtCore as qtc
from PyQt6 import QtGui as qtq
from PyQt6 import uic

# Files
import xlcontrol as xl

# Import from ui file
gui, baseClass = uic.loadUiType('gui.ui')


# Argument can be baseClass
class MainWindow(baseClass):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        # Set up gui
        self.ui = gui()
        self.ui.setupUi(self)

        # Code starts here
        self.ui.min_knapp.clicked.connect(self.get_data)

        self.ui.addTabButton.clicked.connect(self.generate_tab)

        self.ui.removeTabButton.clicked.connect(self.remove_tab)
        # Code stops here

    # Functions
    def generate_tab(self):
        self.ui.myTabs.addTab(qtw.QWidget(), 'Generated tab')

    def remove_tab(self):
        self.ui.myTabs.removeTab(0)

    def get_data(self):
        inData = self.ui.reqXl.text()
        print(inData)
        outData = xl.ws[inData].value
        print(outData)
        self.ui.min_text.setText(outData)


if __name__ == '__main__':
    app = qtw.QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())

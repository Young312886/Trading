import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QAxContainer import *

class MyWindow(QMainWindow):
    def __init__ (self):
        super().__init__()
        self.setWindowTitle("Stock")
        self.setGeometry(300, 300, 300, 400)

        self.kiwom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")

        btn1 = QPushButton("Login", self)
        btn1.move(20, 20)
        btn1.clicked.connect(self.btn2_clicked)

    def btn1_clicked(self):
        ret = self.kiwoom.dynamicCal("CommConnect()")

    def btn2_clicked(self):
        if self.kiwoom.dynamicCal("GetConnectState()") == 0:
            self.statusBar().showMessage("Not Connected")
        else:
            self.statusBar().showMessage("Connected")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    mywindow = MyWindow()
    mywindow.show()
    app.exec_()


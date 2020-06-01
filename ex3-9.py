import sys
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtCore import QDate, Qt

class MyApp(QMainWindow):

    def __init__(self):
        super().__init__()
        self.mydate = QDate.currentDate()
        self.initUI()

    def initUI(self):
        self.statusBar().showMessage(self.mydate.toString(Qt.DefaultLocaleLongDate))

        self.setWindowTitle('Date')
        self.setGeometry(300, 300, 300, 200)
        self.show()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())


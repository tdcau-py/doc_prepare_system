from PyQt5 import QtWidgets, QtCore, QtGui, uic


class SettingsWin(QtWidgets.QMainWindow):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        gui, base = uic.loadUiType('forms/settings_form.ui')
        self.ui = gui()
        self.ui.setupUi(self)
        self.setWindowTitle('Сервер системы формирования документов')


if __name__ == '__main__':
    import sys

    app = QtWidgets.QApplication(sys.argv)
    win = SettingsWin()
    win.show()
    sys.exit(app.exec_())

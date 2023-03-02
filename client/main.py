from PyQt5 import QtWidgets, QtCore, QtGui, uic
from PyQt5 import sip
import openpyxl
from datetime import datetime
from modules.notifications_creator import FormWindow
from modules.receipt_generator import FormReceipt
from modules.inn_search import InnSearchForm


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        gui, base = uic.loadUiType('forms/main_form.ui')
        self.ui = gui()
        self.ui.setupUi(self)
        self.setWindowTitle('Система формирования документов')

        with open('static/main_win/style.css', 'r') as file:
            style = file.read()

        self.setStyleSheet(style)

        year = datetime.today().strftime('%Y')
        self.ui.lbl_copy.setText(f'<html><body><address>&copy; {year}</address></body></html>')

        self.ui.about.triggered.connect(self.about)

        self.ui.btn_search_inn.clicked.connect(self.run_inn_search)
        self.ui.btn_notification.clicked.connect(self.run_notification_creator)
        self.ui.btn_receipt.clicked.connect(self.run_receipt_generator)

    def run_inn_search(self):
        self.form_window = InnSearchForm()
        self.form_window.exec_()

    def run_notification_creator(self):
        self.form_window = FormWindow()
        self.form_window.exec_()

    def run_receipt_generator(self):
        self.form_window = FormReceipt()
        self.form_window.exec_()

    def about(self):
        year_today = datetime.today().strftime('%Y')
        text_message = f"""
        <html>
            <body>
                <table>
                    <tr>
                        <td>Название:</td>
                        <td>Система формирования документов</td>
                    </tr>

                    <br>

                    <tr>
                        <td>Автор:</td>
                        <td>
                            <p>
                               ведущий инженер-программист <br>Филиала ГБУ "МФЦ Владимирской области" в г. Киржач<br>
                               Борисов Павел Валерьевич
                            </p>
                        </td>
                    </tr>
                </table> 
            </body>
        </html>"""

        QtWidgets.QMessageBox.about(self, 'О программе', text_message)


if __name__ == '__main__':
    import sys
    app = QtWidgets.QApplication(sys.argv)
    app.setApplicationName('Система формирования документов')
    ico = QtGui.QIcon('static/image/logo.png')
    app.setWindowIcon(ico)
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())

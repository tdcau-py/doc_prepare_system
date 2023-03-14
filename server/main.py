from PyQt5 import QtWidgets, uic
from modules.searching_inn import search_inn
import configparser
import socket
import json


class SettingsWin(QtWidgets.QMainWindow):
    CONFIG_FILE = 'settings.ini'

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        gui, base = uic.loadUiType('forms/settings_form.ui')
        self.ui = gui()
        self.ui.setupUi(self)
        self.setWindowTitle('Сервер системы формирования документов')

        config = configparser.ConfigParser()
        config.read(self.CONFIG_FILE, encoding='utf8')
        self.HOST = config['default']['host']
        self.PORT = config['default']['port']

    def client_connections(self):
        """Соединение с клиентом и обмен данными"""
        while True:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as server_socket:
                server_socket.bind((self.HOST, self.PORT))
                server_socket.listen(1)
                connect, address = server_socket.accept()

                with connect:
                    print(f'Connected by {address}')

                    while True:
                        data = connect.recv(1024)
                        if not data: break
                        decode_data = json.loads(data.decode())
                        result = search_inn(decode_data)
                        connect.send(result.encode('utf8'))


if __name__ == '__main__':
    import sys

    app = QtWidgets.QApplication(sys.argv)
    win = SettingsWin()
    win.show()
    sys.exit(app.exec_())

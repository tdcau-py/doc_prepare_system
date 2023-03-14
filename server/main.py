from PyQt5 import QtWidgets, QtCore, uic
from modules.searching_inn import search_inn
import configparser
import socket
import json


class ClientConnections(QtCore.QThread):
    CONFIG_FILE = 'settings.ini'
    config = configparser.ConfigParser()
    config.read(CONFIG_FILE, encoding='utf8')
    HOST = config['default']['host']
    PORT = int(config['default']['port'])

    server_status = QtCore.pyqtSignal(str)

    def __init__(self, parent = None) -> None:
        super().__init__(parent)
        self.running = False

    def run(self):
        """Соединение с клиентом и обмен данными"""
        # self.server_status.emit('Сервер работает.')
        self.running = True

        while self.running:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as server_socket:
                server_socket.bind((self.HOST, self.PORT))
                server_socket.listen(1)
                connect, address = server_socket.accept()

                with connect:
                    # print(f'Connected by {address}')

                    while True:
                        data = connect.recv(1024)
                        if not data: break
                        decode_data = json.loads(data.decode())
                        result = search_inn(decode_data)
                        connect.send(result.encode('utf8'))


class SettingsWin(QtWidgets.QMainWindow):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        gui, base = uic.loadUiType('forms/settings_form.ui')
        self.ui = gui()
        self.ui.setupUi(self)
        self.setWindowTitle('Сервер системы формирования документов')

        self.ui.serverStatus.setText('Сервер остановлен.')
        self.connections_thread = ClientConnections()

        self.ui.btnStopServer.setDisabled(True)
        self.ui.btnStartServer.clicked.connect(self.start_server)
        self.ui.btnStopServer.clicked.connect(self.stop_server)

        self.connections_thread.started.connect(self.server_started)
        self.connections_thread.finished.connect(self.server_stoped)

    def start_server(self):
        """Соединение с клиентом и обмен данными"""
        self.ui.btnStartServer.setDisabled(True)
        self.ui.btnStopServer.setEnabled(True)
        self.connections_thread.start()

    def server_started(self):
        self.ui.serverStatus.setText('Сервер работает.')

    def stop_server(self):
        """Остановка потока, обрыв соединения с клиентами"""
        self.ui.btnStartServer.setEnabled(True)
        self.ui.btnStopServer.setDisabled(True)
        self.connections_thread.running = False

    def server_stoped(self):
        self.ui.serverStatus.setText('Сервер остановлен.')


if __name__ == '__main__':
    import sys

    app = QtWidgets.QApplication(sys.argv)
    win = SettingsWin()
    win.show()
    sys.exit(app.exec_())

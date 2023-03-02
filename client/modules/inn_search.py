from PyQt5 import QtWidgets, QtCore, uic
import configparser
import socket
import json

config_file = 'app_settings.ini'
config = configparser.ConfigParser()


def send_to_server(data: bytes) -> str:
	"""Соединяется с сервером, передает ему данные и получает ответ"""
	config.read(config_file, encoding='utf8')
	HOST = config['default']['host']
	PORT = int(config['default']['port'])

	with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as client_socket:
		client_socket.connect((HOST, PORT))
		client_socket.send(data)
		result = client_socket.recv(1024)

	return result


class InnSearchForm(QtWidgets.QDialog):
	DOCS_TYPE = [
		'',
		'01 - Паспорт гражданина СССР',
		'03 - Свидетельство о рождении',
		'10 - Паспорт иностранного гражданина',
		'12 - Вид на жительство в Российской Федерации',
		'15 - Разрешение на временное проживание в Российской Федерации',
		'19 - Свидетельство о предоставлении временного убежища на территории Российской Федерации',
		'21 - Паспорт гражданина Российской Федерации',
		'23 - Свидетельство о рождении, выданное уполномоченным органом иностранного государства',
		'62 - Вид на жительство иностранного гражданина',
	]

	def __init__(self):
		QtWidgets.QDialog.__init__(self)
		gui, base = uic.loadUiType('forms/inn_search_form.ui')
		self.ui = gui()
		self.ui.setupUi(self)

		with open('static/inn_search_form/style.css', 'r') as file:
			style = file.read()

		self.setStyleSheet(style)
		self.setWindowTitle('Поиск ИНН')

		self.ui.last_name.setFocus()
		self.ui.result_widget.setHidden(True)

		self.ui.doc_type.addItems(self.DOCS_TYPE)
		self.ui.doc_type.setCurrentText('21 - Паспорт гражданина Российской Федерации')
		self.ui.doc_type.currentIndexChanged.connect(self.set_docnumber_mask)

		self.ui.check_otch.stateChanged.connect(self.middle_name_disabled)

		self.set_docnumber_mask()

		self.ui.btn_send.clicked.connect(self.on_send)
		self.ui.btn_clear.clicked.connect(self.on_clear)

	def set_docnumber_mask(self):
		self.ui.doc_num.setText('')
		doc_type = self.ui.doc_type.currentText()

		if doc_type == '21 - Паспорт гражданина Российской Федерации':
			self.ui.doc_num.setInputMask('99 99 9999990;_')

		elif doc_type == '01 - Паспорт гражданина СССР':
			self.ui.doc_num.setInputMask('>AAAAAAAAAA-AA 999999;_')

		elif doc_type == '03 - Свидетельство о рождении':
			self.ui.doc_num.setInputMask('>AAAAAAAAAA-AA 999999;_')
		
		else:
			self.ui.doc_num.setInputMask('')

	def on_send(self):
		data = {
			'c': 'find',
			'captcha': '',
			'captchaToken': '',
		}

		fam = self.ui.last_name.text()
		nam = self.ui.first_name.text()
		bdate = self.ui.date_birth.text()
		doctype = self.ui.doc_type.currentText()[:2]
		docno = self.ui.doc_num.text()
		docdt = self.ui.doc_date.text()

		if not self.ui.check_otch.isChecked():
			otch = self.ui.middle_name.text()
		else:
			otch = ''

		if fam and nam and otch and bdate and doctype and docno:
			data.update([('fam', fam), ('nam', nam), ('otch', otch), ('bdate', bdate), ('doctype', doctype), ('docno', docno)])

			if ''.join(i for i in docdt.split('.')).isdigit():
				data.update([('docdt', docdt)])
			else:
				data.update([('docdt', '')])
			
			encode_data = json.dumps(data).encode('utf8')

			try:
				result = send_to_server(encode_data)

			except ConnectionRefusedError:
				return QtWidgets.QMessageBox.critical(self, 'Ошибка подключения', 'Сервер недоступен')

			except TimeoutError:
				return QtWidgets.QMessageBox.critical(self, 'Тайм-аут', 'Нет соединения с сервером')

			self.ui.result_widget.setVisible(True)
			return self.ui.result_inn.setText(result.decode())

		return QtWidgets.QMessageBox.critical(self, 'Не заполнены поля', 'Заполните обязательные поля - *')

	def middle_name_disabled(self):
		if self.ui.check_otch.isChecked():
			return self.ui.middle_name.setDisabled(True)
		
		return self.ui.middle_name.setEnabled(True)

	def on_clear(self):
		self.ui.last_name.setText('')
		self.ui.first_name.setText('')
		self.ui.middle_name.setText('')
		self.ui.date_birth.setText('')
		self.ui.doc_type.setCurrentText('21 - Паспорт гражданина Российской Федерации')
		self.ui.doc_num.setText('')
		self.ui.doc_date.setText('')
		self.ui.result_inn.setText('')
		self.ui.result_widget.setHidden(True)

		if self.ui.check_otch.isChecked():
			self.ui.check_otch.setCheckState(QtCore.Qt.Unchecked)

		self.ui.last_name.setFocus()


if __name__ == '__main__':
	import sys
	app = QtWidgets.QApplication(sys.argv)
	win = InnSearchForm()
	win.show()
	sys.exit(app.exec_())

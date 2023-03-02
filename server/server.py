import socket
import pip._vendor.requests as requests
import json


def search_inn(data: dict) -> str:
	"""Производит поиск ИНН"""
	url_inn_do = 'https://service.nalog.ru/inn-new-proc.do'
	url_inn_json = 'https://service.nalog.ru/inn-new-proc.json'

	response = requests.post(url_inn_do, data)
	request_id = response.json()['requestId']

	data_to_json = {
		'c': 'get',
		'requestId': request_id,
	}

	inn_resp = requests.post(url_inn_json, data_to_json)

	try:
		result = inn_resp.json()['inn']

	except KeyError:
		return 'ИНН не найден. Проверьте правильность введенных данных.'

	return f'ИНН: {result}'


if __name__ == '__main__':
	import configparser

	config_file = 'app_settings.ini'
	config = configparser.ConfigParser()
	config.read()

	HOST = config['default']['host']
	PORT = config['default']['port']

	while True:
		with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as server_socket:
			server_socket.bind((HOST, PORT))
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

from PyQt5 import QtWidgets, QtCore, QtGui, uic
from docxtpl import DocxTemplate, InlineImage
import qrcode, configparser, datetime
from docx.shared import Mm
import os

config_file = 'config.ini'
config = configparser.ConfigParser()


def qr_generate(name, personal_acc, bank_name, bic, corresp_acc, **kwargs):
    if corresp_acc == '':
        corresp_acc = 0

    coding_pay = f'ST00012|Name={name}|PersonalAcc={personal_acc}|BankName={bank_name}|BIC={bic}|CorrespAcc={corresp_acc}'
    
    for k in kwargs:
        coding_pay += f'|{k}={kwargs[k]}'
    img = qrcode.make(coding_pay)

    return img


class FormReceipt(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super(FormReceipt, self).__init__(parent)
        gui, base = uic.loadUiType('forms/receipt_form.ui')
        self.ui = gui()
        self.ui.setupUi(self)
        self.setWindowFlags(QtCore.Qt.Window)

        with open('static/style_receipt/style.css', 'r') as file:
            style = file.read()

        self.setStyleSheet(style)
        self.setWindowTitle('Генератор квитанций')

        self.services = self.ui.comboBox
        self.name = self.ui.name
        self.personal_acc = self.ui.personalAcc
        self.bank_name = self.ui.bankName
        self.bic = self.ui.bic
        self.corresp_acc = self.ui.correspAcc
        self.purpose = self.ui.purpose
        self.payee_inn = self.ui.payeeInn
        self.kpp = self.ui.kpp
        self.cbc = self.ui.cbc
        self.oktmo = self.ui.oktmo
        self.uin = self.ui.uin
        self.payer_fio = self.ui.fioPayer
        self.address_payer = self.ui.addressPayer
        self.pers_acc = self.ui.persAcc
        self.reg_type = self.ui.regType
        self.sum = self.ui.sum
        self.payer_id_num = self.ui.payerIdNum
        self.payer_id_type = self.ui.payerIdType

        ID_TYPES = ['',
                    'ПАСПОРТ РФ',
                    'СВИД О РОЖДЕНИИ',
                    'ИНН',
                    'ВОДИТ УДОСТОВЕРЕНИЕ',
                    'ПАСПОРТ МОРЯКА',
                    'ВОЕННЫЙ БИЛЕТ',
                    'ВРЕМЕН УДОСТОВЕР',
                    'ПАСПОРТ ИН ГРАЖД',
                    'ВИД НА ЖИТЕЛЬСТВО',
                    'УДОСТОВЕР БЕЖЕНЦА',
                    'МИГРАЦИОННАЯ КАРТА',
                    'СНИЛС',
                    'ЗАГРАНПАСПОРТ']

        self.payer_id_type.addItems(ID_TYPES)

        self.btn_gen = self.ui.form
        self.btn_reset = self.ui.reset
        self.btn_add_pattern = self.ui.add_pattern

        self.btn_add_pattern.setDisabled(True)
        self.name.textChanged.connect(self.on__btn_enable)

        config.read(config_file, encoding='utf8')

        self.model()
        self.services.activated.connect(self.on_fill)

        self.btn_gen.clicked.connect(self.check_required_fields)
        self.btn_add_pattern.clicked.connect(self.add_pattern)
        self.btn_reset.clicked.connect(self.on_clear)

    def model(self):
        purposes_pay = config.sections()
        model_purposes = QtGui.QStandardItemModel(self)
        model_purposes.appendRow(QtGui.QStandardItem('Выберите назначение платежа...'))
        for purpose in purposes_pay:
            item = QtGui.QStandardItem(purpose)
            model_purposes.appendRow(item)

        self.services.setModel(model_purposes)

    def on_fill(self):
        service_index = self.services.currentIndex()
        if service_index == 0:
            self.on_clear()
        else:
            service = self.services.currentText()
            self.name.setText(config[service]['name'])
            self.personal_acc.setText(config[service]['personalacc'])
            self.bank_name.setText(config[service]['bankname'])
            self.bic.setText(config[service]['bic'])
            self.corresp_acc.setText(config[service]['correspacc'])
            self.purpose.setText(config[service]['purpose'])
            self.payee_inn.setText(config[service]['payeeinn'])
            self.kpp.setText(config[service]['kpp'])
            self.cbc.setText(config[service]['cbc'])
            self.oktmo.setText(config[service]['oktmo'])

            try:
                self.reg_type.setText(config[service]['RegType'])
            except KeyError:
                self.reg_type.setText('')

            try:
                self.sum.setText(config[service]['SUM'])
            except KeyError:
                self.sum.setText('')

    def check_required_fields(self):
        if (self.name.text() == '' or 
            self.personal_acc.text() == '' or 
            self.bank_name.text() == '' or 
            self.bic.text() == '' or 
            self.corresp_acc == ''):
            QtWidgets.QMessageBox.information(self, 'Предупреждение', 'Заполните обязательные поля - *')
            
        else:
            self.pay_receipt_generate()

    def pay_receipt_generate(self):
        name = self.name.text()
        personal_acc = self.personal_acc.text()
        bank_name = self.bank_name.text()
        bic = self.bic.text()
        corresp_acc = self.corresp_acc.text()

        additional_det = {}

        kpp = self.kpp.text()
        payee_inn = self.payee_inn.text()
        oktmo = self.oktmo.text()
        cbc = self.cbc.text()
        uin = self.uin.text()
        purpose = self.purpose.text()
        payer_fio = self.payer_fio.text()
        address_payer = self.address_payer.text()
        sum_pay = self.sum.text()
        payer_id_type = self.payer_id_type.currentText()
        payer_id_num = self.payer_id_num.text()

        if sum_pay:
            sum_pay = float(sum_pay).__format__('.2f')
            additional_det['SUM'] = float(sum_pay) * 100
            sum_r = sum_pay.split('.')[0]
            sum_k = sum_pay.split('.')[1]
        else:
            sum_r = ''
            sum_k = ''

        if purpose:
            additional_det['Purpose'] = purpose

        if payee_inn:
            additional_det['PayeeINN'] = payee_inn

        if kpp:
            additional_det['KPP'] = kpp

        if cbc:
            additional_det['CBC'] = cbc

        if oktmo:
            additional_det['OKTMO'] = oktmo

        if payer_id_type:
            additional_det['PayerIdType'] = payer_id_type

        if payer_id_num:
            additional_det['PayerIdNum'] = payer_id_num

        if payer_fio:
            if len(payer_fio.split()) < 2:
                return QtWidgets.QMessageBox.critical(self, 'Ошибка реквизитов',
                    'Неверно указаны реквизиты ФИО.\nПожалуйста, введите фамилию, имя, отчество (при наличии) полностью.')

            else:
                fio = payer_fio.split()

                if len(fio) == 2:
                    additional_det['LastName'] = fio[0]
                    additional_det['FirstName'] = fio[1]

                else:
                    additional_det['LastName'] = fio[0]
                    additional_det['FirstName'] = fio[1]
                    additional_det['MiddleName'] = fio[2]

        if address_payer:
            additional_det['PayerAddress'] = address_payer

        if uin:
            additional_det['UIN'] = uin

        pay_receipt = DocxTemplate('templates/templ_pay_receipt.docx')
        
        context = {
            'name': name,
            'kpp': kpp,
            'inn': payee_inn,
            'oktmo': oktmo,
            'personalAcc': personal_acc,
            'bankName': bank_name,
            'bic': bic,
            'correspAcc': corresp_acc,
            'cbc': cbc,
            'uin': uin,
            'purpose': purpose,
            'sum_r': sum_r,
            'sum_k': sum_k,
            'fio': payer_fio,
            'address': address_payer,
            'items': [],
        }

        if self.reg_type.text():
            context['items'].append({'bank_detail': 'Код НО администратора платежа',
                                     'value': self.reg_type.text()})

        if self.payer_id_type.currentText():
            context['items'].append({'bank_detail': 'Вид документа, удостоверяющего личность',
                                    'value': self.payer_id_type.currentText()})

            context['items'].append({'bank_detail': 'Серия и номер документа',
                                    'value': self.payer_id_num.text()})

        qr_code = qr_generate(name=name, personal_acc=personal_acc, bank_name=bank_name, bic=bic, corresp_acc=corresp_acc, **additional_det)
        qr_code.save('qr.png')
        qr_img = InlineImage(pay_receipt, image_descriptor='qr.png', width=Mm(40), height=Mm(40))
        context['qr_img'] = qr_img

        pay_receipt.render(context)

        home_dir = os.path.expanduser('~')
        folder_name = 'квитанции'
        file_name = f'Квитанция_{purpose}'
        another_file_name = ''

        if os.path.exists(home_dir + '\\Desktop'):
            desktop_dir = home_dir + '\\Desktop'
            if not self.payer_fio.text():
                try:
                    pay_receipt.save(f'{desktop_dir}\\{folder_name}\\{file_name}.docx')
                except FileNotFoundError:
                    os.mkdir(f'{desktop_dir}\\{folder_name}')
                    pay_receipt.save(f'{desktop_dir}\\{folder_name}\\{file_name}.docx')

            else:
                another_file_name = f'{self.payer_fio.text().split()[0].upper()} {self.payer_fio.text().split()[1][0].upper()}_{purpose}'
                try:
                    pay_receipt.save(f'{desktop_dir}\\{folder_name}\\{another_file_name}.docx')
                except FileNotFoundError:
                    os.mkdir(f'{desktop_dir}\\{folder_name}')
                    pay_receipt.save(f'{desktop_dir}\\{folder_name}\\{another_file_name}.docx')

        elif os.path.exists(home_dir + '\\Рабочий стол'):
            desktop_dir = home_dir + '\\Рабочий стол'
            if not self.payer_fio.text():
                try:
                    pay_receipt.save(f'{desktop_dir}\\{folder_name}\\{file_name}.docx')
                except FileNotFoundError:
                    os.mkdir(f'{desktop_dir}\\{folder_name}')
                    pay_receipt.save(f'{desktop_dir}\\{folder_name}\\{file_name}.docx')

            else:
                another_file_name = f'{self.payer_fio.text().split()[0].upper()} {self.payer_fio.text().split()[1][0].upper()}_{purpose}'
                try:
                    pay_receipt.save(f'{desktop_dir}\\{folder_name}\\{another_file_name}.docx')
                except FileNotFoundError:
                    os.mkdir(f'{desktop_dir}\\{folder_name}')
                    pay_receipt.save(f'{desktop_dir}\\{folder_name}\\{another_file_name}.docx')

        os.system('del qr.png')

        if another_file_name:
            self.open_file(f'{desktop_dir}\\{folder_name}\\{another_file_name}.docx')
        else:
            self.open_file(f'{desktop_dir}\\{folder_name}\\{file_name}.docx')

    def open_file(self, doc_file: str) -> None:
        os.startfile(doc_file)

    def add_pattern(self):
        bank_details = {}

        # изменение реквизитов при существующем шаблоне платежа
        # if self.purpose.text() and self.purpose.text() in config.sections():
        #     message = QtWidgets.QMessageBox.question(self, 'Добавление шаблона', 
        #         'Шаблон с таким названием уже существует.\nВнести изменения в реквизиты?')

        #     if message == QtWidgets.QMessageBox.StandardButtos.No:
        #         return self.model()

        #     else:
        #         for k, v in config[self.purpose.text()].items():

        # else:
        #     ...

        bank_details['Name'] = self.name.text()
        bank_details['PersonalAcc'] = self.personal_acc.text()
        bank_details['BankName'] = self.bank_name.text()
        bank_details['BIC'] = self.bic.text()
        bank_details['CorrespAcc'] = self.corresp_acc.text()

        if self.kpp.text():
            bank_details['KPP'] = self.kpp.text()
        if self.payee_inn.text():
            bank_details['PayeeINN'] = self.payee_inn.text()
        if self.oktmo.text():
            bank_details['OKTMO'] = self.oktmo.text()
        if self.cbc.text():
            bank_details['CBC'] = self.cbc.text()
        if self.purpose.text():
            bank_details['Purpose'] = self.purpose.text()
        if self.sum.text():
            bank_details['SUM'] = self.sum.text()
        if self.reg_type.text():
            bank_details['RegType'] = self.reg_type.text()

        config[self.purpose.text()] = bank_details

        with open(config_file, mode='w', encoding='utf8') as configfile:
            config.write(configfile)

        return self.model()

    def on_clear(self):
        self.services.setCurrentIndex(0)
        self.name.setText('')
        self.personal_acc.setText('')
        self.bank_name.setText('')
        self.bic.setText('')
        self.corresp_acc.setText('')
        self.purpose.setText('')
        self.payee_inn.setText('')
        self.kpp.setText('')
        self.cbc.setText('')
        self.oktmo.setText('')
        self.uin.setText('')
        self.payer_fio.setText('')
        self.address_payer.setText('')
        self.pers_acc.setText('')
        self.reg_type.setText('')
        self.sum.setText('')
        self.payer_id_type.setCurrentIndex(0)
        self.payer_id_num.setText('')

    def on__btn_enable(self):
        self.btn_add_pattern.setEnabled(True)


if __name__ == '__main__':
    import sys

    app = QtWidgets.QApplication(sys.argv)
    win = FormReceipt()
    win.show()
    sys.exit(app.exec_())

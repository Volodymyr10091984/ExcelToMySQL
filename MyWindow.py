from PyQt5 import uic
from PyQt5.QtWidgets import QDialog, QFileDialog, QMessageBox
from PyQt5.QtCore import Qt, pyqtSlot
from PyQt5.QtGui import QIcon
from ConvertFile import ConvertFile


# class window
class MyWindow(QDialog):
    def __init__(self, parent=None):
        QDialog.__init__(self, parent)

        # download the .ui file
        uic.loadUi('WindowConvert.ui', self)
        # setting the options, title and window icon
        self.setWindowFlag(Qt.WindowMinimizeButtonHint)
        self.setWindowTitle('Конвертер Excel to MySQL')
        self.setWindowIcon(QIcon('MyIcon.png'))

        self.base_name = ''
        self.user_name = ''
        self.password_base = ''
        self.host_base = ''
        self.table_name = ''
        self.work_shit = ''
        self.ExcelFile = None

        # creation slots for processing data
        self.ChangeExelFile.clicked.connect(self.change_file)
        self.ButtonNewTable.clicked.connect(lambda: self.create_new_table(False))
        self.ButtonEditTable.clicked.connect(lambda: self.create_new_table(True))

    # Excel file selection slot
    @pyqtSlot()
    def change_file(self):
        self.ExcelFile = QFileDialog.getOpenFileUrl(caption='Виберіть файл', filter='All (*);;Exes (*.xls *.xlsx)',
                                                    initialFilter='Exes (*.xls *.xlsx)')
        # display the selected file in the window
        self.VievExelFile.setText(self.ExcelFile[0].toLocalFile())

    # a slot for writing data from Excel to MySQL
    def create_new_table(self, flag):
        # function getting data from the window fields
        self.get_data()
        # checking that all fields are filled
        result_label = self.info_label()

        if result_label:
            # if passed data, call the class to write to the database
            table = ConvertFile(self.base_name, self.user_name, self.host_base, self.table_name,
                                self.ExcelFile[0].toLocalFile(), self.work_shit, self.password_base)
            result = table.edit_table(replace=flag)
            if result:
                # if the recording fails, an error message is displayed
                QMessageBox(QMessageBox.Information, 'OK', 'Запис проведено успішно', QMessageBox.Ok).exec()
            else:
                # if the recording is successful, a successful completion message is displayed
                QMessageBox(QMessageBox.Critical,
                            'Помилка', 'Сталась помилка при записі данних. Перевірте параметри під\'єднання',
                            QMessageBox.Ok).exec()

    # method for outputting error messages
    @staticmethod
    def dialog_error(data: str):
        return QMessageBox(QMessageBox.Critical, 'Увага', data, QMessageBox.Ok)

    # function to retrieve data from window fields
    def get_data(self):
        self.base_name = self.BaseName.text()
        self.user_name = self.User.text()
        self.password_base = self.PasswordBase.text()
        self.host_base = self.host.text()
        self.table_name = self.Table.text()
        self.work_shit = self.workShit.text()

    # function for displaying data error messages
    def info_label(self):
        if len(self.base_name) == 0:
            self.dialog_error('Ви не ввели назву бази данних').exec()
            return False
        elif len(self.user_name) == 0:
            self.dialog_error('Ви не ввели ім\'я користувача').exec()
            return False
        elif len(self.host_base) == 0:
            self.dialog_error('Ви не ввели host бази данних').exec()
            return False
        elif len(self.table_name) == 0:
            self.dialog_error('Ви не ввели ім\'я таблиці данних').exec()
            return False
        elif self.ExcelFile is None:
            self.dialog_error('Ви не вибрали файл Excel').exec()
            return False
        elif len(self.work_shit) == 0:
            self.dialog_error('Ви не вибрали робочий лист Excel').exec()
            return False
        else:
            return True

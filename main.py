from PySide6 import QtWidgets, QtGui
import pandas as pd


class CPP_MakeSender(QtWidgets.QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.spisokFileName = 'spisok.xlsx'
        self.dataFileName = 'data.xlsx'
        self.listNames = 'Управление'

        self.initUi()
        self.initSignals()

        self.initFiles()

        self.comboBoxListNamesSet()

        self.shortcut_open = QtGui.QShortcut(QtGui.QKeySequence('Ctrl+c'), self)
        self.shortcut_open.activated.connect(self.onOpen)

    def initUi(self):
        self.setWindowTitle("CPP_MakeSender")

        self.setMinimumSize(900, 600)

        # result table ------------------------------------------------------------------------------------------------
        self.resultTable = QtWidgets.QTableWidget(0, 5)
        self.headlerName = ['ФИО', 'Подразделение', 'Должность', 'Логин', 'Пароль']
        self.resultTable.setHorizontalHeaderLabels(self.headlerName)
        self.resultTable.horizontalHeader().resizeSection(0, 250)
        self.resultTable.horizontalHeader().resizeSection(1, 200)
        self.resultTable.horizontalHeader().resizeSection(2, 200)

        # spisok layout -----------------------------------------------------------------------------------------------
        self.labelSpisokFileName = QtWidgets.QLabel("Список сотрудников. Название файла (en):")
        self.lineEditSpisokFileName = QtWidgets.QLineEdit()
        self.lineEditSpisokFileName.setText(self.spisokFileName)

        spisokLayout = QtWidgets.QHBoxLayout()
        spisokLayout.addWidget(self.labelSpisokFileName)
        spisokLayout.addWidget(self.lineEditSpisokFileName)

        # data layout -------------------------------------------------------------------------------------------------
        self.labelDataFileName = QtWidgets.QLabel("Данные Олимпокс. Название файла (en):")
        self.lineEditDataFileName = QtWidgets.QLineEdit()
        self.lineEditDataFileName.setText(self.dataFileName)

        dataLayout = QtWidgets.QHBoxLayout()
        dataLayout.addWidget(self.labelDataFileName)
        dataLayout.addWidget(self.lineEditDataFileName)

        # spisokData layout -------------------------------------------------------------------------------------------
        spisokDataLayout = QtWidgets.QVBoxLayout()
        spisokDataLayout.addLayout(spisokLayout)
        spisokDataLayout.addLayout(dataLayout)

        # listNames layout --------------------------------------------------------------------------------------------
        self.labelListNames = QtWidgets.QLabel("Название листов:")
        self.comboBoxListNames = QtWidgets.QComboBox()

        listNamesLayout = QtWidgets.QHBoxLayout()
        listNamesLayout.addWidget(self.labelListNames)
        listNamesLayout.addWidget(self.comboBoxListNames)

        # input layout ------------------------------------------------------------------------------------------------
        inputLayout = QtWidgets.QHBoxLayout()
        inputLayout.addLayout(spisokDataLayout)
        inputLayout.addLayout(listNamesLayout)

        # filter layout -----------------------------------------------------------------------------------------------
        self.comboBoxFilter = QtWidgets.QComboBox()
        self.filterPushButton = QtWidgets.QPushButton("Отфильтровать")

        filterLayout = QtWidgets.QHBoxLayout()
        filterLayout.addWidget(self.comboBoxFilter)
        filterLayout.addWidget(self.filterPushButton)

        # main layout ------------------------------------------------------------------------------------------------
        self.savePushButton = QtWidgets.QPushButton("Записать учётные данные выбранного предприятия в Excel")

        mainLayout = QtWidgets.QVBoxLayout()
        mainLayout.addLayout(inputLayout)
        # mainLayout.addWidget(self.resultPlainTextEdit)
        mainLayout.addWidget(self.resultTable)
        mainLayout.addLayout(filterLayout)
        mainLayout.addWidget(self.savePushButton)

        self.setLayout(mainLayout)

    def initSignals(self):
        self.filterPushButton.clicked.connect(self.onFilterPushButtonClicked)
        self.savePushButton.clicked.connect(self.onSavePushButtonClicked)

        self.lineEditSpisokFileName.textChanged.connect(self.changeSpisokFileName)
        self.lineEditDataFileName.textChanged.connect(self.changeDataFileName)
        self.comboBoxListNames.currentTextChanged.connect(self.changeListNames)


    def onOpen(self):
        print('Вы нажали Ctrl+C')

    def initFiles(self):
        data = pd.read_excel(self.dataFileName, sheet_name=self.listNames)
        spisok = pd.read_excel(self.spisokFileName, sheet_name=self.listNames)

        self.tabs = pd.ExcelFile(self.spisokFileName).sheet_names

        # обработчик параметров data
        usernames = data['ФИО'].tolist()
        logins = data['Логин'].tolist()
        passwords = data['Пароль'].tolist()

        # обработчик параметров spisok
        head = [column for column in spisok]

        self.departments = spisok[head[1]].tolist()
        workers = spisok['ФИО'].tolist()
        post = spisok['Должность'].tolist()

        # создание единого файла параметров ---------------------------------------------------------------------------
        self.resultTable.setRowCount(0)

        self.result_data = []
        for ind_worker, worker in enumerate(workers):
            for ind_user, username in enumerate(usernames):
                if username == worker:
                    res = [username,
                           self.departments[ind_worker],
                           post[ind_worker],
                           logins[ind_user],
                           passwords[ind_user]]
                    self.result_data.append(res)

        departments = []
        for department in self.departments:
            if isinstance(department, str):
                departments.append(department)

        self.filter_list_departments = list(set(departments))
        self.filter_list_departments.sort()

        self.result_data_depart = []
        for department in self.result_data:
            self.result_data_depart.append(department[1])

        self.comboBoxFilterSet()

    def changeSpisokFileName(self):
        self.spisokFileName = self.lineEditSpisokFileName.text()

    def changeDataFileName(self):
        self.dataFileName = self.lineEditDataFileName.text()

    def changeListNames(self):
        self.listNames = self.comboBoxListNames.currentText()
        self.initFiles()

    def comboBoxFilterSet(self):
        self.comboBoxFilter.clear()
        self.comboBoxFilter.addItems(self.filter_list_departments)

    def comboBoxListNamesSet(self):
        self.comboBoxListNames.addItems(self.tabs)

    def onFilterPushButtonClicked(self):
        self.resultTable.setRowCount(0)

        self.output = []
        for ind_department, department in enumerate(self.result_data_depart):
            if self.comboBoxFilter.currentText() == department:
                x = 0
                text_list = self.result_data[ind_department]
                self.output.append(', '.join(text_list))
                self.resultTable.insertRow(0)
                for column in text_list:
                    item = QtWidgets.QTableWidgetItem()
                    item.setText(column)
                    self.resultTable.setItem(0, x, item)
                    x += 1

    def onSavePushButtonClicked(self):
        df = pd.DataFrame(self.result_data, columns=self.headlerName)
        with pd.ExcelWriter('result.xlsx', engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name=f'{self.comboBoxListNames.currentText()}')


if __name__ == "__main__":
    app = QtWidgets.QApplication()

    window = CPP_MakeSender()
    window.show()

    app.exec()

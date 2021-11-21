# -----------------------------------------------------------
# Coding: UTF-8
# Generator of acts
#
# (C) 2021 Grif Ivan, Noyabrsk, Russia
# Released under GNU Public License (GPL)
# email grif.iv@acwiki.ru
# -----------------------------------------------------------

import sys                                                      # sys нужен для передачи argv в QApplication
import pandas as pd                                             # Импортируем pandas для работы с DataFrame
from PyQt5 import QtWidgets, QtCore                             # Импортируем PyQt5 для работы интерфейсом
from PyQt5.QtWidgets import QMessageBox, QStyledItemDelegate    # Импортируем QMessageBox для окон ошибок
from docxtpl import DocxTemplate                                # Импортируем docxtpl для работы с Excel
from PyQt5.QtGui import QColor, QPalette
import design                                                   # Импортируем файл дизайна PyQT5


# Создаем DataFrame из export.xlsx
df_1 = pd.read_excel(r'export.xlsx', sheet_name='Лист1')
# Парсим таблицу и вынимаем столбец "Пользователь" в виде списка
Users = list((df_1['Пользователь']))
# Добавляем в начало списка пустую строку
Users.insert(0, "")
# Парсим таблицу и вынимаем столбец "Тип" в виде списка
Type = list((df_1['Тип']))
# Парсим таблицу и вынимаем столбец "Модель" в виде списка
Model = list((df_1['Модель']))
# ...
SN = list((df_1['Серийный номер']))


# Делагат для центрирования данных в таблице
class AlignDelegate(QtWidgets.QItemDelegate):
    # Доступ к переменным, методам и т.д.
    def paint(self, painter, option, index):
        option.displayAlignment = QtCore.Qt.AlignCenter
        QtWidgets.QItemDelegate.paint(self, painter, option, index)


#class ColorDelegate(QStyledItemDelegate):
#    def paint(self, painter, option, index):
#        if index.data() == SN:
#            option.palette.setColor(QPalette.Text, QColor("green"))
#        elif index.data() == 'Offline':
#            option.palette.setColor(QPalette.Text, QColor("red"))
#        QStyledItemDelegate.paint(self, painter, option, index)


class ExampleApp(QtWidgets.QMainWindow, design.Ui_MainWindow):
    # Доступа к переменным, методам и т.д. в файле design.py
    def __init__(self):
        super().__init__()
        # Инициализация нашего дизайна
        self.setupUi(self)
        # Привязываем кнопку "Удалить строку" к событию Delete_Row_1
        self.deleteRow_1.clicked.connect(self.delete_row_1)
        # Привязываем кнопку "Добавить строку" к событию Add_Row_1
        self.addRow_1.clicked.connect(self.add_row_1)
        # Привязываем кнопку "Удалить строку" к событию Delete_Row_2
        self.deleteRow_2.clicked.connect(self.delete_row_2)
        # Привязываем кнопку "Добавить строку" к событию Add_Row_2
        self.addRow_2.clicked.connect(self.add_row_2)
        # Определяем QtWidgets.QComboBox
        self.comboBox1 = QtWidgets.QComboBox()
        # Добавляем данные из переменной Users, Users = list(set(df['Пользователь']))
        self.comboBox.addItems(Users)
        # Центрирование текста с помощью QItemDelegate
        self.DataTable_1.setItemDelegate(AlignDelegate())
        #self.DataTable_1.setItemDelegate(ColorDelegate())
        # Центрирование текста с помощью QItemDelegate
        self.DataTable_2.setItemDelegate(AlignDelegate())
        # Привязываем событие comboBox.activated к функции on_activated
        self.comboBox.activated.connect(self.on_activated)
        # Привязываем событие pushButton.clicked к функции Gen_Akt
        self.pushButton.clicked.connect(self.create_act)
        # Блок корректировки даты dateEdit
        self.dateEdit.setDateTime(QtCore.QDateTime.currentDateTime())
        self.dateEdit.setMaximumDate(QtCore.QDate(7999, 12, 28))
        self.dateEdit.setMaximumTime(QtCore.QTime(23, 59, 59))
        self.dateEdit.setCalendarPopup(True)

        self.tw = None
        self.copied_cells = None
        self.DataTable_1.cellPressed.connect(
            lambda row, col, tw=self.DataTable_1: self.selection_changed(row, col, tw))
        self.DataTable_2.cellPressed.connect(
            lambda row, col, tw=self.DataTable_2: self.selection_changed(row, col, tw))

    def selection_changed(self, row, col, tw):
        self.tw = tw
        #print(f'tw = {tw.objectName()} - {row} - {col}')

    def keyPressEvent(self, event):
        super().keyPressEvent(event)
        if not self.tw:
            return
        if event.key() == QtCore.Qt.Key_C and (event.modifiers() & QtCore.Qt.ControlModifier):
            self.copied_cells = sorted(self.tw.selectedIndexes())
            # print(f'copied_cells = {self.copied_cells}')
        elif event.key() == QtCore.Qt.Key_V and (event.modifiers() & QtCore.Qt.ControlModifier):
            if not self.copied_cells:
                return
                # print(f'objectName = {self.tw.objectName()}')
            r = self.tw.currentRow() - self.copied_cells[0].row()
            c = self.tw.currentColumn() - self.copied_cells[0].column()
            # print(f'{r} - {c}')
            for cell in self.copied_cells:
                self.tw.setItem(
                    cell.row() + r,
                    cell.column() + c,
                    QtWidgets.QTableWidgetItem(cell.data())
                )

    # Функция добавления данных из DataFrame, добавление данных при выборе элементов в comboBox
    def on_activated(self):
        # Создаем переменную selected_box с текстом comboBox
        selected_box = self.comboBox.currentText()
        # Создаем переменную df_name и добавляем в нее выбранного пользователя из comboBox
        df_name = df_1['Пользователь'].isin([selected_box])
        # Создаем переменную df_2 и добавляем в нее значения из файла export.xlsx согласно переменной df_name
        df_2 = df_1[df_name][['Тип', 'Модель', 'Серийный номер', 'Место установки', 'Пользователь']].reset_index(
            drop=True)
        # Создаем переменную headers и добавляем в нее имена заголовков переменной df_2
        headers = df_2.columns.values.tolist()
        # Блок отвечающий за заполнение DataTable_2 данными согласно заголовкам(headers) и значениям файла export.xlsx
        self.DataTable_2.setRowCount(0)
        self.DataTable_2.setColumnCount(len(headers))
        self.DataTable_2.setHorizontalHeaderLabels(headers)
        for i, row in df_2.iterrows():
            row_add = self.DataTable_2.rowCount()
            self.DataTable_2.setRowCount(row_add + 1)
            for j in range(self.DataTable_2.columnCount()):
                self.DataTable_2.setItem(i, j, QtWidgets.QTableWidgetItem(str(row[j])))

    # Функция добавления строк в таблице DataTable_1
    def add_row_1(self):
        row_position = self.DataTable_1.rowCount()
        self.DataTable_1.insertRow(row_position)

    # Функция удаления строк в таблице DataTable_1
    def delete_row_1(self):
        rows = set()
        for index in self.DataTable_1.selectedIndexes():
            rows.add(index.row())
        for row in sorted(rows, reverse=True):
            self.DataTable_1.removeRow(row)

    # Функция добавления строк в таблице DataTable_2
    def add_row_2(self):
        row_position = self.DataTable_2.rowCount()
        self.DataTable_2.insertRow(row_position)

    # Функция удаления строк в таблице DataTable_2
    def delete_row_2(self):
        rows = set()
        for index in self.DataTable_2.selectedIndexes():
            rows.add(index.row())
        for row in sorted(rows, reverse=True):
            self.DataTable_2.removeRow(row)

    # Функция создания акта, при нажатии кнопки "Сгенерировать акт"
    def create_act(self):
        # Создаем переменную ticket содержащую № заявки из lineEdit
        ticket = self.lineEdit.text()
        # Создаем переменную doc в которой содержится шаблон для модуля docxtpl
        doc = DocxTemplate('word.docx')
        # Создаем переменную rows_1 которая содержит количество строк в таблице DataTable_1
        rows_1 = self.DataTable_1.rowCount()
        # Создаем переменную rows_2 которая содержит количество строк в таблице DataTable_2
        rows_2 = self.DataTable_2.rowCount()
        # Создаем стандартный массив(list) mount_data с пустыми значением
        mount_data = []

        rows = self.DataTable_1.rowCount()
        listi = []
        for i in range(rows):
            it = self.DataTable_1.item(i, 0)
            if it and it.text():
                listi.append(it.text())
                print("Что происходит?")

        for i in range(6):
            try:
                if self.DataTable_1.item(0, i).text() == 0 or self.DataTable_1.item(0, i).text() == "":
                    msg = QMessageBox()
                    msg.setWindowTitle("Ошибка")
                    msg.setText('Заполните все ячейки в таблице "Монтаж"')
                    msg.setIcon(QMessageBox.Warning)
                    msg.exec_()
                else:
                    print("Ячейки не пусты")
            except:
                print("Ошибка ТуреЕррор")
        # Цикл добавления данных с помощью модуля docxtpl
        for row in range(rows_1):
            # Начало отслеживания ошибок
            try:
                # Создаем переменную tmp_1 в которой содержатся строки из DataTable_1
                # Приведенные к виду ТИП: Модель (SN: Серийный номер)
                tmp_1 = self.DataTable_1.item(row, 0).text().upper() + ': ' + self.DataTable_1.item(row, 1).text() + \
                        ' (SN: ' + self.DataTable_1.item(row, 2).text() + ')'
                # Добавляем в стандартный массив данные из tmp_1
                mount_data.append(tmp_1)
            # Отслеживаем исключение AttributeError
            except AttributeError:
                pass
                # Выводим сообщение об ошибке
                #msg = QMessageBox()
                #msg.setWindowTitle("Ошибка")
                #msg.setText('Заполните все ячейки в таблице "Монтаж"')
                #msg.setIcon(QMessageBox.Warning)
                #msg.exec_()
        # Создаем стандартный массив(list) demount_data с пустыми значением
        demount_data = []
        # Цикл добавления данных с помощью модуля docxtpl
        for row in range(rows_2):
            # Начало отслеживания ошибок
            try:
                # Создаем переменную tmp_2 в которой содержатся строки из DataTable_2
                # Приведенные к виду ТИП: Модель (SN: Серийный номер)
                tmp_2 = self.DataTable_2.item(row, 0).text().upper() + ': ' + self.DataTable_2.item(row, 1).text() + \
                        ' (SN: ' + self.DataTable_2.item(row, 2).text() + ')'
                # Добавляем в стандартный массив данные из tmp_2
                demount_data.append(tmp_2)
            except AttributeError:
                pass
                # Выводим сообщение об ошибке
                #msg = QMessageBox()
                #msg.setWindowTitle("Ошибка")
                #msg.setText('Заполните все ячейки в таблице "Демонтаж"')
                #msg.setIcon(QMessageBox.Warning)
                #msg.exec_()
                #demount_data = []

        tbl_contents_1 = mount_data
        tbl_contents_2 = demount_data
        context = {'tbl_contents_1': tbl_contents_1,
                   'tbl_contents_2': tbl_contents_2,
                   'ticket': ticket
                   }

        doc.render(context)
        doc.save('result.docx')


def main():
    app = QtWidgets.QApplication(sys.argv)  # Новый экземпляр QApplication
    window = ExampleApp()  # Создаём объект класса ExampleApp
    window.show()  # Показываем окно
    app.exec_()  # и запускаем приложение


if __name__ == '__main__':  # Если мы запускаем файл напрямую, а не импортируем
    main()  # то запускаем функцию main()

'''
rows = self.DataTable_2.rowCount()
cols = self.DataTable_2.columnCount()
data = []
for row in range(rows):
    tmp = []
    for col in range(cols):
        try:
            tmp.append(self.DataTable_2.item(row, col).text())
        except IndexError:
            tmp.append('No data')
    data.append(tmp)
    #print(tmp[1])
    array = [tmp]
    context = []

    for x in array:
        context.append({"Tip": x[0], "Model": x[1], "SN": x[2]})
    print(context)


#context1 = {i: context[i] fo   r i in range(0, len(context))}
#print(context1)
tpl.render(context)
tpl.save('table1.docx')


col_count = self.DataTable_2.columnCount()
row_count = self.DataTable_2.rowCount()
headers1 = [str(self.DataTable_2.horizontalHeaderItem(i).text()) for i in range(col_count)]

df_list = []
for row in range(row_count):
    df_list2 = []
    for col in range(col_count):
        table_item = self.DataTable_2.item(row, col)
        df_list2.append('' if table_item is None else str(table_item.text()))
    df_list.append(df_list2)

df = pd.DataFrame(df_list, columns=headers1)
print(df)
print(list(set(df['1'])))
'''

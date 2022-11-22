import sys, pypyodbc, os
from PyQt6 import QtCore, QtGui, QtWidgets
from form.form import Ui_Dialog


def db_connect() -> pypyodbc.connect:
    pypyodbc.lowercase = False

    conn = pypyodbc.connect(
        "Driver={Microsoft Access Driver (*.mdb, *.accdb)};" +
        rf"Dbq={os.path.dirname(os.path.realpath(__file__))}\data\OC.accdb;")
    return conn


def get_query_table(cursor :pypyodbc.Cursor, os_code :str) -> list:
    cursor.execute(f"SELECT Программы.Программа, Пакеты.Пакеты, Категории_ОС.ОС, Программы.Код \
                    FROM Пакеты INNER JOIN (Категории_ОС INNER JOIN Программы ON Категории_ОС.Код = Программы.Код_ПК) \
                    ON Пакеты.Код = Программы.Код_пакета \
                    WHERE Категории_ОС.Код = ?;", (os_code,))

    return cursor.fetchall()

def get_os_table(cursor :pypyodbc.Cursor) -> list:
    cursor.execute("SELECT * FROM Категории_ОС;")

    return cursor.fetchall()


def get_package_table(cursor :pypyodbc.Cursor) -> list:
    cursor.execute("SELECT * FROM Пакеты;")

    return cursor.fetchall()


def get_program_table(cursor :pypyodbc.Cursor) -> list:
    cursor.execute("SELECT * FROM Программы;")

    return cursor.fetchall()


def insert_program(conn :pypyodbc.connect, program :str, os :str, package :str):
    cursor = conn.cursor()
    cursor.execute(f"INSERT INTO Программы (Программа, Код_ПК, Код_пакета) VALUES (?, ?, ?);", (program, os, package))
    conn.commit()


def delete_program(conn :pypyodbc.connect, program_code :int):
    cursor = conn.cursor()

    cursor.execute("DELETE * FROM Программы WHERE Программы.Код = ?", (program_code,))
    conn.commit()


class App(Ui_Dialog):
    def __init__(self, MainWindow, conn) -> None:
        super().__init__()
        self.setupUi(MainWindow)
        
        self.init_os_dict(conn)
        self.init_package_dict(conn)

        self.load_button.clicked.connect(lambda: self.load_data(conn))
        self.add_button.clicked.connect(lambda: self.add_program(conn))
        self.delete_button.clicked.connect(lambda: self.del_program(conn))
    
    def load_data(self, conn):
        os = self.oc_combobox.currentText()
        program_table = get_query_table(conn.cursor(), self.os_dict[os])
        self.set_table_data(program_table)
        

    def init_package_dict(self, conn):
        package_table = get_package_table(conn.cursor())

        self.package_dict = {}
        for row in package_table:
            self.package_dict[row[1]] = row[0]


    def init_os_dict(self, conn):
        os_table = get_os_table(conn.cursor())

        self.os_dict = {}    
        for row in os_table:
            self.os_dict[row[1]] = row[0]


    def set_table_data(self, data):
        self.data_table.setRowCount(len(data))

        for i, row in enumerate(data):
            for j, item in enumerate(row):
                new_item = QtWidgets.QTableWidgetItem(str(item))
                self.data_table.setItem(i, j, new_item)
        
        header = self.data_table.horizontalHeader()       
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QtWidgets.QHeaderView.ResizeMode.Stretch)


    def add_program(self, conn):
        program_table = get_program_table(conn.cursor())

        program_list = [row[1] for row in program_table]

        program = self.program_input.text()
        os = self.oc_input.text()
        package = self.package_input.text()

        if program not in program_list and os in self.os_dict and package in self.package_dict:
            insert_program(conn, program, self.os_dict[os], self.package_dict[package])
            self.error_label.setText("")
        else:
            self.error_label.setText("Ошибка ввода, попробуйте снова")

    
    def del_program(self, conn):
        program_table = get_program_table(conn.cursor())
        program_code_list = [row[0] for row in program_table]

        program_code = int(self.program_code_input.text())

        if program_code in program_code_list:
            delete_program(conn, program_code)
            self.error_label.setText("")
        else:
            self.error_label.setText("Ошибка удаления, неверный код программы")

        

if __name__ == "__main__":
    conn = db_connect()
    
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = App(MainWindow, conn)
    MainWindow.show()
    sys.exit(app.exec())
import sys
import datetime
import logging
import pandas as pd
import oracledb
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QVBoxLayout, QPushButton, QProgressBar, QMessageBox
from PyQt5.QtCore import Qt, QThread, pyqtSignal

class WorkerThread(QThread):
    progress_updated = pyqtSignal(int)
    status_updated = pyqtSignal(str)
    completed = pyqtSignal()
    error_occurred = pyqtSignal(str)

    def run(self):
        try:
            self.status_updated.emit("Подключение к БД...")
            self.progress_updated.emit(10)
            conn = oracledb.connect("Пароль и Логин к БД")
            cur = conn.cursor()
            query = """
        select distinct s.application_id,
                        s.code,
                        cast(s.status_date as date),
                        pc.product_class, cht.description
        from cl_la.status s
        left join product_content pc
            on pc.application_id = s.application_id
        left join trade_point tp on tp.application_id = s.application_id and tp.type ='registration'
        left join cl_catalog.channel_type cht on cht.code = tp.channel_code
        where s.current_status = 1
            --and status_date > to_date('04.07.2023 7', 'dd.mm.yyyy hh24')
            and status_date < sysdate - 1 / 24 / 5
            and pc.type = '0'
            and pc.product_class = 'GP'
            and s.code not in ('Denied.Done',
                                'Completed.Done',
                                'Refused.Done',
                                'LoanDetailsReceiving',
                                'CCLoanDetailsReceiving',
                                'AdditionalFilling',
                                'ApplicationDataReceiving',
                                'EnterEAN',
                                'ESigningDocs',
                                'SigningDocuments')
        """
            self.status_updated.emit("Выполнение SQL запроса...")
            self.progress_updated.emit(45)
            cur.execute(query)
            results = cur.fetchall()

            # Сохранение данных...
            self.status_updated.emit("Сохранение Excel...")
            self.progress_updated.emit(95)
            current_time = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            file_name = f"unloading_{current_time}.xlsx"

            df = pd.DataFrame(results, columns=['application_id', 'code', 'status_date', 'product_class', 'description'])

            # Создание сводной таблицы
            pivot_table = df.groupby('code')['application_id'].count().reset_index()
            pivot_table.columns = ['Названия строк', 'Количество по полю Application ID']

            # Добавление строки "Общий итог"
            total_count = pivot_table['Количество по полю Application ID'].sum()
            total_row = pd.DataFrame({'Названия строк': ['Общий итог'], 'Количество по полю Application ID': [total_count]})

            # Объединение сводной таблицы и строки "Общий итог"
            pivot_table = pd.concat([pivot_table, total_row], ignore_index=True)

            # Сохранение данных в EXCEL...

            current_time = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            file_name = f"unloading_{current_time}.xlsx"

            with pd.ExcelWriter(file_name) as writer:
            # Форматирование для листа Table NKK
                df.to_excel(writer, index=False, sheet_name='Table NKK')
                worksheet = writer.sheets['Table NKK']

                number_format = writer.book.add_format({'num_format': '0'})
                worksheet.set_column('A:A', 35, number_format)
                worksheet.set_column('B:B', 25)
                worksheet.set_column('C:C', 25)
                worksheet.set_column('D:D', 25)
                worksheet.set_column('E:E', 25)
                
            # Форматирование для листа Pivot Table NKK
                pivot_table.to_excel(writer, index=False, sheet_name='Pivot Table NKK')
                worksheet_pivot = writer.sheets['Pivot Table NKK']
                worksheet_pivot.set_column('A:A' , 45)
                worksheet_pivot.set_column('B:B' , 45)

            self.status_updated.emit("Готово...")
            self.progress_updated.emit(100)
            # Закрытие БД...
            cur.close()
            conn.close()

            self.completed.emit()
        except Exception as e:
            logging.exception(e)
            self.error_occurred.emit(str(e))

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setWindowTitle("Выгрузка зависших заявок НКК")
        self.setFixedSize(420, 280)
        self.setStyleSheet("background-color: #bab5af;")
        
        self.status_label3 = QLabel(self)
        self.status_label3.setText("Выгрузка зависших заявок НКК" )
        self.status_label3.setStyleSheet("font: normal 14px;")
        
        self.status_label4 = QLabel(self)
        self.status_label4.setText(' 1.1 ver. ')
        self.status_label4.setOpenExternalLinks(True)
        
        self.status_label2 = QLabel(self)
        self.status_label2.setText('<a href="https://dit.rencredit.ru/confluence/pages/viewpage.action?pageId=241107200">Confluece</a>')
        self.status_label2.setOpenExternalLinks(True)
        
        
        self.status_label = QLabel(self)
        self.status_label.setText("Статус выгрузки:")
        self.status_label.setStyleSheet("background-color: #474546; color: #ffffff; font: bold;")

        self.progressbar = QProgressBar(self)
        self.progressbar.setOrientation(Qt.Horizontal)
        self.progressbar.setRange(0, 100)
        self.progressbar.setTextVisible(True)
        
        self.start_button = QPushButton(self)
        self.start_button.setText("Нажать для выгрузки")
        self.start_button.setFixedWidth(170)
        self.start_button.clicked.connect(self.load)
        self.start_button.setStyleSheet("""
                                        QPushButton {
                                            background-color: #a826d4;
                                            border: none;
                                            color: white;
                                            padding: 15px;
                                            text-align: center;
                                            border-radius: 15px;
                                            font: normal 14px;
                                        }""")

        self.exit_button = QPushButton(self)
        self.exit_button.setText("Выход")
        self.exit_button.setFixedWidth(100)
        self.exit_button.clicked.connect(self.show_exit_confirmation)
        self.exit_button.setStyleSheet("""
                                        QPushButton {
                                            background-color: #a826d4;
                                            border: none;
                                            color: white;
                                            padding: 15px;
                                            text-align: center;
                                            font: normal 14px;
                                            border-radius: 15px;
                                        }""")

        layout = QVBoxLayout(self)
        layout.addWidget(self.status_label4, alignment=Qt.AlignRight)
        layout.addWidget(self.status_label2, alignment=Qt.AlignRight)
        layout.addWidget(self.status_label3, alignment=Qt.AlignCenter)
        layout.addSpacing(50)

        layout.addWidget(self.status_label, alignment=Qt.AlignCenter)
        layout.addWidget(self.progressbar)
        layout.addWidget(self.start_button, alignment=Qt.AlignCenter)
        layout.addWidget(self.exit_button, alignment=Qt.AlignCenter)

        self.worker_thread = WorkerThread()
        self.worker_thread.progress_updated.connect(self.update_progress)
        self.worker_thread.status_updated.connect(self.update_status)
        self.worker_thread.completed.connect(self.unloading_completed)
        self.worker_thread.error_occurred.connect(self.show_error_message)
        self.update_progress(0)  # Add this line


    def load(self):
        self.progressbar.setValue(0)  
        self.start_button.setEnabled(False)
        self.start_button.setStyleSheet("""
                                        QPushButton {
                                            border: none;
                                            color: #bab5af;
                                            padding: 15px;
                                            text-align: center;
                                            border-radius: 15px;
                                            font: normal 14px;
                                        }""")
        self.worker_thread.start()

    def show_exit_confirmation(self):
        confirmation = QMessageBox.question(
            self, "Подтверждение выхода",
            "Вы уверены, что хотите выйти?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if confirmation == QMessageBox.Yes:
            self.worker_thread.quit()
            self.close()

    def update_progress(self, value):
        self.progressbar.setValue(value)
        self.progressbar.setFormat(f"{value}%")

    def update_status(self, status):
        self.status_label.setText(status)

    def unloading_completed(self):
        self.progressbar.reset()
        self.status_label.setText("Готово")
        QMessageBox.information(None, "Выгрузка", "Выполнение выгрузки - готово")
        self.worker_thread.quit()
        self.close()

    def show_error_message(self, error_message):
        self.progressbar.reset()
        self.status_label.setText("Ошибка")
        QMessageBox.critical(None, "Ошибка", error_message)

    def closeEvent(self, event):
        self.worker_thread.quit()
        super().closeEvent(event)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.moving = True
            self.offset = event.pos()

    def mouseMoveEvent(self, event):
        if self.moving:
            self.move(event.globalPos() - self.offset)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.moving = False

if __name__ == "__main__":
    app = QApplication(sys.argv)
    QApplication.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
import os , logging , datetime , oracledb
import pandas as pd
from threading import Thread
import tkinter as tk
from tkinter import messagebox
from tkinter.ttk import Progressbar


log_dir = 'C:\\log' # Проверка Пути для логов

if not os.path.exists(log_dir):
    os.makedirs(log_dir, exist_ok=True)

log_file = os.path.join(log_dir, "unloading.log") # Запись лога в файл

logging.basicConfig(filename=log_file, level=logging.INFO, format='%(asctime)s %(levelname)s %(funcName)s || %(message)s')  # Конфигурация логов

root = tk.Tk()
root.title("Выгрузка")
root.geometry("420x150")
root.resizable(False, False)
root.configure(bg="#474546")

status_label = tk.Label(root, text="Статус выгрузки :", background="#474546", foreground="#ffffff", font="TkFixedFont")
status_label.pack()

progress_label = tk.Label(root, text='0%', background="#474546", foreground="#ffffff", font="TkFixedFont")
progress_label.pack(pady=5)

def excel():
    try:
        progress_label.configure(text='10%')
        conn = oracledb.connect("Пароль и Логин к БД")
        cur = conn.cursor()
        progress_label.configure(text='35%')
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
        progress_label.configure(text='55%')
        cur.execute(query)
        progress_label.configure(text='65%')
        results = cur.fetchall()
        
        # Сохранение данных...
        progress_label.configure(text='99%')
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
            workbook  = writer.book
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

        # Закрытие БД...
        cur.close()
        conn.close()

        messagebox.showinfo("Выгрузка", "Выполнение выгрузки - готово")
        status_label.configure(text="Готово")
        progress_label.configure(text='100%')
        progressbar.stop()
        os.system("taskkill /f /im unloading.exe")
    except:
        # Если всё не поплану =(
        logging.exception(excel)
        messagebox.showerror("Ошибка", "Нету подключения к БД")
        os.system("taskkill /f /im unloading.exe")

def load():
    progressbar.start()
    start_button.config(state="disabled")
    status_label.configure(text="Выполнение выгрузки...")
    root.update_idletasks()
    start = Thread(target=excel)
    start.start()

progressbar = Progressbar(root, orient=tk.HORIZONTAL, mode="indeterminate")
progressbar.pack(fill=tk.X)

start_button = tk.Button(root, text="Нажать для выгрузки", command=load, width=20)
start_button.pack(pady=10)

root.mainloop()
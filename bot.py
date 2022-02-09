import os
import pandas as pd
import xlrd
import xlsxwriter

import time
from sys import platform

if platform == 'win32':
    import ctypes
    kernel32 = ctypes.windll.kernel32
    kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)

success_message = '\033[2;30;42m [SUCCESS] \033[0;0m' 
warning_message = '\033[2;30;43m [WARNING] \033[0;0m'


class SalesDetails():
    def __init__(self):
        self.month_names = {
                '01': 'January',
                '02': 'February',
                '03': 'March',
                '04': 'April',
                '05': 'May',
                '06': 'June',
                '07': 'July',
                '08': 'August',
                '09': 'September',
                '10': 'October',
                '11': 'November',
                '12': 'December',}

        os.makedirs('Upload Excel', exist_ok=True)
        os.makedirs('Report', exist_ok=True)
    
    def run(self):
        print(warning_message + '\tСперва объединим все файлы из папки Upload Excel')
        self.merge_files()

    def merge_files(self):
        files = os.listdir('Upload Excel')  
        df = pd.DataFrame()
 
        for file in files:
            if file.endswith('.xlsx'):
                df = df.append(pd.read_excel(f'Upload Excel/{file}'), ignore_index=True) 
 
        df.head() 
        df.to_excel('Upload Excel/Total_excel.xlsx')
        print(success_message + '\t Все excel-файлы были собраны в один Upload Excel/Total_Excel.xlsx')

        self.pack_to_file_by_month()
    
    def pack_to_file_by_month(self):
        wb = xlrd.open_workbook('Upload Excel/Total_excel.xlsx')
        ws = wb.sheet_by_index(0)
        
        table_titles = ws.row_values(0)

        for col_num in range(len(table_titles)):
            if table_titles[col_num] == 'Дата продажи':
                date_col = col_num
                all_sale_date = list(set([item.split('-')[1] for item in  ws.col_values(col_num)[1:]]))
        
        for date_month in all_sale_date:
            month_collect = [table_titles]
            for row_num in range(1, ws.nrows-1):
                date = ws.cell(row_num, date_col).value.split('-')[1]
        
                if date_month == date:
                    month_collect.append(ws.row_values(row_num))

            if date_month in self.month_names:
                dirname = os.path.join('Report',self.month_names[date_month])
                os.makedirs(dirname, exist_ok=True)
                self.save_excel(f"{dirname}/Month_data.xlsx", month_collect)
                print(success_message + f'\tСоздал папку {dirname}')


    def save_excel(self, filename, data):
        writer_book = xlsxwriter.Workbook(filename)
        sheet = writer_book.add_worksheet()
        
        for row_num, row_group in enumerate(data):
            for col_num, value in enumerate(row_group):
                sheet.write(row_num, col_num, value)

        writer_book.close()  


    def walk_in_folders(self):
        folders = os.listdir('Report')
        for folder in folders:
            
            file = f'Report/{folder}/' + 'Month_data.xlsx'
            self.find_penalties(folder=f'Report/{folder}', filename=file)
            self.find_logistic(folder=f'Report/{folder}', filename=file)
            self.find_sales_stock(folder=f'Report/{folder}', filename=file)
            self.find_refund(folder=f'Report/{folder}', filename=file)
    
    def find_penalties(self, folder, filename):
        print(filename)
        table_reader = xlrd.open_workbook(filename)
        sheet_reader = table_reader.sheet_by_index(0)

        table_titles = sheet_reader.row_values(0)
        for col_num in range(len(table_titles)):
            if table_titles[col_num] == 'Обоснование для оплаты':
                justifications_for_payment = sheet_reader.col_values(col_num)
        
        penalties = [table_titles, ]
        for item in range(len(justifications_for_payment)):
            if justifications_for_payment[item] in ('Штраф', 'Штраф МП'):
                penalties.append(sheet_reader.row_values(item))
        if len(penalties) > 1: 
            self.save_excel(f'{folder}/Penalties.xlsx', penalties)
            print(success_message + '\tСоздал файл Штрафов')


   
    def find_refund(self, folder, filename):
        table_reader = xlrd.open_workbook(filename)
        sheet_reader = table_reader.sheet_by_index(0)

        table_titles = sheet_reader.row_values(0)
        for col_num in range(len(table_titles)):
            if table_titles[col_num] == 'Обоснование для оплаты':
                justifications_for_payment = sheet_reader.col_values(col_num)
        
        refunded = [table_titles, ]
        for item in range(len(justifications_for_payment)):
            if justifications_for_payment[item] == 'Возврат':
                refunded.append(sheet_reader.row_values(item))

        if len(refunded) > 1:
            self.save_excel(f'{folder}/Refund.xlsx', refunded)
            print(success_message + '\t Создал файл Возвратов')


   
    def find_logistic(self, folder, filename):
        table_reader = xlrd.open_workbook(filename)
        sheet_reader = table_reader.sheet_by_index(0)

        table_titles = sheet_reader.row_values(0)
        for col_num in range(len(table_titles)):
            if table_titles[col_num] == 'Обоснование для оплаты':
                justifications_for_payment = sheet_reader.col_values(col_num)
            elif table_titles[col_num] == 'Количество возврата':
                count_refund_col = col_num
        
        logistics = [table_titles]
        for item in range(len(justifications_for_payment)):
            if justifications_for_payment[item] == 'Логистика':
                logistics.append(sheet_reader.row_values(item))
        
        refunded_rows = [table_titles]

        for item in range(1, (len(logistics) - 1)):
            refund_count = int(logistics[item][count_refund_col])
            if refund_count != 0:
                refunded_rows.append(logistics[item])
        
        writer_book = xlsxwriter.Workbook(f'{folder}/Logistic.xlsx')
        sheet = writer_book.add_worksheet('Логистика')
        sheet2 = writer_book.add_worksheet('Возвраты')
        
        
        for row_num, row_group in enumerate(logistics):
            for col_num, value in enumerate(row_group):
                sheet.write(row_num, col_num, value)
        
        for row_num, row_group in enumerate(refunded_rows):
            for col_num, value in enumerate(row_group):
                sheet2.write(row_num, col_num, value)

        writer_book.close()    
        print(success_message + '\tСоздал файл Логистики')

   
    def find_sales_stock(self, folder, filename):
        table_reader = xlrd.open_workbook(filename)
        sheet_reader = table_reader.sheet_by_index(0)

        table_titles = sheet_reader.row_values(0)
        for col_num in range(len(table_titles)):
            if table_titles[col_num] == 'Обоснование для оплаты':
                justifications_for_payment = sheet_reader.col_values(col_num)
            elif table_titles[col_num] == 'Склад':
                stock = col_num
        
        # print(sheet_reader.col_values(stock))

        sales_stock = [table_titles]
        for item in range(len(justifications_for_payment)):
            if justifications_for_payment[item] == 'Продажа':
                sales_stock.append(sheet_reader.row_values(item))
        
        my_stock = [table_titles]
        another_stock = [table_titles]

        for item in range(1, (len(sales_stock) - 1)):
            stock_type = sales_stock[item][stock]
            if stock_type.strip() not in  ('Склад поставщика', 'Склад поставщика 72 часа'):
                another_stock.append(sales_stock[item])
            else:
                my_stock.append(sales_stock[item])
        
        writer_book = xlsxwriter.Workbook(f'{folder}/Sales_stock.xlsx')
        sheet = writer_book.add_worksheet('Склад поставщика')
        sheet2 = writer_book.add_worksheet('Другие склады')
        
        
        for row_num, row_group in enumerate(my_stock):
            for col_num, value in enumerate(row_group):
                sheet.write(row_num, col_num, value)
        
        for row_num, row_group in enumerate(another_stock):
            for col_num, value in enumerate(row_group):
                sheet2.write(row_num, col_num, value)

        writer_book.close()   
        print(success_message + '\tСоздал файл Поставщиков')

    

bot = SalesDetails()
start_time = time.time()
bot.run()
bot.walk_in_folders()
end_time = time.time()
print(f'Бот потратил {round(end_time-start_time, 1)} секунд.')

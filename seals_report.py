import xlwings as xw
import os
import sys
import re

from customers import get_sub_table_name, Customers_Book, Customers_Book_Exception


seals_report_template_name = os.path.join(os.getenv("OneDrive"), 'BSAT', 'Прокладки и пр', 'РЕЕСТР прокладок.xltx') 
path2save = os.path.join(os.getenv("OneDrive"), 'BSAT', 'Прокладки и пр') 
correspondence_sheet = "Соответствие"
correspondence_table_name = "ТаблицаСоответствий"
s_orders = "заказ-наряды"
s_orders_table = "ТаблицаЗаказНарядов"

report_columns = {
    'dates': "Дата",
    'customers': "Заказчик",
    'drivers': "Водитель",
    'numbers_1': "Гос № (1)",
    'numbers_2': "Гос № (2)",
    'add_works': "Доп работы",
    'qty': "Количество",
}

class Order_Book_Exception(Exception):
    pass

def table_add_row(table_sheet: xw.Sheet, table_name: str):
    last_row = table_sheet.range(table_name).row + len(table_sheet.range(table_name).rows)
    line = f'{str(last_row)}:{str(last_row)}'
    table_sheet.range(line).insert()

def store_in_table(table_sheet: xw.Sheet, table_name: str, orders_sheet: xw.Sheet, orders_table_name: str, source_row_number: int):
    table_last_row = len(table_sheet.range(table_name).rows)
        
    table_add_row(table_sheet, table_name)

    for column_name in report_columns:
        dest_range = table_sheet.range(get_sub_table_name(table_name, report_columns[column_name]))
        source_range = orders_sheet.range(get_sub_table_name(orders_table_name, report_columns[column_name]))
        dest_range [table_last_row-1,0].value = source_range[source_row_number,0].value

    try:
        r_number = table_sheet.range(get_sub_table_name(table_name, "№"))
        r_number [table_last_row-1,0].value = table_last_row
        q_unit_name = table_sheet.range(get_sub_table_name(table_name, "Единица измерения"))
        q_unit_name[table_last_row-1,0].value = "шт"
        pass
    except :
        pass
        
    
class Report_Line():
    #work_number = "№"
        
    def __init__(self):
        self.column_names = {}
        for column_name in report_columns:
            self.column_names[column_name] = report_columns[column_name]

        pass
    def set_values(self, **kwargs):
        for kwarg in kwargs:
            self[kwarg] = kwarg


class Seals_Report():

    def __init__(self):
        '''
        create workbook with report from template
        '''
        self.wb_report = xw.Book(seals_report_template_name)
        self.seal_tables = {}
        self.order_columns = {}

    def get_order_book (self, order_book_name):
        '''
        get workbook with orders
        '''
        try:
            self.wb_orders = xw.Book(order_book_name)
            self.sh_orders = self.wb_orders.sheets[s_orders]
            self.r_orders = self.sh_orders.range(s_orders_table)
        except:
            raise Order_Book_Exception

        for column in report_columns:
            self.order_columns[column] = self.sh_orders.range(get_sub_table_name(s_orders_table, report_columns[column]))
        
    def load_corresponding_table(self):
        '''
        load seals tables
        '''
        self.corresponding_table = self.wb_report.sheets[correspondence_sheet].range(correspondence_table_name)
        for i in range(len(self.corresponding_table.rows)):
            self.seal_tables[self.corresponding_table[i,0].value] = (self.corresponding_table[i,1].value, self.corresponding_table[i,2].value)
        
            
    def load_tables(self):
        '''
        '''
        pass
    def load_orders_with_seals(self):
        '''
        '''
        for row_index in range(len(self.r_orders.rows)):
            add_work = self.sh_orders.range(get_sub_table_name(s_orders_table, report_columns['add_works']))[row_index,0].value
            if add_work is not None:
                try:
                    report_sheet_name = self.seal_tables[add_work][0]
                    report_sheet = self.wb_report.sheets[report_sheet_name]
                    report_table_name = self.seal_tables[add_work][1]
                    # report_table = self.wb_report.sheets[report_sheet].range(report_table_name)

                    store_in_table(report_sheet, report_table_name, self.sh_orders, s_orders_table, row_index)

                    #
                    
                except:
                    pass

    def set_dates(self):
        self.sh_params = self.wb_report.sheets["params"]
        self.r_month = self.sh_params.range("Month")
        self.r_year = self.sh_params.range("Year")
        s_dates =  self.wb_orders.name.split("_")
        self.r_month[0,0].value = int(s_dates[0])
        self.r_year[0,0].value = int(s_dates[1].split('.')[0])

    def save(self):
        try:
            self.wb_report.save(os.path.join(path2save, 'реестр прокладок '+ self.wb_orders.name))
        except:
            print ("some error while saving...")

    def replace_names(self):
        customers = Customers_Book()
        for table in self.seal_tables:
            sh_name = self.seal_tables[table][0]
            table_name = self.seal_tables[table][1]
            r_table_names = self.wb_report.sheets[sh_name].range(get_sub_table_name(table_name, report_columns['customers']))
            for row in range(len(r_table_names.rows)):
                name = r_table_names[row, 0].value
                if  name is not None:
                    short_name = customers.replace_name(name)
                    r_table_names[row, 0].value = short_name


if __name__ == '__main__':
    order_book_name = sys.argv[1]
    print(order_book_name)
    report = Seals_Report()
   
    try:
        report.get_order_book(order_book_name)
        report.load_corresponding_table()
        report.load_orders_with_seals()
        report.replace_names()
        report.set_dates()
        report.save()
    except Order_Book_Exception:
        print('some problem with order book')


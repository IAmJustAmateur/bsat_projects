import xlwings as xw
import sys
import os
customers_path =  os.path.join(os.getenv("OneDrive"), 'BSAT', 'Договора на мойку новые')

customers_book_name = "Заказчики.xlsx"
customers_sheet_name = "Заказчики Филиал"
customers_range_name = "ТаблицаКонтрагентов"
customer_names = "Наименование"
customer_short_names = "Наименование краткое"

def get_sub_table_name(table_name, column_name):
    return f'{table_name}[{column_name}]'

class Customers_Book_Exception(Exception):
    pass

class Customers_Book():
    def __init__(self):
        try:
            self.wb_customers = xw.Book(customers_book_name)
        except:
            try:
                self.wb_customers = xw.Book(os.path.join(customers_path, customers_book_name))
            except:
                raise Customers_Book_Exception
        try:
            self.sh_customers = self.wb_customers.sheets[customers_sheet_name]
            self.r_customers = self.sh_customers.range(customers_range_name)
            self.r_names = self.sh_customers.range(get_sub_table_name(customers_range_name, customer_names))
            self.r_short_names = self.sh_customers.range(get_sub_table_name(customers_range_name, customer_short_names))
        except:
            raise  Customers_Book_Exception

    def customer_short_name(self, name: str) -> str:
        customers_qty = len (self.r_customers.rows)
        for i in range(customers_qty):
            if self.r_names[i,0].value.lower() == name.lower():
                return self.r_short_names[i,0].value
    
    def replace_name(self, name: str) -> str:
        if name.lower() == "noname":
            return "Физлица"
        else:
            return self.customer_short_name(name)

    

        
        
    


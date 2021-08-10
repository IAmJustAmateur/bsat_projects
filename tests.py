from seals_report import Seals_Report
import os

report = Seals_Report()

#order_book_name = os.path.join(os.getenv("OneDrive"), 'BSAT', 'мойки филиал', '2021', '07_2021.xlsx')
order_book_short_name = '08_2021.xlsx'
report.get_order_book(order_book_short_name)
report.load_corresponding_table()
report.load_orders_with_seals()
report.replace_names()
report.set_dates()
report.save()
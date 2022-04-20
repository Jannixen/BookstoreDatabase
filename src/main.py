from dateutil.parser import parse
import warnings
from database_input import read_file
from output_reports import make_storage_report, make_sales_report, make_genres__publishers_report

warnings.filterwarnings("ignore")

try:
    read_file("input_file.xlsx")
    date_first_report = input("Podaj datę dla raportu o stanie magazynu:")
    make_storage_report(parse("24-12-2021").strftime('%Y-%m-%d %H:%M:%S'))
    date_second_report = input("Raport o stanie magazynu wygenerowany. Podaj datę dla raportu o sprzedaży miesięcznej:")
    make_sales_report(parse("2022-04-18").strftime('%Y-%m-%d %H:%M:%S'))
    print("Raport o miesięcznej sprzedaży wygenerowany.")
    make_genres__publishers_report()
    print("Raport statystyk ,,top sprzedaży'' wygenerowany.")
except Exception as e:
    print(e)

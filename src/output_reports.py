import datetime
from datetime import timedelta

import numpy as np
import pandas as pd
import pyodbc
from dateutil.parser import parse

from keys import SERVER, DATABASE, USERNAME, PASSWORD


def make_storage_report(report_date):
    cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL SERVER};SERVER='+SERVER+';DATABASE='+DATABASE+';UID='+USERNAME+';PWD='+ PASSWORD)
    sql = "Select * FROM [dbo].deliveries"
    data_deliveries = pd.read_sql(sql,cnxn)
    sql = "Select isbn, item_count, order_date FROM [dbo].order_items oi, [dbo].orders o WHERE o.id_order = oi.id_order"
    data_orders_items = pd.read_sql(sql,cnxn)
    results = {}
    grouped_deliveries = data_deliveries[data_deliveries["delivery_date"] <= report_date].groupby("isbn").sum()
    grouped_orders_items = data_orders_items[data_orders_items["order_date"] <= report_date].groupby("isbn").sum()
    concatenated = pd.merge(grouped_deliveries, grouped_orders_items, on=['isbn'], how="left")
    concatenated.replace({np.nan: 0}, inplace=True)
    concatenated['balance'] = concatenated['delivery_count'] - concatenated['item_count']
    writer = pd.ExcelWriter("StorageBalance%s.xlsx" % parse(report_date).strftime("%Y-%m-%d"))
    storage_report = concatenated['balance']
    storage_report.to_excel(writer, sheet_name="Storage balance per isbn", header=["Stan"], startcol = 0, startrow = 5)

    worksheet = writer.sheets["Storage balance per isbn"]
    worksheet.write(0, 0, 'Stan magazynu na dzień %s' %report_date)
    writer.close()


def make_sales_report(report_date):
    cnxn = pyodbc.connect(
        'DRIVER={ODBC Driver 17 for SQL SERVER};SERVER=' + SERVER + ';DATABASE=' + DATABASE + ';UID=' + USERNAME + ';PWD=' + PASSWORD)
    cursor = cnxn.cursor()
    cursor.execute(
        "Select o.id_order, DATEFROMPARTS(YEAR(o.order_date), MONTH(o.order_date), DAY(o.order_date)), oi.item_count, b.price FROM [dbo].orders o, [dbo].order_items oi, [dbo].books b WHERE o.id_order = oi.id_order AND oi.isbn = b.isbn AND o.order_date >= ? AND  o.order_date < DATEADD(month, 1, ?) ",
        report_date, report_date)
    data = [[j for j in i] for i in cursor.fetchall()]
    columns = ["id_order", "order_date", "item_count", "price"]

    order_items_with_price = pd.DataFrame(data=data, columns=columns)
    order_items_with_price['transaction_value'] = order_items_with_price['item_count'] * order_items_with_price['price']
    order_items_with_price.drop(columns=['item_count', 'price', 'id_order'], inplace=True)
    order_items_with_price = order_items_with_price.groupby("order_date").sum()
    order_items_with_price.reset_index(inplace=True)
    order_items_with_price['order_date'] = order_items_with_price['order_date'].apply(lambda x: str(x))

    dates = [parse(report_date) + timedelta(days=x) for x in range(31)]
    dates = [x.strftime('%Y-%m-%d') for x in dates]
    sales_days = pd.DataFrame(data=dates, columns=["order_date"])

    results = pd.merge(sales_days, order_items_with_price, on=['order_date'], how="outer")
    results.replace({np.nan: 0}, inplace=True)
    sum_sale_month = sum(results['transaction_value'])

    writer = pd.ExcelWriter("Sales%s.xlsx" % parse(report_date).strftime("%Y-%m-%d"))
    results.to_excel(writer, sheet_name="Sales", header=["Data", "Sprzedaż suma dzienna"], startcol=0, startrow=5)

    worksheet = writer.sheets["Sales"]
    worksheet.write(1, 0, 'Sprzedaż dla kolejnych 30 dni od %s' % report_date)

    worksheet.write(40, 2, 'Podsumowanie dla miesiąca: %f zł' % sum_sale_month)
    writer.close()


def make_genres__publishers_report():
    report_date = datetime.date.today()
    cnxn = pyodbc.connect(
        'DRIVER={ODBC Driver 17 for SQL SERVER};SERVER=' + SERVER + ';DATABASE=' + DATABASE + ';UID=' + USERNAME + ';PWD=' + PASSWORD)
    sql = "Select b.genre, SUM(oi.item_count * b.price) AS sales_total FROM [dbo].orders o, [dbo].order_items oi, [dbo].books b WHERE o.id_order = oi.id_order AND oi.isbn = b.isbn AND o.order_date > DATEADD(year,-1,GETDATE()) GROUP BY b.genre ORDER BY SUM(oi.item_count * b.price) DESC"
    data_genres = pd.read_sql(sql, cnxn)

    cnxn = pyodbc.connect(
        'DRIVER={ODBC Driver 17 for SQL SERVER};SERVER=' + SERVER + ';DATABASE=' + DATABASE + ';UID=' + USERNAME + ';PWD=' + PASSWORD)
    sql = "Select b.publisher, SUM(oi.item_count * b.price) AS sales_total FROM [dbo].orders o, [dbo].order_items oi, [dbo].books b WHERE o.id_order = oi.id_order AND oi.isbn = b.isbn AND o.order_date > DATEADD(year,-1,GETDATE()) GROUP BY b.publisher ORDER BY SUM(oi.item_count * b.price) DESC"
    data_publishers = pd.read_sql(sql, cnxn)

    cnxn = pyodbc.connect(
        'DRIVER={ODBC Driver 17 for SQL SERVER};SERVER=' + SERVER + ';DATABASE=' + DATABASE + ';UID=' + USERNAME + ';PWD=' + PASSWORD)
    sql = "Select b.title, SUM(oi.item_count * b.price) AS sales_total FROM [dbo].orders o, [dbo].order_items oi, [dbo].books b WHERE o.id_order = oi.id_order AND oi.isbn = b.isbn AND o.order_date > DATEADD(year,-1,GETDATE()) GROUP BY b.title ORDER BY SUM(oi.item_count * b.price) DESC"
    data_titles = pd.read_sql(sql, cnxn)

    writer = pd.ExcelWriter("GenresPublishers%s.xlsx" % str(report_date))
    data_genres.to_excel(writer, sheet_name="Genres", header=["Gatunek", "Sprzedaż suma łączna"], startcol=0,
                         startrow=5)

    worksheet = writer.sheets["Genres"]
    worksheet.write(1, 0, 'Sprzedaż na gatunki za ostatni rok stan na dzień %s' % report_date)

    data_publishers.to_excel(writer, sheet_name="Publishers", header=["Wydawca", "Sprzedaż suma łączna"], startcol=0,
                             startrow=5)

    worksheet = writer.sheets["Publishers"]
    worksheet.write(1, 0, 'Sprzedaż na wydawce  za ostatni rok stan na dzień %s' % report_date)

    data_titles.to_excel(writer, sheet_name="Titles", header=["Tytuł", "Sprzedaż suma łączna"], startcol=0, startrow=5)

    worksheet = writer.sheets["Titles"]
    worksheet.write(1, 0, 'Sprzedaż na wydawce za ostatni rok stan na dzień %s' % report_date)

    writer.close()
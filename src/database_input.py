import pandas as pd
import pyodbc
from dateutil.parser import parse

import keys


class BooksTableReadException(Exception):

    def __init__(self, s, idx):
        self.message = "Problem with book table input: " + s + " in row " + str(idx +1)
        super().__init__(self.message)


class DeliveriesTableReadException(Exception):

    def __init__(self, s, idx):
        self.message = "Problem with deliveries table input: " + s + " in row " + str(idx+1)
        super().__init__(self.message)


class ClientTableReadException(Exception):

    def __init__(self, s, idx):
        self.message = "Problem with clients table input: " + s + " in row " + str(idx+1)
        super().__init__(self.message)


class OrdersTableReadException(Exception):

    def __init__(self, s, idx):
        self.message = "Problem with orders table input: " + s + " in row " + str(idx+1)
        super().__init__(self.message)


class ReturnsTableReadException(Exception):

    def __init__(self, s, idx):
        self.message = "Problem with returns table input: " + s + " in row " + str(idx+1)
        super().__init__(self.message)


def read_file(input_file):
    cnxn = pyodbc.connect(
        'DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + keys.SERVER + ';DATABASE=' + keys.DATABASE + ';UID=' + keys.USERNAME + ';PWD=' + keys.PASSWORD)

    df_books = pd.read_excel(input_file, sheet_name="nowe_ksiazki")
    df_deliveries = pd.read_excel(input_file, sheet_name="dostawa")
    df_clients = pd.read_excel(input_file, sheet_name="nowi_klienci")
    df_orders = pd.read_excel(input_file, sheet_name="nowe_zamowienia")
    df_orders_items = df_orders
    df_orders = df_orders.drop_duplicates(subset=["Email", "Data i godzina"], keep='first')

    check_books(df_books)
    check_deliveries(df_deliveries)
    check_clients(df_clients)
    check_orders(df_orders)

    add_books_to_DATABASE(cnxn, df_books)
    add_deliveries_to_DATABASE(cnxn, df_deliveries)
    add_clients_to_DATABASE(cnxn, df_clients)
    add_orders_to_DATABASE(cnxn, df_orders)
    add_orders_items_to_DATABASE(cnxn, df_orders_items)


def check_books(df):
    for i, row in df.iterrows():
        if len(str(row['isbn'])) != 13:
            raise BooksTableReadException("ISBN should have 13 digits.", i)
        if pd.isna(row['tytuł']) or len(row['tytuł']) > 200:
            raise BooksTableReadException("Title should be not null and should be shorter than 200", i)
        if pd.isna(row['autor']) or len(row['autor']) > 100:
            raise BooksTableReadException("Author should be not null and should be shorter than 100", i)
        if not pd.isna(str(row['rok'])):
            try:
                date = parse(str(row['rok']))
            except ValueError:
                raise BooksTableReadException("Date format is incorrect.", i)
        if not pd.isna(row['opis']):
            if len(row['opis']) > 500:
                raise BooksTableReadException("Description too long", i)
        if not pd.isna(row['wydawca']):
            if len(row['wydawca']) > 50:
                raise BooksTableReadException("Publisher name too long", i)
        if pd.isna(row['cena']):
            raise BooksTableReadException("You have to give valid price", i)
        if not pd.isna(row['cena']):
            try:
                "{:.2f}".format(float(row['cena']))
            except ValueError:
                raise BooksTableReadException("Price format is incorrect.", i)
        if not pd.isna(row['gatunek']):
            if len(row['gatunek']) > 50:
                raise BooksTableReadException("Genre too long.", i)


def check_deliveries(df):
    for i, row in df.iterrows():
        if len(str(row['isbn'])) != 13:
            raise DeliveriesTableReadException("ISBN should have 13 digits.", i)
        try:
            int(row['Liczba sztuk'])
        except ValueError:
            raise DeliveriesTableReadException("Number of delivered products must be integer", i)
        if not pd.isna(str(row['Data'])):
            try:
                date = parse(str(row['Data']))
            except ValueError:
                raise DeliveriesTableReadException("Date format is incorrect.", i)


def check_clients(df):
    for i, row in df.iterrows():
        if pd.isna(row['Imię']) or len(row['Imię']) > 30:
            raise ClientTableReadException("Name should be not null and should be shorter than 30", i)
        if pd.isna(row['Nazwisko']) or len(row['Nazwisko']) > 30:
            raise ClientTableReadException("Surname should be not null and should be shorter than 30", i)
        if pd.isna(row['Email']) or len(row['Email']) > 40:
            raise ClientTableReadException("Email should be not null and should be shorter than 40", i)
        if not pd.isna(row['Adres']):
            if len(row['Adres']) > 60:
                raise ClientTableReadException("Adres should be shorter than 60", i)
        if not pd.isna(row['Telefon']):
            if len(str(row['Telefon'])) > 12:
                raise ClientTableReadException("Adres should be shorter than 12", i)


def check_orders(df):
    for i, row in df.iterrows():
        if not pd.isna(str(row['Data i godzina'])):
            try:
                date = parse(str(row['Data i godzina']))
            except ValueError:
                raise OrdersTableReadException("Date format is incorrect.", i)
        if not pd.isna(row['Adres']):
            if len(row['Adres']) > 60:
                raise OrdersTableReadException("Adres should be shorter than 60", i)
        if not pd.isna(row['Telefon']):
            if len(str(row['Telefon'])) > 12:
                raise OrdersTableReadException("Adres should be shorter than 12", i)
        try:
            int(row['Liczebność'])
        except ValueError:
            raise OrdersTableReadException("Order item count must be integer", i)


def add_books_to_DATABASE(cnxn, df_books):
    cursor = cnxn.cursor()
    try:
        for index, row in df_books.iterrows():
            cursor.execute(
                "INSERT INTO [dbo].books (isbn, title, author,published_year, descr, publisher, price, genre) values(?,?, ?,?,?,?,?,?)",
                row['isbn'], row['tytuł'], row['autor'], parse(str(row['rok'])).year, str(row['opis']),
                str(row['wydawca']), row['cena'], str(row['gatunek']))
        cnxn.commit()
    except pyodbc.Error as ex:
        sqlstate = ex.args[0]
        if sqlstate == '23000':
            raise BooksTableReadException("Book with this ISBN already in the DATABASE", index)
        else:
            print(ex.args)
    finally:
        cursor.close()


def add_deliveries_to_DATABASE(cnxn, df_deliveries):
    cursor = cnxn.cursor()
    try:
        for index, row in df_deliveries.iterrows():
            cursor.execute("INSERT INTO [dbo].deliveries (isbn, delivery_count, delivery_date) values(?,?, ?)",
                           str(row['isbn']), int(row['Liczba sztuk']),
                           parse(str(row['Data'])).strftime('%Y-%m-%d %H:%M:%S'))
        cnxn.commit()
    except pyodbc.Error as ex:
        sqlstate = ex.args[0]
        print(ex.args)
        raise DeliveriesTableReadException("Problem with Delivery data", index)
    finally:
        cursor.close()


def add_clients_to_DATABASE(cnxn, df_clients):
    cursor = cnxn.cursor()
    try:
        for index, row in df_clients.iterrows():
            if not pd.isna(row['Telefon']):
                cursor.execute(
                    "INSERT INTO [dbo].clients (name_client, surname_client, client_email, client_addres, client_phone) values(?,?, ?, ?, ?)",
                    str(row['Imię']), str(row['Nazwisko']), str(row['Email']), str(row['Adres']), row['Telefon'])
            else:
                cursor.execute(
                    "INSERT INTO [dbo].clients (name_client, surname_client, client_email, client_addres) values(?,?, ?, ?)",
                    str(row['Imię']), str(row['Nazwisko']), str(row['Email']), str(row['Adres']))
        cnxn.commit()
    except pyodbc.Error as ex:
        sqlstate = ex.args[0]
        if sqlstate == '23000':
            raise ClientTableReadException("Client with this Email already in the DATABASE", index)
        else:
            print(ex.args)
            raise ClientTableReadException("Problem with Client data", index)
    finally:
        cursor.close()


def add_orders_to_DATABASE(cnxn, df_orders):
    cursor = cnxn.cursor()
    try:
        for index, row in df_orders.iterrows():
            cursor.execute("SELECT id_client FROM [dbo].clients where client_email = ?", str(row['Email']))
            results = cursor.fetchall()
            if not results:
                raise OrdersTableReadException("Email not in the DATABASE ", index)
            for res in results:
                id_client = res[0]
            cursor.execute(
                "INSERT INTO [dbo].orders (id_client, order_date, order_addres, order_phone) values(?,?,?, ?)",
                id_client, parse(str(row['Data i godzina'])).strftime('%Y-%m-%d %H:%M:%S'), row['Adres'],
                row['Telefon'])
        cnxn.commit()
    except pyodbc.Error as ex:
        print(ex)
        raise OrdersTableReadException("Problem with order table", index)
    finally:
        cursor.close()


def add_orders_items_to_DATABASE(cnxn, df_orders_items):
    cursor = cnxn.cursor()
    try:
        for index, row in df_orders_items.iterrows():
            cursor.execute("SELECT id_client FROM [dbo].clients where client_email = ?", str(row['Email']))
            results = cursor.fetchall()
            if not results:
                raise OrdersTableReadException("Email not in the DATABASE ", index)
            for res in results:
                id_client = res[0]
            cursor.execute("SELECT id_order FROM [dbo].orders where id_client = ? and order_date = ?", id_client,
                           parse(str(row['Data i godzina'])).strftime('%Y-%m-%d %H:%M:%S'))
            results = cursor.fetchall()
            if not results:
                raise OrdersTableReadException("Client not in the DATABASE ", index)
            for res in results:
                id_order = res[0]
            cursor.execute("INSERT INTO [dbo].order_items (isbn, item_count, id_order) values(?,?,?)", row['isbn'],
                           row['Liczebność'], id_order)
        cnxn.commit()
    except pyodbc.Error as ex:
        print(ex)
        raise OrdersTableReadException("Problem with order table", index)
    finally:
        cursor.close()

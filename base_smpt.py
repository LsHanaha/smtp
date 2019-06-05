# -*- coding: utf-8 -*-

import time
import sqlite3
from sqlite3 import Error
from datetime import datetime

database = r"C:/Users/Kirill/Desktop/smtp/smtp.db"

def create_connection(db_file):
    """ create a database connection to a SQLite database """
    try:
        conn = sqlite3.connect(db_file)
        return conn
    except Error as e:
        with open("log.txt", 'a') as log:
            log.write(str(e))
    return None


def create_table(conn, create_table_sql):
    """ create a table from the create_table_sql statement
    :param conn: Connection object
    :param create_table_sql: a CREATE TABLE statement
    :return:
    """
    try:
        c = conn.cursor()
        c.execute(create_table_sql)
    except Error as e:
        with open("log.txt", 'a') as log:
            log.write(str(e))

###################################################################################################################

def create_tables():
    """
    creating (if needed) and filling db
    :return: None
    """
    sql_create_birthday_table = """ CREATE TABLE IF NOT EXISTS persons (
                                        id integer PRIMARY KEY,
                                        name text NOT NULL,
                                        date text NOT NULL,
                                        email text NOT NULL
                                    ); """

    # create a database tables
    conn = create_connection(database)
    if conn is not None:
        create_table(conn, sql_create_birthday_table)
    else:
        with open("log.txt", 'a') as log:
            log.write("Error! cannot create the database connection.")


def add_new_person(conn, name):

    sql = ''' INSERT INTO persons (name, date, email)
              VALUES(?,?,?) '''
    cur = conn.cursor()
    cur.execute(sql, name)
    return cur.lastrowid

def select_in_persons(conn, today_date):
    sql = "SELECT name, email FROM persons WHERE date LIKE '%'||?||'%'"
    cur = conn.cursor()
    res = cur.execute(sql, (today_date,))
    return res


def select_all_persons():
    conn = create_connection(database)
    today_date = datetime.today().strftime("%m-%d")
    with conn:
        request = select_in_persons(conn, today_date)
        res = []
        for i in request:
            res.append(i)
        print(res)
        return res


def fill_games():
    conn = create_connection(database)
    with conn:
        create_tables()
        names = (
                ("john hohnson", "1999-06-05", "kirik193@yandex.ru"),
                ("bill billson", "2000-10-11", "kirik193@yandex.ru"),
                ("Tom Tomasson", "2001-06-04", "kirik193@yandex.ru"),
                ("Corey Taylor", "2002-10-11", "kirtrishin@gmail.com"),
                ("Who Whoser", "1800-06-05", "kirtrishin@gmail.com"),
                )
        for name in names:
            add_new_person(conn, name)

def start():
    create_tables()

if __name__ == "__main__":
    start()
    fill_games()
    select_all_persons()

import mysql.connector

mydb = mysql.connector.connect(
    host="localhost",
    user="root",
    password="tik-71.Lion",
    database="inventory",
    )

my_cursor = mydb.cursor()

sql_formula = "INSERT INTO revisia (goods, quantity, price, good_total) VALUES (%s, %s, %s, %s)"







# import mysql.connector
#
# mydb = mysql.connector.connect(
#     host="localhost",
#     user="root",
#     password="tik-71.Lion",
#     database="inventory",
#     )
#
# my_cursor = mydb.cursor()
#
# sql_formula = "INSERT INTO revisia (good, quantity, price, good_total) VALUES (%s, %s, %s, %s)"

import random
import mysql.connector

mydb = mysql.connector.connect(
    host="localhost",
    user="root",
    password="tik-71.Lion",
    database="accommodation",
    )

my_cursor = mydb.cursor()

sql_formula_hot = "INSERT INTO booking_hotel (nickname, country, city) VALUES (%s, %s, %s)"
sql_formula_ap = "INSERT INTO booking_apartment (nickname, country, city, price)" \
                  " VALUES (%s, %s, %s, %s)"
sql_formula_rt = "INSERT INTO booking_roomtype(name, price, hotel_id)" \
                  " VALUES (%s, %s, %s)"


names = {"hotel": ["Hotel Royal", "Hotel Inter", "Hotel Ambassador"], "apartment": ["Best View", "Worst View"]}
countries = [("UK", "London"), ("Bulgaria", "Sofia"), ("Italy", "Milan"), ("France", "Paris")]
prices = [60.00, 80.00, 100.00, 150.00, 200.00]
room_types = ["single", "double"," tripple", "lux", "superior"]
rooms = [103, 104, 105, 106, 107]
n = 4

def create_test_properties(n):
    for i in range(n):

        property_type = random.choice(list(names.keys()))
        nickname = random.choice(names[property_type])
        c = random.choice(countries)
        country, city = c[0], c[1]
        if property_type == "hotel":
            val = (nickname, country, city)
            my_cursor.execute(sql_formula_hot, val)
        else:
            price = random.choice(prices)
            val = (nickname, country, city, price)
            my_cursor.execute(sql_formula_ap, val)
        mydb.commit()

    my_cursor.execute("SELECT * FROM booking_hotel")
    my_result = my_cursor.fetchall()
    for i in my_result:
            print(i)

    my_cursor.execute("SELECT * FROM booking_apartment")
    my_result = my_cursor.fetchall()
    for i in my_result:
            print(i)

create_test_properties(10)

def create_room_type():
    for i in range(4):
        list_ids = []
        price_type = random.choice(list(zip(prices, room_types)))
        my_cursor.execute("Select id from booking_hotel")
        my_result = my_cursor.fetchall()
        for i in range(len(my_result)):
            list_ids.append(my_result[i][0])
            hotel = random.choice(list_ids)
        val = (price_type[1], price_type[0], hotel)
        my_cursor.execute(sql_formula_rt, val)
        mydb.commit()
    my_cursor.execute("SELECT * FROM booking_roomtype")
    my_result = my_cursor.fetchall()
    for i in my_result:
        print(i)

# create_room_type()








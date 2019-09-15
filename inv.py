import os
import tkinter as tk
from tkinter import messagebox
from tkinter.filedialog import askopenfilename, asksaveasfilename
from openpyxl import Workbook, utils, load_workbook
from openpyxl.styles import PatternFill, Font
import logging
# import MySQL

# fetch current/working directory
directory = str(os.path.dirname(os.path.abspath(__file__)))

# logger
format_log = "%(levelname)s %(asctime)s - %(message)s"
logging.basicConfig(filename=directory + "\loging.txt",
                    level=logging.DEBUG, format=format_log, filemode='w')
logger = logging.getLogger()


class Goods:
    def __init__(self, name, quantity, price):
        self.name = name
        self.quantity = quantity
        self.price = price
        self.total_price = self.quantity * self.price


good_list = []
good_list_names = []
# list of good objects names for quicker reference to check if good exists




def select_dest_file():
    """Selects file to save the current session operations. """
    logger.info("Selecting destination file(select_dest_file)")
    file = asksaveasfilename(filetypes=[("excel", ".xlsx")], initialdir=directory,
                             initialfile=".xlsx")
    if file:
        root.title(file)
        logger.info(f"Selected destination file (select_dest_file) - {file}  ")
        return file
        # if Canceled nothing happens


def close_with_x():
    """ Closing by closing main window by clicking on 'X' """
    if not ((entries[(1, 1)].get() and entries[(2, 1)].get()) or good_list):
        # If no data is on screen and no goods are created exits
        logger.info("If no data is on screen and no goods are created exits")
        exit()
    else:
        if check_on_screen() == "Good to be entered":
            enter_good()
            logger.info("Exiting data exported to excel(close_with_x)")
            return "Good entered"
        elif good_list:
            logger.info("No onscreen entry but goods in the list (close_with_x)")
            write_or_add()
            return "Goods in the list entered no onscreen entered (close_with_x)"


def write_or_add():
    """Selecting the output file new or existing"""
    logger.info("Selecting the output file new or existing (write_or_add)")
    if root.title() == "Програма за ревизия":
        # destination file is not preselected
        logger.info("Writing the data to excel before closing with button (write_or_add)")
        write_to_excel()
            # create/select file and write the data (exiting data will be over written)
        return "written to excel"

    else:
        add_to_excel()
        # destination file is not preselected
        logger.info("Adding the data to excel before closing with button(write_or_add)")
        return "added to excel"


def close_with_button():  # event is given by pushing the button 'Запази и излез'
    """Closing with button 'Запази и излез'"""
    logger.info("starting close with button (close_with_button)")
    x = check_on_screen()
    if x == "No data" and good_list != []:  # No data onscreen and no goods entered yet
        write_or_add()
        return "Closing with button no onscreen"
        logger.info("Closing with button no onscreen (close_with_button)")
    elif x == "Saving onscreen rejected" and good_list != []:
        # Rejected entering onscreen data and no goods entered yet++
        write_or_add()
        return "Saving onscreen rejected"
        logger.info("Saving onscreen rejected(close_with_button)")
    elif x == "Good to be entered":
        # Entering onscreen confirmed:
        enter_good()
        write_or_add()
        return "onscreen entered"
        logger.info("Saving onscreen entered(close_with_button)")
    elif x == "Not valid data on screen":  # Onscreen entry fails due to invalid data
        logger.info("Onscreen entry fails due to invalid data(close_with_button)")
        return
    else:  # No onscreen data but goods are entered already
        write_or_add()
        logger.info("No onscreen entered the goods in the list(close_with_button)")


def check_on_screen():
    if entries[1, 1].get() and entries[2, 1].get():
    # checks if there are unsaved numbers on the screen
        logger.info("Unsaved data found on the screen (check_on_screen)")
        if messagebox.askyesno(message="Искате ли да "
                                       "запазим операцията на екрана?"):
            logger.info("Saving to file selected (check_on_screen)")
            if check_if_entries_are_numbers():
                logger.info("Data is valid and good is entered (check_on_screen)")
                # checks if the entered info is valid ( the numbers are floats) and creates goods
                return "Good to be entered"
            else:
                highlight_wrong()
                return "Not valid data on screen"
        else:
            return "Saving onscreen rejected"
    else:
        return "No data"


def check_if_entries_are_numbers():
    try:
        quantity = float(entries[1, 1].get().replace(",", "."))
        price = float(entries[2, 1].get().replace(",", "."))
        if quantity>10000.00:
            messagebox.showinfo(message="Максимално количество на операция е 1000000,00")
            logger.info("Quantity over 10000,00")
            return False
        elif price > 1000000.00:
            messagebox.showinfo(message="Максимално единична цена е 1000000,00")
            logger.info("Price over 1000000,00")
            return False
        return True
    except ValueError:
        messagebox.showinfo(title="Резултат от операция", message="Полета количество и "  
                            "цена трябва да са валидни числа")
        logger.info("not valid numbers found (check_if_entries_are_numbers)")
        highlight_wrong()
        return False


def highlight_wrong():
    # highlights the incorrect entry
    logger.info("highlights the incorrect entry(highlight_wrong)")
    for q in entries.keys():
        if entries[q].get():
            try:
                float(entries[q].get()) < 1000000
                entries[q].configure(highlightthickness="0")
            except ValueError:
                entries[q].configure(highlightthickness="2",
                                     highlightcolor="red", highlightbackground="red")
            entries[0, 1].configure(highlightthickness="0")
    return "Wrong entries highlighted"


def on_return(event):
    # moves to next entries.
    logger.info("moves to next entries (on_return_entries")
    en = root.focus_get()
    for q in entries:
        if en == entries[q]:
            n = q[0]
            # current row of the active entry found
            logger.info("current row of the active entry found (on_return_entries")
            if n+1 < (len(entries)):
                # checks if not the last entry
                entries[n, 1].event_generate("<<TraverseOut>>")
                entries[n + 1, 1].focus()
                entries[n + 1, 1].event_generate("<<TraverseIn>>")
                # moved to next entries.
                logger.info("moved to next entries (on_return_entries")
                return entries[n + 1, 1]

            else:
                # it is the last enrty
                entries[n, 1].event_generate("<<TraverseOut>>")
                entries[0, 1].focus()
                entries[0, 1].event_generate("<<TraverseIn>>")
                enter_good()
                # moved to first entries and entered the operation
                logger.info("moved to first entries and entered the goods(*if valid numbers), (on_return_entries)")
                return entries[0, 1]


def set_focus_on_grid_slave(row, column):
    # used to set focus when moving with arrows
    for wid in (root.grid_slaves()):
        if wid.grid_info()["row"] == row and wid.grid_info()["column"] == column:
            wid.focus_set()
            return wid


def arrow_moves(event):
    column = event.widget.grid_info()["column"]
    row = event.widget.grid_info()["row"]
    # fetched coordinates of the event widget
    logger.info("fetched coordinates of the event widget (arrow_moves)")
    if event.keysym == "Left":
        if column > 1:
            column -= 1
            # moved to left as it is not the first column
            logger.info("moved to left as it is not the first column(arrow_moves)")

    elif event.keysym == "Right":
        if column < root.grid_size()[0]-1:
            column += 1
            # moved to Right as it is not the last column
            logger.info("moved to Right as it is not the last column(arrow_moves)")

    elif event.keysym == "Up":
        if row > 0:
            row -= 1
            # moved up  as it is not the top row
            logger.info("moved up  as it is not the top row(arrow_moves)")

    elif event.keysym == "Down":
        if row < 2:
            row += 1
            # moved Down bottom row
            logger.info("moved up  as it is not the bottom row(arrow_moves)")

    set_focus_on_grid_slave(row, column)
    return row, column


def enter_good():
    """Enters good not existing in inventory or adds quantity to existing"""
    if entries[1, 1].get() and entries[(2, 1)].get():
        logger.info("Quantity and price fields contain info (enter_good)")
        if check_if_entries_are_numbers() is False:
            logger.info("Entry of good failed as the numbers are not valid(enter_good)")
            return "Failed"
        else:
            # creates the good or adds quantity to existing
            create_good()
            logger.info("Good entered successfully(enter_good)")
            return "Good entered"
        for q in range(len(entries)):
            entries[q, 1].delete(0, "end")
            entries[q, 1].configure(highlightthickness="0")
        logger.info("Formatting and content cleared (enter_good)")


def create_good():
    if entries[0, 1].get():
        logger.info("Checking for name entry  (create_good)")
        # checks if name will be entered
        a = entries[0, 1].get()
        if a not in good_list_names:
            logger.info("Checking if name exist  (create_good)")
            good_list_names.append(a)
            a = Goods(a, float(entries[1, 1].get().replace(",", ".")),
                      float(entries[(2, 1)].get().replace(",", ".")))
            good_list.append(a)
            logger.info("Good created and name added to list(create_good)")
            m = f" created {a.name, a.quantity, a.price, a.total_price}"

        else:
            q = add_quantity()
            logger.info("Quantity added for existing good  (create_good)")
            m = q

    else:
        good = Goods("No Name good", float(entries[(1, 1)].get().replace(",", ".")),
                     float(entries[(2, 1)].get().replace(",", ".")))
        good_list.append(good)
        logger.info("Good 'No Name good' created and name added to list (create_good)")
        m = f" No name created {good.name, good.quantity, good.price, good.total_price}"

    for q in range(len(entries)):
        entries[q, 1].delete(0, "end")
        logger.info(f"Entry box {entries[q, 1]} cleaned (create_good)")

    return m


def add_quantity():
    a = entries[0, 1].get()
    b = entries[1, 1].get()
    if v.get() == 1:
        # depending on choice accumulate or not the quantity for existing good.
        logger.info("Adding accumulated quantity(add_quantity)")
        for i in good_list:
            logger.info(f"checking {i.name}")
            if a == i.name:
                i.quantity = float(b) + i.quantity
                i.total_price = i.quantity * i.price
                logger.info(f'checked {i}')
        return f"added {i.name, i.quantity, i.price, i.total_price}"

    else:
        # creating same good but found in a different place as non
        # accumulating option is not selected
        logger.info("creating same good, found in a different place"
                    " as non accumulating option is selected(add_quantity)")
        good_list_names.append(a)
        a = Goods(a, float(entries[1, 1].get().replace(",", ".")),
                  float(entries[(2, 1)].get().replace(",", ".")))
        good_list.append(a)
        return f"added {a.name, a.quantity, a.price, a.total_price}"


def calculate_total():
    """calculate totals of all goods"""
    product_total = 0
    for i in good_list:
        logger.info(f"checking {i}")
        product_total += i.total_price
    return product_total


def create_file_structure():

    wb = Workbook()
    ws = wb.active
    # ws.column_dimensions.width = 180

    ws.title = "Inventory"
    ws["A1"] = "РЕВИЗИЯ"
    ws["B1"] = "ДАТА:"
    font_1 = Font(name="Calibri", size=11, color="FFFF00")
    fill_1 = PatternFill(fill_type="solid")

    ws["A1"].font = font_1
    ws["B1"].font = font_1
    ws["A1"].fill = fill_1
    ws["B1"].fill = fill_1

    list_head_names = ["СТОКА", "БРОЙ", "ЦЕНА", "ОБЩО", "РЕЗУЛТАТ"]
    font_2 = Font(name="Calibri", size=20, underline="single")
    fill_2 = PatternFill(fill_type="solid", start_color="ADFF2F")
    # setting column and headings and patterns
    logger.info("setting column and headings and patterns (create_file_structure)")
    for r in (ws["A2":"E2"]):
        for n, c in enumerate(r):
            c.fill = fill_2
            c.font = font_2
            c.value = list_head_names[n]
    logger.info("Loop for formatting and header names (create_file_structure)")

    ws["E3"] = "=SUM(D3:D526)"
    ws["E3"].font = Font(name="Calibri", size=28)
    for i, r in enumerate(ws["A3":"E3"]):
        for n, c in enumerate(r):
            c.fill = PatternFill(fill_type="solid", start_color="B22222")

    ws["F4"] = "СТАРА РЕВИЗИЯ"
    ws["F6"] = "СТОКОВИ ОТ КОЧАН"
    ws["F8"] = "СТОКОВИ ОТ КОМПЮТЪР"
    ws["F10"] = "ОТЧЕТИ"
    ws["F12"] = "НОВА РЕВИЗИЯ"
    ws["F13"] = "=E3"
    ws["F14"] = "КРАЕН РЕЗУЛТАТ"
    ws["F14"].font = Font(name="Calibri", size=16, underline="single", italic=True, bold=True )
    ws["F15"] = "=D5+D7+D9-D13-D11"
    ws["F15"].font = Font(name="Calibri", size=48)
    ws["F15"].fill = fill_2
    logger.info("filling formulas (create_file_structure)")
    for i, r in enumerate(ws["F4":"F14"]):
        if i % 2 == 0:
            for n, c in enumerate(r):
                c.fill = fill_1
                c.font = font_1
    logger.info("Loop for column F (create_file_structure)")
    for i, r in enumerate(ws["C4":"C100"]):
        for n, c in enumerate(r):
            if c.value:
                c.fill = fill_2
    logger.info("Loop for column C (create_file_structure)")
    for col in ws.columns:
        max_length = 0
        column = utils.get_column_letter(col[0].column)  # Get the column name
        for cell in col:
            try:  # Necessary to avoid error on empty cells
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 5) * 1.2
        ws.column_dimensions[column].width = adjusted_width
    logger.info("Loop for column width (create_file_structure)")

    for i in range(len(good_list)):
        n = str(i+4)
        ws[("A" + str(n))] = good_list[i].name
        ws[("B" + str(n))] = good_list[i].quantity
        ws["C" + str(n)] = good_list[i].price
        ws["D" + str(n)] = good_list[i].quantity*good_list[i].price
    logger.info("Loop goods attribute (create_file_structure)")
    logger.info("Saving part of the function (create_file_structure)")
    return wb


def write_to_excel():
    """ writes the data in a new file, ir there is existing data it is overwritten"""
    # write_to_database()
    if root.title() != "Програма за ревизия":
        # Destination is preselcted
        logger.info("Destination is preselcted (write_to_excel)")
        add_to_excel()
    else:
        file = select_dest_file()
        # Destination is selected
        logger.info("Destination is selected (write_to_excel)")
        if file:
            try:
                create_file_structure().save(file)
                logger.info("Check if file is open 2 (write_to_excel)")
            except :
                FileExistsError
                messagebox.showinfo(message="Файлът е отворен, затворете го за да запазите операциите ")
                logger.info("Prompted to close the file 2 (write_to_excel)")

        else:
            return "Saving canceled"
            logger.info("Canceled (write_to_excel")

    os.startfile(file)
    logger.info("Closing program (write_to_excel)")
    if __name__ == "__main__":
        exit()


# def write_to_database():
#     """adds the goods data to MySQL"""
#     MySQL.my_cursor.execute("DROP TABLE IF EXISTS revisia ")
#     MySQL.my_cursor.execute("CREATE TABLE revisia (name VARCHAR(100), "
#                                     "quantity FLOAT(20,2), price FLOAT(20,2), total_price FLOAT (255,2))")
#     for good in good_list:
#         entr = (good.name, good.quantity, good.price, good.total_price)
#         sqlf = "INSERT INTO revisia (name, quantity, price, total_price) VALUES(%s, %s, %s, %s)"
#         MySQL.my_cursor.execute(sqlf, entr)

    MySQL.mydb.commit()


def add_to_excel():
    """Adding to excel - keeping old records. """
    # write_to_database()
    file = root.title()
    wb = load_workbook(file)
    ws = wb.active
    if ws["A1"].value != "РЕВИЗИЯ":  # checks if file already have the structure
        create_file_structure()
    n = 4  # the first row to have entries after headers

    for r in ws.iter_rows(min_row=4, max_col=1):
        for c in r:  # c is cell, r is row, checking only first column as this shows if entries on the row.
            logger.info(f"Checking cell {c}, row {r} (ws.iter_rows)")
            if c.value:
                n += 1
            else:
                break  # no filled cell on the row
        else:
            logger.info("Reached continue (ws.iter_rows)")
            continue
        logger.info("Reached final break (ws.iter_rows)")
        break
    logger.info("first empty row found ")

    for i in range(len(good_list)):
        ws[("A" + str(n))] = good_list[i].name
        ws[("B" + str(n))] = good_list[i].quantity
        ws["C" + str(n)] = good_list[i].price
        ws["D" + str(n)] = good_list[i].quantity*good_list[i].price
        n += 1
    logger.info(f"added rows and info in a loop starting row {n-1}")

    try:
        wb.save(file)
        logger.info("Check if file is open 2 (add_to_excel)")
    except FileExistsError:
        messagebox.showinfo(message="Файлът е отворен затворете за да запазите операциите ")
        logger.info("Prompted to close the file 2 (add_to_excel)")
        return
    os.startfile(file)
    logger.info("Closing program (add_to_excel)")
    if __name__ == "__main__":
        exit()

# GUI part

root = tk.Tk()

w = 605  # width for the Tk root
h = 180  # height for the Tk root

# get screen width and height
ws = root.winfo_screenwidth()  # width of the screen
hs = root.winfo_screenheight()  # height of the screen

# calculate x and y coordinates for the Tk root window
x = (ws/2) - (w/2)
y = (hs/2) - (h/2 + 100)

# set the dimensions of the screen
# and where it is placed
root.geometry('%dx%d+%d+%d' % (w, h, x, y))

root.title("Програма за ревизия")
Logo_photo = tk.PhotoImage(file=directory + "\Logo.png")

v = tk.IntVar()  # need to set the quantities with/wo accumulation

labels = {}
entries = {}
buttons = {}
# Dictionaries with tuple keys used for referencing the grid location
for i in range(3):
    label_texts = ["Стока", "Количество", "Цена", " "]
    lab = tk.Label(root, font=("Courier", 12), width=12, fg="black", anchor="e",
                   text=label_texts[i])
    lab.grid(row=i, column=0, sticky="W")
    labels[i, 0] = lab

    e = tk.Entry(root, width=15, font=("Courier", 12), fg="black", bg="white")
    e.grid(row=i, column=1, sticky="W")
    entries[i, 1] = e
# created the labels and entries
logger.info('created the labels and entries')
for i in range(2):
    texts = ["Запази операция", "Запази и излез", " "]
    commands = [enter_good, close_with_button, None]
    b = tk.Button(root, command=commands[i], font=("Courier", 10), width=25, bg="grey15", fg="white",
                  text=texts[i], activeforeground="grey15")
    b.grid(row=i, column=2, sticky="W")
    buttons[i, 2] = b
# created the buttons
logger.info('created the buttons')

label_space = tk.Label(root, anchor="e").grid(row=4, column=0, columnspan=2, sticky="W")
label_logo = tk.Label(root, image=Logo_photo)
label_logo.grid(row=5, column=0, columnspan=2, sticky="W")
label_rights_reserved = tk.Label(root, font=("Courier", 8), fg="black", anchor="e",
                                 text="© 2019, JM ware. All Rights Reserved")
label_rights_reserved.grid(row=5, column=1, columnspan=2, sticky="E")

# created the space labels and images
logger.info('created the space lables and images ')

entries[0, 1].focus()
entries[0, 1].event_generate("<<TraverseIn>>")

# initial focus set
logger.info('initial focus set')

root.bind("<Return>", on_return)
root.bind("<Left>", arrow_moves)
root.bind("<Up>", arrow_moves)
root.bind("<Down>", arrow_moves)
root.bind("<Right>", arrow_moves)

# Bound enter and arrows to root actions
logger.info('Bound enter and arrows to root actions')

root.protocol("WM_DELETE_WINDOW", close_with_x)

# Menu

menubar = tk.Menu(root)

# create a pulldown menu, and add it to the menu bar
filemenu = tk.Menu(menubar, tearoff=0)
open_image = tk.PhotoImage(file=directory + "\\" + "open.png")

save_image = tk.PhotoImage(file=directory + "\\" + "save.png")
filemenu.add_command(label="Избери файл за запзване нa операциите", image=save_image,
                     compound= "left", command=select_dest_file)

filemenu.add_checkbutton(label="количества с натрупване", compound="left", variable = v)

filemenu.add_separator()
filemenu.add_command(label="Exit", command=root.destroy)
menubar.add_cascade(label="File", menu=filemenu)

# create more pulldown menus

editmenu = tk.Menu(menubar, tearoff=0)

menubar.add_cascade(label="Edit", menu=editmenu)

helpmenu = tk.Menu(menubar, tearoff=0)
editmenu.add_command(label="Отвори файл",  image=open_image,
                     compound="left", command=lambda: os.startfile(askopenfilename(filetypes=
                     [("excel", ".xlsx")], initialdir=directory)))
menubar.add_cascade(label="Help", menu=helpmenu)

# display the menu
root.config(menu=menubar)
logger.info("GUI is ready to loop")

if __name__ == "__main__":
    root.mainloop()









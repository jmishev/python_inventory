import logging
import os
import sys
from tkinter.filedialog import askopenfilename

import inv
import tkinter as tk

directory = inv.directory
logger = inv.logger


def on_return(event):
    # moves to next entries.
    logger.info("moves to next entries (on_return_entries")
    en = root.focus_get()
    for q in inv.entries:
        if en == inv.entries[q]:
            n = q[0]
            # current row of the active entry found
            logger.info("current row of the active entry found (on_return_entries")

            if n+1 < (len(inv.entries)):
                # checks if not the last entry
                inv.entries[n, 1].event_generate("<<TraverseOut>>")
                inv.entries[n + 1, 1].focus()
                inv.entries[n + 1, 1].event_generate("<<TraverseIn>>")
                # moved to next entries.
                logger.info("moved to next entries (on_return_entries")
                return inv.entries[n + 1, 1]

            else:
                # it is the last enrty
                inv.entries[n, 1].event_generate("<<TraverseOut>>")
                inv.entries[0, 1].focus()
                inv.entries[0, 1].event_generate("<<TraverseIn>>")
                inv.enter_good()
                # moved to first entries and entered the operation
                logger.info("moved to first entries and entered the goods(*if valid numbers), (on_return_entries)")
                return inv.entries[0, 1]


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

# GUI part

if __name__ == "__main__":

    root = tk.Tk()
    inv.root = root

    w = 505  # width for the Tk root
    h = 170  # height for the Tk root #TODO not looking good on Win10

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
    Logo_photo = tk.PhotoImage(file=directory + "/Logo.PNG")

    v = tk.IntVar()  # need to set the quantities with/wo accumulation
    inv.v = v

    for i in range(3):
        label_texts = ["Стока", "Количество", "Цена", " "]
        lab = tk.Label(root, font=("Courier", 12), width=12, fg="black", anchor="e",
                       text=label_texts[i])
        lab.grid(row=i, column=0, sticky="W")
        inv.labels[i, 0] = lab

        e = tk.Entry(root, width=15, font=("Courier", 12), fg="black", bg="white")
        e.grid(row=i, column=1, sticky="W")
        inv.entries[i, 1] = e
    # created the labels and entries
    logger.info('created the labels and entries')
    for i in range(2):
        texts = ["Запази операция", "Запази и излез", " "]
        commands = [inv.enter_good, inv.close_with_button, None]
        b = tk.Button(root, command=commands[i], font=("Courier", 10), width=25, bg="grey15", fg="white",
                      text=texts[i], activeforeground="grey15")
        b.grid(row=i, column=2, sticky="W")
        inv.buttons[i, 2] = b
    # created the buttons
    logger.info('created the buttons')

    label_space = tk.Label(root, anchor="e").grid(row=4, column=0, columnspan=2, sticky="W")
    # label_logo = tk.Label(root, image=Logo_photo)
    # label_logo.grid(row=5, column=0, columnspan=2, sticky="W")
    label_rights_reserved = tk.Label(root, font=("Courier", 8), fg="black", anchor="e",
                                     text="© 2019, JM ware. All Rights Reserved")
    label_rights_reserved.grid(row=5, column=1, columnspan=2, sticky="E")

    # created the space labels and images
    logger.info('created the space lables and images ')

    inv.entries[0, 1].focus()
    inv.entries[0, 1].event_generate("<<TraverseIn>>")

    # initial focus set
    logger.info('initial focus set')

    root.bind("<Return>", on_return)
    root.bind("<Left>", arrow_moves)
    root.bind("<Up>", arrow_moves)
    root.bind("<Down>", arrow_moves)
    root.bind("<Right>", arrow_moves)

    # Bound enter and arrows to root actions
    logger.info('Bound enter and arrows to root actions')

    root.protocol("WM_DELETE_WINDOW", inv.close_with_x)

    # Menu

    menubar = tk.Menu(root)

    # create a pulldown menu, and add it to the menu bar
    filemenu = tk.Menu(menubar, tearoff=0)
    open_image = tk.PhotoImage(file=directory + "\\" + "open.png")

    save_image = tk.PhotoImage(file=directory + "\\" + "save.png")
    filemenu.add_command(label="Избери файл за запзване нa операциите", image=save_image,
                         compound= "left", command= inv.select_dest_file)

    filemenu.add_checkbutton(label="количества с натрупване", compound="left", variable=v)

    filemenu.add_separator()
    filemenu.add_command(label="Exit", command=sys.exit)
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

    root.mainloop()

"""
GUI FOR AUTOMATION
"""
import tkinter as tk
from tkinter import filedialog


class Automation:
    """Class to create GUI"""

    def __init__(self):
        """Initializing master variable"""
        self.master = None

    def fun1(self, root):
        """function to create window"""
        self.master = root
        root.configure(background='black')
        root.title("Calendar Automation")
        tk.Label(root,
                 text='', bg="black", fg="black"
                 ).pack()
        tk.Label(root, text="WELCOME TO OUR PROJECT",
                 fg="black",
                 bg="salmon1",
                 font="Helvetica 14 bold ").pack()
        tk.Label(root,
                 text='', bg="black", fg="black"
                 ).pack()
        tk.Label(root,
                 text="TITLE:CALENDAR AUTOMATION",
                 fg="yellow",
                 bg="dark green",
                 font="Times 16 bold underline ").pack()
        tk.Label(root,
                 text='', bg="black", fg="black"
                 ).pack()
        tk.Label(root,
                 text="TEAM COMPOSITION",
                 fg="black",
                 bg="tomato2",
                 font="Helvetica 16 bold italic").pack()
        tk.Label(root,
                 text='', bg="black", fg="black"
                 ).pack()
        tk.Label(root,
                 text="1. GOVIND BANSAL",
                 fg="black",
                 bg="indian red",
                 font="Verdana 15 bold").pack()
        tk.Label(root,
                 text='', bg="black", fg="black"
                 ).pack()
        tk.Label(root,
                 text="2. AMRUTHA VARSHINI",
                 fg="black",
                 bg="indian red",
                 font="Verdana 15 bold").pack()
        tk.Label(root,
                 text='', bg="black", fg="black"
                 ).pack()
        tk.Label(root,
                 text="3. MANASA V BHAT",
                 fg="black",
                 bg="indian red",
                 font="Verdana 15 bold").pack()
        tk.Label(root,
                 text='', bg="black", fg="black"
                 ).pack()
        tk.Label(root,
                 text="4. THIRUVIKRAMAN B",
                 fg="black",
                 bg="indian red",
                 font="Verdana 15 bold").pack()
        tk.Label(root,
                 text='', bg="black", fg="black"
                 ).pack()
        tk.Label(root,
                 text="5. ARCHANA ARUN",
                 fg="black",
                 bg="indian red",
                 font="Verdana 15 bold").pack()
        tk.Label(root,
                 text='', bg="black", fg="black"
                 ).pack()
        tk.Label(root,
                 text="MENTOR:PRITHVI SHEKHAR PAGALA",
                 fg="black",
                 bg="indian red",
                 font="Verdana 15 bold").pack()
        tk.Label(root,
                 text='', bg="black", fg="black"
                 ).pack()


def filepath(root):
    """function for submit buttons"""
    path = open_file()
    btn = tk.Button(root, text='Submit', command=lambda: path)
    btn.pack(side=tk.TOP, pady=8)
    button = tk.Button(root, text='Stop', width=25, command=root.destroy)
    button.pack()
    return path


def open_file():
    """function for getting file path"""
    file = filedialog.askopenfilename(title="select excel file",
                                      filetypes=(("Excel file", "*.xlsx"),
                                                 ("All files", "*.*")))
    return file

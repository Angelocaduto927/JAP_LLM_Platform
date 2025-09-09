import os
import sys
import tkinter as tk
from tkinter import messagebox
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from Paper_Generator import test_paper_generation
import mysql.connector


#--------mysql connection--------
db = mysql.connector.connect(
    host="localhost",
    user = 'root',
    password = '123',
    database = 'japgpt'
)
cursor = db.cursor()

#-------build window--------
window = tk.Tk()
window.title("JAP-GPT paper generation")
window.geometry("600x200")

folder_path = tk.StringVar()
folder_path.set("")

#-------build main menu-----------
def setting():
    global folder_path
    subwindow = tk.Tk()
    subwindow.title("setting")
    subwindow.geometry("500x200")
    placeholder = "Please enter file saving path"

    def on_entry_click(event):
        if save_path_entry.get() == placeholder:
            save_path_entry.delete(0,'end')
            save_path_entry.config(fg="black")

    def on_focusout(event):
        if save_path_entry.get() == '':
            save_path_entry.insert(0, placeholder)
            save_path_entry.config(fg = 'grey')

    save_path_entry = tk.Entry(subwindow ,width = 30, fg = "black", font = ('Arial', 10))
    save_path_entry.place(x=120,y=50,anchor='nw')

    label = tk.Label(subwindow, text = "saving path:", fg='black', font=('Arial', 10))
    label.place(x=20,y=47,anchor='nw')

    if folder_path.get() == "":
        save_path_entry.insert(0, placeholder)
        save_path_entry.config(fg = 'grey')
        save_path_entry.bind("<FocusIn>", on_entry_click)
        save_path_entry.bind("<FocusOut>", on_focusout)
    else:
        save_path_entry.insert(0, folder_path.get())

    def save():
        global folder_path
        enter = save_path_entry.get()
        if enter == placeholder:
            messagebox.showerror(title = "Error", message = "Please enter info")
        else:
            folder_path.set(enter)
            subwindow.destroy()

    botton_save = tk.Button(subwindow, text = 'Save', width = 5, height = 1, bg = 'grey', fg = "black", font = ('Arial', 10),command = save)
    botton_cancel = tk.Button(subwindow, text = 'Cancel', width = 5, height = 1, bg = 'grey', fg = "black", font = ('Arial', 10), command = subwindow.destroy)
    botton_save.place(x=150,y=100, anchor='nw')
    botton_cancel.place(x=250,y=100, anchor='nw')

menu = tk.Menu(window, tearoff=0, bg = 'grey', fg = 'black', font = ('Arial', 10))
window.config(menu = menu)
menu.add_command(label = 'setting', command = setting)


#-------build query entry---------
placeholder = "Please enter student's name or SID"

def on_entry_click(event):
    if query_entry.get() == placeholder:
        query_entry.delete(0,'end')
        query_entry.config(fg="black")

def on_focusout(event):
    if query_entry.get() == '':
        query_entry.insert(0, placeholder)
        query_entry.config(fg = 'grey')

query_entry = tk.Entry(window, width = 30, fg = "grey", font = ('Arial', 10))
query_entry.insert(0, placeholder)
query_entry.bind("<FocusIn>", on_entry_click)
query_entry.bind("<FocusOut>", on_focusout)
query_entry.place(x=150,y=50,anchor='nw')

#-------build entry button--------
def update_optionmenu():
    menu = optionmenu['menu']
    menu.delete(0, 'end')
    paper_list.insert(0,"---Please select a paper---")
    for item in paper_list:
        menu.add_command(label=item, command=lambda value=item: test_paper.set(value))

paper_list = []
def query():
    global paper_list, id
    enter = query_entry.get()
    if enter == placeholder:
        messagebox.showerror(title = "Error", message = "Please enter info")
    else:
        id, papers = test_paper_generation.query_id(enter, cursor)
        paper_list = list(papers)
        update_optionmenu()

botton = tk.Button(window, text = "query", bg = 'grey', width = 5, height= 1, command = query)
botton.place(x=370,y=50, anchor='nw')

#--------build paper entry--------
#paper_entry = tk.Entry(window, width = 30, fg = "grey", state = 'readonly', font = ('Arial', 10))
#paper_entry.place()
#paper_entry.insert(0, "Please select a paper")

#-------build option menu--------
test_paper = tk.StringVar(window)
test_paper.set("Please select a paper")
#paper_old.set("Please select a paper")
'''
def show_selection(selected):
    global paper_old
    if selected != "Please select a paper":
        paper_entry.delete(0, 'end')
        paper_entry.insert(0, selected)
        paper_old.set(selected)
        paper_entry.config(fg = 'black')
    elif selected == 'Please select a paper':
        paper_entry.delete(0, 'end')
        paper_entry.insert(0, selected)
        paper_entry.config(fg = 'grey')
'''
paper_list.insert(0, "Please select a paper")
optionmenu = tk.OptionMenu(window, test_paper, *paper_list)#,command = show_selection)
optionmenu.place(x=150,y=100, anchor='nw')



#--------build analysis and generate button--------
def Analysis():
    global id
    if test_paper.get() == "Please select a paper":
        messagebox.showerror(title = "Error", message = "Please select a paper")
    elif folder_path.get() == "":
        messagebox.showerror(title = "Error", message = "Please go to setting and enter folder path")
    else:
        last_file = test_paper_generation.get_latest_paper_id(test_paper.get(), folder_path.get())
        if last_file:
            name = os.path.splitext(os.path.basename(last_file))[0]
            version = name.split("\\")[-1].split(" ")[-1]
            save_path = test_paper.get() + ' ' + str(int(version) + 1)
        else:
            save_path = test_paper.get() +' 1'
        test_paper_generation.general_analysis(id, test_paper.get(), 'ANALYZE', save_path, cursor, folder_path.get())

def Generate():
    global id
    if test_paper.get() == "Please select a paper":
        messagebox.showerror(title = "Error", message = "Please select a paper")
    elif folder_path.get() == "":
        messagebox.showerror(title = "Error", message = "Please go to setting and enter folder path")
    else:
        last_file = test_paper_generation.get_latest_paper_id(test_paper.get(), folder_path.get())
        if last_file:
            name = os.path.splitext(os.path.basename(last_file))[0]
            version = name.split("\\")[-1].split(" ")[-1]
            save_path = test_paper.get() + ' ' + str(int(version) + 1)
        else:
            save_path = test_paper.get() +' 1'
        test_paper_generation.general_analysis(id, test_paper.get(), 'GENERATE', save_path, cursor, folder_path.get())

botton_analysis = tk.Button(window, text = "Analyze", bg = 'grey', fg = 'black', width = 7, height= 1, command = Analysis)
botton_analysis.place(x=150,y=150, anchor='nw')

botton_generate = tk.Button(window, text = "Generate", bg = 'grey', fg = 'black', width = 7, height= 1, command = Generate)
botton_generate.place(x=250,y=150, anchor='nw')

def on_closing():
    try:
        db.close()
    except Exception as e:
        print("Error closing db:", e)
    window.destroy()

# 在进入 mainloop 之前设置好协议回调
window.protocol("WM_DELETE_WINDOW", on_closing)

window.mainloop()


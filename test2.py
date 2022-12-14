from tkinter import *
from tkinter import ttk
from tkinter import messagebox
win = Tk()
win.title("Raspberry Pi UI")
win.geometry('200x100+200+200')
def ok():
    print(radVar.get())
    str = 'nothing selected'
    if radVar.get() == 1:
            str = "Radio 1 selected"
    if radVar.get() == 2:
            str = "Radio 2 selected"
    messagebox.showinfo("Button Clickec", str)
radVar = IntVar()
r1=ttk.Radiobutton(win, text="Radio 1", variable=radVar,  value=1)
r1.grid(column=0, row=0)
r2=ttk.Radiobutton(win, text="Radio 2", variable=radVar, value=2)
r2.grid(column=0, row=1)
action = ttk.Button(win, text = "Click Me", command = ok)
action.grid(column=0, row=2)
win.mainloop()

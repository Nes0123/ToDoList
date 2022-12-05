from tkinter import *


fen = Tk()
fen.geometry('500x500')

a=10

def calcul(a):
    b=a*10
    return b

def DoCalcul():
    return calcul(a)

text1=Label(fen, text=DoCalcul(), width=80) 
text1.pack()



fen.mainloop()



fen.mainloop()
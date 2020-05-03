from tkinter import *
from result1 import results
from tkinter.messagebox import showinfo
def view_command():
    res = results(sem_text.get(),sec_text.get(),join_text.get())
    if (res==True):
        showinfo("Window","Done!")
    else:
        showinfo("Window","Error!")
def show():
    showinfo("Window","Error!")
window=Tk()

window.wm_title("Semester Results")

l1=Label(window,text="Semester",height=2,width=20)
l1.grid(row=0,column=0)

l2=Label(window,text="Section")
l2.grid(row=1,column=0)

sem_text=StringVar()
e1=Entry(window,textvariable=sem_text)
e1.grid(row=0,column=1)

sec_text=StringVar()
e2=Entry(window,textvariable=sec_text)
e2.grid(row=1,column=1)

l5=Label(window,text="Roll Number",height=2,width=20)
l5.grid(row=2,column=0)

join_text=StringVar()
e3=Entry(window,textvariable=join_text)
e3.grid(row=2,column=1)

b1=Button(window,text="Get Result", width=12,command=view_command)
b1.grid(row=3,column=1)

window.geometry("350x150")
window.mainloop()

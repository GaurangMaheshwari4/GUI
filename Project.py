from tkinter import *
from PIL import Image, ImageTk

root = Tk()


root.geometry("500x400")

root.minsize(500,400)

root.maxsize(700,600)

f1 = Frame(root,borderwidth=8,bg="white",relief=SUNKEN)
f1.pack(side=LEFT,fill=Y)
l1 = Label(f1,text="Trying frames",bg="white",font=("Times New Roman",19,"bold"))
l1.pack(pady=10)



f2 = Frame(root,bg="green")
f2.pack(side=TOP,fill=X)


Label1 = Label(f2,text="GUI Project",bg="blue" ,fg="white" ,padx= 40, pady= 40,font=("Times New Roman",19,"bold"),borderwidth=5,relief=SUNKEN)
Label1.pack()

image = Image.open("logo.png")
photo = ImageTk.PhotoImage(image)
l2 = Label(root,image=photo)
l2.pack()

root1 = Canvas(f2,height=200,width=2000)
root1.create_line(0,0,100,100,fill="red")
root1.pack()

root.mainloop()

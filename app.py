from tkinter import *
from tkinter.font import BOLD
from PIL import Image, ImageTk
from tkinter import filedialog

def deselect(event):
    root.focus_set()

def browsetxt():
    filename = filedialog.askopenfilename(initialdir = "/",title = "Select a File",filetypes = [("Text files","*.txt*")])
                                                    
def browseexcel():
    filename = filedialog.askopenfilename(initialdir = "/",title = "Select a File",filetypes = [("Excel","*.xls*;*.xlsx*;*.csv*")])                                                    

def Page1():
    f2.pack_forget()
    f3.pack_forget()
    f1.pack()

def Page2():
    f1.pack_forget()
    f3.pack_forget()
    f2.pack()

def Page3():
    f2.pack_forget()
    f1.pack_forget()
    f3.pack()

def getinfo():
    print(Email.get(),Password.get(),Subject.get())
    Email.delete(0,"end")
    Password.delete(0,"end")
    Subject.delete(0,"end")

root = Tk()

f1 = Frame(root)
f1.pack()

f2 = Frame(root)

f3 = Frame(root)

image = Image.open("APP/Background.jpg")
image1 = image.resize((800,600))
photo = ImageTk.PhotoImage(image1)

left_arrow = Image.open("APP/Left.png")
left_arrow1 = left_arrow.resize((30,30))
left1 = ImageTk.PhotoImage(left_arrow1)
Right_arrow = Image.open("APP/Right.png")
Right_arrow1 = Right_arrow.resize((30,30))
Right1 = ImageTk.PhotoImage(Right_arrow1)
left=Button(root,image=left1,text="Previous",height=50,width=50,compound=TOP,state="disabled")
Right=Button(root,image=Right1,text="Next",height=50,width=50,compound=TOP,command = Page2)

my_canvas = Canvas(f1, width=800, height=600)
my_canvas.pack(fill="both",expand=True)

my_canvas.create_image(0,0, image=photo, anchor="nw")
my_canvas.create_text(400,40,text="Instructions",fill="white",font=("Helvetica",23,BOLD))
my_canvas.create_text(400,130,text="1. Create 2 files : \n\t1) txt file containing the body of the mail  \n\t2) excel sheet containing the values of the placeholder",fill="white",font=("Helvetica",15,BOLD))
my_canvas.create_text(330,210,text="2. Upload these files as asked in this application",fill="white",font=("Helvetica",15,BOLD))
my_canvas.create_text(260,290,text="3. Enter your email and password",fill="white",font=("Helvetica",15,BOLD))
my_canvas.create_text(370,370,text="4. Click on the submit button and your mails will be sent",fill="white",font=("Helvetica",15,BOLD))

left_window = my_canvas.create_window(40,300,window=left)
Right_window = my_canvas.create_window(760,300,window=Right)

#Page 2
my_canvas2 = Canvas(f2, width=800, height=600)
my_canvas2.pack(fill="both",expand=True)
my_canvas2.create_image(0,0, image=photo, anchor="nw")
my_canvas2.create_text(400,40,text="Upload Files",fill="white",font=("Helvetica",23,BOLD))

txt_image = Image.open("APP/Txt.png")
txt_image1 = txt_image.resize((30,30))
txt1 = ImageTk.PhotoImage(txt_image1)

excel_image = Image.open("APP/excel.png")
excel_image1 = excel_image.resize((40,40))
excel1 = ImageTk.PhotoImage(excel_image1)

left2=Button(root,image=left1,text="Previous",height=50,width=50,compound=TOP,command=Page1)
Right2=Button(root,image=Right1,text="Next",height=50,width=50,compound=TOP,command=Page3)
left_window = my_canvas2.create_window(40,300,window=left2)
Right_window = my_canvas2.create_window(760,300,window=Right2)
txt_explore = Button(root,image=txt1,text="Browse Txt file",font=BOLD,pady=10,padx=20,height=80,width=120,compound=BOTTOM,command=browsetxt)
txt_window = my_canvas2.create_window(250,300,window=txt_explore)
excel_explore = Button(root,image=excel1,text="Browse Excel Sheet",font=BOLD,pady=5,padx=20,height=80,width=140,compound=BOTTOM,command=browseexcel)
excel_window = my_canvas2.create_window(550,300,window=excel_explore)

#Page 3
my_canvas3 = Canvas(f3, width=800, height=600)
my_canvas3.pack(fill="both",expand=True)
my_canvas3.create_image(0,0, image=photo, anchor="nw")
my_canvas3.create_text(400,40,text="Information",fill="white",font=("Helvetica",23,BOLD))

left3=Button(root,image=left1,text="Previous",height=50,width=50,compound=TOP,command=Page2)
Right3=Button(root,image=Right1,text="Next",height=50,width=50,compound=TOP)
left_window = my_canvas3.create_window(40,300,window=left3)
Right_window = my_canvas3.create_window(760,300,window=Right3)

Email = Entry(my_canvas3,justify=CENTER,font=("Helvetica",11,BOLD),highlightcolor="red",highlightbackground="green",highlightthickness=2)
Password = Entry(my_canvas3,show="*",justify=CENTER,font=("Helvetica",11,BOLD),highlightcolor="red",highlightbackground="green",highlightthickness=2)
Subject = Entry(my_canvas3,justify=CENTER,font=("Helvetica",11,BOLD),highlightcolor="red",highlightbackground="green",highlightthickness=2)

Email.bind("<Leave>",deselect)
Password.bind("<Leave>",deselect)
Subject.bind("<Leave>",deselect)

my_canvas3.create_window(550,200,window=Email,height=34,width=200)
my_canvas3.create_window(550,300,window=Password,height=34,width=200)
my_canvas3.create_window(550,400,window=Subject,height=34,width=200)
my_canvas3.create_text(330,200,text="Email:",fill="white",font=("Helvetica",15,BOLD))
my_canvas3.create_text(330,300,text="Password:",fill="white",font=("Helvetica",15,BOLD))
my_canvas3.create_text(330,400,text="Subject:",fill="white",font=("Helvetica",15,BOLD))
Submit=Button(root,text="Submit",width=10,pady=10,font=BOLD,command=getinfo)
my_canvas3.create_window(400,500,window=Submit)


root.mainloop()

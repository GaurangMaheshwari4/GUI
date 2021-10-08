from tkinter import *
from PIL import Image, ImageTk
from tkinter.font import BOLD
from tkinter import filedialog
from tkinter import ttk
import re
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
#import win32gui, win32con

#hide = win32gui.GetForegroundWindow()
#win32gui.ShowWindow(hide , win32con.SW_HIDE)

index = 0
pholder_head = dict()
Email = str()

class GUI(Tk):
    def __init__(self):
        super().__init__()
        
        #creating a container
        self.container = Frame(self)
        self.container.pack()

        #bg_image
        image = Image.open("APP/Background.jpg")
        image1 = image.resize((800,600))
        self.photo = ImageTk.PhotoImage(image1)
        
        Nirma_logo = Image.open("APP/Nirma_logo.png")
        Nirma_logo1 = Nirma_logo.resize((150,100))
        self.Nirma_logo2 = ImageTk.PhotoImage(Nirma_logo1)

        #Left Button
        left_arrow = Image.open("APP/Left.png")
        left_arrow1 = left_arrow.resize((30,30))
        self.left1 = ImageTk.PhotoImage(left_arrow1)

        #Right Button
        Right_arrow = Image.open("APP/Right.png")
        Right_arrow1 = Right_arrow.resize((30,30))
        self.Right1 = ImageTk.PhotoImage(Right_arrow1)

        #Text Image
        txt_image = Image.open("APP/Txt.png")
        txt_image1 = txt_image.resize((30,30))
        self.txt1 = ImageTk.PhotoImage(txt_image1) 

        #Excel Image
        excel_image = Image.open("APP/excel.png")
        excel_image1 = excel_image.resize((40,40))
        self.excel1 = ImageTk.PhotoImage(excel_image1)  

        #array of frames
        self.frames = [FirstPage(self.container,self),SecondPage(self.container,self)]
        
        
        
        self.show_frame(0)

    def destroy_frame(self):
        self.frames = self.frames[:2]

    def show_frame(self,cont):
        global index 
        index = cont
        for i in range(len(self.frames)):
            self.frames[i].pack_forget()
        self.frames[cont].pack()
        

    def dynamic(self):
        if(index+1 == len(self.frames)):
            if(index==1):
                self.frames.append(ThirdPage(self.container,self))
            else:    
                self.frames.append(FourthPage(self.container,self))

    def dynamic1(self):
        if(index+1 == len(self.frames)):
            self.frames.append(FifthPage(self.container,self))           
             


class FirstPage(Frame):
    def __init__(self, parent, controller):
        Frame.__init__(self,parent)
        left=Button(controller,image=controller.left1,text="Previous",height=50,width=50,compound=TOP,state="disabled")
        Right=Button(controller,image=controller.Right1,text="Next",height=50,width=50,compound=TOP,command = lambda: controller.show_frame(1))
        my_canvas = Canvas(self, width=800, height=600)
        my_canvas.pack(fill="both",expand=True)

        my_canvas.create_image(0,0, image= controller.photo, anchor="nw")
        my_canvas.create_image(770,20, image= controller.Nirma_logo2, anchor="ne")
        my_canvas.create_text(400,40,text="Instructions",fill="white",font=("Helvetica",23,BOLD))
        my_canvas.create_text(400,130,text="1. Create 2 files : \n\t1) txt file containing the body of the mail  \n\t2) excel sheet containing the values of the placeholder",fill="white",font=("Helvetica",15,BOLD))
        my_canvas.create_text(330,210,text="2. Upload these files as asked in this application",fill="white",font=("Helvetica",15,BOLD))
        my_canvas.create_text(260,290,text="3. Enter your email and password",fill="white",font=("Helvetica",15,BOLD))
        my_canvas.create_text(370,370,text="4. Click on the submit button and your mails will be sent",fill="white",font=("Helvetica",15,BOLD))
        my_canvas.create_text(150,500,text="Made by:\n\nHarsh Ahuja\n\nGaurang Maheshwari",fill="yellow",font=("Helvetica",15,BOLD))
        my_canvas.create_text(750,500,text="Mentored by:\n\nDr. Sachin Gajjar\n\nPurvansh Shah",fill="yellow",font=("Helvetica",15,BOLD),anchor="e")

        my_canvas.create_window(40,300,window=left)
        my_canvas.create_window(760,300,window=Right)

        

class SecondPage(Frame):
    txtfile = ""
    excelfile = ""
    def __init__(self, parent, controller):
        Frame.__init__(self,parent)
        my_canvas1 = Canvas(self, width=800, height=600)
        my_canvas1.pack(fill="both",expand=True)
        my_canvas1.create_image(0,0, image=controller.photo, anchor="nw")
        my_canvas1.create_image(770,20, image= controller.Nirma_logo2, anchor="ne")
        my_canvas1.create_text(400,40,text="Upload Files",fill="white",font=("Helvetica",23,BOLD))

        left=Button(controller,image=controller.left1,text="Previous",height=50,width=50,compound=TOP,command=lambda: controller.show_frame(0))
        Right=Button(controller,image=controller.Right1,text="Next",height=50,width=50,compound=TOP,command=lambda: [controller.dynamic(),controller.show_frame(2)])
        my_canvas1.create_window(40,300,window=left)
        my_canvas1.create_window(760,300,window=Right)

        txt_explore = Button(controller,image=controller.txt1,text="Browse Txt file",font=BOLD,pady=10,padx=20,height=80,width=120,compound=BOTTOM,command=self.browsetxt)
        my_canvas1.create_window(250,300,window=txt_explore)
        excel_explore = Button(controller,image=controller.excel1,text="Browse Excel Sheet",font=BOLD,pady=5,padx=20,height=80,width=140,compound=BOTTOM,command=self.browseexcel)
        my_canvas1.create_window(550,300,window=excel_explore)

    def browsetxt(self):
        self.txtfile = filedialog.askopenfilename(initialdir = "/",title = "Select a File",filetypes = [("Text files","*.txt*")])
                                                    
    def browseexcel(self):
        self.excelfile = filedialog.askopenfilename(initialdir = "/",title = "Select a File",filetypes = [("Excel","*.xls*;*.xlsx*")])   
        

class ThirdPage(Frame):
    def __init__(self, parent, controller):
        Frame.__init__(self,parent)
        my_canvas2 = Canvas(self, width=800, height=600)
        my_canvas2.pack(fill="both",expand=True)
        my_canvas2.create_image(0,0, image=controller.photo, anchor="nw")
        my_canvas2.create_image(770,20, image= controller.Nirma_logo2, anchor="ne")
        my_canvas2.create_text(400,40,text="Information",fill="white",font=("Helvetica",23,BOLD))

        left=Button(controller,image=controller.left1,text="Previous",height=50,width=50,compound=TOP,command=lambda: [controller.show_frame(1),controller.destroy_frame()])
        Right=Button(controller,image=controller.Right1,text="Next",height=50,width=50,compound=TOP,command=lambda: [controller.dynamic(),controller.show_frame(3)])
        my_canvas2.create_window(40,300,window=left)
        my_canvas2.create_window(760,300,window=Right)

        self.Email = Entry(my_canvas2,justify=CENTER,font=("Helvetica",11,BOLD),highlightcolor="red",highlightbackground="green",highlightthickness=2)
        self.Password = Entry(my_canvas2,show="*",justify=CENTER,font=("Helvetica",11,BOLD),highlightcolor="red",highlightbackground="green",highlightthickness=2)
        self.Subject = Entry(my_canvas2,justify=CENTER,font=("Helvetica",11,BOLD),highlightcolor="red",highlightbackground="green",highlightthickness=2)

        self.Email.bind("<Leave>",self.deselect)
        self.Password.bind("<Leave>",self.deselect)
        self.Subject.bind("<Leave>",self.deselect)

        my_canvas2.create_window(550,200,window=self.Email,height=34,width=200)
        my_canvas2.create_window(550,300,window=self.Password,height=34,width=200)
        my_canvas2.create_window(550,400,window=self.Subject,height=34,width=200)
        my_canvas2.create_text(330,200,text="Email:",fill="white",font=("Helvetica",15,BOLD))
        my_canvas2.create_text(330,300,text="Password:",fill="white",font=("Helvetica",15,BOLD))
        my_canvas2.create_text(330,400,text="Subject:",fill="white",font=("Helvetica",15,BOLD))

        Submit=Button(controller,text="Submit",width=10,pady=10,font=BOLD,command=self.getinfo)
        my_canvas2.create_window(400,500,window=Submit)

        file = open(controller.frames[1].txtfile,"r")
        self.body = "".join(file.readlines())
        self.placeholder = list(set(re.findall("\{(.+?)\}",self.body)))
        self.placeholder.sort()
        file.close()
        self.df = pd.read_excel(controller.frames[1].excelfile)
        self.place_hold = self.df.to_dict('list')
        self.head = list(self.df)
    
    def getinfo(self):
        self.Sender = self.Email.get()
        self.Pass = self.Password.get()
        self.Subj = self.Subject.get()
        self.Email.delete(0,"end")
        self.Password.delete(0,"end")
        self.Subject.delete(0,"end")

    def deselect(self,event):
        self.focus_set()

class FourthPage(Frame):
    def __init__(self, parent, controller):
        global pholder_head
        controller.option_add("*TCombobox*Listbox.font",("Helvetica",15,BOLD))
        Frame.__init__(self,parent)
        my_canvas3 = Canvas(self, width=800, height=600)
        my_canvas3.pack(fill="both",expand=True)
        my_canvas3.create_image(0,0, image= controller.photo, anchor="nw")
        my_canvas3.create_image(10,10, image= controller.Nirma_logo2, anchor="nw")
        my_canvas3.create_text(400,40,text= "Place holder",fill="white",font=("Helvetica",23,BOLD))

        left=Button(controller,image=controller.left1,text="Previous",height=50,width=50,compound=TOP,command=lambda: controller.show_frame(index-1))
        if(((len(controller.frames[2].placeholder))//5)>index-2):
            Right=Button(controller,image=controller.Right1,text="Next",height=50,width=50,compound=TOP,command=lambda: [self.getinfo(controller),controller.dynamic(),controller.show_frame(index+1)])
            
        else:
            Right=Button(controller,image=controller.Right1,text="Next",height=50,width=50,compound=TOP,command=lambda: [self.getinfo(controller),controller.dynamic1(),controller.show_frame(index+1)])
        my_canvas3.create_window(40,300,window=left)
        my_canvas3.create_window(760,300,window=Right)
        i = 0
        while(i+(index-2)*5<len(controller.frames[2].placeholder)) and (i<5):
            if(i==0):
                self.pholder_list = [None]*(len(controller.frames[2].placeholder)-(index-2)*5)
            my_canvas3.create_text(250,100+i*90,text=controller.frames[2].placeholder[i+(index-2)*5],fill="white",font=("Helvetica",15,BOLD))
            pholder_head[controller.frames[2].placeholder[i+(index-2)*5]]=""
            self.pholder_list[i] = ttk.Combobox(controller,textvariable=pholder_head[controller.frames[2].placeholder[i+(index-2)*5]],height=30,font=("Helvetica",13,BOLD))
            self.pholder_list[i]['values']= controller.frames[2].head
            self.pholder_list[i].current()
            my_canvas3.create_window(600,100+i*90,window=self.pholder_list[i],height=30)
            i=i+1
        if(((len(controller.frames[2].placeholder))//5)>index-2):
            Lock_nxt=Button(controller,text="LOCK and Next",width=30,pady=10,font=BOLD,command=lambda:[self.getinfo(controller),controller.dynamic(),controller.show_frame(index+1)])
        else:
            Lock_nxt=Button(controller,text="LOCK and Next",width=30,pady=10,font=BOLD,command=lambda:[self.getinfo(controller),controller.dynamic1(),controller.show_frame(index+1)])
        my_canvas3.create_window(400,550,window=Lock_nxt)

    def getinfo(self,controller):
            i=0
            while(i+(index-3)*5<len(controller.frames[2].placeholder)) and (i<5):
                pholder_head[controller.frames[2].placeholder[i+(index-3)*5]] = self.pholder_list[i].get()
                i= i+1
            
    
            

class FifthPage(Frame):
    def __init__(self, parent, controller):
        self.success = 0
        self.unsuccessful = 0
        controller.option_add("*TCombobox*Listbox.font",("Helvetica",15,BOLD))
        Frame.__init__(self,parent)
        self.my_canvas3 = Canvas(self, width=800, height=600)
        self.my_canvas3.pack(fill="both",expand=True)
        self.my_canvas3.create_image(0,0, image= controller.photo, anchor="nw")
        self.my_canvas3.create_image(770,20, image= controller.Nirma_logo2, anchor="ne")
        self.my_canvas3.create_text(400,40,text= "Confirmation",fill="white",font=("Helvetica",23,BOLD))

        left=Button(controller,image=controller.left1,text="Previous",height=50,width=50,compound=TOP,command=lambda: controller.show_frame(index-1))
        Right=Button(controller,image=controller.Right1,text="Next",height=50,width=50,compound=TOP,state="disabled")
        self.my_canvas3.create_window(40,300,window=left)
        self.my_canvas3.create_window(760,300,window=Right)
        self.my_canvas3.create_text(250,200,text="Receiver's email Column",fill="white",font=("Helvetica",15,BOLD))
        self.pholder_list = ttk.Combobox(controller,textvariable=Email,height=30,font=("Helvetica",13,BOLD))
        self.pholder_list['values']= controller.frames[2].head
        self.pholder_list.current()
        self.my_canvas3.create_window(600,200,window=self.pholder_list,height=30)
        Confirm=Button(controller,text="Confirm",width=30,pady=10,font=BOLD,command = lambda: [self.getinfo(),self.sendmail(controller)])
        self.my_canvas3.create_window(400,300,window=Confirm)
            

    def getinfo(self):
        global Email 
        Email = self.pholder_list.get()

    def sendmail(self,controller):
        self.success = 0
        self.unsuccessful = 0
        for a in range(len(controller.frames[2].df)):
            body = controller.frames[2].body
            for x in pholder_head.keys():
                try:
                    body = body.replace("{"+x+"}",str(controller.frames[2].place_hold[pholder_head[x]][a]))
                except:
                    pass
            try:
                email_ver(controller.frames[2].Sender,controller.frames[2].Pass,controller.frames[2].Subj,controller.frames[2].place_hold[Email][a],body)
                self.success = self.success+1
            except:
                self.unsuccessful = self.unsuccessful + 1
        self.my_canvas3.create_text(250,475,text="No. of successful emails sent : " + str(self.success),fill="light green",font=("Helvetica",15,BOLD))
        self.my_canvas3.create_text(250,550,text="No. of unsuccessful emails : " + str(self.unsuccessful),fill="red",font=("Helvetica",15,BOLD))    
                               

def email_ver(From,Password,Subject,To,message):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = Subject
    msg["From"] = From
    msg["To"] = To
    part1 = MIMEText(message,'plain')
    msg.attach(part1)
    server = smtplib.SMTP("smtp.gmail.com",587)
    server.starttls()
    server.login(From,Password)
    server.sendmail(From,To,msg.as_string())
    server.quit()

        
        
if __name__=="__main__":
    root = GUI()
    root.mainloop()
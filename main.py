from tkinter import *
from tkinter import messagebox,filedialog
import os
import sys
import PIL
import pandas as pd
import numpy as np
from PIL import Image, ImageTk
import email_function
import time

class BULK_EMAIL:
    def __init__(self,root):
        self.root=root
        self.root.title("BULK EMAIL APPLICATION")
        self.root.geometry("700x400+100+50")
        # self.root.resizable(False,False)
        self.root.config(bg="white")
        #-------------ICONS-----------------------
        self.email_icon=ImageTk.PhotoImage(file="images/email.png")
        self.setting_icon=ImageTk.PhotoImage(file="images/settings.png")
    
        #-------------title-----------------------
        title=Label(self.root,text="BULK EMAIL SENDER",image=self.email_icon,padx=10,compound=LEFT,font=("Goudy Old Style",48,"bold"),bg="#222A35",fg="white",anchor="w").place(x=0,y=0,relwidth=1)
        desc=Label(self.root,text="Use Excel file to send the Bulk Email, with one click, Make sure the Email column must be Email",font=("Calibri (Body)",14),bg="#FFD966",fg="#262626").place(x=0,y=110,relwidth=1)

        btn_setting=Button(self.root,image=self.setting_icon,compound=RIGHT,command=self.setting_window).place(x=1150,y=0)

        self.var_choice=StringVar()
        single=Radiobutton(self.root,text="Single",value="single",variable=self.var_choice,activebackground="blue",font=("times new roman",30,"bold"),bg="white",fg="#262626",command=self.check_single_or_bulk).place(x=50,y=150)
        bulk=Radiobutton(self.root,text="Bulk",value="bulk",variable=self.var_choice,activebackground="blue",font=("times new roman",30,"bold"),bg="white",fg="#262626",command=self.check_single_or_bulk).place(x=250,y=150)

        self.var_choice.set("single")


        to=Label(self.root,text="TO (Email Address)",font=("times new roman",18,"bold"),bg="white").place(x=50,y=250)
        sub=Label(self.root,text="SUBJECT",font=("times new roman",18,"bold"),bg="white").place(x=50,y=300)
        msg=Label(self.root,text="MESSAGE",font=("times new roman",18,"bold"),bg="white").place(x=50,y=350)

        self.btn_browse=Button(self.root,command=self.browse_file,text="BROWSE",font=("times new roman",18,"bold"),bg="#00B0F0",fg="white",cursor="hand2",state=DISABLED)
        self.btn_browse.place(x=670,y=250,width=120,height=30)

        self.txt_to=Entry(self.root,font=("times new roman",16),bg="lightblue",borderwidth=3)
        self.txt_to.place(x=300,y=250,width=350,height=30)
 
        self.txt_sub=Entry(self.root,font=("times new roman",16),bg="lightblue",borderwidth=3)
        self.txt_sub.place(x=300,y=300,width=700,height=30)

        self.txt_msg=Text(self.root,font=("times new roman",16),bg="lightblue",borderwidth=3)
        self.txt_msg.place(x=300,y=350,width=700,height=140)
    #===============================================================================================

            #-------------------STATUS--------------------------------------
        
        self.lbl_total=Label(self.root,font=("times new roman",18),bg="white",fg="grey")
        self.lbl_total.place(x=50,y=510)

        self.lbl_sent=Label(self.root,font=("times new roman",18),bg="white",fg="green")
        self.lbl_sent.place(x=280,y=510)

        self.lbl_left=Label(self.root,font=("times new roman",18),bg="white",fg="orange")
        self.lbl_left.place(x=420,y=510)

        self.lbl_failed=Label(self.root,font=("times new roman",18),bg="white",fg="red")
        self.lbl_failed.place(x=550,y=510)

        btn_clear=Button(self.root,text="CLEAR",command=self.clear1,font=("times new roman",18,"bold"),bg="#00B0F0",fg="#000000",cursor="hand2").place(x=730,y=510,width=120)
        btn_send=Button(self.root,text="SEND",command=self.send_email,font=("times new roman",18,"bold"),bg="#00B0F0",fg="#000000",cursor="hand2").place(x=880,y=510,width=120)
        self.check_file_exist()



    def setting_window(self):
        self.root2=Toplevel()
        self.root2.title("setting")
        self.root2.geometry("700x450+350+90")
        self.root2.focus_force()
        self.root2.grab_set()
        
        title2=Label(self.root2,text="LOGIN HERE",image=self.setting_icon,compound=LEFT,font=("Goudy Old Style",48,"bold"),bg="#222A35",fg="white",anchor="w").place(x=0,y=0,relwidth=1)

        desc2=Label(self.root2,text="Enter Email Address and Password from which you want to send the emails",font=("Calibri (Body)",14),bg="#FFD966",fg="#262626").place(x=0,y=110,relwidth=1)

        email_=Label(self.root,text="Email Address", font=("times new roman",18)).place(x=50,y=220,height=30)
        pw_=Label(self.root,text="Password", font=("times new roman",18)).place(x=50,y=300,height=30)

        self.txt_email=Entry(self.root2,font=("times new roman",18),bg="lightblue")
        self.txt_email.place(x=300,y=220,width=350,height=40)
        
        self.txt_pw=Entry(self.root2,font=("times new roman",18),bg="lightblue")
        self.txt_pw.place(x=300,y=310,width=350,height=40)

        btn_save=Button(self.root2,command=self.save_setting,text="SAVE",font=("times new roman",18,"bold"),bg="#00B0F0",fg="#000000",cursor="hand2").place(x=500,y=400,width=120)
        btn_clear2=Button(self.root2,command=self.clear2,text="CLEAR",font=("times new roman",18,"bold"),bg="#00B0F0",fg="#000000",cursor="hand2").place(x=350,y=400,width=120)
        
        self.txt_email.insert(0,self.email_)
        self.txt_pw.insert(0,self.pw_)

    def clear2(self):
        self.txt_email.delete(0,END)
        self.txt_pw.delete(0,END)

    def send_email(self):
        x=len(self.txt_msg.get('1.0',END))
        if self.txt_to.get()=="" or self.txt_sub.get()=="" or x==1:
            messagebox.showerror('Error','All fields are required',parent=self.root)
        else:
            if self.var_choice.get()=="single":
                status=email_function.email_send_func(self.txt_to.get(),self.txt_sub.get(),self.txt_msg.get('1.0',END),self.email_,self.pw_)
                if status=='s':
                    messagebox.showinfo("Success","EMAIL HAS BEEN SENT IT SUCCESFULLY",parent=self.root)
                if status=='f':
                    messagebox.showerror("Success","EMAIL Failed, Try Again Later",parent=self.root)

            if self.var_choice.get()=="bulk":
                self.failed=[]
                self.s_count=0
                self.f_count=0
                for x in self.emails:
                    status=email_function.email_send_func(x,self.txt_sub.get(),self.txt_msg.get("1.0",END),self.email_,self.pw_)
                    if status=='s':
                        self.s_count+=1
                    if status=='f':
                        self.f_count+=1
                    self.status_bar()
                    time.sleep(1)
                messagebox.showinfo("SUCCESS", "Email has been sent, Check Status",parent=self.root)
    
    
    def clear1(self):
        self.txt_to.config(state=NORMAL)
        self.txt_to.delete(0,END)
        self.txt_sub.delete(0,END)
        self.txt_msg.delete('1.0',END)
        self.var_choice.set("single")
        self.btn_browse.config(state=DISABLED)
        self.lbl_total.config(text="")
        self.lbl_left.config(text="")
        self.lbl_sent.config(text="")
        self.lbl_failed.config(text="")


    def browse_file(self):
        op=filedialog.askopenfile(initialdir='/',title="Select Excel file for Emails",filetypes=(("All Files","*.*"),("Excel Files",".xlsx")))
        if op!=None:
            data=pd.read_excel(op.name)
            if 'email' in data.columns:
                self.emails=list(data['email'])
                c=[]
                for i in self.emails:
                    if pd.isnull(i)==False:
                        c.append(i)
                self.emails=np.array(c)
                print(self.emails)
                print(type(self.emails))
                if len(self.emails)>0:
                    self.txt_to.config(state=NORMAL)
                    self.txt_to.delete(0,END)
                    self.txt_to.insert(0,str(op.name.split("/")[-1]))
                    self.txt_to.config(state='readonly')
                    self.lbl_total.config(text="TOTAL: "+str(len(self.emails)))
                    self.lbl_sent.config(text="SENT: ")
                    self.lbl_left.config(text="LEFT: ")
                    self.lbl_failed.config(text="FAILED: ")
                else:
                    messagebox.showerror("Error","This file does not have any email column",parent=self.root)
            else:
                messagebox.showerror("Error","Please select the file which have email column",parent=self.root)
    
    def check_single_or_bulk(self):
        if self.var_choice.get()=="single":
            self.btn_browse.config(state=DISABLED)
            self.txt_to.config(state=NORMAL)
            self.txt_to.delete(0,END)
            self.clear1()
        if self.var_choice.get()=="bulk":
            self.btn_browse.config(state=NORMAL)
            self.txt_to.delete(0,END)
            self.txt_to.config(state='readonly')

    def save_setting(self):
        if self.txt_email.get()=="" or self.txt_pw.get()=="":
            messagebox.showerror("Error","All fields are required",parent=self.root2)
        else:
            f=open('important.txt','w')
            f.write(self.txt_email.get()+","+self.txt_pw.get())
            f.close()
            messagebox.showinfo("Success","your email password saved successfully")
            self.check_file_exist()
        

    def status_bar(self):
        self.lbl_total.config(text="Status "+str(len(self.emails))+"->>")
        self.lbl_sent.config(text="Sent "+str(self.s_count))
        self.lbl_left.config(text="Left "+str(len(self.emails)-(self.s_count+self.f_count)))
        self.lbl_failed.config(text="Failed "+str(self.f_count))
        self.lbl_total.update()
        self.lbl_sent.update()
        self.lbl_left.update()
        self.lbl_failed.update()


    def check_file_exist(self):
        if os.path.exists("important.txt")==False:
            f=open("important.txt",'w')
            f.write(",")
            f.close()
        f2=open("important.txt",'r')
        self.credentials=[]
        for i in f2:
            self.credentials.append( [i.split(",")[0],i.split(",")[1]] )
        self.email_=self.credentials[0][0]
        self.pw_=self.credentials[0][1]

root=Tk()
obj=BULK_EMAIL(root)
root.mainloop()
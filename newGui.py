import time,datetime,sqlite3,tkinter.filedialog,tkinter.messagebox,pymysql,importlib,xlwt,openpyxl,xlsxwriter
from tkinter import *
from PIL import Image,ImageTk
from newRead import textCard,photoCard,resizeImg
import newRead 
from xlwt import Workbook  
import tkinter as tk
import mysql.connector
from tkinter import ttk
from tkinter.font import Font
#global variabel ค่าเริ่มต้นสำหรับติดต่อฐานข้อมูล
global  userDB
global  passDB
global  hostDB
global  nameDB

userDB = "root"
passDB = ""
hostDB = "localhost"
nameDB = "smart_card"

def Main():
    root = tk.Tk()
    root.geometry('960x650')
    root.title("ข้อมูลบัตรประชาชน")
    def abt(): #เป็นฟังก์ชันกล่อง pop up แสดงข้อความต่อเมื่อ ฟังก์ชัน subm2.add_command ถูกกด
        tkinter.messagebox.showinfo("เกี่ยวกับ","smart red v1. เป็นโปรแกรมสำหรับอ่านบัตรประชาชน พัฒนาโดยคนไทย")
    def ext_1():  #เป็นฟังก์ชันคำสั่งออกจากโปรแกรม  ซึ่องรอรับคำสั่งจาก subm1.add_command และ but_quit
        #root.quit()
        root.destroy()
        sys.exit()
    # คำสั่ง menu bar
    menu = Menu(root)
    root.config(menu=menu)

    subm1 = Menu(menu)
    menu.add_cascade(label="File",menu=subm1)
    subm1.add_command(label="Exit",command=ext_1)

    subm2 = Menu(menu)
    menu.add_cascade(label="Option",menu=subm2)
    subm2.add_command(label="เกี่ยวกับ",command=abt)

    

    nb = ttk.Notebook(root)
    mygreen = "#d2ffd2"
    myblue = "pale turquoise"
    style = ttk.Style()
    style.theme_create( "preechai", parent="alt", settings={
            "TNotebook": {"configure": {"tabmargins": [0, 0, 0, 0] } },
            "TNotebook.Tab": {
                "configure": {"padding": [5, 1], "background": mygreen },
                "map":       {"background": [("selected", myblue)],
                            "expand": [("selected", [1, 1, 1, 0])] } } } )
    style.theme_use("preechai")

    page1 = ttk.Frame(nb)
    layout1(page1)
    page2 = ttk.Frame(nb)
    layout2(page2)
    page3 = ttk.Frame(nb)
    layout3(page3)
    
    nb.add(page3, text='คำแนะคำ')
    nb.add(page1, text='ข้อมูลแบบอ่านบัตร')
    nb.add(page2, text='ข้อมูลแบบกรอก')
    


    nb.pack(fill=BOTH, expand=1)
    root.mainloop()
def layout1(page):
    def clearTextBox():
        #เป็นส่วนของการ clear text box 
        entry_0.delete(first=0,last=100)
        entry_1.delete(first=0,last=100)
        entry_2.delete(first=0,last=100)
        entry_3.delete(first=0,last=100)
        entry_4.delete(first=0,last=100)
    def ext():#เป็นฟังก์ชันคำสั่งออกจากโปรแกรม  ซึ่องรอรับคำสั่งจาก subm1.add_command และ but_quit
        #root.quit()
        page.destroy()
        sys.exit()
    def refreshData():
        try:
            fileName = a = textCard()
            photoCard(fileName)#รับอากิวเมน fileName ที่เป็น array โดย photoCard()ทำหน้าที่ดึงรูปและบันทึก
            resizeImg(fileName)#resizeImg()ทำหน้าที่ปรับขนาดรูปและบันทึก
            
            path = "temp/"+a[0]+".png"
            img = ImageTk.PhotoImage(Image.open(path))
            panel = tk.Label(page, image = img)
            panel.image = img # keep a reference!
            panel.pack(side = "top", fill = "both", expand = "yes")
            panel.place(x=435,y=25)

            label_10 = Label(page, text=a[0], anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
            label_10.place(x=150,y=240) #cid

            label_11 = Label(page, text=a[1], anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
            label_11.place(x=150,y=270) #ชื่อไทย

            label_12 = Label(page, text=a[2], anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
            label_12.place(x=625,y=270) #ชื่ออิ้ง

            label_20 = Label(page, text=a[13], anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
            label_20.place(x=625,y=300) #วันเกิดอิ้ง

            label_13 = Label(page, text=a[4], anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
            label_13.place(x=150,y=300) #วันเกิดไทย

            label_14 = Label(page, text=a[5], anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
            label_14.place(x=625,y=240) #เพศ

            label_15 = Label(page, text=a[6], anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
            label_15.place(x=150,y=390) #สถานที่ออกบัตร

            label_16 = Label(page, text=a[8], anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
            label_16.place(x=150,y=330) #วันออกบัตร

            label_17 = Label(page, text=a[10], anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
            label_17.place(x=625,y=330) #วันหมดอายุบัตร

            label_18 = Label(page, text=a[11], anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
            label_18.place(x=150,y=360) #ที่อยู่

            label_19 = Label(page, text=a[12], anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
            label_19.place(x=625,y=360) #เวลาอ่านบัตร

            label_21 = Label(page, text=a[14], anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
            label_21.place(x=625,y=390) #อายุ 
            clearTextBox()
            #tkinter.messagebox.showinfo("สถานะของบัตร","xxxxxxxxxx")
        except:
            try: #แสดงสถานะเตือนกณีไม่มีข้อมูล
                path = "user.png"
                img = ImageTk.PhotoImage(Image.open(path))
                panel = tk.Label(page, image = img)
                panel.image = img # keep a reference!
                panel.pack(side = "top", fill = "both", expand = "yes")
                panel.place(x=435,y=25)
                
                label_10 = Label(page, text='ไม่มีข้อมูล', anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
                label_10.place(x=150,y=240) #cid
                label_11 = Label(page, text='ไม่มีข้อมูล', anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
                label_11.place(x=150,y=270) #ชื่อไทย
                label_12 = Label(page, text='ไม่มีข้อมูล', anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
                label_12.place(x=625,y=270) #ชื่ออิ้ง
                label_20 = Label(page, text='ไม่มีข้อมูล', anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
                label_20.place(x=625,y=300) #วันเกิดอิ้ง
                label_13 = Label(page, text='ไม่มีข้อมูล', anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
                label_13.place(x=150,y=300) #วันเกิดไทย
                label_14 = Label(page, text='ไม่มีข้อมูล', anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
                label_14.place(x=625,y=240) #เพศ
                label_15 = Label(page, text='ไม่มีข้อมูล', anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
                label_15.place(x=150,y=390) #สถานที่ออกบัตร
                label_16 = Label(page, text='ไม่มีข้อมูล', anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
                label_16.place(x=150,y=330) #วันออกบัตร
                label_17 = Label(page, text='ไม่มีข้อมูล', anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
                label_17.place(x=625,y=330) #วันหมดอายุบัตร
                label_18 = Label(page, text='ไม่มีข้อมูล', anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
                label_18.place(x=150,y=360) #ที่อยู่
                label_19 = Label(page, text='ไม่มีข้อมูล', anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
                label_19.place(x=625,y=360) #เวลาอ่านบัตร
                label_21 = Label(page, text='ไม่มีข้อมูล', anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
                label_21.place(x=625,y=390) #อายุ
            except:
                pass
            clearTextBox()
            tkinter.messagebox.showinfo("สถานะของบัตร","โปรดรวจสอบบัตรของท่านหรือ เสียบเครื่องอ่านบัตรก่อนเปิดโปรแกรม")
    def saveInputDB():
        try:
            textbox = []
            reli_gion = entry_0.get()
            mail = entry_1.get()
            tel = entry_2.get()
            TagID = entry_3.get()
            depart = entry_4.get()
            userDB
            passDB
            hostDB
            nameDB
            textbox.append(mail)
            textbox.append(tel)
            textbox.append(TagID)
            textbox.append(depart)
            textbox.append(reli_gion)
            textbox.append(userDB)
            textbox.append(passDB)
            textbox.append(hostDB)
            textbox.append(nameDB)
            #ต้องการตรวจสอบข้อมูลในการเชื่อมต่อฐานข้อมูล
            mydb = mysql.connector.connect(user=textbox[5], password=textbox[6], host=textbox[7], database=textbox[8])
            
            if textbox[2] == '':
                tkinter.messagebox.showinfo("แจ้งเตือน","กรุณากรอก TagID")

            else:
                b = textCard()
                mydb = mysql.connector.connect(user=textbox[5], password=textbox[6], host=textbox[7], database=textbox[8])
                mycursor = mydb.cursor()
                sql = "SELECT * FROM user WHERE ID_Card = '"+str(b[0])+"' "
                mycursor.execute(sql)
                myresult = mycursor.fetchall()
                try:
                    print(myresult[0][1])#เพื่อให้เช็คว่ามีข้อมูลใน DB เป็นการดักเพื่อเข้าหรือไม่เข้า ในเงื่อนไข excep
                    print('มีข้อมูลแล้ว')
                    if textbox[0]!='':    
                        mydb = mysql.connector.connect(user=textbox[5], password=textbox[6], host=textbox[7], database=textbox[8])
                        mycursor = mydb.cursor()
                        sql3 = "UPDATE user SET Email='"+str(textbox[0])+"' WHERE ID_Card = "+str(b[0])+ ""
                        mycursor.execute(sql3)
                        mydb.commit() 
                    if textbox[1]!='':    
                        mydb = mysql.connector.connect(user=textbox[5], password=textbox[6], host=textbox[7], database=textbox[8])
                        mycursor = mydb.cursor()
                        sql3 = "UPDATE user SET Telephon='"+str(textbox[1])+"' WHERE ID_Card = "+str(b[0])+ ""
                        mycursor.execute(sql3)
                        mydb.commit()
                    if textbox[3]!='':    
                        mydb = mysql.connector.connect(user=textbox[5], password=textbox[6], host=textbox[7], database=textbox[8])
                        mycursor = mydb.cursor()
                        sql3 = "UPDATE user SET Department='"+str(textbox[3])+"' WHERE ID_Card = "+str(b[0])+ ""
                        mycursor.execute(sql3)
                        mydb.commit() 
                    if textbox[4]!='':    
                        mydb = mysql.connector.connect(user=textbox[5], password=textbox[6], host=textbox[7], database=textbox[8])
                        mycursor = mydb.cursor()
                        sql3 = "UPDATE user SET religion='"+str(textbox[4])+"' WHERE ID_Card = "+str(b[0])+ ""
                        mycursor.execute(sql3)
                        mydb.commit()   
                    mydb = mysql.connector.connect(user=textbox[5], password=textbox[6], host=textbox[7], database=textbox[8])
                    mycursor = mydb.cursor()
                    nameImg = "result/ImagesAll/"+str(b[0])+".png"
                    sql1 = "UPDATE user SET ID_Card="+str(b[0])+ ", Name_TH= '"+b[1]+"', Name_EN= '"+b[2]+"' , Birth= "+str(b[3])+", GEN=  '"+b[5]+"', Card_Issuer= '"+b[6]+"', Issue_Date= "+str(b[7])+", Expire_Date= "+str(b[9])+", Address= '"+b[11]+"', Time_read_card='"+b[12]+"', image= '"+nameImg+"', age= "+str(b[14])+ " WHERE ID_Card = "+str(b[0])+ ""
                    mycursor.execute(sql1)
                    mydb.commit()
                    

                    cnx = pymysql.connect(user=textbox[5], password=textbox[6], host=textbox[7], database=textbox[8])
                    print("1")
                    cur = cnx.cursor()
                    insert3 = "INSERT INTO pairing (ID_Card, Type_ID, TagID, Time_read_card)"
                    #ID_Card Type_ID TagID Type_User Time_read_card
                    value3 =  "VALUES ('"+str(b[0])+ "','ThaiCard', '"+str(textbox[2])+"', '"+b[12]+"')"
                    print(b[12])
                    count2 = cur.execute(insert3 + value3)
                    cnx.commit()
                    cur.close()
                    cnx.close()
                    tkinter.messagebox.showinfo("บันทึกฐานข้อมูล","บันทึกข้อมูลสำเร็จ")
                except:
                    print('ยังไม่มีข้อมูล')
                    cnx = pymysql.connect(user=textbox[5], password=textbox[6], host=textbox[7], database=textbox[8])
                    cur = cnx.cursor()
                    nameImg = "ImagesAll/ImagesAll/"+str(b[0])+".png"
                    insert1 = "INSERT INTO user (  ID_Card,Type_ID, Name_TH, Name_EN, Birth, GEN, Card_Issuer, Issue_Date, Expire_Date, Address,Email, Telephon, Department, Time_read_card, image, religion, age)"
                    value1 =  "VALUES ("+str(b[0])+ ",'ThaiCard', '"+b[1]+"', '"+b[2]+"' , "+str(b[3])+",  '"+b[5]+"', '"+b[6]+"', "+str(b[7])+", "+str(b[9])+", '"+b[11]+"','"+str(textbox[0])+"', '"+str(textbox[1])+"', '"+str(textbox[3])+"','"+b[12]+"', '"+nameImg+"',  '"+str(textbox[4])+"', "+str(b[14])+ ")"
                    insert2 = "INSERT INTO pairing (  ID_Card, Type_ID, TagID, Time_read_card)"
                    value2 =  "VALUES ("+str(b[0])+ ",'ThaiCard', '"+str(textbox[2])+"', '"+b[12]+"')"
                    
                    count1 = cur.execute( insert1 + value1)
                    count2 = cur.execute( insert2 + value2)

                    cnx.commit()
                    cur.close()
                    cnx.close()
                    #print(textbox)
                    im1 = Image.open('temp/'+b[0]+'.png') 
                    im2 = im1.copy() 
                    im2.save('result/ImagesAll/'+b[0]+'.png')
                    tkinter.messagebox.showinfo("บันทึกฐานข้อมูล","บันทึกข้อมูลสำเร็จ")
                    return     
        except:
            tkinter.messagebox.showinfo("เชื่อมต่อฐานข้อมูล","เชื่อมต่อฐานข้อมูลไม่สำเร็จ กรุณาตรวจสอบ Username Password IPAdress NameDB หรือสร้างฐานข้อมูลและตาราง หน้าคำแนะนำ")            
    def saveExcel():
        try:
            textbox = []
            reli_gion = entry_0.get()
            mail = entry_1.get()
            tel = entry_2.get()
            TagID = entry_3.get()
            depart = entry_4.get()
            textbox.append(mail)
            textbox.append(tel)
            textbox.append(TagID)
            textbox.append(depart)
            textbox.append(reli_gion)
            #fileName = textCard()
            b = textCard()
            nameImg = "result/ImagesAll/"+str(b[0])+".png"
            if textbox[2] == '':
                tkinter.messagebox.showinfo("แจ้งเตือน","กรุณากรอก TagID")
            else:
                filename = 'result/data smart card.xlsx'
                wb = openpyxl.load_workbook(filename=filename)
                sheet = wb['Sheet1']
                new_row = [b[0], 'ThaiCard', b[1], b[2], b[4], b[5], b[6], b[8], b[10], b[11],textbox[0],textbox[1],textbox[2],textbox[3],b[12],nameImg,textbox[4],b[14]]
                sheet.append(new_row)
                wb.save(filename)
                
                im1 = Image.open('temp/'+b[0]+'.png') 
                im2 = im1.copy() 
                im2.save('result/ImagesAll/'+b[0]+'.png') 
                tkinter.messagebox.showinfo("บันทึกข้อมูล Excel","บันทึกข้อมูลสำเร็จ")
        except:
            tkinter.messagebox.showinfo("บันทึกข้อมูล Excel","บันทึกไม่สำเร็จ กรุณาตรวจสอบบัตร หรือปิด Excel")
    def saveAll():
        saveInputDB()
        saveExcel()
    try: #---------------------------------------------ส่วนของการแสดงข้อมูล------------------------------------------------------
        
        label_0 = Label(page, text="ข้อมูลบัตรประชาชน",relief="solid",width=20,font=("arial", 19,"bold"))
        label_0.place(x=350,y=200)

        label_0 = Label(page, text="ไม่พบข้อมูล",relief="solid",width=30,height=20,font=("arial", 5,"bold"))
        label_0.place(x=435,y=45)

        #Name of type card ส่วน layout ที่บ่งบอกชนิดของข้อมูล
        label_1 = Label(page, text="เลขประจำตัวประชาชน :",width=20,font=("bold", 10),bg='#d9d9d9')
        label_1.place(x=0,y=240)
        label_2 = Label(page, text="ชื่อ-สกุล :",width=20,font=("bold", 10),bg='#d9d9d9')
        label_2.place(x=-36,y=270)
        label_3 = Label(page, text="Name-LastName:",width=20,font=("bold", 10),bg='#d9d9d9')
        label_3.place(x=480,y=270)
        label_4 = Label(page, text="เกิดวันที่ :",width=20,font=("bold", 10),bg='#d9d9d9')
        label_4.place(x=-38,y=300)
        label_12 = Label(page, text="Date of Birth:",width=20,font=("bold", 10),bg='#d9d9d9')
        label_12.place(x=470,y=300)
        label_5 = Label(page, text="เพศ :",width=20,font=("bold", 10),bg='#d9d9d9')
        label_5.place(x=445,y=240)
        label_6 = Label(page, text="สถานที่ออกบัตร :",width=20,font=("bold", 10),bg='#d9d9d9')
        label_6.place(x=-22,y=390)
        label_7 = Label(page, text="ออกบัตรวันที่ :",width=20,font=("bold", 10),bg='#d9d9d9')
        label_7.place(x=-25,y=330)
        label_8 = Label(page, text="บัตรหมดอายุวันที่ :",width=20,font=("bold", 10),bg='#d9d9d9')
        label_8.place(x=475,y=330)
        label_9 = Label(page, text="ที่อยู่ :",width=20,font=("bold", 10),bg='#d9d9d9')
        label_9.place(x=-46,y=360)
        label_11 = Label(page, text="เวลาอ่านบัตร :",width=20,font=("bold", 10),bg='#d9d9d9')
        label_11.place(x=465,y=360)
        label_17 = Label(page, text="อายุ :",width=20,font=("bold", 10),bg='#d9d9d9')
        label_17.place(x=450,y=390)

        label_13 = Label(page, text="ศาสนา :",width=20,font=("bold", 10),bg='#d9d9d9')
        label_13.place(x=-40,y=450)
        label_13 = Label(page, text="E-mail(อีเมล) :",width=20,font=("bold", 10),bg='#d9d9d9')
        label_13.place(x=-25,y = 510)
        label_14 = Label(page, text="Telephone(หมายเลขโทรศัพท์) :",width=20,font=("bold", 10),bg='#d9d9d9')
        label_14.place(x=515,y=450)
        label_15 = Label(page, text="Tag ID (*จำเป็น):",width=20,font=("bold", 10),bg='#d9d9d9')
        label_15.place(x=-20,y=480)
        label_16 = Label(page, text="Department(หน่วยงาน) :",width=20,font=("bold", 10),bg='#d9d9d9')
        label_16.place(x=500,y=480)

        #box for input กล่องข้อความสำหรับ กรอกข้อมูล
        mail   = StringVar()
        tel    = StringVar()
        TagID   = StringVar() #คือ Tag ID
        depart = StringVar()
        religion = StringVar()
        entry_0 = Entry(page,textvar= religion,width = 45)
        entry_0.place(x=170,y=450)
        entry_1 = Entry(page,textvar= mail,width = 45)
        entry_1.place(x=170,y=510)
        entry_2 = Entry(page,textvar=tel,width = 45)
        entry_2.place(x=680,y=450)
        entry_3 = Entry(page,textvar=TagID,width = 45)
        entry_3.place(x=170,y=480)
        entry_4 = Entry(page,textvar=depart,width = 45)
        entry_4.place(x=680,y=480)
    except:
        pass
    try:
        path = "user.png"
        img = ImageTk.PhotoImage(Image.open(path))
        panel = tk.Label(page, image = img)
        panel.image = img # keep a reference!
        panel.pack(side = "top", fill = "both", expand = "yes")
        panel.place(x=435,y=25)

        label_10 = Label(page, text='ไม่มีข้อมูล', anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
        label_10.place(x=150,y=240) #cid
        label_11 = Label(page, text='ไม่มีข้อมูล', anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
        label_11.place(x=150,y=270) #ชื่อไทย
        label_12 = Label(page, text='ไม่มีข้อมูล', anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
        label_12.place(x=625,y=270) #ชื่ออิ้ง
        label_20 = Label(page, text='ไม่มีข้อมูล', anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
        label_20.place(x=625,y=300) #วันเกิดอิ้ง
        label_13 = Label(page, text='ไม่มีข้อมูล', anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
        label_13.place(x=150,y=300) #วันเกิดไทย
        label_14 = Label(page, text='ไม่มีข้อมูล', anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
        label_14.place(x=625,y=240) #เพศ
        label_15 = Label(page, text='ไม่มีข้อมูล', anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
        label_15.place(x=150,y=390) #สถานที่ออกบัตร
        label_16 = Label(page, text='ไม่มีข้อมูล', anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
        label_16.place(x=150,y=330) #วันออกบัตร
        label_17 = Label(page, text='ไม่มีข้อมูล', anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
        label_17.place(x=625,y=330) #วันหมดอายุบัตร
        label_18 = Label(page, text='ไม่มีข้อมูล', anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
        label_18.place(x=150,y=360) #ที่อยู่
        label_19 = Label(page, text='ไม่มีข้อมูล', anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
        label_19.place(x=625,y=360) #เวลาอ่านบัตร
        label_21 = Label(page, text='ไม่มีข้อมูล', anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
        label_21.place(x=625,y=390) #อายุ
    except:
        pass
    try:
        fileName = textCard()# textCard() ทำหน้าที่รับค่าจาก CheckCard()
        photoCard(fileName)#รับอากิวเมน fileName ที่เป็น array โดย photoCard()ทำหน้าที่ดึงรูปและบันทึก
        resizeImg(fileName)#resizeImg()ทำหน้าที่ปรับขนาดรูปและบันทึก
        #return fileName
        a = fileName
        path = "temp/"+a[0]+".png"
        img = ImageTk.PhotoImage(Image.open(path))
        panel = tk.Label(page, image = img)
        panel.image = img # keep a reference!
        panel.pack(side = "top", fill = "both", expand = "yes")
        panel.place(x=435,y=25)

        label_10 = Label(page, text=a[0], anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
        label_10.place(x=150,y=240) #cid

        label_11 = Label(page, text=a[1], anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
        label_11.place(x=150,y=270) #ชื่อไทย

        label_12 = Label(page, text=a[2], anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
        label_12.place(x=625,y=270) #ชื่ออิ้ง

        label_20 = Label(page, text=a[13], anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
        label_20.place(x=625,y=300) #วันเกิดอิ้ง

        label_13 = Label(page, text=a[4], anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
        label_13.place(x=150,y=300) #วันเกิดไทย

        label_14 = Label(page, text=a[5], anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
        label_14.place(x=625,y=240) #เพศ

        label_15 = Label(page, text=a[6], anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
        label_15.place(x=150,y=390) #สถานที่ออกบัตร

        label_16 = Label(page, text=a[8], anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
        label_16.place(x=150,y=330) #วันออกบัตร

        label_17 = Label(page, text=a[10], anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
        label_17.place(x=625,y=330) #วันหมดอายุบัตร

        label_18 = Label(page, text=a[11], anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
        label_18.place(x=150,y=360) #ที่อยู่

        label_19 = Label(page, text=a[12], anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
        label_19.place(x=625,y=360) #เวลาอ่านบัตร

        label_21 = Label(page, text=a[14], anchor='w',relief="solid",width=40,font=("arial", 10,"bold"))
        label_21.place(x=625,y=390) #อายุ  
    except:
        pass

    #ฟังก์ชันการทำงานของปุ่ม  connectDB createTable
    button_refresh = Button(page,text="แสดงข้อมูล",width=12,bg='blue',fg='white',font=("bold", 10),command=refreshData).place(x=750,y=25)
    button_saveDB = Button(page, text='บันทึกในฐานข้อมูล',width=18,bg='green',fg='white',font=("bold", 10),command=saveInputDB).place(x=750,y=60)
    button_saveEXCEL = Button(page, text='บันทึกไฟล์ excel',width=18,bg='green',fg='white',font=("bold", 10),command=saveExcel).place(x=750,y=95)
    button_saveAll   = Button(page, text='บันทึกไฟล์ ทั้งหมด',width=18,bg='green',fg='white',font=("bold", 10),command=saveAll).place(x=750,y=130)
    button_quit = Button(page, text='ออก',width=12,bg='brown',fg='white',font=("bold", 10),command= ext).place(x=750,y=165)

def layout2(page):
    def clearTextBox():
        #เป็นส่วนของการ clear text box 
        entry_1.delete(first=0,last=100)
        entry_2.delete(first=0,last=100)
        entry_3.delete(first=0,last=100)
        entry_4.delete(first=0,last=100)
        entry_5.delete(first=0,last=100)
        entry_6.delete(first=0,last=100)
        entry_7.delete(first=0,last=100)
        entry_8.delete(first=0,last=100)
        entry_9.delete(first=0,last=100)
        label_1 = Label(page, text='เวลาจะแสดงเมื่อบันทึก', anchor='w',relief="solid",width=33,font=("arial", 10,"bold"))
        label_1.place(x=150,y=455) #เวลาอ่านบัตร
        label_2 = Label(page, text='ประเภทของบัตร อัตโนมัติ', anchor='w',relief="solid",width=33,font=("arial", 10,"bold"))
        label_2.place(x=150,y=480) #ประเภทของบัตร
    def abt(): #เป็นฟังก์ชันกล่อง pop up แสดงข้อความต่อเมื่อ ฟังก์ชัน subm2.add_command ถูกกด
        tkinter.messagebox.showinfo("เกี่ยวกับ","smart red v1. เป็นโปรแกรมสำหรับอ่านบัตรประชาชน พัฒนาโดยคนไทย")
    def ext():  #เป็นฟังก์ชันคำสั่งออกจากโปรแกรม  ซึ่องรอรับคำสั่งจาก subm1.add_command และ but_quit
        #page.quit()
        page.destroy()
        sys.exit()
    def saveInputDB():
        Cid =       entry_1.get()
        Province =    entry_2.get()
        NameTH =  entry_3.get()
        Tel =       entry_4.get()
        Gen =       entry_5.get()
        Religion =     entry_6.get()
        TagID =  entry_7.get()
        Department= entry_8.get()
        Email =     entry_9.get()
        userDB
        passDB
        hostDB
        nameDB
        time = str(datetime.datetime.now())
        textbox = [Cid ,NameTH  ,Gen ,TagID ,Province ,Tel ,Religion ,Department ,Email ,time,userDB ,passDB ,hostDB ,nameDB]
        try:
            #ต้องการตรวจสอบข้อมูลในการเชื่อมต่อฐานข้อมูล
            mydb = mysql.connector.connect(user=userDB, password=passDB, host=hostDB, database=nameDB)
            
            CheckNull = []
            for i in range(4):
                if textbox[i]=='':
                    if i == 0:
                        CheckNull.append(' เลขประจำตัวประชาชน หรือ Passport No,')
                    elif i == 1:
                        CheckNull.append(' ชื่อ-สกุล')
                    elif i == 2:
                        CheckNull.append(' เพศ')
                    elif i == 3:
                        CheckNull.append(' TagID')

            #print(CheckNull) ตรวจสอบว่าผู้ใช้กรอกข้อมูลครบ 4 ช่องหรือยัง

            if len(CheckNull) == 0: #เข้าเงื่อนไขนี้คือ ทุกช่องจะต้องไม่มีค่าว่าง
                CheckCID = textbox[0].strip() #ย่อย str เก็บอยู่ใน list
                EngBig = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
                EngSmall = ['a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z']
                sum1 = []
                #print(CheckCID[0])
                for k in range(len(EngBig)):
                    if EngBig[k] == CheckCID[0]:
                        print(k)
                        sum1.append(1)
                    elif CheckCID[0] == EngSmall[k]:
                        sum1.append(1)
                    else:
                        sum1.append(0)
                sum1.sort()
                if sum1[-1] == 1: #passport เงื่อนไขสำหรับเช็คว่าตัวเลขของชุดเป็นพยัญชนะหรือไม่
                    mydb = mysql.connector.connect(user=textbox[-4], password=textbox[-3], host=textbox[-2], database=textbox[-1])
                    mycursor = mydb.cursor()
                    sql = "SELECT * FROM user WHERE ID_Card = '"+textbox[0]+"' "
                    mycursor.execute(sql)
                    myresult = mycursor.fetchall()
                    try:
                        print(myresult[0][1])#เพื่อให้เช็คว่ามีข้อมูลใน DB เป็นการดักเพื่อเข้าหรือไม่เข้า ในเงื่อนไข excep
                        print('มีข้อมูลแล้ว')
                        if textbox[4]!='':    
                            mydb = mysql.connector.connect(user=textbox[-4], password=textbox[-3], host=textbox[-2], database=textbox[-1])
                            mycursor = mydb.cursor()
                            sql3 = "UPDATE user SET Address= '"+str(textbox[4])+"' WHERE ID_Card = '"+str(textbox[0])+"'"
                            mycursor.execute(sql3)
                            mydb.commit() 
                        if textbox[5]!='':    
                            mydb = mysql.connector.connect(user=textbox[-4], password=textbox[-3], host=textbox[-2], database=textbox[-1])
                            mycursor = mydb.cursor()
                            sql3 = "UPDATE user SET Telephon= '"+str(textbox[5])+"' WHERE ID_Card = '"+str(textbox[0])+"'"
                            mycursor.execute(sql3)
                            mydb.commit()
                        if textbox[6]!='':    
                            mydb = mysql.connector.connect(user=textbox[-4], password=textbox[-3], host=textbox[-2], database=textbox[-1])
                            mycursor = mydb.cursor()
                            sql3 = "UPDATE user SET religion=  '"+str(textbox[6])+"' WHERE ID_Card = '"+str(textbox[0])+"'"
                            mycursor.execute(sql3)
                            mydb.commit() 
                        if textbox[7]!='':    
                            mydb = mysql.connector.connect(user=textbox[-4], password=textbox[-3], host=textbox[-2], database=textbox[-1])
                            mycursor = mydb.cursor()
                            sql3 = "UPDATE user SET Department= '"+str(textbox[7])+"' WHERE ID_Card = '"+str(textbox[0])+"'"
                            mycursor.execute(sql3)
                            mydb.commit()
                        if textbox[8]!='':    
                            mydb = mysql.connector.connect(user=textbox[-4], password=textbox[-3], host=textbox[-2], database=textbox[-1])
                            mycursor = mydb.cursor()
                            sql3 = "UPDATE user SET Email= '"+str(textbox[8])+"' WHERE ID_Card = '"+str(textbox[0])+"'"
                            mycursor.execute(sql3)
                            mydb.commit()

                        mydb = mysql.connector.connect(user=textbox[-4], password=textbox[-3], host=textbox[-2], database=textbox[-1])
                        mycursor = mydb.cursor()
                        
                        #sql = "UPDATE user SET ID_Card="+str(b[0])+ ", Name_TH= '"+b[1]+"', Address= '"+b[11]+"', GEN=  '"+b[5]+"', Telephon= '"+str(textbox[1])+"', Email='"+str(textbox[0])+"', TagID= '"+str(textbox[2])+"', Department= '"+str(textbox[3])+"', Time_read_card='"+b[12]+"', image= '"+nameImg+"', religion=  '"+str(textbox[4])+"', age= "+str(b[14])+ " WHERE ID_Card = "+str(b[0])+ ""
                        sql1 = "UPDATE user SET ID_Card= '"+str(textbox[0])+"', Name_TH= '"+str(textbox[1])+"', GEN=  '"+str(textbox[2])+"', Time_read_card='"+textbox[9]+"' WHERE ID_Card = '"+str(textbox[0])+"'"
                        mycursor.execute(sql1)
                        mydb.commit()

                        cnx = pymysql.connect(user=textbox[-4], password=textbox[-3], host=textbox[-2], database=textbox[-1])
                        cur = cnx.cursor()
                        insert2 = "INSERT INTO pairing (  ID_Card, Type_ID, TagID, Time_read_card)"
                        value2 =  "VALUES ('"+str(textbox[0])+"', 'Passport', '"+str(textbox[3])+"', '"+textbox[9]+"')"
                        count2 = cur.execute( insert2 + value2)
                        cnx.commit()
                        cur.close()
                        cnx.close()
                        label_1 = Label(page, text=textbox[9], anchor='w',relief="solid",width=33,font=("arial", 10,"bold"))
                        label_1.place(x=150,y=455) #เวลาอ่านบัตร
                        label_2 = Label(page, text='พาสปอร์ต(Passport)', anchor='w',relief="solid",width=33,font=("arial", 10,"bold"))
                        label_2.place(x=150,y=480) #ประเภทของบัตร

                        tkinter.messagebox.showinfo("บันทึกฐานข้อมูล","บันทึกข้อมูลสำเร็จ")
                    except:
                        print('ยังไม่มีข้อมูล')
                        cnx = pymysql.connect(user=textbox[-4], password=textbox[-3], host=textbox[-2], database=textbox[-1])
                        cur = cnx.cursor()
                        insert1 = "INSERT INTO user (  ID_Card, Type_ID, Name_TH, GEN, Address, Telephon, religion, Department, Email, Time_read_card)"
                        value1 =  "VALUES ('"+str(textbox[0])+"', '"+'Passport'+"', '"+str(textbox[1])+"', '"+str(textbox[2])+"', '"+str(textbox[4])+"', '"+str(textbox[5])+"', '"+str(textbox[6])+"', '"+str(textbox[7])+"', '"+str(textbox[8])+"', '"+textbox[9]+"')"
                        insert2 = "INSERT INTO pairing (  ID_Card, Type_ID, TagID, Time_read_card)"
                        value2 =  "VALUES ('"+str(textbox[0])+"', 'Passport', '"+str(textbox[3])+"', '"+textbox[9]+"')"
                        count1 = cur.execute( insert1 + value1)
                        count2 = cur.execute( insert2 + value2)
                        cnx.commit()
                        cur.close()
                        cnx.close()
                        label_1 = Label(page, text=textbox[9], anchor='w',relief="solid",width=33,font=("arial", 10,"bold"))
                        label_1.place(x=150,y=455) #เวลาอ่านบัตร
                        label_2 = Label(page, text='พาสปอร์ต(Passport)', anchor='w',relief="solid",width=33,font=("arial", 10,"bold"))
                        label_2.place(x=150,y=480) #ประเภทของบัตร
                        tkinter.messagebox.showinfo("บันทึกฐานข้อมูล","บันทึกข้อมูลสำเร็จ")
                        return
                elif len(CheckCID) == 13: #เงื่อนไขสำหรับเช็ค CID ว่ากรอกครบ 13ตัว ถ้ากดตัวอักษร ก็จะนับด้วย
                    countCID = []
                    for j in range(13):
                        if CheckCID[j] == '0':
                            countCID.append(CheckCID[j])
                        elif CheckCID[j] == '1':
                            countCID.append(CheckCID[j])
                        elif CheckCID[j] == '2':
                            countCID.append(CheckCID[j])
                        elif CheckCID[j] == '3':
                            countCID.append(CheckCID[j])
                        elif CheckCID[j] == '4':
                            countCID.append(CheckCID[j])
                        elif CheckCID[j] == '5':
                            countCID.append(CheckCID[j])
                        elif CheckCID[j] == '6':
                            countCID.append(CheckCID[j])
                        elif CheckCID[j] == '7':
                            countCID.append(CheckCID[j])
                        elif CheckCID[j] == '8':
                            countCID.append(CheckCID[j])
                        elif CheckCID[j] == '9':
                            countCID.append(CheckCID[j])
                    print(len(countCID))
                    if len(countCID) == 13: #ถ้าเลขบัตร ปปช. ครบ 13 เฉพาะตัวเลขจะเข้าเงื่อนไขนี้
                        mydb = mysql.connector.connect(user=textbox[-4], password=textbox[-3], host=textbox[-2], database=textbox[-1])
                        mycursor = mydb.cursor()
                        sql = "SELECT * FROM user WHERE ID_Card = '"+textbox[0]+"' "
                        mycursor.execute(sql)
                        myresult = mycursor.fetchall()
                        try:
                            print(myresult[0][1])#เพื่อให้เช็คว่ามีข้อมูลใน DB เป็นการดักเพื่อเข้าหรือไม่เข้า ในเงื่อนไข excep
                            print('มีข้อมูลแล้ว')
                            if textbox[4]!='':    
                                mydb = mysql.connector.connect(user=textbox[-4], password=textbox[-3], host=textbox[-2], database=textbox[-1])
                                mycursor = mydb.cursor()
                                sql3 = "UPDATE user SET Address= '"+str(textbox[4])+"' WHERE ID_Card = '"+str(textbox[0])+"'"
                                mycursor.execute(sql3)
                                mydb.commit() 
                            if textbox[5]!='':    
                                mydb = mysql.connector.connect(user=textbox[-4], password=textbox[-3], host=textbox[-2], database=textbox[-1])
                                mycursor = mydb.cursor()
                                sql3 = "UPDATE user SET Telephon= '"+str(textbox[5])+"' WHERE ID_Card = '"+str(textbox[0])+"'"
                                mycursor.execute(sql3)
                                mydb.commit()
                            if textbox[6]!='':    
                                mydb = mysql.connector.connect(user=textbox[-4], password=textbox[-3], host=textbox[-2], database=textbox[-1])
                                mycursor = mydb.cursor()
                                sql3 = "UPDATE user SET religion=  '"+str(textbox[6])+"' WHERE ID_Card = '"+str(textbox[0])+"'"
                                mycursor.execute(sql3)
                                mydb.commit() 
                            if textbox[7]!='':    
                                mydb = mysql.connector.connect(user=textbox[-4], password=textbox[-3], host=textbox[-2], database=textbox[-1])
                                mycursor = mydb.cursor()
                                sql3 = "UPDATE user SET Department= '"+str(textbox[7])+"' WHERE ID_Card = '"+str(textbox[0])+"'"
                                mycursor.execute(sql3)
                                mydb.commit()
                            if textbox[8]!='':    
                                mydb = mysql.connector.connect(user=textbox[-4], password=textbox[-3], host=textbox[-2], database=textbox[-1])
                                mycursor = mydb.cursor()
                                sql3 = "UPDATE user SET Email= '"+str(textbox[8])+"' WHERE ID_Card = '"+str(textbox[0])+"'"
                                mycursor.execute(sql3)
                                mydb.commit()


                            mydb = mysql.connector.connect(user=textbox[-4], password=textbox[-3], host=textbox[-2], database=textbox[-1])
                            mycursor = mydb.cursor()
                            sql1 = "UPDATE user SET ID_Card= '"+str(textbox[0])+"', Name_TH= '"+str(textbox[1])+"', GEN=  '"+str(textbox[2])+"', Time_read_card='"+textbox[9]+"' WHERE ID_Card = '"+str(textbox[0])+"'"
                            mycursor.execute(sql1)
                            mydb.commit()

                            cnx = pymysql.connect(user=textbox[-4], password=textbox[-3], host=textbox[-2], database=textbox[-1])
                            cur = cnx.cursor()
                            insert2 = "INSERT INTO pairing (  ID_Card, Type_ID, TagID, Time_read_card)"
                            value2 =  "VALUES ('"+str(textbox[0])+"', 'ThaiCard', '"+str(textbox[3])+"', '"+textbox[9]+"')"
                            count2 = cur.execute( insert2 + value2)
                            cnx.commit()
                            cur.close()
                            cnx.close()

                            label_1 = Label(page, text=textbox[9], anchor='w',relief="solid",width=33,font=("arial", 10,"bold"))
                            label_1.place(x=150,y=455) #เวลาอ่านบัตร
                            label_2 = Label(page, text='บัตรประชาชน(ThaiCard)', anchor='w',relief="solid",width=33,font=("arial", 10,"bold"))
                            label_2.place(x=150,y=480) #ประเภทของบัตร
                            tkinter.messagebox.showinfo("บันทึกฐานข้อมูล","บันทึกข้อมูลสำเร็จ")
                        except:
                            print('ยังไม่มีข้อมูล')
                            cnx = pymysql.connect(user=textbox[-4], password=textbox[-3], host=textbox[-2], database=textbox[-1])
                            cur = cnx.cursor()
                            insert1 = "INSERT INTO user (  ID_Card, Type_ID, Name_TH, GEN, Address, Telephon, religion, Department, Email, Time_read_card)"
                            value1 =  "VALUES ('"+str(textbox[0])+"', '"+'ThaiCard'+"', '"+str(textbox[1])+"', '"+str(textbox[2])+"', '"+str(textbox[4])+"', '"+str(textbox[5])+"', '"+str(textbox[6])+"', '"+str(textbox[7])+"', '"+str(textbox[8])+"', '"+textbox[9]+"')"
                            insert2 = "INSERT INTO pairing (  ID_Card, Type_ID, TagID, Time_read_card)"
                            value2 =  "VALUES ('"+str(textbox[0])+"', 'ThaiCard', '"+str(textbox[3])+"', '"+textbox[9]+"')"
                            count1 = cur.execute( insert1 + value1)
                            count2 = cur.execute( insert2 + value2)
                            cnx.commit()
                            cur.close()
                            cnx.close()
                            label_1 = Label(page, text=textbox[9], anchor='w',relief="solid",width=33,font=("arial", 10,"bold"))
                            label_1.place(x=150,y=455) #เวลาอ่านบัตร
                            label_2 = Label(page, text='บัตรประชาชน(ThaiCard)', anchor='w',relief="solid",width=33,font=("arial", 10,"bold"))
                            label_2.place(x=150,y=480) #ประเภทของบัตร
                            tkinter.messagebox.showinfo("บันทึกฐานข้อมูล","บันทึกข้อมูลสำเร็จ")
                            return
                    else:
                        tkinter.messagebox.showinfo("การบันทึกฐานข้อมูล",'กรุณากรอกเลขบัตรประชาชนให้ถูกต้องเนื่องจากมีพยัญชนะ')
                        
                
                else:
                    tkinter.messagebox.showinfo('กรอกข้อมูลไม่ถูกต้อง','กรุณากรอกข้อมูลเลขบัตรประชาชนให้ครบ 13 หลัก หรือ passport ที่มีอักษรภาษาอังกฤษข้างหน้า')
                    
                
            else:
                tkinter.messagebox.showinfo('กรุณากรอกข้อมูลให้ครบถ้วน',CheckNull)
        except: 
            tkinter.messagebox.showinfo("เชื่อมต่อฐานข้อมูล","เชื่อมต่อฐานข้อมูลไม่สำเร็จ กรุณาตรวจสอบ: Username Password IPAdress NameDB หรือสร้างฐานข้อมูลและตาราง หน้าคำแนะนำ")
    def saveExcel():
              
        Cid =       entry_1.get()
        Province =    entry_2.get()
        NameTH =  entry_3.get()
        Tel =       entry_4.get()
        Gen =       entry_5.get()
        Religion =     entry_6.get()
        TagID =  entry_7.get()
        Department= entry_8.get()
        Email =     entry_9.get()
        time = str(datetime.datetime.now())
        textbox = [Cid ,NameTH  ,Gen ,TagID ,Province ,Tel ,Religion ,Department ,Email ,time]
        
        CheckNull = []
        for i in range(4):
            if textbox[i]=='':
                if i == 0:
                    CheckNull.append(' เลขประจำตัวประชาชน หรือ Passport No,')
                elif i == 1:
                    CheckNull.append(' ชื่อ-สกุล')
                elif i == 2:
                    CheckNull.append(' เพศ')
                elif i == 3:
                    CheckNull.append(' TagID')

        if len(CheckNull) == 0: #เข้าเงื่อนไขนี้คือ ทุกช่องจะต้องไม่มีค่าว่าง
            CheckCID = textbox[0].strip() #ย่อย str เก็บอยู่ใน list
            EngBig = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
            EngSmall = ['a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z']
            sum1 = []
            #print(CheckCID[0])
            for k in range(len(EngBig)):
                if EngBig[k] == CheckCID[0]:
                    print(k)
                    sum1.append(1)
                elif CheckCID[0] == EngSmall[k]:
                    sum1.append(1)
                else:
                    sum1.append(0)
            sum1.sort()

            if sum1[-1] == 1: #เงื่อนไขสำหรับ Passport เช็คว่าตัวเลขของชุดเป็นพยัญชนะหรือไม่
                
                try:
                    filename = 'result/data smart card.xlsx'
                    wb = openpyxl.load_workbook(filename=filename)
                    sheet = wb['Sheet1']
                    new_row = [Cid ,"Passport",NameTH  ," "," ",Gen ," "," "," ",Province ,Email ,Tel ,TagID ,Department ,time," ",Religion ]
                    sheet.append(new_row)
                    wb.save(filename)
                    
                    label_1 = Label(page, text=textbox[9], anchor='w',relief="solid",width=33,font=("arial", 10,"bold"))
                    label_1.place(x=150,y=455) #เวลาอ่านบัตร
                    label_2 = Label(page, text='พาสปอร์ต(Passport)', anchor='w',relief="solid",width=33,font=("arial", 10,"bold"))
                    label_2.place(x=150,y=480) #ประเภทของบัตร
                    tkinter.messagebox.showinfo("บันทึกข้อมูล Excel","บันทึกข้อมูลสำเร็จ")
                except:
                    tkinter.messagebox.showinfo("บันทึกข้อมูล Excel","บันทึกไม่สำเร็จ กรุณาตรวจสอบบัตร หรือปิด Excel")
                    
            elif len(CheckCID) == 13: #เงื่อนไขสำหรับเช็ค CID ว่ากรอกครบ 13ตัว ถ้ากดตัวอักษร ก็จะนับด้วย
                countCID = []
                for j in range(13):
                    if CheckCID[j] == '0':
                        countCID.append(CheckCID[j])
                    elif CheckCID[j] == '1':
                        countCID.append(CheckCID[j])
                    elif CheckCID[j] == '2':
                        countCID.append(CheckCID[j])
                    elif CheckCID[j] == '3':
                        countCID.append(CheckCID[j])
                    elif CheckCID[j] == '4':
                        countCID.append(CheckCID[j])
                    elif CheckCID[j] == '5':
                        countCID.append(CheckCID[j])
                    elif CheckCID[j] == '6':
                        countCID.append(CheckCID[j])
                    elif CheckCID[j] == '7':
                        countCID.append(CheckCID[j])
                    elif CheckCID[j] == '8':
                        countCID.append(CheckCID[j])
                    elif CheckCID[j] == '9':
                        countCID.append(CheckCID[j])
                print(len(countCID))
                if len(countCID) == 13: #ถ้าเลขบัตร ปปช. ครบ 13 เฉพาะตัวเลขจะเข้าเงื่อนไขนี้
                    try:
                        filename = 'result/data smart card.xlsx'
                        wb = openpyxl.load_workbook(filename=filename)
                        sheet = wb['Sheet1']
                        new_row = [Cid ,"ThaiCard",NameTH  ," "," ",Gen ," "," "," ",Province ,Email ,Tel ,TagID ,Department ,time," ",Religion ]
                        sheet.append(new_row)
                        wb.save(filename)
                        
                        label_1 = Label(page, text=textbox[9], anchor='w',relief="solid",width=33,font=("arial", 10,"bold"))
                        label_1.place(x=150,y=455) #เวลาอ่านบัตร
                        label_2 = Label(page, text='บัตรประชาชน(ThaiCard))', anchor='w',relief="solid",width=33,font=("arial", 10,"bold"))
                        label_2.place(x=150,y=480) #ประเภทของบัตร
                        tkinter.messagebox.showinfo("บันทึกข้อมูล Excel","บันทึกข้อมูลสำเร็จ")
                    except:
                        tkinter.messagebox.showinfo("บันทึกข้อมูล Excel","บันทึกไม่สำเร็จ กรุณาตรวจสอบบัตร หรือปิด Excel")
                else:
                    tkinter.messagebox.showinfo("การบันทึกฐานข้อมูล",'กรุณากรอกเลขบัตรประชาชนให้ถูกต้องเนื่องจากมีพยัญชนะ')
                    
            
            else:
                tkinter.messagebox.showinfo('กรอกข้อมูลไม่ถูกต้อง','กรุณากรอกข้อมูลเลขบัตรประชาชนให้ครบ 13 หลัก หรือ passport ที่มีอักษรภาษาอังกฤษข้างหน้า')
                
            
        else:
            tkinter.messagebox.showinfo('กรุณากรอกข้อมูลให้ครบถ้วน',CheckNull)
    def saveAll():
        saveInputDB()
        saveExcel()
    def title():
        #---------------------------------------------ส่วนของการแสดงข้อมูล------------------------------------------------------
        path = "user.png"
        img = ImageTk.PhotoImage(Image.open(path))
        panel = tk.Label(page, image = img)
        panel.image = img # keep a reference!
        panel.pack(side = "top", fill = "both", expand = "yes")
        panel.place(x=435,y=25)
        label_0 = Label(page, text="กรอกข้อมูลผู้ใช้งาน",relief="solid",width=20,font=("arial", 19,"bold"))
        label_0.place(x=350,y=200)
        
        my_font = Font( size=16, weight="bold", underline=1)
        label_13 = Label(page, text="* ข้อมูลจำเป็น", font=my_font,bg='#d9d9d9')
        label_13.place(x=15,y=290)
        label_14 = Label(page, text="สามารถใส่ข้อมูล หรือไม่ใส่ข้อมูล", font=my_font,bg='#d9d9d9')
        label_14.place(x=510,y=290)

        #Name of type card ส่วน layout ที่บ่งบอกชนิดของข้อมูล
        label_1 = Label(page, text="*เลขประจำตัวประชาชน",width=20,font=("bold", 10),bg='#d9d9d9')
        label_1.place(x=-5,y=325)
        label_1 = Label(page, text="หรือ Passport No :",width=20,font=("bold", 10),bg='#d9d9d9')
        label_1.place(x=-5,y=345)
        label_5 = Label(page, text="จังหวัดตามทะเบียนบ้าน :",width=20,font=("bold", 10),bg='#d9d9d9')
        label_5.place(x=490,y=330)

        label_2 = Label(page, text="ชื่อ-สกุล :",width=20,font=("bold", 10),bg='#d9d9d9')
        label_2.place(x=-35,y=380)
        label_3 = Label(page, text="เบอร์โทร :",width=20,font=("bold", 10),bg='#d9d9d9')
        label_3.place(x=455,y=380)

        label_4 = Label(page, text="*เพศ :",width=20,font=("bold", 10),bg='#d9d9d9')
        label_4.place(x=-45,y=405)
        label_12 = Label(page, text="ศาสนา :",width=20,font=("bold", 10),bg='#d9d9d9')
        label_12.place(x=450,y=405)



        label_7 = Label(page, text="*TagID :",width=20,font=("bold", 10),bg='#d9d9d9')
        label_7.place(x=-40,y=430)
        label_8 = Label(page, text="หน่วยงาน :",width=20,font=("bold", 10),bg='#d9d9d9')
        label_8.place(x=455,y=430)

        label_9 = Label(page, text="เวลาบันทึกข้อมูล :",width=20,font=("bold", 10),bg='#d9d9d9')
        label_9.place(x=-15,y=455)
        label_11 = Label(page, text="Email :",width=20,font=("bold", 10),bg='#d9d9d9')
        label_11.place(x=450,y=455)
        label_12 = Label(page, text="ประเภทบัตร :",width=20,font=("bold", 10),bg='#d9d9d9')
        label_12.place(x=-30,y=475)
    title()

    try:  #ประกาศค่าและวางตำแหน่งตัวหนังสือ
        Cid = StringVar()
        NameTH = StringVar()
        Province = StringVar()
        Gen = StringVar()
        Tel = StringVar()
        TagID = StringVar()
        Religion = StringVar()
        Department = StringVar()
        Email = StringVar()

        entry_1 = Entry(page,textvar= Cid,width = 45)
        entry_1.place(x=150,y=330)
        entry_2 = Entry(page,textvar= Province,width = 45)
        entry_2.place(x=650,y=330)
        entry_3 = Entry(page,textvar= NameTH,width = 45)
        entry_3.place(x=150,y=380)
        entry_4 = Entry(page,textvar= Tel,width = 45)
        entry_4.place(x=650,y=380)
        entry_5 = Entry(page,textvar= Gen,width = 45)
        entry_5.place(x=150,y=405)
        entry_6 = Entry(page,textvar= Religion,width = 45)
        entry_6.place(x=650,y=405)
        entry_7 = Entry(page,textvar= TagID,width = 45)
        entry_7.place(x=150,y=430)
        entry_8 = Entry(page,textvar= Department,width = 45)
        entry_8.place(x=650,y=430)
        entry_9 = Entry(page,textvar= Email,width = 45)
        entry_9.place(x=650,y=455)
        label_1 = Label(page, text='เวลาจะแสดงเมื่อบันทึก', anchor='w',relief="solid",width=33,font=("arial", 10,"bold"))
        label_1.place(x=150,y=455) #เวลาอ่านบัตร
        label_2 = Label(page, text='ประเภทของบัตร อัตโนมัติ', anchor='w',relief="solid",width=33,font=("arial", 10,"bold"))
        label_2.place(x=150,y=480) #ประเภทของบัตร
        
    except:
        pass


    #ฟังก์ชันการทำงานของปุ่ม
    button_refresh = Button(page,text="ล้างข้อมูล",width=12,bg='blue',fg='white',font=("bold", 10),command=clearTextBox).place(x=750,y=25)
    button_saveDB = Button(page, text='บันทึกในฐานข้อมูล',width=18,bg='green',fg='white',font=("bold", 10),command=saveInputDB).place(x=750,y=60)
    button_saveEXCEL = Button(page, text='บันทึกไฟล์ excel',width=18,bg='green',fg='white',font=("bold", 10),command=saveExcel).place(x=750,y=95)
    button_saveAll   = Button(page, text='บันทึกไฟล์ ทั้งหมด',width=18,bg='green',fg='white',font=("bold", 10),command=saveAll).place(x=750,y=130)
    button_quit = Button(page, text='ออก',width=12,bg='brown',fg='white',font=("bold", 10),command= ext).place(x=750,y=165)

def layout3(page):
    def login():
        try:
            global  userDB
            global  passDB
            global  hostDB
            global  nameDB
            userDB = entry_5.get()
            passDB = entry_6.get()
            hostDB = entry_7.get()
            nameDB = entry_8.get()
            try:
                mydb = mysql.connector.connect(user=userDB, password=passDB, host=hostDB, database=nameDB)
                mycursor = mydb.cursor()
                mycursor.execute("SHOW TABLES")
                tables = mycursor.fetchall()
                nameTable = [0,0]
                for table in tables:
                    print(table)
                    if table == ('user',):
                        nameTable.append(1)
                    elif table == ('pairing',):
                        nameTable.append(2)
                    else:
                        nameTable.append(0)
                nameTable.sort()
                if nameTable[-2] == 1 and nameTable[-1] == 2:
                    tkinter.messagebox.showinfo("เชื่อมต่อฐานข้อมูล","เชื่อมต่อฐานข้อมูลสำเร็จ")
                else:
                    tkinter.messagebox.showinfo("เชื่อมต่อฐานข้อมูล","เชื่อมต่อฐานข้อมูลสำเร็จ (*ผู้ใช้ยังไม่มีตารางในฐานข้อมูล)")
            except ValueError:
                pass
        except:
            tkinter.messagebox.showinfo("เชื่อมต่อฐานข้อมูล","เชื่อมต่อฐานข้อมูลไม่สำเร็จ กรุณาตรวจสอบ Username Password IPAdress NameDB หรือสร้างฐานข้อมูลและตาราง หน้าคำแนะนำ")
    def createDB():
        global  userDB
        global  passDB
        global  hostDB
        global  nameDB
        userDB = entry_5.get()
        passDB = entry_6.get()
        hostDB = entry_7.get()
        nameDB = entry_8.get()
        try: #สร้างฐานข้อมูลและตารรางอัตโนมัติ   utf8_general_ci
            mydb = mysql.connector.connect(user=userDB, password=passDB, host=hostDB)
            mycursor = mydb.cursor()
            mycursor.execute("CREATE DATABASE "+nameDB+" DEFAULT COLLATE = utf8_general_ci")

        except:
            pass
        try: #สร้างตารรางอัตโนมัติ
            db_connection = mydb = mysql.connector.connect(user=userDB, password=passDB, host=hostDB, database=nameDB)
            db_cursor = db_connection.cursor()
            db_cursor.execute("CREATE TABLE user( id INT AUTO_INCREMENT PRIMARY KEY, ID_Card VARCHAR(50) UNIQUE, Type_ID VARCHAR(50), Name_TH VARCHAR(50), Name_EN VARCHAR(50), Birth DATE, GEN VARCHAR(10), Card_Issuer VARCHAR(100), Issue_Date DATE, Expire_Date DATE, Address VARCHAR(100), Email VARCHAR(80), Telephon VARCHAR(20), Department VARCHAR(30), Time_read_card DATETIME(6), image VARCHAR(100), religion VARCHAR(15), age INT(120)) DEFAULT COLLATE=	utf8_general_ci")
            db_cursor.execute("CREATE TABLE pairing( id INT AUTO_INCREMENT PRIMARY KEY, ID_Card VARCHAR(20), Type_ID VARCHAR(20), TagID VARCHAR(15), Type_User VARCHAR(20), Time_read_card DATETIME(6)) DEFAULT COLLATE=	utf8_general_ci")
            tkinter.messagebox.showinfo("การแจ้งเตือน","สร้างฐานข้อมูลสำเร็จ")
        except:
            try:
                mydb = mysql.connector.connect(user=userDB, password=passDB, host=hostDB, database=nameDB)
                mycursor = mydb.cursor()
                mycursor.execute("SHOW TABLES")
                tables = mycursor.fetchall()
                nameTable = [0,0]
                for table in tables:
                    print(table)
                    if table == ('user',):
                        nameTable.append(1)
                    elif table == ('pairing',):
                        nameTable.append(2)
                    else:
                        nameTable.append(0)
                nameTable.sort()
                if nameTable[-2] == 1 and nameTable[-1] == 2:
                    tkinter.messagebox.showinfo("การแจ้งเตือน","ผู้ใช้มีฐานมูลและตารางแล้ว")
            except:
                tkinter.messagebox.showinfo("การแจ้งเตือน","ชื่อฐานข้อมูลห้ามมีอักขระพิเศษหรือเว้นวรรค")
            
        

    label_0 = Label(page, text="คำแนะนำเบื้องต้น",relief="solid",width=20,font=("arial", 19,"bold"))
    label_0.place(x=0,y=210)
    label_a = Label(page, text="ปุ่มสร้างตารางหรือฐานข้อมูล",width=30,font=("bold", 15),bg='#d9d9d9')
    label_a.place(x=-40,y=330)
    label_a = Label(page, text="กรณีที่ 1 มีฐานข้อมูลแล้ว  ต้องการสร้างตารางสำหรับเก็บข้อมูล ให้ผู้ใช้กรอกข้อมูลให้ครบ",width=80,font=("bold", 15),bg='#d9d9d9')
    label_a.place(x=-70,y=370)
    label_a = Label(page, text="แล้วใส่ชื่อฐานข้อมูลที่ผู้ใช้สร้าง กดปุ่มสร้างตารางหรือฐานข้อมูล",width=60,font=("bold", 15),bg='#d9d9d9')
    label_a.place(x=-60,y=410)
    label_a = Label(page, text="กรณีที่ 2 ยังไม่มีฐานข้อมูล  ต้องการสร้างฐานข้อมูลและตารางสำหรับเก็บข้อมูล ให้ผู้ใช้กรอกข้อมูลให้ครบ",width=80,font=("bold", 15),bg='#d9d9d9')
    label_a.place(x=-20,y=450)
    label_a = Label(page, text="แล้วใส่ชื่อฐานข้อมูลที่ผู้ใช้ต้องการสร้าง กดปุ่มสร้างตารางหรือฐานข้อมูล",width=60,font=("bold", 15),bg='#d9d9d9')
    label_a.place(x=-30,y=490)

    label_b = Label(page, text="ปุ่ม Login",width=30,font=("bold", 15),bg='#d9d9d9')
    label_b.place(x=-110,y=250)
    label_b = Label(page, text="ใส่ข้อมูลให้ครบถ้วน แล้วกดปุ่ม login เพื่อใช้งานฐานข้อมูล ",width=60,font=("bold", 15),bg='#d9d9d9')
    label_b.place(x=-80,y=290)

    label_a = Label(page, text="User Name",width=20,font=("bold", 10),bg='#d9d9d9')
    label_a.place(x=185,y=20)
    label_b = Label(page, text="Password",width=20,font=("bold", 10),bg='#d9d9d9')
    label_b.place(x=182,y=50)
    label_c = Label(page, text="IP Adress",width=20,font=("bold", 10),bg='#d9d9d9')
    label_c.place(x=182,y=80)
    label_d = Label(page, text="Name Database",width=20,font=("bold", 10),bg='#d9d9d9')
    label_d.place(x=200,y=110)
    
    #ประกาศค่าสำหรับ text box
    userDB = StringVar()
    passDB = StringVar()
    hostDB = StringVar()
    nameDB = StringVar()
    #โชว์ค่าเริ่มต้น บน text box
    entry_5 = Entry(page,textvar= userDB,width = 35)
    entry_5.insert(10, "root")
    entry_5.place(x=20,y=20)

    entry_6 = Entry(page,textvar=passDB,width = 35,show = '*')
    entry_6.insert(10, "")
    entry_6.place(x=20,y=50)

    entry_7 = Entry(page,textvar=hostDB,width = 35)
    entry_7.insert(10, "localhost")
    entry_7.place(x=20,y=80)

    entry_8 = Entry(page,textvar=nameDB,width = 35)
    entry_8.insert(10, "smart_card")
    entry_8.place(x=20,y=110)
    button_createTable = Button(page, text='สร้างตารางหรือฐานข้อมูล',width=20,bg='green',fg='white',font=("bold", 10),command= createDB).place(x=200,y=150)
    button_createTable = Button(page, text='login',width=20,bg='blue',fg='white',font=("bold", 10),command= login).place(x=20,y=150)


if __name__ == "__main__":
    Main()

   
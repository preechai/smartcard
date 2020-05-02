import os,io,sys,time,datetime,binascii,codecs,xlwt
#เรียกใช้ module ที่มาจาก tkinter สำหรับการสร้างหน้าตา Gui
import sqlite3,tkinter.messagebox,tkinter.filedialog,glob
import tkinter as tk
from tkinter import *
from PIL import Image,ImageTk
from resizeimage import resizeimage
from PIL import Image
from xlwt import Workbook 
#เรียกใช้ module ที่มาจาก pyscard สำหรับดึงข้อมูลจากบัตร
from smartcard.System import readers
from smartcard.util import toHexString
from smartcard.CardType import AnyCardType
from smartcard.CardRequest import CardRequest
from smartcard.CardConnection import CardConnection
from smartcard.CardMonitoring import CardMonitor, CardObserver
from smartcard.util import HexListToBinString, toHexString, toBytes
from smartcard.CardConnectionObserver import ConsoleCardConnectionObserver

# สร้าง dictionary เพื่อเปรียบเทียบข้อมูลที่ดึงมาจากบัตร
tis620encoding = {

    32:' ', 33:'!', 34:'"', 35:'#', 36:'$', 37:'%', 38:'&', 39:"'", 40:'(', 41:')', 42:'*', 43:'+', 44:',', 45:'-', 46:'.',
    47:'/', 48:'0', 49:'1', 50:'2', 51:'3', 52:'4', 53:'5', 54:'6', 55:'7', 56:'8', 57:'9', 58:':', 59:';', 60:'<', 61:'=',
    62:'>', 63:'?', 64:'@', 65:'A', 66:'B', 67:'C', 68:'D', 69:'E', 70:'F', 71:'G', 72:'H', 73:'I', 74:'J', 75:'K', 76:'L',
    77:'M', 78:'N', 79:'O', 80:'P', 81:'Q', 82:'R', 83:'S', 84:'T', 85:'U', 86:'V', 87:'W', 88:'X', 89:'Y', 90:'Z', 91:'[',
    92:'\\', 93:']', 94:'^', 95:'_', 96:'`', 97:'a', 98:'b', 99:'c', 100:'d', 101:'e', 102:'f', 103:'g', 104:'h', 105:'i',
    106:'j', 107:'k', 108:'l', 109:'m', 110:'n', 111:'o', 112:'p', 113:'q', 114:'r', 115:'s', 116:'t', 117:'u', 118:'v',
    119:'w', 120:'x', 121:'y', 122:'z', 123:'{', 124:'|', 125:'}', 126:'~',

    161:'\u0e01',162:'\u0e02',163:'\u0e03',164:'\u0e04',165:'\u0e05',166:'\u0e06',167:'\u0e07',168:'\u0e08',169:'\u0e09',
    170:'\u0e0a',171:'\u0e0b',172:'\u0e0c',173:'\u0e0d',174:'\u0e0e',175:'\u0e0f',176:'\u0e10',177:'\u0e11',178:'\u0e12',
    179:'\u0e13',180:'\u0e14',181:'\u0e15',182:'\u0e16',183:'\u0e17',184:'\u0e18',185:'\u0e19',186:'\u0e1a',187:'\u0e1b',
    188:'\u0e1c',189:'\u0e1d',190:'\u0e1e',191:'\u0e1f',192:'\u0e20',193:'\u0e21',194:'\u0e22',195:'\u0e23',196:'\u0e24',
    197:'\u0e25',198:'\u0e26',199:'\u0e27',200:'\u0e28',201:'\u0e29',202:'\u0e2a',203:'\u0e2b',204:'\u0e2c',205:'\u0e2d',
    206:'\u0e2e',207:'\u0e2f',208:'\u0e30',209:'\u0e31',210:'\u0e32',211:'\u0e33',212:'\u0e34',213:'\u0e35',214:'\u0e36',
    215:'\u0e37',216:'\u0e38',217:'\u0e39',218:'\u0e3a',

    223:'\u0e3f',224:'\u0e40',225:'\u0e41',226:'\u0e42',227:'\u0e43',228:'\u0e44',229:'\u0e45',230:'\u0e46',231:'\u0e47',
    232:'\u0e48',233:'\u0e49',234:'\u0e4a',235:'\u0e4b',236:'\u0e4c',237:'\u0e4d',238:'\u0e4e',239:'\u0e4f',240:'\u0e50',
    241:'\u0e51',242:'\u0e52',243:'\u0e53',244:'\u0e54',245:'\u0e55',246:'\u0e56',247:'\u0e57',248:'\u0e58',249:'\u0e59',
    250:'\u0e5a',251:'\u0e5b'}

try :
    SELECT = [0x00, 0xA4, 0x04, 0x00, 0x08] # Check card
    THAI_CARD = [0xA0, 0x00, 0x00, 0x00, 0x54, 0x48, 0x00, 0x01]
    CMD_CID = [0x80, 0xb0, 0x00, 0x04, 0x02, 0x00, 0x0d] # CID
    CMD_THFULLNAME = [0x80, 0xb0, 0x00, 0x11, 0x02, 0x00, 0x64] # TH Fullname
    CMD_ENFULLNAME = [0x80, 0xb0, 0x00, 0x75, 0x02, 0x00, 0x64] # EN Fullname
    CMD_BIRTH = [0x80, 0xb0, 0x00, 0xD9, 0x02, 0x00, 0x08] # Date of birth
    CMD_GENDER = [0x80, 0xb0, 0x00, 0xE1, 0x02, 0x00, 0x01] # Gender
    CMD_ISSUER = [0x80, 0xb0, 0x00, 0xF6, 0x02, 0x00, 0x64] # Card Issuer
    CMD_ISSUE = [0x80, 0xb0, 0x01, 0x67, 0x02, 0x00, 0x08] # Issue Date
    CMD_EXPIRE = [0x80, 0xb0, 0x01, 0x6F, 0x02, 0x00, 0x08] # Expire Date
    CMD_ADDRESS = [0x80, 0xb0, 0x15, 0x79, 0x02, 0x00, 0x64] # Address 
    # Get all the available readers
    readerList = readers()
    readerSelectIndex = 0 #int(input("Select reader[0]: ") or "0")
    reader = readerList[readerSelectIndex]
    connection = reader.createConnection()
except:
    pass 

def thai2unicode(data):
    result = ''
    result = bytes(data).decode('tis-620')
    return result.strip()#strip()หมายถึงไม่เอา string ที่ไม่ต้องการ

def getData(cmd, req = [0x00, 0xc0, 0x00, 0x00]):
    data, sw1, sw2 = connection.transmit(cmd)
    data, sw1, sw2 = connection.transmit(req+ [cmd[-1]])
    return [data, sw1, sw2]

def textCard(): #ข้อมูลในบัตรทั้งหมด ยกเว้นรูป
    try:
        cardtype = AnyCardType()
        cardrequest = CardRequest( timeout=1, cardType=cardtype )
        cardservice = cardrequest.waitforcard()
        cardservice.connection.connect()

        connection.connect()
        atr = connection.getATR()
        if (atr[0] == 0x3B & atr[1] == 0x67):
            req = [0x00, 0xc0, 0x00, 0x01]
        else :
            req = [0x00, 0xc0, 0x00, 0x00]
        # Check card
        data, sw1, sw2 = connection.transmit(SELECT + THAI_CARD)
        #print ("Select Applet: %02X %02X" % (sw1, sw2))

        #ตัวเก็บข้อมูลทุกอย่างไว้ใน list
        count = []#เก็บข้อมูลตัวเลข และวันที่แปลงเป็นตัวหนังสือ 
        monthThai = ['null','มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม']
        monthEng = ['null','January','February','March','April','May','June','July','August','September','October','November','December']
        # CID
        data = getData(CMD_CID, req)
        cid = thai2unicode(data[0])
        count.append(cid)
        print ("เลขประจำตัวประชาชน: " + cid)

        # TH Fullname
        data = getData(CMD_THFULLNAME, req)
        TH = thai2unicode(data[0])
        count.append(TH)
        TH = TH.replace('#', ' ')
        #TH.replace('#', ' ')
        print ("ชื่อ-สกุล: " +  TH)
        #print(thai2unicode2(data[0])))

        # EN Fullname
        data = getData(CMD_ENFULLNAME, req)
        EN = thai2unicode(data[0])
        count.append(EN)
        EN = EN.replace('#', ' ')
        print ("Name-LastName: " + EN)

        # Date of birth thai
        data = getData(CMD_BIRTH, req)
        count.append(thai2unicode(data[0]))
        x1 = list(thai2unicode(data[0]))
        a = []
        num = x1[4]+x1[5]
        for i in  range(12):
            if i == int(num):
                a.append(monthThai[i])
                a.append(monthEng[i])
            else:
                pass
            
        DATE_B = x1[-2]+x1[-1]+'/'+a[0]+'/'+x1[0]+x1[1]+x1[2]+x1[3]
        a.append(int(x1[0]+x1[1]+x1[2]+x1[3])-543)
        count.append(DATE_B)
        print( "เกิดวันที่: " +DATE_B)
                            
        # Gender
        data = getData(CMD_GENDER, req)
        GEN = thai2unicode(data[0])
        
        if GEN == '1':
            count.append('ชาย')
            print ("เพศ: " + "ชาย")
        else:
            count.append('หญิง')
            print ("เพศ: " + "หญิง")

        # Card Issuer
        data = getData(CMD_ISSUER, req)
        Card_Is = thai2unicode(data[0])
        count.append(Card_Is)
        print ("สถานที่ออกบัตร: " + Card_Is)

        # Issue Date
        data = getData(CMD_ISSUE, req)
        count.append(thai2unicode(data[0]))
        x2 = list(thai2unicode(data[0]))
        a2 = []
        num2 = x2[4]+x2[5]
        for i in  range(12):
            if i == int(num2):
                a2.append(monthThai[i])
                a2.append(monthEng[i])
            else:
                pass
        Issue_Date = x2[-2]+x2[-1]+'/'+a2[0]+'/'+x2[0]+x2[1]+x2[2]+x2[3]
        count.append(Issue_Date)
        print ("วันออกบัตร: " + Issue_Date)

        # Expire Date
        data = getData(CMD_EXPIRE, req)
        count.append(thai2unicode(data[0]))
        x3 = list(thai2unicode(data[0]))
        a3 = []
        num3 = x3[4]+x3[5]
        for i in  range(12):
            if i == int(num3):
                a3.append(monthThai[i])
                a3.append(monthEng[i])
            else:
                pass
        Expire_Date = x3[-2]+x3[-1]+'/'+a3[0]+'/'+x3[0]+x3[1]+x3[2]+x3[3]
        count.append(Expire_Date)
        print ("วันบัตรหมดอายุ: " + Expire_Date)

        # Address
        data = getData(CMD_ADDRESS, req)
        Add = thai2unicode(data[0])
        count.append(Add)
        Add = Add.replace('#', ' ')
        print ("ที่อยู่: " + Add)

        #Time
        time = str(datetime.datetime.now())
        count.append(time)
        print("เวลาในการอ่านบัตร:"+time)

        #Date of birth ENG
        DATE_C = x1[-2]+x1[-1]+'/'+a[1]+'/'+str(a[2])
        count.append(x1[-2]+x1[-1]+'/'+a[1]+'/'+str(a[2]))
        print("Date of birth: "+DATE_C)

        #Age
        instantTime = list(time)
        countTime = (int(instantTime[0]+instantTime[1]+instantTime[2]+instantTime[3])+543) - int(x1[0]+x1[1]+x1[2]+x1[3])
        count.append(countTime)
        print('อายุ ',countTime ,'ปี')

        count[1] = count[1].replace('#', ' ')
        count[2] = count[2].replace('#', ' ')
        count[11 ] = count[11].replace('#', ' ')
        return count   
    except:
        pass 

def photoCard(fileName): #อ่านไฟล์รูป
    cardtype = AnyCardType()
    cardrequest = CardRequest( timeout=1, cardType=cardtype )
    cardservice = cardrequest.waitforcard()
    cardservice.connection.connect()

    SELECT = [0x00, 0xA4, 0x04, 0x00, 0x08]
    THAI_ID_CARD = [0xA0, 0x00, 0x00, 0x00, 0x54, 0x48, 0x00, 0x01]
    REQ_PHOTO_P1 = [0x80,0xB0,0x01,0x7B,0x02,0x00,0xFF]
    REQ_PHOTO_P2 = [0x80,0xB0,0x02,0x7A,0x02,0x00,0xFF]
    REQ_PHOTO_P3 = [0x80,0xB0,0x03,0x79,0x02,0x00,0xFF]
    REQ_PHOTO_P4 = [0x80,0xB0,0x04,0x78,0x02,0x00,0xFF]
    REQ_PHOTO_P5 = [0x80,0xB0,0x05,0x77,0x02,0x00,0xFF]
    REQ_PHOTO_P6 = [0x80,0xB0,0x06,0x76,0x02,0x00,0xFF]
    REQ_PHOTO_P7 = [0x80,0xB0,0x07,0x75,0x02,0x00,0xFF]
    REQ_PHOTO_P8 = [0x80,0xB0,0x08,0x74,0x02,0x00,0xFF]
    REQ_PHOTO_P9 = [0x80,0xB0,0x09,0x73,0x02,0x00,0xFF]
    REQ_PHOTO_P10 = [0x80,0xB0,0x0A,0x72,0x02,0x00,0xFF]
    REQ_PHOTO_P11 = [0x80,0xB0,0x0B,0x71,0x02,0x00,0xFF]
    REQ_PHOTO_P12 = [0x80,0xB0,0x0C,0x70,0x02,0x00,0xFF]
    REQ_PHOTO_P13 = [0x80,0xB0,0x0D,0x6F,0x02,0x00,0xFF]
    REQ_PHOTO_P14 = [0x80,0xB0,0x0E,0x6E,0x02,0x00,0xFF]
    REQ_PHOTO_P15 = [0x80,0xB0,0x0F,0x6D,0x02,0x00,0xFF]
    REQ_PHOTO_P16 = [0x80,0xB0,0x10,0x6C,0x02,0x00,0xFF]
    REQ_PHOTO_P17 = [0x80,0xB0,0x11,0x6B,0x02,0x00,0xFF]
    REQ_PHOTO_P18 = [0x80,0xB0,0x12,0x6A,0x02,0x00,0xFF]
    REQ_PHOTO_P19 = [0x80,0xB0,0x13,0x69,0x02,0x00,0xFF]
    REQ_PHOTO_P20 = [0x80,0xB0,0x14,0x68,0x02,0x00,0xFF]

    PHOTO = [REQ_PHOTO_P1,REQ_PHOTO_P2,REQ_PHOTO_P3,REQ_PHOTO_P4,REQ_PHOTO_P5,
    REQ_PHOTO_P6,REQ_PHOTO_P7,REQ_PHOTO_P8,REQ_PHOTO_P9,REQ_PHOTO_P10,REQ_PHOTO_P11
    ,REQ_PHOTO_P12,REQ_PHOTO_P13,REQ_PHOTO_P14,REQ_PHOTO_P15,REQ_PHOTO_P16,REQ_PHOTO_P17,
    REQ_PHOTO_P18,REQ_PHOTO_P19,REQ_PHOTO_P20]
    
    apdu = SELECT+THAI_ID_CARD
    response, sw1, sw2 = cardservice.connection.transmit( apdu )

    ### Fetch and write photo บันทึกรูปภาพ
    fphoto = open('temp/'+fileName[0]+".png", "wb")
    for d in PHOTO:
        response, sw1, sw2 = cardservice.connection.transmit( d )
        if sw1 == 0x61:
            GET_RESPONSE = [0X00, 0XC0, 0x00, 0x00 ]
            apdu = GET_RESPONSE + [sw2]
            response, sw1, sw2 = cardservice.connection.transmit( apdu )
            fphoto.write(bytearray(response))
            #BackUpPhoto.write(bytearray(response))

def resizeImg(fileName): #ปรับขนาดรูป
    for name in glob.glob('temp/'+fileName[0]+".png"):
                with open(name, 'r+b') as f: 
                    with Image.open(f) as image:
                        cover = resizeimage.resize_width(image,148)
                        cover = resizeimage.resize_cover(image,[148,165])
                        cover.save(name, image.format)













# from PIL import ImageTk,ImageDraw,ImageFont, Image as PILImage
# import time
# import datetime
# global im
# global image



# def addTime2Img(file):
#     img=PILImage.open(file)
#     font = ImageFont.truetype("cmb10.ttf", 50)
#     draw = ImageDraw.Draw(img)
#     now = int(time.time()) 
#     # time1 = time.localtime(now)
#     time1_str = datetime.datetime.fromtimestamp(now)
#     otherStyleTime = time1_str.strftime("%Y--%m--%d %H:%M:%S")
#     draw.text((200,550),otherStyleTime,(0,0,0),font)
#     draw = ImageDraw.Draw(img)
#     img.save(file)

# addTime2Img("test.png")
import hashlib
import pathlib
import configparser
import wmi
import datetime
import win32com.client

def encrypt(raw:str,key_int):
    raw_bytes:bytes = raw.encode()
    raw_int:int = int.from_bytes(raw_bytes, 'big')
    # key_int:int = random_key(len(raw_bytes))
    return raw_int ^ int.from_bytes(key_int.encode(),'big')


def decrypt(encrypted:str, key_int:str) -> str:
    decrypted:int = int(encrypted) ^ int.from_bytes(key_int.encode(),'big')
    length = (decrypted.bit_length() + 7) // 8
    decrypted_bytes:bytes = int.to_bytes(decrypted, length, 'big') 
    return decrypted_bytes.decode()

def encrypt_file(path, encoding='utf-8'):
    global serialNumber
    with path.open('rt', encoding='utf-8') as f1:
        tmp=encrypt(f1.read(),serialNumber)
        # print(tmp)
        # print(decrypt(tmp,serialNumber))
    path.open('wt', encoding='utf-8').write(str(tmp))

def decrypt_file(path, encoding='utf-8'):
    global serialNumber
    with path.open('rt', encoding='utf-8') as f1:
        tmp=decrypt(f1.read(),serialNumber)
        # print(tmp)
        return tmp
    # path.open('wt', encoding='utf-8').write(tmp)

def getHardDiskNumber():
    c = wmi.WMI()
    for physical_disk in c.Win32_DiskDrive():
        print(physical_disk.SerialNumber)
        return physical_disk.SerialNumber
    

serialNumber=getHardDiskNumber()

curpath=pathlib.Path(__file__)

# import os

def open_excel(path,open_password,write_password):
    excel_path="加密测试.xlsx"
    picture_path="screenshot.png"
    curpath=pathlib.Path(__file__)
    root_path=curpath.parent.joinpath(picture_path)
    xlApp=win32com.client.Dispatch("Excel.Application")
    xlApp.Visible=False
    xlApp.DisplayAlerts=False
    xlwb=xlApp.Workbooks.Open(Filename=curpath.parent.joinpath(excel_path),UpdateLinks=0,
    ReadOnly=False,Format=None,Password=open_password,WriteResPassword=write_password)
    #获取某个Sheet页数据（页数从1开始）
    wtire_Test=xlwb.Worksheets(1)
    wtire_Test.Cells(2,8).Value=""
    picture_left = wtire_Test.Cells(1, 4).Left
    picture_top = wtire_Test.Cells(1, 4).Top
    picture_width = wtire_Test.Cells(1, 4).Width
    picture_height = wtire_Test.Cells(1, 4).Height
    wtire_Test.Shapes.AddPicture(root_path,1,4,picture_left,picture_top,picture_width,picture_height)
    sheet_data=xlwb.Worksheets(1).UsedRange.Value
    # for item in sheet_data:
    #     print(item[0])
    print(len(sheet_data))
    print(sheet_data[0][1]==None)
    # xlwb.SaveAs(os.path.abspath('.')+"\\123.xlsx")
    xlwb.Save()
    xlwb.Close(True)
    
    
    xlApp.Quit()

def getMd5(file):
    m = hashlib.md5()
    # file='加密测试.xlsx'
    with file.open('rb') as f:
        for line in f:            
            m.update(line)    
    md5code = m.hexdigest()    
    return md5code

# curpath = os.path.dirname(os.path.realpath(__file__))
curpath=pathlib.Path(__file__)
# print(curpath.cwd())
open_excel("","123","123")
curpath=pathlib.Path("加密测试.xlsx")
dest=pathlib.Path("test.xlsx")
# print(dest.open("rb").read())
with dest.open("wb") as f1:
    f1.write(curpath.open("rb").read())
# cfgpath = curpath.parent.joinpath("FacilityTest.ini")
# # ini_path=pathlib.Path("FacilityTest.ini")
# print(cfgpath)
# conf = configparser.ConfigParser()
# conf.read_string(decrypt_file(cfgpath))
# # # conf.add_section("20191210")
# sections = conf.sections()

# date = datetime.datetime.now()

# detester = date.strftime('%Y%m%d')
# try:
#     conf.add_section(detester)
# except:
#     pass
# conf[detester]["md5"]="1234567"
# sections = conf.sections()
# # print(conf)
# conf.write(cfgpath.open('wt'))
# encrypt_file(cfgpath)
# decrypt_file(cfgpath)
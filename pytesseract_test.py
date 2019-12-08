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
import wmi
def getHardDiskNumber():
    c = wmi.WMI()
    for physical_disk in c.Win32_DiskDrive():
        print(physical_disk.SerialNumber)

getHardDiskNumber()
import pandas as pd
from selenium.webdriver import Firefox
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from PIL import ImageTk, ImageDraw, ImageFont, Image as PILImage
from tkinter import Tk, Canvas, mainloop
import pathlib
import time
import datetime
import getpass
import hashlib
# import pathlib
import configparser
import wmi
# import datetime
import win32com.client

driver = None
wait = None
width = 0
height = 0
tmp_screenshot = "tmp/screenshot.png"
tmp_captcha = "tmp/captcha.png"


class search4facility:
    # def addTime2Img(file):
    #     img = PILImage.open(file)
    #     font = ImageFont.truetype("cmb10.ttf", 50)
    #     draw = ImageDraw.Draw(img)
    #     now = int(time.time())
    #     # time1 = time.localtime(now)
    #     time1_str = datetime.datetime.fromtimestamp(now)
    #     otherStyleTime = time1_str.strftime("%Y--%m--%d %H:%M:%S")
    #     draw.text((200, 550), otherStyleTime, (0, 0, 0), font)
    #     draw = ImageDraw.Draw(img)
    #     img.save(file)
    date=None
    securityControl=None
    # conf=None
    def showCaptcha(self):
        global driver
        global tmp_captcha
        global tmp_screenshot

        baidu_img = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, 'captchaImg')))
        input_yzm = wait.until(EC.presence_of_element_located((By.ID, "yzm")))
        # input_yzm.send_keys("test")
        input_yzm.clear()
        for i in range(0, 3):
            baidu_img.click()
            driver.implicitly_wait(10)
        time.sleep(0.5)
        baidu_img = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, 'captchaImg')))
        driver.save_screenshot(tmp_screenshot)  # 对整个浏览器页面进行截图
        left = baidu_img.location['x']
        top = baidu_img.location['y']
        right = baidu_img.location['x'] + baidu_img.size['width']
        bottom = baidu_img.location['y'] + baidu_img.size['height']
        im = PILImage.open(tmp_screenshot)
        im = im.crop((left, top, right, bottom))
        im.save(tmp_captcha)

        print("显示验证码图片......")
        image = PILImage.open(tmp_captcha)
        # image.show()
        master = Tk()
        canvas = Canvas(master, width=200, height=34, bg='black')
        canvas.pack()
        im = ImageTk.PhotoImage(image)
        # im.show()
        canvas.create_image(100, 17, image=im)
        # lambda x=ALL: canvas.delete(x)
        # canvas.delete(tmp)
        mainloop()

    def brower_start(self):
        global driver
        global wait
        global width
        global height
        print("正在初始化...")
        if driver is not None:
            driver.close()
        else:
            options = Options()
            options.add_argument('-headless')  # 无头参数
            driver = Firefox(firefox_options=options)
            driver.set_page_load_timeout(12)
            wait = WebDriverWait(driver, timeout=10)
        while True:
            try:
                print("正在载入页面...")
                driver.get('http://zxgk.court.gov.cn/shixin/')
                print("页面载入成功")
                break
            except:
                print("页面载入错误，尝试重载")
                pass
        eles = driver.find_elements_by_xpath("//html")
        # locs = []
        if width == 0 and height == 0 and len(eles) > 0:
            width = int(eles[0].size['width'])
            height = int(eles[0].size['height'])
        driver.set_window_size(width, height)

    def page_restart(self):
        global driver
        global wait
        global width
        global height
        while True:
            try:
                print("正在载入页面...")
                driver.get('http://zxgk.court.gov.cn/shixin/')
                print("页面载入成功")
                break
            except:
                print("页面载入错误，尝试重载")
                pass
        eles = driver.find_elements_by_xpath("//html")
        # locs = []
        if width == 0 and height == 0 and len(eles) > 0:
            width = int(eles[0].size['width'])
            height = int(eles[0].size['height'])
        driver.set_window_size(width, height)

    def search(self, pName, pCardNum):
        global driver
        global wait
        global width
        global height
        global tmp_captcha
        global tmp_screenshot

        if str(pCardNum) == "nan":
            pCardNum = ""
        else:
            pCardNum = str(pCardNum).split(".")[0]
        flag = 0  # 0:无记录，1:有记录，2：查询失败重新查询
        driver.set_window_size(width, height)

        print("当前正在查询:"+pName)
        if not(pathlib.Path("tmp").exists()):
            pathlib.Path("tmp").mkdir()

        if not(pathlib.Path("result").exists()):
            pathlib.Path("result").mkdir()

        input_tag = wait.until(
            EC.presence_of_element_located((By.ID, "pName")))
        input_id = wait.until(
            EC.presence_of_element_located((By.ID, "pCardNum")))
        searchButton = driver.find_element_by_css_selector(
            "div#yzm-group button")
        input_yzm = wait.until(EC.presence_of_element_located((By.ID, "yzm")))
        input_tag.clear()
        input_tag.send_keys(pName)
        input_id.clear()
        input_id.send_keys(pCardNum)

        while(1):
            print("请键入以下指令或直接键入四位验证码:")
            print("1————刷新验证码")
            print("2————刷新页面")
            print("3————退出")
            # input_yzm.clear()
            self.showCaptcha()
            driver.implicitly_wait(10)
            yzm = input()
            if yzm == "":
                continue
            elif yzm == "1":
                try:
                    input_yzm.clear()
                    # showCaptcha()
                    continue
                except Exception as e:
                    print(e)
                    print("页面元素获取错误，请等待页面重载")
                    self.page_restart()
                    input_tag = wait.until(
                        EC.presence_of_element_located((By.ID, "pName")))
                    input_id = wait.until(
                        EC.presence_of_element_located((By.ID, "pCardNum")))
                    searchButton = driver.find_element_by_css_selector(
                        "div#yzm-group button")
                    input_yzm = wait.until(
                        EC.presence_of_element_located((By.ID, "yzm")))
                    input_tag.clear()
                    input_tag.send_keys(pName)
                    input_id.clear()
                    input_id.send_keys(pCardNum)
                    continue
                    # showCaptcha()

            elif yzm == "2":
                self.page_restart()
                input_tag = wait.until(
                    EC.presence_of_element_located((By.ID, "pName")))
                input_id = wait.until(
                    EC.presence_of_element_located((By.ID, "pCardNum")))
                searchButton = driver.find_element_by_css_selector(
                    "div#yzm-group button")
                input_yzm = wait.until(
                    EC.presence_of_element_located((By.ID, "yzm")))
                input_tag.clear()
                input_tag.send_keys(pName)
                input_id.clear()
                input_id.send_keys(pCardNum)
                # showCaptcha()
            elif yzm == "3":
                print("正在退出，请稍候")
                driver.quit()
                exit()
            else:
                input_yzm.send_keys(yzm)
                try:
                    searchButton.click()
                    driver.implicitly_wait(10)
                    # searchButton.click()
                    # driver.implicitly_wait(10)
                    time.sleep(1)
                    driver.save_screenshot("screenshot.png")
                    # try:
                    #     _ = driver.find_element_by_css_selector(
                    #         "div[type='dialog']")
                    #     print("身份证号/组织结构代码不合法！请检查excel文档")
                    #     flag = 3
                    #     break
                    # except Exception as e:
                    #     print(e)
                    #     pass
                    try:
                        # print(driver.find_element_by_css_selector("tbody#tbody-result span[text='验证码错误或验证码已过期。']").text)
                        # print(driver.find_elements_by_xpath(".//span[@class='important']").text)
                        if(driver.find_element_by_css_selector("div#result-block span[class='important']").text != "全国"):
                            print("验证码错误")
                            # input_yzm.clear()
                            # showCaptcha()
                            continue
                    except Exception as e:
                        print(e)
                        flag = 1
                        pass
                    driver.implicitly_wait(10)
                    if(flag == 0):
                        driver.save_screenshot("result/无记录_"+pName+".png")
                        addTime2Img("result/无记录_"+pName+".png")
                    elif flag == 1:
                        driver.save_screenshot("result/有记录_"+pName+".png")
                        addTime2Img("result/有记录_"+pName+".png")
                    print("查询成功，已保存")
                    # driver.quit()
                    break
                except Exception as e:
                    print(e)
                    print("查询按钮获取错误，请等待页面重载")
                    flag = 2
                    page_restart()
                    break
        return flag
        # showCaptcha(tmp_captcha)

        # driver.save_screenshot("screenshot2.png")
        # searchButton = driver.find_element_by_css_selector("div#yzm-group button")
    
    def getMd5(self, file):
        m = hashlib.md5()
        # file='加密测试.xlsx'
        with file.open('rb') as f:
            for line in f:
                m.update(line)
        md5code = m.hexdigest()
        return md5code
    
    def md5_check(self):
        conf=configparser.ConfigParser()
        curpath = pathlib.Path(__file__)
        cfgpath = curpath.parent.joinpath("FacilityTest.ini")
        conf.read(cfgpath)
        try:
            conf.add_section(self.date)
            backup=pathlib.Path("backup.xlsx")
            destination=pathlib.Path(self.date+".xlsx")
            with destination.open("wb") as f1:
                f1.write(backup.open("rb").read())
            conf[self.date]["md5"]=self.getMd5(destination)
            conf.write(cfgpath.open("wt"))
            # self.securityControl.encrypt_file(cfgpath)
            return True
        except:
            pass
        conf.read(cfgpath)
        if conf[self.date]["md5"]==self.getMd5(pathlib.Path(self.date+".xlsx")):
            return True
        else:
            return False
        # conf=configparser.ConfigParser()
        # conf=self.securityControl.decrypt_file(cfgpath)
        
        


    def store_result(self,path,open_password,write_password,pName,pCardNum,result,pic_path):
        if not self.md5_check():
            print("xlsx校验失败")
            return 0
        else:
            pass
        xlApp=win32com.client.Dispatch("Excel.Application")
        xlApp.Visible=False
        xlApp.DisplayAlerts=False
        xlwb=xlApp.Workbooks.Open(Filename=path,UpdateLinks=0,
        ReadOnly=False,Format=None,Password=open_password,WriteResPassword=write_password) 
        sheet=xlwb.Worksheets(1)
        sheet_data=sheet.Value
        index=0
        for i in range(0,len(sheet_data)):
            if sheet_data[i][0]=="None":
                index=i+1
                break
        sheet.Cells(index,1).Value=pName
        sheet.Cells(index,2).Value=pCardNum
        sheet.Cells(index,3).Value=result
        picture_left = sheet.Cells(index, 4).Left
        picture_top = sheet.Cells(index, 4).Top
        picture_width = sheet.Cells(index, 4).Width
        picture_height = sheet.Cells(index, 4).Height
        sheet.Shapes.AddPicture(pic_path,index,4,picture_left,picture_top,picture_width,picture_height)
        xlwb.Save()
        xlwb.Close(True)
        
        
    def __init__(self,securityControl):
        self.brower_start()
        self.securityControl=securityControl
        self.date = datetime.datetime.now()

        self.date = self.date.strftime('%Y%m%d')
        data = pd.read_excel('投标人失信查询.xlsx', sheet_name='Sheet1',
                             dtype={"投标人证件号": object})
        result_col = data["失信被执行人查询结果"].tolist()
        count = len(result_col)
        index = -1
        for i, j in zip(data['投标人名称'], data["投标人证件号"]):
            index += 1
            print(i)
            if(result_col[index] == "有记录" or result_col[index] == "无记录" or result_col[index] == "号码错误"):
                print("已查询,跳过")
            else:
                while(1):
                    print("正在查询第"+str(index+1)+"个,共"+str(count)+"个")
                    flag = self.search(i, j)
                    if(flag == 0):
                        result_col[index] = "无记录"
                        break
                    elif flag == 1:
                        result_col[index] = "有记录"
                        break
                    elif flag == 2:
                        continue
                    elif flag == 3:
                        result_col[index] = "号码错误"
                        self.page_restart()
                        break
                data["失信被执行人查询结果"] = result_col
                data.to_excel('投标人失信查询.xlsx', sheet_name='Sheet1')

        if driver is not None:
            driver.close()


class SecurityControl:

    userFlag = None
    conf=None
    serialNumber=None

    def encrypt(self, raw: str, key_int):
        raw_bytes: bytes = raw.encode()
        raw_int: int = int.from_bytes(raw_bytes, 'big')
        # key_int:int = random_key(len(raw_bytes))
        return raw_int ^ int.from_bytes(key_int.encode(), 'big')

    def decrypt(self, encrypted: str, key_int: str) -> str:
        decrypted: int = int(encrypted) ^ int.from_bytes(
            key_int.encode(), 'big')
        length = (decrypted.bit_length() + 7) // 8
        decrypted_bytes: bytes = int.to_bytes(decrypted, length, 'big')
        return decrypted_bytes.decode()

    def encrypt_file(self, path, encoding='utf-8'):
        # global serialNumber
        with path.open('rt', encoding='utf-8') as f1:
            tmp = self.encrypt(f1.read(), serialNumber)
            # print(tmp)
            # print(decrypt(tmp,serialNumber))
        path.open('wt', encoding='utf-8').write(str(tmp))

    def decrypt_file(self, path, encoding='utf-8'):
        # global serialNumber
        with path.open('rt', encoding='utf-8') as f1:
            tmp = decrypt(f1.read(), serialNumber)
            # print(tmp)
            return tmp
        # path.open('wt', encoding='utf-8').write(tmp)

    def getHardDiskNumber(self):
        c = wmi.WMI()
        for physical_disk in c.Win32_DiskDrive():
            print(physical_disk.SerialNumber)
            return physical_disk.SerialNumber

    # serialNumber=getHardDiskNumber()

    # curpath=pathlib.Path(__file__)

    # import os

    def open_excel(self, path, open_password, write_password):
        excel_path = "加密测试.xlsx"
        picture_path = "screenshot.png"
        root_path = curpath.parent.joinpath(picture_path)
        xlApp = win32com.client.Dispatch("Excel.Application")
        xlApp.Visible = False
        xlApp.DisplayAlerts = False
        xlwb = xlApp.Workbooks.Open(Filename=curpath.parent.joinpath(excel_path), UpdateLinks=0,
                                    ReadOnly=False, Format=None, Password=open_password, WriteResPassword=write_password)
        # 获取某个Sheet页数据（页数从1开始）
        wtire_Test = xlwb.Worksheets(1)
        wtire_Test.Cells(2, 5).Value = ""
        picture_left = wtire_Test.Cells(1, 4).Left
        picture_top = wtire_Test.Cells(1, 4).Top
        picture_width = wtire_Test.Cells(1, 4).Width
        picture_height = wtire_Test.Cells(1, 4).Height
        wtire_Test.Shapes.AddPicture(
            root_path, 1, 4, picture_left, picture_top, picture_width, picture_height)
        sheet_data = xlwb.Worksheets(1).UsedRange.Value
        print(sheet_data)
        # xlwb.SaveAs(os.path.abspath('.')+"\\123.xlsx")
        xlwb.Save()
        xlwb.Close(True)

        xlApp.Quit()

    def getMd5(self, file):
        m = hashlib.md5()
        # file='加密测试.xlsx'
        with file.open('rb') as f:
            for line in f:
                m.update(line)
        md5code = m.hexdigest()
        return md5code

    def setpwd(self):
        while(1):
            pwd1 = getpass.getpass("请输入密码(8位纯数字)")
            if not(pwd1.isdecimal() and len(pwd1) == 8):
                print("输入格式错误")
                continue
            else:
                pwd2 = getpass.getpass("请再次输入密码")
                if pwd1 == pwd2:
                    return pwd1
                else:
                    print("两次输入不一致")
                    continue

    def user_control(self):
        pwd=getpass.getpass("请输入你的密码：")
        if pwd==self.conf["System"]["admin"]:
            self.userFlag="admin"
        elif pwd==self.conf["System"]["user1"]:
            self.userFlag="user1"
        elif pwd==self.conf["System"]["user2"]:
            self.userFlag="user2"
        else:
            print("密码错误，回到上级菜单")
            return 0

        while(1):
                print("请键入以下指令或直接键入四位验证码:")
                print("1————开始查询")
                print("2————结果查看")
                print("3————退出")
                tag = input()
                if(tag == "1"):
                    search4facility()
                elif(tag == "2"):
                    self.open_excel()
                elif(tag == "3"):
                    exit()
        # return 1

    def __init__(self):
        curpath = pathlib.Path(__file__)
        cfgpath = curpath.parent.joinpath("FacilityTest.ini")
        # 判断程序完整性
        if not(cfgpath.exists()):
            print("配置文件不存在，请重新配置程序！")
            exit()
        self.conf = configparser.ConfigParser()
        try:
            # conf.read_string(self.decrypt_file(cfgpath))
            self.conf.read(cfgpath)
        except Exception as e:
            print(e)
            print("配置文件加密错误，可能被修改，请重新配置程序！")
            exit()

        if(self.conf["System"]["firstrun"] == "1"):
            self.conf["System"]["serialnumber"] = self.getHardDiskNumber()
            self.conf["System"]["admin"] = "06218767"
            # print("请输入8位数字密码：")
            print("设置user1的密码")
            self.userFlag = "user1"
            self.conf["System"][self.userFlag] = self.setpwd()
            print("设置user2的密码")
            self.userFlag = "user2"
            self.conf["System"][self.userFlag] = self.setpwd()
            self.conf["System"]["firstrun"] == "0"
            self.conf.write(cfgpath.open("wt"))
            # self.encrypt(cfgpath)

        if(self.conf["System"]["firstrun"] == "0"):

            while(1):
                print("请键入以下指令或直接键入四位验证码:")
                print("1————开始查询")
                print("2————用户登录")
                print("3————退出")
                tag = input()
                if(tag == "1"):
                    search4facility(self)
                elif(tag == "2"):
                    self.user_control()
                elif(tag == "3"):
                    exit()



# curpath=pathlib.Path(__file__)
# # print(curpath.cwd())
# cfgpath = curpath.parent.joinpath("FacilityTest.ini")
# conf = configparser.ConfigParser()
# conf.add_section("System")
# conf["System"]["FirstRun"]="1"
# conf.write(cfgpath.open("wt"))
test = SecurityControl()

import pandas as pd
from selenium.webdriver import Firefox
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from PIL import ImageTk, Image as PILImage
from tkinter import Tk, Canvas, mainloop
import os
import time
import datetime

driver = None
wait = None
width = 0
height = 0
tmp_screenshot = "tmp/screenshot.png"
tmp_captcha = "tmp/captcha.png"


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


def showCaptcha():
    global driver
    global tmp_captcha
    global tmp_screenshot

    baidu_img = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, 'captchaImg')))
    input_yzm = wait.until(EC.presence_of_element_located((By.ID, "yzm")))
    # input_yzm.send_keys("test")
    input_yzm.clear()
    for i in range(0,3):
        baidu_img.click()
        driver.implicitly_wait(10)
    time.sleep(1)
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


def brower_start():
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


def page_restart():
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


def search(pName, pCardNum):
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
    if not(os.path.isdir("tmp")):
        os.mkdir("tmp")

    if not(os.path.isdir("result")):
        os.mkdir("result")

    input_tag = wait.until(EC.presence_of_element_located((By.ID, "pName")))
    input_id = wait.until(EC.presence_of_element_located((By.ID, "pCardNum")))
    searchButton = driver.find_element_by_css_selector("div#yzm-group button")
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
        showCaptcha()
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
                page_restart()
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
            page_restart()
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
                    # addTime2Img("result/无记录_"+pName+".png")
                elif flag == 1:
                    driver.save_screenshot("result/有记录_"+pName+".png")
                    # addTime2Img("result/有记录_"+pName+".png")
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


brower_start()
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
            flag = search(i, j)
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
                page_restart()
                break
        data["失信被执行人查询结果"] = result_col
        data.to_excel('投标人失信查询.xlsx', sheet_name='Sheet1')

if driver is not None:
    driver.close()

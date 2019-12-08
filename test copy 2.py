import pandas as pd
import requests
import json
import urllib.parse
import urllib.request as urllib2
import http.cookiejar as cookielib

from selenium.webdriver import Firefox
from selenium.webdriver.common.by import By 
from selenium.webdriver.common.keys import Keys 
from selenium.webdriver.firefox.options import Options 
from selenium.webdriver.support import expected_conditions as EC 
from selenium.webdriver.support.wait import WebDriverWait 
from PIL import Image,ImageDraw
import pytesseract

def ChangeColor(img):
    w,h=img.size
    for x in range(w):
        for y in range(h):
            # print(img.getpixel((x,y)))
            r,g,b, _=img.getpixel((x,y))
            if 190<=r<=255 and 170<=g<=255 and 0<=b<=140:
                img.putpixel((x,y),(0,0,0))
            if 0<=r<=90 and 210<=g<=255 and 0<=b<=90:
                img.putpixel((x,y),(0,0,0))
    img=img.convert('L').point([0]*150+[1]*(256-150),'1')
    # img.show()
    return img

t2val = {}


def twoValue(image, G):
    for y in range(0, image.size[1]):
        for x in range(0, image.size[0]):
            g = image.getpixel((x, y))
            if g > G:
                t2val[(x, y)] = 1
            else:
                t2val[(x, y)] = 0

def clearNoise(image, N, Z):
    for i in range(0, Z):
        t2val[(0, 0)] = 1
        t2val[(image.size[0] - 1, image.size[1] - 1)] = 1

        for x in range(1, image.size[0] - 1):
            for y in range(1, image.size[1] - 1):
                nearDots = 0
                L = t2val[(x, y)]
                if L == t2val[(x - 1, y - 1)]:
                    nearDots += 1
                if L == t2val[(x - 1, y)]:
                    nearDots += 1
                if L == t2val[(x - 1, y + 1)]:
                    nearDots += 1
                if L == t2val[(x, y - 1)]:
                    nearDots += 1
                if L == t2val[(x, y + 1)]:
                    nearDots += 1
                if L == t2val[(x + 1, y - 1)]:
                    nearDots += 1
                if L == t2val[(x + 1, y)]:
                    nearDots += 1
                if L == t2val[(x + 1, y + 1)]:
                    nearDots += 1

                if nearDots < N:
                    t2val[(x, y)] = 1
def saveImage(filename, size):
    image = Image.new("1", size)
    draw = ImageDraw.Draw(image)

    for x in range(0, size[0]):
        for y in range(0, size[1]):
            draw.point((x, y), t2val[(x, y)])

    image.save(filename)

def recognize_captcha(img):
    num = pytesseract.image_to_string(img)
    return num

def test(file):
    img=ChangeColor(file)
    
    twoValue(img.convert("L"),50)
    clearNoise(img.convert("L"), 1, 1)
    saveImage("test_"+file, img.size)
    d=recognize_captcha(file)
    print(d)


options = Options() 
options.add_argument('-headless') # 无头参数 
driver = Firefox(firefox_options=options)


wait = WebDriverWait(driver, timeout=10) 
CaptchaUrl="http://zxgk.court.gov.cn/shixin/"
driver.get(CaptchaUrl) 

for i in range(0,400):
    eles = driver.find_elements_by_xpath("//html")
    # locs = []
    width = 1920
    height = 1080
    if len(eles) > 0:
        width = int(eles[0].size['width'])
        height = int(eles[0].size['height'])
    driver.set_window_size(width, height)
    # driver.maximize_window()

    # input_tag=wait.until(EC.presence_of_element_located((By.ID,"pName")))
    # input_tag.send_keys('好日子')

    if (i%78)==0 :
        driver.get(CaptchaUrl) 
    baidu_img = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'captchaImg')))
    baidu_img.click()
    driver.save_screenshot("captchas/screenshot.png")  # 对整个浏览器页面进行截图
    left = baidu_img.location['x']
    top = baidu_img.location['y']
    right = baidu_img.location['x'] + baidu_img.size['width']
    bottom = baidu_img.location['y'] + baidu_img.size['height'] 
    im = Image.open("captchas/screenshot.png")
    im = im.crop((left, top, right, bottom)) 
    im=ChangeColor(im)
    im.save("captchas/captcha_"+str(i)+".png")
driver.quit() 

# test('captcha.png')

CaptchaUrl="http://zxgk.court.gov.cn/shixin/captchaNew.do?captchaId=aa08d8b2f6e44a2f9efaf3516a350ed9&random=0.5276334818552565"

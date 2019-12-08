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

def ChangeColor(path):
    img=Image.open(path)
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
driver.get('http://zxgk.court.gov.cn/shixin/') 
eles = driver.find_elements_by_xpath("//html")
# locs = []
width = 1920
height = 1080
if len(eles) > 0:
    width = int(eles[0].size['width'])
    height = int(eles[0].size['height'])
driver.set_window_size(width, height)
# driver.maximize_window()

input_tag=wait.until(EC.presence_of_element_located((By.ID,"pName")))
input_tag.send_keys('好日子')


baidu_img = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'captchaImg')))
driver.save_screenshot("screenshot.png")  # 对整个浏览器页面进行截图
left = baidu_img.location['x']
top = baidu_img.location['y']
right = baidu_img.location['x'] + baidu_img.size['width']
bottom = baidu_img.location['y'] + baidu_img.size['height'] 
im = Image.open('screenshot.bmp')
im = im.crop((left, top, right, bottom)) 
im.save('captcha.png')

test('captcha.png')

input_yzm=wait.until(EC.presence_of_element_located((By.ID,"yzm")))
input_yzm.send_keys('abcd')
driver.save_screenshot("test.png")
driver.quit() 



data = pd.read_excel('投标人失信查询.xlsx', sheet_name='Sheet1')
for i in data['投标人名称']:
    print(i)
    pName=urllib.parse.quote(i)


cookie = cookielib.CookieJar()
handler = urllib2.HTTPCookieProcessor(cookie)
opener = urllib2.build_opener(handler)

CaptchaUrl="http://zxgk.court.gov.cn/shixin/captchaNew.do?captchaId=aa08d8b2f6e44a2f9efaf3516a350ed9&random=0.5276334818552565"
picture = opener.open(CaptchaUrl).read()
# 用openr访问验证码地址,获取cookie
local = open('e:/image.jpg', 'wb')
local.write(picture)
local.close()

codeurl = 'http://zxgk.court.gov.cn/shixin/static/img/headbg.png'
valcode = requests.get(codeurl)
codeurl_val="http://zxgk.court.gov.cn/shixin/captchaNew.do?captchaId=aa08d8b2f6e44a2f9efaf3516a350ed9&random=0.5276334818552565"
header={"Cookies":'JSESSIONID=9117A14D4FB9B50DEF876BCEBAFB6879; _gscu_15322769=740452404ewyf219'}
valcode1 = requests.get(codeurl_val,headers=header)

f = open('valcode.png', 'wb')
# 将response的二进制内容写入到文件中
f.write(valcode.content)
# 关闭文件流对象
f.close()
# conn = requests.session()

# postdata = {
#     'username': '******',
#     'password': '******'
# }

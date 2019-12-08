import os

file_list=os.listdir("./")
res=[]
ip_list=[]
for tmp_file in file_list:
    file_post = str(tmp_file.split('.')[-1])
    if "txt" in file_post:
        print(tmp_file)
        with open('./1.txt', 'r') as f:
            res.extend(f.readlines())

for i in res:
    if i.startswith("SRC IP"):
        ip_list.append(i.split(":")[-1].replace(" ", "").replace("\n", ""))
print(ip_list)
f=open("./res.txt",'w')
f.write(str(ip_list))
f.close()
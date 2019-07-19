#test environment win10_x64 python_3.6.6

import random
import os

import random,string
src = string.ascii_letters + string.digits
# count = input('请确认要生成几条密码： ')
list_passwd_all = random.sample(src, 5) #从字母和数字中随机取5位
list_passwd_all.extend(random.sample(string.digits, 1))  #让密码中一定包含数字
list_passwd_all.extend(random.sample(string.ascii_lowercase, 1)) #让密码中一定包含小写字母
list_passwd_all.extend(random.sample(string.ascii_uppercase, 1)) #让密码中一定包含大写字母
random.shuffle(list_passwd_all) #打乱列表顺序
str_passwd = ''.join(list_passwd_all) #将列表转化为字符串
print(str_passwd)
os.system("pause")

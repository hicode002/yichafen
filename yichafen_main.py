from splinter.browser import Browser #need install splinter
import re
import xlrd
import time
shee=xlrd.open_workbook("D:\workdir\pybug\\test.xls")#the students' information table here


names={}
nums={}
cla={}
t=0
for i in range(0,22):
    table=shee.sheets()[i]
    nrows=table.nrows
    for j in range(1,nrows):
        names[t]=table.cell_value(j,3)
        nums[t]=table.cell_value(j,1)
        cla[t]=str(100+i+1)# class numbers
        t=t+1


b= Browser('edge')#need edge browser and edge webdrivers
url=""#Your yichafen url here(the query page ,not the content page)
s=""

for i in range(0,t):
    b.visit(url)

    b.fill("s_shenfenzhenghao",nums[i])#maybe others,need to grab from the website

    b.fill("s_xingming",names[i])# like the former comment
    b.find_by_id("yiDunSubmitBtn").click()
    if b.is_element_present_by_id("result_content"):
        c=b.find_by_id("result_content")
        ret=re.match ("(.*)[\n](.*)[\n](.*)",c.text) # the rules should match the query result,need to grab from the website
        ret2=ret.group(3)
       # print(ret2.split(' ')[1],' ',ret2.split(' ')[3])''
        s=s+ret2.split(' ')[1]+' '+ret2.split(' ')[3]+' '+cla[i]
        s=s+'\n'


    
    time.sleep(0.1)

with open(".\\senior1.txt","w") as f:#the score details saved address,can be personalized
    
    f.write(s)


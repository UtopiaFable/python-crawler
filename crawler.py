import re
import requests as rq
import random
import json
from pdf import *
#输出文件
FILE='output.xlsx'
if not os.path.exists(FILE):
    table = Workbook()
    sheet = table['Sheet']
    sheet.append(['报告时间','公司名称','项目名称','本年度','上年度'])
    table.save(FILE)
#所需url
index="http://www.szse.cn/api/disc/announcement/annList"
report="http://www.szse.cn/api/disc/announcement/bulletin_detail/"
src="http://disc.static.szse.cn"
head={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36",
      "Content-Type":"application/json"}
#开始遍历
company=""
id=""
time=""
need=3000
for i in range(1000):
    print("第%d页:"%(i+1))
    data={"channelCode":["fixed_disc"],"pageSize":30,"pageNum":i+1,"stock":[]}
    reports=rq.post(index,data=json.dumps(data),headers=head)
    for j in range(30):
        if need==0:
            continue
        info=reports.json()['data'][j]
        if company==info['secName'][0]:
            continue
        if info['title'].find("摘要")!=-1:
            continue

        company=info['secName'][0]
        time=info['publishTime'][:10]
        id=info['id']

        pdf_path=src+rq.get(report+id,headers=head).json()['attachPath']
        pdf_string=rq.get(pdf_path)
        pdf_file=open('tmp.pdf','wb')
        pdf_file.write(pdf_string.content)
        pdf_file.close()

        print("\n正在检索：%s"%(company))
        pdf=pdfplumber.open('tmp.pdf')
        prep=check(pdf)
        if prep==False:
            print("不可支持格式")
            continue
        idx=search(pdf,prep[0],prep[1])
        if idx is None:
            print("未能找到项目")
            continue
        table=extract(pdf,idx[0],idx[1])
        if len(table)==0:
            print("无法提取表格")
            continue
        save(table,time,company)
        print("在第%d页找到第%d个目标"%(idx[0]+1,3001-need))
        need-=1
    if need==0:
        break

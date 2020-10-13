import os.path
import pdfplumber
import openpyxl as opx
from openpyxl import Workbook
import re
import time
import numpy as np
def check(pdf):#检查年报是否为最方便的格式
    hint=pdf.pages[1].extract_text()#重要提示
    index=pdf.pages[2].extract_text()#目录
    rows=hint.split('\n')
    obj=index.split('\n')
    if index.find("....")==-1:#第2页为目录
        return False
    try:#最末一行为页码
        int(rows[-1])
    except:
        return False
    if rows[0]!=obj[0]:#最首一行为页眉
        return False
    num=re.compile('[0-9]+')
    for i in range(2,len(obj)):
        if obj[i].find("财务报告")!=-1:#找到财务报告项
            try:
                page1=int(re.search(num,obj[i],0).group(0))
                page2=min(int(re.search(num,obj[i+1],0).group(0)),len(pdf.pages)-1)
            except:
                return False
            return (page1,page2-1)#返回财务报告页区间
    return False
def search(pdf,page1,page2):
    target=re.compile("(（|\()([0-9a-z]+)(）|\)) ?支付(的)?其他与筹资活动有关的现金")#目标项
    for i in range(page2,page1,-1):#目标项一般靠后
        text=pdf.pages[i].extract_text()
        if text.find("与筹")==-1:#查找关键字
            continue
        rows=text.split('\n')
        for j,row in enumerate(rows):
            if re.match(target,row,0) is not None:#匹配成功返回页数和行数
                return i,j
    return None
def extract(pdf,page,row):
    cash=re.compile("[0-9,]+\.[0-9][0-9]")
    table=[]
    text=pdf.pages[page].extract_text()
    rows=text.split('\n')[1:-1]
    extra_rows=pdf.pages[page+1].extract_text().split('\n')[1:-1]
    rows+=extra_rows
    if rows[row].find("单位")!=-1:#排除“单位：元”行
        row+=1
    dim=len(rows[row+1].split(' '))
    if dim!=6:
        return table
    for j in range(row+1,len(rows)):
        words=rows[j].split(' ')#切分
        if len(words)!=dim or words[0]=="合计" or re.match(cash,words[0],0) is not None:#当出现“合计”、空栏、“合计”被划开时结束
            break
        table+=[np.array(words)]
        if words[0]=="无":#特例
            break
    return np.array(table)
def save(table,time,company):
    FILE='output.xlsx'
    n=len(table)
    for i in range(len(table)):
        for j in range(0,6,2):
            if table[i][j]=='':
                table[i][j]='-'
    table=table[:,[0,2,4]]
    extra=np.array([np.array([time,company]) for i in range(n)])
    table=np.column_stack((extra,table))
    data = opx.load_workbook(FILE)
    sheet=data['Sheet']
    for i in table:
        sheet.append(i.tolist())
    data.save(FILE)
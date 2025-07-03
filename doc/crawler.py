import os
import docx
import time
import requests as rq
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook

def process_doc(doc_name):
    assert doc_name.endswith(('.doc', '.docx'))
    doc = docx.Document(f'../files/{doc_name}')
    texts = []
    for paragraph in doc.paragraphs:
        texts.append(paragraph.text)
    feedback = []
    j = 0
    for i in range(len(texts)):
        if texts[i].endswith('如下反馈意见：'):
            j = i
        elif texts[i].startswith('请主承销商对上述事项进行核查并发表明确核查意见'):
            feedback.append(''.join(texts[j+1:i+1]))
            j = i
    return feedback

def feedback_to_excel(bond_info):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

def crawl_szse():
    index = "https://bond.szse.cn/api/report/ShowReport/data"
    doc_download='https://reportdocs.static.szse.cn'
    data = {
        'SHOWTYPE': 'JSON',
        'CATALOGID': 'xmjdxx',
        'TABKEY': 'tab1',
        'zqlb': '0',
        'selectXmztt': '已反馈,通过,提交注册,注册生效,终止'
    }
    i = 0
    year = 2025
    while year == 2025:
        i += 1
        data['PAGENO'] = str(i)
        reports = rq.get(index,params=data)
        for bond in reports.json()[0]['data']:
            soup = BeautifulSoup(bond['zqmc'], 'html.parser').find('a')
            bond_url, bond_name = soup.get('a-param'), soup.string
            update_time = bond['xmztgxrq']
            state = bond['xmzt']
            print(f'processing {bond_name}, state: {state}, update_time: {update_time}')
            year = update_time
            bond_page = rq.get(index + '?' + bond_url.split('?')[1])
            doc_feedback = {}
            for feedback in bond_page.json()[2]['data']:
                soup = BeautifulSoup(feedback['fkyjh'], 'html.parser').find('a')
                doc_url, doc_name = soup.get('encode-open'), soup.string
                if not os.path.splitext(doc_name)[0].endswith('反馈意见'):
                    continue
                print(f'processing {doc_name}')
                doc_string = rq.get(doc_download + doc_url)
                with open(f'../files/{doc_name}', 'wb') as f:
                    f.write(doc_string.content)
                doc_feedback[feedback['fkyjhgxrq']] = process_doc(doc_name)
                os.remove(f'../files/{doc_name}')
                time.sleep(1)
            bond_info = {
                'name': 'bond_name',
                'state': state,
                'date': update_time,
                'feedback': doc_feedback
            }
            feedback_to_excel(bond_info)
            time.sleep(1)
            break
        break
crawl_szse()
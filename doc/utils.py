import os
import time
import requests as rq

import docx
import pdfplumber
import pandas as pd
from win32com import client

def safe_request(url, **kwargs):
    max_retry = 5
    while max_retry:
        try:
            reports = rq.get(url, **kwargs)
            return reports
        except Exception as e:
            max_retry -= 1
            time.sleep(2)
            continue
    print(f'Request to {url} failed! Check the network.')
    exit(0)


def doc2docx(doc_name):
    word = client.Dispatch("Word.Application")
    abs_doc_path = os.path.abspath(doc_name)
    abs_doc_dir = os.path.dirname(abs_doc_path)
    doc = word.Documents.Open(abs_doc_path)
    doc.SaveAs(f"{abs_doc_dir}/tmp.docx", 12)  # 12代表docx格式
    doc.Close()
    os.rename(f"{abs_doc_dir}/tmp.docx", f'{abs_doc_path}x')
    os.remove(abs_doc_path)


def process_docx(doc_file):
    doc = docx.Document(doc_file)
    texts = []
    for paragraph in doc.paragraphs:
        texts.append(paragraph.text)
    doc_file.close()
    start, end = 0, -1
    for i in range(len(texts)):
        if texts[i].endswith('反馈意见：'):
            start = i
        elif texts[i].find('请你公司对上述问题逐项落实') != -1:
            end = i
    feedback = '\n'.join(texts[start+1:end]).strip()
    return feedback


def process_pdf(pdf_file):
    pdf_reader = pdfplumber.open(pdf_file)
    if not pdf_reader.pages[0].extract_text():
        return '文件处理失败'
    if len(pdf_reader.pages) == 1:
        lines = pdf_reader.pages[0].extract_text()
        text = ''.join(lines.split('\n'))
        index1 = text.find('反馈意见：')
        index2 = text.find('请你公司对上述问题逐项落实')
        return text[index1 + 5:index2]
    else:
        feedback = ''
        begin = False
        end = False
        for page in pdf_reader.pages:
            lines = page.extract_text()
            text = lines.split('\n')
            if text[-1][0] == '-':
                text = text[:-1]
            text = ''.join(text)
            if not begin:
                index = text.find('反馈意见：')
                if index != -1:
                    begin = True
                    feedback += text[index + 5:]
            elif not end:
                index = text.find('请你公司对上述问题逐项落实')
                if index != -1:
                    feedback += text[:index]
                    break
                else:
                    feedback += text
        return feedback


def process_feedback(file, file_format):
    if file_format == '.pdf':
        return process_pdf(file)
    elif file_format == '.docx':
        return process_docx(file)
    elif file_format == '.doc':
        with open('./tmp_file.doc', 'wb') as f:
            f.write(file.getvalue())
        doc2docx('./tmp_file.doc')
        with open('./tmp_file.docx', 'rb') as f:
            feedback = process_docx(f)
        os.remove('./tmp_file.docx')
        return feedback
    else:
        return '文件处理失败'


def feedback_to_excel(bond_info, suffix):
    if os.path.exists(f'./result_{suffix}.xlsx'):
        df = pd.read_excel(f'./result_{suffix}.xlsx')
    else:
        df = pd.DataFrame({
            '债券名': [],
            '状态': [],
            '反馈意见': [],
            '反馈时间': [],
            '债券类别': [],
            '申报规模（亿元）': [],
            '发行人': [],
            '地区': [],
            '承销商/管理人': [],
            '文件文号': [],
            '受理日期': [],
            '更新日期': [],
        })
    feedback = bond_info['feedback'] or {'N/A': 'N/A'}
    for date, text in feedback.items():
        new_row = pd.DataFrame({
            '债券名': [bond_info['name']],
            '状态': [bond_info['state']],
            '反馈意见': [text],
            '反馈时间': [date],
            '债券类别': [bond_info['bond_type']],
            '申报规模（亿元）': [bond_info['scale']],
            '发行人': [bond_info['issuer']],
            '地区': [bond_info['area']],
            '承销商/管理人': [bond_info['underwriter']],
            '文件文号': [bond_info['file_id']],
            '受理日期': [bond_info['accept_date']],
            '更新日期': [bond_info['update_date']],
        })
        df = pd.concat([df, new_row])
    df.to_excel(f'./result_{suffix}.xlsx', index=False)

def sort_excel(suffix):
    if not os.path.exists(f'./result_{suffix}.xlsx'):
        return
    df = pd.read_excel(f'./result_{suffix}.xlsx')
    df = df.sort_values(by='受理日期', ascending=False)
    df.to_excel(f'./result_{suffix}.xlsx', index=False)

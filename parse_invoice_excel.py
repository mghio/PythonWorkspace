#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# 安装如下必要依赖
"""
pip3 install pdfplumber
pip3 install pandas
"""

import pdfplumber
import re
import os
import pandas as pd


# 调试模式(1-开启，0-关闭)
debug = 0

def re_text(bt, text):
    m1 = re.search(bt, text)
    if m1 is not None:
        return re_block(m1[0])

def re_block(text):
    return text.replace(' ', '').replace('　', '').replace('）', '').replace(')', '').replace('：', ':')

def get_pdf(dir_path):
    pdf_file = []
    for root, sub_dirs, file_names in os.walk(dir_path):
        for name in file_names:
            if name.endswith('.pdf'):
                filepath = os.path.join(root, name)
                pdf_file.append(filepath)
    return pdf_file

def read(invoice_dir):
    filenames = get_pdf(invoice_dir)
    results = []
    for filename in filenames:
        print(filename)
        with pdfplumber.open(filename) as pdf:
            cont = {}

            first_page = pdf.pages[0]
            pdf_text = first_page.extract_text()
            if '发票' not in pdf_text:
                continue

            if debug == 1:
                print('--------------------------------------------------------')

            # 发票名称
            general_invoice_name = re_text(re.compile(r'[\u4e00-\u9fa5]+电子普通发票.*?'), pdf_text)
            cont['发票名称'] = general_invoice_name
            if debug == 1:
                print(general_invoice_name)
            special_invoice_name = re_text(re.compile(r'[\u4e00-\u9fa5]+专用发票.*?'), pdf_text)
            if special_invoice_name:
                cont['发票名称'] = special_invoice_name
                if debug == 1:
                    print(special_invoice_name)

            # 发票代码
            invoice_code = re_text(re.compile(r'发票代码(.*\d+)'), pdf_text).split('发票代码:')[1]
            cont['发票代码'] = invoice_code
            if debug == 1:
                print(invoice_code)

            # 发票号码
            invoice_number = re_text(re.compile(r'发票号码(.*\d+)'), pdf_text).split('发票号码:')[1]
            cont['发票号码'] = invoice_number
            if debug == 1:
                print(invoice_number)

            # 校验码
            invoice_verify_code = re_text(re.compile(r'校 验 码(.*\d+)'), pdf_text).split('校验码:')[1]
            cont['校验码'] = invoice_verify_code
            if debug == 1:
                print(invoice_verify_code)
            
            # 开票日期
            invoice_create_time = re_text(re.compile(r'开票日期(.*)'), pdf_text).split('开票日期:')[1]
            cont['开票日期'] = invoice_create_time
            if debug == 1:
                print(invoice_create_time)

            # 购买方名称
            invoice_purchaser_name = re_text(re.compile(r'名\s*称\s*[:：]\s*([\u4e00-\u9fa5]+)'), pdf_text).split('名称:')[1]
            cont['购买方名称'] = invoice_purchaser_name
            if debug == 1:
                print(invoice_purchaser_name)

            # 纳税人识别号
            taxpayer_identify_number = re_text(re.compile(r'纳税人识别号\s*[:：]\s*([a-zA-Z0-9]+)'), pdf_text).split('纳税人识别号:')[1]
            cont['纳税人识别号'] = taxpayer_identify_number
            if debug == 1:
                print(taxpayer_identify_number)
            
            # 发票金额
            invoice_amount = re_text(re.compile(r'小写.*(.*[0-9.]+)'), pdf_text).split('小写¥')[1]
            cont['发票金额(元)'] = invoice_amount
            if debug == 1:
                print(invoice_amount)

            # 销售方名称
            invoice_seller_name = re.findall(re.compile(r'名.*称\s*[:：]\s*([\u4e00-\u9fa5]+)'), pdf_text)
            if invoice_seller_name:
                invoice_seller_name = re_block(invoice_seller_name[len(invoice_seller_name)-1])
                cont['销售方名称'] = invoice_seller_name
                if debug == 1:
                    print(invoice_seller_name)

            results.append(cont)    
            if debug == 1:
                print('--------------------------------------------------------')
    
    if debug == 1:
        print(results)

    return results
        

def save_to_excel(invoice_dir, results):
    pf = pd.DataFrame(results)

    order = ["发票名称", "发票代码", "发票号码", "校验码", "开票日期", "发票代码", "发票号码", "开票日期", "购买方名称", "纳税人识别号", "发票金额(元)", "销售方名称"]  # 指定列的顺序
    pf = pf[order]
    file_path = pd.ExcelWriter(invoice_dir + '/发票.xlsx')  # 打开excel文件
    # 替换空单元格
    pf.fillna(' ', inplace=True)
    # 输出
    pf.to_excel(file_path, encoding='utf-8', index=False, sheet_name="sheet1")
    file_path.save()


if __name__ == '__main__':
    # 发票目录
    invoice_dir = '/Users/mghio/Desktop/发票'
    # read pdf to object array
    results = read(invoice_dir)
    # save to excel
    save_to_excel(invoice_dir, results)
# -*- coding: utf-8 -*-
import requests
from bs4 import BeautifulSoup
import pandas as pd
import openpyxl

def getNDLItemsByAuthor(author):
    target_url = 'https://iss.ndl.go.jp/api/opensearch?creator='+author
    soup = BeautifulSoup(requests.get(target_url).text, 'lxml')
    
    linksToItemPages = []
    for guid in soup.find_all('guid'):
        linksToItemPages.append(guid.text)
    
    itemPages = []
    for link in linksToItemPages:
        itemPages.append(BeautifulSoup(requests.get(link).text, 'lxml'))
    
    items = []
    for itemPage in itemPages:
        item = {}
        item['出版社・発行所'] = ''
        rows = itemPage.find_all('tr')
        for row in rows:
            if not row.th is None:
                if row.th.text.strip() == 'タイトル':
                    item['著書・論文名'] = row.td.text.strip()
                if row.th.text.strip() in ['掲載誌名','掲載誌情報（URI形式）']:
                     item['収録書誌名'] = row.td.text.strip()
                if row.th.text.strip() == '著者':
                     item['著者'] = row.td.text.strip()
                if row.th.text.strip() == '出版社':
                     item['出版社・発行所'] = row.td.text.strip()
                if row.th.text.strip() in ['出版年(W3CDTF)','出版年月日等']:
                     item['出版年'] = row.td.text.strip()
                if row.th.text.strip() in ['掲載号','掲載巻']:
                     item['巻・号'] = row.td.text.strip()
                if item['出版社・発行所'] == '':
                    if row.th.text.strip() in ['掲載誌名','掲載誌情報（URI形式）']:
                        linkToPubPage = row.td.a.get('href')
                        pubPage = BeautifulSoup(requests.get(linkToPubPage).text, 'lxml')
                        rows = pubPage.find_all('tr')
                        for row in rows:
                            if not row.th is None:
                                if row.th.text.strip() == '出版社':
                                    item['出版社・発行所'] = row.td.text.strip()     
                          
        items.append(item)
        
    columns = ['著書・論文名','収録書誌名','出版社・発行所','巻・号','出版年','著者']
    pd.DataFrame(items).to_excel('JPROK1.xlsx', sheet_name=author, columns=columns, encoding="cp932")



getNDLItemsByAuthor('吉澤文寿')

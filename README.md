# 使い方

1. [このページ](https://jupyter.org/try)で「Try Jupiter Lab」を選び、移動先のページの左上メニューにある「＋」を押して、「Notebook」内の「Python 3」を選んでください。

2. グレーの細長いボックスの中に、下のスクリプトをペーストしてください。
```python
!pip install bs4 openpyxl lxml
import requests
import openpyxl
import time
from bs4 import BeautifulSoup
from operator import itemgetter
import pandas as pd

def getNDLItemsByAuthor(author):
    target_url = 'https://iss.ndl.go.jp/api/opensearch?creator='+author
    soup = BeautifulSoup(requests.get(target_url).text, 'lxml')
    print('Search Results Obtained.')
    
    linksToItemPages = []
    for guid in soup.find_all('guid'):
        linksToItemPages.append(guid.text)  
        print('', end=f'\r{len(linksToItemPages)} Items Found.')
        time.sleep(0.2)
    print('')
    
    itemPages = []
    for link in linksToItemPages:
        itemPages.append(BeautifulSoup(requests.get(link).text, 'lxml'))
        print('', end=f'\rPage Data Extracted: {len(itemPages)/len(linksToItemPages)*100}%')
        time.sleep(0.2)
    print('')

    items = []
    for itemPage in itemPages:
        item = {
            '著者': '',
            '著書・論文名': '',
            '収録書誌名': '',
            '巻・号': '',
            '出版社・発行所': '',
            '出版年': '',
        }
        rows = itemPage.find_all('tr')
        for row in rows:
            if not row.th is None:
                if row.th.text.strip() == 'タイトル':
                    item['著書・論文名'] = row.td.text.strip()
                if row.th.text.strip() == '部分タイトル':
                    if author in row.td.text.strip():
                        item['収録書誌名'] = item['著書・論文名']
                        item['著書・論文名'] = row.td.text.strip()
                if row.th.text.strip() in ['掲載誌名','掲載誌情報（URI形式）']:
                    item['収録書誌名'] = row.td.text.strip()
                if row.th.text.strip() == '著者':
                    if item['著者'] == '':
                        item['著者'] = row.td.text.strip()
                    else:
                        item['著者'] = item['著者'] + ',' + row.td.text.strip()
                if row.th.text.strip() == '出版社':
                    item['出版社・発行所'] = row.td.text.strip()
                if row.th.text.strip() in ['出版年(W3CDTF)','出版年月日等']:
                    item['出版年'] = row.td.text.strip()
                if row.th.text.strip() in ['掲載号','掲載巻','掲載通号']:
                    item['巻・号'] = row.td.text.strip()
                if item['出版社・発行所'] == '':
                    if row.th.text.strip() in ['掲載誌名','掲載誌情報（URI形式）']:
                        if not row.td.a is None:
                            linkToPubPage = row.td.a.get('href')
                            pubPage = BeautifulSoup(requests.get(linkToPubPage).text, 'lxml')
                            rows = pubPage.find_all('tr')
                            for row in rows:
                                if not row.th is None:
                                    if row.th.text.strip() == '出版社':
                                        item['出版社・発行所'] = row.td.text.strip() 
        items.append(item)
        print('', end=f'\rPage Data Saved: {len(items)/len(itemPages)*100}%')
        time.sleep(0.2)
        items = sorted(items, key=itemgetter('出版年')) 
    print('')
    
    print('Writing Data to an Excel Sheet.')
    columns = ['著者','著書・論文名','収録書誌名','巻・号','出版社・発行所','出版年']
    pd.DataFrame(items).to_excel(author+'.xlsx', sheet_name=author, columns=columns, encoding="cp932")
    print('Mission Accomplished. Please Download the Output File.')
```

3. 上のメニューバーにある「＋」ボタンを押すと、スクリプトをペーストしたボックスの下に、新しいボックスが追加されます。新しいボックスに、下のコードを貼り付けたのち、「著者名」の部分に検索したい著者名を入れてください。

```python
getNDLItemsByAuthor('著者名')
```

4. 一つ目のボックス（スクリプトをペーストしたボックス）をクリックし、そのボックスが選択された状態にした上で、上のメニューバーにある再生ボタン（三角形のボタン）を2回連続で押すと、処理がスタートします。
. 処理が終わると、画面左のファイルリストに「著者名.xlsx」というExcelファイルが現れます。右クリックで「ダウンロード」を選択してください。

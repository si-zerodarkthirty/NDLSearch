# 使い方

1. [このページ](https://jupyter.org/try)で「Try Jupiter Lab」を選び、移動先のページの左上メニューにある「＋」を押して、「Notebook」内の「Python 3」を選んでください。

2. グレーの細長いボックスの中に、下のスクリプトをペーストしてください。
```python
!pip install bs4 openpyxl lxml
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
                if row.th.text.strip() == '部分タイトル':
                    if author in row.td.text.strip():
                        item['著書・論文名'] = row.td.text.strip()
                if row.th.text.strip() in ['掲載誌名','掲載誌情報（URI形式）']:
                     item['収録書誌名'] = row.td.text.strip()
                if row.th.text.strip() == '著者':
                     item['著者'] = item['著者'] + ' ' + row.td.text.strip()
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
        
    columns = ['著書・論文名','収録書誌名','出版社・発行所','巻・号','出版年','著者']
    pd.DataFrame(items).to_excel(author+'.xlsx', sheet_name=author, columns=columns, encoding="cp932")
```

3. 上のメニューバーにある「＋」ボタンを押すと、スクリプトをペーストしたボックスの下に、新しいボックスが追加されます。新しいボックスに、下のコードを貼り付けたのち、「著者名」の部分に検索したい著者名を入れてください。

```python
getNDLItemsByAuthor('著者名')
```

4. 一つ目のボックス（スクリプトをペーストしたやつ）をクリックして、選択状態にしてください。
5. 上のメニューバーにある再生ボタン（三角形のやつ）を2回連続で押すと、処理がスタートします。
6. 2つ目のボックスの横にある[ ]内が「*」から数字に変わるのが、処理が終わった合図です（だいたい10アイテムにつき1分ほど時間がかかります）。
7. 画面左のファイルリストに、「著者名.xlsx」というExcelファイルが現れるので、右クリックでダウンロードを選択してください。
8. 適宜、フォーマット修正や重複部分の削除などを行ってください。

# 注意事項・補足

- 完成後のExcelファイル内で、重複を削除したい場合は、Excelの「データ」から「重複の削除」を選ぶと自動でやってくれます。
- シートを一つのファイルに統合したい場合には、ファイル内のシートを右クリックして、「移動またはコピー」をクリックし、移動先のファイルを選んでください。
- 同姓同名の人は弾けません。

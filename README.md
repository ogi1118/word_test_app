実行ファイルの作成
```
 pyinstaller --onefile --add-data "vocabulary_books;vocabulary_books" --add-data "C:\\Users\\81705\\src\\wordTest\\poppler\\bin;poppler/bin" app.py
```

/dist/app.exeが実行ファイルです。
起動したら必要項目を選択、入力して印刷を押下します。
選択したプリンターで印刷が開始されます。

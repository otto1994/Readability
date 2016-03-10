####`read10k.py`運作上的邏輯是讀取`.docx`檔案，依照`item`分類，以`POST requests`型態向[可讀性網站](https://readability-score.com/) 中的`https://readability-score.com/ajax.php`送出請求，獲得可讀性數據後以`record.xls`存於工作資料夾底下

####`rea10k.py`以`python2`為運作基礎寫成(可能新增`python3`版本)，邏輯上來說若環境設置完整，`windows`與`macintosh`皆可運作無誤，若有問題麻煩以[郵件](ponienchiang@gmail.com)通知我

* 若`Python`尚未設置完成，可以參閱[環境設置](https://github.com/otto1994/Readability/tree/master/Python-Setting)
* 關於`.docx`檔案內文要求的修改方式，請參閱[範例](https://github.com/otto1994/Readability/tree/master/Example)
* 

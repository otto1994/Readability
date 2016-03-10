* ####`read10k.py`運作上的邏輯是讀取`.docx`檔案，依照`item`分類，以`POST requests`型態向[可讀性網站](https://readability-score.com/) 中的`https://readability-score.com/ajax.php`送出請求，獲得可讀性數據後以`record.xls`存於工作資料夾底下

* ####`rea10k.py`以`python2`為運作基礎寫成(可能新增`python3`版本)，邏輯上來說若環境設置完整，`windows`與`macintosh`皆可運作無誤，若發現問題，麻煩以臉書或郵件(ponienchiang@gmail.com)通知我，謝謝！

* 若`Python`尚未設置完成，可以參閱[環境設置](https://github.com/otto1994/Readability/tree/master/Python-Setting)
* 關於`.docx`檔案內文要求的修改方式，請參閱[範例](https://github.com/otto1994/Readability/tree/master/Example)
* 詳細版的操作說明，請參閱[操作說明](https://github.com/otto1994/Readability/blob/master/%E6%93%8D%E4%BD%9C%E8%AA%AA%E6%98%8E(%E8%A9%B3).md)
* 關於`read10k.py`效能的測試，請參閱[測試](https://github.com/otto1994/Readability/tree/master/Test)  (我目前測試的結果是跟用手貼沒有區別的，應該可以放心使用，但是歡迎進行任何形式的壓力測試，數據就...之後再放上來....)
* 欲一次下載`Readability`資料夾中的所有資料(建議)，可以在右上角找到`Download ZIP`按鈕

###為了讓`Anne.py`能夠自動讀取財報，所需的修改如下，可以以1389修改前後作參考

* 2016/03/17新增`@other@`，用來表示不屬於任何`item`但是在`overall`計算的段落
* 新增`只保留文字`

####`只保留文字`
#####考慮到程式讀取文字的判定，在處理文檔上應`Ctrl+A`全選>`Ctrl+X`剪下>右鍵貼上選項`只保留文字`，以確保產出的數據完全符合手動貼到可讀性網站的數據。

####`item`判定
為了使`Anne`判定每個`item`的範圍，`item`的命名必須一致如下(若該item不存在則可以忽略)
* @item1@
* @item1a@
* @item1b@
* @item2@
* @item3@
* @item4@
* @item4a@
* @item5@
* @item6@
* @item7@
* @ResultsOfOperations@
* @item7a@
* @item8@
* @item9@
* @item9a@
* @item9b@
* @item10@
* @other@
 
#####其他比較需要注意的重點有：

1.空格、大小寫、拼字必須完全符合，否則`Anne.py`會視該`item`為不存在，比如 `@item 1@`、`@Item1@`   、`@ResultOfOperations@`、`@ResultsOfOperation@`都是不合格範例(results of operations正確寫法是每個單字開頭大寫，複數結尾，如果真的覺得麻煩可以在程式碼中進行修正)

2.有些`item`底下的內容是`none`或者`not applicable`，這個類型的內文由於我們不會計算可讀性，所以必須將這些字眼刪除，以防造成程式判斷應該讀取

3.`item10`後面的`item11`到`item14`標題應刪除

4.內文中仍然有可能有`@`出現(雖然機率很低)，應該要用`Ctrl+F`進行搜尋確認




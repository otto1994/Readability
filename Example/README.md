###為了讓`read10k.py`能夠自動讀取財報，所需的修改如下，可以以1389修改前後作參考

####`item`判定
為了使`read10k`判定每個`item`的範圍，`item`的命名必須一致如下(若該item不存在則可以忽略)
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

其中比較需要注意的重點有：

1.空格、大小寫、拼字必須完全符合，否則`read10k.py`會視該`item`為不存在，比如 `@item 1@`、`@Item1@`   、`@ResultOfOperations@`、`@ResultsOfOperation@`都是不合格範例(results of operations正確寫法是每個單字開頭大寫，複數結尾，如果真的覺得麻煩可以在程式碼中進行修正)

2.有些`item`底下的內容是`none`或者`not applicable`，這個類型的內文由於我們不會計算可讀性，所以必須將這些字眼刪除，以防造成程式判斷應該讀取

3.`item10`後面的`item11`到`item14`標題應刪除



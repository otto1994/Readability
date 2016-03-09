##為了讓`read10k.py`能夠自動讀取財報，所需的修改如下，可以以1389修改前後作參考

1.`item`判定
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

1.空格、大小寫、拼字必須完全符合，否則`read10k.py`會視該`item`為不存在，比如`@Item1@`、`@item 1@`、`ResultOfOperations`、`ResultsOfOperation`都是不合格範例()

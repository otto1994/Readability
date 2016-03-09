#此說明分成三步驟解釋Python及相關套件的安裝設置：

###1.Python主程式

###2.pip

###3.套件(package)


#1.Python主程式

此處程式碼是以`Python2`為運作基礎寫成的，所以應該下載的是`Python2.7`。(但是Python將來只會在`Python3`上進行更新，使用`Python2`是為避免需要套件尚未自`Python2`轉移至`Python3`的權宜之計)
![0](https://github.com/otto1994/Readability/blob/master/figure/0.png)

在下載並安裝`Python`主程式後，可按`windows`+`R`鍵後輸入`cmd`打開終端機
![1](https://github.com/otto1994/Readability/blob/master/figure/1.jpg)
(這是`windows`鍵)
![2](https://github.com/otto1994/Readability/blob/master/figure/2.png)
![3](https://github.com/otto1994/Readability/blob/master/figure/3.png)

打開終端機後，輸入`python`，若出現以下畫面代表`Python`主程式之安裝及設置已成功

![4](https://github.com/otto1994/Readability/blob/master/figure/4.png)

如果沒有如上圖顯示`Python`版本，返回如`python is an invalid syntax` 之類的句子(這是正常狀態)，可能原因如下

1.下載了主程式但是還沒點開來安裝(真蠢)

2.終端機找不到`Python`

在終端機找不到`Python`的情況下，我們必須對路徑進行設置

首先，先找到`Python`所在的資料夾確認它的存在(預設上是很大方的躺在C槽)
![5](https://github.com/otto1994/Readability/blob/master/figure/5.png)

接著 `開始`>`控制台`>`系統`>`進階系統設定`>`環境變數` 找到`PATH`後進行`編輯`(不同版本的WINDOWS在位置上可能略有差異)

![6](https://github.com/otto1994/Readability/blob/master/figure/6.png)
![7](https://github.com/otto1994/Readability/blob/master/figure/7.png)
![8](https://github.com/otto1994/Readability/blob/master/figure/8.png)

在現有路徑的最後面打入`;`+`python資料夾位置`(圖中為`C:\Python27`)

![9](https://github.com/otto1994/Readability/blob/master/figure/9.png)

重新開啟終端機後即可執行`Python`(檢驗是否設置成功請參見前面步驟)

#2.pip

`pip`的用途是取得並安裝`Python`套件，首先將`Python-Setting`資料夾中的`get-pip.py`下載並放置於桌面(其實放哪裡都可以，只是為了方便)

![10](https://github.com/otto1994/Readability/blob/master/figure/10.png)

接著打開終端機，將工作資料夾(終端機找檔案的地方)移動到桌面(`cd Desktop`指令，cd的意思是change directory，注意 `D`大寫)，並輸入`python get-pip.py`用`Python`來取得`pip`

![11](https://github.com/otto1994/Readability/blob/master/figure/11.png)

安裝結束後輸入`pip`若出現以下畫面代表成功

![12](https://github.com/otto1994/Readability/blob/master/figure/12.png)

如果又出現`invalid syntax`之類的字眼，代表我們的老朋友路徑設置先生又必須上工了

首先找到`pip`所在的資料夾，預設狀態下是`Python27`>`Script`

![13](https://github.com/otto1994/Readability/blob/master/figure/13.png)

接著再一次 `開始`>`控制台`>`系統`>`進階系統設定`>`環境變數` 找到`PATH`後進行`編輯`，添加路徑

![14](https://github.com/otto1994/Readability/blob/master/figure/14.png)

重新開啟終端機後即可執行`pip`(檢驗是否設置成功請參見前面步驟)

#3.套件(package)

為了運作`read10k.py`，總共有六個套件必須安裝：`regex`、`requests`、`xlwt`、`xlrd`、`xlutils`及`python-docx`

安裝套件的方法是在終端機中輸入`pip` `install` `套件名稱`(圖中即為安裝`regex`的指令)

![15](https://github.com/otto1994/Readability/blob/master/figure/15.png)

在安裝完前五個套件後，應該會在安裝`python-docx`套件時遭遇一些問題(紅字預警)，此時將`Python-Setting`資料夾中的`xml-3.5.0.win32-py2.7.exe`下載並安裝後(如果無法安裝，可以試試`lxml-3.5.0.win-amd64-py2.7.exe`)，再執行一次`pip install python-docx`即可看到`successful`字樣

至此，環境設置部分告一段落












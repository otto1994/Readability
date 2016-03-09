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









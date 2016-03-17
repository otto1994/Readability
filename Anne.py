# -*- coding: utf-8 -*-

'''
安妮工作的重點：

0.為了防止網路不穩造成資料損失，每完成一份.docx數據紀錄即會儲存到工作表上。如此，若在讀取資料夾中大量檔案時因故中斷，可以接續先前的進度。
  只有一個比較重要的重點是，安妮會讀取資料夾中"全部".docx檔，故不欲讀取的.docx檔應移出資料夾外，非.docx檔不在此限。
  讀取.docx檔的順序是照命名排序。
  
1.在目標.docx檔開啟的狀態下無法運行

2.item後有不打算計算的字串(如none, not applicable)，若沒有刪掉仍會計算可讀性

3.對於md&a 1這一欄位，狀況如下：
    1.沒有md&a1 沒有results of operations：印出merged to item 7(與此同時item7那欄是n.a)
	2.有  md&a1 沒有results of operations：印出merged to item 7
	3.沒有md&a1 有  results of operations：印出n.a
	4.有  md&a1 有  results of operations：記錄md&1數據於此欄位

4.此版本已經可以保留item作為標題，若該item底下沒有東西(只有空白格之類)會印出n.a，不必特意將空的item刪除
'''
#介面，要求輸入一個正整數(也沒有非正整數的行數這種東西)
row = input("Specify the starting row (a positive integer) : ")
#設置
import glob
from xlwt import Workbook
from xlrd import open_workbook
from xlutils.copy import copy
from docx import Document
import re
import requests
import os.path
#找到資料夾中所有的.docx檔
file = glob.glob("*.docx")

for d in range(0,len(file)):
    if os.path.exists("record.xls")!= True:
        wb = Workbook()
        sheet1 = wb.add_sheet('Sheet1')
    else:
        rb = open_workbook('record.xls')
        wb = copy(rb)
        sheet1 = wb.get_sheet(0)

    i = row+d-1
    document = Document(file[d])	
    docText = '\n'.join([paragraph.text.encode('utf-8') for paragraph in document.paragraphs])
    n = docText.count('@')+1
    sp = re.split(r'@*',docText)

#item1
    for j in range(0,n):
		if sp[j] == "item1":
			item1 = sp[j+1]
		if "item1" not in sp:
			item1 = ""
    if len(item1.split()) == 0:
        item1 = ""
    if item1 != "":
        payload = {"type":"text","text":item1}
        res = requests.post("https://readability-score.com/ajax.php",data=payload)
        text = str(res.text)
        split = re.split(r'"*',text)
        fkg = split[4].replace(':','')
        f = fkg.replace(',','')
        gf = split[6].replace(':','')
        g = gf.replace(',','')
        cl = split[8].replace(':','')
        c = cl.replace(',','')
        si = split[10].replace(':','')
        s = si.replace(',','')
        ar = split[12].replace(':','')
        a = ar.replace(',','')
        ag = split[15].replace(':','')
        average = ag.replace(',','')
        words = split[27].replace(',','')
        sentences = split[19].replace(',','')
        sheet1.write(i,0,file[d])
        sheet1.write(i,1,float(f))
        sheet1.write(i,2,float(g))
        sheet1.write(i,3,float(c))
        sheet1.write(i,4,float(s))
        sheet1.write(i,5,float(a))
        sheet1.write(i,6,float(average))
        sheet1.write(i,7,float(words))
        sheet1.write(i,8,float(sentences))
    else:
        sheet1.write(i,0,file[d])
        sheet1.write(i,1,"N.A")	

#item1a
		
    for j in range(0,n):
		if sp[j] == "item1a":
			item1a = sp[j+1]
		if "item1a" not in sp:
			item1a = ""
    if len(item1a.split()) == 0:
        item1a = ""
    if item1a != "":
        payload = {"type":"text","text":item1a}
        res = requests.post("https://readability-score.com/ajax.php",data=payload)
        text = str(res.text)
        split = re.split(r'"*',text)
        fkg = split[4].replace(':','')
        f = fkg.replace(',','')
        gf = split[6].replace(':','')
        g = gf.replace(',','')
        cl = split[8].replace(':','')
        c = cl.replace(',','')
        si = split[10].replace(':','')
        s = si.replace(',','')
        ar = split[12].replace(':','')
        a = ar.replace(',','')
        ag = split[15].replace(':','')
        average = ag.replace(',','')
        words = split[27].replace(',','')
        sentences = split[19].replace(',','')
        sheet1.write(i,9,float(f))
        sheet1.write(i,10,float(g))
        sheet1.write(i,11,float(c))
        sheet1.write(i,12,float(s))
        sheet1.write(i,13,float(a))
        sheet1.write(i,14,float(average))
        sheet1.write(i,15,float(words))
        sheet1.write(i,16,float(sentences))
    else:
        sheet1.write(i,9,"N.A")				

#item1b
    for j in range(0,n):
		if sp[j] == "item1b":
			item1b = sp[j+1]
		if "item1b" not in sp:
			item1b = ""
    if len(item1b.split()) == 0:
        item1b = ""
    if item1b != "":
        payload = {	"type":"text","text":item1b}
        res = requests.post("https://readability-score.com/ajax.php",data=payload)
        text = str(res.text)
        split = re.split(r'"*',text)
        fkg = split[4].replace(':','')
        f = fkg.replace(',','')
        gf = split[6].replace(':','')
        g = gf.replace(',','')
        cl = split[8].replace(':','')
        c = cl.replace(',','')
        si = split[10].replace(':','')
        s = si.replace(',','')
        ar = split[12].replace(':','')
        a = ar.replace(',','')
        ag = split[15].replace(':','')
        average = ag.replace(',','')
        words = split[27].replace(',','')
        sentences = split[19].replace(',','')
        sheet1.write(i,17,float(f))
        sheet1.write(i,18,float(g))
        sheet1.write(i,19,float(c))
        sheet1.write(i,20,float(s))
        sheet1.write(i,21,float(a))
        sheet1.write(i,22,float(average))
        sheet1.write(i,23,float(words))
        sheet1.write(i,24,float(sentences))
    else:
        sheet1.write(i,17,"N.A")	
#item2
    for j in range(0,n):
		if sp[j] == "item2":
			item2 = sp[j+1]
		if "item2" not in sp:
			item2 = ""
    if len(item2.split()) == 0:
        item2 = ""
    if item2 != "":
        payload = {	"type":"text","text":item2}
        res = requests.post("https://readability-score.com/ajax.php",data=payload)
        text = str(res.text)
        split = re.split(r'"*',text)
        fkg = split[4].replace(':','')
        f = fkg.replace(',','')
        gf = split[6].replace(':','')
        g = gf.replace(',','')
        cl = split[8].replace(':','')
        c = cl.replace(',','')
        si = split[10].replace(':','')
        s = si.replace(',','')
        ar = split[12].replace(':','')
        a = ar.replace(',','')
        ag = split[15].replace(':','')
        average = ag.replace(',','')
        words = split[27].replace(',','')
        sentences = split[19].replace(',','')
        sheet1.write(i,25,float(f))
        sheet1.write(i,26,float(g))
        sheet1.write(i,27,float(c))
        sheet1.write(i,28,float(s))
        sheet1.write(i,29,float(a))
        sheet1.write(i,30,float(average))
        sheet1.write(i,31,float(words))
        sheet1.write(i,32,float(sentences))
    else:
        sheet1.write(i,25,"N.A")	
#item3
    for j in range(0,n):
		if sp[j] == "item3":
			item3 = sp[j+1]
		if "item3" not in sp:
			item3 = ""
    if len(item3.split()) == 0:
        item3 = ""
    if item3 != "":
        payload = {	"type":"text","text":item3}
        res = requests.post("https://readability-score.com/ajax.php",data=payload)
        text = str(res.text)
        split = re.split(r'"*',text)
        fkg = split[4].replace(':','')
        f = fkg.replace(',','')
        gf = split[6].replace(':','')
        g = gf.replace(',','')
        cl = split[8].replace(':','')
        c = cl.replace(',','')
        si = split[10].replace(':','')
        s = si.replace(',','')
        ar = split[12].replace(':','')
        a = ar.replace(',','')
        ag = split[15].replace(':','')
        average = ag.replace(',','')
        words = split[27].replace(',','')
        sentences = split[19].replace(',','')
        sheet1.write(i,33,float(f))
        sheet1.write(i,34,float(g))
        sheet1.write(i,35,float(c))
        sheet1.write(i,36,float(s))
        sheet1.write(i,37,float(a))
        sheet1.write(i,38,float(average))
        sheet1.write(i,39,float(words))
        sheet1.write(i,40,float(sentences))
    else:
        sheet1.write(i,33,"N.A")
#item4		
    for j in range(0,n):
		if sp[j] == "item4":
			item4 = sp[j+1]
		if "item4" not in sp:
			item4 = ""
    if len(item4.split()) == 0:
        item4 = ""
    if item4 != "":
        payload = {	"type":"text","text":item4}
        res = requests.post("https://readability-score.com/ajax.php",data=payload)
        text = str(res.text)
        split = re.split(r'"*',text)
        fkg = split[4].replace(':','')
        f = fkg.replace(',','')
        gf = split[6].replace(':','')
        g = gf.replace(',','')
        cl = split[8].replace(':','')
        c = cl.replace(',','')
        si = split[10].replace(':','')
        s = si.replace(',','')
        ar = split[12].replace(':','')
        a = ar.replace(',','')
        ag = split[15].replace(':','')
        average = ag.replace(',','')
        words = split[27].replace(',','')
        sentences = split[19].replace(',','')
        sheet1.write(i,41,float(f))
        sheet1.write(i,42,float(g))
        sheet1.write(i,43,float(c))
        sheet1.write(i,44,float(s))
        sheet1.write(i,45,float(a))
        sheet1.write(i,46,float(average))
        sheet1.write(i,47,float(words))
        sheet1.write(i,48,float(sentences))
    else:
        sheet1.write(i,41,"N.A")	
#item4a
    for j in range(0,n):
		if sp[j] == "item4a":
			item4a = sp[j+1]
		if "item4a" not in sp:
			item4a = ""
    if len(item4a.split()) == 0:
        item4a = ""
    if item4a != "":
        payload = {	"type":"text","text":item4a}
        res = requests.post("https://readability-score.com/ajax.php",data=payload)
        text = str(res.text)
        split = re.split(r'"*',text)
        fkg = split[4].replace(':','')
        f = fkg.replace(',','')
        gf = split[6].replace(':','')
        g = gf.replace(',','')
        cl = split[8].replace(':','')
        c = cl.replace(',','')
        si = split[10].replace(':','')
        s = si.replace(',','')
        ar = split[12].replace(':','')
        a = ar.replace(',','')
        ag = split[15].replace(':','')
        average = ag.replace(',','')
        words = split[27].replace(',','')
        sentences = split[19].replace(',','')
        sheet1.write(i,49,float(f))
        sheet1.write(i,50,float(g))
        sheet1.write(i,51,float(c))
        sheet1.write(i,52,float(s))
        sheet1.write(i,53,float(a))
        sheet1.write(i,54,float(average))
        sheet1.write(i,55,float(words))
        sheet1.write(i,56,float(sentences))
    else:
        sheet1.write(i,49,"N.A")	
#item5
    for j in range(0,n):
		if sp[j] == "item5":
			item5 = sp[j+1]
		if "item5" not in sp:
			item5 = ""
    if len(item5.split()) == 0:
        item5 = ""
    if item5 != "":
        payload = {	"type":"text","text":item5}
        res = requests.post("https://readability-score.com/ajax.php",data=payload)
        text = str(res.text)
        split = re.split(r'"*',text)
        fkg = split[4].replace(':','')
        f = fkg.replace(',','')
        gf = split[6].replace(':','')
        g = gf.replace(',','')
        cl = split[8].replace(':','')
        c = cl.replace(',','')
        si = split[10].replace(':','')
        s = si.replace(',','')
        ar = split[12].replace(':','')
        a = ar.replace(',','')
        ag = split[15].replace(':','')
        average = ag.replace(',','')
        words = split[27].replace(',','')
        sentences = split[19].replace(',','')
        sheet1.write(i,57,float(f))
        sheet1.write(i,58,float(g))
        sheet1.write(i,59,float(c))
        sheet1.write(i,60,float(s))
        sheet1.write(i,61,float(a))
        sheet1.write(i,62,float(average))
        sheet1.write(i,63,float(words))
        sheet1.write(i,64,float(sentences))
    else:
        sheet1.write(i,57,"N.A")	
#item6
    for j in range(0,n):
		if sp[j] == "item6":
			item6 = sp[j+1]
		if "item6" not in sp:
			item6 = ""
    if len(item6.split()) == 0:
        item6 = ""
    if item6 != "":
        payload = {	"type":"text","text":item6}
        res = requests.post("https://readability-score.com/ajax.php",data=payload)
        text = str(res.text)
        split = re.split(r'"*',text)
        fkg = split[4].replace(':','')
        f = fkg.replace(',','')
        gf = split[6].replace(':','')
        g = gf.replace(',','')
        cl = split[8].replace(':','')
        c = cl.replace(',','')
        si = split[10].replace(':','')
        s = si.replace(',','')
        ar = split[12].replace(':','')
        a = ar.replace(',','')
        ag = split[15].replace(':','')
        average = ag.replace(',','')
        words = split[27].replace(',','')
        sentences = split[19].replace(',','')
        sheet1.write(i,65,float(f))
        sheet1.write(i,66,float(g))
        sheet1.write(i,67,float(c))
        sheet1.write(i,68,float(s))
        sheet1.write(i,69,float(a))
        sheet1.write(i,70,float(average))
        sheet1.write(i,71,float(words))
        sheet1.write(i,72,float(sentences))
    else:
        sheet1.write(i,65,"N.A")
#item7	roo and item7a	
    for j in range(0,n):
		if sp[j] == "item7":
			mda1 = sp[j+1]
		if "item7" not in sp:
			mda1 = ""
    if len(mda1.split()) == 0:
        mda1 = ""

    for j in range(0,n):
		if sp[j] == "ResultsOfOperations":
			roo = sp[j+1]
		if "ResultsOfOperations" not in sp:
			roo = ""
    if len(roo.split()) == 0:
        roo = ""

    for j in range(0,n):
		if sp[j] == "item7a":
			item7a = sp[j+1]
		if "item7a" not in sp:
			item7a = ""
    if len(item7a.split()) == 0:
        item7a = ""
#item7
    if mda1+roo != "":
        payload = {	"type":"text","text":mda1+roo}
        res = requests.post("https://readability-score.com/ajax.php",data=payload)
        text = str(res.text)
        split = re.split(r'"*',text)
        fkg = split[4].replace(':','')
        f = fkg.replace(',','')
        gf = split[6].replace(':','')
        g = gf.replace(',','')
        cl = split[8].replace(':','')
        c = cl.replace(',','')
        si = split[10].replace(':','')
        s = si.replace(',','')
        ar = split[12].replace(':','')
        a = ar.replace(',','')
        ag = split[15].replace(':','')
        average = ag.replace(',','')
        words = split[27].replace(',','')
        sentences = split[19].replace(',','')
        sheet1.write(i,73,float(f))
        sheet1.write(i,74,float(g))
        sheet1.write(i,75,float(c))
        sheet1.write(i,76,float(s))
        sheet1.write(i,77,float(a))
        sheet1.write(i,78,float(average))
        sheet1.write(i,79,float(words))
        sheet1.write(i,80,float(sentences))
    else:
        sheet1.write(i,73,"N.A")	
		
#MD&A_Def1	
    if roo == "":
        sheet1.write(i,81,"merged to item 7")
    if (roo != "" and mda1 == "") :
        sheet1.write(i,81,"N.A")
    if (roo != "" and mda1 != ""):
        payload = {	"type":"text","text":mda1}
        res = requests.post("https://readability-score.com/ajax.php",data=payload)
        text = str(res.text)
        split = re.split(r'"*',text)
        fkg = split[4].replace(':','')
        f = fkg.replace(',','')
        gf = split[6].replace(':','')
        g = gf.replace(',','')
        cl = split[8].replace(':','')
        c = cl.replace(',','')
        si = split[10].replace(':','')
        s = si.replace(',','')
        ar = split[12].replace(':','')
        a = ar.replace(',','')
        ag = split[15].replace(':','')
        average = ag.replace(',','')
        words = split[27].replace(',','')
        sentences = split[19].replace(',','')
        sheet1.write(i,81,float(f))
        sheet1.write(i,82,float(g))
        sheet1.write(i,83,float(c))
        sheet1.write(i,84,float(s))
        sheet1.write(i,85,float(a))
        sheet1.write(i,86,float(average))
        sheet1.write(i,87,float(words))
        sheet1.write(i,88,float(sentences))
#item7a
    if item7a != "":
        payload = {	"type":"text","text":item7a}
        res = requests.post("https://readability-score.com/ajax.php",data=payload)
        text = str(res.text)
        split = re.split(r'"*',text)
        fkg = split[4].replace(':','')
        f = fkg.replace(',','')
        gf = split[6].replace(':','')
        g = gf.replace(',','')
        cl = split[8].replace(':','')
        c = cl.replace(',','')
        si = split[10].replace(':','')
        s = si.replace(',','')
        ar = split[12].replace(':','')
        a = ar.replace(',','')
        ag = split[15].replace(':','')
        average = ag.replace(',','')
        words = split[27].replace(',','')
        sentences = split[19].replace(',','')
        sheet1.write(i,89,float(f))
        sheet1.write(i,90,float(g))
        sheet1.write(i,91,float(c))
        sheet1.write(i,92,float(s))
        sheet1.write(i,93,float(a))
        sheet1.write(i,94,float(average))
        sheet1.write(i,95,float(words))
        sheet1.write(i,96,float(sentences))
    else:
        sheet1.write(i,89,"N.A")
#MD&A_Def2)
    if mda1+roo+item7a != "":
        payload = {	"type":"text","text":mda1+roo+item7a}
        res = requests.post("https://readability-score.com/ajax.php",data=payload)
        text = str(res.text)
        split = re.split(r'"*',text)
        fkg = split[4].replace(':','')
        f = fkg.replace(',','')
        gf = split[6].replace(':','')
        g = gf.replace(',','')
        cl = split[8].replace(':','')
        c = cl.replace(',','')
        si = split[10].replace(':','')
        s = si.replace(',','')
        ar = split[12].replace(':','')
        a = ar.replace(',','')
        ag = split[15].replace(':','')
        average = ag.replace(',','')
        words = split[27].replace(',','')
        sentences = split[19].replace(',','')
        sheet1.write(i,97,float(f))
        sheet1.write(i,98,float(g))
        sheet1.write(i,99,float(c))
        sheet1.write(i,100,float(s))
        sheet1.write(i,101,float(a))
        sheet1.write(i,102,float(average))
        sheet1.write(i,103,float(words))
        sheet1.write(i,104,float(sentences))	
    else:
        sheet1.write(i,97,"N.A")    	
		
#item8
    for j in range(0,n):
		if sp[j] == "item8":
			item8 = sp[j+1]
		if "item8" not in sp:
			item8 = ""
    if len(item8.split()) == 0:
        item8 = ""
    if item8 != "":
        payload = {	"type":"text","text":item8}
        res = requests.post("https://readability-score.com/ajax.php",data=payload)
        text = str(res.text)
        split = re.split(r'"*',text)
        fkg = split[4].replace(':','')
        f = fkg.replace(',','')
        gf = split[6].replace(':','')
        g = gf.replace(',','')
        cl = split[8].replace(':','')
        c = cl.replace(',','')
        si = split[10].replace(':','')
        s = si.replace(',','')
        ar = split[12].replace(':','')
        a = ar.replace(',','')
        ag = split[15].replace(':','')
        average = ag.replace(',','')
        words = split[27].replace(',','')
        sentences = split[19].replace(',','')
        sheet1.write(i,105,float(f))
        sheet1.write(i,106,float(g))
        sheet1.write(i,107,float(c))
        sheet1.write(i,108,float(s))
        sheet1.write(i,109,float(a))
        sheet1.write(i,110,float(average))
        sheet1.write(i,111,float(words))
        sheet1.write(i,112,float(sentences))
    else:
        sheet1.write(i,105,"N.A")	
#item9
    for j in range(0,n):
		if sp[j] == "item9":
			item9 = sp[j+1]
		if "item9" not in sp:
			item9 = ""
    if len(item9.split()) == 0:
        item9 = ""
    if item9 != "":
        payload = {	"type":"text","text":item9}
        res = requests.post("https://readability-score.com/ajax.php",data=payload)
        text = str(res.text)
        split = re.split(r'"*',text)
        fkg = split[4].replace(':','')
        f = fkg.replace(',','')
        gf = split[6].replace(':','')
        g = gf.replace(',','')
        cl = split[8].replace(':','')
        c = cl.replace(',','')
        si = split[10].replace(':','')
        s = si.replace(',','')
        ar = split[12].replace(':','')
        a = ar.replace(',','')
        ag = split[15].replace(':','')
        average = ag.replace(',','')
        words = split[27].replace(',','')
        sentences = split[19].replace(',','')
        sheet1.write(i,113,float(f))
        sheet1.write(i,114,float(g))
        sheet1.write(i,115,float(c))
        sheet1.write(i,116,float(s))
        sheet1.write(i,117,float(a))
        sheet1.write(i,118,float(average))
        sheet1.write(i,119,float(words))
        sheet1.write(i,120,float(sentences))
    else:
        sheet1.write(i,113,"N.A")	
#item9a
    for j in range(0,n):
		if sp[j] == "item9a":
			item9a = sp[j+1]
		if "item9a" not in sp:
			item9a = ""
    if len(item9a.split()) == 0:
        item9a = ""
    if item9a != "":
        payload = {	"type":"text","text":item9a}
        res = requests.post("https://readability-score.com/ajax.php",data=payload)
        text = str(res.text)
        split = re.split(r'"*',text)
        fkg = split[4].replace(':','')
        f = fkg.replace(',','')
        gf = split[6].replace(':','')
        g = gf.replace(',','')
        cl = split[8].replace(':','')
        c = cl.replace(',','')
        si = split[10].replace(':','')
        s = si.replace(',','')
        ar = split[12].replace(':','')
        a = ar.replace(',','')
        ag = split[15].replace(':','')
        average = ag.replace(',','')
        words = split[27].replace(',','')
        sentences = split[19].replace(',','')
        sheet1.write(i,121,float(f))
        sheet1.write(i,122,float(g))
        sheet1.write(i,123,float(c))
        sheet1.write(i,124,float(s))
        sheet1.write(i,125,float(a))
        sheet1.write(i,126,float(average))
        sheet1.write(i,127,float(words))
        sheet1.write(i,128,float(sentences))

    else:
        sheet1.write(i,121,"N.A")	
#item9b
    for j in range(0,n):
		if sp[j] == "item9b":
			item9b = sp[j+1]
		if "item9b" not in sp:
			item9b = ""
    if len(item9b.split()) == 0:
        item9b = ""
    if item9b != "":
        payload = {	"type":"text","text":item9b}
        res = requests.post("https://readability-score.com/ajax.php",data=payload)
        text = str(res.text)
        split = re.split(r'"*',text)
        fkg = split[4].replace(':','')
        f = fkg.replace(',','')
        gf = split[6].replace(':','')
        g = gf.replace(',','')
        cl = split[8].replace(':','')
        c = cl.replace(',','')
        si = split[10].replace(':','')
        s = si.replace(',','')
        ar = split[12].replace(':','')
        a = ar.replace(',','')
        ag = split[15].replace(':','')
        average = ag.replace(',','')
        words = split[27].replace(',','')
        sentences = split[19].replace(',','')
        sheet1.write(i,129,float(f))
        sheet1.write(i,130,float(g))
        sheet1.write(i,131,float(c))
        sheet1.write(i,132,float(s))
        sheet1.write(i,133,float(a))
        sheet1.write(i,134,float(average))
        sheet1.write(i,135,float(words))
        sheet1.write(i,136,float(sentences))
    else:
        sheet1.write(i,129,"N.A")	
#item10to14
    for j in range(0,n):
		if sp[j] == "item10":
			item10to14 = sp[j+1]
		if "item10" not in sp:
			item10to14 = ""
    if len(item10to14.split()) == 0:
        item10to14 = ""
    if item10to14 != "":
        payload = {	"type":"text","text":item10to14}
        res = requests.post("https://readability-score.com/ajax.php",data=payload)
        text = str(res.text)
        split = re.split(r'"*',text)
        fkg = split[4].replace(':','')
        f = fkg.replace(',','')
        gf = split[6].replace(':','')
        g = gf.replace(',','')
        cl = split[8].replace(':','')
        c = cl.replace(',','')
        si = split[10].replace(':','')
        s = si.replace(',','')
        ar = split[12].replace(':','')
        a = ar.replace(',','')
        ag = split[15].replace(':','')
        average = ag.replace(',','')
        words = split[27].replace(',','')
        sentences = split[19].replace(',','')
        sheet1.write(i,137,float(f))
        sheet1.write(i,138,float(g))
        sheet1.write(i,139,float(c))
        sheet1.write(i,140,float(s))
        sheet1.write(i,141,float(a))
        sheet1.write(i,142,float(average))
        sheet1.write(i,143,float(words))
        sheet1.write(i,144,float(sentences))
    else:
        sheet1.write(i,137,"N.A")	
#other and overall test
    for j in range(0,n):
		if sp[j] == "other":
			other = sp[j+1]
		if "other" not in sp:
			other = ""
    if len(other.split()) == 0:
        other = ""
    if item1+item1a+item1b+item2+item3+item4+item4a+item5+item6+mda1+roo+item7a+item8+item9+item9a+item9b+item10to14+other != "":
        payload = {	"type":"text","text":item1+item1a+item1b+item2+item3+item4+item4a+item5+item6+mda1+roo+item7a+item8+item9+item9a+item9b+item10to14+other}
        res = requests.post("https://readability-score.com/ajax.php",data=payload)
        text = str(res.text)
        split = re.split(r'"*',text)
        fkg = split[4].replace(':','')
        f = fkg.replace(',','')
        gf = split[6].replace(':','')
        g = gf.replace(',','')
        cl = split[8].replace(':','')
        c = cl.replace(',','')
        si = split[10].replace(':','')
        s = si.replace(',','')
        ar = split[12].replace(':','')
        a = ar.replace(',','')
        ag = split[15].replace(':','')
        average = ag.replace(',','')
        words = split[27].replace(',','')
        sentences = split[19].replace(',','')
        sheet1.write(i,145,float(f))
        sheet1.write(i,146,float(g))
        sheet1.write(i,147,float(c))
        sheet1.write(i,148,float(s))
        sheet1.write(i,149,float(a))
        sheet1.write(i,150,float(average))
        sheet1.write(i,151,float(words))
        sheet1.write(i,152,float(sentences))
    else:
        sheet1.write(i,145,"N.A")	

    wb.save('record.xls')
    print "Readabilty scores of ",file[d]," has been recorded at row ",i+1," on record.xls"


print "orange !"
	



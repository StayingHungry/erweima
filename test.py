#coding:utf-8
import qrcode
import os
import win32com.client as win32
urls_file = open("urls.txt")
urls = urls_file.readlines()
word=win32.gencache.EnsureDispatch("Word.Application")
doc=word.Documents.Add()
index =0
for  url in urls:
    img = qrcode.make(url)
    img.save(r"D:\\erweima.png")
    myRange = doc.Range(0,index)
    myRange.InsertBefore('Hello from Python!\n')
    word.Selection.InlineShapes.AddPicture(FileName=r"D:\\erweima.png",LinkToFile= False,SaveWithDocument=True)
    index = index + 1
    # word.Selection.InlineShapes.AddPicture(FileName=r"D:\\erweima.png",LinkToFile= False,SaveWithDocument=True)
doc.SaveAs(r"D:\\test.docx")
doc.Close(True)
word.Application.Quit(-1)

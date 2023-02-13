# docx文件批量转txt文件,传入docx文件夹路径
import os

from win32com import client as wc

def docx2txt(input):
    docx_name_list = os.listdir(input)
    for dn in docx_name_list:
        if not os.path.splitext(dn)[1] == ".docx":  # 筛选文件类型,注意”.“
            continue
        wordapp = wc.Dispatch('Word.Application')
        path1 = os.path.join(input, dn)
        doc = wordapp.Documents.Open(path1)
        output = os.path.splitext(dn)[0]
        output = os.path.join(input, output)
        doc.SaveAs(output, 4)  # 为了让python可以在后续操作中r方式读取txt和不产生乱码，参数为4
        doc.Close()


path1 = r'D:\1'
docx2txt(path1)
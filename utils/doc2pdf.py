# -*- coding: utf-8 -*-
from __future__ import unicode_literals

import os
import sys
from pythoncom import CoInitialize, CoUninitialize
from win32com.client import Dispatch, constants, gencache

ROOT_PATH = os.path.abspath('static')

# 定义文件的存储位置


def fetch_all_files(path):
    """
    获取目录下所有的doc文件
    """
    files=[]
    for dir_path, dir_names, filenames in os.walk(path):
        for file in filenames:
            ext = os.path.splitext(file)[1].lower()
            if(ext=='.docx' or ext=='.doc'):
                fullpath=os.path.join(dir_path,file)
                files.append(fullpath)
    return files

def convert_word_to_pdf(docx_path, pdf_path):
    """
    将word文档转换成pdf文件
    """
    CoInitialize()
    word=gencache.EnsureDispatch('Word.Application')
    word.Visible=0
    word.DisplayAlerts=0
    try:
        doc=word.Documents.Open(docx_path, ReadOnly=1)
        doc.ExportAsFixedFormat(pdf_path, 
            constants.wdExportFormatPDF, 
            Item=constants.wdExportDocumentWithMarkup, 
            CreateBookmarks=constants.wdExportCreateHeadingBookmarks)
    except Exception , e:
        print(e)
    finally:
        word.Quit(constants.wdDoNotSaveChanges)
        CoUninitialize()
def ensure_module():
    gencache.EnsureModule('{00020905-0000-0000-C000-0000000000046}', 0, 8, 0)

def main(simple_name, docpath=None):
    if docpath is None:
        docpath = os.path.join(ROOT_PATH, 'data')
    else:
        docpath = os.path.abspath(docpath)   
    print(docpath)
    inputfilename = simple_name+'.docx'
    outputfilename = simple_name+'.pdf'
    src_path = os.path.join(docpath,inputfilename)
    des_path = os.path.join(docpath, outputfilename)
    print(src_path,des_path)
    if not os.path.exists(src_path):
        print(simple_name+" not exists")
        return
    if not os.path.exists(des_path):
        convert_word_to_pdf(src_path,des_path)
    else:
        print(des_path+' already exists')
    

if __name__=="__main__":
    samplename=''
    docxpath=None
    samplename=sys.argv[1]
    if len(sys.argv) is 3:
        docxpath=sys.argv[2]
    main(samplename,docxpath)


# [call ]
    

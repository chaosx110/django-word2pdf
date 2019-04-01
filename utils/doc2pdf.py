# -*- coding: utf-8 -*-
from __future__ import unicode_literals

import os
import sys
from time import sleep, localtime, strftime
from shutil import move
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

def save_doc(fs, docx_path):
    with open(docx_path) as f:
        f.write(fs)

def ensure_module():
    gencache.EnsureModule('{00020905-0000-0000-C000-0000000000046}', 0, 8, 0)

def make_sure_file_exist(src_path, max_try_times = 10):
    """
    确保源文件存在，尝试次数最大位10，即10s

    Parameters:
    -----------
    src_path: string
        源文件地址
    max_try_times: number
        尝试次数，默认值为20
    """
    times = 0
    while not os.path.exists(src_path):
        print('can not find file : {0}, try {1}'.format(src_path, times+1))
        if times >= max_try_times:
            return False
        sleep(1)
        times +=1
    return True


def main(simple_name, docpath=None):
    if docpath is None:
        docpath = os.path.join(ROOT_PATH, 'data')
    else:
        docpath = os.path.abspath(docpath)   
    inputfilename = simple_name+'.docx'
    outputfilename = simple_name+'.pdf'
    src_path = os.path.join(docpath,inputfilename)
    des_path = os.path.join(docpath, outputfilename)
    bak_path = os.path.join(ROOT_PATH,'backup', outputfilename+strftime("%Y%m%d%H%M%S", localtime()))
    print(src_path,des_path)
    if not make_sure_file_exist(src_path):
        print(simple_name+" not exists")
        return
    if os.path.exists(des_path):
        move(des_path, bak_path)
        print('backup file: {0} success'.format(bak_path))
    convert_word_to_pdf(src_path,des_path)
    

if __name__=="__main__":
    samplename=''
    docxpath=None
    samplename=sys.argv[1]
    if len(sys.argv) is 3:
        docxpath=sys.argv[2]
    main(samplename,docxpath)


# [call ]
    

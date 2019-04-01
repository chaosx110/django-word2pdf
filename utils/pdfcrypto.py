# -*- coding: utf-8 -*-
from __future__ import unicode_literals

from PDFlib.PDFlib import PDFlib
from PDFlib.PDFlib import PDFlibException

def add_watermark(pdf_file_in, pdf_file_out, p_w_picpath_file):
    p = PDFlib()
    p.set_option("license=xxxxx")   #your key
    p.set_option("errorpolicy=return");
                 
    if (p.begin_document(pdf_file_out, "") == -1):
        raise PDFlibException("Error: " + p.get_errmsg())
    p.set_info("Author", "walker");
    p.set_info("Title", "");
    p.set_info("Creator", "walker");
    p.set_info("Subject", "");
    p.set_info("Keywords", "");
    #p.set_info("Producer", "walker");
    #输入文件
    indoc = p.open_pdi_document(pdf_file_in, "");
    if (indoc == -1):
        raise PDFlibException("Error: " + p.get_errmsg())
                 
    endpage = p.pcos_get_number(indoc, "length:pages");
    endpage = int(endpage)
                 
    p_w_picpath = p.load_p_w_picpath("auto", p_w_picpath_file, "")
    if p_w_picpath == -1:
        raise PDFlibException("Error: " + p.get_errmsg())
                 
    for pageno in range(1, endpage+1):
        page = p.open_pdi_page(indoc, pageno, "");
        if (page == -1):
            raise PDFlibException("Error: " + p.get_errmsg())
        p.begin_page_ext(0, 0, "");     #添加一页
                     
        p.fit_pdi_page(page, 0, 0, "adjustpage")
        page_width = p.get_value("pagewidth", 0)    #单位为像素72dpi下像素值
        page_height = p.get_value("pageheight", 0)  #单位为像素72dpi下像素值
                     
        p_w_picpathwidth = p.info_p_w_picpath(p_w_picpath, "p_w_picpathwidth", "");
        p_w_picpathheight = p.info_p_w_picpath(p_w_picpath, "p_w_picpathheight", "");
                     
        margin = 1000   #用于设置水印边距
                     
        optlist_top = "boxsize={" + str(page_width) + " " + str(page_height) + "} "
        optlist_top += "position={" + str(margin/page_width) + " " + str(margin/ page_height) + "} "
        optlist_top += " fitmethod=clip dpi=96"
                     
        optlist_bottom = "boxsize={" + str(page_width) + " " + str(page_height) + "} "
        optlist_bottom += "position={" + str(100 - margin/page_width) + " " + str(100 - margin/ page_height) + "} "
        optlist_bottom += " fitmethod=clip dpi=96"
                     
        p.fit_p_w_picpath(p_w_picpath, 0, 0, optlist_bottom)
        p.fit_p_w_picpath(p_w_picpath, 0, 0, optlist_top)
                     
        p.close_pdi_page(page);
        p.end_page_ext("");
                 
    p.close_p_w_picpath(p_w_picpath)
    p.end_document("")

if __name__ == "__main__":
    # create_watermark('yaoming')
    pdfsigner()
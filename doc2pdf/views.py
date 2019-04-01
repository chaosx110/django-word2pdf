# -*- coding: utf-8 -*-
from __future__ import unicode_literals

import base64
import json
import os
import random
import string
import tempfile

from django.shortcuts import render
from django.http import HttpResponse, JsonResponse, FileResponse
from django.views.decorators.csrf import csrf_exempt
from subprocess import call
from utils import doc2pdf

# Create your views here.


def index(request):
    if 'doc_name' in request.GET:
        doc_name = request.GET['doc_name']
        try:
            doc2pdf.main(doc_name)
        except Exception as err:
            print(err)
        return HttpResponse('Y')

@csrf_exempt
def convert(request):
    '''
    method: post

    params:
        file

    return:
        resp:
            file
            or 403
s    '''
    if request.method == 'POST':
        file_obj = request.FILES.get('file')
        doc_file = os.path.join(doc2pdf.ROOT_PATH, 'converts', file_obj.name)
        pdf_file = os.path.join(
            doc2pdf.ROOT_PATH, 'converts', file_obj.name.rstrip('.docx')+'.pdf')
        with open(doc_file, 'wb') as docx:
            for chunk in file_obj.chunks():
                docx.write(chunk)
        doc2pdf.convert_word_to_pdf(doc_file, pdf_file)
        print(pdf_file)
        pdf =  open(pdf_file, 'rb')
        resp = FileResponse(pdf)
        resp['Content-Type'] = 'application/octet-stream'
        resp['Content-Disposition'] = 'attachment;filename="{}"'.format(pdf.name)
        return resp
    return HttpResponse("N")

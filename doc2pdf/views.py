# -*- coding: utf-8 -*-
from __future__ import unicode_literals

from django.shortcuts import render
from django.http import HttpResponse
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
			return HttpResponse('N')
        return HttpResponse('Y')

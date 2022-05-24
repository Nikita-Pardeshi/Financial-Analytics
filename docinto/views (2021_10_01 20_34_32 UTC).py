# -*- coding: utf-8 -*-
from __future__ import unicode_literals

from django.shortcuts import render
from django.shortcuts import render_to_response

from django.template import RequestContext
# Create your views here.
from smtplib import SMTP
from django.views.decorators.csrf import csrf_exempt
import json
from django.http import HttpResponseRedirect, HttpResponse
from coverpage_2 import merge
from django.core.files.storage import FileSystemStorage



global data
data = dict({
    'name':''
    
})


def page1(request):
    context = RequestContext(request)
    return render(request, 'Page1.html')


def page2(request):
    context = RequestContext(request)
    
    data['name']=request.POST['name']
    print("Company: ",data['name'])
    

    path=merge(data)

    return path
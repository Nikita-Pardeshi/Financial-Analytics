# -*- coding: utf-8 -*-
# Create your models here.

from django import template
from django.contrib.auth.models import Group

register = template.Library()



@register.filter(name ='range2')
def range2(n):
	print("n is ", n)
	print("rg is ", range(n))
	return range(n)
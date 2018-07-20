# -*- coding: utf-8 -*-

from django import forms

	
class DocumentForm(forms.Form):

	docfile = forms.FileField(widget=forms.ClearableFileInput(attrs={'multiple': True}))
	# your_name = forms.CharField(max_length=100)
	# vehicle_platform = forms.CharField(max_length=100)
	# ps_platform = forms.CharField(max_length=100)
	# my = forms.CharField(max_length=100)
	# phase = forms.CharField(max_length=100)
	
class CAEForm(forms.Form):
	docfile = forms.FileField(label='Select a file')
	your_name = forms.CharField(max_length=100)

class NVHForm(forms.Form):
	docfile = forms.FileField(label='Select a file')
	your_name = forms.CharField(max_length=100)
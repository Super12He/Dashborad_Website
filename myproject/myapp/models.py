# -*- coding: utf-8 -*-
from django.db import models

def user_directory_path(instance, filename):
	# file will be uploaded to MEDIA_ROOT/user_<id>/<filename>
    return 'documents'.format(filename)


class Document(models.Model):
    # docfile = models.FileField(upload_to='documents/%Y/%m/%d')
    docfile = models.FileField(upload_to=user_directory_path)

class CAE(models.Model):
    # docfile = models.FileField(upload_to='documents/%Y/%m/%d')
    docfile = models.FileField(upload_to='CAE')	

class NVH(models.Model):
    # docfile = models.FileField(upload_to='documents/%Y/%m/%d')
    docfile = models.FileField(upload_to='NVH')	
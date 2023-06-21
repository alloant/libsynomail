#!/bin/python
# -*- coding: utf-8 -*-

import sys
import os

from getpass import getpass

import logging

from libsynomail.classes import File
from libsynomail.nas_def import init_connection,files_path,move_path,get_info,copy_path,rename_path,convert_office


PASS = getpass()
init_connection('vInd1',PASS)

#files = files_path("/mydrive/00 - Admin")
#for file in files:
#    print(file)

path = "/mydrive/00 - Admin/Patata.odoc"
#print(get_info(path))
#copy_path(path,"/mydrive/00 - Admin/tests/Patata.odoc")
#copy_path(path,"/mydrive/00 - Admin/tests")
#rename_path(path,"potato.odoc")
convert_office("/mydrive/00 - Admin/Patata.docx")


#!/bin/python
# -*- coding: utf-8 -*-

import sys
import os

from getpass import getpass

import logging

import libsynomail.connection as con

PASS = getpass()
con.init_nas(PASS) 

path = '/team-folders/File Sharing/Antonio/Tests'

#con4.nas.convert_office(f"{path4}/test.docx")
#con4.nas.copy(f"{path4}/test.docx",f"{path4}/Despacho")

#con.nas.change_name(f"{path}/Untitled.odoc","patata.odoc")
#con4.nas.change_name(f"{path4}/patata.odoc","Untitled.odoc")

"""
files = con.nas.get_file_list('/team-folders/Folders Ind/Mail')
files4 = con4.nas.get_file_list('/Folders Ind/Mail')


for i in range(len(files)):
    print(files[i]['display_path'])
    print(files4[i]['display_path'])
    print('------------')
"""



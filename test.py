#!/bin/python
# -*- coding: utf-8 -*-

import sys
import os

from getpass import getpass

import logging

from libsynomail.classes import File
import libsynomail.connection as con

PASS = getpass()
con.init_nas(PASS) 

path = '/team-folders/File Sharing/Antonio/Tests/Despacho/Inbox Despacho'

#con4.nas.convert_office(f"{path4}/test.docx")
#con4.nas.copy(f"{path4}/test.docx",f"{path4}/Despacho")

#con.nas.change_name(f"{path}/Untitled.odoc","patata.odoc")
#con4.nas.change_name(f"{path4}/patata.odoc","Untitled.odoc")

#files = con.nas.get_file_list('/team-folders/Folders Ind/Mail')
files = con.nas.get_file_list(path)
fls = []
for i in range(len(files)):
    fls.append(File(files[i]))
    #print(files[i]['display_path'],files[i]['file_id'],files[i]['name'],files[i]['path'],files[i]['permanent_link'],files[i]['type'])
    print(fls[-1].getLinkSheet(),fls[-1].getLinkMessage()) 
    print('------------')



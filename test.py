#!/bin/python
# -*- coding: utf-8 -*-

import sys
import os

from getpass import getpass

import logging

from libsynomail.classes import File
import libsynomail.connection as con

PASS = getpass()
con.init_nas('vInd1',PASS,'')

#path = '/team-folders/Despacho/Inbox Despacho/scc-sg_0101.odoc'
#print(con.nas.get_info(path))

#info = con.nas.get_info("/team-folders/Mail cg/Mail from cg/cg_0011.docx")
#info = con.nas.get_info("/team-folders/File Sharing/Antonio/Tests/Despacho/Inbox Despacho/cg_0011.docx")
#print(info)


file_id = '758152172627081340'
#757845037896743694
#con.nas.move(file_id,"/team-folders/File Sharing/Antonio/Tests/Despacho/Inbox Despacho")
#con.nas.convert_office(fid)

#info = con.nas.get_info("/team-folders/Despacho/ToSend/2682.docx")
con.nas.convert_office("/team-folders/Despacho/ToSend/2682.docx")
print(info)


#con4.nas.convert_office(f"{path4}/test.docx")
#con4.nas.copy(f"{path4}/test.docx",f"{path4}/Despacho")

#con.nas.change_name(f"{path}/Untitled.odoc","patata.odoc")
#con4.nas.change_name(f"{path4}/patata.odoc","Untitled.odoc")

#files = con.nas.get_file_list('/team-folders/Folders Ind/Mail')
#files = con.nas.get_file_list(path)
#fls = []
#for i in range(len(files)):
#    fls.append(File(files[i]))
    #print(files[i]['display_path'],files[i]['file_id'],files[i]['name'],files[i]['path'],files[i]['permanent_link'],files[i]['type'])
#    print(fls[-1].getLinkSheet(),fls[-1].getLinkMessage()) 
#    print('------------')



#!/usr/bin/python
# -*- coding: utf-8 -*-

import os

_ROOT = os.path.abspath(os.path.dirname(__file__))

INV_EXT = {'osheet':'xlsx','odoc':'docx'}
EXT = {'xls':'osheet','xlsx':'osheet','docx':'odoc'}

#SETTINGS = QSettings("synochat","config")

#if not str(SETTINGS.fileName()).endswith('.conf'): # We are on Windows probably
#	SETTINGS = QSettings(QSettings.IniFormat, QSettings.UserScope,'synomail', 'config')


def txt2dict(file):
    DICT = {}
    
    with open(file,mode='r') as inp:
        lines = inp.read().splitlines()

        DICT = {ln.split(':',1)[0]:ln.split(':',1)[1] for ln in lines if ":" in ln}
    
    return DICT


CONFIG = txt2dict("config.txt")
GROUPS = txt2dict("groups.txt")

if 'DEBUG' in CONFIG:
    DEBUG = CONFIG['DEBUG']
else:
    DEBUG = False



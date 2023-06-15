from libsynomail.nas import prome

def init_nas(user,PASS,UI_CONFIG):
    global nas
    nas = prome(user,PASS)
    global CONFIG
    CONFIG = UI_CONFIG
    

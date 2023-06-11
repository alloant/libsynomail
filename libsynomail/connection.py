from libsynomail.nas import prome

def init_nas(user,PASS):
    global nas
    nas = prome(user,PASS)
    

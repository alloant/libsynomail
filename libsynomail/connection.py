from synomail.nas import prome

def init_nas(PASS):
    global nas
    nas = prome(PASS)

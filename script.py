#!/usr/bin/python
# -*- coding: utf8 -*-
import sys
import imp

def check_requirements():
    #check Python version
    res = True
    if sys.version_info[0] != 2:
        print("Cerema Script est uniquement compatible Python 2, vous utilisez Python " + str(sys.version_info[0]))
        return False

    #check instalation lpod
    try:
        imp.find_module('lpod')
    except ImportError:
        print("Vous avez besoin de la librairie lpod pour exécuter Cerema Script")
        return False
    return res

if __name__ == '__main__':
    check_requirements()
    
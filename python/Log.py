#!/usr/bin/python
# -*- coding: UTF-8 -*-

import os
import time
import PySimpleGUI as sg

class Log:
    'Log util'

    # Log info
    @staticmethod
    def info(msg):
        # print (f'\033-----{time.time()}')
        sg.cprint(msg)
        # print (f'\033-----{time.time()}')

    # Log error
    @staticmethod
    def error(msg):
        # print ('\033[1;31;50m')
        sg.cprint(msg)
        # print ('\033[0m')

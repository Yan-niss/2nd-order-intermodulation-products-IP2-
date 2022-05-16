# -*- coding: utf-8 -*-
"""
Created on Fri Jan  8 18:43:17 2021

@author: yannis.nissopoulos
"""
import sys
import visa 
import time
import os
from timeit import default_timer as timer
from datetime import datetime
from multiprocessing import Process

from Freq_Conv import MHz_to_Hz

def get_valid_freq(prompt,choice):
    if choice == 'freq':    
        init = -1
        while init < 0:
            try:
                value = float(input(prompt))
                print('The frequency is set to ',value,' MHz.')
                value = MHz_to_Hz(value)
            except ValueError:
                print('\nThe input is wrong.')
                sys.exit()
            
            if value < 0:
                print('\n---> Frequency cannot be negative. Please enter again. <---')
            else:
                init = 1
    else:
        try:
            value = float(input(prompt))
            print('\nThe power is set to ',value,' dBm.')
        except ValueError:
            print('\nThe input is wrong.')
            sys.exit()
    return value


def process_query(instrument):
    tic = time.perf_counter() # timer
    instrument.write('*OPC')
    ESR_value = 0
    while (ESR_value & 1) == 0:
        ESR_value = int(instrument.query('*ESR?'))
        time.sleep(0.5)
    toc = time.perf_counter()   
    test_duration = toc - tic
#    print(f'\nThe sweep took {toc - tic:0.2f} seconds.')
    return
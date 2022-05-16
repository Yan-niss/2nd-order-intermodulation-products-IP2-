# -*- coding: utf-8 -*-
"""
Created on Wed Jan 13 21:48:59 2021

@author: yannis.nissopoulos
"""

import xlsxwriter
import time
import os
import sys
import math
import errno
import numpy as np

from Freq_Conv import Hz_to_MHz

def data_to_list_IP3(pos, value, freq_LO, race): # put data into list
    
#    global fund_freq, IF_freq, P_IM_low, P_fund_low, P_IM_high, IIP3, IM3, IP2_RF_low, IP2_RF_high, IP2_IF_1, Compr_1dB

    if race == 'i': # zero the list
        row_length = value
        fund_freq = np.zeros(shape=(row_length), dtype=object)
        IF_freq = np.zeros(shape=(row_length), dtype=object)
        P_IM_low = np.zeros(shape=(row_length), dtype=object) 
        P_fund_low = np.zeros(shape=(row_length), dtype=object)
        P_fund_high = np.zeros(shape=(row_length), dtype=object)
        P_IM_high = np.zeros(shape=(row_length), dtype=object)
        IIP3 = np.zeros(shape=(row_length), dtype=object)
        IM3 = np.zeros(shape=(row_length), dtype=object)
        IP2_RF_low = np.zeros(shape=(row_length), dtype=object)
        IP2_RF_high = np.zeros(shape=(row_length), dtype=object)
        IP2_IF_1 = np.zeros(shape=(row_length), dtype=object)
        Compr_1dB = np.zeros(shape=(row_length), dtype=object)
    elif race == 'freq':
        fund_freq[pos] = int(Hz_to_MHz(value)) # first column is fundamental frequency
        IF_freq[pos] = int(Hz_to_MHz(freq_LO - value))
#        print('Fundamental frequency:    ', fund_freq)
#        print('IF frequency:             ', IF_freq)
    elif race == 'IM_low':
        P_IM_low[pos] = value # for power columns
        print('Low IM product:           ', P_IM_low)
    elif race == 'fund_low':
        P_fund_low[pos] = value
        print('Low Fundamental product:  ', P_fund_low)
    elif race == 'fund_high':
        P_fund_high[pos] = value
        print('High Fundamental product: ', P_fund_high)
    elif race == 'IM_high':
        P_IM_high[pos] = value
        print('High IM product:          ', P_IM_high)
#        print('')
    # IP2
    elif race == 'IP2_RF_low':
        IP2_RF_low[pos] = value
        print('Low IP2 product:          ', IP2_RF_low)
    elif race == 'IP2_RF_high':
        IP2_RF_high[pos] = value
        print('High IP2 product:         ', IP2_RF_high)
    elif race == 'IP2_IF_1':
        IP2_IF_1[pos] = value
        print('IP2_IF_1 product:         ', IP2_IF_1)
    elif race == 'IP2_IF_2':
        IP2_IF_2[pos] = value
        print('IP2_IF_2 product:         ', IP2_IF_2)
    elif race == 'IP2_IF_3':
        IP2_IF_3[pos] = value
        print('IP2_IF_3 product:         ', IP2_IF_3)
    elif race == 'IP2_IF_4':
        IP2_IF_4[pos] = value
        print('IP2_IF_4 product:         ', IP2_IF_4)
        
    # 1dB Compression Point
    elif race == '1dB':
        Compr_1dB[pos] = value
        print('1dB Compression Point:    ', Compr_1dB)
    else:
        sys.exit()
        
    return (fund_freq, IF_freq, P_IM_low, P_fund_low, P_fund_high, P_IM_high, Compr_1dB)



def data_to_list(pos, value, freq_LO, choice): # put data into list
    
    global fund_freq, IF_freq, power

    if choice == 'i': # zero the list
        row_length = value
        fund_freq = np.zeros(shape=(row_length), dtype=object)
        IF_freq = np.zeros(shape=(row_length), dtype=object)
        power = np.zeros(shape=(row_length), dtype=object)
    elif choice == 'freq':
        fund_freq[pos] = int(Hz_to_MHz(value)) # first column is fundamental frequency
        IF_freq[pos] = int(Hz_to_MHz(freq_LO - value))
#        print('Fundamental frequency:    ', fund_freq)
        print('IF frequency:     ', IF_freq)        
    elif choice == 'IP':
        power[pos] = float("{:.2f}".format(value))
#        print('power = ',power)
    else:
        sys.exit()
        
    return (fund_freq, IF_freq, power[pos])
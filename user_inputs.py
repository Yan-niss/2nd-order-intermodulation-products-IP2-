import pyvisa as visa
import time
import os
import sys
import math
import errno
import numpy as np
from datetime import datetime
import configparser
#from configparser import ConfigParser

from Valid_Input import get_valid_freq

def User_Inputs():
    
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
        
    config = configparser.ConfigParser(allow_no_value=True)
    config.read('IP3_config.ini')
    config.set('SigGen', '; This is a comment.', None)                        
    
    # Enter non-compensated for cable loss values
    freq_range_start = config.getfloat('SigGen', 'start_freq') # in Hz
    freq_range_stop = config.getfloat('SigGen', 'stop_freq') # the last value is not executed, so put one more!
    freq_range_step = config.getfloat('SigGen', 'freq_step')
    freq_Spacing_IP3 = config.getfloat('SigGen', 'freq_Spacing') # usually 2 or 5 MHz
    freq_Spacing_IP2 = 154E6 # freq_spacing for IP2
    freq_RF1 = freq_range_start
    freq_RF2 = freq_RF1 + freq_Spacing_IP3
    Pin_DUT_1 = config.getfloat('SigGen', 'Pin_DUT_1') # it is the input power at DUTs input in dBm - both RF generators have the same power (adjustmennts for cable loss performed locally)
    Pin_DUT_2 = config.getfloat('SigGen', 'Pin_DUT_2') # in dBm
    freq_LO = config.getfloat('SigGen', 'LO_Freq') # in Hz
    Pow_LO = config.getfloat('SigGen', 'LO_Pow') # in dBm

    freq_range_start = int(freq_range_start)
    freq_range_stop = int(freq_range_stop)
    freq_range_step = int(freq_range_step)
    
    FolderName = config.get('Output_dir', 'Output_dir')
    dir_path = os.path.dirname(os.path.abspath(__file__))
    path = dir_path + "/" + FolderName
    if not os.path.exists(path):
        os.makedirs(path)
    os.chdir(path) # change path to input path
    
    return (freq_RF1, freq_RF2, Pin_DUT_1, Pin_DUT_2, freq_LO, Pow_LO, freq_range_start, freq_range_stop, freq_range_step, freq_Spacing_IP3, freq_Spacing_IP2)


def Cable_Loss():
    
    global Cable_Loss_RF1, Cable_Loss_RF2, Cable_Loss_IF2SpecAn, Cable_Loss_LO, Pigts_loss_RF, Pigts_loss_IF, Pigts_loss_LO, Cable_Loss_IF2PowMeter
    
    # Initializing the arrays
    Cable_Loss_RF1 = np.zeros(shape=(14), dtype=object) # manually enter the number of columns
    Cable_Loss_RF2 = np.zeros(shape=(14), dtype=object) 
    Cable_Loss_IF2SpecAn = np.zeros(shape=(14), dtype=object) 
    Pigts_loss_RF = np.zeros(shape=(14), dtype=object)
    Pigts_loss_IF = np.zeros(shape=(14), dtype=object)
    Cable_Loss_IF2PowMeter = np.zeros(shape=(14), dtype=object)
    
    Cable_Loss_RF1 = [5.55, 5.55, 5.11, 5.11, 5.2, 5.2, 5.25, 5.25, 5.46, 5.46, 5.58, 5.58, 5.66, 5.66] # cable loss SMU: short cable + circulator #1 + combiner + medium cable  
#    Cable_Loss_RF1 = [2.04, 2.04, 1.57, 1.57, 1.64, 1.64, 1.66, 1.66, 5.46, 1.86, 1.86, 1.88, 1.61, 1.61] # cable loss SMU: short cable + medium cable  
    Cable_Loss_RF2 = [5.64, 5.64, 5.23, 5.23, 5.43, 5.43, 5.42, 5.42, 5.63, 5.63, 5.67, 5.67, 6.29, 6.29] # cable loss SMIQ: short cable + circulator #1 + combiner + medium cable
    Cable_Loss_IF2SpecAn = [7.0, 7.0, 7.1, 7.1, 7.1, 7.39, 7.39, 7.39, 7.65, 7.65, 7.65, 7.85, 7.85, 7.85] # cable loss: long cable + 6dB 3-way splitter + short cable
    Cable_Loss_IF2PowMeter = [12.68, 12.68, 12.9, 12.9, 12.9, 13, 13, 13, 13.1, 13.1, 13.1, 13.2, 13.2, 13.2] # cable loss: long cable + 6dB 3-way splitter + 6dB attenuator
#    Cable_Loss_LO = 2.2 # cable loss @ 2.4 GHz
    Cable_Loss_LO = 0 # no external LO
    Pigts_loss_RF = [0.16, 0.16, 0.16, 0.16, 0.16, 0.16, 0.16, 0.16, 0.16, 0.19, 0.19, 0.19, 0.19, 0.19] # pigtail loss RF
    Pigts_loss_IF = [0.06, 0.06, 0.06, 0.06, 0.06, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.12, 0.12, 0.12] # pigtail loss IF
#    Pigts_loss_LO = 0.2
    Pigts_loss_LO = 0 # no external LO
    
    return (Cable_Loss_LO, Pigts_loss_LO, Cable_Loss_RF1, Cable_Loss_RF2, Cable_Loss_IF2SpecAn, Pigts_loss_RF, Pigts_loss_IF, Cable_Loss_IF2PowMeter)
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
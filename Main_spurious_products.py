import pyvisa as visa
import xlsxwriter
import time
import os
import sys
import math
import errno
from timeit import default_timer as timer
from datetime import datetime
from multiprocessing import Process
from excel_plot import excel_plot

from User_Inputs import User_Inputs
from Initialization import Initialization
from Instr_Setup import SigGen_Setup
from Instr_Setup import SpecAn_Setup
from IP3_meas import IP3_meas_PeakZoomIn
from IP2_meas import IP2_meas_PeakZoomIn
from Compres_1dB import Compres_1dB



# ==========================================================================================================================

if __name__=='__main__':
    
#    print('\n############################################################################################')
#    print('##############################         PROGRAM STARTS        #################################')
#    print('##############################################################################################\n')
    
    total_tic = time.perf_counter() # timer
    
    IP3_meas_PeakZoomIn()
           
    total_toc = time.perf_counter()   
    print(f'\n\n\n\n(Total test time: {total_toc - total_tic:0.2f} secs)')

# -*- coding: utf-8 -*-
"""
Created on Thu Jan 14 11:45:38 2021

@author: yannis.nissopoulos
"""
import pyvisa as visa
import time
import os
import sys
import errno
import numpy as np
import xlsxwriter

#from data_to_list import data_to_list
from data_to_list import data_to_list_IP3
from Valid_Input import process_query
from User_Inputs import User_Inputs
from User_Inputs import Cable_Loss
from Freq_Conv import Hz_to_MHz
from Freq_Conv import MHz_to_Hz
#from Calculations import IP3_Products
from Calculations import Mixer_IP2_Products
from Initialization import Initialization
from Instr_Setup import SigGen_Setup
from Instr_Setup import SpecAn_Setup
#from Calculations import Gain_Calc
from Calculations import IP2_Calc
from excel_plot import excel_plot

def IP2_meas_PeakZoomIn():

#    SigGen_RF_1, SigGen_RF_2, SigGen_LO, SpecAn, PowMeter = Initialization()
    SigGen_RF_1, SigGen_RF_2, SpecAn, PowMeter = Initialization()
    freq_RF1, freq_RF2, Pin_DUT_1, Pin_DUT_2, freq_LO, Pow_LO, RF_freq_start, RF_freq_stop, RF_freq_step, freq_Spacing_IP3, freq_Spacing_IP2 = User_Inputs()
    Cable_Loss_LO, Pigts_loss_LO, Cable_Loss_RF1, Cable_Loss_RF2, Cable_Loss_IF2SpecAn, Pigts_loss_RF, Pigts_loss_IF, Cable_Loss_IF2PowMeter = Cable_Loss()
    
    Pow_LO_comp = Pow_LO + Cable_Loss_LO + Pigts_loss_LO # LO power compensation for cable loss
#    SigGen_Setup(SigGen_LO, str(freq_LO), str(Pow_LO_comp))
        
    row_length = int((RF_freq_stop - RF_freq_start)/RF_freq_step) # if you enable this the add these arguments (RF_freq_start, RF_freq_stop, RF_freq_step)
    data_to_list(0, row_length, freq_LO, 'i') # initializatioin
#    Gain_Calc(0, row_length, 0, 0, 0, 0, SigGen_RF_1, SpecAn, 'i')
    IP2_Calc(0, row_length, 0, 0, 'i')


    print('\n\n##############    IP2    ##############')    
    count = 1
    pos = 0 # position
    for RF_freq in range(RF_freq_start, RF_freq_stop, RF_freq_step): # RF_freq is the frequencies that the measurements for IP3, IP2, etc. are wanted
        
        # Cable Loss compensation for RF cables
        Cable_Loss_LO, Pigts_loss_LO, Cable_Loss_RF1, Cable_Loss_RF2, Cable_Loss_IF2SpecAn, Pigts_loss_RF, Pigts_loss_IF, Cable_Loss_IF2PowMeter = Cable_Loss()
        Pout_SigGen1 = Pin_DUT_1 + Cable_Loss_RF1[pos] + Pigts_loss_RF[pos] # power level that the SigGen#1 has been set 
#        print('Pout_SigGen1', Pout_SigGen1)
        Pout_SigGen2 = Pin_DUT_2 + Cable_Loss_RF2[pos] + Pigts_loss_RF[pos] # power level that the SigGen#1 has been set

#        Gain = Gain_Calc(pos, row_length, RF_freq, Pout_SigGen1, Cable_Loss_IF2SpecAn, Pigts_loss_IF, SigGen_RF_1, SpecAn, 'G')
#        print('Gain (dB): ', Gain)
        
        SpecAn_Setup(SpecAn, 4000, 100, 1) # SPAN, RBW, AVG
        SigGen_Setup(SigGen_RF_1, RF_freq, Pout_SigGen1)
        SigGen_Setup(SigGen_RF_2, RF_freq + freq_Spacing_IP3, Pout_SigGen2)
        print('\n\n            Iteration #', count)
        print('------------------------------------')
        count = count + 1
        fund_f1, fund_f3, IP2_IF_1, IP2_IF_2, IP2_IF_3, IP2_IF_4 = Mixer_IP2_Products(int(RF_freq),float(freq_Spacing_IP2),float(freq_LO))
        
        print(f'\n\nThe current fundamental frequency (freq_RF1) is {Hz_to_MHz(RF_freq):0.0f} MHz.')
        print('-----------------------------------------------------------')
        data_to_list(pos, RF_freq, freq_LO, 'freq') 

        for center_freq in [fund_f1, IP2_IF_1]:
#            print(f'The center freq in {center_freq}')
            if center_freq == fund_f1:
                print('\nstart fund\n')
                marker_title = 'fund_f1_low'
                SpecAn.write('DISP:TRAC:Y:RLEV 0') # set the ref. level for lower signal level
                SpecAn.write('FREQ:CENT '+ str(center_freq) +'MHz') # sets center freq  
                SpecAn.write('CALC1:MARK1:X '+ str(center_freq) +'MHz') # Positions marker to frequency
                SpecAn.write('CALC:MARK1:MAX')
                time.sleep(4)
                SpecAn.write('INIT') # Starts the measurement and waits for the end of the sweeps
                process_query(SpecAn) # checks when the sweep finishes
#                SpecAn.write('CALC:MARK1:MAX')
                Marker_Freq = Hz_to_MHz(float(SpecAn.query('CALC:MARK1:X?'))) # Outputs the measured value of the marker in screen A
                Marker_Pow = float(SpecAn.query('CALC:MARK1:Y?')) # Outputs the measured value of marker 2 in screen A
                IF_fund_low = Marker_Pow + Cable_Loss_IF2SpecAn[pos] + Pigts_loss_IF[pos] # IF power compensation for various cable loss
                IF_fund_low = float("{:.2f}".format(IF_fund_low))
                fund_freq, IF_freq, IF_fund_low = data_to_list(pos, IF_fund_low, freq_LO, 'IP')
                print('Low IF Fundamental product:', IF_fund_low)
#                time.sleep(5)
#            elif center_freq == fund_f3:
#                marker_title = 'fund_f3_high'
#                SpecAn.write('DISP:TRAC:Y:RLEV -30') # set the ref. level for lower signal level
#                SpecAn.write('FREQ:CENT '+ str(center_freq) +'MHz') # sets center freq  
#                SpecAn.write('CALC1:MARK1:X '+ str(center_freq) +'MHz') # Positions marker to frequency
#                SpecAn.write('INIT') # Starts the measurement and waits for the end of the sweeps
#                process_query(SpecAn) # checks when the sweep finishes
#                Marker_Freq = Hz_to_MHz(float(SpecAn.query('CALC:MARK1:X?'))) # Outputs the measured value of the marker in screen A
#                Marker_Pow = float(SpecAn.query('CALC:MARK1:Y?')) # Outputs the measured value of marker 2 in screen A
#                IF_fund_high = Marker_Pow + Cable_Loss_IF2SpecAn[pos] + Pigts_loss_IF[pos] # IF power compensation for various cable loss
#                IF_fund_high = float("{:.2f}".format(IF_fund_high))
#                fund_freq, IF_freq, IP2_fund_high = data_to_list(pos, IF_fund_high, freq_LO, 'IP')
#                print('High Fundamental product:   ', P_fund_low)
#            elif center_freq == IP2_IF_1:
#                print('\nstart IP2\n')
#                marker_title = 'IP2_IF_1'
#                SpecAn.write('DISP:TRAC:Y:RLEV -40') # set the ref. level for lower signal level
#                SpecAn.write('FREQ:CENT '+ str(center_freq) +'MHz') # sets center freq  
#                SpecAn.write('CALC1:MARK1:X '+ str(center_freq) +'MHz') # Positions marker to frequency
#                time.sleep(3)
#                SpecAn.write('INIT') # Starts the measurement and waits for the end of the sweeps
#                process_query(SpecAn) # checks when the sweep finishes
#                SpecAn.write('CALC:MARK1:MAX')
#                Marker_Freq = Hz_to_MHz(float(SpecAn.query('CALC:MARK1:X?'))) # Outputs the measured value of the marker in screen A
#                Marker_Pow = float(SpecAn.query('CALC:MARK1:Y?')) # Outputs the measured value of marker 2 in screen A
#                IP2_1 = Marker_Pow + Cable_Loss_IF2SpecAn[pos] + Pigts_loss_IF[pos] # IF power compensation for various cable loss
#                IP2_1 = float("{:.2f}".format(IP2_1))
#                fund_freq, IF_freq, IP2_IF_1 = data_to_list(pos, IP2_1, freq_LO, 'IP')
#                print('Low IP2 product:           ', IP2_IF_1)
#                time.sleep(5)
#            elif center_freq == IP2_IF_2:
#                marker_title = 'IP2_IF_2'
#                SpecAn.write('DISP:TRAC:Y:RLEV -10') # set the ref. level for lower signal level
#                SpecAn.write('FREQ:CENT '+ str(center_freq) +'MHz') # sets center freq  
#                SpecAn.write('CALC1:MARK1:X '+ str(center_freq) +'MHz') # Positions marker to frequency
#                SpecAn.write('INIT') # Starts the measurement and waits for the end of the sweeps
#                process_query(SpecAn) # checks when the sweep finishes
#                Marker_Freq = Hz_to_MHz(float(SpecAn.query('CALC:MARK1:X?'))) # Outputs the measured value of the marker in screen A
#                Marker_Pow = float(SpecAn.query('CALC:MARK1:Y?')) # Outputs the measured value of marker 2 in screen A
#                IP2_2 = Marker_Pow + Cable_Loss_IF2SpecAn[pos] + Pigts_loss_IF[pos] # IF power compensation for various cable loss
#                IP2_2 = float("{:.2f}".format(IP2_2))
#                fund_freq, IF_freq, IP2_IF_2 = data_to_list(pos, IP2_2, freq_LO, 'IP')
#            elif center_freq == IP2_IF_3:
#                marker_title = 'IP2_IF_3'
#                SpecAn.write('DISP:TRAC:Y:RLEV -10') # set the ref. level for lower signal level
#                SpecAn.write('FREQ:CENT '+ str(center_freq) +'MHz') # sets center freq  
#                SpecAn.write('CALC1:MARK1:X '+ str(center_freq) +'MHz') # Positions marker to frequency
#                SpecAn.write('INIT') # Starts the measurement and waits for the end of the sweeps
#                process_query(SpecAn) # checks when the sweep finishes
#                Marker_Freq = Hz_to_MHz(float(SpecAn.query('CALC:MARK1:X?'))) # Outputs the measured value of the marker in screen A
#                Marker_Pow = float(SpecAn.query('CALC:MARK1:Y?')) # Outputs the measured value of marker 2 in screen A
#                IP2_3 = Marker_Pow + Cable_Loss_IF2SpecAn[pos] + Pigts_loss_IF[pos] # IF power compensation for various cable loss
#                IP2_3 = float("{:.2f}".format(IP2_3))
#                fund_freq, IF_freq, IP2_IF_3 = data_to_list(pos, IP2_3, freq_LO, 'IP')
#            elif center_freq == IP2_IF_4:
#                marker_title = 'IP2_IF_4'
#                SpecAn.write('DISP:TRAC:Y:RLEV -10') # set the ref. level for lower signal level
#                SpecAn.write('FREQ:CENT '+ str(center_freq) +'MHz') # sets center freq  
#                SpecAn.write('CALC1:MARK1:X '+ str(center_freq) +'MHz') # Positions marker to frequency
#                SpecAn.write('INIT') # Starts the measurement and waits for the end of the sweeps
#                process_query(SpecAn) # checks when the sweep finishes
#                Marker_Freq = Hz_to_MHz(float(SpecAn.query('CALC:MARK1:X?'))) # Outputs the measured value of the marker in screen A
#                Marker_Pow = float(SpecAn.query('CALC:MARK1:Y?')) # Outputs the measured value of marker 2 in screen A
#                IP2_4 = Marker_Pow + Cable_Loss_IF2SpecAn[pos] + Pigts_loss_IF[pos] # IF power compensation for various cable loss
#                IP2_4 = float("{:.2f}".format(IP2_4))
#                fund_freq, IF_freq, IP2_IF_4 = data_to_list(pos, IP2_4, freq_LO, 'IP')                
            else:
                sys.exit()
#            print(f'\nThe frequency and power for the {marker_title} is: {Marker_Freq:0.0f} MHz & {Marker_Pow:0.2f} dBm.')
                
        IM_Low, IIP2 = IP2_Calc(pos, row_length, IP2_IF_1, IF_fund_low, 'calc_IP2') 

#        excel_plot(fund_freq, IF_freq, Gain, P_IM_low, P_fund_low, P_fund_high, P_IM_high, Delta_P_Low, Delta_P_High, OIP3_Low, OIP3_High, Cable_Loss_RF1, Cable_Loss_RF2, Cable_Loss_IF2SpecAn, Cable_Loss_IF2PowMeter, Pigts_loss_RF, Pigts_loss_IF)
    
        print('\n\n= = = = = = = = = = = = = = = = = = = = = = = ')
        pos = pos + 1
        
    return

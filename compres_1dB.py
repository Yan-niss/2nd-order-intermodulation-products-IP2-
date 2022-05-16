# -*- coding: utf-8 -*-
"""
Created on Sun Jan 17 20:05:26 2021

@author: yannis.nissopoulos
"""

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

#from data_to_list import data_to_list
from data_to_list import data_to_list_IP3
from Valid_Input import process_query
from User_Inputs import User_Inputs
from User_Inputs import Cable_Loss
from Freq_Conv import Hz_to_MHz
from Freq_Conv import MHz_to_Hz
from Initialization import Initialization
from Instr_Setup import SigGen_Setup
from Instr_Setup import SpecAn_Setup
from Calculations import Gain_Calc
from excel_plot import excel_plot

def Compres_1dB():
    
#    SigGen_RF_1, SigGen_RF_2, SigGen_LO, SpecAn, PowMeter = Initialization()
    SigGen_RF_1, SigGen_RF_2, SpecAn, PowMeter = Initialization()
    freq_RF1, freq_RF2, Pin_DUT_1, Pin_DUT_2, freq_LO, Pow_LO, RF_freq_start, RF_freq_stop, RF_freq_step, freq_Spacing_IP3, freq_Spacing_IP2 = User_Inputs()
    Cable_Loss_LO, Pigts_loss_LO, Cable_Loss_RF1, Cable_Loss_RF2, Cable_Loss_IF2SpecAn, Pigts_loss_RF, Pigts_loss_IF, Cable_Loss_IF2PowMeter = Cable_Loss()
    SigGen_Setup(SigGen_RF_2, 1E6, -70) # not part of the setup but I need to set the ref clock to ext.
    SigGen_RF_2.write('OUTP OFF')
    Pow_LO_comp = Pow_LO + Cable_Loss_LO + Pigts_loss_LO # LO power compensation for cable loss
#    SigGen_Setup(SigGen_LO, str(freq_LO), str(Pow_LO_comp))
    
    row_length = int((RF_freq_stop - RF_freq_start)/RF_freq_step) # if you enable this the add these arguments (RF_freq_start, RF_freq_stop, RF_freq_step)
    data_to_list(0, row_length, freq_LO, 'i') # initializatioin
    Gain = Gain_Calc(0, row_length, 0, 0, 0, 0, SigGen_RF_1, SpecAn, 'i')
   
    print('\n\n##############    1dB Compression Point    ##############\n\n')    
    count = 1
    
    pos = 0 # position
    for RF_freq in range(RF_freq_start, RF_freq_stop, RF_freq_step): # RF_freq is the frequencies that the measurements for IP3, IP2, etc. are wanted
        No_points = 0 # number of point of the nested loop
        RF_FREQ, IF_FREQ, OOI = data_to_list(pos, RF_freq, freq_LO, 'freq') # OOI = Out Of Interest
        # Cable Loss compensation for RF cables
        freq_RF1, freq_RF2, Pin_DUT_1, Pin_DUT_2, freq_LO, Pow_LO, RF_freq_start, RF_freq_stop, RF_freq_step, freq_Spacing_IP3, freq_Spacing_IP2 = User_Inputs()
        Cable_Loss_LO, Pigts_loss_LO, Cable_Loss_RF1, Cable_Loss_RF2, Cable_Loss_IF2SpecAn, Pigts_loss_RF, Pigts_loss_IF, Cable_Loss_IF2PowMeter = Cable_Loss()
        Pout_SigGen1 = Pin_DUT_1 + Cable_Loss_RF1[pos] + Pigts_loss_RF[pos] # power level that the SigGen#1 has been set
        Gain = Gain_Calc(pos, row_length, RF_freq, Pout_SigGen1, Cable_Loss_IF2SpecAn, Pigts_loss_IF, SigGen_RF_1, SpecAn, 'G')
        print('Gain (dB): ', Gain)
#        SigGen_Setup(SigGen_RF_1, RF_freq, Pout_SigGen1)
#        time.sleep(1.5)
#        SpecAn_Setup(SpecAn, 4000, 100, 1)
        
        print('\n\n               Iteration #', count)
        print('----------------------------------------------------------------')
        count = count + 1
        step = 0.1 # step to increase SigGens input power in every loop
        diff = 0 # difference between theoritical value and measured value - here initial value
        low_limit = 0.9
        high_limit = 1.1
        while diff<low_limit or diff>high_limit:
            print(No_points,'.')
            print(f'Signal Generators output is: {Pout_SigGen1:0.2f}dBm.\n')
            print(f'The power at DUTs RF input is: {Pin_DUT_1:0.2f}dBm.\n')
            SigGen_Setup(SigGen_RF_1, RF_freq, Pout_SigGen1)
            IF_Freq = freq_LO - RF_freq # the IF frequency wherethe 1dB point will be found 
            SpecAn.write('DISP:TRAC:Y:RLEV 10') # set the ref. level for lower signal level
            SpecAn.write('FREQ:CENT '+ str(IF_Freq) +'MHz') # sets center freq  
            SpecAn.write('CALC1:MARK1:X '+ str(IF_Freq) +'MHz') # Positions marker to frequency
            SpecAn.write('INIT') # Starts the measurement and waits for the end of the sweeps
            process_query(SpecAn) # checks when the sweep finishes
            SpecAn.write('CALC:MARK1:MAX') # Switches the search limit function on for screen A                           
            Marker_Freq = Hz_to_MHz(float(SpecAn.query('CALC:MARK1:X?'))) # Outputs the measured value
            Marker_Pow = float(SpecAn.query('CALC:MARK1:Y?')) # Outputs the measured power of SpecAn
            Pout_DUT = Marker_Pow + Cable_Loss_IF2SpecAn[pos] + Pigts_loss_IF[pos] # power at DUT's IF output
            Pout_DUT = float("{:.2f}".format(Pout_DUT))
            print(f'DUT Output Power = {Pout_DUT:0.2f}dBm.')
#            OOI1, OOI2, Array_Pout_DUT = data_to_list(No_points, Pout_DUT, freq_LO, 'IP')
            
            Theoritical_Output_Response = Pin_DUT_1 + Gain[pos]
#            print(type(Theoritical_Output_Response))
            print('Calculated Theoritical_Output_Response = ', float("{:.2f}".format(Theoritical_Output_Response)))
#            OOI1, OOI2, Array_1dB = data_to_list(No_points, Theoritical_Output_Response, freq_LO, 'IP')  
            
            diff = abs(Pout_DUT - Theoritical_Output_Response)
            print('\nThe difference is: ', float("{:.2f}".format(diff)))
#            OOI1, OOI2, Array_diff = data_to_list(No_points, diff, freq_LO, 'IP')
            
#            data_to_list(pos, Pin_DUT_1, freq_LO, 'IP')
            
            Pout_SigGen1 = Pout_SigGen1 + step
            Pin_DUT_1 = Pin_DUT_1 + step          
            No_points = No_points + 1

            print('***********************************************\n')
            # DUT's Safety
            if Pin_DUT_1 >= 17.8: # max power at DUT is 18dBm
                print('\n\nTHE INPUT POWER REACHED THE MAXIMUM POWER RATING OF THE DUT WITHOUT CALCULATING THE 1dB COMPREASSION POINT!')
                SigGen_RF_1.write('OUTP OFF') # turns off the SigGen
                break
        Pin_DUT_1 = Pin_DUT_1 - step # Power at DUT's input
        OOI1, OOI2, Compr_1dB = data_to_list(pos, Pin_DUT_1, freq_LO, 'IP')
        print('\n\nThe 1dB (input) Compression Point is: ', Compr_1dB, 'dBm.')
        
        print('\n\n= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = ')
        pos = pos + 1
    SigGen_RF_1.write('OUTP OFF')
    
    ##### Pass data to Excel #####
    workbook = xlsxwriter.Workbook('1dB_Compression_Point_22oC.xlsx')
    Compression_tab = workbook.add_worksheet('1dB_Compression_Point')
    Compression_tab.set_tab_color('green')
    Compression_Details_tab = workbook.add_worksheet('1dB_Compression_Details')
    Compression_Details_tab.set_tab_color('green')        
    bold = workbook.add_format({'bold': 1})
    
    CompressionP_heads = ['RF Frequency (MHz)', 'IF Frequency (MHz)', '1dB Input Compression Point (dBm)']           

    Compression_tab.write_row('A1', CompressionP_heads, bold) # the first cell where the headings will be placed
    Compression_tab.write_column('A2', RF_FREQ) # the cell where the data will be placed & the data column
    Compression_tab.write_column('B2', IF_FREQ)
    Compression_tab.write_column('C2', Compr_1dB)
    #######################################################################
    #
    # Create a scatter chart sub-type with straight lines and no markers.
    #
    chart1 = workbook.add_chart({'type': 'scatter',
                                 'subtype': 'straight'})
    
    # Configure the first series.
    chart1.add_series({
        'name':       '1dB Point (input)',
        'categories': '=1dB_Compression_Point!$A$2:$A$15',
        'values':     '=1dB_Compression_Point!$C$2:$C$15',
    })
    
    # Add a chart title and some axis labels.
    chart1.set_title ({'name': '1dB Compression Point (22°C)'})
    chart1.set_x_axis({'name': 'Frequency (MHz)'})
    chart1.set_y_axis({'name': '1dB Compression Point (dBm)'})
    
    # Set an Excel chart style.
    chart1.set_style(13)
    
    # Insert the chart into the worksheet (with an offset).
    Compression_tab.insert_chart('D2', chart1, {'x_offset': 0, 'y_offset': 0})
    
    chart1 = workbook.add_chart({'type': 'scatter',
                                 'subtype': 'straight'
    })
    
#    # Configure second series.
#    chart2.add_series({
#        'name':       '=Sheet1!$R$1',
#        'categories': '=Sheet1!$B$2:$B$15',
#        'values':     '=Sheet1!$R$2:$R$15',
#    })
#    
#    # Add a chart title and some axis labels.
#    chart2.set_title ({'name': '1dB Compression Point'})
#    chart2.set_x_axis({'name': 'IIP3 (dBm)'})
#    chart2.set_y_axis({'name': 'OIP3 (dBm)'})
#    
#    # Set an Excel chart style.
#    chart2.set_style(13)
#    
#    # Insert the chart into the worksheet (with an offset).
#    worksheet.insert_chart('R18', chart2, {'x_offset': 0, 'y_offset': 0})
    
    workbook.close()
    
    return

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
import configparser

from data_to_list import data_to_list
from datetime import datetime
from Valid_Input import process_query
from User_Inputs import User_Inputs
from User_Inputs import Cable_Loss
from Freq_Conv import Hz_to_MHz
from Freq_Conv import MHz_to_Hz
from Calculations import IP3_Products
from Initialization import Initialization
from Instr_Setup import SigGen_Setup
from Instr_Setup import SpecAn_Setup
from Calculations import Gain_Calc
from Calculations import IP3_Calc
from excel_plot import excel_plot



def IP3_meas_PeakZoomIn():
    
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    config = configparser.ConfigParser(allow_no_value=True)
    config.read('IP3_config.ini')
    config.set('SpecAn', '; This is a comment.', None)
    
#    SigGen_RF_1, SigGen_RF_2, SigGen_LO, SpecAn, PowMeter = Initialization()
    SigGen_RF_1, SigGen_RF_2, SpecAn, PowMeter = Initialization()
    freq_RF1, freq_RF2, Pow_RF1, Pow_RF2, freq_LO, Pow_LO, RF_freq_start, RF_freq_stop, RF_freq_step, freq_Spacing_IP3, freq_Spacing_IP2 = User_Inputs()
    Cable_Loss_LO, Pigts_loss_LO, Cable_Loss_RF1, Cable_Loss_RF2, Cable_Loss_IF2SpecAn, Pigts_loss_RF, Pigts_loss_IF, Cable_Loss_IF2PowMeter = Cable_Loss()
    
    Pow_LO_comp = Pow_LO + Cable_Loss_LO + Pigts_loss_LO # LO power compensation for cable loss
#    SigGen_Setup(SigGen_LO, str(freq_LO), str(Pow_LO_comp))
        
    row_length = int((RF_freq_stop - RF_freq_start)/RF_freq_step) # if you enable this the add these arguments (RF_freq_start, RF_freq_stop, RF_freq_step)
#    data_to_list_IP3(0, row_length, freq_LO, 'i') # initializatioin
    data_to_list(0, row_length, freq_LO, 'i')
    Gain_Calc(0, row_length, 0, 0, 0, 0, SigGen_RF_1, SpecAn, 'i')
    IP3_Calc(0, row_length, 0, 0, 0, 0, 0, 'i')
    P_IM_low = np.zeros(shape=(row_length), dtype=object)
    P_IM_high = np.zeros(shape=(row_length), dtype=object)
    P_fund_low = np.zeros(shape=(row_length), dtype=object)
    P_fund_high = np.zeros(shape=(row_length), dtype=object)
#    OIP3_Low = np.zeros(shape=(row_length), dtype=object)
#    OIP3_High = np.zeros(shape=(row_length), dtype=object)
#    IIP3_Low = np.zeros(shape=(row_length), dtype=object)
#    IIP3_High = np.zeros(shape=(row_length), dtype=object)
    
    print('\n\n##############    IP3    ##############\n\n')    
    count = 1
    pos = 0 # position
    for RF_freq in range(RF_freq_start, RF_freq_stop, RF_freq_step): # RF_freq is the frequencies that the measurements for IP3, IP2, etc. are wanted
        
        # Cable Loss compensation for RF cables
        Cable_Loss_LO, Pigts_loss_LO, Cable_Loss_RF1, Cable_Loss_RF2, Cable_Loss_IF2SpecAn, Pigts_loss_RF, Pigts_loss_IF, Cable_Loss_IF2PowMeter = Cable_Loss()
        Pow_RF1_comp = Pow_RF1 + Cable_Loss_RF1[pos] + Pigts_loss_RF[pos] # RF1 power compensation for various cable loss
#        print('Pow_RF1_comp', Pow_RF1_comp)
        Pow_RF2_comp = Pow_RF1 + Cable_Loss_RF2[pos] + Pigts_loss_RF[pos]
        SigGen_Setup(SigGen_RF_2, 1E6, -70) # not part of the setup but I need to set the ref clock to ext.
        SigGen_RF_2.write('OUTP OFF')
        Gain = Gain_Calc(pos, row_length, RF_freq, Pow_RF1_comp, Cable_Loss_IF2SpecAn, Pigts_loss_IF, SigGen_RF_1, SpecAn, 'G')
        print('Gain (dB): ', Gain)
        
        SpecAn_Span = config.getfloat('SpecAn', 'Span')
        SpecAn_RBW = config.getfloat('SpecAn', 'RBW')
        SpecAn_Avg = config.getfloat('SpecAn', 'Averaging')
        SpecAn_Setup(SpecAn, SpecAn_Span, SpecAn_RBW, SpecAn_Avg)
        
        SigGen_Setup(SigGen_RF_1, RF_freq, Pow_RF1_comp)
        SigGen_Setup(SigGen_RF_2, RF_freq + freq_Spacing_IP3, Pow_RF2_comp)

        print('\n\n                                                                Iteration #', count)
        print('----------------------------------------------------------------------------------------------------------------------------------------')
        count = count + 1
        IP3_f1, fund_f1, fund_f3, IP3_f3 = IP3_Products(int(RF_freq),float(freq_Spacing_IP3),float(freq_LO)) # <----- ENABLED MEASUREMENT
        
        print(f'\n\nThe current fundamental frequency (freq_RF1) is {Hz_to_MHz(RF_freq):0.0f} MHz.')
        print('-----------------------------------------------------------')
#        data_to_list_IP3(pos, RF_freq, freq_LO, 'freq') 
        data_to_list(pos, RF_freq, freq_LO, 'freq')
        for center_freq in [IP3_f1, fund_f1, fund_f3, IP3_f3]:
#            print(f'The center freq in {center_freq}')
            if center_freq == IP3_f1:
                marker_title = 'IP3 component'
                SpecAn.write('DISP:TRAC:Y:RLEV -30') # set the ref. level for lower signal level
                SpecAn.write('FREQ:CENT '+ str(center_freq) +'MHz') # sets center freq  
                SpecAn.write('CALC1:MARK1:X '+ str(center_freq) +'MHz') # Positions marker to frequency
                SpecAn.write('INIT') # Starts the measurement and waits for the end of the sweeps
                process_query(SpecAn) # checks when the sweep finishes
                Marker_Freq = Hz_to_MHz(float(SpecAn.query('CALC:MARK1:X?'))) # Outputs the measured value of the marker in screen A
                Marker_Pow = float(SpecAn.query('CALC:MARK1:Y?')) # Outputs the measured value of marker 2 in screen A
                Power_IP3_f1 = Marker_Pow + Cable_Loss_IF2SpecAn[pos] + Pigts_loss_IF[pos] # IF power compensation for various cable loss
                Power_IP3_f1 = float("{:.2f}".format(Power_IP3_f1))
#                fund_freq, IF_freq, P_IM_low, P_fund_low, P_fund_high, P_IM_high, Compr_1dB = data_to_list_IP3(pos, Power_IP3_f1, freq_LO, 'IM_low')
                fund_freq, IF_freq, P_IM_low[pos] = data_to_list(pos, Power_IP3_f1, freq_LO, 'IP')
#                print('P_IM_low main = ',P_IM_low)
            elif center_freq == IP3_f3:
                marker_title = 'IP3 component'
                SpecAn.write('DISP:TRAC:Y:RLEV -30') # set the ref. level for lower signal level
                SpecAn.write('FREQ:CENT '+ str(center_freq) +'MHz') # sets center freq  
                SpecAn.write('CALC1:MARK1:X '+ str(center_freq) +'MHz') # Positions marker to frequency
                SpecAn.write('INIT') # Starts the measurement and waits for the end of the sweeps
                process_query(SpecAn) # checks when the sweep finishes
                Marker_Freq = Hz_to_MHz(float(SpecAn.query('CALC:MARK1:X?'))) # Outputs the measured value of the marker in screen A
                Marker_Pow = float(SpecAn.query('CALC:MARK1:Y?')) # Outputs the measured value of marker 2 in screen A
                Power_IP3_f3 = Marker_Pow + Cable_Loss_IF2SpecAn[pos] + Pigts_loss_IF[pos] # IF power compensation for various cable loss
                Power_IP3_f3 = float("{:.2f}".format(Power_IP3_f3))
#                fund_freq, IF_freq, P_IM_low, P_fund_low, P_fund_high, P_IM_high, Compr_1dB = data_to_list_IP3(pos, Power_IP3_f3, freq_LO, 'IM_high')
                fund_freq, IF_freq, P_IM_high[pos] = data_to_list(pos, Power_IP3_f3, freq_LO, 'IP')
            elif center_freq == fund_f1:
                marker_title = 'fundamental component'
                SpecAn.write('DISP:TRAC:Y:RLEV -10') # set the ref. level for lower signal level
                SpecAn.write('FREQ:CENT '+ str(center_freq) +'MHz') # sets center freq  
                SpecAn.write('CALC1:MARK1:X '+ str(center_freq) +'MHz') # Positions marker to frequency
                SpecAn.write('INIT') # Starts the measurement and waits for the end of the sweeps
                process_query(SpecAn) # checks when the sweep finishes
                Marker_Freq = Hz_to_MHz(float(SpecAn.query('CALC:MARK1:X?'))) # Outputs the measured value of the marker in screen A
                Marker_Pow = float(SpecAn.query('CALC:MARK1:Y?')) # Outputs the measured value of marker 2 in screen A
                Power_fund_f1 = Marker_Pow + Cable_Loss_IF2SpecAn[pos] + Pigts_loss_IF[pos] # IF power compensation for various cable loss
                Power_fund_f1 = float("{:.2f}".format(Power_fund_f1))
#                fund_freq, IF_freq, P_IM_low, P_fund_low, P_fund_high, P_IM_high, Compr_1dB = data_to_list_IP3(pos, Power_fund_f1, freq_LO, 'fund_low')
                fund_freq, IF_freq, P_fund_low[pos] = data_to_list(pos, Power_fund_f1, freq_LO, 'IP')
            elif center_freq == fund_f3:
                marker_title = 'fundamental component'
                SpecAn.write('DISP:TRAC:Y:RLEV -10') # set the ref. level for lower signal level
                SpecAn.write('FREQ:CENT '+ str(center_freq) +'MHz') # sets center freq  
                SpecAn.write('CALC1:MARK1:X '+ str(center_freq) +'MHz') # Positions marker to frequency
                SpecAn.write('INIT') # Starts the measurement and waits for the end of the sweeps
                process_query(SpecAn) # checks when the sweep finishes
                Marker_Freq = Hz_to_MHz(float(SpecAn.query('CALC:MARK1:X?'))) # Outputs the measured value of the marker in screen A
                Marker_Pow = float(SpecAn.query('CALC:MARK1:Y?')) # Outputs the measured value of marker 2 in screen A
                Power_fund_f3 = Marker_Pow + Cable_Loss_IF2SpecAn[pos] + Pigts_loss_IF[pos] # IF power compensation for various cable loss
                Power_fund_f3 = float("{:.2f}".format(Power_fund_f3))
#                fund_freq, IF_freq, P_IM_low, P_fund_low, P_fund_high, P_IM_high, Compr_1dB = data_to_list_IP3(pos, Power_fund_f3, freq_LO, 'fund_high')
                fund_freq, IF_freq, P_fund_high[pos] = data_to_list(pos, Power_fund_f3, freq_LO, 'IP')
            else:
                sys.exit()
#            print(f'\nThe frequency and power for the {marker_title} is: {Marker_Freq:0.0f} MHz & {Marker_Pow:0.2f} dBm.')
        print('P_IM_low:     ', P_IM_low)
        print('P_fund_low:   ', P_fund_low)
        print('P_fund_high:  ', P_fund_high)
        print('P_IM_high:    ', P_IM_high)       
        
        Delta_P_Low, Delta_P_High, OIP3_Low, OIP3_High, IIP3_Low, IIP3_High = IP3_Calc(pos, row_length, Gain, P_IM_low, P_fund_low, P_fund_high, P_IM_high, 'calc')

#        excel_plot(fund_freq, IF_freq, Gain, P_IM_low, P_fund_low, P_fund_high, P_IM_high, Delta_P_Low, Delta_P_High, OIP3_Low, OIP3_High, IIP3_Low, IIP3_High, Cable_Loss_RF1, Cable_Loss_RF2, Cable_Loss_IF2SpecAn, Cable_Loss_IF2PowMeter, Pigts_loss_RF, Pigts_loss_IF)
    
        print('\n\n= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = ')
        pos = pos + 1
        
    ##### Pass data to Excel #####
    Temperature = str('22°C')
    Timestamp = datetime.now().strftime('%d-%m-%Y--%H-%M-%S') #write date and time of test to folder name
    workbook = xlsxwriter.Workbook('IP3_Products_' + Temperature + ' - ' + Timestamp + '.xlsx')
    IP3_tab = workbook.add_worksheet('IP3')
    IP3_tab.set_tab_color('green')
    IP3_Details_tab = workbook.add_worksheet('IP3_Details')
    IP3_Details_tab.set_tab_color('red')        
    bold = workbook.add_format({'bold': 1})
    
    IP3_heads_tab1 = ['RF Freq (MHz)', 'IF Freq (MHz)', 
                      'IIP3_Low (dBm)', 'IIP3_High (dBm)', 'OIP3_Low (dBm)', 'OIP3_High (dBm)',
                      'P_IM_low (dBm)', 'P_fund_low (dBm)', 'P_fund_high (dBm)', 'P_IM_high (dBm)'] 
    IP3_heads_tab2 = ['RF Freq (MHz)', 'IF Freq (MHz)', 'Gain (dB)',
                      'IM_Low-both (dBc)', 'IM_High-both (dBc)',
                      'Cable_Loss_RF1 (dB)', 'Cable_Loss_RF2 (dB)', 'Cable_Loss_SpecAn (dB)', 'Cable_Loss_PowMeter (dB)', 
                      'Pigtail_Atten_RF (dB)', 'Pigtail_Atten_IF (dB)']           

    IIP3_Low = IIP3_Low[row_length-1]
    IIP3_High = IIP3_High[row_length-1]
    OIP3_Low = OIP3_Low[row_length-1]
    OIP3_High = OIP3_High[row_length-1]
    Delta_P_Low = Delta_P_Low[row_length-1]
    Delta_P_High = Delta_P_High[row_length-1]
    
    IP3_tab.write_row('A1', IP3_heads_tab1, bold) # the first cell where the headings will be placed
    IP3_tab.write_column('A2', fund_freq) # the cell where the data will be placed & the data column
    IP3_tab.write_column('B2', IF_freq)
    IP3_tab.write_column('C2', IIP3_Low)
    IP3_tab.write_column('D2', IIP3_High)
    IP3_tab.write_column('E2', OIP3_Low)
    IP3_tab.write_column('F2', OIP3_High)
    IP3_tab.write_column('G2', P_IM_low)
    IP3_tab.write_column('H2', P_fund_low)
    IP3_tab.write_column('I2', P_fund_high)
    IP3_tab.write_column('J2', P_IM_high)
    
    IP3_Details_tab.write_row('A1', IP3_heads_tab2, bold) # the first cell where the headings will be placed
    IP3_Details_tab.write_column('A2', fund_freq) # the cell where the data will be placed & the data column
    IP3_Details_tab.write_column('B2', IF_freq)
    IP3_Details_tab.write_column('C2', Gain)
    IP3_Details_tab.write_column('D2', Delta_P_Low)
    IP3_Details_tab.write_column('E2', Delta_P_High)
    IP3_Details_tab.write_column('F2', Cable_Loss_RF1)
    IP3_Details_tab.write_column('G2', Cable_Loss_RF2)
    IP3_Details_tab.write_column('H2', Cable_Loss_IF2SpecAn)
    IP3_Details_tab.write_column('I2', Cable_Loss_IF2PowMeter)
    IP3_Details_tab.write_column('J2', Pigts_loss_RF)
    IP3_Details_tab.write_column('K2', Pigts_loss_IF)
    
    #######################################################################
    #
    # Create a scatter chart sub-type with straight lines and no markers.
    #
    chart1 = workbook.add_chart({'type': 'scatter',
                                 'subtype': 'straight'})
    
    # Configure the first series.
    chart1.add_series({
        'name':       'Input_IP3 Low',
        'categories': 'IP3!$A$2:$A$14',
        'values':     'IP3!$C$2:$C$15',
    })
    
    # Configure second series.
    chart1.add_series({
        'name':       'Input_IP3 High',
        'categories': 'IP3!$A$2:$A$15',
        'values':     'IP3!$D$2:$D$15',
    })
    
    # Add a chart title and some axis labels.
    chart1.set_title ({'name': 'Input IP3 (22°C)'})
    chart1.set_x_axis({'name': 'Frequency (MHz)'})
    chart1.set_y_axis({'name': 'IIP3 (dBm)'})
    
    # Set an Excel chart style.
    chart1.set_style(13)
    
    # Insert the chart into the worksheet (with an offset).
    IP3_tab.insert_chart('A16', chart1, {'x_offset': 0, 'y_offset': 0})
    
    #### 2nd ####
    chart2 = workbook.add_chart({'type': 'scatter',
                                 'subtype': 'straight'})
    
    # Configure the first series.
    chart2.add_series({
        'name':       'Output_IP3 Low',
        'categories': 'IP3!$B$2:$B$15',
        'values':     'IP3!$E$2:$E$15',
    })
    
    # Configure the scond series.
    chart2.add_series({
        'name':       'Output_IP3 High',
        'categories': 'IP3!$B$2:$B$15',
        'values':     'IP3!$F$2:$F$15',
    })
    
    # Add a chart title and some axis labels.
    chart2.set_title ({'name': 'Output IP3 (22°C)'})
    chart2.set_x_axis({'name': 'Frequency (MHz)'})
    chart2.set_y_axis({'name': 'OIP3 (dBm)'})
    
    # Set an Excel chart style.
    chart2.set_style(13)
    
    # Insert the chart into the worksheet (with an offset).
    IP3_tab.insert_chart('I16', chart2, {'x_offset': 0, 'y_offset': 0})
   
    #######################################################################
    #
    # Create a scatter chart sub-type with straight lines and no markers.
    #
    chart3 = workbook.add_chart({'type': 'scatter',
                                 'subtype': 'straight'})
    
    # Configure the first series.
    chart3.add_series({
        'name':       'DUT Gain',
        'categories': 'IP3_Details!$B$2:$B$14',
        'values':     'IP3_Details!$C$2:$C$15',
    })
    
    # Add a chart title and some axis labels.
    chart3.set_title ({'name': 'DUT Gain'})
    chart3.set_x_axis({'name': 'Frequency (MHz)'})
    chart3.set_y_axis({'name': 'Gain (dB)'})
    
    # Set an Excel chart style.
    chart3.set_style(13)
    
    # Insert the chart into the worksheet (with an offset).
    IP3_Details_tab.insert_chart('A16', chart3, {'x_offset': 0, 'y_offset': 0})
    
    workbook.close()
        
    return

def IP3_meas_TOI_function():

    # Option 3 - TOI marker search
#    SpecAn.write('BAND 10kHz') # Sets the resolution bandwidth to X MHz
#    SpecAn.write('FREQ:SPAN 8MHz') # set freq span
#    SpecAn.write('FREQ:CENT 899MHz') # sets center freq  <-------- THIS WILL BE VARIABLE
#    SpecAn.write('CALC:MARK:FUNC:TOI ON') # Switches on the measurement of the third-order intercept
#    #    A two-tone signal with equal carrier levels is expected at the RF input of the instrument. Marker 1 and
#    #    marker 2 (both normal markers) are set to the maximum of the two signals. Delta marker 3 and delta
#    #    marker 4 are positioned to the intermodulation products. The delta markers can be modified separately
#    #    afterwards with the commands CALCulate:DELTamarker3:X and CALCulate:DELTamarker4:
#    #    X. The third-order intercept is calculated from the level spacing between the normal markers and the delta
#    #    markers.
#    SpecAn.write('CALC:MARK:FUNC:TOI:MARK SEAR') # selects TOI marker search mode
#
#    SpecAn.write('INIT') # Starts the measurement and waits for the end of the sweeps
#    process_query(SpecAn) # checks when the sweep finishes
#    
#    Var = SpecAn.write('CALC:MARK:FUNC:TOI:RES?') # Outputs the 3rd order intercept point - requires long time due to wide freq range and low RBW
#    print(f'\nThe markers from TOI is: {Var}')
    
    return









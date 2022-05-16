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


def excel_plot(fund_freq, IF_freq, Gain, P_IM_low, P_fund_low, P_fund_high, P_IM_high, Delta_P_Low, Delta_P_High, OIP3_Low, OIP3_High, IIP3_Low, IIP3_High, Cable_Loss_RF1, Cable_Loss_RF2, Cable_Loss_IF2SpecAn, Cable_Loss_IF2PowMeter, Pigts_loss_RF, Pigts_loss_IF):
        
    workbook = xlsxwriter.Workbook('NonLinear_data.xlsx')
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': 1})
    
    headings = ['RF Freq (MHz)', 'IF Freq (MHz)', 'P_IM_low (dBm)', 'P_fund_low (dBm)', 'P_fund_high (dBm)', 'P_IM_high (dBm)', 'Gain (dB)', 'IM_Low-both (dBc)', 'IM_High-both (dBc)', 'OIP3_Low (dBm)', 'OIP3_High (dBm)', 'IIP3_Low (dBm)', 'IIP3_High (dBm)', 'Cable_Loss_RF1 (dB)', 'Cable_Loss_RF2 (dB)', 'Cable_Loss_SpecAn (dB)', 'Cable_Loss_PowMeter (dB)', 'Pigtail_Atten_RF (dB)', 'Pigtail_Atten_IF (dB)', '1dB Point'] # move it to IP3_meas                
    print('P_IM_high')
    print(type(P_IM_high))
    print('Delta_P_Low')
    print(type(Delta_P_Low))
    worksheet.write_row('A1', headings, bold) # the first cell where the headings will be placed
    worksheet.write_column('A2', fund_freq) # the cell where the data will be placed & the data column
    worksheet.write_column('B2', IF_freq)
    worksheet.write_column('C2', P_IM_low)
    worksheet.write_column('D2', P_fund_low)
    worksheet.write_column('E2', P_fund_high)
    worksheet.write_column('F2', P_IM_high)
    worksheet.write_column('G2', Gain)
    worksheet.write_column('H2', Delta_P_Low)
    worksheet.write_column('I2', Delta_P_High)
    worksheet.write_column('J2', OIP3_Low)
    worksheet.write_column('K2', OIP3_High)
    worksheet.write_column('L2', IIP3_Low)
    worksheet.write_column('M2', IIP3_High)   
    worksheet.write_column('N2', Cable_Loss_RF1)
    worksheet.write_column('O2', Cable_Loss_RF2)
    worksheet.write_column('P2', Cable_Loss_IF2SpecAn) 
    worksheet.write_column('Q2', Cable_Loss_IF2PowMeter)
    worksheet.write_column('R2', Pigts_loss_RF)
    worksheet.write_column('S2', Pigts_loss_IF)
#    worksheet.write_column('R2', Compres_1dB)
#    worksheet.write_column('S2', Cable_Loss_IF)
#    
    #######################################################################
    #
    # Create a scatter chart sub-type with straight lines and no markers.
    #
    chart1 = workbook.add_chart({'type': 'scatter',
                                 'subtype': 'straight'})
    
    # Configure the first series.
    chart1.add_series({
        'name':       '=Sheet1!$K$1',
        'categories': '=Sheet1!$B$2:$B$15',
        'values':     '=Sheet1!$K$2:$K$15',
    })
    
    # Configure second series.
    chart1.add_series({
        'name':       '=Sheet1!$J$1',
        'categories': '=Sheet1!$B$2:$B$15',
        'values':     '=Sheet1!$J$2:$J$15',
    })
    
    # Add a chart title and some axis labels.
    chart1.set_title ({'name': 'OIP3'})
    chart1.set_x_axis({'name': 'Frequency (MHz)'})
    chart1.set_y_axis({'name': 'Power (dBm)'})
    
    # Set an Excel chart style.
    chart1.set_style(13)
    
    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('A16', chart1, {'x_offset': 0, 'y_offset': 0})
    
    chart2 = workbook.add_chart({'type': 'scatter',
                                 'subtype': 'straight'
    })
    
    # Configure second series.
    chart2.add_series({
        'name':       '=Sheet1!$M$1',
        'categories': '=Sheet1!$B$2:$B$15',
        'values':     '=Sheet1!$M$2:$M$15',
    })
    
    # Configure second series.
    chart2.add_series({
        'name':       '=Sheet1!$L$1',
        'categories': '=Sheet1!$B$2:$B$15',
        'values':     '=Sheet1!$L$2:$L$15',
    })
    
    # Add a chart title and some axis labels.
    chart2.set_title ({'name': 'IIP3'})
    chart2.set_x_axis({'name': 'Frequency (MHz)'})
    chart2.set_y_axis({'name': 'Power (dBm)'})
    
    # Set an Excel chart style.
    chart2.set_style(13)
    
    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('N16', chart2, {'x_offset': 0, 'y_offset': 0})
    
    workbook.close()
    
    return

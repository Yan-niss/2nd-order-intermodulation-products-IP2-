import pyvisa as visa
import time
import os
import sys
import errno
import numpy as np
import xlsxwriter

from Valid_Input import process_query
from User_Inputs import User_Inputs
from Freq_Conv import Hz_to_MHz
from Freq_Conv import MHz_to_Hz
from Instr_Setup import SpecAn_Setup
from Instr_Setup import SigGen_Setup


def Gain_Calc(pos, row_length, RF_freq, Pout_SigGen1, Cable_Loss_IF, Pigts_loss_IF, SigGen_RF_1, SpecAn, race):
    
    global Gain
        
    freq_RF1, freq_RF2, Pin_DUT_1, Pin_DUT_2, freq_LO, Pow_LO, RF_freq_start, RF_freq_stop, RF_freq_step, freq_Spacing_IP3, freq_Spacing_IP2 = User_Inputs()
    SigGen_Setup(SigGen_RF_1, RF_freq, Pout_SigGen1)    

    SpecAn.write('CALC:MARK1:MAX')
    
    if race == 'i':
        Gain = np.zeros(shape=(row_length), dtype=object)
    elif race == 'G':      
        SpecAn_Setup(SpecAn, 3000, 100, 3) # SPAN, RBW, AVG
        print('\n\n##############    Gain Calculation    ##############')
        freq_LO = float(Hz_to_MHz(freq_LO))
        RF_freq = float(Hz_to_MHz(RF_freq))
#        SpecAn.write('INP:ATT 30dB') # Sets the attenuation on the attenuator to **dB and switches off the coupling to the reference level
        SpecAn.write('DISP:TRAC:Y:RLEV 0') # set the ref. level for lower signal level
        SpecAn.write('FREQ:CENT '+ str(freq_LO - RF_freq) +'MHz') # sets center freq  
        SpecAn.write('CALC1:MARK1:X '+ str(freq_LO - RF_freq) +'MHz') # Positions marker to frequency
        SpecAn.write('INIT') # Starts the measurement and waits for the end of the sweeps
        process_query(SpecAn) # checks when the sweep finishes
        SpecAn.write('CALC:MARK1:MAX') # Switches the search limit function on for screen A
#        Marker_Freq = Hz_to_MHz(float(SpecAn.query('CALC:MARK1:X?'))) # Outputs the measured value of the marker in screen A
#        print(f'The measured IF frequency in the SpecAn is (Marker_Freq): {Marker_Freq:0.6f}MHz')
        Marker_Pow = float(SpecAn.query('CALC:MARK1:Y?')) # Outputs the measured value from SpecAn
        Marker_fund_IF = float("{:.2f}".format(Marker_Pow))
        Pout_DUT = Marker_fund_IF + Cable_Loss_IF[pos] + Pigts_loss_IF[pos] # power at DUT's IF output
#        print('Gain Pout_DUT = ', Pout_DUT)
#        print('Gain Pin_DUT_1 = ', Pin_DUT_1)
        Gain[pos] = Pout_DUT -  float(Pin_DUT_1) # Conversion Gain calculation
        Gain[pos] = float("{:.2f}".format(Gain[pos]))
    else:
        sys.exit()
    
    return (Gain)


def IP3_Calc(pos, row_length, Gain, P_IM_low, P_fund_low, P_fund_high, P_IM_high, race):
    
    global IM_low, IM_high, OIP3_Low, OIP3_High, IIP3_Low, IIP3_High, Delta_P_Low, Delta_P_High
    
    if race == 'i':
        IM_low = np.zeros(shape=(row_length), dtype=object)
        IM_high = np.zeros(shape=(row_length), dtype=object)
        OIP3_Low = np.zeros(shape=(row_length), dtype=object)
        OIP3_High = np.zeros(shape=(row_length), dtype=object)
        IIP3_Low = np.zeros(shape=(row_length), dtype=object)
        IIP3_High = np.zeros(shape=(row_length), dtype=object)
        Delta_P_Low = np.zeros(shape=(row_length), dtype=object)
        Delta_P_High = np.zeros(shape=(row_length), dtype=object)
    elif race == 'calc':
    #    IM_low[pos] =  P_IM_low[pos] - P_fund_low[pos] # dBc
    #    IM_high[pos] =  P_IM_high[pos] - P_fund_high[pos] # dBc
        
        # 3rd Order Intermodulation Point - TOI
        
        # OPTION 1 (R&S)
        # Based on "Intermodulation Distortion Measurements on Modern Spectrum Analyzers - Application Note - R&S"
    #    # This method assumes that the two intermodulation products (Hi & Lo) have the same power
    #    P_fund_Out = *TBD* # measured power of the fundamantal tone (dBm) - (Mentioned as P_Tone in the AN)
    #    IP3 = *TBD* # measured absolute power of intermodulation products (dBm)
    #    P_Delta = P_fund_Out - P_IM3 # relative power of the intermodulation products referenced to P_fund_Out (dBc)
    #    OIP3 = P_fund_Out + P_Delta/2 # IP3 referenced to the output of the DUT
        
#        print('P_fund_high', P_fund_high[pos])
#        print('P_fund_low', P_fund_low[pos])
#        print('P_fund_high', P_IM_low[pos])

        # OPTION 2 (Keysight)
        # Based on "Keysight X-Series Signal Analyser N9060-90041 manual" - p. 1456
        # This method assumes that the two intermodulation products (Hi & Lo) have different power with each other and takes both into consideration
#        print('P_fund_high[pos]', P_fund_high)
#        print('P_fund_low[pos]', P_fund_low)
#        print('P_IM_low[pos]', P_IM_low)
        OIP3_Low[pos] = P_fund_high/2 + P_fund_low - P_IM_low/2 # TOI calculation of the low interm. product refered at the output of the DUT
        OIP3_High[pos] = P_fund_low/2 + P_fund_high - P_IM_high/2 # TOI calculation of the high interm. product refered at the output of the DUT
        IIP3_Low[pos] = OIP3_Low[pos]-Gain[pos] # TOI calculation of the low interm. product refered at the output of the DUT
        IIP3_High[pos] = OIP3_High[pos]-Gain[pos] # TOI calculation of the high interm. product refered at the output of the DUT        
#        print('OIP3_Low')
#        print(type(OIP3_Low))
#        print('IIP3_Low')
#        print(type(IIP3_Low))
        Delta_P_Low[pos] = P_IM_low - (2*P_fund_low + P_fund_high)/3 # relative power of the low side intermodulation products referenced to OIP3_Low (dBc)
        Delta_P_High[pos] = P_IM_high - (2*P_fund_high + P_fund_low)/3 # relative power of the high side intermodulation products referenced to OIP3_High (dBc)
#        print('Delta_P_Low')
#        print(type(Delta_P_Low))
    else:
        sys.exit()
        
#    return (IM_low, IM_high, OIP3_Low, OIP3_High, Delta_P_Low, Delta_P_High)
    return (Delta_P_Low, Delta_P_High, OIP3_Low, OIP3_High, IIP3_Low, IIP3_High)


def IP2_Calc(pos, row_length, P_IM_low, IF_fund_low, race):
    
    global IM_low, IM_high, IIP2
    
    if race == 'i':
        IM_low = np.zeros(shape=(row_length), dtype=object)
#        IM_high = np.zeros(shape=(row_length), dtype=object)
        IIP2 = np.zeros(shape=(row_length), dtype=object)
#        OIP3_High = np.zeros(shape=(row_length), dtype=object)
#        Delta_P_Low = np.zeros(shape=(row_length), dtype=object)
#        Delta_P_High = np.zeros(shape=(row_length), dtype=object)
    elif race == 'calc_IP2':
        IM_low[pos] =  P_IM_low[pos] - IF_fund_low[pos] # dBc
#        IM_high[pos] =  P_IM_high[pos] - P_fund_high[pos] # dBc
        
        # 3rd Order Intermodulation Point - TOI
        
        # OPTION 1 (R&S)
        # Based on "Intermodulation Distortion Measurements on Modern Spectrum Analyzers - Application Note - R&S"
    #    # This method assumes that the two intermodulation products (Hi & Lo) have the same power
    #    P_fund_Out = *TBD* # measured power of the fundamantal tone (dBm) - (Mentioned as P_Tone in the AN)
    #    IP3 = *TBD* # measured absolute power of intermodulation products (dBm)
    #    P_Delta = P_fund_Out - P_IM3 # relative power of the intermodulation products referenced to P_fund_Out (dBc)
    #    OIP3 = P_fund_Out + P_Delta/2 # IP3 referenced to the output of the DUT

        
        # OPTION 2 (Keysight)
        # Based on "Keysight X-Series Signal Analyser N9060-90041 manual" - p. 1456
        # This method assumes that the two intermodulation products (Hi & Lo) have different power with each other and takes both into consideration
        IIP2[pos] = float("{:.2f}".format(2*IF_fund_low[pos] - IM_low[pos])) # https://www.microwavejournal.com/articles/print/3411-two-tone-vs-single-tone-measurement-of-second-order-nonlinearity
#        OIP3_High[pos] = float("{:.2f}".format(P_fund_low[pos]/2 + P_fund_high[pos] - P_IM_high[pos]/2)) # TOI calculation of the high interm. product refered at the output of the DUT
#        Delta_P_Low[pos] = float("{:.2f}".format(P_IM_low[pos] - (2*P_fund_low[pos] + P_fund_high[pos])/3)) # relative power of the low side intermodulation products referenced to OIP3_Low (dBc)
#        Delta_P_High[pos] = float("{:.2f}".format(P_IM_high[pos] - (2*P_fund_high[pos] + P_fund_low[pos])/3)) # relative power of the high side intermodulation products referenced to OIP3_High (dBc)
    else:
        sys.exit()
        
#    return (IM_low, IM_high, OIP3_Low, OIP3_High, Delta_P_Low, Delta_P_High)
    return (IM_low, IIP2)




def IP3_Products(freq_RF1, freq_spacing, freq_LO):
#    print(f'The freq_RF2 is {Hz_to_MHz(freq_RF1):0.0f} MHz.')
#    print(f'The freq_RF2 is {Hz_to_MHz(freq_RF2):0.0f} MHz.')
#    print('The freq_LO is ',freq_LO,' MHz.')

    freq_RF2 = freq_RF1 + freq_spacing
    
    # Fundamental freqs
    fund_f1 = Hz_to_MHz(freq_LO-freq_RF2) # in MHz
    fund_f2 = Hz_to_MHz(freq_LO + freq_RF2)
    fund_f3 = Hz_to_MHz(freq_LO-freq_RF1)
    fund_f4 = Hz_to_MHz(freq_LO + freq_RF1)
    
    # IP3 Products
    IP3_f1 = Hz_to_MHz(freq_LO-(2*freq_RF2-freq_RF1))
    IP3_f2 = Hz_to_MHz(freq_LO + (2*freq_RF2-freq_RF1))
    IP3_f3 = Hz_to_MHz(freq_LO-(2*freq_RF1-freq_RF2))
    IP3_f4 = Hz_to_MHz(freq_LO + (2*freq_RF1-freq_RF2))
    
#    print('\nThe fundamental components are:')
#    print(f'Fund_f1 = {fund_f1:0.0f} MHz')
#    print(f'Fund_f2 = {fund_f2:0.0f} MHz')
#    print(f'Fund_f3 = {fund_f3:0.0f} MHz')
#    print(f'Fund_f4 = {fund_f4:0.0f} MHz')
    
#    print('\nThe 3rd order products are:')
#    print(f'IP3_f1 = {IP3_f1:0.0f} MHz')
#    print(f'IP3_f2 = {IP3_f2:0.0f} MHz')
#    print(f'IP3_f3 = {IP3_f3:0.0f} MHz')
#    print(f'IP3_f4 = {IP3_f4:0.0f} MHz')   

    print(' - - - - - - - IP3_Low - - - - - - - - - - - - - - Fund_Low - - - - - - - - Fund_High - - - - - - - - - - - - - - IP3_High - - - - - - - ')
    print(' - freq_LO - (2*freq_RF2 - freq_RF1) - - - - freq_LO - freq_RF2 - - - - freq_LO - freq_RF1 - - - - - freq_LO - (2*freq_RF1 - freq_RF2) - ')
    print(f' - - - - - - - {IP3_f1:0.0f}MHz - - - - - - - - - - - - - - {fund_f1:0.0f}MHz - - - - - - - - - - - {fund_f3:0.0f}MHz - - - - - - - - - - - - - - {IP3_f3:0.0f}MHz - - - - - - - - ')
    print('----------------------------------------------------------------------------------------------------------------------------------------')
    return(IP3_f1, fund_f1, fund_f3, IP3_f3)
    
    
def Mixer_IP2_Products(freq_RF1, freq_spacing, freq_LO):

    freq_RF2 = freq_RF1 + freq_spacing
    
    # Fundamental freqs
    fund_f1 = Hz_to_MHz(freq_LO-freq_RF2) # in MHz
    fund_f3 = Hz_to_MHz(freq_LO-freq_RF1)
    
    # IP2 Products
#    IP2_RF_high = Hz_to_MHz((freq_RF1 + freq_RF2) - freq_LO) # f1 - f2
#    IP2_RF_low = Hz_to_MHz((freq_RF1 + freq_RF2) - freq_LO) # f1 + f2
    IP2_IF_1 = Hz_to_MHz((freq_RF1 + freq_RF2) - freq_LO) # (f1+f2)-f_LO 
    IP2_IF_2 = Hz_to_MHz(freq_LO - (freq_RF1 - freq_RF2)) # f_LO-(f1-f2)
    IP2_IF_3 = Hz_to_MHz((freq_RF1 - freq_RF2) + freq_LO) # (f1-f2)+f_LO
    IP2_IF_4 = Hz_to_MHz((freq_RF1 + freq_RF2) + freq_LO) # (f1+f2)+f_LO
    
    print(' - - - (f1+f2)-f_LO - - - ') #- - f_LO-(f1-f2) - - - (f1-f2)+f_LO  - - -  (f1+f2)+f_LO - ')
    print(f' - - - - - {IP2_IF_1:0.0f}MHz - - - - - ') #- - - {IP2_IF_2:0.0f}MHz - - - - {IP2_IF_3:0.0f}MHz - - - - - {IP2_IF_4:0.0f}MHz - ')
    print('------------------------------------')
    return(fund_f1, fund_f3, IP2_IF_1, IP2_IF_2, IP2_IF_3, IP2_IF_4)
    
    
    
    
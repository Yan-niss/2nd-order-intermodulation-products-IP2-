import visa
import os
import sys
import math
import errno
import numpy
import time
import configparser

def Initialization():

    os.chdir(os.path.dirname(os.path.abspath(__file__)))
        
    config = configparser.ConfigParser(allow_no_value=True)
    config.read('IP3_config.ini')
    config.set('SigGen', '; This is a comment.', None)
    
    # Listing connected resources
    rm = visa.ResourceManager()
    device_list = rm.list_resources()



    # GPIB addresses
    addr_RF_1 = config.get('GPIB_addr', 'addr_RF_1') # R&S SMU200A
    addr_RF_2 = config.get('GPIB_addr', 'addr_RF_2') # R&S SMIQ 03b
    addr_LO = config.get('GPIB_addr', 'addr_LO') # R&S SMU200A
    addr_SpecAn = config.get('GPIB_addr', 'addr_SpecAn') # R&S FSP-3(0)
    addr_PM = config.get('GPIB_addr', 'addr_PM') # Keysight E4418B

    try:
        SigGen_RF_1 = rm.open_resource(addr_RF_1) # RF Signal Generator 1
    except Exception:
        print('Unable to connect to the Signal Generator (RF #1) at '+ str(addr_RF_1) + '. Aborting script.')
        sys.exit()

    try:
        SigGen_RF_2 = rm.open_resource(addr_RF_2) # RF Signal Generator 2
    except Exception:
        print('Unable to connect to Signal Generator (RF #2) at '+ str(addr_RF_2) + '. Aborting script.')
        sys.exit()

#    try:
#        SigGen_LO = rm.open_resource(addr_LO) # LO Signal Generator - optional for mixers
#    except Exception:
#        print('Unable to connect to Signal Generator (LO) at '+ str(addr_LO) + '. Aborting script.')
#        sys.exit()
	
    try:
        SpecAn = rm.open_resource(addr_SpecAn) # Spectrum Analyser
    except Exception:
        print('Unable to connect to Spectrum Analyser at '+ str(addr_SpecAn) + '. Aborting script.')
        sys.exit()

    try:
        PowMeter = rm.open_resource(addr_PM) # Power Meter for calibration
    except Exception:
        print('Unable to connect to Power Meter at '+ str(addr_PM) + '. Aborting script.')
        sys.exit()
			
    time.sleep(0.1)
#    print('\nThe connected devices are:\n')

	# Query - Identify instruments
#    print(SigGen_RF_1.query('*IDN?'))
#    print(SigGen_RF_2.query('*IDN?'))
#    print(SigGen_LO.query('*IDN?'))
#    print(SpecAn.query('*IDN?'))
#    print(PowMeter.query('*IDN?'))

	# Save IDN to variable for later use in Excel
    SigGen_RF_1_IDN= SigGen_RF_1.query('*IDN?')
    SigGen_RF_2_IDN= SigGen_RF_2.query('*IDN?')
#    SigGen_LO_IDN= SigGen_LO.query('*IDN?')
    SpecAn_IDN= SpecAn.query('*IDN?')
    PowMeter_IDN= PowMeter.query('*IDN?')

	# Resetting devices to default state
    SigGen_RF_1.write('*RST')
    SigGen_RF_2.write('*RST')
#    SigGen_LO.write('*RST')
    SpecAn.write('*RST')
    PowMeter.write('*RST')

	# Clears registers
    SigGen_RF_1.write('*CLS')
    SigGen_RF_2.write('*CLS')
#    SigGen_LO.write('*CLS')
    SpecAn.write('*CLS')
    PowMeter.write('*CLS')

#    print('\nThe initialization of the instruments has been done!\n')


#    return (SigGen_RF_1, SigGen_RF_2, SigGen_LO, SpecAn, PowMeter)
    return (SigGen_RF_1, SigGen_RF_2, SpecAn, PowMeter)






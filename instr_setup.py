import pyvisa as visa
import time
import os
import sys
#import math
import errno
import numpy as np
import xlsxwriter

from Valid_Input import process_query
#from Initialization import Initialization
#from Valid_Input import process_query
#from Freq_Conv import Hz_to_MHz
#from Freq_Conv import MHz_to_Hz
#from Calculations import IP3_Products
#from Calculations import Mixer_IP2_Products
#from Initialization import Initialization


# R&S Signal Generator setup
def SigGen_Setup(SigGen, freq, power):
    
    SigGen.write('ROSC:SOUR EXT') # Enables external 10MHz reference
    SigGen.write('ROSC:EXT 10 MHz')
    SigGen.write('MOD:STAT OFF') # Disables all modulation
    SigGen.write('FREQ '+ str(freq)) # set frequency in Hz
    SigGen.write('FREQ:MODE CW') # SigGen sends a CW
    SigGen.write('FREQ:OFFS 0kHz') # freq offset = 0
    SigGen.write('FREQ:MULT 1') # multiplication factor = 1
    SigGen.write(':POW '+ str(power)) # sets the power in dBm
    SigGen.write('POW:MODE CW') # SigGen sends a CW
    SigGen.write('OUTP ON') # turns on the SigGen
    return

# R&S Spectrum Analyzer setup
def SpecAn_Setup(SpecAn, span, RBW, average):
    
    span = str(span)
    RBW = str(RBW)
    average = str(average)
    
    SpecAn.write('ROSC:SOUR INT') # Enables internal 10MHz reference
    SpecAn.write('CALC:UNIT:POW DBM') # Sets the power unit for screen A to dBm
    SpecAn.write('FREQ:MODE SWE') # enables freq doamain
    SpecAn.write('INIT:CONT OFF') # Switching to single-sweep mode
    SpecAn.write('AVER:STAT ON') # Switches on the calculation of average
    SpecAn.write('BAND:VID:AUTO ON') # Video BW auto
    SpecAn.write('DET RMS') # Sets detector type to RMS
    SpecAn.write('SWE:TIME:AUTO ON') # automatic sweep time
    SpecAn.write('DISP:TRAC:Y:RLEV:OFFS 0dB') # defines the offset of the reference level
    SpecAn.write('CALC:MARK1 ON') # Switches the marker in screen A on.
    SpecAn.write('CALC:MARK:MAX:AUTO ON')
    SpecAn.write('CALC:MARK1:MAX')
    SpecAn.write('CALC:MARK1:TRAC 1') # sets marker 1 in screen A to trace 1
    SpecAn.write('CALC:MARK1:COUN ON') # Switches the frequency counter for the marker 1
    SpecAn.write('CALC:MARK1:X:SLIM ON') # Switches the search limit function on for screen A
    SpecAn.write('CALC:MARK1:X:SLIM:LEFT 300Hz') # Sets the left limit of the search range in screen A to 1kHz
    SpecAn.write('CALC:MARK1:X:SLIM:RIGH 300Hz') # Sets the right limit of the search range in screen A to 1kHz
    SpecAn.write('INP:ATT 20dB') # Sets the attenuation on the attenuator to **dB and switches off the coupling to the reference level
#    SpecAn.write('INP:ATT:AUTO ON') # automatically couples the input attenuation to the reference level
#    time.sleep(0.5)
#    print('\nATTEN = ', int(SpecAn.query('INP:ATT?')),'dB.')    
    SpecAn.write('FREQ:SPAN '+ span +'Hz') # SPAN
    SpecAn.write('BAND '+ RBW +'Hz') # RBW
    SpecAn.write('AVER:COUN '+ average) # ANERAGE
    process_query(SpecAn)

    
    return
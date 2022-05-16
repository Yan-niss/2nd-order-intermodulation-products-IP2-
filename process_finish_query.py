# -*- coding: utf-8 -*-
"""
Created on Sun Jan 10 18:53:25 2021

@author: yannis.nissopoulos
"""

def process_query(instrument):
    instrument.write('*OPC')
    ESR_value = 0
    while ESR_value & 1 == 0:
        ESR_value = instrument.query('*ESR?')

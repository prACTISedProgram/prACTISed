#!/usr/bin/env python3
# -*- coding: utf-8 -*-                 

# ACTIS - Kd Determination Program
# July 8, 2022

#20220708 JL    NOTE: Prgram no longer generates temporary Excel files, description needs revision 

# This program extracts ACTIS titration data in a Microsoft Excel file (.xlsx) organized in a particular way*,
# and determines the signal (average peak height within a detection window) for each concentration and a corresponding
# R value for each concentration, then plots a binding isotherm for R vs Protein Concentration and performs
# non-linear curve fitting to calculate and output the Kd value for the experiment

# *The program code as shown by default below requires the Microsoft Excel workbook to be organized in the following format:
# - Data for each concentration must be contained in separate worksheets within the Excel file, with each sheet being named in the following format: "# µM".
# - The time intervals are written in column A.
# - Row 2 of each worksheet must be the first row containing data
# - Cell A1 is denoted as "raw time"
# - Cells A# are denoted as "Experiment #"
# - The signal measurement for each run is written in each corresponding column of the worksheet.
# - Important note: temporary Excel files called "TempFile6.xlsx" and "ASDF.xlsx" will be generated as part of the execution of this program
# should your device contain a file with important data in a file with this name, please note that the contents of this file will be
# overwritten. Therefore, either rename this pre-existing file and/or adjust the name of the temporary file created in the code
# below.

# Testing script execution time         ~ 6 seconds without Verbose
import time
start = time.time()

## Part 1 - Importing the required libraries and sub-libraries required below
import argparse                                   
import pathlib
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils import FORMULAE
from openpyxl.utils import get_column_letter, column_index_from_string
import math
from scipy.optimize import curve_fit

   
## Part 2 - Argument setup and parsing
parser = argparse.ArgumentParser(description = 'prACTISed! This program analyzes ACTIS data and extracts the Kd-value.')
parser.add_argument('inputfile', action = 'store', nargs = 1, type = str)
parser.add_argument('--version', help = 'prints version information', action = 'version', version = 'prACTISed written by Shiv Jain.')
parser.add_argument('-v', '--verbose', help = 'prints detailed output while analyzing', action = 'store_true')

args = vars(parser.parse_args())

## Part 3 - Locating the raw data file and establishing important inputs
fileName = args['inputfile'][0]                         # It's a list but we can only work with one file                           
if not pathlib.Path(fileName).is_file():
        print("Given file '%s' is not a file or does not exist." % fileName)
        exit(-1)

name = pathlib.PurePath(fileName).name                 

# Set explicitly the engine to use openpyxl, otherwise it might use xlrd, which has removed support for Excel's xlsx format (only supports the old binary format)
data = pd.read_excel(fileName, engine='openpyxl')
inputBook = load_workbook(fileName, data_only=True)         
inputBooknames = inputBook.sheetnames


# Confirm Inputs sheet with correct formatting before preceding
if "Inputs" in inputBook.sheetnames:                                            
    idealSheet = inputBook["Inputs"]
    percentage = float((idealSheet.cell(1,2).value)/100)
    numberOfConcs = (idealSheet.cell(3,2).value)
    injectionTime = int(idealSheet.cell(5,2).value)
    ligandConc = float(idealSheet.cell(7,2).value)

else:
        print("Inputfile Formatting Error: The script expects inputdata.xlsx to be in a certain format, see next section for generation or the provided idealinputs.xlsx as an example.")
        exit

# Temporary variables
concentration = []
signal = []
stddev = []
relstddev = []
Rvalue = []
Rstddev = []
forDF = [concentration,signal,stddev,relstddev,Rvalue,Rstddev]
DFnames = ["concentration","signal","stddev","relstddev","Rvalue","Rstddev"]


## Part 3 - Calculating signal information for each concentration
for x in range(1,int(numberOfConcs)+1):         
    conc1 = idealSheet.cell(x,5).value
    numberOfRuns = int(idealSheet.cell(x,7).value)
    
    avgSigConc = 0
    avgSigsRun = []

    # Reading in the whole data frame (= all runs for one particular concentration) and dropping all lines that are blank, i.e. that would produce "NaN"s
    data = pd.read_excel(fileName, str(conc1) + " µM", engine='openpyxl')
    data = data.dropna(how='all')


    # Calculate background signals for each run
    for runNumber in range(1, numberOfRuns+1):  
        ###### FOR VERBOSE
            #print("-------- " + str(conc1) + " µM" + " ---- Experimental Run " + str(runNumber) + " --------")

        xvalues = data['raw time']
        yvalues = data['Experiment ' + str(runNumber)]
        
        # All y-values before the injection time are considered background signal 
        background_yvalues = yvalues[xvalues < injectionTime]
        background_average = np.average(background_yvalues)
        background_stdev = np.std(background_yvalues)


        if args['verbose']:
            print("Background signal (first %d values): %.4f±%.4f" % (len(background_yvalues), background_average, background_stdev))
            exit(0)


        # First run at lowest concentration used to calculate peak time and time window
        if x==1 and runNumber==1:
                peakSignal = max(yvalues[xvalues>=injectionTime])
                peakIndex = yvalues[yvalues == peakSignal].index
                peakTime = xvalues[peakIndex]
                ###### FOR VERBOSE
                        #print ("The peak time is " + str(peakTime) + "seconds.")


        # Set time window parameters and determine the average signal within window for each run             
        windowLow = float(peakTime - (percentage * peakTime))
        windowHigh = float(peakTime + (percentage * peakTime))
        windowIndex = xvalues.between(windowLow,windowHigh)
        windowTimes = xvalues[windowIndex]
        windowSignals = yvalues[windowIndex]
        avgSigsRun.append(np.average(windowSignals))
        ###### FOR VERBOSE
                #print("The time window for Experimental Run " str(runNumber) "is " + str(windowLow) + "seconds to " + str(windowHigh) + "seconds.")
        

        # Graph the signal for each run and show time window
        #plt.scatter(xvalues, yvalues)
        #plt.xlabel('Propagation time (s)')
        #plt.ylabel('Fluorescence signal (RFU)')
        #plt.title('Graph')
        #plt.xlim(xmin=0)
        #plt.vlines(windowLow, 0, (float(yvalues.max()+ 0.05*float(yvalues.max()))), linestyles='dashed',color='black')
        #plt.vlines(windowHigh, 0, (float(yvalues.max()+ 0.05*float(yvalues.max()))), linestyles='dashed',color='black')     
        #plt.show()

    # Calculating average signal for each concentration, stdev and relative stdev        
    avgSigConc = np.average(avgSigsRun)
    avgSigConc_stdev = np.std(avgSigsRun)
    avgSigConc_relstdev = (avgSigConc_stdev/avgSigConc)*100

    concentration.append(conc1)
    signal.append(avgSigConc)    
    stddev.append(avgSigConc_stdev)
    relstddev.append(avgSigConc_relstdev)
    
###### FOR VERBOSE
    #df = pd.DataFrame (forDF).transpose()
    #df.columns = DFnames
    #print(df)
    

## Part 5 - Calculate R values and standard deviation of R values for each concentration
LowProt_sig = signal[0]
HighProt_sig = signal[numberOfConcs-1]
LowProt_stddev = stddev[0]
HighProt_stdDev = stddev[numberOfConcs-1]

for y in range(0, numberOfConcs):
        avgSigConc_R = (signal[y] - HighProt_sig)/ (LowProt_sig - HighProt_sig) 
        Rvalue.append(avgSigConc_R)

        avgSiglConc_Rstddev = ( 1/(LowProt_sig - HighProt_sig) * math.sqrt( (stddev[y]**2) + ((signal[y] - LowProt_sig)/ (LowProt_sig - HighProt_sig) * HighProt_stdDev)**2 +
                                                                       ((HighProt_sig - signal[y])/(LowProt_sig - HighProt_sig) * LowProt_stddev)**2))
        Rstddev.append(avgSiglConc_Rstddev)

###### FOR VERBOSE
        #df = pd.DataFrame (forDF).transpose()
        #df.columns = DFnames
        #print(df)
    
## Part 6 - Plotting the binding isotherm R vs P[0] with curve of best fit
# Plotting data points for each concentration
plt.scatter(concentration, Rvalue, color = 'black', label='Experiment')               
plt.errorbar(concentration, Rvalue, linestyle = 'None', yerr = Rstddev, ecolor = 'black', elinewidth=1, capsize=2, capthick=1)
plt.xscale("log")

# Define the Levenberg Marquardt algorithm
def LevenMarqu(x,a):          
    return -((a + x - ligandConc)/(2*ligandConc)) + ((((a + x - ligandConc)/(2*ligandConc))**2) + (a/ligandConc))**(0.5)

# Curve fitting and plotting the curve of best fit
popt, pcov = curve_fit(LevenMarqu, concentration, Rvalue)
error = np.sqrt(np.diag(pcov))

xFit = np.arange(0.0, float(conc1), min(concentration))
plt.plot(xFit, LevenMarqu(xFit, popt), 'r', label='Fit (Kd =%5.3f)' % tuple(popt), linewidth=2)       
plt.ylabel('R')
plt.xlabel('Concentration of BSA (μM)')
plt.legend()
plt.xscale("log")
plt.show(block=False)

# Statistics
residuals = Rvalue - LevenMarqu(concentration, *popt)
ss_res = np.sum(residuals**2)
ss_tot = np.sum((Rvalue - np.mean(Rvalue))**2)
r_squared = 1 - (ss_res/ss_tot)

chiSquared = sum((((Rvalue - LevenMarqu(concentration, *popt))**2) / LevenMarqu(concentration, *popt)))

# Returned/printed values
print("The Kd is " + str(popt) + " ± " + str(error) + " μM")
print("The R² value is " + str(r_squared))
print("The χ² value is " + str(chiSquared))

# Testing script execution time
end = time.time()
print("\n"+str(end-start))


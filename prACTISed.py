#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# ACTIS - Kd Determination Program
# August 6, 2021

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


## Part X - Argument setup and parsing
parser = argparse.ArgumentParser(
			description = 'prACTISed! This program analyzes ACTIS data and extracts the Kd-value.')

parser.add_argument('inputfile', action = 'store', nargs = 1, type = str)
parser.add_argument('--version', help = 'prints version information', action = 'version', 
            version = 'prACTISed written by Shiv Jain.')
parser.add_argument('-v', '--verbose', help = 'prints detailed output while analyzing', action = 'store_true')

args = vars(parser.parse_args())


## Part 2 - Locating the raw data file
workbook = Workbook()          # Establishing the initial workbook which will be used
worksheet = workbook.active

fileName = args['inputfile'][0] # It's a list but we only can work with one file
if not pathlib.Path(fileName).is_file():
    print("Given file '%s' is not a file or does not exist." % fileName)
    exit(-1)

# Set explicitly the engine to use openpyxl - otherwise it might use xlrd, which has removed
# support for Excel's xlsx format (only supports the old binary format)
data = pd.read_excel(fileName, engine='openpyxl')

idealBook =  load_workbook(fileName, data_only=True)          # Locating the peak position from the cell in the temporary Excel file
idealSheet = idealBook["Inputs"]

name = pathlib.PurePath(fileName).name


## Part 3 - Establishing important inputs
percentage = (idealSheet.cell(1,2).value)/100
numberOfConcs = str(idealSheet.cell(3,2).value)
injectionTime = idealSheet.cell(5,2).value

def stdev(data):          # Defining a function and performing the function to calculate the standard deviation of the data
        n=len(data)
        mean=sum(data)/n
        deviations=[(x-mean)**2 for x in data]
        variance=sum(deviations)/(n-1)
        stdev=math.sqrt(variance)
        return stdev
    
    
noise = []


## Part 3 - Data collection
for x in range(1,int(numberOfConcs) + 1):          
    conc1 = idealSheet.cell(x,5).value    
    protConc1 = conc1    
    runNumber =1 
    numberOfRuns = idealSheet.cell(x,7).value    
    
    totalAverage = 0          # Temporary variables and classes used in determining the average peak height within the detection window
    averages = []    
    
    expRunNumber = 1   

    peakOnsetTimes = []

    # Reading in the whole data frame (= all runs for one particular concentration)
    # and dropping all lines that are blank, i.e. that would produce "NaN"s
    data = pd.read_excel(fileName, str(conc1) + " µM", engine='openpyxl')
    data = data.dropna(how='all')

    for runNumber in range(1, numberOfRuns+1):
        print("")
        print("-------- " + str(conc1) + " µM" + " ---- Experimental Run " + str(runNumber) + " --------")

        xvalues = data['raw time']
        yvalues = data['Experiment ' + str(runNumber)]

        # All y-values before the injection time are considered background signal 
        background_yvalues = yvalues[xvalues < injectionTime]
        background_average = np.average(background_yvalues)
        background_stdev = np.std(background_yvalues)

        if args['verbose']:
            print("Background signal (first %d values): %.4f±%.4f" % (len(background_yvalues), 
                                                                      background_average,
                                                                      background_stdev))

        exit(0)
        # TODO: Modify rest of the code to get rid of the temporary worksheets, adding verbose statements, etc.

    
        wb = Workbook()          # Selecting and activating the particular workbook/worksheet chosen for the code to run
        sheet = wb.active

        runColumn = get_column_letter(int(runNumber) + 1)          # Locating column A of the spreadsheet to obtain time data
        columnA = get_column_letter(1)

        
        
        xValues = pd.DataFrame(data,columns = ['raw time'])          # Used to plot the separagram for each run as well as the bounds for the detection window selected
        xValues = (xValues)      
        yValues = pd.DataFrame(data,columns = ['Experiment ' + str(runNumber)])
        
        maxVal = float(yValues.max())
        maxValForGraph = maxVal + (0.05*maxVal)
        
        print("")
        print("-------- " + str(conc1) + " µM" + " ---- Experimental Run " + str(expRunNumber) + " --------")
        expRunNumber = expRunNumber + 1        

        plt.scatter(xValues, yValues)          # All of these lines, including those below can be activated in order to display each plot
        plt.xlabel('Propagation time (s)')
        plt.ylabel('Fluorescence signal (RFU)')
        plt.title('Graph')
        plt.xlim(xmin=0)
        #plt.ylim(ymax=11)
        #plt.show()    
        
        #########################################
        excelRows = str(1154) # ADJUSTABLE
        
        injectionTime = idealSheet.cell(5,2).value     
        excelBook =  load_workbook(fileName, data_only=True)          # Locating the peak position from the cell in the temporary Excel file
        wsheet = excelBook[str(conc1) + " µM"]
            
        numOfTimes = 0
        timeSpecificTotal = 0
        backgroundSigs = []       
        
        for variable1 in range(2, int(excelRows)+1):
            timeColumn = "A" + str(variable1)
            timeSpecific = wsheet[timeColumn].value
            
            if timeSpecific < injectionTime:
                backgroundCell = str(runColumn) + str(variable1)                
                background1 = wsheet[backgroundCell].value
                backgroundSigs.append(background1)
                
        sumOfBackgroundSigs = sum(backgroundSigs)
        lenOfBackgroundSigs = len(backgroundSigs)
        avgOfBackgroundSigs = sumOfBackgroundSigs/lenOfBackgroundSigs
        stdevOfBackgroundSigs = stdev(backgroundSigs)
                
        for rowNumber1 in range(lenOfBackgroundSigs+2,int(excelRows)+1):
            firstSig = wsheet.cell(rowNumber1,int(runNumber)+1).value 
            factor = 5 # ADJUSTABLE
            boundary = avgOfBackgroundSigs + (factor*stdevOfBackgroundSigs)
            
            if firstSig > boundary:
                onsetTime = wsheet.cell(rowNumber1,1).value        
                break
            
        peakOnsetTimes.append(onsetTime)
        
        if x == 1 and runNumber == 1:
            sheet[columnA + "2"] = "=INDEX('[" + name + "]" + str(conc1) + " µM" + "'!$A2:$A" + (excelRows) + ",MATCH(MAX('[" + name + "]" + str(conc1) + " µM" + "'!" + str(runColumn) + "2:" + str(runColumn) + str(excelRows) + "), '[" + name + "]" + str(conc1) + " µM" + "'!" + str(runColumn) + "2:" + str(runColumn) + str(excelRows) + ",0))"          # This function is put into the temporary Excel file in order to determine the average peak height within the detection window            
            wb.save("TempFile6.xlsx")

            def pause():        # A function in python is defined that pauses the program until the user enters "space" (i.e., presses Spacebar on a keyboard); this key can be altered to anything desired
                print("Press ENTER to continue")
                input()

            print("")
            print("The program is currently paused.")        # This section of code informs the user that they must open up the temporary file, save it, close it and press space back on Python to allow the program to proceed
            print("Please open up \"TempFile6.xlsx\", save it and close it.")
            print("")
            pause()

            book =  load_workbook("TempFile6.xlsx", data_only=True)          # Locating the peak position from the cell in the temporary Excel file
            sheet = book.active        
        
            maxTime = sheet.cell(2,1).value
       
            recommendedTimeDiff = maxTime - onsetTime
            
        recommendedTime = onsetTime + recommendedTimeDiff     
        peakPosition = recommendedTime
        ####################################################
    
        peakPositionPlusX = peakPosition + (float(percentage) * peakPosition)          # Determining the bounds of the detection window
        peakPositionMinusX = peakPosition - (float(percentage) * peakPosition)                           

#############################################################################################
        data = pd.read_excel(fileName, str(conc1) + " µM", engine='openpyxl')
        
        wb = load_workbook(fileName)          # Selecting and activating the particular workbook/worksheet chosen for the code to run
        sheet = wb[str(conc1) + " µM"]
        
        periods = int(excelRows)          # This indicates the number of rows in the titration data Excel file containing signal data; NOTE: This number may be changed for different data sets
        count = 0
        peakHeightTotal = 0

        for variable in range(2, periods+1):
            cell1 = "A" + str(variable)
            times = sheet[cell1].value

            if times >= peakPositionMinusX and times <= peakPositionPlusX:          # Determining the signal values within the detection window and calculating an average
                count = count + 1
        
                cell2 = str(runColumn) + str(variable)
                peakHeightAtTime = sheet[cell2].value
        
                peakHeightTotal = peakHeightTotal + peakHeightAtTime
        
            variable = variable+1

############################################################################################
        xValues = pd.DataFrame(data,columns = ['raw time'])          # Used to plot the separagram for each run as well as the bounds for the detection window selected
        xValues = (xValues)
        yValues = pd.DataFrame(data,columns = ['Experiment ' + str(runNumber)])

        plt.scatter(xValues, yValues)          # All of these lines, including those below can be activated in order to display each plot
        plt.xlabel('Propagation time (s)')
        plt.ylabel('Fluorescence signal (RFU)')
        plt.title('Graph')
        plt.xlim(xmin=0)

        peakPositionMinusXSec = (peakPositionMinusX)
        peakPositionPlusXSec = (peakPositionPlusX)
        plt.vlines(peakPositionMinusXSec, 0, maxValForGraph, linestyles='dashed',color='black')
        plt.vlines(peakPositionPlusXSec, 0, maxValForGraph, linestyles='dashed',color='black')     
   
        #plt.show()

#############################################################################################
        averageOfRange = peakHeightTotal/count          # Determining the average peak height within the detection window for all of the runs for each concentration and outputting it
        totalAverage = totalAverage + averageOfRange
        averages.append(averageOfRange)
#######################################################################################
    print("")
    print("-------- " + str(conc1) + " µM --------")
    print("Signal (total average): " + str(totalAverage/numberOfRuns))

    def relstdev(data):          # Defining a function and performing the function to calculate the relative standard deviation of the data
        n=len(data)
        mean=sum(data)/n
        deviations = [(x - mean)**2 for x in data]
        variance = sum(deviations)/(n-1)
        stdev = math.sqrt(variance)
        relstdev = (stdev/averageOfRange)*100
        return relstdev
    
    if numberOfRuns == 1:
        relativeStdDev = 0
        standardDeviation = 0
        
    else:
        relativeStdDev = relstdev(averages)
        standardDeviation = stdev(averages)
  
    print("Standard deviation: " + str((relativeStdDev/100)*averageOfRange))
    print("Relative standard deviation: " + str(relativeStdDev) + "%")
    print("Peak onset times: " + str(peakOnsetTimes) + " seconds")
    print("")

    
## Part 4 - Preparing the Excel File with R values
    worksheet["A" + str(x+1)] = str(conc1) + " µM"          # Organizing the data contained within the Excel file (temporary "ASDF.xlsx" file)
    worksheet["H" + str(x+1)] = float(protConc1)
    worksheet["C" + str(x+1)] = float(totalAverage/numberOfRuns)
    worksheet["E" + str(x+1)] = standardDeviation
    worksheet["F" + str(x+1)] = relativeStdDev    
    worksheet["I" + str(x+1)] = "=(C" + (str(x+1)) + "-$C$" + (str(int(numberOfConcs)+1)) + ")/($C$2" + "-$C$" + (str(int(numberOfConcs)+1)) + ")"
    worksheet["J" + str(x+1)] = "=1/($C$2 - $C$" + str(int(numberOfConcs)+1) + ")*SQRT(E" + str(x+1) + "^2+((C" + str(x+1) + "-$C$2)/($C$2-$C$" + str(int(numberOfConcs)+1) + ")*$E$" + str(int(numberOfConcs)+1) + ")^2+(($C$" + str(int(numberOfConcs)+1) + "-C" + str(x+1) + ")/($C$2-$C$" + str(int(numberOfConcs)+1) + ")*$E$2)^2)"
       
worksheet["A1"] = "Concentration (µM)"
worksheet["C1"] = "Signal"
worksheet["E1"] = "StDev"
worksheet["F1"] = "Relative StDev (%)"
worksheet["H1"] = "[P]0 (µM)"
worksheet["I1"] = "R expr"
worksheet["J1"] = "σ R expr"

from openpyxl.styles import Font
bold_font = Font(bold=True)
worksheet["A1"].font = bold_font
worksheet["C1"].font = bold_font
worksheet["E1"].font = bold_font
worksheet["F1"].font = bold_font
worksheet["H1"].font = bold_font
worksheet["I1"].font = bold_font
worksheet["J1"].font = bold_font
    
workbook.save("ASDF.xlsx")

##############################################################################################
## Part 5 - Performing non-linear curve fitting

print("")
print("The program is currently paused.")        # This section of code informs the user that they must open up the temporary file, save it, close it and press space back on Python to allow the program to proceed
print("Please open up \"ASDF.xlsx\", save it and close it.")
print("Once this is done, press Space to resume the program")
print("")
pause()

xData = []          # Temporary classes to hold data for concentrations and R values
yData = []
yError = []

excelFile = load_workbook("ASDF.xlsx", data_only=True)          # Opening up the temporary file containing the concentration and R values
activeBook = excelFile.active

for a in range(2,(int(numberOfConcs)+2)):          # Storing the R values and concentration values in the temporary classes created above
    xValues = activeBook.cell(a,8).value
    yValues = activeBook.cell(a,9).value
    yErrors = activeBook.cell(a,10).value
    xData.append(xValues)
    yData.append(yValues)
    yError.append(yErrors)
      
##############
minimumConc = min(xData)
##############

def func(x,a):          # Defining a function to perform non-linear curve fitting according to the Levenberg Marquardt algorithm
    return -((a + x - ligandConc)/(2*ligandConc)) + ((((a + x - ligandConc)/(2*ligandConc))**2) + (a/ligandConc))**(0.5)
    
plt.scatter(xData,yData, color = 'black', label='Experiment')          # Plotting the R vs concentration data as points with error bars
plt.errorbar(xData, yData, linestyle = 'None', yerr = yError, ecolor = 'black', elinewidth=1, capsize=2, capthick=1)
plt.xscale("log")          # Creating a logarithmic x-axis scale

ligandConc = idealSheet.cell(7,2).value
popt, pcov = curve_fit(func, xData, yData)          # Performing curve fitting

##########################
error = np.sqrt(np.diag(pcov))          # Calculates the uncertainty associated with the curve fitting process
##############################


residuals = yData - func(xData, *popt)
ss_res = np.sum(residuals**2)
ss_tot = np.sum((yData - np.mean(yData))**2)
r_squared = 1 - (ss_res/ss_tot)

expected = func(xData, *popt)
observed = yData

chiSquared = (((observed - expected)**2)/expected).sum()

print("")          # Outputting the Kd value and uncertainty associated with it, both determined based on the non-linear curve fit

xFit = np.arange(0.0, float(conc1), minimumConc)

plt.plot(xFit, func(xFit, popt), 'r', label='Fit (Kd =%5.3f)' % tuple(popt), linewidth=2)           # Plotting the curve of best fit
plt.ylabel('R')
plt.xlabel('Concentration of BSA (μM)')
plt.legend()
plt.xscale("log")
plt.show()
print("The Kd is " + str(popt) + " ± " + str(error) + " μM")
print("The R² value is " + str(r_squared))
print("The χ² value is " + str(chiSquared))




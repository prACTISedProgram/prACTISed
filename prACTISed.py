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
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils import FORMULAE
from openpyxl.utils import get_column_letter, column_index_from_string
import keyboard
import math
from scipy.optimize import curve_fit

## Part 2 - Locating the raw data file
workbook = Workbook()          # Establishing the initial workbook which will be used
worksheet = workbook.active

fileName = (input("File Path: "))          # Gathering input from the user on the file name containing titration data
#data = pd.read_excel(r'C:\Users\Rajin\Desktop' + "\\" + fileName)          # Locating the file on the computer - NOTE: This must be changed for different devices
data = pd.read_excel(fileName, engine='openpyxl')

idealBook =  load_workbook(fileName, data_only=True)          # Locating the peak position from the cell in the temporary Excel file
idealSheet = idealBook["Inputs"]


from pathlib import PurePath
  

location = PurePath(fileName)
  

name = location.name



## Part 3 - Establishing important inputs
percentage = (idealSheet.cell(1,2).value)/100
#print(percentage)

#percentage = input("What +/- decimal range (e.g. 0.04)? ")          # Gathering input from the user on what detection window would like to be used

numberOfConcs = str(idealSheet.cell(3,2).value)
#print(numberOfConcs)
#numberOfConcs = int(input("How many concentrations? "))          # Gathering input from the user on how many initial protein concentrations will be used
print("")


def stdev(data):          # Defining a function and performing the function to calculate the standard deviation of the data
        n=len(data)
        mean=sum(data)/n
        deviations=[(x-mean)**2 for x in data]
        variance=sum(deviations)/(n-1)
        stdev=math.sqrt(variance)
        return stdev
    
    
noise = []


for x in range(1,int(numberOfConcs) + 1):          # For each concentration, the user is asked what the initial protein concentration (i.e. [P]0) is
    conc1 = idealSheet.cell(x,5).value
    #print(conc1)
    #conc1 = input("What is concentration #" + str(x) + " (in µM)? ")
    protConc1 = conc1
    #protConc1 = input("What is protein concentration #" + str(x) + " (in µM)? ")

    runNumber =1 
    numberOfRuns = idealSheet.cell(x,7).value
    #numberOfRuns = int(input("How many runs? "))          # Gathering input for how many runs of each particular protein concentration are performed

    
## Part 3 - Data collection
    totalAverage = 0          # Temporary variables and classes used in determining the average peak height within the detection window
    averages = []
    
    
    expRunNumber = 1
    


    peakOnsetTimes = []

    for runNumber in range(1, numberOfRuns+1):
    
        wb = Workbook()          # Selecting and activating the particular workbook/worksheet chosen for the code to run
        sheet = wb.active

        runColumn = get_column_letter(int(runNumber) + 1)          # Locating column A of the spreadsheet to obtain time data
        columnA = get_column_letter(1)

        
        #data = pd.read_excel(r'C:\Users\Rajin\Desktop' + "\\" + fileName, conc1 + " µM")
        data = pd.read_excel(fileName, str(conc1) + " µM", engine='openpyxl')
        
        xValues = pd.DataFrame(data,columns = ['raw time'])          # Used to plot the separagram for each run as well as the bounds for the detection window selected
        xValues = (xValues)
        yValues = pd.DataFrame(data,columns = ['Experiment ' + str(runNumber)])
        
        #
        maxVal = float(yValues.max())
        maxValForGraph = maxVal + (0.05*maxVal)
        #print(maxValForGraph)
        
        
        
        
        #expRunNumber = 1
        print("")
        print("-------- " + str(conc1) + " µM" + " ---- Experimental Run " + str(expRunNumber) + " --------")
        expRunNumber = expRunNumber + 1
        

        plt.scatter(xValues, yValues)          # All of these lines, including those below can be activated in order to display each plot
        plt.xlabel('Propagation time (s)')
        plt.ylabel('Fluorescence signal (RFU)')
        plt.title('Graph')
        plt.xlim(xmin=0)
        #plt.ylim(ymax=11)
        plt.show()
        
        
        
        
        #########################################
        excelRows = str(1154) # ADJUSTABLE
        
        
        
        
        injectionTime = idealSheet.cell(5,2).value
        #print(injectionTime)
        #injectionTime = int(input("How long is injection time (in seconds, enter an integer)? "))
        
        excelBook =  load_workbook(fileName, data_only=True)          # Locating the peak position from the cell in the temporary Excel file
        wsheet = excelBook[str(conc1) + " µM"]
        
        
        
#        rowNumb = 2
#        #timeColumn = 1
#        timeColumn = wsheet.cell(rowNumb, 1).value
        
#        print(timeColumn)
#        #if timeColumn < injectionTime:
            
        
#        for timeColumn in range(injectionTime):
#            timeColumn = wsheet.cell(rowNumb, 1).value
#            print(timeColumn)
#            rowNumb = rowNumb + 1
    
        numOfTimes = 0
        timeSpecificTotal = 0
        backgroundSigs = []
        
        
        
        for variable1 in range(2, int(excelRows)+1):
            timeColumn = "A" + str(variable1)
            timeSpecific = wsheet[timeColumn].value

            #print(timeSpecific)
            
            if timeSpecific < injectionTime:
                backgroundCell = str(runColumn) + str(variable1)
                #print(backgroundCell)
                
                background1 = wsheet[backgroundCell].value
                #print(background1)
                backgroundSigs.append(background1)
                
        #print(backgroundSigs)
        sumOfBackgroundSigs = sum(backgroundSigs)
        lenOfBackgroundSigs = len(backgroundSigs)
        #print(sumOfBackgroundSigs)
        #print(lenOfBackgroundSigs)
        avgOfBackgroundSigs = sumOfBackgroundSigs/lenOfBackgroundSigs
        stdevOfBackgroundSigs = stdev(backgroundSigs)
        
        #print(avgOfBackgroundSigs)
        #print(stdevOfBackgroundSigs)
        
        
        #print("lenOfBackgroundSigs " + str(lenOfBackgroundSigs))
        #print("excelRows " + str(excelRows))
        
        for rowNumber1 in range(lenOfBackgroundSigs+2,int(excelRows)+1):
            #print(rowNumber1)
            firstSig = wsheet.cell(rowNumber1,int(runNumber)+1).value
            #print(firstSig)

            factor = 5 # ADJUSTABLE
            boundary = avgOfBackgroundSigs + (factor*stdevOfBackgroundSigs)
            
            if firstSig > boundary:
                #print("done")
                #print("firstSig " + str(firstSig))
                #print("rowNumber1 " + str(rowNumber1))
                #print("boundary " + str(boundary))
                
                onsetTime = wsheet.cell(rowNumber1,1).value
                
                #print("Peak onset time: " + str(onsetTime))        # ADJUSTABLE as of Aug 6; can remove "#" to show peak onset time for each run
        
                break
            
        peakOnsetTimes.append(onsetTime)
   
        
        
        
        if x == 1 and runNumber == 1:
            #sheet[columnA + "2"] = "1"
            sheet[columnA + "2"] = "=INDEX('[" + name + "]" + str(conc1) + " µM" + "'!$A2:$A" + (excelRows) + ",MATCH(MAX('[" + name + "]" + str(conc1) + " µM" + "'!" + str(runColumn) + "2:" + str(runColumn) + str(excelRows) + "), '[" + name + "]" + str(conc1) + " µM" + "'!" + str(runColumn) + "2:" + str(runColumn) + str(excelRows) + ",0))"          # This function is put into the temporary Excel file in order to determine the average peak height within the detection window
            #sheet[columnA + "2"] = "=INDEX('[" + fileName + "]" + conc1 + " µM" + "'!$A2:$A" + (excelRows) + ",MATCH(MAX('[" + fileName + "]" + conc1 + " µM" + "'!" + str(runColumn) + "2:" + str(runColumn) + (excelRows) + "), '[" + fileName + "]" + conc1 + " µM" + "'!" + str(runColumn) + "2:" + str(runColumn) + (excelRows) + ",0))"          # This function is put into the temporary Excel file in order to determine the average peak height within the detection window
            wb.save("TempFile6.xlsx")

            def pause():        # A function in python is defined that pauses the program until the user enters "space" (i.e., presses Spacebar on a keyboard); this key can be altered to anything desired
                while True:
                    if keyboard.read_key() == 'space':
                        break

            print("")
            print("The program is currently paused.")        # This section of code informs the user that they must open up the temporary file, save it, close it and press space back on Python to allow the program to proceed
            print("Please open up \"TempFile6.xlsx\", save it and close it.")
            print("Once this is done, press Space to resume the program")
            print("")
            pause()

            book =  load_workbook("TempFile6.xlsx", data_only=True)          # Locating the peak position from the cell in the temporary Excel file
            sheet = book.active
        
        
            maxTime = sheet.cell(2,1).value
            #print(peakPosition)
        
        
        
            recommendedTimeDiff = maxTime - onsetTime
            
        recommendedTime = onsetTime + recommendedTimeDiff
                
        
        #peakPosition = float(input("What time for peak? (Recommended: " + str(recommendedTime) + " seconds) "))
        peakPosition = recommendedTime
        ####################################################
    
        peakPositionPlusX = peakPosition + (float(percentage) * peakPosition)          # Determining the bounds of the detection window
        peakPositionMinusX = peakPosition - (float(percentage) * peakPosition)                           

#############################################################################################

        data = pd.read_excel(fileName, str(conc1) + " µM", engine='openpyxl')
        #data = pd.read_excel(r'C:\Users\Rajin\Desktop' + "\\" + fileName, conc1 + " µM")          # Locating the signal values from the titration data Excel file, NOTE: This file path must be changed for different devices

        wb = load_workbook(fileName)          # Selecting and activating the particular workbook/worksheet chosen for the code to run
        #sheet = wb.active
        sheet = wb[str(conc1) + " µM"]
        

        periods = int(excelRows)          # This indicates the number of rows in the titration data Excel file containing signal data; NOTE: This number may be changed for different data sets
        count = 0
        peakHeightTotal = 0

        for variable in range(2, periods+1):
            cell1 = "A" + str(variable)
            times = sheet[cell1].value

            #print(times)
            #print(peakPositionMinusX)
            #print(peakPositionPlusX)
    
            if times >= peakPositionMinusX and times <= peakPositionPlusX:          # Determining the signal values within the detection window and calculating an average
                count = count + 1
        
                cell2 = str(runColumn) + str(variable)
                peakHeightAtTime = sheet[cell2].value
        
                peakHeightTotal = peakHeightTotal + peakHeightAtTime
        
            variable = variable+1

        #print(peakHeightTotal)
        #print(count)
    
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
        #plt.vlines(peakPositionMinusXSec, 0, 10, linestyles='dashed',color = 'black')
        #plt.vlines(peakPositionPlusXSec, 0, 10, linestyles='dashed',color = 'black')

        
        #plt.vlines(peakPositionMinusXSec, 0, 0.0001, linestyles='dashed',color = 'black')
        #plt.vlines(peakPositionPlusXSec, 0, 0.0001, linestyles='dashed',color = 'black')
        plt.vlines(peakPositionMinusXSec, 0, maxValForGraph, linestyles='dashed',color='black')
        plt.vlines(peakPositionPlusXSec, 0, maxValForGraph, linestyles='dashed',color='black')
        

        plt.show()
#############################################################################################
        averageOfRange = peakHeightTotal/count          # Determining the average peak height within the detection window for all of the runs for each concentration and outputting it
        #print("The average Peak Height for run " + str(runNumber) + " is "+ str(averageOfRange))
    
        totalAverage = totalAverage + averageOfRange
        averages.append(averageOfRange)
#######################################################################################
    print("")
    #print(averageOfRange)
    print("-------- " + str(conc1) + " µM --------")
    #print("The total average is " + str(totalAverage/numberOfRuns))
    #print("totalAverage " + str(totalAverage))
    #print("numberOfRuns " + str(numberOfRuns))
    print("Signal (total average): " + str(totalAverage/numberOfRuns))
    #print(averages)

    def relstdev(data):          # Defining a function and performing the function to calculate the relative standard deviation of the data
        n=len(data)
        mean=sum(data)/n
        deviations = [(x - mean)**2 for x in data]
        variance = sum(deviations)/(n-1)
        stdev = math.sqrt(variance)
        relstdev = (stdev/averageOfRange)*100
        return relstdev
        #return stdev
    
    
    
    if numberOfRuns == 1:
        relativeStdDev = 0
        standardDeviation = 0
        
    else:
        relativeStdDev = relstdev(averages)
        standardDeviation = stdev(averages)
        
#    relativeStdDev = relstdev(averages)

    #relativeStdDev = (float(stdev(averages)))/averageOfRange
    #print("and the relative standard deviation is " + str(relativeStdDev) + "%")          # Outputting the relative standard deviation for all of the runs for each concentration
    print("Standard deviation: " + str((relativeStdDev/100)*averageOfRange))
    print("Relative standard deviation: " + str(relativeStdDev) + "%")
    print("Peak onset times: " + str(peakOnsetTimes) + " seconds")
    print("")
    
    

#    standardDeviation = stdev(averages)
    
    #workbook =  load_workbook("ASDF.xlsx")
    
    #workbook.save("ASDF.xlsx")

    #workbook =  load_workbook("ASDF.xlsx")
    
## Part 4 - Preparing the Excel File with R values
    worksheet["A" + str(x+1)] = str(conc1) + " µM"          # Organizing the data contained within the Excel file (temporary "ASDF.xlsx" file)
    worksheet["H" + str(x+1)] = float(protConc1)
    worksheet["C" + str(x+1)] = float(totalAverage/numberOfRuns)
    worksheet["E" + str(x+1)] = standardDeviation
    worksheet["F" + str(x+1)] = relativeStdDev
    
#    print(str(int(numberOfConcs)+1))
    
    worksheet["I" + str(x+1)] = "=(C" + (str(x+1)) + "-$C$" + (str(int(numberOfConcs)+1)) + ")/($C$2" + "-$C$" + (str(int(numberOfConcs)+1)) + ")"
    worksheet["J" + str(x+1)] = "=1/($C$2 - $C$" + str(int(numberOfConcs)+1) + ")*SQRT(E" + str(x+1) + "^2+((C" + str(x+1) + "-$C$2)/($C$2-$C$" + str(int(numberOfConcs)+1) + ")*$E$" + str(int(numberOfConcs)+1) + ")^2+(($C$" + str(int(numberOfConcs)+1) + "-C" + str(x+1) + ")/($C$2-$C$" + str(int(numberOfConcs)+1) + ")*$E$2)^2)"
    #workbook.save("ASDF.xlsx")
    
#workbook =  load_workbook("ASDF.xlsx")        
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
    #print(xValues)
    yValues = activeBook.cell(a,9).value
    yErrors = activeBook.cell(a,10).value
    xData.append(xValues)
    yData.append(yValues)
    yError.append(yErrors)
       
#print(xData)
#print(yData)

##############
#print(xData)
minimumConc = min(xData)
#print(minimumConc)
##############

def func(x,a):          # Defining a function to perform non-linear curve fitting according to the Levenberg Marquardt algorithm
    #ligandConc = float(input("What is the initial ligand concentration? "))
    #print(ligandConc)
    return -((a + x - ligandConc)/(2*ligandConc)) + ((((a + x - ligandConc)/(2*ligandConc))**2) + (a/ligandConc))**(0.5)
    
plt.scatter(xData,yData, color = 'black', label='Experiment')          # Plotting the R vs concentration data as points with error bars
plt.errorbar(xData, yData, linestyle = 'None', yerr = yError, ecolor = 'black', elinewidth=1, capsize=2, capthick=1)
plt.xscale("log")          # Creating a logarithmic x-axis scale

ligandConc = idealSheet.cell(7,2).value
#print(ligandConc)
#ligandConc = float(input("What is the initial ligand concentration (in µM)? "))          # Inputting the initial ligand concentration (i.e. [L]0) for the non-linear curve fitting, Kd equation
popt, pcov = curve_fit(func, xData, yData)          # Performing curve fitting

##########################
error = np.sqrt(np.diag(pcov))          # Calculates the uncertainty associated with the curve fitting process
#print(error)
##############################


residuals = yData - func(xData, *popt)
ss_res = np.sum(residuals**2)
ss_tot = np.sum((yData - np.mean(yData))**2)
r_squared = 1 - (ss_res/ss_tot)
#print(ss_res)
#print(ss_tot)
#print(ss_res/ss_tot)
#print("")

expected = func(xData, *popt)
#print(expected)
observed = yData
#print(observed)

chiSquared = (((observed - expected)**2)/expected).sum()
#print(chiSquared)



print("")          # Outputting the Kd value and uncertainty associated with it, both determined based on the non-linear curve fit


# xFit = np.arange(0.0, float(conc1), 0.10)          # Creating the x data for the non-linear curve fit equation (starts at 0, ends at the highest protein concentration (i.e. conc1) and in steps of 0.10)
#xFit = np.arange(0.0, 25, 1)
xFit = np.arange(0.0, float(conc1), minimumConc)

plt.plot(xFit, func(xFit, popt), 'r', label='Fit (Kd =%5.3f)' % tuple(popt), linewidth=2)           # Plotting the curve of best fit
plt.ylabel('R')
plt.xlabel('Concentration of BSA (μM)')
plt.legend()
plt.xscale("log")
plt.show()
print("The Kd is " + str(popt) + " ± " + str(error) + " μM")
#print("The R^2 value is " + str(r_squared))
print("The R² value is " + str(r_squared))
print("The χ² value is " + str(chiSquared))

#20220725 JL    NOTE: Switched to PySimpleGUI, added data analysis and working file preparation as functions (compensation & report WIP)

import PySimpleGUI as sg

import os
import pathlib
import glob

from PIL import Image, ImageTk

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils import FORMULAE
from openpyxl.utils import get_column_letter, column_index_from_string
import math
from scipy.optimize import curve_fit
from datetime import date
import time
from natsort import natsorted

sg.theme('DarkBlue3')
    
inputCol = [
    [sg.Frame('Fluidics Experimental Parameters',
              [[sg.Text('Propagation flow rate (µM/min)', size=(25,1), key ='propFlow'),
                sg.Input(size=(10,1), key='propFlow_val')],
               [sg.Text('Injection flow rate (µM/min)', size=(25,1),key='injectFlow'),
                sg.Input(size=(10,1), key='injectFlow_val')],
               [sg.Text('Injection time (s)', size=(25,1), key='injectTime'),
                sg.Input(size=(10,1), key='injectTime_val')],
               [sg.Text('Separation capillary length (cm)', size=(25,1), key='sepLength'),
                sg.Input(size=(10,1), key='sepLength_val')],
               [sg.Text('Injection loop length (cm)', size=(25,1), key='injectLength'),
                sg.Input(size=(10,1), key='injectLength_val')]]
              )],


    [sg.Frame('Concentrations',
              [[sg.Text('Protein name', size=(25,1), key ='protName'),
                sg.Input(size=(10,1), key='protName_val')],
               [sg.Text('Ligand name', size=(25,1), key='ligName'),
                sg.Input(size=(10,1), key='ligName_val')],
               [sg.Text('Protein concentrations [P]0 \n (with decimal and separated by commas)', key='protConcs')],
               [sg.Input(size=(37,1), key='protConcs_val')],
               [sg.Text('Initial ligand concentration [L]0', size=(25,1), key='ligConc'),
                sg.Input(size=(10,1), key='ligConc_val')]]
              )],
    
    [sg.Frame('Data Analyis Parameters',
              [[sg.Text('Compensation procedure \n (recommended for MS data)', key ='comp'),
                sg.Radio('Yes', 'compChoice', key ='compYes', enable_events=True),
                sg.Radio('No', 'compChoice', key ='compNo', enable_events=True)],
               [sg.Text('[P]0 reference for normalization', size=(30,1), key='normConc', visible=False),
                sg.Input(size=(10,1), key='normConc_val', visible=False)],
               [sg.Text('Window width (%)', size=(30,1), key='window'),
                sg.Input(size=(10,1), key='window_val')],
               [sg.Text('Determination of peak', size=(17,1), key ='peak'),
                sg.Radio('Manual', 'peakChoice', key ='peakM', enable_events=True),
                sg.Radio('Program', 'peakChoice', key ='peakP', enable_events=True)],
               [sg.Text('Peaks for concentrations [P]0 \n (same order as [P]0, separated by commas)', key='manPeaks', visible=False)],
               [sg.Input(size=(42,1), key='manPeaks_val', visible=False)],
               [sg.Text('Specify [P]0 used to determine peak', size=(30,1), key='progPeak', visible=False),
                sg.Input(size=(10,1), key='progPeak_val', visible=False)]]
              )],
    
    [sg.Button('Calculate Kd', key='calculate'),
     sg.Button('Report', key='report', visible=False)]
    ]
        
outputCol= [
    [sg.Table(values=[],headings=['prACTISed output'], key='Kd', hide_vertical_scroll=True, def_col_width=20, auto_size_columns=False)],

    [sg.Table(values=[], headings=['Conc','Avg Signal', 'Std Dev', 'Rel Std Dev', 'R value', 'R Std Dev'], key='summary')],

    [sg.Frame('Graphs',
              [[sg.Image(key='graphImage')],
               [sg.Button('Prev', key='back'),
                sg.Button('Next', key='fwd')]]
               )]           
    ]
    

layout = [
    [sg.Text('File path', key ='filePath'),
     sg.Input(size=(50,1), key='filePath_val')],
    
    [sg.Column(inputCol), sg.Column(outputCol, visible=False, key='out')]
    ]

window = sg.Window("prACTISed", layout, finalize=True)


######### DEFINE FUNCTIONS ##############

# Display image from file path
def load_image(path,window):
    img = Image.open(path)
    img.thumbnail((300, 300))
    photo_img = ImageTk.PhotoImage(img)
    window['graphImage'].update(data=photo_img)

location = 0


# Generate working file from directory containing raw data (.txt or .asc)
def workingfileprep(inputPath, propFlow, injectFlow, injectTime, sepLength, injectLength,
                    proteinName, ligandName, proteinConcs, ligandConc, compYN, normalConc,
                    windowWidth, peakDet, manualPeaks, peakConc):


    # Testing script execution time
    start = time.time()


    d = {}

    def isfloat(num):
        try:
            float(num)
            return True
        except ValueError:
            return False

    workingFileName = "%s_%s.xlsx" % (proteinName, date.today())         

        ### Reading in raw data files ###
    for file in os.listdir(inputPath):

        if file.endswith((".txt", ".asc")):
            # Extract concentration prefix
            molar = file.find("M")
            prefix = file[molar-1]

            # Extract concentration and run number from file name (see naming conventions)
            conc = float(file.partition(prefix)[0])

            if file[-5].isnumeric:              
                runNumber= int(file[-6:-4])         

            else:
                runNumber=1

            # Extract the preamble information
            preamble = []
            with open("%s/%s" %(inputPath, file), "r", encoding="latin-1") as fileCheck:
                csv_reader = reader(fileCheck, delimiter="\t")
                for row in csv_reader:
                       
                    if row[0].isalpha() == True or isfloat(row[0])==False:
                        preamble.append(row)
                           
                    elif row[0].isalpha() == False or isfloat(row[0])==True:
                        break

            # If no multiplier extract time and signal
            if len(preamble) <= 1:
                run = pd.read_csv("%s/%s" % (inputPath,file), sep="\t", encoding="latin-1", keep_default_na=True, na_values=str(0))
                run = run.dropna(how="all")
                run = run.fillna(0)
                run = run.iloc[:,[0,1]]
                run.columns = ["raw time", "Experiment " + str(runNumber)]


                # Create or add experiment to dataframe if it exists
                if conc in d:
                    d[conc].insert(len(d[conc].columns), "Experiment " + str(runNumber), run["Experiment " + str(runNumber)])

                elif conc not in d:
                    d[conc]=run

            # If multiplier needed extract signals, signal multipler, and iterate over signals
            elif len(preamble) > 1:
                signalMult_line = list(filter(lambda x: "Y Axis Multiplier:" in x[0], preamble))
                signalMult = float(signalMult_line[0][1])

                run = pd.read_csv("%s/%s" %(inputPath, file), sep="\t", encoding="latin-1", skiprows=len(preamble), header=None, keep_default_na=True, na_values=str(0))
                run = run.dropna(how="all")
                run = run.fillna(0)
                run.columns = ["Experiment" + str(runNumber)]
                run.loc[:"Experiment" + str(runNumber)] *= signalMult

                # Create or add experiment to dataframe if it exists
                if conc in d:
                    d[conc].insert(len(d[conc].columns), "Experiment " + str(runNumber), run)

                elif conc not in d:
                    # Reconstruct raw time from sampling rate
                    hz_line = list(filter(lambda x: "Sampling Rate:" in x[0], preamble))
                    hz = float(hz_line[0][1])
                    second_Gap = 1/hz

                    rawTime = [0]

                    for x in range(0,len(run)-1):
                        rawTime.append(rawTime[x]+second_Gap)

                    run.insert(0, "raw time", rawTime)
                    d[conc]=run
                        

    for xConc in d:
        orderedExp = natsorted(list(d[xConc].columns))
        orderedExp.remove("raw time")
        orderedExp.insert(0,"raw time")
        d[xConc] = d[xConc].reindex(columns = orderedExp)
            
    orderedDict = natsorted(d.keys())

    if prefix == "u":
        prefix = "µ"
            
    # Create Excel Working File
    wb = Workbook()
    ws = wb.active

    # Generate Inputs Sheet
    ws.title = "Inputs"
    ws["A1"] = "Propogation flow rate (%sM/min)" % prefix
    ws["B1"] = propFlow
    ws["A2"] = "Injection flow rate (%sM/min)" % prefix
    ws["B2"] = injectFlow
    ws["A3"] = "Injection time"
    ws["B3"] = injectTime
    ws["A4"] = "Separation capillary length (cm)"
    ws["B4"] = sepLength
    ws["A5"] = "Injection loop length (cm)"
    ws["B5"] = injectLength
    ws["A6"] = "Protein name"
    ws["B6"] = proteinName
    ws["A7"] = "Ligand name"
    ws["B7"] = ligandName
    ws["A8"] = "Number of Concentrations"
    ws["B8"] = len(d.keys())
    ws["A9"] = "Protein Concentration Unit"
    ws["B9"] = "%sM" % prefix
    ws["A10"] = "Initial Ligand concentration [L]0 (%sM)" % prefix
    ws["B10"] = ligandConc
    ws["A11"] = "Compensation procedure (Y) or (N) - recommended for MS data"
    ws["B11"] = compYN
    ws["A12"] = "[P]0 reference for MS normalization (%sM)" % prefix
    ws["B12"] = normalConc
    ws["A13"] = "Window width (%)"
    ws["B13"] = windowWidth
    ws["A14"] = "Determination of peak (M) or (P)"
    ws["B14"] = peakDet
    ws["A15"] = "Manually determined peaks"
    ws["B15"] = manualPeaks 
    ws["A16"] = "[P]0 used to programmaticlly determine peak"
    ws["B16"] = peakConc  


    for x in range(1,len(d)+1):
        ws["D"+str(x)] = "Protein Conc. #" + str(x)
        ws["F"+str(x)] = "Runs"

        ws["E"+str(x)] = orderedDict[x-1]
        ws["G"+str(x)] = len(d[orderedDict[x-1]].columns)-1

    # Write each concentration to a new sheet with all runs and raw time
    writer = pd.ExcelWriter("%s/%s" % (inputPath, workingFileName), engine = 'openpyxl')          ## User could name Excel file in GUI
    writer.book = wb

    for y in orderedDict:
        d[y].to_excel(writer, sheet_name = str(float(y)) + " %sM" % prefix, index=False)
            
    writer.save()
    writer.close()
    wb.save("%s/%s" % (inputPath,workingFileName))


    # Testing script execution time
    end = time.time()
    print("Script run time: %.2f seconds" %(end-start))

    return "%s/%s" % (inputPath, workingFileName)



# Analyze data in working file, add outputs and generate graph subfolder of pngs
def dataanalysis(fileName):
        
        # Testing script execution time
        start = time.time()

        ## Part 3 - Locating the raw data file and establishing important inputs                         
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
            percentage = int(idealSheet.cell(13,2).value)/100
            numberOfConcs = int(idealSheet.cell(8,2).value)
            injectionTime = float(idealSheet.cell(3,2).value)
            ligandConc = float(idealSheet.cell(10,2).value)
            
            proteinName = str(idealSheet.cell(6,2).value)
            compYN = str(idealSheet.cell(11,2).value)
            peakDet = str(idealSheet.cell(14,2).value)
            concUnit = str(idealSheet.cell(9,2).value)

            
            subdirect = "%s/%s_graphs_%s" %(os.path.dirname(fileName),proteinName, date.today())
            if os.path.exists(subdirect)== False:
                    os.mkdir(subdirect)
                    
            elif os.path.exists(subdirect)== True:
                    for duplicate in range(2,9):
                            temp = "%s_%d" % (subdirect, duplicate)
                            
                            if os.path.exists(temp)== False:
                                    subdirect=temp
                                    os.mkdir(subdirect)
                                    break
            
        elif "Inputs" not in inputBook.sheetnames:  
                sys.exit("Input file Formatting Error: The script expects inputdata.xlsx to be in a certain format, see provided idealinputs.xlsx as an example.") 

        # Temporary variables
        concentration = []
        signal = []
        stddev = []
        relstddev = []
        Rvalue = []
        Rstddev = []
        graphs = []
        graphNames = []
        forDF = [concentration,signal,stddev,relstddev,Rvalue,Rstddev]
        DFnames = ["Conc (%s)" % concUnit,"Avg Signal","Std Dev","Rel Std Dev","R value","R Std Dev"]

        # If programmatic determination of peak use first run at specified concentration to calculate peak time and time window
        if peakDet == "P":
                    windowCalcConc = float(idealSheet.cell(16,2).value)

                    for x in range(1,int(numberOfConcs)+1):
                            conc1 = float(idealSheet.cell(x,5).value)
                            if  conc1 == windowCalcConc:
                                    data = pd.read_excel(fileName, sheet_name=x, engine='openpyxl')
                                    data = data.dropna(how='all')
                                    xvalues = data['raw time']
                                    yvalues = data.iloc[:,1]

                                    peakSignal = max(yvalues[xvalues>=injectionTime])
                                    peakIndex = yvalues[yvalues == peakSignal].index
                                    peakTime = xvalues[peakIndex]


                
        ## Part 4 - Calculating signal information for each concentration and generating separagram graphs
        for x in range(1,int(numberOfConcs)+1):         
            conc1 = float(idealSheet.cell(x,5).value)
            numberOfRuns = idealSheet.cell(x,7).value
            
            avgSigConc = 0
            avgSigsRun = []


            # Reading in the whole data frame (= all runs for one particular concentration) and dropping all lines that are blank, i.e. that would produce "NaN"s
            data = pd.read_excel(fileName, sheet_name=x, engine='openpyxl')
            data = data.dropna(how='all')

            # Calculate background signals for each run
            for col in range(1, numberOfRuns):

                runName = data.columns[col]
                

                xvalues = data['raw time']
                yvalues = data[runName]
                
                # All y-values before the injection time are considered background signal 
                background_yvalues = yvalues[xvalues < injectionTime]
                background_average = np.average(background_yvalues)
                background_stdev = np.std(background_yvalues)


                # If manual determination, extract the manually set peak time for given concentration
                if peakDet == "M":
                    if x==1:
                            peakSignal = max(yvalues[xvalues>=injectionTime])
                               
                    manualTimes = idealSheet.cell(15,2).value.split(",")
                    peakIndex = xvalues.searchsorted(float(manualTimes[x-1]), side='left')
                    peakTime = xvalues[peakIndex-1]

                # Set time window parameters using peak time and determine the average signal within window for each run             
                windowLow = float(peakTime - (percentage * peakTime))
                windowHigh = float(peakTime + (percentage * peakTime))
                windowIndex = xvalues.between(windowLow,windowHigh)
                windowTimes = xvalues[windowIndex]
                windowSignals = yvalues[windowIndex]
                avgSigsRun.append(np.average(windowSignals))
          

                # Graph the signal for each run and with time window indicated
                plt.plot(xvalues, yvalues)

            # Appending a figure with all experimental runs for concentration
            plt.xlabel('Propagation time (s)', fontweight='bold')
            if compYN == "Y":
                    plt.ylabel('MS intensity (a.u.)', fontweight='bold')
            elif compYN == "N":
                    plt.ylabel('Fluorescence (a.u.)', fontweight='bold')
            plt.text(0,peakSignal*1.05, r"[%s]$\mathbf{_0}$ = %s %s" % (proteinName, conc1, concUnit), fontweight='bold')
            plt.vlines(windowLow, 0, peakSignal*1.05, linestyles='dashed',color='gray')
            plt.vlines(windowHigh, 0, peakSignal*1.05, linestyles='dashed',color='gray')
            plt.savefig("%s/%s%s.png" % (subdirect, conc1, concUnit))    # save separagram graphs
            graphs.append(plt.figure())
            plt.clf()


            # Calculating average signal for each concentration, stdev and relative stdev        
            avgSigConc = np.average(avgSigsRun)
            avgSigConc_stdev = np.std(avgSigsRun)
            avgSigConc_relstdev = (avgSigConc_stdev/avgSigConc)*100

                    
            concentration.append(conc1)
            signal.append(avgSigConc)    
            stddev.append(avgSigConc_stdev)
            relstddev.append(avgSigConc_relstdev)

            
        # Graphing separagrams for the first run for every concentration
        for x in range(1,int(numberOfConcs),2):         
            conc1 = idealSheet.cell(x,5).value

            # Reading in the whole data frame (= all runs for one particular concentration) and dropping all lines that are blank, i.e. that would produce "NaN"s
            data = pd.read_excel(fileName, sheet_name=x, engine='openpyxl')
            data = data.dropna(how='all')
            xvalues = data['raw time']
            yvalues = data.iloc[:, 2]

            p = plt.plot(xvalues, yvalues, label= '%s %s' % (conc1,concUnit))   
            #plt.text(xvalues[int(0.7*len(xvalues))], max(yvalues), '%.1f %s' % (conc1,concUnit), color=p[0].get_color())
            

        plt.xlabel('Propagation time (s)', fontweight='bold')
        if compYN == "Y":
                plt.ylabel('MS intensity (a.u.)', fontweight='bold')
        elif compYN == "N":
                plt.ylabel('Fluorescence (a.u.)', fontweight='bold')
        plt.text(0,peakSignal*1.05, r"[%s]$\mathbf{_0}$"  % (proteinName), fontweight='bold')
        plt.legend()
        if peakDet == "P":
                plt.vlines(windowLow, 0, peakSignal*1.05, linestyles='dashed',color='gray')
                plt.vlines(windowHigh, 0, peakSignal*1.05, linestyles='dashed',color='gray')
        elif peakDet == "M":
                plt.vlines(0, peakSignal, peakSignal*1.05, linestyles='dashed',color='white')

        # Appending a figure with first experimental run for all concentrations
        plt.savefig("%s/allconcentration.png" % subdirect)           # save separagram graph
        graphs.append(plt.figure())
        plt.clf()


        ## Part 5 - Calculate R values and standard deviation of R values for each concentration
        LowProt_sig = signal[0]
        HighProt_sig = signal[numberOfConcs-1]
        LowProt_stddev = stddev[0]
        HighProt_stdDev = stddev[numberOfConcs-1]

        for y in range(0, numberOfConcs):
                conc2 = idealSheet.cell(y+1,5).value
                avgSigConc_R = (signal[y] - HighProt_sig)/ (LowProt_sig - HighProt_sig) 
                Rvalue.append(avgSigConc_R)

                avgSiglConc_Rstddev = ( 1/(LowProt_sig - HighProt_sig) * math.sqrt( (stddev[y]**2) + ((signal[y] - LowProt_sig)/ (LowProt_sig - HighProt_sig) * HighProt_stdDev)**2 +
                                                                               ((HighProt_sig - signal[y])/(LowProt_sig - HighProt_sig) * LowProt_stddev)**2))
                Rstddev.append(avgSiglConc_Rstddev)


            
        ## Part 6 - Plotting the binding isotherm R vs P[0] with curve of best fit
        # Plotting data points for each concentration
        plt.scatter(concentration, Rvalue, c='white', edgecolor='black', label="R", zorder=10)
        plt.errorbar(concentration, Rvalue, yerr = Rstddev, linestyle="none", ecolor = 'black', elinewidth=1, capsize=2, capthick=1, zorder=0)
        plt.xscale("log")
                      
        # Define the Levenberg Marquardt algorithm
        def LevenMarqu(x,a):          
            return -((a + x - ligandConc)/(2*ligandConc)) + ((((a + x - ligandConc)/(2*ligandConc))**2) + (a/ligandConc))**(0.5)

        # Curve fitting and plotting curve of best fit
        popt, pcov = curve_fit(LevenMarqu, concentration, Rvalue)
        error = np.sqrt(np.diag(pcov))

        step=0
        if concentration[step] == float(0):
                step = 1

        xFit = np.arange(0.0, max(concentration), concentration[step])
        plt.plot(xFit, LevenMarqu(xFit, popt), linewidth=1.5, color='black', label="Best Fit")
        plt.text(0.2, 0.2, r'K$\mathbf{_d}$ = %.2f ± %.2f %s' % (popt, error, concUnit), fontweight='bold')
        plt.ylabel('R', fontweight='bold')
        plt.xlabel(r'[%s]$\mathbf{_0}$ (%s)' % (proteinName,concUnit), fontweight='bold')
        plt.xscale("log")
        plt.legend()
        plt.savefig("%s/bindingisotherm.png" % subdirect)           # save binding isotherm graph
        graphs.append(plt.figure())
        plt.close()

        # Statistics
        residuals = Rvalue - LevenMarqu(concentration, *popt)
        ss_res = np.sum(residuals**2)
        ss_tot = np.sum((Rvalue - np.mean(Rvalue))**2)
        r_squared = 1 - (ss_res/ss_tot)

        chiSquared = sum((((Rvalue - LevenMarqu(concentration, *popt))**2) / LevenMarqu(concentration, *popt)))

        # Returned/printed values
        print("Kd: %.4f ± %.4f %s" % (popt,error,concUnit))
        print("R²: %.4f" % (r_squared))
        print("χ²: %.4f" % (chiSquared))


        ## Part 7 - Returning summary data and graphs
        # Summary dataframe of average signal per concentration with standard deviation, relative standard deviation, R value and standard deviation
        df = pd.DataFrame (forDF).transpose()
        df.columns = DFnames

        # Create new output sheet in input Excel file with summary data and Kd, R² and χ²
        writer = pd.ExcelWriter(fileName, engine = 'openpyxl')
        writer.book = inputBook
        df.to_excel(writer, sheet_name = "Outputs", index=False, startcol=8)        # does not overwrite if a sheet named Outputs already exists
        writer.save()
        writer.close()

        outputSheet = inputBook.worksheets[len(inputBook.sheetnames)-1]

        maxr = idealSheet.max_row
        maxc = idealSheet.max_column
        for r in range(1, maxr+1):
                for c in range (1, maxc+1):
                        outputSheet.cell(row=r, column=c).value = idealSheet.cell(row=r, column=c).value

        
        outputSheet["P1"] = "Kd: %.4f ± %.4f %s" % (popt,error,concUnit)
        outputSheet["P2"] = "R²: %.4f" % (r_squared)
        outputSheet["P3"] = "χ²: %.4f" % (chiSquared)

        inputBook.save(fileName)


        # Testing script execution time
        end = time.time()
        print("Script run time: %.2f seconds" %(end-start))

        # Returning all graphs (separagrams and binding isotherm)
        #plt.show()

        return subdirect
    



######### EVENT LOOP ##############
while True:
    event, values = window.read() 
    if event == sg.WIN_CLOSED:
        break

# Radiobuttons displaying relevant secondary fields
    # If compensation needed request concentration for normalization
    if event == 'compYes':
        window['normConc'].update(visible=False)
        window['normConc_val'].update(visible=False)
        
    if event == 'compNo':
        window['normConc'].update(visible=False)
        window['normConc_val'].update(visible=False)

    # If manual peak determination, enter peaks    
    if event == 'peakM':
        window['manPeaks'].update(visible=True)
        window['manPeaks_val'].update(visible=True)
        window['progPeak'].update(visible=False)
        window['progPeak_val'].update(visible=False)
        
    # If programmatic peak determination, enter concentration to use
    if event == 'peakP':
        window['manPeaks'].update(visible=False)
        window['manPeaks_val'].update(visible=False)
        window['progPeak'].update(visible=True)
        window['progPeak_val'].update(visible=True)
        
    if event == 'calculate':
        
        window['out'].update(visible=False)

        filePath = str(values['filePath_val'])

        if os.path.exists(filePath):
            if os.path.isfile(filePath) and filePath.endswith(".xlsx"):
                workingFile = filePath
                
            elif os.path.isdir(filePath):
                propFlow = float(values['propFlow_val'])
                injectFlow = float(values['injectFlow_val'])
                injectTime = float(values['injectTime_val'])
                sepLength = float(values['sepLength_val'])
                injectLength = float(values['injectLength_val'])
                
                proteinName = str(values['protName_val'])
                ligandName = str(values['ligName_val'])
                protConcs = str(values['protConcs_val'])
                ligandConc = float(values['ligConc_val'])


                if window['compYes'].get()== True:
                    compYN = "Y"
                    normalConc = float(values['normConc_val'])
                elif window['compNo'].get()== True:
                    compYN = "N"
                    normalConc = None

                windowWidth = float(values['window_val'])

                if values['peakM'] == True:
                    peakDet = "M"
                    manualPeaks = str(values['manPeaks_val'])
                    peakConc=None
                elif values['peakP'] == True:
                    peakDet = "P"
                    manualPeaks = None
                    peakConc = float(values['progPeak_val'])

                workingFile = workingfileprep(filePath, propFlow, injectFlow, injectTime, sepLength, injectLength,
                                              proteinName, ligandName, protConcs, ligandConc, compYN, normalConc,
                                              windowWidth, peakDet, manualPeaks, peakConc)

        graphPath = dataanalysis(workingFile)
        images = natsorted(glob.glob('%s/*.png' % graphPath))
        load_image(images[0],window)
        
        df = pd.read_excel(workingFile, sheet_name=-1, header=None, usecols="I:N", engine='openpyxl')
        df = df.dropna(how='any')
        headers = df.iloc[0].values.tolist()
        data = df.iloc[1:].values.tolist()
        window['summary'].update(values=data, num_rows=min(10,len(data)))

        df = pd.read_excel(workingFile, sheet_name=-1, header=None, usecols="P", engine='openpyxl')
        df = df.dropna(how='any')
        data = df.values.tolist()
        window['Kd'].update(values=data, num_rows=3)
        
        window['out'].update(visible=True)
        window['report'].update(visible=True)

        
    if event == 'fwd':
        if location == len(images)-1:
            location=0
        else:
            location +=1
        load_image(images[location], window)
    

    if event == 'back':
        if location == 0:
            location=len(images)-1
        else:
             location -=1
        load_image(images[location], window)

    
window.close()

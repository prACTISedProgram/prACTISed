#20220725 JL    NOTE: Switched to PySimpleGUI, added data analysis and working file preparation as functions (compensation & report WIP)
#20220809 JL    NOTE: Added compensation with interpolation from simulated, small adjustments to working file and graph formatting (TODO: report)
#20220912 JL    NOTE: Added pdf report option to GUI (TODO: unicode for Chi character, open pdf)

import PySimpleGUI as sg

import os
import pathlib
import glob
import sys

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
from scipy import interpolate
from datetime import date
import time
from natsort import natsorted
from fpdf import FPDF
#from working import workingfileprep
#from analysis import dataanalysis
#from compensation import compensate
#from pdf import report

sg.theme('DarkBlue3')
    
inputCol = [
    [sg.Frame('Fluidics Experimental Parameters', 
              [[sg.Text('Propagation flow rate', size=(25,1), key ='propFlow'),
                sg.Input(size=(10,1), key='propFlow_val')],
               [sg.Text('Injection flow rate', size=(25,1),key='injectFlow'),
                sg.Input(size=(10,1), key='injectFlow_val')],
               [sg.Text('Injection time (s)', size=(25,1), key='injectTime'),
                sg.Input(size=(10,1), key='injectTime_val')],
               [sg.Text('Separation capillary length', size=(25,1), key='sepLength'),
                sg.Input(size=(10,1), key='sepLength_val')],
               [sg.Text('Separation capillary diameter', size=(25,1), key='sepDiam'),
                sg.Input(size=(10,1), key='sepDiam_val')],
               [sg.Text('Injection loop length', size=(25,1), key='injectLength'),
                sg.Input(size=(10,1), key='injectLength_val')],
               [sg.Text('Injection loop diameter', size=(25,1), key='injectDiam'),
                sg.Input(size=(10,1), key='injectDiam_val')]],
              key = 'fluidParam')],


    [sg.Frame('Concentrations', 
              [[sg.Text('Protein name', size=(25,1), key ='protName'),
                sg.Input(size=(10,1), key='protName_val')],
               [sg.Text('Ligand name', size=(25,1), key='ligName'),
                sg.Input(size=(10,1), key='ligName_val')],
               [sg.Text('Initial ligand concentration [L]0 \n (same units as [P]0)', key='ligConc'),
                sg.Input(size=(10,1), key='ligConc_val')]],
              key='concFrame')],
    
    [sg.Frame('Data Analyis Parameters',
              [[sg.Text('Type of data', size=(11,1), key ='dataType'),
                sg.Radio('Fluoresence', 'dataChoice', key ='dataF', enable_events=True),
                sg.Radio('Mass Spec', 'dataChoice', key ='dataMS', enable_events=True)],
               [sg.Text('Compensation procedure \n (recommended for MS data)', key ='comp'),
                sg.Radio('Yes', 'compChoice', key ='compYes', enable_events=True),
                sg.Radio('No', 'compChoice', key ='compNo', enable_events=True)],
               [sg.Text('[P]0 reference for normalization', size=(30,1), key='normConc', visible=False),
                sg.Input(size=(10,1), key='normConc_val', visible=False)],
               [sg.Text('Window width (%)', size=(30,1), key='window'),
                sg.Input(size=(10,1), key='window_val')],
               [sg.Text('Determination of peak', size=(17,1), key ='peak'),
                sg.Radio('Manual', 'peakChoice', key ='peakM', enable_events=True),
                sg.Radio('Program', 'peakChoice', key ='peakP', enable_events=True)],
               [sg.Text('Peaks for concentrations [P]0 \n (ascending [P]0, separated by commas)', key='manPeaks', visible=False)],
               [sg.Input(size=(42,1), key='manPeaks_val', visible=False)],
               [sg.Text('Specify [P]0 used to determine peak', size=(30,1), key='progPeak', visible=False),
                sg.Input(size=(10,1), key='progPeak_val', visible=False)]],
              key='analysisParam')],
    
    [sg.Button('Calculate Kd', key='calculate', disabled=True, disabled_button_color='gray'),
     sg.Button('Report', key='report', visible=False)]
    ]

        
outputCol= [
    [sg.Table(values=[],headings=['prACTISed output'], key='Kd', hide_vertical_scroll=True, def_col_width=20, auto_size_columns=False)],

    [sg.Table(values=[], headings=['Conc','Avg Signal', 'Std Dev', 'Rel Std Dev', 'R value', 'R Std Dev'], key='summary', def_col_width=10,auto_size_columns=False)],

    [sg.Frame('Graphs',
              [[sg.Image(key='graphImage')],
               [sg.Button('Prev', key='back'),
                sg.Button('Next', key='fwd')]]
               )]           
    ]
    

layout = [
    [sg.Text('File path', key ='filePath'),
     sg.Input(size=(50,1), key='filePath_val'),
     sg.Button('Validate', key='validate')],
    [sg.Text('Working File Entered', key='work', visible=False),
     sg.Text('Raw Data Directory Entered', key='rawData', visible=False),
     sg.Text('Invalid File Path Entered', key='invalid', visible=False)],
    
    [sg.Column(inputCol), sg.Column(outputCol, visible=False, key='out')]
    ]

window = sg.Window("prACTISed", layout, finalize=True)


######### DEFINE FUNCTIONS ##############

# Display image from file path
def load_image(path,window):
    img = Image.open(path)
    img.thumbnail((350, 350))
    photo_img = ImageTk.PhotoImage(img)
    window['graphImage'].update(data=photo_img)

location = 0


### Generate working file from directory containing raw data (.txt or .asc)

def workingfileprep(inputPath, propFlow, injectFlow, injectTime, sepLength, sepDiam,
                    injectLength, injectDiam, proteinName, ligandName, ligandConc,
                    dataType, compYN, normalConc, windowWidth, peakDet, manualPeaks, peakConc):

    
    ## Part 1 - Importing the required libraries and sub-libraries required below
    import argparse                                   
    import pathlib
    import sys
    import pandas as pd
    import matplotlib.pyplot as plt
    import numpy as np
    from openpyxl import load_workbook
    from openpyxl import Workbook
    from openpyxl.utils import FORMULAE
    from openpyxl.utils import get_column_letter, column_index_from_string
    import math
    from scipy.optimize import curve_fit

    import os
    from csv import reader
    from natsort import natsorted
    from datetime import date

    # Testing script execution time
    import time
    start = time.time()


    d = {}
    prefix=""

    def isfloat(num):
        try:
            float(num)
            return True
        except ValueError:
            return False     

    ### Reading in files from directory ###
    for file in os.listdir(inputPath):

        ## Read in simulated protein profile if compensation procedure was selected
        if compYN == "Y" and file.endswith((".txt")) and file.startswith('simulated'):

             # Extract the preamble information
             preamble = []
             with open("%s/%s" %(inputPath, file), "r", encoding="latin-1") as fileCheck:
                 csv_reader = reader(fileCheck, delimiter=" ")
         
                 for row in csv_reader:
                     if row[0].isalpha() == True or isfloat(row[0])==False:
                         preamble.append(row)
                           
                     elif row[0].isalpha() == False or isfloat(row[0])==True:
                         break

             # Read in simulated protein profile if compensation procedure was selected
             simulated = pd.read_csv("%s/%s" % (inputPath,file), sep="\s+", encoding="latin-1", skiprows=len(preamble), header=None, keep_default_na=True, na_values=str(0))
             simulated = simulated.fillna(0)
             simulated.columns = ['raw time', 'signal']

             # Normalize simulated protein signals
             maxSim = simulated['signal'].max()
             simulated['signal'] = simulated['signal'].div(maxSim)


        ## Read in raw data files 
        if file.endswith((".txt", ".asc")) and not file.startswith('simulated'):

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
                run.iloc[:,0] = run.iloc[:,0].mul(60)


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
            
    workingFileName = "%s_%s.xlsx" % (proteinName, date.today())

    if os.path.exists("%s/%s" % (inputPath, workingFileName)) == True:
        wb = load_workbook("%s/%s" % (inputPath, workingFileName))
        writer = pd.ExcelWriter("%s/%s" % (inputPath, workingFileName), engine='openpyxl')
        writer.book = wb

        writer.sheets = dict((ws.title, ws) for ws in wb.worksheets)
        ws = wb["Inputs"]

    elif os.path.exists("%s/%s" % (inputPath, workingFileName)) == False:
        wb = Workbook()
        ws = wb.active
        ws.title = "Inputs"

        writer = pd.ExcelWriter("%s/%s" % (inputPath, workingFileName), engine = 'openpyxl')
        writer.book = wb
        
        
    # Generate Inputs Sheet
    ws["A1"] = "Propogation flow rate"
    ws["B1"] = propFlow
    ws["A2"] = "Injection flow rate"
    ws["B2"] = injectFlow
    ws["A3"] = "Injection time (s)"
    ws["B3"] = injectTime
    ws["A4"] = "Separation capillary length"
    ws["B4"] = sepLength
    ws["A5"] = "Separation capillary diameter"
    ws["B5"] = sepDiam
    ws["A6"] = "Injection loop length"
    ws["B6"] = injectLength
    ws["A7"] = "Injection loop diameter"
    ws["B7"] = injectDiam
    ws["A8"] = "Protein name"
    ws["B8"] = proteinName
    ws["A9"] = "Ligand name"
    ws["B9"] = ligandName
    ws["A10"] = "Number of Concentrations"
    ws["B10"] = len(d.keys())
    ws["A11"] = "Initial Ligand concentration [L]0 (%sM)" % prefix
    ws["B11"] = ligandConc
    ws["A12"] = "Type of Data"
    ws["B12"] = dataType
    ws["A13"] = "Compensation procedure"
    ws["B13"] = compYN
    ws["A14"] = "[P]0 reference for MS normalization (%sM)" % prefix
    ws["B14"] = normalConc
    ws["A15"] = "Window width (%)"
    ws["B15"] = windowWidth
    ws["A16"] = "Determination of peak"
    ws["B16"] = peakDet
    ws["A17"] = "Manually determined peaks"
    ws["B17"] = manualPeaks 
    ws["A18"] = "[P]0 to programmaticlly determine peak"
    ws["B18"] = peakConc


    for x in range(1,len(d)+1):
        ws["D"+str(x)] = "Protein Conc. #" + str(x)
        ws["E"+str(x)] = "%s %sM" % (orderedDict[x-1], prefix)


    if compYN == "Y":
        simulated.to_excel(writer, sheet_name = "P_simulated", index=False)

    for y in orderedDict:
        d[y].to_excel(writer, sheet_name = "%s %sM" % (y, prefix), index=False)
            
    writer.save()
    writer.close()
    wb.save("%s/%s" % (inputPath,workingFileName))


    # Testing script execution time
    end = time.time()
    print("Script run time: %.2f seconds" %(end-start))

    return "%s/%s" % (inputPath, workingFileName)

### Compensation procedure to unmask signal

def compensate (fileName):
#fileName = "/Users/jess/Documents/practised/ALP.xlsx"

        d ={}

        data = pd.read_excel(fileName, engine='openpyxl')
        inputBook = load_workbook(fileName, data_only=True)
        idealSheet = inputBook["Inputs"]

        injectTime = idealSheet.cell(3,2).value
        numberOfConcs = idealSheet.cell(10,2).value
        normalConc = float(idealSheet.cell(14,2).value)
        unit = idealSheet.cell(1,5).value.partition(" ")[2]
        normalConc = "%s %s" % (normalConc, unit)
        compYN = str(idealSheet.cell(13,2).value)

        if compYN == 'Compensated':
                sys.exit("Compensation procedure has alredy been applied to this working file")
        

        # Get dimensionless simulated separagram of pure protein, S̃p, and interpolate signals
        simulated = pd.read_excel(fileName, sheet_name='P_simulated', engine='openpyxl')

        # Interpolate signals from simulated protein profile
        timeSim = simulated['raw time']
        sigSim = simulated['signal']
        interp = interpolate.splrep(timeSim, sigSim)

        # Isolate interrpolated signals at for times in raw data files
        rawData = pd.read_excel(fileName, sheet_name=normalConc, engine='openpyxl')
        rawData = rawData.dropna(how='all')
        rawTime = rawData['raw time']
        interpSigs = interpolate.splev(rawTime, interp, der=0)

        # Normalize signals, remove negative values and generate data frame
        interpSigs = pd.Series(interpSigs, name='signal')
        interpSigs = interpSigs.div(max(interpSigs))
        interpSigs[interpSigs < 0] = 0
        simulatedSigs = pd.concat([rawTime,interpSigs], axis=1)

        # Get integrated signal of normalization concentration
        integratedNorm = []

        # Isolate first run of concentration used to normalize
        rawSignal = pd.read_excel(fileName, sheet_name=normalConc, engine='openpyxl')
        rawSignal = rawSignal.dropna(how='all')
        rawSignal = rawSignal.iloc[:,:2]

        time = rawSignal['raw time']
        sig = rawSignal[rawSignal.columns[1]]
        colName = str(rawSignal.columns[1])

        # Average the background signals for first 5 secs, subtract from all signals
        background = sig[time < injectTime].mean()
        rawSignal.iloc[:,1] = rawSignal.iloc[:,1].sub(background)

        # Multiply raw signal by Sp
        rawSignal.iloc[:,1] = rawSignal.iloc[:,1].mul(simulatedSigs.iloc[:,1])

        # Normalize and remove negative signals
        rawSignal.iloc[:,1] = rawSignal.iloc[:,1].mul(1)
        rawSignal.loc[(rawSignal[colName] < 0), colName] = 0
        normArea = rawSignal.iloc[:,1].sum()

        # Add concentration to dictionary
        d[normalConc]= rawSignal
                

        # Apply Sp and normalization to all other concentrations and runs
        for x in range(1,int(numberOfConcs)+1):
            
                conc1 = str(idealSheet.cell(x,5).value)

                # Read in all data for concentration
                rawSignal = pd.read_excel(fileName, sheet_name=conc1, engine='openpyxl')
                rawSignal = rawSignal.dropna(how='all')

                time = rawSignal['raw time']
                numberOfRuns = len(rawSignal.columns)

                for run in range(1, int(numberOfRuns)):

                        # Skip normalization concentration run 1
                        if conc1.partition(" ")[0] != "%s" % normalConc or run != 1:
                        
                                sig = rawSignal[rawSignal.columns[run]]

                                # Average the background signals for frist 5 secs, subtract from all signals
                                background = sig[time < injectTime].mean()
                                sig = sig.sub(background)

                                # Multiply raw signal by Sp
                                sig = sig.mul(simulatedSigs.iloc[:,1])

                                # Normalize
                                sigArea = sig.sum()
                                sig = sig.mul(normArea)
                                sig = sig.div(sigArea)
                                sig[sig < 0] = 0

                                # Add concentration to dictionary or append run to exisiting entry
                                if run == 1:
                                        sig = pd.DataFrame(sig, columns=[rawSignal.columns[run]])
                                        sig.insert(0, "raw time", time)
                                        d[conc1]= sig

                                else:
                                        d[conc1].insert(len(d[conc1].columns), rawSignal.columns[run], sig)
        

        writer = pd.ExcelWriter(fileName, engine='openpyxl')
        writer.book = inputBook

        writer.sheets = dict((ws.title, ws) for ws in inputBook.worksheets)

        for y in d.keys():
                d[y].to_excel(writer, sheet_name = y, index=False)

        simulatedSigs.to_excel(writer, sheet_name = 'P_simulated', index=False)

        idealSheet["B13"] = "Compensated"

        writer.save()
        writer.close()
        inputBook.save(fileName)
        


### Analyze data in working file, add outputs and generate graph subfolder of pngs

def dataanalysis(fileName):
        
        # Testing script execution time
        import time
        start = time.time()

        ## Part 1 - Importing the required libraries and sub-libraries required below
        import argparse                                   
        import pathlib
        import sys
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
        import os


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
            percentage = int(idealSheet.cell(15,2).value)/100
            numberOfConcs = int(idealSheet.cell(10,2).value)
            injectionTime = float(idealSheet.cell(3,2).value)
            ligandConc = float(idealSheet.cell(11,2).value)
            
            proteinName = str(idealSheet.cell(8,2).value)
            dataType = str(idealSheet.cell(12,2).value)
            peakDet = str(idealSheet.cell(16,2).value)

            
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
        DFnames = ["Conc","Avg Signal","Std Dev","Rel Std Dev","R value","R Std Dev"]

        # If programmatic determination of peak use first run at specified concentration to calculate peak time and time window
        if peakDet == "P":
                windowCalcConc = float(idealSheet.cell(18,2).value)
                windowCalcConc = "%s %s" % (windowCalcConc, idealSheet.cell(1,5).value.partition(" ")[2])

                data = pd.read_excel(fileName, sheet_name=windowCalcConc, engine='openpyxl')
                data = data.dropna(how='all')
                xvalues = data['raw time']
                yvalues = data.iloc[:,1]

                peakSignal = max(yvalues[xvalues>=injectionTime])
                peakIndex = yvalues[yvalues == peakSignal].index
                peakTime = xvalues[peakIndex]

        # Determine maximum signal value (used to set graph parameters)
        maxSig = 0 
        for x in range(1,int(numberOfConcs)+1):         
                conc1 = idealSheet.cell(x,5).value
                data = pd.read_excel(fileName, sheet_name=conc1, engine='openpyxl')
                data = data.dropna(how='all')
                data = data.iloc[:,1:]
                colMax = data.max()
                currentMax = colMax.max()

                if currentMax > maxSig:
                        maxSig = currentMax

                
        ## Part 4 - Calculating signal information for each concentration and generating separagram graphs
        for x in range(1,int(numberOfConcs)+1):         
            conc1 = idealSheet.cell(x,5).value
            
            avgSigConc = 0
            avgSigsRun = []


            # Reading in the whole data frame (= all runs for one particular concentration) and dropping all lines that are blank, i.e. that would produce "NaN"s
            data = pd.read_excel(fileName, sheet_name=conc1, engine='openpyxl')
            data = data.dropna(how='all')
            numberOfRuns = len(data.columns)
            

            # Calculate background signals for each run
            for col in range(1, numberOfRuns):

                runName = data.columns[col]

                xvalues = data['raw time']
                yvalues = data[runName]
                minTime = xvalues[0]
                
                # All y-values before the injection time are considered background signal 
                background_yvalues = yvalues[xvalues < injectionTime]
                background_average = np.average(background_yvalues)
                background_stdev = np.std(background_yvalues)


                # If manual determination, extract the manually set peak time for given concentration
                if peakDet == "M":
                    if x==1:
                            peakSignal = max(yvalues[xvalues>=injectionTime])
                               
                    manualTimes = idealSheet.cell(17,2).value.split(",")
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
            if dataType == "MS":
                    plt.ylabel('MS intensity (a.u.)', fontweight='bold')
            elif dataType == "F":
                    plt.ylabel('Fluorescence (a.u.)', fontweight='bold')
            plt.text(minTime,maxSig*1.05, r"[%s]$\mathbf{_0}$ = %s" % (proteinName, conc1), fontweight='bold')
            plt.vlines(windowLow, 0, maxSig*1.05, linestyles='dashed',color='gray')
            plt.vlines(windowHigh, 0, maxSig*1.05, linestyles='dashed',color='gray')
            plt.savefig("%s/%s.png" % (subdirect, conc1))    # save separagram graphs
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
            data = pd.read_excel(fileName, sheet_name=conc1, engine='openpyxl')
            data = data.dropna(how='all')
            xvalues = data['raw time']
            yvalues = data.iloc[:, 2]

            p = plt.plot(xvalues, yvalues, label= '%s' % conc1)   

        plt.xlabel('Propagation time (s)', fontweight='bold')
        if dataType == "MS":
                plt.ylabel('MS intensity (a.u.)', fontweight='bold')
        elif dataType == "F":
                plt.ylabel('Fluorescence (a.u.)', fontweight='bold')
        plt.text(minTime, maxSig*1.05, r"[%s]$\mathbf{_0}$"  % (proteinName), fontweight='bold')
        plt.legend()
        if peakDet == "P":
                plt.vlines(windowLow, 0, maxSig*1.05, linestyles='dashed',color='gray')
                plt.vlines(windowHigh, 0, maxSig*1.05, linestyles='dashed',color='gray')
        elif peakDet == "M":
                plt.vlines(0, maxSig, maxSig*1.05, linestyles='dashed',color='white')

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
        # Convert concentration strings to floats
        concs = []
        for element in range(0, len(concentration)):
                num = concentration[element].partition(" ")[0]
                concs.append(float(num))
        unit = concentration[0].partition(" ")[2]

        # Plotting data points for each concentration
        plt.scatter(concs, Rvalue, c='white', edgecolor='black', label="R", zorder=10)
        plt.errorbar(concs, Rvalue, yerr = Rstddev, linestyle="none", ecolor = 'black', elinewidth=1, capsize=2, capthick=1, zorder=0)
        plt.xscale("log")
                      
        # Define the Levenberg Marquardt algorithm
        def LevenMarqu(x,a):          
            return -((a + x - ligandConc)/(2*ligandConc)) + ((((a + x - ligandConc)/(2*ligandConc))**2) + (a/ligandConc))**(0.5)

        # Curve fitting and plotting curve of best fit    
        popt, pcov = curve_fit(LevenMarqu, concs, Rvalue)
        error = np.sqrt(np.diag(pcov))

        step=0
        if concs[step] == float(0):
                step = 1

        xFit = np.arange(0.0, max(concs), concs[step])
        plt.plot(xFit, LevenMarqu(xFit, popt), linewidth=1.5, color='black', label="Best Fit")
        plt.text((concs[step]), 0.2, r'K$\mathbf{_d}$ = %.2f ± %.2f %s' % (popt, error, unit), fontweight='bold')
        plt.ylabel('R', fontweight='bold')
        plt.xlabel(r'[%s]$\mathbf{_0}$ (%s)' % (proteinName,unit), fontweight='bold')
        plt.xscale("log")
        plt.legend()
        plt.savefig("%s/bindingisotherm.png" % subdirect)           # save binding isotherm graph
        graphs.append(plt.figure())
        plt.close()

        # Statistics
        residuals = Rvalue - LevenMarqu(concs, *popt)
        ss_res = np.sum(residuals**2)
        ss_tot = np.sum((Rvalue - np.mean(Rvalue))**2)
        r_squared = 1 - (ss_res/ss_tot)

        chiSquared = sum((((Rvalue - LevenMarqu(concs, *popt))**2) / LevenMarqu(concs, *popt)))


        ## Part 7 - Returning summary data and graphs
        # Summary dataframe of average signal per concentration with standard deviation, relative standard deviation, R value and standard deviation
        df = pd.DataFrame (forDF).transpose()
        df.columns = DFnames

        # Create new output sheet in input Excel file with summary data
        writer = pd.ExcelWriter(fileName, engine = 'openpyxl')
        writer.book = inputBook
        df.to_excel(writer, sheet_name = "Outputs", float_format='%.4f', index=False, startcol=3, engine = 'openpyxl')        # does not overwrite if a sheet named Outputs already exists
        writer.save()
        writer.close()

        outputSheet = inputBook.worksheets[len(inputBook.sheetnames)-1]

        # Duplicate input sheet information onto output sheet for reproducibility
        for r in range(1, 18):
                for c in range (1, 3):
                        outputSheet.cell(row=r, column=c).value = idealSheet.cell(row=r, column=c).value

        
        # Include Kd, R² and χ² values on output sheet
        outputSheet["K1"] = "Kd: %.4f ± %.4f %s" % (popt,error,unit)
        outputSheet["K2"] = "R²: %.4f" % (r_squared)
        outputSheet["K3"] = "χ²: %.4f" % (chiSquared)

        inputBook.save(fileName)

        # Testing script execution time
        end = time.time()
        print("Script run time: %.2f seconds" %(end-start))

        # Returning all graphs (separagrams and binding isotherm)
        #plt.show()
        plt.close('all')

        return subdirect

### Generate PDF report summary

def report (workingFile, graphFolder):
    # Read in data tables from working file
    userInputs = pd.read_excel(workingFile, sheet_name=-1, header=None, usecols="A:B", engine='openpyxl')
    proteinName = str(userInputs.iloc[7,1])
    userLength = userInputs.shape[0]
    userInputs = userInputs.values.tolist()

    summaryTable = pd.read_excel(workingFile, sheet_name=-1, header=None, usecols="D:I", engine='openpyxl')
    summaryTable = summaryTable.dropna(how='any')
    sumLength = summaryTable.shape[0]
    summaryTable = summaryTable.values.tolist()

    kdTable = pd.read_excel(workingFile, sheet_name=-1, header=None, usecols="K", engine='openpyxl')
    kdTable = kdTable.dropna(how='any')
    kdTable = kdTable.values.tolist()

    pdf = FPDF()
    pdf.add_page()
    pdf.set_font('Times', size=10)

    pdf.cell(pdf.epw/2, pdf.font_size*2, '%s_%s' % (proteinName,date.today()))

    current = pdf.get_y() +5
    pdf.set_y(current)
    pdf.image("%s/bindingisotherm.png" % (graphFolder), w = pdf.epw/2)
    pdf.set_y(current)
    pdf.image("%s/allconcentration.png" % (graphFolder), w = pdf.epw/2, x = pdf.epw/2)

    # Add summary table to pdf
    current = pdf.get_y() + 5
    pdf.set_y(current)
    for row in summaryTable:
        pdf.set_x(2.5*(pdf.epw/5))
        
        for col in row:
            pdf.multi_cell(pdf.epw/12 , pdf.font_size*2, str(col), border=1,
                    new_x="RIGHT", new_y="TOP", align="L", max_line_height=pdf.font_size)
        pdf.ln(pdf.font_size*2)

    # Add Kd and statistics to pdf
    pdf.set_y(current + pdf.font_size*2*sumLength + 5)
    for row in kdTable:
        pdf.set_x(2.5*(pdf.epw/5))
        
        for col in row:
            if col.find('χ') != -1:
                col = col.replace("χ", "Chi")
            pdf.multi_cell(pdf.epw/4 , pdf.font_size*2, str(col), border=1,
                    new_x="RIGHT", new_y="TOP", align="L", max_line_height=pdf.font_size)
        pdf.ln(pdf.font_size*2)

    # Add user input table to pdf
    pdf.set_y(current)

    for row in userInputs:
        for col in row:
            pdf.multi_cell(pdf.epw/5, pdf.font_size*2, str(col), border=1,
                    new_x="RIGHT", new_y="TOP", align="L", max_line_height=pdf.font_size)
        pdf.ln(pdf.font_size*2)

    # Add separagram images to pdf
    images = natsorted(glob.glob('%s/*.png' % graphFolder))
    pdf.add_page()
    pos = 0.2
    graphY = pdf.get_y()

    for img in images:
        if img.endswith('M.png'):
            if pos < 3:
                pdf.set_y(graphY)
                pdf.image(img, w = pdf.epw/3, x = pos*pdf.epw/3)
                pdf.ln(0)
                pos +=1
                
                
            elif pos == 3.2:
                pos = 0.2
                graphY = pdf.get_y()+5
                pdf.set_y(graphY)
                pdf.image(img, w = pdf.epw/3, x = pos*pdf.epw/3)
                pos = 1.2

    currentFolder = os.path.dirname(workingFile)
    pdfNAME = '%s/%s_%s' % (currentFolder, proteinName, date.today())
    pdf.output(pdfNAME)



######### EVENT LOOP ##############
while True:
    event, values = window.read() 
    if event == sg.WIN_CLOSED:
        break

# Radiobuttons displaying relevant secondary fields
    # If compensation needed request concentration for normalization
    if event == 'compYes':
        window['normConc'].update(visible=True)
        window['normConc_val'].update(visible=True)
        
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

    if event == 'validate':
        filePath = str(values['filePath_val'])
        window['out'].update(visible=False)
        window['report'].update(visible=False)

        if not os.path.exists(filePath):
            window['fluidParam'].update(visible=False)
            window['concFrame'].update(visible=False)
            window['analysisParam'].update(visible=False)
            window['rawData'].update(visible=False)
            window['work'].update(visible=False)
            window['invalid'].update(visible=True)
            window['calculate'].update(disabled=True)

        elif os.path.exists(filePath):
            if os.path.isfile(filePath) and filePath.endswith(".xlsx"):
                workingFile = filePath
                window['fluidParam'].update(visible=False)
                window['concFrame'].update(visible=False)
                window['analysisParam'].update(visible=False)
                window['rawData'].update(visible=False)
                window['invalid'].update(visible=False)
                window['work'].update(visible=True)
                window['calculate'].update(disabled=False)
                

            elif os.path.isdir(filePath):
                window['fluidParam'].update(visible=True)
                window['concFrame'].update(visible=True)
                window['analysisParam'].update(visible=True)
                window['rawData'].update(visible=True)
                window['invalid'].update(visible=False)
                window['work'].update(visible=False)
                window['calculate'].update(disabled=False)
   
    
    if event == 'calculate':
        
        window['out'].update(visible=False)

        filePath = str(values['filePath_val'])

        if os.path.exists(filePath):
            if os.path.isfile(filePath) and filePath.endswith(".xlsx"):
                workingFile = filePath
                data = pd.read_excel(workingFile, engine='openpyxl')
                inputBook = load_workbook(workingFile, data_only=True)         
                idealSheet = inputBook["Inputs"]
                compYN = str(idealSheet.cell(13,2).value)
                
            # If file path is directory read in user inputs
            elif os.path.isdir(filePath):
                
                # Read in user inputs for fluidics parameters
                propFlow = values['propFlow_val']
                injectFlow = values['injectFlow_val']
                injectTime = float(values['injectTime_val'])
                sepLength = values['sepLength_val']
                sepDiam = values['sepDiam_val']
                injectLength = values['injectLength_val']
                injectDiam = values['injectDiam_val']

                # Read in user inputs for concentrations
                proteinName = str(values['protName_val'])
                ligandName = values['ligName_val']
                ligandConc = float(values['ligConc_val'])

                # Read in user inputs for data analysis parameters
                if window['dataF'].get() == True:
                    dataType = "F"
                elif window['dataMS'].get() == True:
                    dataType = "MS"
                    
                if window['compYes'].get()== True:
                    compYN = "Y"
                    normalConc = float(values['normConc_val'])
                elif window['compNo'].get()== True:
                    compYN = "N"
                    normalConc = None

                windowWidth = values['window_val']

                if values['peakM'] == True:
                    peakDet = "M"
                    manualPeaks = values['manPeaks_val']
                    peakConc=None
                elif values['peakP'] == True:
                    peakDet = "P"
                    manualPeaks = None
                    peakConc = float(values['progPeak_val'])

                workingFile = workingfileprep(filePath, propFlow, injectFlow, injectTime, sepLength, sepDiam,
                                              injectLength, injectDiam, proteinName, ligandName, ligandConc, dataType,
                                              compYN, normalConc, windowWidth, peakDet, manualPeaks, peakConc)



            if compYN == "Y" :
                compensate(workingFile)
                

            graphPath = dataanalysis(workingFile)
            images = natsorted(glob.glob('%s/*.png' % graphPath))
            load_image(images[0],window)
            
            df = pd.read_excel(workingFile, sheet_name=-1, header=None, usecols="D:I", engine='openpyxl')
            df = df.dropna(how='any')
            headers = df.iloc[0].values.tolist()
            data = df.iloc[1:].values.tolist()
            window['summary'].update(values=data, num_rows=min(10,len(data)))

            df = pd.read_excel(workingFile, sheet_name=-1, header=None, usecols="K", engine='openpyxl')
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

    if event == 'report':
        report(workingFile, graphPath) 

    
window.close()

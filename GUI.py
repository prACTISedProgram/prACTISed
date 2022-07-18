
# First attempt at GUI layout for prACTISed

# 20220717 JL NOTE: graph jpegs from prACTISed.py(idealinputs.xlsx), missing concentration to determine peak, manual peak determination missing from prACTISed.py

from tkinter import *
from tkinter import ttk
from PIL import ImageTk, Image
import pandas as pd

root = Tk()
root.title("prACTISed User Interface")
##root.iconbitmap

main = Frame(root)
main.grid(padx=15, pady=15)

### Raw Data Directory ### 
directory = Frame(main, bg="white", padx=5, pady=5)
directory.grid(column=0, row=0, columnspan=2, padx=5, pady=5)

dataDirect_prompt = Label(directory, text="Raw Data Directory")
dataDirect_prompt.grid(column=0, row=1)
dataDirect_entry = Entry(directory, width=50)
dataDirect_entry.grid(column=1, row=1)
dataDirect_entry.insert(0,'C://')


### Fluidics Experimental Parameters ###
fluid = LabelFrame(main, text="Fluidics Experimental Parameters", bg="white", padx=5, pady=5)
fluid.grid(column=0,row=1,padx=5, pady=5)

flowRate_prompt = Label(fluid, text="Propagation flow rate (µL/min)")
flowRate_prompt.grid(column=0, row=0, sticky=E)
flowRate_entry = Entry(fluid, width=5)
flowRate_entry.grid(column=1, row=0)
flowRate = flowRate_entry.get()

injectRate_prompt = Label(fluid, text="Injection flow rate (µL/min)")
injectRate_prompt.grid(column=0, row=1, sticky=E)
injectRate_entry = Entry(fluid, width=5)
injectRate_entry.grid(column=1, row=1)
injectRate = injectRate_entry.get()

injectTime_prompt = Label(fluid, text="Injection time (s)")
injectTime_prompt.grid(column=0, row=2, sticky=E)
injectTime_entry = Entry(fluid, width=5)
injectTime_entry.grid(column=1, row=2)
injectTime = injectTime_entry.get()

capLength_prompt = Label(fluid, text="Separation capillary length (cm)")
capLength_prompt.grid(column=0, row=3, sticky=E)
capLength_entry = Entry(fluid, width=5)
capLength_entry.grid(column=1, row=3)
capLength = capLength_entry.get()

injectLength_prompt = Label(fluid, text="Injection loop length (cm)")
injectLength_prompt.grid(column=0, row=4, sticky=E)
injectLength_entry = Entry(fluid, width=5)
injectLength_entry.grid(column=1, row=4)
injectLength = injectLength_entry.get()


### Concentrations Used ###
concs = LabelFrame(main, text="Concentrations Used", bg="white", padx=5, pady=5)
concs.grid(column=0,row=2, padx=5, pady=5)

protConcs_prompt = Label(concs, text="Protein concentrations [P]0 (µM)\n(separated by a comma)")
protConcs_prompt.grid(column=0, row=0,columnspan=2)
protConcs_entry = Entry(concs, width=30)
protConcs_entry.grid(column=0, row=1, columnspan=2)
protConcs = protConcs_entry.get()

ligandConc_prompt = Label(concs, text="Initial ligand concentration [L]0 (µM)")
ligandConc_prompt.grid(column=0, row=2, sticky=E)
ligandConc_entry = Entry(concs, width=4)
ligandConc_entry.grid(column=1, row=2)
ligandConc = ligandConc_entry.get()


### Data Analysis Parameters ###
analysis = LabelFrame(main,text="Data Analysis Parameters", bg="white", padx=5, pady=5)
analysis.grid(column=0, row=3, padx=5, pady=5)

        # Might use radiobutton instead
dataType_prompt = Label(analysis, text="MS (S) or Fluorescense (F) data?")
dataType_prompt.grid(column=0, row=0, sticky=E)
dataType_entry = Entry(analysis, width=5)
dataType_entry.grid(column=1, row=0)
dataType = dataType_entry.get()

protRef_prompt = Label(analysis, text="If MS (S), select the P[0] \n reference for normalization (µM)")
protRef_prompt.grid(column=0, row=1, sticky=E)
protRef_entry = Entry(analysis, width=5)
protRef_entry.grid(column=1, row=1)
protRef = protRef_entry.get()

        # Might use radiobutton instead
        # EXTRA: Need to add manual determination of peak to prACTISed.py code
        # MISSING: prACTISed.py currently allows choice for the concentration used to determine peak
detPeak_prompt = Label(analysis, text="Manual (M) or programmic (P)\n determination of peak?")
detPeak_prompt.grid(column=0, row=2, sticky=E)
detPeak_entry = Entry(analysis, width=5)
detPeak_entry.grid(column=1, row=2)
detPeak = detPeak_entry.get()

window_prompt = Label(analysis, text="Window width (%)")
window_prompt.grid(column=0, row=3, sticky=E)
window_entry = Entry(analysis, width=5)
window_entry.grid(column=1, row=3)
window = window_entry.get()


### Signals ###             #Unsure if this is an input or an output
signal = LabelFrame(main,text="Signals", bg="white", padx=5, pady=5)
signal.grid(column=0, row=4, padx=5, pady=5)

protConc_prompt = Label(signal, text="[P]0 (µM)")
protConc_prompt.grid(column=0, row=0)
protConc1_entry = Entry(signal, width=5)
protConc1_entry.grid(column=0, row=1)
protConc1 = dataType_entry.get()
protConc2_entry = Entry(signal, width=5)
protConc2_entry.grid(column=0, row=2)
protConc2 = dataType_entry.get()
protConc3_entry = Entry(signal, width=5)
protConc3_entry.grid(column=0, row=3)
protConc3 = dataType_entry.get()

recTime_prompt = Label(signal, text="Recommended \n times (s)")
recTime_prompt.grid(column=1, row=0)
recTime1_entry = Entry(signal, width=7)
recTime1_entry.grid(column=1, row=1)
recTime1 = dataType_entry.get()
recTime2_entry = Entry(signal, width=7)
recTime2_entry.grid(column=1, row=2)
recTime2 = dataType_entry.get()
recTime3_entry = Entry(signal, width=7)
recTime3_entry.grid(column=1, row=3)
recTime3 = dataType_entry.get()

inputTime_prompt = Label(signal, text="Input \n times (s)")
inputTime_prompt.grid(column=2, row=0)
inputTime1_entry = Entry(signal, width=5)
inputTime1_entry.grid(column=2, row=1)
inputTime1 = dataType_entry.get()
inputTime2_entry = Entry(signal, width=5)
inputTime2_entry.grid(column=2, row=2)
inputTime2 = dataType_entry.get()
inputTime3_entry = Entry(signal, width=5)
inputTime3_entry.grid(column=2, row=3)
inputTime3 = dataType_entry.get()

sig_prompt = Label(signal, text="Signal")
sig_prompt.grid(column=3, row=0)
sig1_entry = Entry(signal, width=5)
sig1_entry.grid(column=3, row=1)
sig1 = dataType_entry.get()
sig2_entry = Entry(signal, width=5)
sig2_entry.grid(column=3, row=2)
sig2 = dataType_entry.get()
sig3_entry = Entry(signal, width=5)
sig3_entry.grid(column=3, row=3)
sig3 = dataType_entry.get()


### Calculate Kd Button ###
kd = Frame(main, bg="white", padx=5, pady=5)
kd.grid(column=0,row=5)
kd_button = Button(kd,text="Calculate Kd", padx=30, pady=10, bg="blue")
kd_button.grid(column=0, row=0)

### Output Window ###       NOTE: all data was taken manually from prACTISed for idealinputs.xlsx, needs adjustment to pass values
output = LabelFrame(main, text="Output", bg="white", padx=5, pady=5)
output.grid(column=1, row=1, rowspan=4, padx=5, pady=5)

# Return Kd and Statistic values
kd_output = Label(output, text="Kd: 27.7146 ± 2.6802 μM")
kd_output.grid(column=0, row=0)
Rsq_output = Label(output, text="R²: 0.9887")
Rsq_output.grid(column=1, row=0)
Chisq_output = Label(output, text="χ²: 0.1010")
Chisq_output.grid(column=2, row=0)

# Display Graph Images
# Get names of graphs from prACTISed.py         #Alternatively, use the figures directly from prACTISed.py (save optional)
graphNames = ('bindingisotherm.jpeg','allconcentration.jpeg','0.1µM.jpeg', '1µM.jpeg', '5µM.jpeg', '10µM.jpeg', '25µM.jpeg', '50µM.jpeg', '100µM.jpeg', '250µM.jpeg', '500µM.jpeg', '1000µM.jpeg')
graphImages = []

for img in graphNames:
    this_img = ImageTk.PhotoImage(Image.open(img))
    graphImages.append(this_img)

view = Label(output, image=graphImages[0])
view.grid(column=0, row=1, columnspan=3)
status = Label(output, text="1/"+str(len(graphImages)), pady=10)
status.grid(column=1,row=2)

# Back button for graph viewer
def back(imageNumber):
    global view
    global button_back
    global button_forward

    view.grid_forget()
    view = Label(output, image=graphImages[imageNumber])
    button_forward = Button(output, text=">>", command=lambda:forward(imageNumber+1))
    button_back = Button(output, text="<<", command=lambda:back(imageNumber-1))

    if imageNumber == 0:
        button_back = Button(output, text="<<", state=DISABLED)

    view.grid(column=0, row=1, columnspan=3)
    button_back.grid(column=0, row=2)
    button_forward.grid(column=2, row=2)

    status = Label(output, text=str(imageNumber+1)+"/"+str(len(graphImages)), pady=10)
    status.grid(column=1,row=2)

# Forward button for graph viewer
def forward(imageNumber):
    global view
    global button_back
    global button_forward

    view.grid_forget()
    view = Label(output, image=graphImages[imageNumber])
    button_forward = Button(output, text=">>", command=lambda:forward(imageNumber+1))
    button_back = Button(output, text="<<", command=lambda:back(imageNumber-1))

    if imageNumber == len(graphImages)-1:
        button_forward = Button(output, text=">>", state=DISABLED)
        
    view.grid(column=0, row=1, columnspan=3)
    button_back.grid(column=0, row=2)
    button_forward.grid(column=2, row=2)

    status = Label(output, text=str(imageNumber+1)+"/"+str(len(graphImages)), pady=10)
    status.grid(column=1,row=2)

button_back = Button(output, text="<<", state=DISABLED, command=back)
button_back.grid(column=0, row=2)                               
button_forward = Button(output, text=">>", command=lambda:forward(1))
button_forward.grid(column=2, row=2)



root.mainloop()

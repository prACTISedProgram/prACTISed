# prACTISed
PRocessing ACTIS Experimental Data

```prACTISed``` is a collection of scripts that allow ACTIS experimental data to be analyzed with minimal user input and a straightforward graphical user interface. prACTISed can generate a formatted Excel working file from Beckman Coulter/SCIEX fluorescence detector (Karat32) and API 5000 mass spectrometry/SCIEX detector (Analyst) experimental data. prACTISed also supportes general ASCII files or a correctly formatted Excel working file. ```prACTISed``` can be used on any platform that supports Python is licensed under the GNU General Public License 3 (GPLv3).

# Usage

ACTIS data can be analyzed by simply providing file path to a folder containing experimental run data or providing the file path to a formatted Excel working file such as. The script expects the working file to be organized in a certain format, see next section for generation or the provided ```idealinputs.xlsx``` as an example.

# Generating input files

If you wish to generate your own input file manually it must be as follows:
* The first sheet is called Inputs
  * It contains at minimum ```injection time```, ```protein name```, ```number of concentrations```, ```initial ligand concentration```, ```type of data``` and ```window width``` in column B
  * Additionally you must specify whether the ```compensation procedure``` must be applied and if so, indicate which ```concentration for normalization```
  * The method of ```peak determination``` must be indicated along with the peak to be used for ```programmatic peak determination``` or the times to use for ```manual peak determination```
  * Finally, all concentrations with units must be indicated in column E
* All subsequent sheets are too be named as concentration with units
  * Column A should have the raw times, and cell A1 must be labeled ```raw time```
  * All following columns should have signals for an experimental run and be titled ```Experiment 1``` and so on
* If compensation is required, a sheet titled ```P_simulated``` is required with ```raw time``` in column A and ```signal``` in column B

# Dependencies
## Python
* ```prACTISed``` requires Python version 3
## Libraries
* ```PySimpleGUI```
* ```os```
* ```pathlib```
* ```glob```
* ```sys```
* ```PIL``` 
* ```pandas```
* ```matplotlib```
* ```numpy```
* ```openpyxl```
* ```math```
* ```scipy```
* ```natsort```
* ```fpdf2```
* ```webbrowser```

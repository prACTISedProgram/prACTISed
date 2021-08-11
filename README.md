# prACTISed
Processing ACTIS experimental data

# Usage

ACTIS data can be analyzed by simply providing the input file to prACTISed:

```bash
prACTISed.py inputdata.xlsx
```

The script expects ```inputdata.xlsx``` to be in a certain format, see next section for generation or the provided ```idealinputs.xlsx``` as an example.

There are some options available, e.g. to show more detailed output while analyzing the data:

```bash
usage: prACTISed.py [-h] [--version] [-v] inputfile

prACTISed! This program analyzes ACTIS data and extracts the Kd-value.

positional arguments:
  inputfile

optional arguments:
  -h, --help     show this help message and exit
  --version      prints version information
  -v, --verbose  prints detailed output while analyzing
```


# Generating input files

tbd
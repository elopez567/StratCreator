# StratCreator

Allows the user to load a pool of loans and creates a full collateral table strat of the loan pool's characteristics.

## Description

PyStrat_V1.py is a Tkinter GUI python script that is used to automate the strat creation proccess for the collateral characteristics of a given loan pool. Fill out the Python Strat Input file and run it through the application in order to create a full strat with extensive details within seconds. Includes adjustable ranges and weighted average characteristics per individual range.

## Getting Started

### Dependencies

- Libraries needed: pandas, tkinter, openpyxl, and win32com
- Files needed: Python Strat Input

### Installing

- pip install the libraries stated under dependencies
- download the python strat input template and move to an appropriate folder
- move script into desired IDE

### Executing program

- Manually fill out Python Strat Input file (see example)
- Run script within desired IDE
- Click "Open File" Button
- Select the completed Python Strat Input file
- Change data elements within entry boxes as you please
- After you select a file wait until you see the path to the file show up in the text box below "Strat Input File"
- Click "Create Strat"
- The script will save the completed strat in the same folder where the loaded input file is locate
- Once done, the finished strat file will be automatically opened in excel

## Help

Common Issue: To prevent data elements from being excluded from the strat creation, please be sure to give some cushion above your ceiling cut-off. 

## Authors

Emmanuel Lopez - LinkedIn: (https://www.linkedin.com/in/lopez-emmanuel/)

## Version History

* 1.0
    * Initial Release

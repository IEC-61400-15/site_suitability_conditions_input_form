### IEC61400-15-1 site suitability input form converter (from *.xlsx to *.json)
## Written in June 15 2022 to June 17 2022 during the meeting in Copenhagen 
## VERY TEMPORARY CODE without much error treatment
## Only tested for the wind direction sector width is 30 degree and wind speed bin width is 1 m/s.
## Unexpected error may happen without any warnings.

## Atsushi Yamaguchi :  yamaguchi.atsushi@g.ashikaga-u.ac.jp

## Usage:
## xlsx2json.py [Input xlsx file] [Output json file]

## TODO
## 1) Avoid redundant code. Should be put in functions in nice way (without using global variables...)
## 2) Test with some more examples, especially with the case of wind direction sector number is
##    16 or wind speed bin width is 2.
## 3) The LFs between array elements should be removed in the ooutpuf file for readability.
## 4) Prepare GUI for easier use ??
## 5) Error and exeption treatment should be enhaced.


import openpyxl
import json
import sys

# 
args = sys.argv
InputXlsxFile = args[1]
OutputJsonFile = args[2]

wb = openpyxl.load_workbook(InputXlsxFile,data_only=True)
ws = wb['Project Information']


print("Analyzing the excel file...")

# Analyze the wind speed bin width (wsbw) and number of wind direction sectors (nwds)
ws = wb['WS Frequency']
wsbw = int(ws.cell(5,2).value)
nwds = int(360 / ws.cell(int (40 / wsbw + 5),1).value)

print("The wind speed bin width is",wsbw)
print("The number of wind direction sector is",nwds)

nwsb = int(40 / wsbw) + 1


# Analyze the number of wind turbines and their IDs
ws = wb['Turbine Layout Summary']
wtids = []
row = 2
while ws.cell(row,2).value != None:
    wtids.append(str(ws.cell(row,2).value))
    row = row + 1

nwt = len(wtids)

print("The number of wind turbine is",nwt)
print("Their IDs are",wtids)

# Analyse the number of measurement devices and their IDs
ws = wb['Measurement Device Summary']
mdids = []
row = 2
while ws.cell(row,1).value != None:
    mdids.append(str(ws.cell(row,1).value))
    row = row + 1

nmd = len(mdids)

print("The number of measurement devices is",nmd)
print("Their IDs are",mdids)

# Make the whole the list of WT + meas. device
mdwtids = mdids + wtids


# Meta data
OjMetaData = {}
OjMetaData["Number of wind direction secttors"] = nwds
OjMetaData["Wind speed bin width"] = wsbw
OjMetaData["Number of measurement devices"] = nmd
OjMetaData["Measurement device IDs"] = mdids
OjMetaData["Number of wind turbines"] = nwt
OjMetaData["Wind turbine IDs"] = wtids

# Project inforamtion
OjPI = {}
OjPI

# Turbine Layout Summary
OjTLS = {}
ws = wb['Turbine Layout Summary']
for wt in wtids:
    row = 2
    while (row < nwt + 2):
        if (str(ws.cell(row,2).value)) == wt:
            trow = row
            exit
        row = row + 1

    tmpobj = {}
    tmpobj["Project Name"] = ws.cell(trow,1).value
    tmpobj["Easting or Latitude"] = ws.cell(trow,3).value
    tmpobj["Northing or Latitude"] = ws.cell(trow,4).value
    tmpobj["Ground Elevation Above Sea Level"] = ws.cell(trow,5).value
    tmpobj["Wind Turbine Manufacturer"] = ws.cell(trow,6).value
    tmpobj["Wind Turbine Model"] = ws.cell(trow,7).value
    tmpobj["Wind Turbine Rated Power"] = ws.cell(trow,8).value
    tmpobj["Rotor Diameter"] = ws.cell(trow,9).value
    tmpobj["Hub Height"] = ws.cell(trow,10).value
    tmpobj["Associated Data Source"] = ws.cell(trow,11).value
    tmpobj["Ve50"] = ws.cell(trow,12).value
    tmpobj["V50"] = ws.cell(trow,13).value
    tmpobj["CoV"] = ws.cell(trow,14).value
    tmpobj["Air Density"] = ws.cell(trow,15).value
    tmpobj["Annual Average Wind Speed at Hub Height"] = ws.cell(trow,16).value
    tmpobj["Scale Parameter of Weibull Function"] = ws.cell(trow,17).value
    tmpobj["Shape Parameter of Weibull Function"] = ws.cell(trow,18).value
    tmpobj["Turbulence Structure Correction Parameter"] = ws.cell(trow,19).value
    tmpobj["Annual Mean Wind Shear"] = ws.cell(trow,20).value
    tmpobj["Annual Average Turbulence Intensity at 15 m/s"] = ws.cell(trow,21).value
    tmpobj["Standard Deviation of Turbulence Intensity at 15 m/s"] = ws.cell(trow,22).value
    tmpobj["Inflow Angle"] = ws.cell(trow,23).value
        
    OjTLS[wt] = tmpobj

# Measurement Device Summary
OjMDS = {}
ws = wb['Measurement Device Summary']
for md in mdids:
    row = 2
    while (row <= nmd + 2):
        if (str(ws.cell(row,1).value)) == md:
            trow = row
            exit
        row = row + 1
        
    EachMDS = {}
    EachMDS["Easting or Latitude"] = ws.cell(trow,2).value
    EachMDS["Northing or Latitude"] = ws.cell(trow,3).value
    EachMDS["Ground Elevation Above Sea Level"] = ws.cell(trow,4).value
    EachMDS["Measurement Device Height"] = ws.cell(trow,5).value

    OjMDS[md] = EachMDS


# WS Frequency
OjWSF = {}
ws = wb['WS Frequency']

# For measurement device
for md in mdids:
    column = 3
    while (column <= nwt + nmd*2 + 3):
        if (str(ws.cell(3,column).value)) == md:
            tcolumn = column
            exit
        column = column + 1

    tmpobj = {}

    # WS Frequency    
    tmparray2 = []    
    iwd = 1
    while iwd <= nwds:
        tmparray = []
        iws = 0
        while iws <= 40:
            row = 4+(iwd-1)*nwsb + iws
            tmparray.append(ws.cell(row,tcolumn).value * 100.0)
            iws = iws + wsbw
        tmparray2.append(tmparray)
        iwd = iwd + 1
        
    tmpobj['WS Frequency'] = tmparray2

    # WS number of samples
    tmparray2 = []    
    iwd = 1
    while iwd <= nwds:
        tmparray = []
        iws = 0
        while iws <= 40:
            row = 4+(iwd-1)*nwsb + iws
            tmparray.append(int(ws.cell(row,tcolumn + 1).value))
            iws = iws + wsbw
        tmparray2.append(tmparray)
        iwd = iwd + 1
        
    tmpobj['WS Number of samples'] = tmparray2

    OjWSF[md]=tmpobj


# For wind turbines
for wt in wtids:
    column = 3
    while (column <= nwt + nmd*2 + 3):
        if (str(ws.cell(3,column).value)) == wt:
            tcolumn = column
            exit
        column = column + 1

    tmpobj = {}
    # WS Frequency    
    tmparray2 = []    
    iwd = 1
    while iwd <= nwds:
        tmparray = []
        iws = 0
        while iws <= 40:
            row = 4+(iwd-1)*nwsb + iws
            tmparray.append(ws.cell(row,tcolumn).value * 100.0)
            iws = iws + wsbw
        tmparray2.append(tmparray)
        iwd = iwd + 1
        
    tmpobj['WS Frequency'] = tmparray2

    OjWSF[wt]=tmpobj
    

# WS Weibull
ws = wb['WS Weibull']
OjWSW = {}

# For both meas. device and wt
for mdwt in mdwtids:
    column = 3
    tmpobj={}
    while (column <= nwt + nmd + 2):
        if (str(ws.cell(3,column).value)) == mdwt:
            tcolumn = column
            exit
        column = column + 1

    tmpobj['WS Weibull scale parameter all directions'] = ws.cell(4,tcolumn).value
    tmpobj['WS Weibull shape parameter all directions'] = ws.cell(5 + nwds,tcolumn).value

    tmparray1=[]
    tmparray2=[]
    tmparray3=[]
    iwd = 1
    while(iwd <= nwds):
        tmparray1.append(ws.cell(3 + iwd,tcolumn).value)
        tmparray2.append(ws.cell(5 + nwds + iwd,tcolumn).value)
        tmparray3.append(ws.cell(5 + 2 * nwds + iwd, tcolumn).value * 100.0)
        iwd = iwd + 1
        
    tmpobj['WS Weibull scale parameter'] = tmparray1
    tmpobj['WS Weibull shape parameter'] = tmparray2
    tmpobj['WS Weibull frequency'] = tmparray3
    
    OjWSW[mdwt]=tmpobj

    
## Ambient Mean TI
ws = wb['Ambient Mean TI']
OjAMTI = {}

# For meas. device and WT
for mdwt in mdwtids:
    column = 3
    while (column <= nwt + nmd + 2):
        if (str(ws.cell(3,column).value)) == mdwt:
            tcolumn = column
            exit
        column = column + 1

    tmpobj = {}

    # For all wind direction
    tmparray = []
    iws = 0
    while iws <= 40:
        row = 4 + iws
        tmparray.append(ws.cell(row,tcolumn).value*100.0)
        iws = iws + wsbw
    tmpobj["Ambient mean TI all directions"] = tmparray

    
    # For each wind direction 
    tmparray2 = []    
    iwd = 1
    while iwd <= nwds:
        tmparray = []
        iws = 0
        while iws <= 40:
            row = 4 + iwd * nwsb + iws
            tmparray.append(ws.cell(row,tcolumn).value * 100.0)
            iws = iws + wsbw
        tmparray2.append(tmparray)
        iwd = iwd + 1
    tmpobj['Ambient mean TI'] = tmparray2

    
    OjAMTI[mdwt] = tmpobj


## SD TI
ws = wb['SD TI']
OjSDTI = {}

# For meas. device and WT
for mdwt in mdwtids:
    column = 3
    while (column <= nwt + nmd + 2):
        if (str(ws.cell(3,column).value)) == mdwt:
            tcolumn = column
            exit
        column = column + 1

    tmpobj = {}

    # For all wind direction
    tmparray = []
    iws = 0
    while iws <= 40:
        row = 4 + iws
        tmparray.append(ws.cell(row,tcolumn).value*100.0)
        iws = iws + wsbw
    tmpobj["SD TI all directions"] = tmparray

    
    # For each wind direction 
    tmparray2 = []    
    iwd = 1
    while iwd <= nwds:
        tmparray = []
        iws = 0
        while iws <= 40:
            row = 4 + iwd * nwsb + iws
            tmparray.append(ws.cell(row,tcolumn).value * 100.0)
            iws = iws + wsbw
        tmparray2.append(tmparray)
        iwd = iwd + 1
    tmpobj['SD TI'] = tmparray2

    OjSDTI[mdwt] = tmpobj

## Extreme Ambient TI
ws = wb['Extreme Ambient TI']
OjEATI = {}

# For meas. device and WT
for mdwt in mdwtids:
    column = 3
    while (column <= nwt + nmd + 2):
        if (str(ws.cell(3,column).value)) == mdwt:
            tcolumn = column
            exit
        column = column + 1

    tmpobj = {}

    # For all wind direction
    tmparray = []
    iws = 0
    while iws <= 40:
        row = 4 + iws
        tmparray.append(ws.cell(row,tcolumn).value * 100)
        iws = iws + wsbw
    tmpobj["SD TI all directions"] = tmparray
    
    OjEATI[mdwt]= tmpobj
    
## Temperature
ws = wb['Temperature']
OjTEMP = {}

# For meas. device and WT
for md in mdids:
    ## Prepare empty dictionary
    tmpobj = {}
    
    ## Scalar variables
    column = 2
    while (column <= nmd + 1):
        if (str(ws.cell(1,column).value)) == md:
            tcolumn = column
            exit
        column = column + 1

        
    tmpobj['Yearly mean ambient Temperature'] = ws.cell(2,tcolumn).value
    tmpobj['Days per year with at least 1 hour below -20 deg'] = ws.cell(3,tcolumn).value

    ## Vector variables
    column = 2
    while (column <= nmd * 2 + 1):
        if (str(ws.cell(5,column).value)) == md:
            tcolumn = column
            exit
        column = column + 1


    tmparray1 = []
    tmparray2 = []
    itemp = 1
    while itemp <= 91:
        row = 5 + itemp
        tmparray1.append(ws.cell(row,tcolumn).value)
        tmparray2.append(int(ws.cell(row,tcolumn+1).value))
        itemp = itemp + 1
    tmpobj["Temperature frequency"] = tmparray1
    tmpobj["Number of sanples"] = tmparray2
       
    OjTEMP[md] = tmpobj

## Shear    
ws = wb['Shear']
OjSHEAR = {}

# For meas. device and WT
for mdwt in mdwtids:
    tmpobj = {}
    column = 3
    while (column <= nwt + nmd + 2):
        if (str(ws.cell(3,column).value)) == mdwt:
            tcolumn = column
            exit
        column = column + 1

    tmpobj["Shear all directions"] = ws.cell(4,tcolumn).value
    tmparray = []
    iwd = 1
    while iwd <= nwds:
        row = 4 + iwd
        tmparray.append(ws.cell(row,tcolumn).value)
        iwd = iwd + 1

    tmpobj["Directional shear"] = tmparray

    OjSHEAR[mdwt]=tmpobj
    

## Inflow Angle
ws = wb['Inflow Angle']
OjIA = {}

# For meas. device and WT
for mdwt in mdwtids:
    tmpobj = {}
    column = 2
    while (column <= nwt + nmd + 1):
        if (str(ws.cell(3,column).value)) == mdwt:
            tcolumn = column
            exit
        column = column + 1

    tmpobj["Inflow angle all directions"] = ws.cell(4,tcolumn).value
    tmpobj["Inflow angle max"] = ws.cell(5,tcolumn).value

    tmparray = []
    iwd = 1
    while iwd <= nwds:
        row = 5 + iwd
        tmparray.append(ws.cell(row,tcolumn).value)
        iwd = iwd + 1

    tmpobj["Directional Inflow angle"] = tmparray

    OjIA[mdwt]=tmpobj

##CcT
ws = wb['CcT']
OjCCT = {}

# For meas. device and WT
for mdwt in mdwtids:
    tmpobj = {}
    column = 2
    while (column <= nwt + nmd + 1):
        if (str(ws.cell(3,column).value)) == mdwt:
            tcolumn = column
            exit
        column = column + 1

    tmpobj["sigma 3/sigma 1"] = ws.cell(4,tcolumn).value
    tmpobj["sigma 2/sigma 1"] = ws.cell(5,tcolumn).value
    tmpobj["CcT"] = ws.cell(6,tcolumn).value

    OjCCT[mdwt]=tmpobj
    
    
# Make top level keys
top = {}
top["Meta data"] = OjMetaData
top["Turbine Layout Summary"] = OjTLS
top["Measurement Device Summary"] = OjMDS
top["WS Frequency"] = OjWSF
top["WS Weibull"] = OjWSW
top["Ambient Mean TI"] = OjAMTI
top["SD TI"] = OjSDTI
top["Extreme Ambient TI"] = OjEATI
top["Temperature"] = OjTEMP
top["Shear"] = OjSHEAR
top["Inflow Angle"] = OjIA
top["CcT"] = OjCCT


# Damp the json file
with open(OutputJsonFile, 'w') as f:
    json.dump(top, f, ensure_ascii=False, indent=4)
    
    
    

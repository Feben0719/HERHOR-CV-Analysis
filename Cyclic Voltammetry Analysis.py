# -*- coding: utf-8 -*-
"""
Created on Mon Jan 23 14:07:27 2023

@author: feben
"""

import openpyxl
from scipy import interpolate
import matplotlib.pyplot as plt
import numpy as np
from CorrectionLibrary import*

#Before using this code, create two sheets within one file for your raw data
#One sheet should contain the test using the bare working electrode
#The second sheet should contain the test the working electrode with the additive
#Potential data in column A; Current data in column B

wb = openpyxl.load_workbook('31323.xlsx') #Opens the workbook
Sheet_Bare = wb['Bare'] #Opens the sheet using the bare electrode data
Sheet_Additive = wb['Add'] #Opens the sheet using the additive electrode data

BareV_list = [] #Creates empty list for potential data from bare test
BareI_list = [] #Creates empty list for current data from bare test

AddV_list = [] #Creates empty list for potential data additive test
AddI_list = [] #Creates empty list for current data from additive test


BareV_cell_value = Sheet_Bare['A2']
BareI_cell_value = Sheet_Bare['B2']
AddV_cell_value = Sheet_Additive['A2']
AddI_cell_value = Sheet_Additive['B2']
empty_cellB = Sheet_Bare['Z1']
empty_cellA = Sheet_Additive['Z1']
ratio = []


#Following two loops count the number of data points
count = 0
P_1 = Sheet_Bare['A2']
Q_1 = Sheet_Additive['A2']



for i in range(1, 999):
    if P_1.value == empty_cellB.value:
        break
    else:
        count += 1
        P_1 = Sheet_Bare['A{}'.format(i)]
        
for i in range(1, 999):
    if Q_1.value == empty_cellA.value:
        break
    else:
        count += 1
        Q_1 = Sheet_Additive['A{}'.format(i)]

#Following two functions add data to the empty lists 
for i in range(2, count):
    BareV_cell_value  = Sheet_Bare['A{}'.format(i)]
    BareI_cell_value = Sheet_Bare['B{}'.format(i)]
    BareV_list.append(BareV_cell_value.value)
    BareI_list.append(BareI_cell_value.value)
    i += 1

for i in range(2, count):
    AddV_cell_value  = Sheet_Additive['A{}'.format(i)]
    AddI_cell_value = Sheet_Additive['B{}'.format(i)]
    AddV_list.append(AddV_cell_value.value)
    AddI_list.append(AddI_cell_value.value)
    i += 1
# Following functions will create data lists without 'None' 
xb = []
for item in BareV_list:
    if item != None:
        xb.append(item)
        

yb = []
for item in BareI_list:
    if item != None:
        yb.append(item)       
 
xa = []
for item in AddV_list:
    if item != None:
        xa.append(item) 
       
ya = []
for item in AddI_list:
    if item != None:
        ya.append(item)  
        

  
#Following code will correct for RHE and current density

xb = np.subtract(xb, P115AgCl)
yb = np.multiply(yb, CD) - 1.48

xa = np.subtract(xa, P11AgCl)
ya = np.multiply(ya, CD)

#Following code will calculate slope at every data point
slopeb = np.diff(yb)/np.diff(xb)
iob = slopeb/(96485/(298*8.3145))  
slopea = np.diff(ya)/np.diff(xa)
ioa = slopea/(96485/(298*8.3145))

#for i in range(len(slopeya)-1):
#    a = slopeya[i]-slopeya[i+1]
#    ratio.append(a)
#print((max(ratio)), ratio.index(max(ratio)))


 #Following will create two tables summarizing potential, current density, and instataneous slopes      
Line = "\n{:4s}{:8s}{:20s}{:20s}{:20s}".format('','x','y', 'slope', 'Exchange Current Density')
print('Bare Working Electrode')
print(Line)
for (xb_k,yb_k,slopeb_k, iob_k) in zip(xb,yb, slopeb, iob):
    Line = "{:8.4f}{:20.4f}{:20.4f}{:20.4f}".format(xb_k,yb_k, slopeb_k, iob_k)
    print(Line)
'''    
Line2 = "\n{:4s}{:8s}{:20s}{:20s}{:20s}".format('','x','y', 'slope', 'Exchange Current Density')
print('Additive Working Electrode')
print(Line2)
for (xa_k,ya_k, slopea_k, ioa_k) in zip(xa,ya, slopea, ioa):
    Line = "{:8.4f}{:20.4f}{:20.4f}{:20.4f}".format(xa_k,ya_k, slopea_k, ioa_k)
    print(Line)    
'''
#manually enter tangent line points here (read from table):
#look at where y = 0 and take +-10 to automate it
xb_t = 0. # x coordinate of tangent with bare data
yb_t = 0.# y coordinate of tangent with bare data
mb_t = 6.778 #slope of tangent with bare data

xa_t = 0. # x coordinate of tangent with additive data
ya_t = 0. # y coordinate of tangent with additive data
ma_t = 9.4553 # slope of tangent with additive data


plt.plot(xb,yb,color='black',label='Bare')
plt.plot(xa,ya,color='green',label='Additive')
plt.plot(xb_t,yb_t, "or")
plt.plot(xa_t, ya_t, "or")
plt.xlabel('Potential v RHE (V)',fontsize=12)
plt.ylabel('Current Density (mA/cm^2)',fontsize=12)
plt.title('HER/HOR pH 11: Phosphate Buffer on Platinum')
#plt.xlim([0.22,0.24])
plt.legend()
plt.show()

#Equation tangent line y = mx - mx1 + y1

#plotting tangent lines
plt.plot(xb,yb,color='black',label='Bare')
plt.plot(xb_t,yb_t, "or")
xbb = np.linspace(-0.2,0.2,5) #change depending on xb_t
tangentB = mb_t * xbb - mb_t*xb_t + yb_t
plt.plot(xbb, tangentB, color='grey', label='Tangent Bare', linestyle='dashed')
plt.xlabel('Potential v RHE (V)',fontsize=12)
plt.ylabel('Current Density (mA/cm^2)',fontsize=12)
plt.title('HER/HOR pH 11: Bare Platinum Electrode')
plt.legend()
plt.show()

plt.plot(xa,ya,color='green',label='Additive')
plt.plot(xa_t, ya_t, "or")
xaa = np.linspace(-0.2,0.2,5) #change depending on yb_t
tangentA = ma_t * xaa - ma_t*xa_t + ya_t
plt.plot(xaa, tangentA, color='grey', label='Tangent Additive', linestyle='dashed')
plt.xlabel('Potential v RHE (V)',fontsize=12)
plt.ylabel('Current Density (mA/cm^2)',fontsize=12)
plt.title('HER/HOR pH 11: Additive on Platinum Electrode')
plt.legend()
plt.show()

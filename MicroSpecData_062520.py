# -*- coding: utf-8 -*-
"""
Created on Thu May 14 11:19:03 2020

@author: Jacob Nichols

This is a script to be used for compiling SEE Microspectrophotometer spectra files into an 
excel workbook containing spectral and color data. All spectra files should be corrected using the
SEE software before being compiled by this script. 

"""

import os
import math
import numpy as np
import spc
import pandas as pd
from tkinter import filedialog
from tkinter import *

x_bar = [0.0002, 0.0024, 0.0191, 0.0847, 0.2045, 0.3147, 0.3837, 0.3707, 0.3023, 0.1956, 0.0805, 0.0162, 0.0038, 0.0375, 0.1177, 0.2365, 0.3768, 0.5298, 0.7052, 0.8787, 1.0142, 1.1185, 1.124, 1.0305, 0.8563, 0.6475, 0.4316, 0.2683, 0.1526, 0.0813, 0.0409, 0.0199, 0.0096, 0.0046, 0.0022]
y_bar = [0, 0.0003, 0.002, 0.0088, 0.0214, 0.0387, 0.0621, 0.0895, 0.1282, 0.1852, 0.2536, 0.3391, 0.4608, 0.6067, 0.7618, 0.8752, 0.962, 0.9918, 0.9973, 0.9556, 0.8689, 0.7774, 0.6583, 0.528, 0.3981, 0.2835, 0.1798, 0.1076, 0.0603, 0.0318, 0.0159, 0.0077, 0.0037, 0.0018, 0.0008]
z_bar = [0.0007, 0.0105, 0.086, 0.3894, 0.9725, 1.5535, 1.9673, 1.9948, 1.7454, 1.3176, 0.7721, 0.4153, 0.2185, 0.112, 0.0607, 0.0305, 0.0137, 0.004, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
D65_10 = [49.98, 54.65, 82.75, 91.49, 93.43, 86.68, 104.86, 117.01, 117.81, 114.86, 115.92, 108.81, 109.35, 107.80, 104.79, 107.69, 104.41, 104.05, 100.00, 96.33, 95.79, 88.69, 90.01, 89.60, 87.70, 83.29, 83.70, 80.03, 80.21, 82.28, 78.28, 69.72, 71.61, 74.35, 61.60]
normalization_constant = 0

index = ['L*', 'a*', 'b*', 'C*', 'Hue', 'x', 'y', 'X', 'Y', 'Z']

df_obs_funcs = pd.DataFrame()
df_init = pd.DataFrame()
df_final = pd.DataFrame()
df_color = pd.DataFrame()
df_colorAnalysis = pd.DataFrame()
df_fullSpec = pd.DataFrame()
file_list = []


def set_constants():
    global x_bar
    global y_bar
    global z_bar
    global D65_10
    global normalization_constant
    global Xn
    global Yn
    global Zn
    total = 0
    for num in range(0,len(y_bar)):
        temp = y_bar[num] * D65_10[num]
        total+=temp
    normalization_constant = 100/total
    total_x = 0
    total_y = 0
    total_z = 0
    for num in range(0,len(y_bar)):
        temp_x = x_bar[num] * D65_10[num]
        temp_y = y_bar[num] * D65_10[num]
        temp_z = z_bar[num] * D65_10[num]
        total_x += temp_x
        total_y += temp_y
        total_z += temp_z
    Xn = total_x * normalization_constant
    Yn = total_y * normalization_constant
    Zn = total_z * normalization_constant


def fill_df():
    global df_init
    global file_list
    count = 0
    #scan is each file name as a str
    #loop through each file in the selected directory
    for scan in os.listdir():
        fileName = scan
        if fileName[-3:] == 'spc':
            file_list.append(fileName)
            count += 1
            f = spc.File(fileName)   #Initialize spectra object f. Use f to call the x and y components.
            wave_list = f.x   #list of all the wavelengths
            df_init['wavelength' + str(count)] = wave_list   #casting the wavelength list into the dataframe
            percentR = f.sub[0].y   #list of all the R% values
            df_init[fileName] = percentR   #casting the R% list into the dataframe
        else:
            pass
    return count

def check_wavelength(num_files):
    global df_init
    for i in range(1,num_files):
        if df_init['wavelength' + str(i)].equals(df_init['wavelength' + str(i+1)]) == False:
            print('The wavelength range needs to be the same for all files in the selected directory.')
            return False
        else:
            return True

def set_index(boolean, num_files):
    global df_init
    if boolean == True:
        for i in range(1,num_files+1):
            if i==1:
                df_init['Wavelength'] = df_init['wavelength'+ str(i)]
                df_init.drop('wavelength' + str(i), axis=1, inplace=True)
            else:
                df_init.drop('wavelength' + str(i), axis=1, inplace=True)
    else:
        print('Error setting index.')
        exit()

def make_new_wavelist():
    global df_init
    new_wavelist = []
    for num in df_init['Wavelength']:
        waveByTen = int(num)%10
        if waveByTen == 0:
            new_wavelist.append(int(num))
        else:
            pass
    wavelist_byTen = [380]
    for item in new_wavelist:
        if item not in wavelist_byTen and item < 730:
            wavelist_byTen.append(item)
    return wavelist_byTen

def compress_data(wavelength):
    global df_init
    global df_final
    global file_list
    df_final['Wavelength'] = wavelength
    count = 0
    for i in df_init['Wavelength']:
        count+=1
    for f in file_list: #loop through each col of R% values
        filename = f
        valueholder = [0]
        for j in wavelength:    #loop through each wavelist_byTen value
            for k in range(0,count):    #loop through each R% value in the col
                if k==0:            #skip the first R% value
                    continue
                elif df_init['Wavelength'][k-1] <= j and j <= df_init['Wavelength'][k]: 
                    adjusted_r = (((j-df_init['Wavelength'][k-1])/(df_init['Wavelength'][k]-df_init['Wavelength'][k-1])) * (df_init[f][k]-df_init[f][k-1])) + (df_init[f][k-1])
                    valueholder.append(adjusted_r)
                else:
                    pass
        valueholder[0] = valueholder[1]        
        df_final[filename] = valueholder
    df_final['x_bar'] = x_bar
    df_final['y_bar'] = y_bar
    df_final['z_bar'] = z_bar
    df_final['D65_10'] = D65_10
        
def fill_df_color():
    global df_final
    global file_list
    global df_color
    global index
    global normalization_constant
    global Xn
    global Yn
    global Zn
    for f in file_list:
        X = (df_final['x_bar'] * df_final['D65_10'] * df_final[f] / 100).sum() * normalization_constant
        Y = (df_final['y_bar'] * df_final['D65_10'] * df_final[f] / 100).sum() * normalization_constant
        Z = (df_final['z_bar'] * df_final['D65_10'] * df_final[f] / 100).sum() * normalization_constant
        x = X/(X+Y+Z)
        y = Y/(X+Y+Z)
        L = 116*((Y/Yn)**(1/3)) - 16
        a = 500*(((X/Xn)**(1/3))-((Y/Yn)**(1/3)))
        b = 200*(((Y/Yn)**(1/3))-((Z/Zn)**(1/3)))
        C = math.sqrt((a**2 + b**2))
        if b < 0:
            h = 360 + np.degrees(np.arctan2(b,a))
        else:
            h = np.degrees(np.arctan2(b,a))
        df_color[f] = [L, a, b , C, h, x, y, X, Y, Z]
    df_color['Index'] = index
    df_color.set_index('Index', inplace = True)

def trim_df_final():
    global df_final
    df_final.drop('x_bar', axis = 'columns', inplace = True)
    df_final.drop('y_bar', axis = 'columns', inplace = True)
    df_final.drop('z_bar', axis = 'columns', inplace = True)
    df_final.drop('D65_10', axis = 'columns', inplace = True)
    df_final.set_index('Wavelength', inplace = True)

def fix_df_init():
    global df_init
    global file_list
    cols = ['Wavelength']
    cols = cols + file_list
    df_init = df_init[cols]
    df_init.set_index('Wavelength', inplace = True)


'''
def create_df_colorAnalysis():
    global df_color
    global file_list
    L = []
    a = []
    b = []
    C = []
    h = []
    X = []
    Y = []
    Z = []
    for file in filelist:
        
    
    
    
'''
root = Tk()
root.direct =  filedialog.askdirectory(initialdir = '/', title = "Select Directory")
print()
print("Select the directory with your Spectra Files.")
os.chdir(root.direct)
#root.filename = filedialog.asksaveasfile(initialdir = root.direct, title = 'Save As...')
#print(root.filename)
#new_filename = root.filename
#root.destroy()

set_constants()
num_Files = fill_df()
set_index(check_wavelength(num_Files),num_Files)
compress_data(make_new_wavelist())
fill_df_color()
trim_df_final()
fix_df_init()
#print(file_list)
#print(df_init)
writer = pd.ExcelWriter('MicroSpecData.xlsx')
df_final.to_excel(writer, sheet_name = 'Spectra')
df_color.to_excel(writer, sheet_name = 'Color')
df_init.to_excel(writer, sheet_name = 'Full Spectra')
writer.save()
print()
print('Script Complete')
print()
print('Your data file, "MicroSpecData.xlsx" is saved here:')
print()
print(root.direct)
#print('Your data file is saved in the same directory as your spectral files.')
root.destroy()

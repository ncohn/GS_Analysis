#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Jun  3 15:04:45 2021

@author: rdchlntc
"""

import numpy as np
import os
import xlsxwriter

#Put the Directory here with all the files you want to load
directory = '/Users/rdchlntc/Documents/Projects/Other/For Liz Davis'

#first loop through to determine how many xle files are present
os.chdir(directory)
count = -1
for filename in os.listdir(directory):
    f = os.path.join(directory, filename)
    if os.path.isfile(f) and filename.endswith('.xle'):
        count = count+1
        
#set up corresponding variables for storage + output
d2 = np.zeros([count+1])  
d5 = np.zeros([count+1])  
d10 = np.zeros([count+1])  
d25 = np.zeros([count+1])  
d50 = np.zeros([count+1])  
d75 = np.zeros([count+1])  
d90 = np.zeros([count+1])  
d95 = np.zeros([count+1])  
d98 = np.zeros([count+1])  
bl_avg = np.zeros([count+1])  
spht_avg = np.zeros([count+1])  
symm_avg = np.zeros([count+1])  
perc_very_coarse_sand = np.zeros([count+1])  
perc_coarse_sand = np.zeros([count+1])  
perc_medium_sand = np.zeros([count+1])  
perc_fine_sand = np.zeros([count+1])  
perc_very_fine_sand = np.zeros([count+1])  


#set up excel spreadsheet
workbook = xlsxwriter.Workbook('camsizer_outputs.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write(0,0, 'Filename')
worksheet.write(0,1, 'Site')
worksheet.write(0,2, 'CoreNumber')
worksheet.write(0,3, 'SampleNumber')
worksheet.write(0,4, 'D2(mm)')
worksheet.write(0,5, 'D5(mm)')
worksheet.write(0,6, 'D10(mm)')
worksheet.write(0,7, 'D25(mm)')
worksheet.write(0,8, 'D50(mm)')
worksheet.write(0,9, 'D75(mm)')
worksheet.write(0,10, 'D90(mm)')
worksheet.write(0,11, 'D95(mm)')
worksheet.write(0,12, 'D98(mm)')
worksheet.write(0,13, 'Symmetry_Avg')
worksheet.write(0,14, 'Sphericity_Avg')
worksheet.write(0,15, 'B/L_Avg')
worksheet.write(0,16, 'Perc_VeryCoarseSand')
worksheet.write(0,17, 'Perc_CoarseSand')
worksheet.write(0,18, 'Perc_MediumSand')
worksheet.write(0,19, 'Perc_FineSand')
worksheet.write(0,20, 'Perc_VeryFineSand')

#loop through again to actually calculate data
count = -1
for filename in os.listdir(directory):
    f = os.path.join(directory, filename)
    if os.path.isfile(f) and filename.endswith('.xle'):
                
        data = np.loadtxt(filename, delimiter='\t', skiprows=31, encoding = "utf-16")
        count = count + 1
        
        #determine grain size bounds
        lower_bound_size_mm = np.array(data[:,0])
        upper_bound_size_mm = np.array(data[:,1])
        upper_bound_size_mm[-1] = 10
        grain_size_mm = lower_bound_size_mm/2 + upper_bound_size_mm/2
        
        #store other shape data
        p= data[:,2]
        q=data[:,3]
        spht = data[:,4]
        symm = data[:,5]
        bl = data[:,6]
        pdn = data[:,7]
             
        #calc relevant stats
        d2[count] = np.interp(2, q, grain_size_mm)
        d5[count] = np.interp(5, q, grain_size_mm)
        d10[count] = np.interp(10, q, grain_size_mm)
        d25[count] = np.interp(25, q, grain_size_mm)
        d50[count] = np.interp(50, q, grain_size_mm)
        d75[count] = np.interp(75, q, grain_size_mm)
        d90[count] = np.interp(90, q, grain_size_mm)
        d95[count] = np.interp(95, q, grain_size_mm)
        d98[count] = np.interp(98, q, grain_size_mm)       
        ifind = np.where((d25[count] <= grain_size_mm) & (grain_size_mm <= d75[count]))
        symm_avg[count] = np.nanmean(symm[ifind])        
        spht_avg[count] = np.nanmean(spht[ifind])        
        bl_avg[count] = np.nanmean(bl[ifind])   
                
        
        #ifind = np.where((grain_size_mm >= 1) & (grain_size_mm < 2))
        perc_very_coarse_sand[count] = np.sum(p[np.where((grain_size_mm >= 1) & (grain_size_mm < 2))])
        perc_coarse_sand[count] = np.sum(p[np.where((grain_size_mm >= 0.5) & (grain_size_mm < 1))])
        perc_medium_sand[count] = np.sum(p[np.where((grain_size_mm >= 0.25) & (grain_size_mm < 0.5))])
        perc_fine_sand[count] = np.sum(p[np.where((grain_size_mm >= 0.125) & (grain_size_mm < 0.25))])
        perc_very_fine_sand[count] = np.sum(p[np.where((grain_size_mm >= 0.0625) & (grain_size_mm < 0.125))])
        
        
        #write out data
        site = filename[0:3]
        core = np.array(filename[3:5])
        samplenum = np.array(filename[8:11])
        worksheet.write(count+1,0, filename)
        worksheet.write(count+1,1, site)
        worksheet.write(count+1,2, core)
        worksheet.write(count+1,3, samplenum)        
        worksheet.write(count+1,4, d2[count])
        worksheet.write(count+1,5, d5[count])
        worksheet.write(count+1,6, d10[count])
        worksheet.write(count+1,7, d25[count])
        worksheet.write(count+1,8, d50[count])
        worksheet.write(count+1,9, d75[count])
        worksheet.write(count+1,10, d90[count])
        worksheet.write(count+1,11, d95[count])
        worksheet.write(count+1,12, d98[count])
        worksheet.write(count+1,13, symm_avg[count])
        worksheet.write(count+1,14, spht_avg[count])
        worksheet.write(count+1,15, bl_avg[count])
        worksheet.write(count+1,16, perc_very_coarse_sand[count])
        worksheet.write(count+1,17, perc_coarse_sand[count])
        worksheet.write(count+1,18, perc_medium_sand[count])
        worksheet.write(count+1,19, perc_fine_sand[count])
        worksheet.write(count+1,20, perc_very_fine_sand[count])
                                
#close and save excel spreadsheet        
workbook.close()
        

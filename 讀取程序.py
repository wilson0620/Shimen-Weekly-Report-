'''
Actually, I'm not good at programing Q Q
'''

import os
import csv
import pandas as pd

All_file_list = os.listdir()
All_dic = {}


def read_file(file_name):
    file = pd.read_excel(str(file_name), header=2)
    Head = file.columns
    for i in range(len(Head)):
        
        if Head[i] not in All_dic:
            All_dic[Head[i]] = [str(Head[i])]
        
        All_dic[Head[i]].append(file[Head[i]][0])
    
    print('%s done.'%file_name)
    print('')



for i in range(len(All_file_list)):
    if All_file_list[i] != '讀取程序.py':
        if 'xlsx' in All_file_list[i]:
            read_file(All_file_list[i])
        
print('Startup complete.')

def Yixing():
    target = ['時間' ,'CR1', 'CR2', 'CR3', 'CR4', 'CR5', 'CR6', 'N1', 'N2', 'N3', 'TI1_X', 'TI1_Y', 'TI2_X', 'TI2_Y', 'TI3_X', 'TI3_Y', 'L1', 'L2', 'L3', 'L4', 'L5', 'L6', 'L7', 'L8', 'S4_5_X', 'S6_1_X', 'S13_6_X', 'S14_3_X', 'LD1', 'LD2', 'LD3', 'LD4']
    out_list = []
    
    for i in target :
        out_list.append(All_dic[i])
    with open('Yixing.csv', 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerows(out_list)
    
    print('義興輸出成功')

def Ronghua():
    target = ['PL1_Y', 'PL2_Y', 'PL3_Y', 'PL4_Y', 'WS1', 'WS2', 'WS3', 'WS4', 'ETL1', 'ETL2', 'ETL3', 'ETR1', 'ETR2', 'ETR3']
    out_list = []
    
    for i in target :
        out_list.append(All_dic[i])
    with open('Ronghua.csv', 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerows(out_list)
    
    print('榮華輸出成功')
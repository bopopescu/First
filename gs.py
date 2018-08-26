# -*- coding: utf-8 -*-
"""
Created on Sat Aug  4 21:56:52 2018
test
@author: Administrator
"""
import numpy as np
import os
import openpyxl as xl
print('load gs')
A = np.ones(3)
B = np.ones(3)
E = np.ones(3)
F = np.ones(3)
Ap = np.ones(3)
Fp2 = np.ones(3)
Fp = np.ones(3)
A2 = np.ones(3)
B2 = np.ones(3)
E2 = np.ones(3)
F2 = np.ones(3)

xm1 = 0.
xm2 = 0.
xm3 = 0.
xm12 = 0.
xm22 = 0.

xs1 = 0.
xs2 = 0.
xs3 = 0.
xs12 = 0.
xs22 = 0.
BC = 50.
CD = 250.
ED = 62.
BC2 = 50.
CD2 = 450.
ED2 = 62.
Distance = 10.
Distance2 = 10.
errorP = 0.
errorN = 0.
Delta = 0.
excel_out = ''
w2Target = 0.
w2cal = 0.
w2opt = 0.
w3Target = 0.
w3cal = 0.
w3opt = 0.
minArray = []
maxArray = []
minArray2 = []
maxArray2 = []

Alfa = 360.
num = 0
Delta = 0.
Delta2 = 0.
BC2 = 0.
styleTime = ''
w2Target = 0.
w2cal = 0.
w3Target = 0.
w3cal = 0.
w2actual = 0.
w3actual = 0.
outKPIall = 0.
step = 0.
aa = 1.
bb = 1.
Customer = ''
ProjectName = ''
Drawing = ''
ValidDate = ''
Name = ''
Comment = ''
ParkPosition = ''
w2 = 0.
w3 = 0.
errorP = ''
errorN = ''
isClock = 1
isStandard = ''
Measured = ''
zeroAngle = 0.
APS1 = 0.
APS2 = 0
IPL = 0.
UWL = 0.
ALP = 0.
SP = 0.
KMP = 0.
Offset = 0.
ParkWhere = ''
DriveType = ''
MechanicType = ''
Ra_Rb = 0.
Ra_Rb2 = 0.
KBEW = '-x'
KBEW2 = '-x'
startFrom = ''
A_X = 0.
A_Y = 0.
A_Y = 0.
A_Z = 0.
B_X = 0.
B_Y = 0.
B_Z = 0.
A_X2 = 0.
A_Y2 = 0.
A_Z2 = 0.
B_X2 = 0.
B_Y2 = 0.
B_Z2 = 0.
E_X = 0.
E_Y = 0.
E_Z = 0.
F_X = 0.
F_Y = 0.
F_Z = 0.
E_X2 = 0.
E_Y2 = 0.
E_Z2 = 0.
F_X2 = 0.
F_Y2 = 0.
F_Z2 = 0.
Ap_X = 0.
Ap_Y = 0.
Ap_Z = 0.
Ap_X2 = 0.
Ap_Y2 = 0.
Ap_Z2 = 0.
Fp_X = 0.
Fp_Y = 0.
Fp_Z = 0.
Fp_X2 = 0.
Fp_Y2 = 0.
Fp_Z2 = 0.

line_ani = ''
listWrite = []
template_path = os.getcwd()+'/template.xlsx'  # read
template_design1 = os.getcwd()+'/template_DesignTable1.xlsx'  # re.ad
template_design2 = os.getcwd()+'/template_DesignTable2.xlsx'  # read
template_design1_Series = os.getcwd()+'/template_DesignTable1_Series.xlsx'  # re.ad
excel_out = ''
excel_design1 = ''
excel_design2 = ''
listWrite2 = []
wb = xl.load_workbook(template_path)
sheet1 = wb['Input']
sheet2 = wb['Optimization']
sheet3 = wb['Input']
sheet4 = wb['Output-Alphanumeric']
sheet5 = wb['Output-Alphanumeric (2)']
sheet6 = wb['Output-Tolerance calcuation']

wb1 = ''
sheetDesign1 = ''
wb2 = ''
sheetDesign2 = ''
wb1Series = ''
sheetDesign1Series = ''
ordinatePath = 'outputCordinate.xlsx'
sheetNew = ''
pause = False
offsetAngle = 0.
Master = 'Driver Side'
cwd = os.getcwd()
DesignPath = {'Center': [cwd+'\LinkLines\Center\Center_LinkLines_DesignTable1.xlsx', cwd+'\LinkLines\Center\Center_LinkLines_DesignTable2.xlsx'],
              'Serial DS-PS': [cwd+'\LinkLines\Serial DS-PS\Serial_DS-PS_Linklines_DesignTable1.xlsx', cwd+'\LinkLines\Serial DS-PS\Serial_DS-PS_Linklines_DesignTable2.xlsx'],
              'Serial PS-DS': [cwd+'\LinkLines\Serial PS-DS\Serial_PS-DS_Linklines_DesignTable1.xlsx', cwd+'\LinkLines\Serial PS-DS\Serial_PS-DS_Linklines_DesignTable2.xlsx']}

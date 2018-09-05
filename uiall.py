# -*- coding: utf-8 -*-
"""
Created on Mon Jul 23 19:37:08 2018
test test
@author: Administrator
""" 
import logging
import math
import os
import os.path
import sys
import time
import uuid
#import hashlib
import time
import base64
import os
import sys


import matplotlib.animation as animation
import matplotlib.pyplot as plt
import mpl_toolkits.mplot3d.axes3d as p3
import numpy as np
import openpyxl as xl
import pandas as pd
import sympy as sp
import scipy.optimize as op
from mpl_toolkits.mplot3d import proj3d
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from PyQt5.QtWidgets import QTableWidgetItem as Qitem
import ui 
import Func
import parameters
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QBrush, QColor
import pyomo.environ as pe

from PyQt5.QtWidgets import QApplication, QMainWindow, QMenu, QVBoxLayout, QSizePolicy, QMessageBox, QWidget,QTabWidget

import Fig

class MainWindow(QTabWidget, ui.Ui_MainWindow):
    def __init__(self,parent=None):
        super(MainWindow,self).__init__(parent)
        self.setWindowTitle("Kinematics  Calculation of wiper linkage")
        self.resize(1000, 1000)
        
        # window_pale = QtGui.QPalette()
        # window_pale.setColor(self.backgroundRole(),QColor(0,0,255))
        # self.setPalette(window_pale)
        # self.setAutoFillBackground(True)
        #self.setupUi(self)
        self.WindowInput = MainWindow_Input()
        self.WindowOpt = MainWindow_Opt()
        self.WindowAlpha = MainWindow_alphaNumeric()
        self.WindowExtreme = MainWindow_extreme()
        self.WindowTolerance = MainWindow_Tolerance()
        self.addTab(self.WindowInput,u"input")
        self.addTab(self.WindowOpt,u"optimization")
        self.addTab(self.WindowAlpha,u"Output-AlphaNumeric")
        self.addTab(self.WindowExtreme,u"Output-ExtremeValues")
        self.addTab(self.WindowTolerance,u"Tolerance Calculation")

        self.WindowInput.LoadButton.clicked.connect(self.WindowInput.LoadClicked) # laod data
        self.WindowInput.LoadButton.clicked.connect(self.WindowOpt.OptClicked) # optimization
        self.WindowInput.LoadButton.clicked.connect(self.WindowInput.UpdateOptValue) # set items in input sheet
        self.WindowInput.LoadButton.clicked.connect(self.WindowOpt.setOptItem) # set items in optimization sheet

        self.WindowInput.OutputButton.clicked.connect(self.WindowInput.close)
        self.WindowInput.OutputButton.clicked.connect(self.WindowAlpha.show)
        self.WindowInput.OutputButton.clicked.connect(self.resetAlpha)       # move to output sheet
        self.WindowOpt.pushButton.clicked.connect(self.WindowOpt.OptClicked) # opt
        self.WindowOpt.pushButton.clicked.connect(self.WindowOpt.setOptItem) # set itmes in optimization sheet

        self.WindowAlpha.TableButton.clicked.connect(self.WindowAlpha.alphaTable)  # do alphanumeric calculation
        #self.WindowAlpha.plotButton.clicked.connect(lambda: self.WindowAlpha.plot(gs.outKPIall))  # do alphanumeric calculation
        self.WindowAlpha.plotButton.clicked.connect(self.WindowAlpha.plot)  # do alphanumeric calculation
        #self.WindowAlpha.plotButton.clicked.connect(figs.show)  # do alphanumeric calculation
       # self.WindowAlpha.AnimationButton.clicked.connect(self.WindowAlpha.animate)  # do alphanumeric calculation
        
        self.WindowAlpha.ExtremeButton.clicked.connect(self.WindowExtreme.ExtremeTable) # do extremetable calculation
        self.WindowAlpha.ExtremeButton.clicked.connect(self.WindowAlpha.close)          
        self.WindowAlpha.ExtremeButton.clicked.connect(self.WindowExtreme.show)         # move to extreme tab
        self.WindowAlpha.ExtremeButton.clicked.connect(self.reset2)
        
        self.WindowInput.ToleranceButton.clicked.connect(self.WindowInput.close)
        self.WindowInput.ToleranceButton.clicked.connect(self.WindowTolerance.show)      # move to tolerance tab
        self.WindowInput.ToleranceButton.clicked.connect(self.resetTolerance)
        
        self.WindowTolerance.ButtonCal.clicked.connect(self.WindowTolerance.Tolerance)   # do tolerance calculation

        self.WindowInput.SaveButton.clicked.connect(self.writeMainCordinate)             # write cordinate
        self.WindowInput.SaveButton.clicked.connect(self.writeDetailCordinate)
        self.WindowInput.SaveButton.clicked.connect(self.WindowInput.updateNewsheet)
        self.WindowInput.SaveButton.clicked.connect(self.saveExcel)                      # save excel
        self.WindowInput.SaveButton.clicked.connect(self.writejson)                      # save excel

    def writejson(self):
        outputjson ='output.json'
        Func.writejson(outputjson,outputs,gs)
    def saveExcel(self):
        gs.wb.save(gs.excel_out)
        print('save completed')
    def writeMainCordinate(self):
        outM1=Func.Output(gs.BC,gs.CD,gs.ED,gs.xm1,gs.A,gs.B,gs.E,gs.F,KBEW=gs.KBEW)  #Master -UWL,[alpha,beta,NYS_T,NYS_A,NYK_T,NYK_A,C[0][0],C[0][1],C[0][2],Db[0][0],Db[0][1],Db[0][2]]
        outM2=Func.Output(gs.BC,gs.CD,gs.ED,gs.xm2,gs.A,gs.B,gs.E,gs.F,KBEW=gs.KBEW) #Master-OWL [alpha,beta,NYS_T,NYS_A,NYK_T,NYK_A,C[0][0],C[0][1],C[0][2],Db[0][0],Db[0][1],Db[0][2]]
        outS3=Func.Output(gs.BC2,gs.CD2,gs.ED2,gs.xm12,gs.A2,gs.B2,gs.E2,gs.F2,KBEW=gs.KBEW2) #slave UWL
        outS4=Func.Output(gs.BC2,gs.CD2,gs.ED2,gs.xm22,gs.A2,gs.B2,gs.E2,gs.F2,KBEW=gs.KBEW2) #slave OWL
        
        if gs.MechanicType =='Center':
            outM3=Func.Output(gs.BC,gs.CD,gs.ED,gs.xm12,gs.A,gs.B,gs.E,gs.F,KBEW=gs.KBEW)  #Slave -UWL, 
            outM4=Func.Output(gs.BC,gs.CD,gs.ED,gs.xm22,gs.A,gs.B,gs.E,gs.F,KBEW=gs.KBEW)  #Slave -OWL, 
            outS1=Func.Output(gs.BC2,gs.CD2,gs.ED2,gs.xm1,gs.A2,gs.B2,gs.E2,gs.F2,KBEW=gs.KBEW2)  #Master -UWL,
            outS2=Func.Output(gs.BC2,gs.CD2,gs.ED2,gs.xm2,gs.A2,gs.B2,gs.E2,gs.F2,KBEW=gs.KBEW2)  #Master -OWL,
            if gs.Master =='driver side':
                array =  [gs.A,gs.Ap,gs.B,gs.B2 , gs.F, gs.Fp, gs.E, gs.F2,gs.Fp2,gs.E2,outM1[6:12],outS1[6:12],
                outM1[6:12],outS1[6:12],outM2[6:12],outS2[6:12],outM3[6:12],outS3[6:12],outM4[6:12],outS4[6:12]]
            else :
                array =  [gs.A,gs.Ap,gs.B2,gs.B , gs.F2, gs.Fp2, gs.E2, gs.F,gs.Fp,gs.E,outS1[6:12],outM1[6:12],
                outS1[6:12],outM1[6:12],outS2[6:12],outM2[6:12],outS3[6:12],outM3[6:12],outS4[6:12],outM4[6:12]]
            gs.CordinateList =[y for x in array for y in x]
            num = len(gs.CordinateList)
    
            for i in range(num):
                Func.write(gs.sheetDesign1,i+1,2,'%.4f'%gs.CordinateList[i])
            gs.wb1.save(filename=gs.excel_design1)
        
        else:
            if gs.Master =='driver side':
                array = [gs.A,gs.Ap,gs.F,gs.Fp,gs.F2,gs.Fp2 ,gs.B, gs.E, gs.B2, gs.E2,outM1[6:12],outS3[6:12],
                outM1[6:12],outS3[6:12],outM2[6:12],outS4[6:12]]
            else:

                array = [gs.A,gs.Ap,gs.F2,gs.Fp2,gs.F,gs.Fp ,gs.B2, gs.E2, gs.B, gs.E,outS3[6:12],outM1[6:12],
                outS3[6:12],outM1[6:12],outS4[6:12],outM2[6:12]]
            gs.CordinateList =[y for x in array for y in x]
            num = len(gs.CordinateList)
    
            for i in range(num):
                Func.write(gs.sheetDesign1,i+1,2,'%.4f'%gs.CordinateList[i])
            gs.wb1.save(filename=gs.excel_design1)

    def writeDetailCordinate (self):
            alphaList = np.linspace(gs.xm1,gs.xm1+360,12,endpoint =False)
            for alpha in alphaList:
                if gs.MechanicType =='Center' :
                    outM=   Func.Output(gs.BC,gs.CD,gs.ED,alpha,gs.A,gs.B,gs.E,gs.F,KBEW=gs.KBEW)  #
                    outS =  Func.Output(gs.BC2,gs.CD2,gs.ED2,alpha,gs.A2,gs.B2,gs.E2,gs.F2,KBEW=gs.KBEW2)  #
                else:
                    outM=   Func.Output(gs.BC, gs.CD,  gs.ED, alpha, gs.A, gs.B , gs.E,  gs.F, KBEW=gs.KBEW)  #
                    alpha2 =outM[1]+gs.Delta2 
                    outS =  Func.Output(gs.BC2,gs.CD2, gs.ED2,alpha2,gs.A2,gs.B2, gs.E2, gs.F2, KBEW=gs.KBEW2)  #
                if gs.Master =='driver side':                     
                    gs.CordinateDetailList.extend(outM[6:12])
                    gs.CordinateDetailList.extend(outS[6:12])
                else:
                    gs.CordinateDetailList.extend(outS[6:12])
                    gs.CordinateDetailList.extend(outM[6:12])
            num = len(gs.CordinateDetailList)
            for i in range(num):
                Func.write(gs.sheetDesign2,i+1,2,'%.4f'%gs.CordinateDetailList[i])

            gs.wb2.save(filename=gs.excel_design2)
                    
                    
    def resetTolerance(self):
        self.addTab(self.WindowTolerance, u"Tolerance Calculation")
        self.addTab(self.WindowInput,u"input")
        self.addTab(self.WindowOpt,u"optimization")
        self.addTab(self.WindowAlpha,u"AlphaNumeric")
        self.addTab(self.WindowExtreme,u"Extreme values")
        
    def resetAlpha(self):
        self.addTab(self.WindowAlpha,u"AlphaNumeric")
        self.addTab(self.WindowExtreme,u"Extreme values")
        self.addTab(self.WindowTolerance, u"Tolerance Calculation")
        self.addTab(self.WindowInput,u"input")
        self.addTab(self.WindowOpt,u"optimization")
    def reset(self):
        self.addTab(self.WindowOpt,u"optimization")
        self.addTab(self.WindowAlpha,u"AlphaNumeric")
        self.addTab(self.WindowExtreme,u"Extreme values")
        self.addTab(self.WindowTolerance, u"Tolerance Calculation")
        self.addTab(self.WindowInput,u"input")

    def reset2(self):
        self.addTab(self.WindowExtreme,u"Extreme values")
        self.addTab(self.WindowTolerance, u"Tolerance Calculation")
        self.addTab(self.WindowInput,u"input")
        self.addTab(self.WindowOpt,u"optimization")
        self.addTab(self.WindowAlpha,u"AlphaNumeric")
# =============================================================================
class MainWindow_Tolerance(QMainWindow,ui.Ui_Tolerance):
    def __init__(self,parent=None):
        super(MainWindow_Tolerance,self).__init__(parent)
        self.setupUi(self)
        
    def Tolerance(self):
        gs.noCrank = self.comboLink.currentText() #Master,Slave
        obj = self.comboObj.currentText() #wipping angle,NYS_T

        listToleranceStrall = ["BC", "ED", 'Delta','CD',"F_X", "F_Y", "F_Z","Fp_X", "Fp_Y", "Fp_Z",'FE',"A_X", "A_Y", "A_Z","Ap_X", "Ap_Y", "Ap_Z",'Distance',
                             "BC2", "ED2", 'Delta2','CD2',"F_X2", "F_Y2", "F_Z2", "Fp_X2", "Fp_Y2", "Fp_Z2",'FE2',"A_X2", "A_Y2", "A_Z2","Ap_X2", "Ap_Y2", "Ap_Z2",'Distance2'] #TBD:DElta,Distance
        numTolerance = int(len(listToleranceStrall)/2)
        if  obj =='outputAngle':
                t = 0
                gs.index = 1 # to write out
        elif obj =='tangent force angle':
                t = 1
                gs.index =4 # to write out
        else:
                print('please configure objective function')
                
        if gs.noCrank=='Master':
                listToleranceStr = listToleranceStrall[0:numTolerance]
                Base = [eval('gs.'+st) for st in listToleranceStr]
                Target = Func.Tolerance(Base,gs.xm1,gs.xm2)[t]
                gs.no = 2
        elif gs.noCrank == 'Slave':
                listToleranceStr = listToleranceStrall[numTolerance:]
                Base = [eval('gs.'+st) for st in listToleranceStr]
                Target = Func.Tolerance(Base,gs.xm12,gs.xm22)[t]
                gs.no=3
        else:
                print('please configure noCrank')

        errorPList = [ 'gs.errorP.'+t for t in listToleranceStr]
        errorPositive = [eval(t) for t in errorPList]
        errorNList = [ 'gs.errorN.'+t for t in listToleranceStr]
        errorNegative = [eval(t) for t in errorNList]

        
        
        ToleranceValue = Base.copy()  # to be changed
        if gs.noCrank=='Master': #master
                for i in range(numTolerance):
                        Func.updateTolerance(ToleranceValue, gs.w2ErrorList, gs.w2List, gs.kpiList, i,  t,Target, errorPositive[i],
                                errorNegative[i], gs.xm1, gs.xm2)
                arrayi = ToleranceValue.copy()
                print(ToleranceValue)
                for i in range(numTolerance):
                        arrayi[i] = Base[i]
                        Func.updateTolerance(arrayi, gs.w2ErrorList, gs.w2List, gs.kpiList, i, t,Target, errorPositive[i],
                                errorNegative[i], gs.xm1, gs.xm2)
        elif gs.noCrank == 'Slave':
                for i in range(numTolerance):
                        Func.updateTolerance(ToleranceValue, gs.w2ErrorList, gs.w2List, gs.kpiList, i, t,Target,
                                             errorPositive[i],
                                             errorNegative[i], gs.xm12, gs.xm22)
                
                arrayi = ToleranceValue.copy()
                for i in range(numTolerance):
                        arrayi[i] = Base[i]
                        Func.updateTolerance(arrayi,gs.w2ErrorList, gs.w2List, gs.kpiList, i, t,Target,
                                             errorPositive[i],
                                             errorNegative[i], gs.xm12, gs.xm22)
        gs.Parameter = np.tile(listToleranceStr, 2)
        #gs.w2ErrorList=np.array(gs.w2ErrorList).cumsum()
        # start wrtie to ui/excel
        self.tableTolerance.setRowCount(2*numTolerance)
        for i in range(2*numTolerance):
                self.tableTolerance.setItem(i, 0, Qitem(str(gs.no)))
                self.tableTolerance.setItem(i, 1, Qitem(gs.Parameter[i]))
                self.tableTolerance.setItem(i, 2, Qitem('%.4f'%gs.kpiList[i]))
                self.tableTolerance.setItem(i, 3, Qitem('%.4f'%gs.w2List[i]))
                self.tableTolerance.setItem(i, 4, Qitem('%.4f'%gs.w2ErrorList[i]))
        Func.write(gs.sheet6,3,4,gs.no)
        Func.write(gs.sheet6,3,6,gs.index)
        startColumn = 2
        startRow = 10
        for i in range(2*numTolerance):
                Func.write(gs.sheet6, startRow + i, startColumn, gs.no)  # TBD:N2
                Func.write(gs.sheet6, startRow + i, startColumn + 1, gs.Parameter[i])  # parameter
                Func.write(gs.sheet6, startRow + i, startColumn + 2, '%.4f'%gs.kpiList[i])
                Func.write(gs.sheet6, startRow + i, startColumn + 3, '%.4f'%gs.w2List[i])
                Func.write(gs.sheet6, startRow + i, startColumn + 4, '%.4f'%gs.w2ErrorList[i])

        print('Tolerance calculation completed')

class MainWindow_extreme(QMainWindow,ui.Ui_extreme):
    def __init__(self,parent=None):
        super(MainWindow_extreme,self).__init__(parent)
        self.setupUi(self)
    def ExtremeTable(self):
        if gs.MechanicType =='Center':
                gs.UWL1 = gs.xm2+90
                gs.UWL2 = gs.xm2+90
        else:
                gs.UWL1 = gs.xm2 + 90
                gs.UWL2 = gs.xm22 + 90 #TBD
        if gs.DriveType =='Standard':
                gs.Park = gs.xm1+90
        else:
                gs.Park = gs.xm1+90 #TBD
        self.tableExtreme1.setItem(0, 2, Qitem('%.2f' % gs.Park))
        self.tableExtreme1.setItem(0, 3, Qitem('%.2f' % gs.UWL1))
        self.tableExtreme1_2.setItem(0, 2, Qitem('%.2f' % gs.Park))
        self.tableExtreme1_2.setItem(0, 3, Qitem('%.2f' % gs.UWL2))

        self.tableExtreme1.setItem(0,0,Qitem('%.2f'%gs.w2Target))
        self.tableExtreme1.setItem(0,1,Qitem('%.2f'%gs.w2cal))
        self.tableExtreme1_2.setItem(0,0,Qitem('%.2f'%gs.w3Target))
        self.tableExtreme1_2.setItem(0,1,Qitem('%.2f'%gs.w3cal))
        for i in range(10):
            self.tableExtreme2.setItem(0,i,Qitem('%.2f'%gs.maxArray[i]))
            self.tableExtreme2.setItem(1,i,Qitem('%.2f'%gs.minArray[i]))
            self.tableExtreme2_2.setItem(0,i,Qitem('%.2f'%gs.maxArray2[i]))
            self.tableExtreme2_2.setItem(1,i,Qitem('%.2f'%gs.minArray2[i]))
        for i in range(4):
            self.tableExtreme3.setItem(0,6+i,Qitem('%.2f'%gs.maxArray[10+i]))
            self.tableExtreme3.setItem(1,6+i,Qitem('%.2f'%gs.minArray[10+i]))
            self.tableExtreme3_2.setItem(0,6+i,Qitem('%.2f'%gs.maxArray2[10+i]))
            self.tableExtreme3_2.setItem(1,6+i,Qitem('%.2f'%gs.minArray2[10+i]))
        startRow = [10, 15, 27, 32]
        startCol = [3, 9, 3, 9]
        Func.write(gs.sheet5, 3, 6, gs.w2Target)
        Func.write(gs.sheet5, 4, 6, gs.w2cal)
        Func.write(gs.sheet5, 20, 6, gs.w3Target)
        Func.write(gs.sheet5, 21, 6, gs.w3cal)

        Func.write(gs.sheet5, 3, 11, '%.2f' % gs.Park)
        Func.write(gs.sheet5, 4, 11, '%.2f' % gs.UWL1)
        Func.write(gs.sheet5, 20, 11, '%.2f' % gs.Park)
        Func.write(gs.sheet5, 21, 11, '%.2f' % gs.UWL2)

        for i in range(10):
                Func.write(gs.sheet5, startRow[0], startCol[0] + i, gs.maxArray[i])
                Func.write(gs.sheet5, startRow[0] + 1, startCol[0] + i, gs.minArray[i])
                Func.write(gs.sheet5, startRow[2], startCol[2] + i, gs.maxArray2[i])
                Func.write(gs.sheet5, startRow[2] + 1, startCol[2] + i, gs.minArray2[i])
        for i in range(4):
                Func.write(gs.sheet5, startRow[1], startCol[1] + i, gs.maxArray[10 + i])
                Func.write(gs.sheet5, startRow[1] + 1, startCol[1] + i, gs.minArray[10 + i])
                Func.write(gs.sheet5, startRow[3], startCol[3] + i, gs.maxArray2[10 + i])
                Func.write(gs.sheet5, startRow[3] + 1, startCol[3] + i, gs.minArray2[10 + i])
        # write alphanumeric


        print('extreme values complete')

class MainWindow_alphaNumeric(QMainWindow,ui.Ui_alphaNumeric):
    def __init__(self,parent=None):
        super(MainWindow_alphaNumeric,self).__init__(parent)
        self.setupUi(self)
    @staticmethod
    def animate():
        number=gs.num

        Atemp=np.array([gs.A]*number).T
        Btemp=np.array([gs.B]*number).T
        Etemp=np.array([gs.E]*number).T
        Ftemp=np.array([gs.F]*number).T
        A2temp=np.array([gs.A2]*number).T
        B2temp=np.array([gs.B2]*number).T
        E2temp=np.array([gs.E2]*number).T
        F2temp=np.array([gs.F2]*number).T
        Ctemp= np.array((gs.outCordinateall['Cx'],gs.outCordinateall['Cy'],gs.outCordinateall['Cz']))
        Dtemp= np.array((gs.outCordinateall['Dx'],gs.outCordinateall['Dy'],gs.outCordinateall['Dz']))
        C2temp= np.array((gs.outCordinateall['Cx2'],gs.outCordinateall['Cy2'],gs.outCordinateall['Cz2']))
        D2temp= np.array((gs.outCordinateall['Dx2'],gs.outCordinateall['Dy2'],gs.outCordinateall['Dz2']))
        dataAB=np.r_[Atemp,Btemp]
        dataAB2=np.r_[A2temp,B2temp]
        dataEF=np.r_[Etemp,Ftemp]
        dataEF2=np.r_[E2temp,F2temp]
        dataCD =np.r_[Ctemp,Dtemp]
        dataCD2=np.r_[C2temp,D2temp]
        dataBC=np.r_[Btemp,Ctemp]
        dataDE=np.r_[Dtemp,Etemp]
        dataBC2=np.r_[B2temp,C2temp]
        dataDE2=np.r_[D2temp,E2temp]
        dataCD2=np.r_[C2temp,D2temp]
        data=np.ones((10,6,number))
        data[0]=dataAB
        data[1]=dataBC
        data[2]=dataCD
        data[3]=dataDE
        data[4]=dataEF
        data[5]=dataAB2
        data[6]=dataBC2
        data[7]=dataCD2
        data[8]=dataDE2
        data[9]=dataEF2
        #TBD
        # initial point
        def orthogonal_proj(zfront, zback):
            a = (zfront+zback)/(zfront-zback)
            b = -2*(zfront*zback)/(zfront-zback)
            return np.array([[1,0,0,0],
                                [0,1,0,0],
                                [0,0,a,b],
                                [0,0, -0.0001,zback]])
        proj3d.persp_transformation = orthogonal_proj

        fig2 = plt.figure()
        ax2 = p3.Axes3D(fig2)
        ax2.set_xlabel('x')
        ax2.set_ylabel('y')
        ax2.set_zlabel('z')
        Label=['AB','BC','CD','DE','EF','AB2','BC2','CD2','DE2','EF2']
        colors ='bgcmkbgcmk'
        linestyles=['-','-','-','-','-','--','--','--','--','--']

        cordinateMin = np.max(data[0],axis=1)
        cordinateMax = np.max(data[0],axis=1)
        #gs.data =data
        for item in data:
            cordinateMaxNew = np.max(item,axis=1)
            cordinateMinNew = np.min(item,axis=1)
            for i in range(6):
                if cordinateMaxNew[i] > cordinateMax[i]:
                    cordinateMax[i] = cordinateMaxNew[i]
                if cordinateMinNew[i] < cordinateMin[i]:
                    cordinateMin[i] = cordinateMinNew[i]
        xmin = min (cordinateMin[0],cordinateMin[3])
        ymin = min (cordinateMin[1],cordinateMin[4])
        zmin = min (cordinateMin[2],cordinateMin[5])
        xmax = max (cordinateMax[0],cordinateMax[3])
        ymax = max (cordinateMax[1],cordinateMax[4])
        zmax = max (cordinateMax[2],cordinateMax[5])
        
        xxrange = xmax-xmin
        yrange = ymax-ymin
        zrange = zmax-zmin


        ab=gs.B-gs.A
        lines = [ax2.plot([data[i][0,0],data[i][3,0]],[data[i][1,0],data[i][4,0]],[data[i][2,0],data[i][5,0]],label=Label[i],color=colors[i],linestyle=linestyles[i])[0] for i  in range(10)]
        gs.ab = [ab[0]/xxrange,ab[1]/yrange,ab[2]/zrange] #scale axis equal
        az=-90-Func.degree(math.atan(gs.ab[0]/gs.ab[1]))
        el=-np.sign(ab[1])*Func.degree(math.atan(gs.ab[2]/math.sqrt(gs.ab[0]**2+gs.ab[1]**2)))
        ax2.set_xlim3d(xmin,xmax)
        ax2.set_ylim3d(ymin,ymax)
        ax2.set_zlim3d(zmin,zmax)
        
        ax2.set_title('initial position')
        #ax2.axis('scaled')
        ax2.view_init(azim=az, elev=el)
        
        # Attaching 3D axis to the figure
        gs.pause =False
        def onClick(event):
            gs.pause ^= True
        def update_lines ( num, dataLines, lines, ax):
            if not gs.pause:
                for line, data in zip(lines, dataLines):
                # NOTE: there is no .set_data() for 3 dim data.
                    Cx = data[0, num]
                    Cy = data[1, num]
                    Cz = data[2, num]
                    Dx = data[3, num]
                    Dy = data[4, num]
                    Dz = data[5, num]
                    temp = [[Cx, Dx], [Cy, Dy]]
    
                    line.set_data(temp)
                    line.set_3d_properties([Cz, Dz])
            return lines
        fig = plt.figure()
        ax = p3.Axes3D(fig)
        # =============================================================================
        ax.set_xlabel('x')
        ax.set_ylabel('y')
        ax.set_zlabel('z')
        ax.set_title('kinematics')
        #ax.set_xbound(0.9*xmin,1.1*xmax)
        #ax.set_ybound(0.9*ymin,1.1*ymax)
        #ax.set_zbound(0.9*zmin,1.1*zmax)
        # =============================================================================
        #ax.axis('equal')
        ax.view_init(azim=az,elev=el)
        
        # =============================================================================
        txt=['A','B','C','D','E','F','A2','B2','C2','D2','E2','F2']
        # =============================================================================
        # for i in range(10):
        #     ax.text(data[i][0][0],data[i][1][0],data[i][2][0],txt[i])
        # =============================================================================
        # Creating the Animation object
        Label=['AB','BC','CD','DE','EF','AB2','BC2','CD2','DE2','EF2']
        lines = [ax.plot([data[i][0,0],data[i][3,0]],[data[i][1,0],data[i][4,0]],[data[i][2,0],data[i][5,0]],label=Label[i],color=colors[i],linestyle=linestyles[i])[0] for i  in range(10)]
        #start of each line
        #fig.canvas.mpl_connect('button_press_event', onClick)
        gs.line_ani=animation.FuncAnimation(fig, update_lines, number, fargs=(data, lines,ax),interval=int(3600/gs.num), blit=True)
        
        #==============================================================================
        # plt.rcParams['animation.ffmpeg_path']='G:\\wai\\ffmpeg\\bin\\ffmpeg.exe'
        #==============================================================================
        plt.legend()
        plt.show()
        #return gs.line_ani  
    def plot(self):
        self.figs = Fig.ApplicationWindow(gs)
        self.figs.show( )

    
    def getextreme(self):
        n = len(gs.outKPIall.columns) 
        
        assert (n%2==1)
        n1 = int((n+1)/2) # master part
        for i in range(1,n1):
            listT = Func.getMax(gs.outKPIall,i)
            gs.maxArray.append(listT[0])
            gs.maxArray.append(listT[1])
            listT = Func.getMin(gs.outKPIall,i)
            gs.minArray.append(listT[0])
            gs.minArray.append(listT[1])

        for i in range(n1,n):
            listT = Func.getMax(gs.outKPIall,i)
            gs.maxArray2.append(listT[0])
            gs.maxArray2.append(listT[1])
            listT = Func.getMin(gs.outKPIall,i)
            gs.minArray2.append(listT[0])
            gs.minArray2.append(listT[1])
            

        gs.w2cal=gs.maxArray[1]-gs.minArray[1]
        gs.w3cal=gs.maxArray2[1]-gs.minArray2[1]
        print('w2=%.2f'%gs.w2cal+'\tw3=%.2f'%gs.w3cal)
        
        return [gs.maxArray,gs.minArray,gs.maxArray2,gs.minArray2]
        #outCordinate.to_excel('outputCordinate.xlsx')
        
        #Range=[1.2,50,8,90,8]# 90 TBD
        #write to excel
   
    def getoutall(self): # alpha calculation ,return outkpi,outcordinate dataframe
            print('-----------------------start map calculation-------------------------')
            output=[]
            output2=[]
            # =============================================================================
            alpha0=gs.xm1
            gs.step=int(self.Textstep.text())
            if gs.DriveType=='Standard':    
                gs.num =int(abs(360/gs.step))
            else:
                gs.num=math.ceil(abs(gs.Alfa/gs.step))+1
            alphaList = [alpha0 +gs.step * x  for x in range(gs.num-1)]
            alphaList.append(alpha0+gs.Alfa)
            for alpha in alphaList:
                if   gs.MechanicType =='Center':
                    outM=   Func.Output(gs.BC,gs.CD,gs.ED,alpha,gs.A,gs.B,gs.E,gs.F,KBEW= gs.KBEW)  #
                    outS =  Func.Output(gs.BC2,gs.CD2,gs.ED2,alpha,gs.A2,gs.B2,gs.E2,gs.F2,KBEW = gs.KBEW2)  #
                else:
                    outM=   Func.Output(gs.BC,gs.CD,gs.ED,alpha,gs.A,gs.B,gs.E,gs.F,KBEW = gs.KBEW)  #
                    alpha2 =outM[1]+gs.Delta2
                    outS =  Func.Output(gs.BC2,gs.CD2,gs.ED2,alpha2,gs.A2,gs.B2,gs.E2,gs.F2,KBEW = gs.KBEW2)  #
                output.append(outM)
                output2.append(outS)

            beta_sList=[]
            beta_ssList=[]
            beta_s2List=[]
            beta_ss2List=[]

            out = pd.DataFrame(output)
            out2= pd.DataFrame(output2)
            out.columns =['alpha','beta','NYS_T','NYS_A','NYK_T','NYK_A','Cx','Cy','Cz','Dx','Dy','Dz']
            out2.columns=['alpha2','beta2','NYS_T2','NYS_A2','NYK_T2','NYK_A2','Cx2','Cy2','Cz2','Dx2','Dy2','Dz2']

            series1=pd.Series([gs.zeroAngle]*gs.num)
            series2=pd.Series([out['beta'][0]]*gs.num)
            series3=pd.Series([out2['alpha2'][0]]*gs.num)
            series4=pd.Series([out2['beta2'][0]]*gs.num)
            out['alpha']=out['alpha'].sub(series1,axis=0)
            out['beta'] = out['beta'].sub(series2,axis=0)
            out2['alpha2'] = out2['alpha2'].sub(series3,axis=0)
            out2['beta2'] = out2['beta2'].sub(series4,axis=0)

            for i in range(gs.num-1):
                beta_s =(out['beta'][i+1]-out['beta'][i])/(alphaList[i+1]-alphaList[i])
                beta_sList.append(beta_s)
                beta_s2 =(out2['beta2'][i+1]-out2['beta2'][i])/(alphaList[i+1]-alphaList[i])
                beta_s2List.append(beta_s2)
            beta_sList.append(0) # fill with 0
            beta_s2List.append(0) # fill with 0
            for i in range(gs.num-1):
                beta_ss =(beta_sList[i+1]-beta_sList[i])/(alphaList[i+1]-alphaList[i])
                beta_ssList.append(beta_ss)
                beta_ss2 =(beta_s2List[i+1]-beta_s2List[i])/(alphaList[i+1]-alphaList[i])
                beta_ss2List.append(beta_ss2)
            beta_ssList.append(0) # fill with 0
            beta_ss2List.append(0) # fill with 0
            out['beta_s']=beta_sList
            out['beta_ss']=beta_ssList
            out2['beta_s2']=beta_s2List
            out2['beta_ss2']=beta_ss2List
            
            #outKPIall=pd.concat([outKPI,outKPI2],axis=1)
            outall = pd.concat([out,out2],axis=1)
            outall = outall.astype('float64')
            cols =['alpha','beta','beta_s','beta_ss','NYS_T','NYS_A','NYK_T','NYK_A']
            cols2=['beta2','beta_s2','beta_ss2','NYS_T2','NYS_A2','NYK_T2','NYK_A2']
            colskpi = cols+cols2
            colsCordinate=['Cx','Cy','Cz','Dx','Dy','Dz']
            colsCordinate2=['Cx2','Cy2','Cz2','Dx2','Dy2','Dz2']
            colsCordinateall = colsCordinate + colsCordinate2
            # =============================================================================
            gs.outKPIall = outall.loc[:,colskpi]
            gs.outCordinateall = outall.loc[:,colsCordinateall]
           
            # outKPI['NYS_T']=outKPI['NYS_T'].astype('float64')
            # outKPI['NYS_A']=outKPI['NYS_A'].astype('float64')
            # outKPI['NYK_T']=outKPI['NYK_T'].astype('float64')
            # outKPI['NYK_A']=outKPI['NYK_A'].astype('float64')
            # outKPI2['NYS_T2']=outKPI2['NYS_T2'].astype('float64')
            # outKPI2['NYS_A2']=outKPI2['NYS_A2'].astype('float64')
            # outKPI2['NYK_T2']=outKPI2['NYK_T2'].astype('float64')
            # outKPI2['NYK_A2']=outKPI2['NYK_A2'].astype('float64')
            return [gs.outKPIall, gs.outCordinateall]
    
    def writealphaNumeric(self):
        startRow = 10
        startCol = 1
        self.tableAlfa.setRowCount(gs.num)
        for i in range(gs.num):
            for j in range(gs.outKPIall.shape[1]):
                self.tableAlfa.setItem(i,j,Qitem('%.2f'%gs.outKPIall.iloc[i,j]))
                Func.write(gs.sheet4,startRow+i , startCol+j , gs.outKPIall.iloc[i,j])
        Func.write(gs.sheet4,3,6,gs.step)
    def alphaTable(self):
        [gs.outKPIall, gs.outCordinateall] = self.getoutall()
        self.writealphaNumeric()
        self.getextreme()
        #self.plot(gs.outKPIall)

    #animation

    
class MainWindow_Opt(QMainWindow,ui.Ui_Optimization):
    def __init__(self,parent=None):
        super(MainWindow_Opt,self).__init__(parent)
        self.setupUi(self)

    def setOptItem(self):
                gs.listOpt1 = [[gs.EDb,gs.CDb,gs.xs1b,gs.xs2b,gs.w2optb,gs.M1NYS_Tb,gs.M2NYS_Tb,gs.M1NYS_Tb+gs.M2NYS_Tb]]
                gs.listOpt1.append([gs.ED,gs.CD,gs.xs1,gs.xs2,gs.w2opt,gs.M1NYS_T,gs.M2NYS_T,gs.M1NYS_T+gs.M2NYS_T])

                gs.listOpt2 = [[gs.ED2b,gs.CD2b,gs.xs12b,gs.xs22b,gs.w3optb,gs.S1NYS_Tb,gs.S2NYS_Tb,gs.S1NYS_Tb+gs.S2NYS_Tb]]
                gs.listOpt2.append([gs.ED2,gs.CD2,gs.xs12,gs.xs22,gs.w3opt,gs.S1NYS_T,gs.S2NYS_T,gs.S1NYS_T+gs.S2NYS_T])
                if gs.MechanicType =='Center':
                    gs.parameterOpt =='ED'
                    
                else:
                    gs.parameterOpt =='BC'
                    gs.listOpt2[0][0] = gs.BC2b
                    gs.listOpt2[1][0] = gs.BC2
                for i,item in enumerate(gs.listOpt1):
                    for j,par in enumerate(item):
                        self.TableOpt.setItem(i,j,Qitem('%.4f'%par))
                self.TableOpt.setItem(2,1,Qitem('%.2f'%gs.w2opt))
                self.TableOpt.setItem(2,6,Qitem('%.2f'%gs.M1NYS_T))
                self.TableOpt.setItem(2,7,Qitem('%.2f'%gs.M2NYS_T))


                for i,item in enumerate(gs.listOpt2):
                    for j,par in enumerate(item):
                        self.TableOpt2.setItem(i,j,Qitem('%.4f'%par))
                self.TableOpt2.setItem(2,1,Qitem('%.2f'%gs.w3opt))
                self.TableOpt2.setItem(2,6,Qitem('%.2f'%gs.S1NYS_T))
                self.TableOpt2.setItem(2,7,Qitem('%.2f'%gs.S2NYS_T))

                self.textEdit_3.setText(gs.parameterOpt)
                self.TextW2.setText('%.4f'%gs.w2Target)
                self.TextW3.setText('%.4f'%gs.w3Target)
                
    @staticmethod
    def rounding2():
        gs.CD2=round(gs.CD2b,1)
        if gs.MechanicType =='Center':
            gs.ED2= round(gs.ED2b,1)
        else:
            gs.BC2=round(gs.BC2b,1)
        gs.xm12 = gs.xm12b
        gs.xm22 = gs.xm22b
        gs.Ra_Rb2=float(gs.BC2/gs.CD2)
        outS3=Func.Output(gs.BC2,gs.CD2,gs.ED2,gs.xm12,gs.A2,gs.B2,gs.E2,gs.F2,KBEW=gs.KBEW2) #slave UWL
        outS4=Func.Output(gs.BC2,gs.CD2,gs.ED2,gs.xm22,gs.A2,gs.B2,gs.E2,gs.F2,KBEW=gs.KBEW2) #slave OWL
        #[alpha,beta,NYS_T,NYS_A,NYK_T,NYK_A,C[0][0],C[0][1],C[0][2],D[0][0],D[0][1],D[0][2]]
        gs.xs12 = outS3[1]
        gs.xs22 = outS4[1]
        gs.w3opt=Func.angleDiff(gs.xs22,gs.xs12)
        gs.S1NYS_T=outS3[2]
        gs.S2NYS_T=outS4[2]
    @staticmethod
    def rounding():
        gs.CD=round(gs.CDb,1)
        gs.ED=round(gs.EDb,1)
        gs.xm1 = gs.xm1b
        gs.xm2 = gs.xm2b
        if gs.DriveType=='Reversing':
            gs.Alfa = round(gs.Alfab,1)
            gs.offsetAngle = round(gs.offsetAngleb,1)
            [equal1,equal2,NYK_T,NYS_T,Dbx,Dby,Dbz] = Func.OutputSymbol  (gs.A,gs.B,gs.E,gs.F)
            Equal1inline = repr(equal1[0]).replace('xCrank','gs.ED').replace('xLink','gs.CD').replace('xm','angle[0]').replace('xs','angle[1]').replace('sin','sp.sin').replace('cos','sp.cos').replace('sqrt','sp.sqrt').replace('BC','gs.BC') # link length equation for master link
            Equal2inline = repr(equal2).replace('xCrank','gs.ED').replace('xLink','gs.CD').replace('xm','angle[0]').replace('xs','angle[1]').replace('sin','sp.sin').replace('cos','sp.cos').replace('sqrt','sp.sqrt').replace('BC','gs.BC')
            EqualInline = [Equal1inline,Equal2inline]
            def Finline(angle):
                F1 = eval(Equal1inline)
                F2 = 1000*(eval(Equal2inline)**2-1)
                return([F1,F2])
            if gs.KBEW =='+x':
                [xim1,xis1]= op.fsolve(Finline,[0,0])
            else :
                [xim1,xis1] = op.fsolve(Finline,[math.pi,math.pi])

            #[xim2,xis2]= op.fsolve(Finline,[math.pi,math.pi])
            #gs.xm1 = gs.xm3+gs.offsetAngle
            gs.xm1 = gs.xm1b
            gs.xm2 = gs.xm1+gs.Alfa
        else:
            gs.offsetAngle = 0
            gs.Alfa = 360
        gs.Ra_Rb = float(gs.BC/gs.CD)
        outM1=Func.Output(gs.BC,gs.CD,gs.ED,gs.xm1,gs.A,gs.B,gs.E,gs.F,KBEW=gs.KBEW)  #Master -UWL,[alpha,beta,NYS_T,NYS_A,NYK_T,NYK_A,C[0][0],C[0][1],C[0][2],Db[0][0],Db[0][1],Db[0][2]]
        outM2=Func.Output(gs.BC,gs.CD,gs.ED,gs.xm2,gs.A,gs.B,gs.E,gs.F,KBEW= gs.KBEW) #Master-OWL [alpha,beta,NYS_T,NYS_A,NYK_T,NYK_A,C[0][0],C[0][1],C[0][2],Db[0][0],Db[0][1],Db[0][2]]
        gs.xs1 = outM1[1]
        gs.xs2 = outM2[1]
        # cxs1 = math.cos(Func.rad(gs.xs1))
        # cxs2 = math.cos(Func.rad(gs.xs2))
        # sxs1 = math.sin(Func.rad(gs.xs1))
        # sxs2 = math.sin(Func.rad(gs.xs2))
        gs.w2opt = Func.wipAngle(gs.xs2,gs.xs1,gs.KBEW)
        gs.M1NYS_T=float(outM1[2])
        gs.M2NYS_T=float(outM2[2])
        gs.M1NYK_T=float(outM1[4])
        gs.M2NYK_T=float(outM2[4])

    def OptClicked(self):
        ab = gs.B-gs.A
        ab2 = gs.B2-gs.A2
        BC = gs.BC
        Distance = np.linalg.norm(ab)
        gs.Distance = np.sign(ab[2])*Distance
        Distance2 = np.linalg.norm(ab2)
        gs.Distance2 = np.sign(ab2[2]) * Distance2
        #clock means if clockwise for negative crank TBD
        if (gs.ParkWhere=='+y'):
                Clock = -1 if gs.KBEW=='-x' else 1
                Clock2 = -1 if gs.KBEW2=='-x' else 1
        elif (gs.ParkWhere=='-y'):
                Clock = -1 if gs.KBEW=='+x' else 1
                Clock2 = -1 if gs.KBEW2=='+x' else 1
        else:
                print('park where error')

        # =============================================================================
        # self.TextW2.setText(str(gs.w2Target))
        # self.TextW3.setText(str(gs.w3Target))
        # self.Text20.setText(str('0'))
        # self.Text30.setText(str('0'))
        # =============================================================================
        print('start optimization')
        time_start=time.time()

        Delta2Rad = Func.rad(gs.Delta2)
        cDelta2 = math.cos(Delta2Rad)
        sDelta2 = math.sin(Delta2Rad)
        # cal
        [equal1,equal2,NYK_T,NYS_T,Dbx,Dby,Dbz] = Func.OutputSymbol  (gs.A,gs.B,gs.E,gs.F)
        [equal12,equal22,NYK_T2,NYS_T2,Dbx2,Dby2,Dbz2] = Func.OutputSymbol(gs.A2,gs.B2,gs.E2,gs.F2)

        # Dbx=repr(Dbx2).replace('xCrank','model2.outCrank2').replace('xLink','model2.link2').replace('sin(xs)','model2.sxs12').replace('cos(xs)','model2.cxs12').replace('sin(xm)','model2.sxm12').replace('cos(xm)','model2.cxm12')
        # Dby=repr(Dby2).replace('xCrank','model2.outCrank2').replace('xLink','model2.link2').replace('sin(xs)','model2.sxs12').replace('cos(xs)','model2.cxs12').replace('sin(xm)','model2.sxm12').replace('cos(xm)','model2.cxm12')
        # Dbz=repr(Dbz2).replace('xCrank','model2.outCrank2').replace('xLink','model2.link2').replace('sin(xs)','model2.sxs12').replace('cos(xs)','model2.cxs12').replace('sin(xm)','model2.sxm12').replace('cos(xm)','model2.cxm12')

        # generating equation
        #optimize
        Equal1 = repr(equal1[0]).replace('xCrank','model.outCrank').replace('xLink','model.link') # link length equation for master link
        Equal2 = repr(equal2).replace('xCrank','model.outCrank').replace('sqrt','pe.sqrt') # bc parrel bd equation for master link
                                                                            
        F_M11=Equal1.replace('sin(xs)','model.sxs1').replace('cos(xs)','model.cxs1').replace('sin(xm)','model.sxm1').replace('cos(xm)','model.cxm1') # 11 means equation 1 for start position
        F_M12=Equal1.replace('sin(xs)','model.sxs2').replace('cos(xs)','model.cxs2').replace('sin(xm)','model.sxm2').replace('cos(xm)','model.cxm2')
        F_M13=Equal1.replace('sin(xs)','model.sxs3').replace('cos(xs)','model.cxs3').replace('sin(xm)','model.sxm3').replace('cos(xm)','model.cxm3') # 11 means equation 1 for start position
        F_M21=Equal2.replace('sin(xs)','model.sxs1').replace('cos(xs)','model.cxs1').replace('sin(xm)','model.sxm1').replace('cos(xm)','model.cxm1')
        F_M22=Equal2.replace('sin(xs)','model.sxs2').replace('cos(xs)','model.cxs2').replace('sin(xm)','model.sxm2').replace('cos(xm)','model.cxm2')
        F_M23=Equal2.replace('sin(xs)','model.sxs3').replace('cos(xs)','model.cxs3').replace('sin(xm)','model.sxm3').replace('cos(xm)','model.cxm3')
        F_M1NYK_T= repr(NYK_T).replace('sin(xm)','model.sxm1').replace('cos(xm)','model.cxm1').replace('acos','pe.acos').replace('xCrank','model.outCrank').replace('acos','pe.acos').replace('sin(xm)','model.sxm1').replace('cos(xm)','model.cxm1').replace('sqrt','pe.sqrt').replace('sin(xs)','model.sxs1').replace('cos(xs)','model.cxs1')
        F_M2NYK_T= repr(NYK_T).replace('sin(xm)','model.sxm2').replace('cos(xm)','model.cxm2').replace('acos','pe.acos').replace('xCrank','model.outCrank').replace('acos','pe.acos').replace('sin(xm)','model.sxm1').replace('cos(xm)','model.cxm1').replace('sqrt','pe.sqrt').replace('sin(xs)','model.sxs2').replace('cos(xs)','model.cxs2')
        F_M1NYS_T= repr(NYS_T).replace('sin(xs)','model.sxs1').replace('cos(xs)','model.cxs1').replace('xCrank','model.outCrank').replace('acos','pe.acos').replace('sin(xm)','model.sxm1').replace('cos(xm)','model.cxm1').replace('sqrt','pe.sqrt')
        F_M2NYS_T= repr(NYS_T).replace('sin(xs)','model.sxs2').replace('cos(xs)','model.cxs2').replace('xCrank','model.outCrank').replace('acos','pe.acos').replace('sin(xm)','model.sxm2').replace('cos(xm)','model.cxm2').replace('sqrt','pe.sqrt')
        
        if gs.MechanicType =='Center':# only one model , master and slave is optimizaed together
            Equal12 = repr(equal12[0]).replace('xCrank','model.outCrank2').replace('xLink','model.link2').replace('BC','model.bc2') #link length equation for slave link
            Equal22 = repr(equal22).replace('xCrank', 'model.outCrank2').replace('sqrt','pe.sqrt').replace('BC','model.bc2') #model.bc2 TBD for center
            F_S11=Equal12.replace('sin(xs)','model.sxs12').replace('cos(xs)','model.cxs12').replace('sin(xm)','model.sxm12').replace('cos(xm)','model.cxm12')
            F_S12=Equal12.replace('sin(xs)','model.sxs22').replace('cos(xs)','model.cxs22').replace('sin(xm)','model.sxm22').replace('cos(xm)','model.cxm22') # two position start and OWL
            F_S1NYS_T= repr(NYS_T2).replace('sin(xs)','model.sxs12').replace('cos(xs)','model.cxs12').replace('xCrank','model.outCrank2').replace('acos','pe.acos').replace('sin(xm)','model.sxm12').replace('cos(xm)','model.cxm12').replace('sqrt','pe.sqrt').replace('BC','model.bc2')
            F_S2NYS_T= repr(NYS_T2).replace('sin(xs)','model.sxs22').replace('cos(xs)','model.cxs22').replace('xCrank','model.outCrank2').replace('acos','pe.acos').replace('sin(xm)','model.sxm22').replace('cos(xm)','model.cxm22').replace('sqrt','pe.sqrt').replace('BC','model.bc2')
            F_S1NYK_T= repr(NYK_T2).replace('sin(xm)','model.sxm12').replace('cos(xm)','model.cxm12').replace('sin(xs)','model.sxs12').replace('cos(xs)','model.cxs12').replace('acos','pe.acos').replace('xCrank','model.outCrank2').replace('sqrt','pe.sqrt').replace('BC','model.bc2')
            F_S2NYK_T= repr(NYK_T2).replace('sin(xm)','model.sxm22').replace('cos(xm)','model.cxm22').replace('sin(xs)','model.sxs22').replace('cos(xs)','model.cxs22').replace('acos','pe.acos').replace('xCrank','model.outCrank2').replace('sqrt','pe.sqrt').replace('BC','model.bc2')
            F_S21 = Equal22.replace('sin(xs)', 'model.sxs12').replace('cos(xs)', 'model.cxs12').replace('sin(xm)','model.sxm12').replace('cos(xm)', 'model.cxm12')
            F_S22 = Equal22.replace('sin(xs)', 'model.sxs22').replace('cos(xs)', 'model.cxs22').replace('sin(xm)', 'model.sxm22').replace('cos(xm)', 'model.cxm22')
        else:# two model, master and slave is optimized seperately
            Equal12 = repr(equal12[0]).replace('xCrank','model2.outCrank2').replace('xLink','model2.link2').replace('BC','model2.bc2') #link length equation for slave link
            Equal22 = repr(equal22).replace('xCrank', 'model2.outCrank2').replace('sqrt','pe.sqrt').replace('BC','model2.bc2') #model.bc2 TBD for center
            F_S11 = Equal12.replace('sin(xs)','model2.sxs12').replace('cos(xs)','model2.cxs12').replace('sin(xm)','model2.sxm12').replace('cos(xm)','model2.cxm12')
            F_S12 = Equal12.replace('sin(xs)','model2.sxs22').replace('cos(xs)','model2.cxs22').replace('sin(xm)','model2.sxm22').replace('cos(xm)','model2.cxm22') # two position start and OWL
            F_S1NYS_T= repr(NYS_T2).replace('sin(xs)','model2.sxs12').replace('cos(xs)','model2.cxs12').replace('xCrank','model2.outCrank2').replace('acos','pe.acos').replace('sin(xm)','model2.sxm12').replace('cos(xm)','model2.cxm12').replace('sqrt','pe.sqrt').replace('BC','model2.bc2')
            F_S2NYS_T= repr(NYS_T2).replace('sin(xs)','model2.sxs22').replace('cos(xs)','model2.cxs22').replace('xCrank','model2.outCrank2').replace('acos','pe.acos').replace('sin(xm)','model2.sxm22').replace('cos(xm)','model2.cxm22').replace('sqrt','pe.sqrt').replace('BC','model2.bc2')
            F_S1NYK_T= repr(NYK_T2).replace('sin(xm)','model2.sxm12').replace('cos(xm)','model2.cxm12').replace('sin(xs)','model2.sxs12').replace('cos(xs)','model2.cxs12').replace('acos','pe.acos').replace('xCrank','model2.outCrank2').replace('sqrt','pe.sqrt').replace('BC','model2.bc2')
            F_S2NYK_T= repr(NYK_T2).replace('sin(xm)','model2.sxm22').replace('cos(xm)','model2.cxm22').replace('sin(xs)','model2.sxs22').replace('cos(xs)','model2.cxs22').replace('acos','pe.acos').replace('xCrank','model2.outCrank2').replace('sqrt','pe.sqrt').replace('BC','model2.bc2')
        # general equation
        model = pe.ConcreteModel()
        model.link = pe.Var(initialize = gs.CD,bounds=(3*gs.BC,10*gs.BC) ) #length of link 
        model.outCrank = pe.Var(initialize =gs.ED,bounds = (0.5*gs.BC,2*gs.BC)) # length of output crank 
        
        model.cxm1 = pe.Var(initialize = math.sqrt(3)/2,bounds=(-1,1))
        model.cxm2 = pe.Var(initialize = -math.sqrt(3)/2,bounds=(-1,1))
        if gs.ParkWhere =='+y':
            model.sxm1 =  pe.Var(initialize =-0.5,bounds=(-1,0)) 
            model.sxm2 =  pe.Var(initialize =0.5,bounds=(-1,1)) 
        else:
            model.sxm1 =  pe.Var(initialize =0.5,bounds=(0,1)) 
            model.sxm2 =  pe.Var(initialize =-0.5,bounds=(-1,1)) 
        
        if gs.KBEW =='-x':
            model.sxs1 = pe.Var(initialize=0.5, bounds=(0, 1))
            model.cxs1 = pe.Var(initialize=-math.sqrt(3) / 2, bounds=(-1, 0))
            model.sxs2 = pe.Var(initialize=-0.5, bounds=(-1, 0))
            model.cxs2 = pe.Var(initialize=-math.sqrt(3) / 2, bounds=(-1, 0))
        elif gs.KBEW=='+x':
            model.sxs1 = pe.Var(initialize=-0.5, bounds=(-1, 0))
            model.cxs1 = pe.Var(initialize=math.sqrt(3) / 2, bounds=(0, 1))
            model.sxs2 = pe.Var(initialize=0.5, bounds=( 0 , 1))
            model.cxs2 = pe.Var(initialize=math.sqrt(3) / 2, bounds=(0, 1))    
        model.Con = pe.ConstraintList()
        model.Con.add(expr=(model.sxm2 ** 2 + model.cxm2 ** 2 - 1) == 0)
        model.Con.add(expr=(model.sxm1 ** 2 + model.cxm1 ** 2 - 1) == 0)
        model.Con.add(expr=(model.sxs1 ** 2 + model.cxs1 ** 2 - 1) == 0)
        model.Con.add(expr=(model.sxs2 ** 2 + model.cxs2 ** 2 - 1) == 0)
        model.Con.add(eval(F_M11)==0) # meet basic physical constrain
        model.Con.add(eval(F_M12)==0) # meet basic physical constrain

        if (Clock==1):
            model.Con.add(model.sxs2*model.cxs1-model.cxs2*model.sxs1>=0)
        elif(Clock==-1) :
            model.Con.add(model.sxs2*model.cxs1-model.cxs2*model.sxs1<=0)
        else:
            print(gs.KBEW)

        if (gs.DriveType=='Standard'):
            # standard system, 1 and 2 are in inlien position
            model.Con.add(1000*(eval(F_M22)**2-1)==0) #
            model.Con.add(1000*(eval(F_M21)**2-1)==0)
            
        elif (gs.DriveType=='Reversing'): #
            # reversing system , 1 and 2 not in inline position , 3 is inline position
            model.cxm3 = pe.Var(initialize = -math.sqrt(3)/2,bounds=(-1,1))
            if gs.ParkWhere =='+y':
                model.sxm3 =  pe.Var(initialize =-0.5,bounds=(-1,0)) 
            else:
                model.sxm3 =  pe.Var(initialize =0.5,bounds=(0,1)) 
            

            model.sxs3 = pe.Var(initialize=-0.5, bounds=(-1, 1))
            model.cxs3 = pe.Var(initialize=math.sqrt(3) / 2, bounds=(-1, 1))

            model.Con.add(eval(F_M13)==0)
            model.Con.add(1000*(eval(F_M23)**2-1)==0) 
            model.Con.add(expr=(model.sxm3**2+model.cxm3**2-1)==0)
            model.Con.add(expr=(model.sxs3**2+model.cxs3**2-1)==0)
            
            model.Con.add(eval(F_M1NYK_T)+eval(F_M2NYK_T)==0)
            cOff =model.cxm1*model.cxm3+model.sxm1*model.sxm3 #offset =xm1-xm3
            sOff = model.sxm1*model.cxm3- model.cxm1*model.sxm3

            cAlfa =model.cxm2*model.cxm1+model.sxm2*model.sxm1 #alpha =xm2-xm1
            sAlfa = model.sxm2*model.cxm1-model.cxm2*model.sxm1
            
            #NYK_T = eval(F_M1NYK_T)
            #sNYK_T = pe.sqrt(1-eval(F_M1NYK_T)*eval(F_M1NYK_T))
            model.Con.add(expr = cAlfa>=math.cos(170*math.pi/180.)) # alfa <170
            model.Con.add(expr = sOff*cAlfa+cOff*sAlfa >=0) # off+alfa<180
            model.Con.add(expr = sOff>=0  ) # off alfa must be the same sign to avoid cross inline position
            #model.Con.add(expr = sNYK_T*cxm21+cNYK_T*sxm21 >=0) # alpha+offsetangle <180, cannot cover inline position
            #model.Con.add(eval(F_M1NYK_T) ** 2 <= (math.cos(14 * math.pi / 180)) ** 2) #NYK_T<76

        if gs.MechanicType =='Center':                                   
            model.link2 = pe.Var(initialize=gs.CD2, bounds=(4.4 * gs.BC2, 10 * gs.BC2))  # length of link
            model.outCrank2 = pe.Var(initialize=gs.ED2, bounds=(0.25*gs.BC2, 4*gs.BC2))  # length of output crank
            model.sxm12 = pe.Var(initialize=0, bounds=(-1, 1))
            model.cxm12 = pe.Var(initialize=1, bounds=(-1, 1))
            model.sxm22 = pe.Var(initialize=0, bounds=(-1, 1))
            model.cxm22 = pe.Var(initialize=-1, bounds=(-1, 1))
            if gs.KBEW2 =='-x':
                model.sxs12 = pe.Var(initialize=0.5, bounds=(0, 1))
                model.cxs12 = pe.Var(initialize=-math.sqrt(3) / 2, bounds=(-1, 0))
                model.sxs22 = pe.Var(initialize=-0.5, bounds=(-1, 0))
                model.cxs22 = pe.Var(initialize=-math.sqrt(3) / 2, bounds=(-1, 0))
            elif gs.KBEW2 =='+x':
                model.sxs12 = pe.Var(initialize=-0.5, bounds=(-1, 0))
                model.cxs12 = pe.Var(initialize=math.sqrt(3) / 2, bounds=(0, 1))
                model.sxs22 = pe.Var(initialize=0.5, bounds=( 0 , 1))
                model.cxs22 = pe.Var(initialize=math.sqrt(3) / 2, bounds=(0, 1))  
            model.bc2 = pe.Var(initialize=gs.BC, bounds=(gs.BC, gs.BC))
            model.obj = pe.Objective(
                expr=((model.cxs2 * model.cxs1 + model.sxs2 * model.sxs1 - math.cos(gs.w2)) ** 2)  # minmize wipping angle error
            +(model.cxs22 * model.cxs12 + model.sxs22 * model.sxs12 - math.cos(gs.w3)) ** 2)  # the lowest wipping angle requirement
            # =============================================================================
            model.Con.add ((model.cxs2 * model.cxs1 + model.sxs2 * model.sxs1 - math.cos(gs.w2)) ** 2<=0.01)
            model.Con.add((model.cxs22 * model.cxs12 + model.sxs22 * model.sxs12 - math.cos(gs.w3)) ** 2<=0.01)
            model.Con.add(expr=(model.sxm22 ** 2 + model.cxm22 ** 2 - 1) == 0)
            model.Con.add(expr=(model.sxm12 ** 2 + model.cxm12 ** 2 - 1) == 0)
            model.Con.add(expr=(model.sxs12 ** 2 + model.cxs12 ** 2 - 1) == 0)
            model.Con.add(expr=(model.sxs22 ** 2 + model.cxs22 ** 2 - 1) == 0)
            if Clock2:
                model.Con.add(model.sxs22*model.cxs12-model.sxs12*model.cxs22>=0)
            else:
                print('pls specicy clock')

            model.Con.add(eval(F_S11) == 0)
            model.Con.add(eval(F_S12) == 0)  # physical requirement
            model.Con.add(1000 * (eval(F_S21)**2 - 1) == 0)
            model.Con.add(1000 * (eval(F_S22)**2- 1) == 0)  #
            model.Con.add((eval(F_M1NYS_T)) + (eval(F_M2NYS_T)) == 0)
            model.Con.add((eval(F_S1NYS_T)) + (eval(F_S2NYS_T)) == 0)
            model.Con.add(eval(F_M1NYS_T) ** 2 <= (math.cos(40 * math.pi / 180)) ** 2)  # NYS_T<50
            model.Con.add(eval(F_S1NYS_T) ** 2 <= (math.cos(40 * math.pi / 180)) ** 2)  # NYS_T<50
            # model.Con.add(expr=model.cxm1*model.cxm12+model.sxm1*model.sxm12>=math.cos(8*math.pi/180))
            # =============================================================================
            M1NYS_Tv = eval(F_M1NYS_T)
            M2NYS_Tv = eval(F_M2NYS_T)
            M1NYK_Tv = eval(F_M1NYK_T)
            M2NYK_Tv = eval(F_M2NYK_T)
            
            S1NYS_Tv = eval(F_S1NYS_T)
            S2NYS_Tv = eval(F_S2NYS_T)
            S1NYK_Tv = eval(F_S1NYK_T)
            S2NYK_Tv = eval(F_S2NYK_T)
            # =============================================================================
            opt = pe.SolverFactory('ipopt')
            result_obj1 = opt.solve(model)
            str1 = "The solver returned a status of: " + str(result_obj1.solver.status)
            str2 = "The solver terminated when: " + str(result_obj1.solver.termination_condition)
            print('------------------------------master link info------------------------')
            print(str1)
            print(str2)
            if str(result_obj1.solver.status).strip() !='ok':
                model.display()
            sxm1 = pe.value(model.sxm1)
            cxm1 = pe.value(model.cxm1)
            sxs1 = pe.value(model.sxs1)
            cxs1 = pe.value(model.cxs1)
            sxm2 = pe.value(model.sxm2)
            cxm2 = pe.value(model.cxm2)
            sxs2 = pe.value(model.sxs2)
            cxs2 = pe.value(model.cxs2)
            
            gs.xm1b = Func.getDegree(cxm1, sxm1)
            gs.xm2b = Func.getDegree(cxm2, sxm2)
            gs.xs1b = Func.getDegree(cxs1, sxs1)
            gs.xs2b = Func.getDegree(cxs2, sxs2)
            
            
            gs.CDb = (pe.value((model.link)))
            gs.EDb = (pe.value(model.outCrank))
            M1NYS_Tb = Func.degree(math.pi / 2 - math.acos(pe.value(M1NYS_Tv)))
            M2NYS_Tb = Func.degree(math.pi / 2 - math.acos(pe.value(M2NYS_Tv)))
            M1NYK_Tb = Func.degree(math.pi / 2 - math.acos(pe.value(M1NYK_Tv)))
            M1NYK_Tb = Func.degree(math.pi / 2 - math.acos(pe.value(M1NYK_Tv)))

            sxm12 = pe.value(model.sxm12)
            cxm12 = pe.value(model.cxm12)
            sxm22 = pe.value(model.sxm22)
            cxm22 = pe.value(model.cxm22)
            sxs12 = pe.value(model.sxs12)
            cxs12 = pe.value(model.cxs12)
            sxs22 = pe.value(model.sxs22)
            cxs22 = pe.value(model.cxs22)

            gs.xs22b = Func.getDegree(cxs22, sxs22)
            gs.xs12b = Func.getDegree(cxs12, sxs12)
            gs.xm22b = Func.getDegree(cxm22, sxm22)
            gs.xm12b = Func.getDegree(cxm12, sxm12)
            gs.CD2b = (pe.value((model.link2)))
            gs.ED2b = (pe.value(model.outCrank2))
         
            # =============================================================================
            gs.S1NYS_Tb = Func.degree(math.pi / 2 - math.acos(pe.value(S1NYS_Tv)))
            gs.S2NYS_Tb = Func.degree(math.pi / 2 - math.acos(pe.value(S2NYS_Tv)))
            gs.S1NYK_Tb = Func.degree(math.pi / 2 - math.acos(pe.value(S1NYK_Tv)))
            gs.S2NYK_Tb = Func.degree(math.pi / 2 - math.acos(pe.value(S2NYK_Tv)))
            
            gs.w2optb = Func.wipAngle(gs.xs2b,gs.xs1b,gs.KBEW)
            gs.w3optb = Func.wipAngle(gs.xs22b,gs.xs12b,gs.KBEW)
            self.rounding()
            self.rounding2()

        else: # Master-Slave
                    #--- master link
            model.obj=pe.Objective(expr=(model.cxs2*model.cxs1+model.sxs2*model.sxs1-math.cos(gs.w2))**2) # minmize wipping angle error
            model.Con.add(expr=(model.cxs2 * model.cxs1 + model.sxs2 * model.sxs1 - math.cos(
                    gs.w2)) ** 2 <= 0.001)  # the lowest wipping angle requirement
            # =============================================================================
            model.Con.add((eval(F_M1NYS_T)) + (eval(F_M2NYS_T)) == 0)
            model.Con.add(eval(F_M1NYS_T)**2<=(math.cos(40*math.pi/360))**2)#NYS_T<50

            # =============================================================================
            opt = pe.SolverFactory('ipopt')
            result_obj1 = opt.solve(model) 
            str1="The solver returned a status of: " + str(result_obj1.solver.status)
            str2= "The solver terminated when: " + str(result_obj1.solver.termination_condition)
            print('------------------------------master link info------------------------')
            print(str1)
            print(str2)
            if str(result_obj1.solver.status)!='ok':
                model.display()
                
            sxm1=pe.value(model.sxm1)
            cxm1=pe.value(model.cxm1)
            sxs1=pe.value(model.sxs1)
            cxs1=pe.value(model.cxs1)
            sxm2=pe.value(model.sxm2)
            cxm2=pe.value(model.cxm2)
            sxs2=pe.value(model.sxs2)
            cxs2=pe.value(model.cxs2)

            gs.xs1b=Func.getDegree(cxs1,sxs1)
            gs.xm1b=Func.getDegree(cxm1,sxm1)
            gs.xm2b=Func.getDegree(cxm2,sxm2)
            gs.xs2b=Func.getDegree(cxs2,sxs2)
            if gs.DriveType=='Reversing':
                sxm3=pe.value(model.sxm3)
                cxm3=pe.value(model.cxm3)
                sxs3=pe.value(model.sxs3)
                cxs3=pe.value(model.cxs3)
                gs.xs3b=Func.getDegree(cxs3,sxs3)
                gs.xm3b=Func.getDegree(cxm3,sxm3)

            M1NYS_Tv = eval(F_M1NYS_T)
            M2NYS_Tv = eval(F_M2NYS_T)
            M1NYK_Tv = eval(F_M1NYK_T)
            M2NYK_Tv = eval(F_M2NYK_T)

            gs.CDb=  (pe.value((model.link)))
            gs.EDb=(pe.value(model.outCrank))
            gs.M1NYS_Tb=Func.degree(math.pi/2-math.acos(pe.value(M1NYS_Tv)))
            gs.M2NYS_Tb=Func.degree(math.pi/2-math.acos(pe.value(M2NYS_Tv)))
            gs.M1NYK_Tb=Func.degree(math.pi/2-math.acos(pe.value(M1NYK_Tv)))
            gs.M2NYK_Tb=Func.degree(math.pi/2-math.acos(pe.value(M2NYK_Tv)))

            if gs.DriveType=='Reversing':
                

                gs.Alfab = gs.xm2b-gs.xm1b
                gs.offsetAngleb = gs.xm1b-gs.xm3b
            else:
                gs.offsetAngleb = 0
                gs.Alfab = 360

            gs.w2optb = Func.wipAngle(gs.xs2b,gs.xs1b,gs.KBEW)
            self.rounding()
            # calculate cos,sin of  S angle after rounding
            cxs1 = math.cos(Func.rad(gs.xs1))
            cxs2 = math.cos(Func.rad(gs.xs2))
            sxs1 = math.sin(Func.rad(gs.xs1))
            sxs2 = math.sin(Func.rad(gs.xs2))
        
            #=============================================================================

            #slave  link
            print('------------------slave crank info---------------')
            model2 = pe.ConcreteModel()
            model2.link2 = pe.Var(initialize = gs.CD2,bounds=(4.5*gs.BC,10*gs.BC) ) # length of link 
            model2.outCrank2 = pe.Var(initialize =gs.ED2,bounds = (gs.ED2,gs.ED2)) # length of output crank
            model2.cDelta2 = pe.Var(initialize =cDelta2,bounds = (-1,1)) # coupling angle
            model2.sDelta2 = pe.Var(initialize =sDelta2,bounds = (-1,1)) # coupling angle
            model2.sxm12 =  pe.Var(initialize =0,bounds=(-1,1)) 
            model2.cxm12 = pe.Var(initialize = 1,bounds=(-1,1))
            model2.sxm22 =  pe.Var(initialize =0,bounds=(-1,1)) 
            model2.cxm22 = pe.Var(initialize = -1,bounds=(-1,1))
            if gs.KBEW2 =='-x':
                model2.sxs12 = pe.Var(initialize=0.5, bounds=(0, 1))
                model2.cxs12 = pe.Var(initialize=-math.sqrt(3) / 2, bounds=(-1, 0))
                model2.sxs22 = pe.Var(initialize=-0.5, bounds=(-1, 0))
                model2.cxs22 = pe.Var(initialize=-math.sqrt(3) / 2, bounds=(-1, 0))
            elif gs.KBEW2=='+x':
                model2.sxs12 = pe.Var(initialize=-0.5, bounds=(-1, 0))
                model2.cxs12 = pe.Var(initialize=math.sqrt(3) / 2, bounds=(0, 1))
                model2.sxs22 = pe.Var(initialize=0.5, bounds=( 0 , 1))
                model2.cxs22 = pe.Var(initialize=math.sqrt(3) / 2, bounds=(0, 1))    
                
            model2.bc2 = pe.Var(initialize =gs.BC,bounds=(0.25*BC,4*BC) ) #link length
            model2.Con = pe.ConstraintList()
            model2.obj = pe.ObjectiveList()
            # =============================================================================
            model2.obj.add(expr=(100*(model2.cxs22*model2.cxs12+model2.sxs22*model2.sxs12-math.cos(gs.w3))**2+(eval(F_S1NYS_T)+eval(F_S2NYS_T))**2+(eval(F_S1NYK_T)+eval(F_S2NYK_T))**2)) # minize wipping angel 
            # =============================================================================
            #==============================================================================
            # model2.Con.add(expr=1000*(model2.cxs22*model2.cxs12+model2.sxs22*model2.sxs12-math.cos(w3))>=0.0) #wipping angle requirement 
            # model2.Con.add(eval(F12NYS_T)+eval(F22NYS_T)>=0)
            # model2.Con.add(eval(F12NYK_T)+eval(F22NYK_T)>=0) # move to objective function
            #==============================================================================
            if (Clock2==1):
                model.Con.add(model.sxs2*model.cxs1-model.cxs2*model.sxs1>=0)
            elif(Clock2==-1) :
                model.Con.add(model.sxs2*model.cxs1-model.cxs2*model.sxs1<=0)
            else:
                print(gs.KBEW)
            # =============================================================================
            model2.Con.add(eval(F_S11)==0)
            model2.Con.add(eval(F_S12)==0)
            #==============================================================================
            # model2.Con.add(expr=model2.sxm12**2+model2.cxm12**2-1==0)
            # model2.Con.add(expr=model2.sxm22**2+model2.cxm22**2-1==0)
            # =============================================================================
            model2.Con.add(expr=(model2.sxs12**2+model2.cxs12**2)==1)
            model2.Con.add(expr=(model2.sxs22**2+model2.cxs22**2)==1)
            
            model2.Con.add(expr=(model2.sDelta2**2+model2.cDelta2**2)==1)
            model2.Con.add(expr=(cxs1*model2.cDelta2-sxs1*model2.sDelta2-model2.cxm12)==0) 
            model2.Con.add(expr=(sxs1*model2.cDelta2+cxs1*model2.sDelta2-model2.sxm12)==0)
            model2.Con.add(expr=(cxs2*model2.cDelta2-sxs2*model2.sDelta2-model2.cxm22)==0) 
            model2.Con.add(expr=(sxs2*model2.cDelta2+cxs2*model2.sDelta2-model2.sxm22)==0)
            
            opt2 = pe.SolverFactory('ipopt')
            result_obj2 = opt2.solve(model2)
            str3="The solver returned a status of: " + str(result_obj2.solver.status)
            str4="The solver terminated when: " + str(result_obj2.solver.termination_condition)
            print(str3)
            print(str4)
            if str(result_obj2.solver.status)!='ok':
                model2.display()
            sxm12=pe.value(model2.sxm12)
            cxm12=pe.value(model2.cxm12)
            sxm22=pe.value(model2.sxm22)
            cxm22=pe.value(model2.cxm22)
            sxs12=pe.value(model2.sxs12)
            cxs12=pe.value(model2.cxs12)
            sxs22=pe.value(model2.sxs22)
            cxs22=pe.value(model2.cxs22)
            # =============================================================================
            sDelta2=pe.value(model2.sDelta2)
            cDelta2=pe.value(model2.cDelta2)
            # =============================================================================
            gs.BC2b=pe.value(model2.bc2)
            gs.xs22b=Func.getDegree(cxs22,sxs22)
            gs.xs12b=Func.getDegree(cxs12,sxs12)
            gs.xm22b=Func.getDegree(cxm22,sxm22)
            gs.xm12b=Func.getDegree(cxm12,sxm12)
            gs.Delta2=Func.getDegree(cDelta2,sDelta2)
            gs.CD2b=  (pe.value((model2.link2)))
            gs.ED2b=(pe.value(model2.outCrank2))
            S1NYS_Tv = eval(F_S1NYS_T)
            S2NYS_Tv = eval(F_S2NYS_T)
            S1NYK_Tv = eval(F_S1NYK_T)
            S2NYK_Tv = eval(F_S2NYK_T)
            
            gs.S1NYS_Tb=Func.degree(math.pi/2-math.acos(pe.value(S1NYS_Tv)))
            gs.S2NYS_Tb=Func.degree(math.pi/2-math.acos(pe.value(S2NYS_Tv)))
            gs.S1NYK_Tb=Func.degree(math.pi/2-math.acos(pe.value(S1NYK_Tv)))
            gs.S2NYK_Tb=Func.degree(math.pi/2-math.acos(pe.value(S2NYK_Tv)))
            
            gs.w3optb = Func.wipAngle(gs.xs22b,gs.xs12b,gs.KBEW2)

            ###rounding
            self.rounding2()

            # outM1=[float(s) for s in outM1]
            # outM2=[float(s) for s in outM2]
            # outS3=[float(s) for s in outS3]
            # outS4=[float(s) for s in outS4]
            # write to output sheet
        
            print('--------optimization completed------------')

        gs.printGS()
        gs.writeGS()
        gs.startCal()

class MainWindow_Input(QMainWindow,ui.Ui_Input):
    def __init__(self,parent=None):
        super(MainWindow_Input,self).__init__(parent)
        self.setupUi(self)
        self.initiInput()
       # self.LoadClicked()

    def UpdateOptValue(self):
        print('update opt value')
        GeneralBounds = [0,1000]
        Func.writeNewValueGUI(self.tableMasterCranklInfo,4,0,gs.CD,bounds = GeneralBounds)
        Func.writeNewValueGUI(self.tableMasterCranklInfo,5,0,gs.ED,bounds = GeneralBounds)
        Func.writeNewValueGUI(self.tableMasterCrankInfo2,0,1,'%.4f'%gs.Ra_Rb,bounds = [0,1/4.5])
        Func.writeNewValueGUI(self.tableMasterCrankInfo2,0,2,'%.4f'%gs.w2opt,bounds=[gs.w2Target-0.1,gs.w2Target+0.1])
        Func.writeNewValueGUI(self.tableSlaveCrankInfo2,0,1,'%.4f'%gs.Ra_Rb2,bounds=[0,1/4.5])
        Func.writeNewValueGUI(self.tableSlaveCrankInfo2,0,2,'%.4f'%gs.w3opt,bounds =[gs.w3Target-0.1, gs.w3Target+0.1])
        Func.writeNewValueGUI(self.tableMasterCranklInfo,1,0,'%.4f'%gs.Delta,bounds = [-90,90])
        Func.writeNewValueGUI(self.tableSlaveCrankInfo,1,0,'%.4f'%gs.Delta2,bounds = [-90,90])
        #self.tableMasterCranklInfo.setItem(4, 0 ,Qitem(str(gs.CD)))

        if gs.DriveType == 'Reversing':
            Func.writeNewValueGUI(self.tableMotorInfo,0,0,'%.4f'%gs.Alfa,bounds=[140,170])
            Func.writeNewValueGUI(self.tableMotorInfo,1,0,'%.4f'%gs.offsetAngle,bounds=[-180,180])



        if gs.MechanicType =='Center':
            Func.writeNewValueGUI(self.tableSlaveCrankInfo,4,0,gs.CD2,bounds=GeneralBounds)
            Func.writeNewValueGUI(self.tableSlaveCrankInfo,5,0,gs.ED2,bounds = GeneralBounds)

        else :
            Func.writeNewValueGUI(self.tableSlaveCrankInfo,4,0,gs.CD2,bounds = GeneralBounds)
            Func.writeNewValueGUI(self.tableSlaveCrankInfo,3,0,gs.BC2,bounds = GeneralBounds)
            
    def updateNewsheet(self):
        print('create new output sheet')
        gs.sheetNew = gs.wb.copy_worksheet(gs.sheet1)
        gs.sheetNew.title = 'Outputs'

        Func.writeResult(gs.sheetNew, 47, 2, '%.4f'%gs.Delta)
        Func.writeResult(gs.sheetNew, 64, 2, '%.4f'%gs.Delta2)

        Func.writeResult(gs.sheetNew, 47, 9, gs.CD)
        Func.writeResult(gs.sheetNew, 47, 15, gs.ED)
        Func.writeResult(gs.sheetNew, 49, 9, '%.4f'%gs.Ra_Rb)
        Func.writeResult(gs.sheetNew, 49, 17, '%.4f'%gs.w2opt)
        Func.writeResult(gs.sheetNew, 66, 9, '%.4f'%gs.Ra_Rb2)
        Func.writeResult(gs.sheetNew, 66, 17, '%.4f'%gs.w3opt)
        Func.writeResult(gs.sheetNew, 66, 17, '%.4f'%gs.w3opt)
        if gs.DriveType =='Reversing':
            Func.writeResult(gs.sheetNew, 22, 2, '%.4f'%gs.Alfa)
            Func.writeResult(gs.sheetNew, 22, 11, '%.4f'%gs.offsetAngle)


        if gs.MechanicType =='Center':

            Func.writeResult(gs.sheetNew, 64, 9, gs.CD2)
            Func.writeResult(gs.sheetNew, 64, 15, gs.ED2)
        else :

            Func.writeResult(gs.sheetNew, 64, 9, gs.CD2)
            Func.writeResult(gs.sheetNew, 61, 9, gs.BC2)
        
    def initiInput(self):
        print ('start initiInput')
    
        self.TextCustomer.setText(Func.read(gs.sheet1,5,2))
        self.TextProject.setText(Func.read(gs.sheet1,5,9))
        self.TextDepartment.setText(Func.read(gs.sheet1,5,15))
        self.TextDrawing.setText(Func.read(gs.sheet1,7,2))
        self.TextValid.setText(Func.read(gs.sheet1,7,9))
        self.TextName.setText(Func.read(gs.sheet1,7,15))
        self.TextComment.setText(Func.read(gs.sheet1,9,2))
        self.comboPark.setCurrentText(Func.read(gs.sheet1,14,7).strip())
        self.comboDirection.setCurrentText(Func.read(gs.sheet1,14,15).strip())
        self.comboDrive.setCurrentText ( Func.read(gs.sheet1,12,5).strip())
        self.comboMechanic.setCurrentText ( Func.read(gs.sheet1,12,15).strip())
        self.TextW2.setText(Func.read(gs.sheet1,16,6))
        self.TextW3.setText(Func.read(gs.sheet1,16,10))
        self.comboMeasured.setCurrentText (Func.read(gs.sheet1,16,17).strip())
        self.comboNo2.setCurrentText   (Func.read(gs.sheet1,35,7).strip())
        self.comboNo3.setCurrentText   (Func.read(gs.sheet1,52,7).strip())
        self.tableMotorAngle.setItem(0,0,Qitem(Func.read(gs.sheet1,24,4)))
        self.tableMotorAngle.setItem(0,1,Qitem(Func.read(gs.sheet1,24,9)))
        self.tableMotorAngle.setItem(0,2,Qitem(Func.read(gs.sheet1,24,13)))
        self.tableMotorAngle.setItem(0,3,Qitem(Func.read(gs.sheet1,26,4)))
        self.tableMotorAngle.setItem(0,4,Qitem(Func.read(gs.sheet1,26,9)))
        self.tableMotorAngle.setItem(0,5,Qitem(Func.read(gs.sheet1,26,13)))
        self.tableMotorAngle.setItem(0,6,Qitem(Func.read(gs.sheet1,26,18)))
        
        self.tableMotorInfo.setItem(0,0,Qitem(Func.read(gs.sheet1,22,2)))
        self.tableMotorInfo.setItem(1,0,Qitem(Func.read(gs.sheet1,22,11)))
        self.tableMotorInfo.setItem(2,0,Qitem(Func.read(gs.sheet1,29,2)))
        self.tableMotorInfo.setItem(3,0,Qitem(Func.read(gs.sheet1,29,9)))
        self.tableMotorInfo.setItem(4,0,Qitem(Func.read(gs.sheet1,29,15)))
        self.tableMotorInfo.setItem(5,0,Qitem(Func.read(gs.sheet1,32,2)))
        self.tableMotorInfo.setItem(6,0,Qitem(Func.read(gs.sheet1,32,9)))
        self.tableMotorInfo.setItem(7,0,Qitem(Func.read(gs.sheet1,32,15)))
        self.tableMotorInfo.setItem(0,1,Qitem(Func.read(gs.sheet1,21,7)))
        self.tableMotorInfo.setItem(1,1,Qitem(Func.read(gs.sheet1,21,16)))
        self.tableMotorInfo.setItem(2,1,Qitem(Func.read(gs.sheet1,28,6)))
        self.tableMotorInfo.setItem(3,1,Qitem(Func.read(gs.sheet1,28,13)))
        self.tableMotorInfo.setItem(4,1,Qitem(Func.read(gs.sheet1,28,19)))
        self.tableMotorInfo.setItem(5,1,Qitem(Func.read(gs.sheet1,31,6)))
        self.tableMotorInfo.setItem(6,1,Qitem(Func.read(gs.sheet1,31,13)))
        self.tableMotorInfo.setItem(7,1,Qitem(Func.read(gs.sheet1,31,19)))
        self.tableMotorInfo.setItem(0,2,Qitem(Func.read(gs.sheet1,22,7)))        
        self.tableMotorInfo.setItem(1,2,Qitem(Func.read(gs.sheet1,22,16)))
        self.tableMotorInfo.setItem(2,2,Qitem(Func.read(gs.sheet1,29,6)))
        self.tableMotorInfo.setItem(3,2,Qitem(Func.read(gs.sheet1,29,13)))
        self.tableMotorInfo.setItem(4,2,Qitem(Func.read(gs.sheet1,29,19)))
        self.tableMotorInfo.setItem(5,2,Qitem(Func.read(gs.sheet1,32,6)))
        self.tableMotorInfo.setItem(6,2,Qitem(Func.read(gs.sheet1,32,13)))
        self.tableMotorInfo.setItem(7,2,Qitem(Func.read(gs.sheet1,32,19)))
        
        self.tableMasterCranklInfo.setItem(0,0,Qitem(Func.read(gs.sheet1,44,2)))
        self.tableMasterCranklInfo.setItem(1,0,Qitem(Func.read(gs.sheet1,47,2)))
        self.tableMasterCranklInfo.setItem(2,0,Qitem(Func.read(gs.sheet1,44,15)))
        self.tableMasterCranklInfo.setItem(3,0,Qitem(Func.read(gs.sheet1,44,9)))
        self.tableMasterCranklInfo.setItem(4,0,Qitem(Func.read(gs.sheet1,47,9)))
        self.tableMasterCranklInfo.setItem(5,0,Qitem(Func.read(gs.sheet1,47,15)))
        self.tableMasterCranklInfo.setItem(6,0,Qitem(Func.read(gs.sheet1,38,2)))
        self.tableMasterCranklInfo.setItem(7,0,Qitem(Func.read(gs.sheet1,38,9)))
        self.tableMasterCranklInfo.setItem(8,0,Qitem(Func.read(gs.sheet1,38,15)))
        self.tableMasterCranklInfo.setItem(9,0,Qitem(Func.read(gs.sheet1,41,2)))
        self.tableMasterCranklInfo.setItem(10,0,Qitem(Func.read(gs.sheet1,41,9)))
        self.tableMasterCranklInfo.setItem(11,0,Qitem(Func.read(gs.sheet1,41,15)))
        self.tableMasterCranklInfo.setItem(0,1,Qitem(Func.read(gs.sheet1,43,6)))
        self.tableMasterCranklInfo.setItem(1,1,Qitem(Func.read(gs.sheet1,46,6)))
        self.tableMasterCranklInfo.setItem(2,1,Qitem(Func.read(gs.sheet1,43,19)))
        self.tableMasterCranklInfo.setItem(3,1,Qitem(Func.read(gs.sheet1,43,13)))
        self.tableMasterCranklInfo.setItem(4,1,Qitem(Func.read(gs.sheet1,46,13)))
        self.tableMasterCranklInfo.setItem(5,1,Qitem(Func.read(gs.sheet1,46,19)))
        self.tableMasterCranklInfo.setItem(6,1,Qitem(Func.read(gs.sheet1,37,6)))
        self.tableMasterCranklInfo.setItem(7,1,Qitem(Func.read(gs.sheet1,37,13)))
        self.tableMasterCranklInfo.setItem(8,1,Qitem(Func.read(gs.sheet1,37,19)))
        self.tableMasterCranklInfo.setItem(9,1,Qitem(Func.read(gs.sheet1,40,6)))
        self.tableMasterCranklInfo.setItem(10,1,Qitem(Func.read(gs.sheet1,40,13)))
        self.tableMasterCranklInfo.setItem(11,1,Qitem(Func.read(gs.sheet1,40,19)))
        self.tableMasterCranklInfo.setItem(0,2,Qitem(Func.read(gs.sheet1, 44,6)))
        self.tableMasterCranklInfo.setItem(1,2,Qitem(Func.read(gs.sheet1, 47,6)))
        self.tableMasterCranklInfo.setItem(2,2,Qitem(Func.read(gs.sheet1,44,19)))
        self.tableMasterCranklInfo.setItem(3,2,Qitem(Func.read(gs.sheet1,44,13)))
        self.tableMasterCranklInfo.setItem(4,2,Qitem(Func.read(gs.sheet1,47,13)))
        self.tableMasterCranklInfo.setItem(5,2,Qitem(Func.read(gs.sheet1,47,19)))
        self.tableMasterCranklInfo.setItem(6,2,Qitem(Func.read(gs.sheet1,38,6)))
        self.tableMasterCranklInfo.setItem(7,2,Qitem(Func.read(gs.sheet1,38,13)))
        self.tableMasterCranklInfo.setItem(8,2,Qitem(Func.read(gs.sheet1,38,19)))
        self.tableMasterCranklInfo.setItem(9,2,Qitem(Func.read(gs.sheet1,41,6)))
        self.tableMasterCranklInfo.setItem(10,2,Qitem(Func.read(gs.sheet1,41,13)))
        self.tableMasterCranklInfo.setItem(11,2,Qitem(Func.read(gs.sheet1,41,19)))

        self.tableMasterCrankInfo2.setItem(0,0,Qitem(Func.read(gs.sheet1,49,4)))
        self.tableMasterCrankInfo2.setItem(0,1,Qitem(Func.read(gs.sheet1,49,9)))
        self.tableMasterCrankInfo2.setItem(0,2,Qitem(Func.read(gs.sheet1,49,17)))

        self.tableSlaveCrankInfo.setItem(0,0,Qitem(Func.read(gs.sheet1,61,2)))
        self.tableSlaveCrankInfo.setItem(1,0,Qitem(Func.read(gs.sheet1,64,2)))
        self.tableSlaveCrankInfo.setItem(2,0,Qitem(Func.read(gs.sheet1,61,15)))
        self.tableSlaveCrankInfo.setItem(3,0,Qitem(Func.read(gs.sheet1,61,9)))
        self.tableSlaveCrankInfo.setItem(4,0,Qitem(Func.read(gs.sheet1,64,9)))
        self.tableSlaveCrankInfo.setItem(5,0,Qitem(Func.read(gs.sheet1,64,15)))
        self.tableSlaveCrankInfo.setItem(6,0,Qitem(Func.read(gs.sheet1,55,2)))
        self.tableSlaveCrankInfo.setItem(7,0,Qitem(Func.read(gs.sheet1,55,9)))
        self.tableSlaveCrankInfo.setItem(8,0,Qitem(Func.read(gs.sheet1,55,15)))
        self.tableSlaveCrankInfo.setItem(9,0,Qitem(Func.read(gs.sheet1,58,2)))
        self.tableSlaveCrankInfo.setItem(10,0,Qitem(Func.read(gs.sheet1,58,9)))
        self.tableSlaveCrankInfo.setItem(11,0,Qitem(Func.read(gs.sheet1,58,15)))
        self.tableSlaveCrankInfo.setItem(0,1,Qitem(Func.read(gs.sheet1,60,6)))
        self.tableSlaveCrankInfo.setItem(1,1,Qitem(Func.read(gs.sheet1,63,6)))
        self.tableSlaveCrankInfo.setItem(2,1,Qitem(Func.read(gs.sheet1,60,19)))
        self.tableSlaveCrankInfo.setItem(3,1,Qitem(Func.read(gs.sheet1,60,13)))
        self.tableSlaveCrankInfo.setItem(4,1,Qitem(Func.read(gs.sheet1,63,13)))
        self.tableSlaveCrankInfo.setItem(5,1,Qitem(Func.read(gs.sheet1,63,19)))
        self.tableSlaveCrankInfo.setItem(6,1,Qitem(Func.read(gs.sheet1,54,6)))
        self.tableSlaveCrankInfo.setItem(7,1,Qitem(Func.read(gs.sheet1,54,13)))
        self.tableSlaveCrankInfo.setItem(8,1,Qitem(Func.read(gs.sheet1,54,19)))
        self.tableSlaveCrankInfo.setItem(9,1,Qitem(Func.read(gs.sheet1,57,6)))
        self.tableSlaveCrankInfo.setItem(10,1,Qitem(Func.read(gs.sheet1,57,13)))
        self.tableSlaveCrankInfo.setItem(11,1,Qitem(Func.read(gs.sheet1,57,19)))
        self.tableSlaveCrankInfo.setItem(0,2,Qitem(Func.read(gs.sheet1,61,6)))
        self.tableSlaveCrankInfo.setItem(1,2,Qitem(Func.read(gs.sheet1,64,6)))
        self.tableSlaveCrankInfo.setItem(2,2,Qitem(Func.read(gs.sheet1,61,19)))
        self.tableSlaveCrankInfo.setItem(3,2,Qitem(Func.read(gs.sheet1,61,13)))
        self.tableSlaveCrankInfo.setItem(4,2,Qitem(Func.read(gs.sheet1,64,13)))
        self.tableSlaveCrankInfo.setItem(5,2,Qitem(Func.read(gs.sheet1,64,19)))
        self.tableSlaveCrankInfo.setItem(6,2,Qitem(Func.read(gs.sheet1,55,6)))
        self.tableSlaveCrankInfo.setItem(7,2,Qitem(Func.read(gs.sheet1,55,13)))
        self.tableSlaveCrankInfo.setItem(8,2,Qitem(Func.read(gs.sheet1,55,19)))
        self.tableSlaveCrankInfo.setItem(9,2,Qitem(Func.read(gs.sheet1,58,6)))
        self.tableSlaveCrankInfo.setItem(10,2,Qitem(Func.read(gs.sheet1,58,13)))
        self.tableSlaveCrankInfo.setItem(11,2,Qitem(Func.read(gs.sheet1,58,19)))

        self.tableSlaveCrankInfo2.setItem(0,0,Qitem(Func.read(gs.sheet1,66,4)))
        self.tableSlaveCrankInfo2.setItem(0,1,Qitem(Func.read(gs.sheet1,66,9)))
        self.tableSlaveCrankInfo2.setItem(0,2,Qitem(Func.read(gs.sheet1,66,17)))
        gs.DriveType = (Func.read(gs.sheet1,12,5))
        gs.MechanicType =  (Func.read(gs.sheet1,12,15))
        gs.Alfa = float(Func.read(gs.sheet1,22,2))
        gs.offsetAngle = float(Func.read(gs.sheet1,22,11))
        gs.CD = float(Func.read(gs.sheet1,47,9))
        gs.ED = float(Func.read(gs.sheet1,47,15))
        gs.CD2 = float(Func.read(gs.sheet1,64,9))
        gs.BC2 = float(Func.read(gs.sheet1,61,9))
        gs.ED2 = float(Func.read(gs.sheet1,64,15))
        gs.Ra_Rb = float(Func.read(gs.sheet1,49,9))
        gs.w2opt = float(Func.read(gs.sheet1,49,17))
        gs.Ra_Rb2 = float(Func.read(gs.sheet1,66 ,9))
        gs.w3actual = float(Func.read(gs.sheet1,66,17))

        self.UpdateOptValue()     

    def LoadClicked(self):
        print('start loading data')
        # loading from excel
        print('start loading')
        gs.Customer    = self.TextCustomer.text()
        gs.ProjectName = self.TextProject.text()
        gs.Department    = self.TextDepartment.text()
        gs.Drawing     = self.TextDrawing.text()
        gs.ValidDate   = self.TextValid.text()
        gs.Name        = self.TextName.text()
        gs.Comment     = self.TextComment.text()
        gs.ParkWhere = self.comboPark.currentText()
        gs.DriveType = (self.comboDrive.currentText())
        gs.MechanicType =  (self.comboMechanic.currentText())
        # ParkPosition= float(self.TextPark.text())

        gs.isClock     = int(self.comboDirection.currentText())
        gs.startFrom     =  self.comboMeasured.currentText() #TBD
        
        gs.APS1   = float(self.tableMotorAngle.item(0,0).text())
        gs.APS2   = float(self.tableMotorAngle.item(0,1).text())
        gs.IPL    = float(self.tableMotorAngle.item(0,2).text())
        gs.UWL    = float(self.tableMotorAngle.item(0,3).text())
        gs.ALP    = float(self.tableMotorAngle.item(0,4).text())
        gs.SP     = float(self.tableMotorAngle.item(0,5).text())
        gs.KMP    = float(self.tableMotorAngle.item(0,6).text())
        
        gs.Alfa   = float(self.tableMotorInfo.item(0,0).text())
        gs.Offset = float(self.tableMotorInfo.item(1,0).text())
        gs.A_X   = float(self.tableMotorInfo.item(2,0).text())
        gs.A_Y   = float(self.tableMotorInfo.item(3,0).text())
        gs.A_Z   = float(self.tableMotorInfo.item(4,0).text())
        gs.Ap_X   = float(self.tableMotorInfo.item(5,0).text())
        gs.Ap_Y   = float(self.tableMotorInfo.item(6,0).text())
        gs.Ap_Z   = float(self.tableMotorInfo.item(7,0).text())
        gs.errorP.Alfa   = float(self.tableMotorInfo.item(0,1).text())
        gs.errorP.Offset = float(self.tableMotorInfo.item(1,1).text())
        gs.errorP.A_X   = float(self.tableMotorInfo.item(2,1).text())
        gs.errorP.A_Y   = float(self.tableMotorInfo.item(3,1).text())
        gs.errorP.A_Z   = float(self.tableMotorInfo.item(4,1).text())
        gs.errorP.Ap_X   = float(self.tableMotorInfo.item(5,1).text())
        gs.errorP.Ap_Y   = float(self.tableMotorInfo.item(6,1).text())
        gs.errorP.Ap_Z   = float(self.tableMotorInfo.item(7,1).text())
        gs.errorN.Alfa   = float(self.tableMotorInfo.item(0,2).text())
        gs.errorN.Offset = float(self.tableMotorInfo.item(1,2).text())
        gs.errorN.A_X   = float(self.tableMotorInfo.item(2,2).text())
        gs.errorN.A_Y   = float(self.tableMotorInfo.item(3,2).text())
        gs.errorN.A_Z   = float(self.tableMotorInfo.item(4,2).text())
        gs.errorN.Ap_X   = float(self.tableMotorInfo.item(5,2).text())
        gs.errorN.Ap_Y   = float(self.tableMotorInfo.item(6,2).text())
        gs.errorN.Ap_Z   = float(self.tableMotorInfo.item(7,2).text())

        gs.Master =  QtWidgets.QComboBox.currentText(self.comboNo2)
        gs.w2          = float(self.TextW2.text())
        gs.w3          = float(self.TextW3.text())
        
        gs.Distance   = float(self.tableMasterCranklInfo.item(0,0).text())
        gs.Delta = float(self.tableMasterCranklInfo.item(1,0).text())
        gs.FE   = float(self.tableMasterCranklInfo.item(2,0).text())
        gs.BC   = float(self.tableMasterCranklInfo.item(3,0).text())
        gs.CD   = float(self.tableMasterCranklInfo.item(4,0).text())
        gs.ED   = float(self.tableMasterCranklInfo.item(5,0).text())
        gs.F_X   = float(self.tableMasterCranklInfo.item(6,0).text())
        gs.F_Y   = float(self.tableMasterCranklInfo.item(7,0).text())
        gs.F_Z   = float(self.tableMasterCranklInfo.item(8,0).text())
        gs.Fp_X   = float(self.tableMasterCranklInfo.item(9,0).text())
        gs.Fp_Y   = float(self.tableMasterCranklInfo.item(10,0).text())
        gs.Fp_Z   = float(self.tableMasterCranklInfo.item(11,0).text())
        gs.errorP.Distance   = float(self.tableMasterCranklInfo.item(0,1).text())
        gs.errorP.Delta = float(self.tableMasterCranklInfo.item(1,1).text())
        gs.errorP.FE   = float(self.tableMasterCranklInfo.item(2,1).text())
        gs.errorP.BC   = float(self.tableMasterCranklInfo.item(3,1).text())
        gs.errorP.CD   = float(self.tableMasterCranklInfo.item(4,1).text())
        gs.errorP.ED   = float(self.tableMasterCranklInfo.item(5,1).text())
        gs.errorP.F_X   = float(self.tableMasterCranklInfo.item(6,1).text())
        gs.errorP.F_Y   = float(self.tableMasterCranklInfo.item(7,1).text())
        gs.errorP.F_Z   = float(self.tableMasterCranklInfo.item(8,1).text())
        gs.errorP.Fp_X   = float(self.tableMasterCranklInfo.item(9,1).text())
        gs.errorP.Fp_Y   = float(self.tableMasterCranklInfo.item(10,1).text())
        gs.errorP.Fp_Z   = float(self.tableMasterCranklInfo.item(11,1).text())
        gs.errorN.Distance   = float(self.tableMasterCranklInfo.item(0,2).text())
        gs.errorN.Delta = float(self.tableMasterCranklInfo.item(1,2).text())
        gs.errorN.FE   = float(self.tableMasterCranklInfo.item(2,2).text())
        gs.errorN.BC   = float(self.tableMasterCranklInfo.item(3,2).text())
        gs.errorN.CD   = float(self.tableMasterCranklInfo.item(4,2).text())
        gs.errorN.ED   = float(self.tableMasterCranklInfo.item(5,2).text())
        gs.errorN.F_X   = float(self.tableMasterCranklInfo.item(6,2).text())
        gs.errorN.F_Y   = float(self.tableMasterCranklInfo.item(7,2).text())
        gs.errorN.F_Z   = float(self.tableMasterCranklInfo.item(8,2).text())
        gs.errorN.Fp_X   = float(self.tableMasterCranklInfo.item(9,2).text())
        gs.errorN.Fp_Y   = float(self.tableMasterCranklInfo.item(10,2).text())
        gs.errorN.Fp_Z   = float(self.tableMasterCranklInfo.item(11,2).text())
       
        gs.KBEW= (self.tableMasterCrankInfo2.item(0,0).text())

        
        gs.Distance2   = float(self.tableSlaveCrankInfo.item(0,0).text())
        gs.Delta2 = float(self.tableSlaveCrankInfo.item(1,0).text())
        gs.FE2   = float(self.tableSlaveCrankInfo.item(2,0).text())
        gs.BC2   = float(self.tableSlaveCrankInfo.item(3,0).text())
        gs.CD2   = float(self.tableSlaveCrankInfo.item(4,0).text())
        gs.ED2   = float(self.tableSlaveCrankInfo.item(5,0).text())
        gs.F_X2   = float(self.tableSlaveCrankInfo.item(6,0).text())
        gs.F_Y2   = float(self.tableSlaveCrankInfo.item(7,0).text())
        gs.F_Z2   = float(self.tableSlaveCrankInfo.item(8,0).text())
        gs.Fp_X2   = float(self.tableSlaveCrankInfo.item(9,0).text())
        gs.Fp_Y2   = float(self.tableSlaveCrankInfo.item(10,0).text())
        gs.Fp_Z2   = float(self.tableSlaveCrankInfo.item(11,0).text())
        gs.errorP.Distance2   = float(self.tableSlaveCrankInfo.item(0,1).text())
        gs.errorP.Delta2 = float(self.tableSlaveCrankInfo.item(1,1).text())
        gs.errorP.FE2   = float(self.tableSlaveCrankInfo.item(2,1).text())
        gs.errorP.BC2   = float(self.tableSlaveCrankInfo.item(3,1).text())
        gs.errorP.CD2   = float(self.tableSlaveCrankInfo.item(4,1).text())
        gs.errorP.ED2   = float(self.tableSlaveCrankInfo.item(5,1).text())
        gs.errorP.F_X2   = float(self.tableSlaveCrankInfo.item(6,1).text())
        gs.errorP.F_Y2   = float(self.tableSlaveCrankInfo.item(7,1).text())
        gs.errorP.F_Z2   = float(self.tableSlaveCrankInfo.item(8,1).text())
        gs.errorP.Fp_X2   = float(self.tableSlaveCrankInfo.item(9,1).text())
        gs.errorP.Fp_Y2   = float(self.tableSlaveCrankInfo.item(10,1).text())
        gs.errorP.Fp_Z2   = float(self.tableSlaveCrankInfo.item(11,1).text())
        gs.errorN.Distance2   = float(self.tableSlaveCrankInfo.item(0,2).text())
        gs.errorN.Delta2 = float(self.tableSlaveCrankInfo.item(1,2).text())
        gs.errorN.FE2   = float(self.tableSlaveCrankInfo.item(2,2).text())
        gs.errorN.BC2   = float(self.tableSlaveCrankInfo.item(3,2).text())
        gs.errorN.CD2   = float(self.tableSlaveCrankInfo.item(4,2).text())
        gs.errorN.ED2   = float(self.tableSlaveCrankInfo.item(5,2).text())
        gs.errorN.F_X2   = float(self.tableSlaveCrankInfo.item(6,2).text())
        gs.errorN.F_Y2   = float(self.tableSlaveCrankInfo.item(7,2).text())
        gs.errorN.F_Z2   = float(self.tableSlaveCrankInfo.item(8,2).text())
        gs.errorN.Fp_X2   = float(self.tableSlaveCrankInfo.item(9,2).text())
        gs.errorN.Fp_Y2   = float(self.tableSlaveCrankInfo.item(10,2).text())
        gs.errorN.Fp_Z2   = float(self.tableSlaveCrankInfo.item(11,2).text())
        gs.KBEW2 = (self.tableSlaveCrankInfo2.item(0,0).text())

        gs.w2Target=gs.w2
        gs.w3Target=gs.w3
        Func.write(gs.sheet1,5,2, gs.Customer)
        Func.write(gs.sheet1,5,9, gs.ProjectName)
        Func.write(gs.sheet1,5,15,gs.Department)
        Func.write(gs.sheet1,7,2, gs.Drawing)
        Func.write(gs.sheet1,7,9, gs.ValidDate)
        Func.write(gs.sheet1,7,15,gs.Name)
        Func.write(gs.sheet1,9,2,  gs.Comment)
        Func.write(gs.sheet1,14,7, gs.ParkWhere)
        Func.write(gs.sheet1,14,15,gs.isClock)
        Func.write(gs.sheet1,12,5, gs.DriveType)
        Func.write(gs.sheet1, 12, 15, gs.MechanicType)
        Func.write(gs.sheet1,16,6,   gs.w2)
        Func.write(gs.sheet1,16,10,  gs.w3)
        Func.write(gs.sheet1,16,17,  gs.startFrom)
        
        Func.write(gs.sheet1,22,2,    gs.Alfa)
        Func.write(gs.sheet1,22,11,   gs.Offset)
        Func.write(gs.sheet1,24,4,    gs.APS1)
        Func.write(gs.sheet1,24,9,    gs.APS2)
        Func.write(gs.sheet1,24,13,   gs.IPL)
        Func.write(gs.sheet1,26,4,    gs.UWL)
        Func.write(gs.sheet1,26,9,    gs.ALP)
        Func.write(gs.sheet1,26,13,   gs.SP)
        Func.write(gs.sheet1,26,18,   gs.KMP)
        Func.write(gs.sheet1,29,2, gs.A_X)
        Func.write(gs.sheet1,29,9, gs.A_Y)
        Func.write(gs.sheet1,29,15,gs.A_Z)
        Func.write(gs.sheet1,32,2, gs.Ap_X)
        Func.write(gs.sheet1,32,9, gs.Ap_Y)
        Func.write(gs.sheet1,32,15,gs.Ap_Z)
                                   
        Func.write(gs.sheet1,38,2,   gs.F_X)
        Func.write(gs.sheet1,38,9,   gs.F_Y)
        Func.write(gs.sheet1,38,15,  gs.F_Z)
        Func.write(gs.sheet1,41,2,   gs.Fp_X)
        Func.write(gs.sheet1,41,9,   gs.Fp_Y)
        Func.write(gs.sheet1,41,15,  gs.Fp_Z)
        Func.write(gs.sheet1,44,2, gs.Distance)
        Func.write(gs.sheet1,44,9, gs.BC)
        Func.write(gs.sheet1,44,15,gs.FE)
        Func.write(gs.sheet1,47,2, gs.Delta)
        Func.write(gs.sheet1,47,9, gs.CD)
        Func.write(gs.sheet1,47,15,gs.ED)
        Func.write(gs.sheet1,49,4,gs.KBEW)
        Func.write(gs.sheet1,49,9,gs.Ra_Rb)
        Func.write(gs.sheet1,49,17,gs.w2opt)

        Func.write(gs.sheet1,55,2,  gs.F_X2)
        Func.write(gs.sheet1,55,9,  gs.F_Y2)
        Func.write(gs.sheet1,55,15, gs.F_Z2)
        Func.write(gs.sheet1,58,2,  gs.Fp_X2)
        Func.write(gs.sheet1,58,9,  gs.Fp_Y2)
        Func.write(gs.sheet1,58,15, gs.Fp_Z2)
        Func.write(gs.sheet1,61,2, gs.Distance2)
        Func.write(gs.sheet1,61,9, gs.BC2)
        Func.write(gs.sheet1,61,15,gs.FE2)
        Func.write(gs.sheet1,64,2, gs.Delta2)
        Func.write(gs.sheet1,64,9, gs.CD2)
        Func.write(gs.sheet1,64,15,gs.ED2)
        Func.write(gs.sheet1,66,4,gs.KBEW2)
        Func.write(gs.sheet1,66,9,gs.Ra_Rb2)
        Func.write(gs.sheet1,66,17,gs.w3actual)

        
        Func.write(gs.sheet1,21,7, gs.errorP.Alfa)
        Func.write(gs.sheet1,22,7, gs.errorN.Alfa)
        Func.write(gs.sheet1,21,16,gs.errorP.Offset)
        Func.write(gs.sheet1,22,16,gs.errorN.Offset)
        Func.write(gs.sheet1,28,6, gs.errorP.A_X)
        Func.write(gs.sheet1,29,6, gs.errorN.A_X)
        Func.write(gs.sheet1,28,13,gs.errorP.A_Y)
        Func.write(gs.sheet1,29,13,gs.errorN.A_Y)
        Func.write(gs.sheet1,28,19,gs.errorP.A_Z)
        Func.write(gs.sheet1,29,19,gs.errorN.A_Z)
        Func.write(gs.sheet1,31,6, gs.errorP.Ap_X)
        Func.write(gs.sheet1,32,6, gs.errorN.Ap_X)
        Func.write(gs.sheet1,31,13,gs.errorP.Ap_Y)
        Func.write(gs.sheet1,32,13,gs.errorN.Ap_Y)
        Func.write(gs.sheet1,31,19,gs.errorP.Ap_Z)
        Func.write(gs.sheet1,32,19,gs.errorN.Ap_Z)
                                
        Func.write(gs.sheet1,37,6, gs.errorP.F_X)
        Func.write(gs.sheet1,38,6, gs.errorN.F_X)
        Func.write(gs.sheet1,37,13,gs.errorP.F_Y)
        Func.write(gs.sheet1,38,13,gs.errorN.F_Y)
        Func.write(gs.sheet1,37,19,gs.errorP.F_Z)
        Func.write(gs.sheet1,38,19,gs.errorN.F_Z)
        Func.write(gs.sheet1,40,6, gs.errorP.Fp_X)
        Func.write(gs.sheet1,41,6, gs.errorN.Fp_X)
        Func.write(gs.sheet1,40,13,gs.errorP.Fp_Y)
        Func.write(gs.sheet1,41,13,gs.errorN.Fp_Y)
        Func.write(gs.sheet1,40,19,gs.errorP.Fp_Z)
        Func.write(gs.sheet1,41,19,gs.errorN.Fp_Z)
        Func.write(gs.sheet1,43,6, gs.errorP.Distance)
        Func.write(gs.sheet1,44,6, gs.errorN.Distance)
        Func.write(gs.sheet1,43,13,gs.errorP.BC)
        Func.write(gs.sheet1,44,13,gs.errorN.BC)
        Func.write(gs.sheet1,43,19,gs.errorP.FE)
        Func.write(gs.sheet1,44,19,gs.errorN.FE)
        Func.write(gs.sheet1,46,6, gs.errorP.Delta)
        Func.write(gs.sheet1,47,6, gs.errorN.Delta)
        Func.write(gs.sheet1,46,13,gs.errorP.CD)
        Func.write(gs.sheet1,47,13,gs.errorN.CD)
        Func.write(gs.sheet1,46,19,gs.errorP.ED)
        Func.write(gs.sheet1,47,19,gs.errorN.ED)
        Func.write(gs.sheet1,54,6, gs.errorP.F_X2)
        Func.write(gs.sheet1,55,6, gs.errorN.F_X2)
        Func.write(gs.sheet1,54,13,gs.errorP.F_Y2)
        Func.write(gs.sheet1,55,13,gs.errorN.F_Y2)
        Func.write(gs.sheet1,54,19,gs.errorP.F_Z2)
        Func.write(gs.sheet1,55,19,gs.errorN.F_Z2)
        Func.write(gs.sheet1,57,6, gs.errorP.Fp_X2)
        Func.write(gs.sheet1,58,6, gs.errorN.Fp_X2)
        Func.write(gs.sheet1,57,13,gs.errorP.Fp_Y2)
        Func.write(gs.sheet1,58,13,gs.errorN.Fp_Y2)
        Func.write(gs.sheet1,57,19,gs.errorP.Fp_Z2)
        Func.write(gs.sheet1,58,19,gs.errorN.Fp_Z2)
        Func.write(gs.sheet1,60,6, gs.errorP.Distance2)
        Func.write(gs.sheet1,61,6, gs.errorN.Distance2)
        Func.write(gs.sheet1,60,13,gs.errorP.BC2)
        Func.write(gs.sheet1,61,13,gs.errorN.BC2)
        Func.write(gs.sheet1,60,19,gs.errorP.FE2)
        Func.write(gs.sheet1,61,19,gs.errorN.FE2)
        Func.write(gs.sheet1,63,6, gs.errorP.Delta2)
        Func.write(gs.sheet1,64,6, gs.errorN.Delta2)
        Func.write(gs.sheet1,63,13,gs.errorP.CD2)
        Func.write(gs.sheet1,64,13,gs.errorN.CD2)
        Func.write(gs.sheet1,63,19,gs.errorP.ED2)
        Func.write(gs.sheet1,64,19,gs.errorN.ED2)

        gs.preTreat()
        Func.writejson(inputjson,inputs,gs)
        gs.wb.save(gs.template_path)
        print('save input configuration')
                # pre-treatment
                
        
        if not gs.BC2:
            logger.warning('please give a correct BC2 value')

        if gs.startFrom=='EPS':
            gs.startAngle=gs.Offset
        elif gs.startFrom=="APS1":
            gs.startAngle=gs.Offset+gs.APS1
        elif gs.startFrom=='Park':
            logger.warning('TO be defined what park position mean')
        else:
            logger.warning('pls configure where measured from')
            
        gs.Alfa=np.sign(gs.isClock)*gs.Alfa

        print('load completed')

    
def get_file_name(dir, file_extension):
    f_list = os.listdir(dir)

    result_list = []

    for file_name in f_list:
        if os.path.splitext(file_name)[1] == file_extension:
            result_list.append(os.path.join(dir, file_name))
    return result_list

def get_mac_address(): 
    mac=uuid.UUID(int = uuid.getnode()).hex[-12:] 
    return mac
    #return ":".join([mac[e:e+2] for e in range(0,11,2)])


def createLicense(mac,deadLine):
    name = mac+'_license.lic'

    str1 = 'Name:'+mac
    #str1 = str1.encode(encoding="utf-8")
    strb2 = base64.encodestring(bytes(mac,'utf8'))
    str2 = strb2.decode()
    strb3 = base64.encodestring(bytes(deadLine,'utf8'))
    str3 = strb3.decode()
    #string = str(str1+str2+str3)
    fp = open(name, 'w')
    strList = [str1,str2,str3]
    for item in strList:
        fp.writelines(item)
        fp.write('\t')
        #fp.write('\n')
    fp.close

def checkLicense():
    second = time.time()
    cwd = os.getcwd()
    mac = get_mac_address()
    licenses = get_file_name(cwd, '.lic')[0]

    file_object = open(licenses) 
    try:
        file_context = file_object.read() 

    finally:
        file_object.close()
    contents=file_context.split('\t')
    [name, macCode,deadLineCode] = contents[0:3]
    deadLineBytes = base64.decodestring(deadLineCode.encode())
    macBytes = base64.decodestring(macCode.encode())
    macL = macBytes.decode()
    deadLine = float(deadLineBytes.decode())
    if (second<=deadLine) :# ok (mac == macL)&
        print('ok')
        return 1
    else:
        print ('pls update license')
        return 0

# generate license
# mac = get_mac_address()
# deadLine = str(time.mktime(time.strptime("2019-01-01","%Y-%m-%d")))  #deadline date
# createLicense(mac,deadLine)

# interpret license

# m = hashlib.md5()
# m.update(mac.encode('utf8'))
# macMD5 = m.hexdigest()



if __name__ == "__main__":

    gs = parameters.GS()
    inputs= parameters.Inputs()
    outputs= parameters.Outputs()
    logger = Func.logSetting()

    timeStamp=int(time.time())
    timeArray=time.localtime(timeStamp)
    gs.styleTime=time.strftime("%Y%m%d_%H%M%S",timeArray)
    gs.loadtemplate() #template is fixed file
    flag = checkLicense()
    # inputjson = sys.argv[1]
    # outputjson = sys.argv[2]
    inputjson='input.json'
    outputjson = 'output.json'
    if len(sys.argv)>1:
        backend = True # flag if run without gui
    else:
        backend = False
    if backend:
        app = QtWidgets.QApplication(sys.argv)
        Func.LoadJson(inputjson,inputs,gs)
        gs.preTreat()

        mainWindow = MainWindow()
        mainWindow.WindowOpt.OptClicked()
        mainWindow.WindowAlpha.alphaTable()
        mainWindow.WindowExtreme.ExtremeTable()
        
        Func.writejson(outputjson,outputs,gs)
    

    else:
        app = QtWidgets.QApplication(sys.argv)
        mainWindow = MainWindow()
        mainWindow.show()
    



    sys.exit(app.exec_())
    sys.exit(0)
# =============================================================================



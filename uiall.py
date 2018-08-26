# -*- coding: utf-8 -*-
"""
Created on Mon Jul 23 19:37:08 2018

@author: Administrator
"""
import logging
import math
import os
import os.path
import sys
import time

import matplotlib
import matplotlib.animation as animation
import matplotlib.pyplot as plt
import mpl_toolkits.mplot3d.axes3d as p3
import numpy as np
import openpyxl as xl
import pandas as pd
import sympy as sp
from mpl_toolkits.mplot3d import proj3d
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QBrush, QColor
from PyQt5.QtWidgets import QTableWidgetItem as Qitem

import Func
import gs
import pyomo.environ as pe
import scipy.optimize as op


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("Kinematics of Calculation")
        MainWindow.resize(1013, 813)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.widget = QtWidgets.QWidget(self.centralwidget)
        self.widget.setGeometry(QtCore.QRect(2, 660, 841, 25))
        self.widget.setObjectName("widget")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.widget)
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.LoadButton = QtWidgets.QPushButton(self.widget)
        font = QtGui.QFont()
        font.setBold(True)
        font.setUnderline(False)
        font.setWeight(75)
        self.LoadButton.setFont(font)
        self.LoadButton.setCheckable(False)
        self.LoadButton.setAutoDefault(True)
        self.LoadButton.setDefault(True)
        self.LoadButton.setFlat(False)
        self.LoadButton.setObjectName("LoadButton")
        self.horizontalLayout.addWidget(self.LoadButton)
        self.CheckButton = QtWidgets.QPushButton(self.widget)
        font = QtGui.QFont()
        font.setBold(True)
        font.setUnderline(False)
        font.setWeight(75)
        self.CheckButton.setFont(font)
        self.CheckButton.setCheckable(False)
        self.CheckButton.setAutoDefault(True)
        self.CheckButton.setDefault(True)
        self.CheckButton.setFlat(False)
        self.CheckButton.setObjectName("CheckButton")
        self.horizontalLayout.addWidget(self.CheckButton)
        self.OutputButton = QtWidgets.QPushButton(self.widget)
        font = QtGui.QFont()
        font.setBold(True)
        font.setUnderline(False)
        font.setWeight(75)
        self.OutputButton.setFont(font)
        self.OutputButton.setCheckable(False)
        self.OutputButton.setAutoDefault(True)
        self.OutputButton.setDefault(True)
        self.OutputButton.setFlat(False)
        self.OutputButton.setObjectName("OutputButton")
        self.horizontalLayout.addWidget(self.OutputButton)
        self.ToleranceButton = QtWidgets.QPushButton(self.widget)
        font = QtGui.QFont()
        font.setBold(True)
        font.setUnderline(False)
        font.setWeight(75)
        self.ToleranceButton.setFont(font)
        self.ToleranceButton.setCheckable(False)
        self.ToleranceButton.setAutoDefault(True)
        self.ToleranceButton.setDefault(True)
        self.ToleranceButton.setFlat(False)
        self.ToleranceButton.setObjectName("ToleranceButton")
        self.horizontalLayout.addWidget(self.ToleranceButton)
        self.SaveButton = QtWidgets.QPushButton(self.widget)
        font = QtGui.QFont()
        font.setBold(True)
        font.setUnderline(False)
        font.setWeight(75)
        self.SaveButton.setFont(font)
        self.SaveButton.setCheckable(False)
        self.SaveButton.setAutoDefault(True)
        self.SaveButton.setDefault(True)
        self.SaveButton.setFlat(False)
        self.SaveButton.setObjectName("SaveButton")
        self.horizontalLayout.addWidget(self.SaveButton)
        #MainWindow.setCentralWidget(self.centralwidget)
        #self.statusbar = QtWidgets.QStatusBar(MainWindow)
        #self.statusbar.setObjectName("statusbar")
        #MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.LoadButton.setText(_translate("MainWindow", "Apply"))
        self.CheckButton.setText(_translate("MainWindow", "Check"))
        self.OutputButton.setText(_translate("MainWindow", "Output"))
        self.ToleranceButton.setText(_translate("MainWindow", "Tolerance"))
        self.SaveButton.setText(_translate("MainWindow", "Update to Catia"))

class Ui_Input(object):
    def setupUi(self, Input):
        Input.setObjectName("Input")
        Input.resize(1025, 921)
        self.centralwidget = QtWidgets.QWidget(Input)
        self.centralwidget.setObjectName("centralwidget")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(0, 0, 461, 31))
        self.label.setObjectName("label")
        self.widget = QtWidgets.QWidget(self.centralwidget)
        self.widget.setGeometry(QtCore.QRect(0, 40, 441, 141))
        self.widget.setObjectName("widget")
        self.label_2 = QtWidgets.QLabel(self.widget)
        self.label_2.setGeometry(QtCore.QRect(21, 1, 64, 18))
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(self.widget)
        self.label_3.setGeometry(QtCore.QRect(160, 1, 56, 18))
        self.label_3.setObjectName("label_3")
        self.label_4 = QtWidgets.QLabel(self.widget)
        self.label_4.setGeometry(QtCore.QRect(299, 1, 80, 18))
        self.label_4.setObjectName("label_4")
        self.label_5 = QtWidgets.QLabel(self.widget)
        self.label_5.setGeometry(QtCore.QRect(21, 51, 88, 18))
        self.label_5.setObjectName("label_5")
        self.label_6 = QtWidgets.QLabel(self.widget)
        self.label_6.setGeometry(QtCore.QRect(160, 51, 80, 62))
        self.label_6.setObjectName("label_6")
        self.label_7 = QtWidgets.QLabel(self.widget)
        self.label_7.setGeometry(QtCore.QRect(299, 51, 32, 18))
        self.label_7.setObjectName("label_7")
        self.label_8 = QtWidgets.QLabel(self.widget)
        self.label_8.setGeometry(QtCore.QRect(20, 90, 101, 21))
        self.label_8.setObjectName("label_8")
        self.TextCustomer = QtWidgets.QLineEdit(self.widget)
        self.TextCustomer.setGeometry(QtCore.QRect(21, 25, 133, 20))
        self.TextCustomer.setObjectName("TextCustomer")
        self.TextProject = QtWidgets.QLineEdit(self.widget)
        self.TextProject.setGeometry(QtCore.QRect(160, 25, 133, 20))
        self.TextProject.setObjectName("TextProject")
        self.TextDepartment = QtWidgets.QLineEdit(self.widget)
        self.TextDepartment.setGeometry(QtCore.QRect(299, 25, 133, 20))
        self.TextDepartment.setObjectName("TextDepartment")
        self.TextDrawing = QtWidgets.QLineEdit(self.widget)
        self.TextDrawing.setGeometry(QtCore.QRect(20, 70, 133, 20))
        self.TextDrawing.setObjectName("TextDrawing")
        self.TextValid = QtWidgets.QLineEdit(self.widget)
        self.TextValid.setGeometry(QtCore.QRect(160, 70, 133, 20))
        self.TextValid.setObjectName("TextValid")
        self.TextName = QtWidgets.QLineEdit(self.widget)
        self.TextName.setGeometry(QtCore.QRect(300, 70, 133, 20))
        self.TextName.setObjectName("TextName")
        self.TextComment = QtWidgets.QLineEdit(self.widget)
        self.TextComment.setGeometry(QtCore.QRect(20, 110, 411, 20))
        self.TextComment.setObjectName("TextComment")
        self.frame_2 = QtWidgets.QFrame(self.centralwidget)
        self.frame_2.setGeometry(QtCore.QRect(0, 180, 441, 91))
        self.frame_2.setStyleSheet("border：2px solid red;\n"
"background-color: yellow;")
        self.frame_2.setFrameShape(QtWidgets.QFrame.Box)
        self.frame_2.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_2.setObjectName("frame_2")
        self.label_14 = QtWidgets.QLabel(self.frame_2)
        self.label_14.setGeometry(QtCore.QRect(180, 70, 16, 16))
        self.label_14.setObjectName("label_14")
        self.TextW2 = QtWidgets.QLineEdit(self.frame_2)
        self.TextW2.setGeometry(QtCore.QRect(110, 70, 51, 20))
        self.TextW2.setObjectName("TextW2")
        self.TextW3 = QtWidgets.QLineEdit(self.frame_2)
        self.TextW3.setGeometry(QtCore.QRect(210, 70, 51, 20))
        self.TextW3.setObjectName("TextW3")
        self.label_15 = QtWidgets.QLabel(self.frame_2)
        self.label_15.setGeometry(QtCore.QRect(280, 70, 78, 16))
        self.label_15.setObjectName("label_15")
        self.comboMeasured = QtWidgets.QComboBox(self.frame_2)
        self.comboMeasured.setGeometry(QtCore.QRect(380, 70, 50, 20))
        self.comboMeasured.setObjectName("comboMeasured")
        self.comboMeasured.addItem("")
        self.comboMeasured.addItem("")
        self.comboMeasured.addItem("")
        self.label_12 = QtWidgets.QLabel(self.frame_2)
        self.label_12.setGeometry(QtCore.QRect(11, 71, 66, 16))
        self.label_12.setObjectName("label_12")
        self.label_13 = QtWidgets.QLabel(self.frame_2)
        self.label_13.setGeometry(QtCore.QRect(83, 71, 16, 16))
        self.label_13.setObjectName("label_13")
        self.layoutWidget = QtWidgets.QWidget(self.frame_2)
        self.layoutWidget.setGeometry(QtCore.QRect(20, 10, 401, 22))
        self.layoutWidget.setObjectName("layoutWidget")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.layoutWidget)
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label_11 = QtWidgets.QLabel(self.layoutWidget)
        self.label_11.setObjectName("label_11")
        self.horizontalLayout.addWidget(self.label_11)
        self.comboDrive = QtWidgets.QComboBox(self.layoutWidget)
        self.comboDrive.setObjectName("comboDrive")
        self.comboDrive.addItem("")
        self.comboDrive.addItem("")
        self.horizontalLayout.addWidget(self.comboDrive)
        self.label_19 = QtWidgets.QLabel(self.layoutWidget)
        self.label_19.setObjectName("label_19")
        self.horizontalLayout.addWidget(self.label_19)
        self.comboMechanic = QtWidgets.QComboBox(self.layoutWidget)
        self.comboMechanic.setObjectName("comboMechanic")
        self.comboMechanic.addItem("")
        self.comboMechanic.addItem("")
        self.comboMechanic.addItem("")
        self.horizontalLayout.addWidget(self.comboMechanic)
        self.label_9 = QtWidgets.QLabel(self.frame_2)
        self.label_9.setGeometry(QtCore.QRect(15, 40, 101, 20))
        self.label_9.setObjectName("label_9")
        self.comboDirection = QtWidgets.QComboBox(self.frame_2)
        self.comboDirection.setGeometry(QtCore.QRect(366, 40, 51, 20))
        self.comboDirection.setObjectName("comboDirection")
        self.comboDirection.addItem("")
        self.comboDirection.addItem("")
        self.label_10 = QtWidgets.QLabel(self.frame_2)
        self.label_10.setGeometry(QtCore.QRect(230, 40, 126, 16))
        self.label_10.setObjectName("label_10")
        self.comboPark = QtWidgets.QComboBox(self.frame_2)
        self.comboPark.setGeometry(QtCore.QRect(138, 40, 81, 20))
        self.comboPark.setObjectName("comboPark")
        self.comboPark.addItem("")
        self.comboPark.addItem("")
        self.label_9.raise_()
        self.comboDirection.raise_()
        self.label_10.raise_()
        self.comboPark.raise_()
        self.label_14.raise_()
        self.TextW2.raise_()
        self.TextW3.raise_()
        self.label_15.raise_()
        self.comboMeasured.raise_()
        self.label_12.raise_()
        self.label_13.raise_()
        self.layoutWidget.raise_()
        self.frame_3 = QtWidgets.QFrame(self.centralwidget)
        self.frame_3.setGeometry(QtCore.QRect(0, 270, 441, 361))
        self.frame_3.setFrameShape(QtWidgets.QFrame.Box)
        self.frame_3.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_3.setObjectName("frame_3")
        self.tableMasterCranklInfo = QtWidgets.QTableWidget(self.frame_3)
        self.tableMasterCranklInfo.setGeometry(QtCore.QRect(0, 30, 461, 321))
        self.tableMasterCranklInfo.setObjectName("tableMasterCranklInfo")
        self.tableMasterCranklInfo.setColumnCount(3)
        self.tableMasterCranklInfo.setRowCount(12)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setVerticalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setVerticalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setVerticalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setVerticalHeaderItem(4, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setVerticalHeaderItem(5, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setVerticalHeaderItem(6, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setVerticalHeaderItem(7, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setVerticalHeaderItem(8, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setVerticalHeaderItem(9, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setVerticalHeaderItem(10, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setVerticalHeaderItem(11, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(0, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(0, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(0, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(1, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(1, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(1, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(2, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(2, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(2, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(3, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(3, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(3, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(4, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(4, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(4, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(5, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(5, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(5, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(6, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(6, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(6, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(7, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(7, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(7, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(8, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(8, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(8, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(9, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(9, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(9, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(10, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(10, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(10, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(11, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(11, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCranklInfo.setItem(11, 2, item)
        self.tableMasterCranklInfo.horizontalHeader().setDefaultSectionSize(130)
        self.tableMasterCranklInfo.verticalHeader().setDefaultSectionSize(20)
        self.tableMasterCrankInfo2 = QtWidgets.QTableWidget(self.frame_3)
        self.tableMasterCrankInfo2.setGeometry(QtCore.QRect(0, 300, 461, 83))
        self.tableMasterCrankInfo2.setObjectName("tableMasterCrankInfo2")
        self.tableMasterCrankInfo2.setColumnCount(3)
        self.tableMasterCrankInfo2.setRowCount(1)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCrankInfo2.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCrankInfo2.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCrankInfo2.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCrankInfo2.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCrankInfo2.setItem(0, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCrankInfo2.setItem(0, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMasterCrankInfo2.setItem(0, 2, item)
        self.tableMasterCrankInfo2.horizontalHeader().setDefaultSectionSize(130)
        self.tableMasterCrankInfo2.verticalHeader().setDefaultSectionSize(20)
        self.layoutWidget1 = QtWidgets.QWidget(self.frame_3)
        self.layoutWidget1.setGeometry(QtCore.QRect(10, 10, 294, 22))
        self.layoutWidget1.setObjectName("layoutWidget1")
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout(self.layoutWidget1)
        self.horizontalLayout_3.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.label_17 = QtWidgets.QLabel(self.layoutWidget1)
        self.label_17.setObjectName("label_17")
        self.horizontalLayout_3.addWidget(self.label_17)
        self.comboBoxNo2 = QtWidgets.QComboBox(self.layoutWidget1)
        self.comboBoxNo2.setObjectName("comboBoxNo2")
        self.comboBoxNo2.addItem("")
        self.comboBoxNo2.addItem("")
        self.horizontalLayout_3.addWidget(self.comboBoxNo2)
        self.frame = QtWidgets.QFrame(self.centralwidget)
        self.frame.setGeometry(QtCore.QRect(460, 0, 461, 271))
        self.frame.setFrameShape(QtWidgets.QFrame.Box)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame.setObjectName("frame")
        self.label_16 = QtWidgets.QLabel(self.frame)
        self.label_16.setGeometry(QtCore.QRect(10, 10, 171, 16))
        self.label_16.setObjectName("label_16")
        self.tableMotorInfo = QtWidgets.QTableWidget(self.frame)
        self.tableMotorInfo.setGeometry(QtCore.QRect(0, 80, 431, 191))
        self.tableMotorInfo.setObjectName("tableMotorInfo")
        self.tableMotorInfo.setColumnCount(3)
        self.tableMotorInfo.setRowCount(8)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setVerticalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setVerticalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setVerticalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setVerticalHeaderItem(4, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setVerticalHeaderItem(5, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setVerticalHeaderItem(6, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setVerticalHeaderItem(7, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setItem(0, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setItem(0, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setItem(0, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setItem(1, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setItem(1, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setItem(1, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setItem(2, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setItem(2, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setItem(2, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setItem(3, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setItem(3, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setItem(3, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setItem(4, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setItem(4, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setItem(4, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setItem(5, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setItem(5, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setItem(5, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setItem(6, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setItem(6, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setItem(6, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setItem(7, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setItem(7, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorInfo.setItem(7, 2, item)
        self.tableMotorInfo.verticalHeader().setDefaultSectionSize(20)
        self.tableMotorAngle = QtWidgets.QTableWidget(self.frame)
        self.tableMotorAngle.setGeometry(QtCore.QRect(0, 30, 431, 51))
        self.tableMotorAngle.setObjectName("tableMotorAngle")
        self.tableMotorAngle.setColumnCount(7)
        self.tableMotorAngle.setRowCount(1)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorAngle.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorAngle.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorAngle.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorAngle.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorAngle.setHorizontalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorAngle.setHorizontalHeaderItem(4, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorAngle.setHorizontalHeaderItem(5, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorAngle.setHorizontalHeaderItem(6, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorAngle.setItem(0, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorAngle.setItem(0, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorAngle.setItem(0, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorAngle.setItem(0, 3, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorAngle.setItem(0, 4, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorAngle.setItem(0, 5, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableMotorAngle.setItem(0, 6, item)
        self.tableMotorAngle.horizontalHeader().setDefaultSectionSize(55)
        self.tableMotorAngle.verticalHeader().setDefaultSectionSize(20)
        self.frame_4 = QtWidgets.QFrame(self.centralwidget)
        self.frame_4.setGeometry(QtCore.QRect(460, 270, 461, 361))
        self.frame_4.setFrameShape(QtWidgets.QFrame.Box)
        self.frame_4.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_4.setObjectName("frame_4")
        self.tableSlaveCrankInfo = QtWidgets.QTableWidget(self.frame_4)
        self.tableSlaveCrankInfo.setGeometry(QtCore.QRect(0, 30, 461, 331))
        self.tableSlaveCrankInfo.setObjectName("tableSlaveCrankInfo")
        self.tableSlaveCrankInfo.setColumnCount(3)
        self.tableSlaveCrankInfo.setRowCount(12)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setVerticalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setVerticalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setVerticalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setVerticalHeaderItem(4, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setVerticalHeaderItem(5, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setVerticalHeaderItem(6, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setVerticalHeaderItem(7, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setVerticalHeaderItem(8, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setVerticalHeaderItem(9, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setVerticalHeaderItem(10, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setVerticalHeaderItem(11, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(0, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(0, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(0, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(1, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(1, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(1, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(2, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(2, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(2, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(3, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(3, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(3, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(4, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(4, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(4, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(5, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(5, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(5, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(6, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(6, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(6, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(7, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(7, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(7, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(8, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(8, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(8, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(9, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(9, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(9, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(10, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(10, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(10, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(11, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(11, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo.setItem(11, 2, item)
        self.tableSlaveCrankInfo.horizontalHeader().setDefaultSectionSize(130)
        self.tableSlaveCrankInfo.verticalHeader().setDefaultSectionSize(20)
        self.tableSlaveCrankInfo2 = QtWidgets.QTableWidget(self.frame_4)
        self.tableSlaveCrankInfo2.setGeometry(QtCore.QRect(0, 300, 461, 61))
        self.tableSlaveCrankInfo2.setObjectName("tableSlaveCrankInfo2")
        self.tableSlaveCrankInfo2.setColumnCount(3)
        self.tableSlaveCrankInfo2.setRowCount(1)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo2.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo2.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo2.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo2.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo2.setItem(0, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo2.setItem(0, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableSlaveCrankInfo2.setItem(0, 2, item)
        self.tableSlaveCrankInfo2.horizontalHeader().setDefaultSectionSize(130)
        self.tableSlaveCrankInfo2.verticalHeader().setDefaultSectionSize(20)
        self.widget1 = QtWidgets.QWidget(self.frame_4)
        self.widget1.setGeometry(QtCore.QRect(11, 10, 271, 22))
        self.widget1.setObjectName("widget1")
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout(self.widget1)
        self.horizontalLayout_4.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.label_18 = QtWidgets.QLabel(self.widget1)
        self.label_18.setObjectName("label_18")
        self.horizontalLayout_4.addWidget(self.label_18)
        self.comboBoxNo3 = QtWidgets.QComboBox(self.widget1)
        self.comboBoxNo3.setObjectName("comboBoxNo3")
        self.comboBoxNo3.addItem("")
        self.comboBoxNo3.addItem("")
        self.horizontalLayout_4.addWidget(self.comboBoxNo3)
        self.layoutWidget2 = QtWidgets.QWidget(self.centralwidget)
        self.layoutWidget2.setGeometry(QtCore.QRect(0, 640, 911, 25))
        self.layoutWidget2.setObjectName("layoutWidget2")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout(self.layoutWidget2)
        self.horizontalLayout_2.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.LoadButton = QtWidgets.QPushButton(self.layoutWidget2)
        font = QtGui.QFont()
        font.setBold(True)
        font.setUnderline(False)
        font.setWeight(75)
        self.LoadButton.setFont(font)
        self.LoadButton.setCheckable(False)
        self.LoadButton.setAutoDefault(True)
        self.LoadButton.setDefault(True)
        self.LoadButton.setFlat(False)
        self.LoadButton.setObjectName("LoadButton")
        self.horizontalLayout_2.addWidget(self.LoadButton)
        self.OutputButton = QtWidgets.QPushButton(self.layoutWidget2)
        font = QtGui.QFont()
        font.setBold(True)
        font.setUnderline(False)
        font.setWeight(75)
        self.OutputButton.setFont(font)
        self.OutputButton.setCheckable(False)
        self.OutputButton.setAutoDefault(True)
        self.OutputButton.setDefault(True)
        self.OutputButton.setFlat(False)
        self.OutputButton.setObjectName("OutputButton")
        self.horizontalLayout_2.addWidget(self.OutputButton)
        self.ToleranceButton = QtWidgets.QPushButton(self.layoutWidget2)
        font = QtGui.QFont()
        font.setBold(True)
        font.setUnderline(False)
        font.setWeight(75)
        self.ToleranceButton.setFont(font)
        self.ToleranceButton.setCheckable(False)
        self.ToleranceButton.setAutoDefault(True)
        self.ToleranceButton.setDefault(True)
        self.ToleranceButton.setFlat(False)
        self.ToleranceButton.setObjectName("ToleranceButton")
        self.horizontalLayout_2.addWidget(self.ToleranceButton)
        self.SaveButton = QtWidgets.QPushButton(self.layoutWidget2)
        font = QtGui.QFont()
        font.setBold(True)
        font.setUnderline(False)
        font.setWeight(75)
        self.SaveButton.setFont(font)
        self.SaveButton.setCheckable(False)
        self.SaveButton.setAutoDefault(True)
        self.SaveButton.setDefault(True)
        self.SaveButton.setFlat(False)
        self.SaveButton.setObjectName("SaveButton")
        self.horizontalLayout_2.addWidget(self.SaveButton)
        Input.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(Input)
        self.statusbar.setObjectName("statusbar")
        Input.setStatusBar(self.statusbar)

        self.retranslateUi(Input)
        QtCore.QMetaObject.connectSlotsByName(Input)

    def retranslateUi(self, Input):
        _translate = QtCore.QCoreApplication.translate
        Input.setWindowTitle(_translate("Input", "MainWindow"))
        self.label.setText(_translate("Input", "<html><head/><body><p><span style=\" font-size:12pt; font-weight:600;\">Kinematics calculation of wiper linkage</span></p></body></html>"))
        self.label_2.setText(_translate("Input", "<html><head/><body><p><span style=\" font-size:12pt;\">Customer</span></p></body></html>"))
        self.label_3.setText(_translate("Input", "<html><head/><body><p><span style=\" font-size:12pt;\">Project</span></p></body></html>"))
        self.label_4.setText(_translate("Input", "<html><head/><body><p><span style=\" font-size:12pt;\">Department</span></p></body></html>"))
        self.label_5.setText(_translate("Input", "<html><head/><body><p><span style=\" font-size:12pt;\">Drawing No.</span></p></body></html>"))
        self.label_6.setText(_translate("Input", "<html><head/><body><p><span style=\" font-size:12pt;\">Valid Date</span></p><p><span style=\" font-size:12pt;\"><br/></span></p></body></html>"))
        self.label_7.setText(_translate("Input", "<html><head/><body><p><span style=\" font-size:12pt;\">Name</span></p></body></html>"))
        self.label_8.setText(_translate("Input", "<html><head/><body><p><span style=\" font-size:12pt;\">Comment</span></p></body></html>"))
        self.TextCustomer.setText(_translate("Input", "BMW"))
        self.TextProject.setText(_translate("Input", "C1UG"))
        self.TextDepartment.setText(_translate("Input", "Eng"))
        self.TextDrawing.setText(_translate("Input", "12"))
        self.TextValid.setText(_translate("Input", "2018/07/26"))
        self.TextName.setText(_translate("Input", "Qian"))
        self.TextComment.setText(_translate("Input", "BMW"))
        self.label_14.setText(_translate("Input", "<html><head/><body><p>w3</p></body></html>"))
        self.label_15.setText(_translate("Input", "<html><head/><body><p>measured from</p></body></html>"))
        self.comboMeasured.setItemText(0, _translate("Input", "APS1"))
        self.comboMeasured.setItemText(1, _translate("Input", "EPS"))
        self.comboMeasured.setItemText(2, _translate("Input", "Park"))
        self.label_12.setText(_translate("Input", "<html><head/><body><p>Wipe angle:</p></body></html>"))
        self.label_13.setText(_translate("Input", "<html><head/><body><p>w2</p></body></html>"))
        self.label_11.setText(_translate("Input", "<html><head/><body><p>Drive type</p></body></html>"))
        self.comboDrive.setItemText(0, _translate("Input", "Standard"))
        self.comboDrive.setItemText(1, _translate("Input", "Reversing"))
        self.label_19.setText(_translate("Input", "Mechanism Type"))
        self.comboMechanic.setItemText(0, _translate("Input", "Center"))
        self.comboMechanic.setItemText(1, _translate("Input", "Serial DS-PS"))
        self.comboMechanic.setItemText(2, _translate("Input", "Serial PS-DS"))
        self.label_9.setText(_translate("Input", "<html><head/><body><p>Park position at</p></body></html>"))
        self.comboDirection.setItemText(0, _translate("Input", "1"))
        self.comboDirection.setItemText(1, _translate("Input", "-1"))
        self.label_10.setText(_translate("Input", "<html><head/><body><p>Direction of rotation</p><p><br/></p></body></html>"))
        self.comboPark.setItemText(0, _translate("Input", "-y"))
        self.comboPark.setItemText(1, _translate("Input", "+y"))
        item = self.tableMasterCranklInfo.verticalHeaderItem(0)
        item.setText(_translate("Input", "D"))
        item = self.tableMasterCranklInfo.verticalHeaderItem(1)
        item.setText(_translate("Input", "Delta"))
        item = self.tableMasterCranklInfo.verticalHeaderItem(2)
        item.setText(_translate("Input", "FE"))
        item = self.tableMasterCranklInfo.verticalHeaderItem(3)
        item.setText(_translate("Input", "BC"))
        item = self.tableMasterCranklInfo.verticalHeaderItem(4)
        item.setText(_translate("Input", "CD"))
        item = self.tableMasterCranklInfo.verticalHeaderItem(5)
        item.setText(_translate("Input", "ED"))
        item = self.tableMasterCranklInfo.verticalHeaderItem(6)
        item.setText(_translate("Input", "F-X"))
        item = self.tableMasterCranklInfo.verticalHeaderItem(7)
        item.setText(_translate("Input", "F-Y"))
        item = self.tableMasterCranklInfo.verticalHeaderItem(8)
        item.setText(_translate("Input", "F-Z"))
        item = self.tableMasterCranklInfo.verticalHeaderItem(9)
        item.setText(_translate("Input", "F\'-X"))
        item = self.tableMasterCranklInfo.verticalHeaderItem(10)
        item.setText(_translate("Input", "F\'-Y"))
        item = self.tableMasterCranklInfo.verticalHeaderItem(11)
        item.setText(_translate("Input", "F\'-Z"))
        item = self.tableMasterCranklInfo.horizontalHeaderItem(0)
        item.setText(_translate("Input", "Value"))
        item = self.tableMasterCranklInfo.horizontalHeaderItem(1)
        item.setText(_translate("Input", "+"))
        item = self.tableMasterCranklInfo.horizontalHeaderItem(2)
        item.setText(_translate("Input", "-"))
        __sortingEnabled = self.tableMasterCranklInfo.isSortingEnabled()
        self.tableMasterCranklInfo.setSortingEnabled(False)
        item = self.tableMasterCranklInfo.item(0, 0)
        item.setText(_translate("Input", "0"))
        item = self.tableMasterCranklInfo.item(0, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableMasterCranklInfo.item(0, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableMasterCranklInfo.item(1, 0)
        item.setText(_translate("Input", "0"))
        item = self.tableMasterCranklInfo.item(1, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableMasterCranklInfo.item(1, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableMasterCranklInfo.item(2, 0)
        item.setText(_translate("Input", "0.2"))
        item = self.tableMasterCranklInfo.item(2, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableMasterCranklInfo.item(2, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableMasterCranklInfo.item(3, 0)
        item.setText(_translate("Input", "50"))
        item = self.tableMasterCranklInfo.item(3, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableMasterCranklInfo.item(3, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableMasterCranklInfo.item(4, 0)
        item.setText(_translate("Input", "250"))
        item = self.tableMasterCranklInfo.item(4, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableMasterCranklInfo.item(4, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableMasterCranklInfo.item(5, 0)
        item.setText(_translate("Input", "71"))
        item = self.tableMasterCranklInfo.item(5, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableMasterCranklInfo.item(5, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableMasterCranklInfo.item(6, 0)
        item.setText(_translate("Input", "1826.507"))
        item = self.tableMasterCranklInfo.item(6, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableMasterCranklInfo.item(6, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableMasterCranklInfo.item(7, 0)
        item.setText(_translate("Input", "-158.845"))
        item = self.tableMasterCranklInfo.item(7, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableMasterCranklInfo.item(7, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableMasterCranklInfo.item(8, 0)
        item.setText(_translate("Input", "993.065"))
        item = self.tableMasterCranklInfo.item(8, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableMasterCranklInfo.item(8, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableMasterCranklInfo.item(9, 0)
        item.setText(_translate("Input", "1836.793"))
        item = self.tableMasterCranklInfo.item(9, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableMasterCranklInfo.item(9, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableMasterCranklInfo.item(10, 0)
        item.setText(_translate("Input", "-158.854"))
        item = self.tableMasterCranklInfo.item(10, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableMasterCranklInfo.item(10, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableMasterCranklInfo.item(11, 0)
        item.setText(_translate("Input", "973.617"))
        item = self.tableMasterCranklInfo.item(11, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableMasterCranklInfo.item(11, 2)
        item.setText(_translate("Input", "0"))
        self.tableMasterCranklInfo.setSortingEnabled(__sortingEnabled)
        item = self.tableMasterCrankInfo2.verticalHeaderItem(0)
        item.setText(_translate("Input", "Value"))
        item = self.tableMasterCrankInfo2.horizontalHeaderItem(0)
        item.setText(_translate("Input", "KBEW"))
        item = self.tableMasterCrankInfo2.horizontalHeaderItem(1)
        item.setText(_translate("Input", "Ra/Rb"))
        item = self.tableMasterCrankInfo2.horizontalHeaderItem(2)
        item.setText(_translate("Input", "wiping  actual angle"))
        __sortingEnabled = self.tableMasterCrankInfo2.isSortingEnabled()
        self.tableMasterCrankInfo2.setSortingEnabled(False)
        item = self.tableMasterCrankInfo2.item(0, 0)
        item.setText(_translate("Input", "-x"))
        item = self.tableMasterCrankInfo2.item(0, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableMasterCrankInfo2.item(0, 2)
        item.setText(_translate("Input", "0"))
        self.tableMasterCrankInfo2.setSortingEnabled(__sortingEnabled)
        self.label_17.setText(_translate("Input", "<html><head/><body><p><span style=\" font-size:12pt; font-weight:600;\">Coupled Link No.2</span></p></body></html>"))
        self.comboBoxNo2.setItemText(0, _translate("Input", "passenger side"))
        self.comboBoxNo2.setItemText(1, _translate("Input", "driver side"))
        self.label_16.setText(_translate("Input", "<html><head/><body><p><span style=\" font-size:12pt; font-weight:600;\">Motor crank axis</span></p></body></html>"))
        item = self.tableMotorInfo.verticalHeaderItem(0)
        item.setText(_translate("Input", "Reversing angle"))
        item = self.tableMotorInfo.verticalHeaderItem(1)
        item.setText(_translate("Input", "Offset angle"))
        item = self.tableMotorInfo.verticalHeaderItem(2)
        item.setText(_translate("Input", "A-X"))
        item = self.tableMotorInfo.verticalHeaderItem(3)
        item.setText(_translate("Input", "A-Y"))
        item = self.tableMotorInfo.verticalHeaderItem(4)
        item.setText(_translate("Input", "A-Z"))
        item = self.tableMotorInfo.verticalHeaderItem(5)
        item.setText(_translate("Input", "A\'-X"))
        item = self.tableMotorInfo.verticalHeaderItem(6)
        item.setText(_translate("Input", "A\'-Y"))
        item = self.tableMotorInfo.verticalHeaderItem(7)
        item.setText(_translate("Input", "A\'-Z"))
        item = self.tableMotorInfo.horizontalHeaderItem(0)
        item.setText(_translate("Input", "Value"))
        item = self.tableMotorInfo.horizontalHeaderItem(1)
        item.setText(_translate("Input", "+"))
        item = self.tableMotorInfo.horizontalHeaderItem(2)
        item.setText(_translate("Input", "-"))
        __sortingEnabled = self.tableMotorInfo.isSortingEnabled()
        self.tableMotorInfo.setSortingEnabled(False)
        item = self.tableMotorInfo.item(0, 0)
        item.setText(_translate("Input", "360"))
        item = self.tableMotorInfo.item(0, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableMotorInfo.item(0, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableMotorInfo.item(1, 0)
        item.setText(_translate("Input", "0"))
        item = self.tableMotorInfo.item(1, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableMotorInfo.item(1, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableMotorInfo.item(2, 0)
        item.setText(_translate("Input", "1838.976"))
        item = self.tableMotorInfo.item(2, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableMotorInfo.item(2, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableMotorInfo.item(3, 0)
        item.setText(_translate("Input", "-418.576"))
        item = self.tableMotorInfo.item(3, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableMotorInfo.item(3, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableMotorInfo.item(4, 0)
        item.setText(_translate("Input", "1005.54"))
        item = self.tableMotorInfo.item(4, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableMotorInfo.item(4, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableMotorInfo.item(5, 0)
        item.setText(_translate("Input", "1872.567"))
        item = self.tableMotorInfo.item(5, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableMotorInfo.item(5, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableMotorInfo.item(6, 0)
        item.setText(_translate("Input", "-411.091"))
        item = self.tableMotorInfo.item(6, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableMotorInfo.item(6, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableMotorInfo.item(7, 0)
        item.setText(_translate("Input", "960.479"))
        item = self.tableMotorInfo.item(7, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableMotorInfo.item(7, 2)
        item.setText(_translate("Input", "0"))
        self.tableMotorInfo.setSortingEnabled(__sortingEnabled)
        item = self.tableMotorAngle.verticalHeaderItem(0)
        item.setText(_translate("Input", "value"))
        item = self.tableMotorAngle.horizontalHeaderItem(0)
        item.setText(_translate("Input", "APS1"))
        item = self.tableMotorAngle.horizontalHeaderItem(1)
        item.setText(_translate("Input", "APS2"))
        item = self.tableMotorAngle.horizontalHeaderItem(2)
        item.setText(_translate("Input", "IPL"))
        item = self.tableMotorAngle.horizontalHeaderItem(3)
        item.setText(_translate("Input", "UWL"))
        item = self.tableMotorAngle.horizontalHeaderItem(4)
        item.setText(_translate("Input", "ALP"))
        item = self.tableMotorAngle.horizontalHeaderItem(5)
        item.setText(_translate("Input", "SP"))
        item = self.tableMotorAngle.horizontalHeaderItem(6)
        item.setText(_translate("Input", "KMP"))
        __sortingEnabled = self.tableMotorAngle.isSortingEnabled()
        self.tableMotorAngle.setSortingEnabled(False)
        item = self.tableMotorAngle.item(0, 0)
        item.setText(_translate("Input", " 0"))
        item = self.tableMotorAngle.item(0, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableMotorAngle.item(0, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableMotorAngle.item(0, 3)
        item.setText(_translate("Input", "0"))
        item = self.tableMotorAngle.item(0, 4)
        item.setText(_translate("Input", "0"))
        item = self.tableMotorAngle.item(0, 5)
        item.setText(_translate("Input", "0"))
        item = self.tableMotorAngle.item(0, 6)
        item.setText(_translate("Input", "0"))
        self.tableMotorAngle.setSortingEnabled(__sortingEnabled)
        item = self.tableSlaveCrankInfo.verticalHeaderItem(0)
        item.setText(_translate("Input", "D"))
        item = self.tableSlaveCrankInfo.verticalHeaderItem(1)
        item.setText(_translate("Input", "Delta"))
        item = self.tableSlaveCrankInfo.verticalHeaderItem(2)
        item.setText(_translate("Input", "FE"))
        item = self.tableSlaveCrankInfo.verticalHeaderItem(3)
        item.setText(_translate("Input", "BC"))
        item = self.tableSlaveCrankInfo.verticalHeaderItem(4)
        item.setText(_translate("Input", "CD"))
        item = self.tableSlaveCrankInfo.verticalHeaderItem(5)
        item.setText(_translate("Input", "ED"))
        item = self.tableSlaveCrankInfo.verticalHeaderItem(6)
        item.setText(_translate("Input", "F-X"))
        item = self.tableSlaveCrankInfo.verticalHeaderItem(7)
        item.setText(_translate("Input", "F-Y"))
        item = self.tableSlaveCrankInfo.verticalHeaderItem(8)
        item.setText(_translate("Input", "F-Z"))
        item = self.tableSlaveCrankInfo.verticalHeaderItem(9)
        item.setText(_translate("Input", "F\'-X"))
        item = self.tableSlaveCrankInfo.verticalHeaderItem(10)
        item.setText(_translate("Input", "F\'-Y"))
        item = self.tableSlaveCrankInfo.verticalHeaderItem(11)
        item.setText(_translate("Input", "F\'-Z"))
        item = self.tableSlaveCrankInfo.horizontalHeaderItem(0)
        item.setText(_translate("Input", "Value"))
        item = self.tableSlaveCrankInfo.horizontalHeaderItem(1)
        item.setText(_translate("Input", "+"))
        item = self.tableSlaveCrankInfo.horizontalHeaderItem(2)
        item.setText(_translate("Input", "-"))
        __sortingEnabled = self.tableSlaveCrankInfo.isSortingEnabled()
        self.tableSlaveCrankInfo.setSortingEnabled(False)
        item = self.tableSlaveCrankInfo.item(0, 0)
        item.setText(_translate("Input", "16.10572"))
        item = self.tableSlaveCrankInfo.item(0, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableSlaveCrankInfo.item(0, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableSlaveCrankInfo.item(1, 0)
        item.setText(_translate("Input", "-10"))
        item = self.tableSlaveCrankInfo.item(1, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableSlaveCrankInfo.item(1, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableSlaveCrankInfo.item(2, 0)
        item.setText(_translate("Input", "0.144"))
        item = self.tableSlaveCrankInfo.item(2, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableSlaveCrankInfo.item(2, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableSlaveCrankInfo.item(3, 0)
        item.setText(_translate("Input", "67.2"))
        item = self.tableSlaveCrankInfo.item(3, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableSlaveCrankInfo.item(3, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableSlaveCrankInfo.item(4, 0)
        item.setText(_translate("Input", "466.5"))
        item = self.tableSlaveCrankInfo.item(4, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableSlaveCrankInfo.item(4, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableSlaveCrankInfo.item(5, 0)
        item.setText(_translate("Input", "70"))
        item = self.tableSlaveCrankInfo.item(5, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableSlaveCrankInfo.item(5, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableSlaveCrankInfo.item(6, 0)
        item.setText(_translate("Input", "1897.865"))
        item = self.tableSlaveCrankInfo.item(6, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableSlaveCrankInfo.item(6, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableSlaveCrankInfo.item(7, 0)
        item.setText(_translate("Input", "-618.136"))
        item = self.tableSlaveCrankInfo.item(7, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableSlaveCrankInfo.item(7, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableSlaveCrankInfo.item(8, 0)
        item.setText(_translate("Input", "974.678"))
        item = self.tableSlaveCrankInfo.item(8, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableSlaveCrankInfo.item(8, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableSlaveCrankInfo.item(9, 0)
        item.setText(_translate("Input", "1904.541"))
        item = self.tableSlaveCrankInfo.item(9, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableSlaveCrankInfo.item(9, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableSlaveCrankInfo.item(10, 0)
        item.setText(_translate("Input", "-616.724"))
        item = self.tableSlaveCrankInfo.item(10, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableSlaveCrankInfo.item(10, 2)
        item.setText(_translate("Input", "0"))
        item = self.tableSlaveCrankInfo.item(11, 0)
        item.setText(_translate("Input", "971.32"))
        item = self.tableSlaveCrankInfo.item(11, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableSlaveCrankInfo.item(11, 2)
        item.setText(_translate("Input", "0"))
        self.tableSlaveCrankInfo.setSortingEnabled(__sortingEnabled)
        item = self.tableSlaveCrankInfo2.verticalHeaderItem(0)
        item.setText(_translate("Input", "Value"))
        item = self.tableSlaveCrankInfo2.horizontalHeaderItem(0)
        item.setText(_translate("Input", "KBEW"))
        item = self.tableSlaveCrankInfo2.horizontalHeaderItem(1)
        item.setText(_translate("Input", "Ra/Rb"))
        item = self.tableSlaveCrankInfo2.horizontalHeaderItem(2)
        item.setText(_translate("Input", "wipping actual angle"))
        __sortingEnabled = self.tableSlaveCrankInfo2.isSortingEnabled()
        self.tableSlaveCrankInfo2.setSortingEnabled(False)
        item = self.tableSlaveCrankInfo2.item(0, 0)
        item.setText(_translate("Input", "-x"))
        item = self.tableSlaveCrankInfo2.item(0, 1)
        item.setText(_translate("Input", "0"))
        item = self.tableSlaveCrankInfo2.item(0, 2)
        item.setText(_translate("Input", "0"))
        self.tableSlaveCrankInfo2.setSortingEnabled(__sortingEnabled)
        self.label_18.setText(_translate("Input", "<html><head/><body><p><span style=\" font-size:12pt; font-weight:600;\">Coupled Link No.3</span></p></body></html>"))
        self.comboBoxNo3.setItemText(0, _translate("Input", "driver side"))
        self.comboBoxNo3.setItemText(1, _translate("Input", "passenger side"))
        self.LoadButton.setText(_translate("Input", "Apply"))
        self.OutputButton.setText(_translate("Input", "Output"))
        self.ToleranceButton.setText(_translate("Input", "Tolerance"))
        self.SaveButton.setText(_translate("Input", "Update to Catia"))

class Ui_Optimization(object):
    def setupUi(self, Optimization):
        Optimization.setObjectName("Optimization")
        Optimization.resize(800, 789)
        self.centralwidget = QtWidgets.QWidget(Optimization)
        self.centralwidget.setObjectName("centralwidget")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(10, 10, 471, 31))
        self.label.setObjectName("label")
        self.frame = QtWidgets.QFrame(self.centralwidget)
        self.frame.setGeometry(QtCore.QRect(0, 40, 591, 261))
        self.frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame.setObjectName("frame")
        self.label_5 = QtWidgets.QLabel(self.frame)
        self.label_5.setGeometry(QtCore.QRect(11, 61, 72, 16))
        self.label_5.setObjectName("label_5")
        self.label_6 = QtWidgets.QLabel(self.frame)
        self.label_6.setGeometry(QtCore.QRect(11, 86, 54, 16))
        self.label_6.setObjectName("label_6")
        self.label_7 = QtWidgets.QLabel(self.frame)
        self.label_7.setGeometry(QtCore.QRect(11, 111, 96, 16))
        self.label_7.setObjectName("label_7")
        self.label_13 = QtWidgets.QLabel(self.frame)
        self.label_13.setGeometry(QtCore.QRect(199, 111, 16, 16))
        self.label_13.setObjectName("label_13")
        self.label_14 = QtWidgets.QLabel(self.frame)
        self.label_14.setGeometry(QtCore.QRect(388, 111, 16, 16))
        self.label_14.setObjectName("label_14")
        self.TableOpt = QtWidgets.QTableWidget(self.frame)
        self.TableOpt.setGeometry(QtCore.QRect(10, 130, 571, 121))
        self.TableOpt.setObjectName("TableOpt")
        self.TableOpt.setColumnCount(8)
        self.TableOpt.setRowCount(3)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt.setVerticalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt.setVerticalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt.setHorizontalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt.setHorizontalHeaderItem(4, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt.setHorizontalHeaderItem(5, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt.setHorizontalHeaderItem(6, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt.setHorizontalHeaderItem(7, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt.setItem(2, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt.setItem(2, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt.setItem(2, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt.setItem(2, 3, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt.setItem(2, 5, item)
        self.TableOpt.horizontalHeader().setDefaultSectionSize(60)
        self.TextW2 = QtWidgets.QTextEdit(self.frame)
        self.TextW2.setGeometry(QtCore.QRect(270, 60, 131, 21))
        self.TextW2.setObjectName("TextW2")
        self.Text20 = QtWidgets.QTextEdit(self.frame)
        self.Text20.setGeometry(QtCore.QRect(410, 60, 161, 21))
        self.Text20.setObjectName("Text20")
        self.layoutWidget = QtWidgets.QWidget(self.frame)
        self.layoutWidget.setGeometry(QtCore.QRect(10, 10, 561, 50))
        self.layoutWidget.setObjectName("layoutWidget")
        self.gridLayout = QtWidgets.QGridLayout(self.layoutWidget)
        self.gridLayout.setContentsMargins(0, 0, 0, 0)
        self.gridLayout.setObjectName("gridLayout")
        self.label_2 = QtWidgets.QLabel(self.layoutWidget)
        self.label_2.setObjectName("label_2")
        self.gridLayout.addWidget(self.label_2, 0, 0, 1, 1)
        self.label_3 = QtWidgets.QLabel(self.layoutWidget)
        self.label_3.setObjectName("label_3")
        self.gridLayout.addWidget(self.label_3, 1, 0, 1, 1)
        self.label_8 = QtWidgets.QLabel(self.layoutWidget)
        self.label_8.setObjectName("label_8")
        self.gridLayout.addWidget(self.label_8, 0, 2, 1, 1)
        self.label_10 = QtWidgets.QLabel(self.layoutWidget)
        self.label_10.setObjectName("label_10")
        self.gridLayout.addWidget(self.label_10, 1, 2, 1, 1)
        self.label_4 = QtWidgets.QLabel(self.layoutWidget)
        self.label_4.setObjectName("label_4")
        self.gridLayout.addWidget(self.label_4, 2, 0, 1, 1)
        self.label_16 = QtWidgets.QLabel(self.layoutWidget)
        self.label_16.setObjectName("label_16")
        self.gridLayout.addWidget(self.label_16, 0, 3, 1, 1)
        self.label_12 = QtWidgets.QLabel(self.layoutWidget)
        self.label_12.setObjectName("label_12")
        self.gridLayout.addWidget(self.label_12, 2, 2, 1, 1)
        self.label_18 = QtWidgets.QLabel(self.layoutWidget)
        self.label_18.setObjectName("label_18")
        self.gridLayout.addWidget(self.label_18, 2, 3, 1, 1)
        self.label_17 = QtWidgets.QLabel(self.layoutWidget)
        self.label_17.setObjectName("label_17")
        self.gridLayout.addWidget(self.label_17, 1, 3, 1, 1)
        self.label_9 = QtWidgets.QLabel(self.layoutWidget)
        self.label_9.setObjectName("label_9")
        self.gridLayout.addWidget(self.label_9, 0, 1, 1, 1)
        self.textEdit = QtWidgets.QTextEdit(self.frame)
        self.textEdit.setGeometry(QtCore.QRect(270, 90, 131, 21))
        self.textEdit.setObjectName("textEdit")
        self.textEdit_2 = QtWidgets.QTextEdit(self.frame)
        self.textEdit_2.setGeometry(QtCore.QRect(410, 90, 161, 21))
        self.textEdit_2.setObjectName("textEdit_2")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(360, 570, 221, 31))
        self.pushButton.setAutoDefault(True)
        self.pushButton.setDefault(True)
        self.pushButton.setObjectName("pushButton")
        self.frame_2 = QtWidgets.QFrame(self.centralwidget)
        self.frame_2.setGeometry(QtCore.QRect(10, 310, 571, 251))
        self.frame_2.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_2.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_2.setObjectName("frame_2")
        self.label_22 = QtWidgets.QLabel(self.frame_2)
        self.label_22.setGeometry(QtCore.QRect(11, 61, 72, 16))
        self.label_22.setObjectName("label_22")
        self.label_23 = QtWidgets.QLabel(self.frame_2)
        self.label_23.setGeometry(QtCore.QRect(11, 86, 54, 16))
        self.label_23.setObjectName("label_23")
        self.label_24 = QtWidgets.QLabel(self.frame_2)
        self.label_24.setGeometry(QtCore.QRect(11, 111, 96, 16))
        self.label_24.setObjectName("label_24")
        self.label_30 = QtWidgets.QLabel(self.frame_2)
        self.label_30.setGeometry(QtCore.QRect(196, 111, 16, 16))
        self.label_30.setObjectName("label_30")
        self.label_31 = QtWidgets.QLabel(self.frame_2)
        self.label_31.setGeometry(QtCore.QRect(381, 111, 16, 16))
        self.label_31.setObjectName("label_31")
        self.TableOpt2 = QtWidgets.QTableWidget(self.frame_2)
        self.TableOpt2.setGeometry(QtCore.QRect(0, 130, 571, 121))
        self.TableOpt2.setObjectName("TableOpt2")
        self.TableOpt2.setColumnCount(8)
        self.TableOpt2.setRowCount(3)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt2.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt2.setVerticalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt2.setVerticalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt2.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt2.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt2.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt2.setHorizontalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt2.setHorizontalHeaderItem(4, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt2.setHorizontalHeaderItem(5, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt2.setHorizontalHeaderItem(6, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt2.setHorizontalHeaderItem(7, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt2.setItem(2, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt2.setItem(2, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt2.setItem(2, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt2.setItem(2, 3, item)
        item = QtWidgets.QTableWidgetItem()
        self.TableOpt2.setItem(2, 5, item)
        self.TableOpt2.horizontalHeader().setDefaultSectionSize(60)
        self.TextW3 = QtWidgets.QTextEdit(self.frame_2)
        self.TextW3.setGeometry(QtCore.QRect(270, 60, 121, 21))
        self.TextW3.setObjectName("TextW3")
        self.Text30 = QtWidgets.QTextEdit(self.frame_2)
        self.Text30.setGeometry(QtCore.QRect(393, 60, 161, 21))
        self.Text30.setObjectName("Text30")
        self.layoutWidget1 = QtWidgets.QWidget(self.frame_2)
        self.layoutWidget1.setGeometry(QtCore.QRect(10, 10, 551, 50))
        self.layoutWidget1.setObjectName("layoutWidget1")
        self.gridLayout_2 = QtWidgets.QGridLayout(self.layoutWidget1)
        self.gridLayout_2.setContentsMargins(0, 0, 0, 0)
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.label_19 = QtWidgets.QLabel(self.layoutWidget1)
        self.label_19.setObjectName("label_19")
        self.gridLayout_2.addWidget(self.label_19, 0, 0, 1, 1)
        self.label_26 = QtWidgets.QLabel(self.layoutWidget1)
        self.label_26.setObjectName("label_26")
        self.gridLayout_2.addWidget(self.label_26, 0, 1, 1, 1)
        self.label_25 = QtWidgets.QLabel(self.layoutWidget1)
        self.label_25.setObjectName("label_25")
        self.gridLayout_2.addWidget(self.label_25, 0, 2, 1, 1)
        self.label_33 = QtWidgets.QLabel(self.layoutWidget1)
        self.label_33.setObjectName("label_33")
        self.gridLayout_2.addWidget(self.label_33, 0, 3, 1, 1)
        self.label_20 = QtWidgets.QLabel(self.layoutWidget1)
        self.label_20.setObjectName("label_20")
        self.gridLayout_2.addWidget(self.label_20, 1, 0, 1, 1)
        self.label_27 = QtWidgets.QLabel(self.layoutWidget1)
        self.label_27.setObjectName("label_27")
        self.gridLayout_2.addWidget(self.label_27, 1, 2, 1, 1)
        self.label_34 = QtWidgets.QLabel(self.layoutWidget1)
        self.label_34.setObjectName("label_34")
        self.gridLayout_2.addWidget(self.label_34, 1, 3, 1, 1)
        self.label_21 = QtWidgets.QLabel(self.layoutWidget1)
        self.label_21.setObjectName("label_21")
        self.gridLayout_2.addWidget(self.label_21, 2, 0, 1, 1)
        self.label_29 = QtWidgets.QLabel(self.layoutWidget1)
        self.label_29.setObjectName("label_29")
        self.gridLayout_2.addWidget(self.label_29, 2, 2, 1, 1)
        self.label_35 = QtWidgets.QLabel(self.layoutWidget1)
        self.label_35.setObjectName("label_35")
        self.gridLayout_2.addWidget(self.label_35, 2, 3, 1, 1)
        self.textEdit_3 = QtWidgets.QTextEdit(self.frame_2)
        self.textEdit_3.setGeometry(QtCore.QRect(270, 90, 121, 21))
        self.textEdit_3.setObjectName("textEdit_3")
        self.textEdit_4 = QtWidgets.QTextEdit(self.frame_2)
        self.textEdit_4.setGeometry(QtCore.QRect(390, 90, 161, 21))
        self.textEdit_4.setObjectName("textEdit_4")
        Optimization.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(Optimization)
        self.statusbar.setObjectName("statusbar")
        Optimization.setStatusBar(self.statusbar)

        self.retranslateUi(Optimization)
        QtCore.QMetaObject.connectSlotsByName(Optimization)

    def retranslateUi(self, Optimization):
        _translate = QtCore.QCoreApplication.translate
        Optimization.setWindowTitle(_translate("Optimization", "MainWindow"))
        self.label.setText(_translate("Optimization", "<html><head/><body><p><span style=\" font-size:12pt; font-weight:600;\">Optimization --Wipping angle,tangential force angle</span></p></body></html>"))
        self.label_5.setText(_translate("Optimization", "target value"))
        self.label_6.setText(_translate("Optimization", "parameter"))
        self.label_7.setText(_translate("Optimization", "coupled link No."))
        self.label_13.setText(_translate("Optimization", "2"))
        self.label_14.setText(_translate("Optimization", "2"))
        item = self.TableOpt.verticalHeaderItem(0)
        item.setText(_translate("Optimization", "optimization"))
        item = self.TableOpt.verticalHeaderItem(1)
        item.setText(_translate("Optimization", "rounded"))
        item = self.TableOpt.horizontalHeaderItem(0)
        item.setText(_translate("Optimization", "dim.1"))
        item = self.TableOpt.horizontalHeaderItem(1)
        item.setText(_translate("Optimization", "dim.2"))
        item = self.TableOpt.horizontalHeaderItem(2)
        item.setText(_translate("Optimization", "min.1"))
        item = self.TableOpt.horizontalHeaderItem(3)
        item.setText(_translate("Optimization", "max.1"))
        item = self.TableOpt.horizontalHeaderItem(4)
        item.setText(_translate("Optimization", "target.1"))
        item = self.TableOpt.horizontalHeaderItem(5)
        item.setText(_translate("Optimization", "min.2"))
        item = self.TableOpt.horizontalHeaderItem(6)
        item.setText(_translate("Optimization", "max.2"))
        item = self.TableOpt.horizontalHeaderItem(7)
        item.setText(_translate("Optimization", "target.2"))
        __sortingEnabled = self.TableOpt.isSortingEnabled()
        self.TableOpt.setSortingEnabled(False)
        item = self.TableOpt.item(2, 0)
        item.setText(_translate("Optimization", "w="))
        item = self.TableOpt.item(2, 2)
        item.setText(_translate("Optimization", "B_ss="))
        item = self.TableOpt.item(2, 5)
        item.setText(_translate("Optimization", "NYS_T"))
        self.TableOpt.setSortingEnabled(__sortingEnabled)
        self.TextW2.setHtml(_translate("Optimization", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">80</p></body></html>"))
        self.Text20.setHtml(_translate("Optimization", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">0</p></body></html>"))
        self.label_2.setText(_translate("Optimization", "driven crank No."))
        self.label_3.setText(_translate("Optimization", "objective function"))
        self.label_8.setText(_translate("Optimization", "[DS output crank]"))
        self.label_10.setText(_translate("Optimization", "output angle"))
        self.label_4.setText(_translate("Optimization", "max/min calculation"))
        self.label_16.setText(_translate("Optimization", "wipping angle measured from"))
        self.label_12.setText(_translate("Optimization", "Diff|External|"))
        self.label_18.setText(_translate("Optimization", "Diff|External|"))
        self.label_17.setText(_translate("Optimization", "tangential force angle"))
        self.label_9.setText(_translate("Optimization", "2"))
        self.textEdit.setHtml(_translate("Optimization", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">ED</p></body></html>"))
        self.textEdit_2.setHtml(_translate("Optimization", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">CD</p></body></html>"))
        self.pushButton.setText(_translate("Optimization", "Optimization"))
        self.label_22.setText(_translate("Optimization", "target value"))
        self.label_23.setText(_translate("Optimization", "parameter"))
        self.label_24.setText(_translate("Optimization", "coupled link No."))
        self.label_30.setText(_translate("Optimization", "3"))
        self.label_31.setText(_translate("Optimization", "3"))
        item = self.TableOpt2.verticalHeaderItem(0)
        item.setText(_translate("Optimization", "optimization"))
        item = self.TableOpt2.verticalHeaderItem(1)
        item.setText(_translate("Optimization", "rounded"))
        item = self.TableOpt2.horizontalHeaderItem(0)
        item.setText(_translate("Optimization", "dim.1"))
        item = self.TableOpt2.horizontalHeaderItem(1)
        item.setText(_translate("Optimization", "dim.2"))
        item = self.TableOpt2.horizontalHeaderItem(2)
        item.setText(_translate("Optimization", "min.1"))
        item = self.TableOpt2.horizontalHeaderItem(3)
        item.setText(_translate("Optimization", "max.1"))
        item = self.TableOpt2.horizontalHeaderItem(4)
        item.setText(_translate("Optimization", "target.1"))
        item = self.TableOpt2.horizontalHeaderItem(5)
        item.setText(_translate("Optimization", "min.2"))
        item = self.TableOpt2.horizontalHeaderItem(6)
        item.setText(_translate("Optimization", "max.2"))
        item = self.TableOpt2.horizontalHeaderItem(7)
        item.setText(_translate("Optimization", "target.2"))
        __sortingEnabled = self.TableOpt2.isSortingEnabled()
        self.TableOpt2.setSortingEnabled(False)
        item = self.TableOpt2.item(2, 0)
        item.setText(_translate("Optimization", "w="))
        item = self.TableOpt2.item(2, 2)
        item.setText(_translate("Optimization", "B_ss="))
        item = self.TableOpt2.item(2, 5)
        item.setText(_translate("Optimization", "NYS_T"))
        self.TableOpt2.setSortingEnabled(__sortingEnabled)
        self.TextW3.setHtml(_translate("Optimization", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">80</p></body></html>"))
        self.Text30.setHtml(_translate("Optimization", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">0</p></body></html>"))
        self.label_19.setText(_translate("Optimization", "driven crank No."))
        self.label_26.setText(_translate("Optimization", "3"))
        self.label_25.setText(_translate("Optimization", "[DS output crank]"))
        self.label_33.setText(_translate("Optimization", "wipping angle measured from"))
        self.label_20.setText(_translate("Optimization", "objective function"))
        self.label_27.setText(_translate("Optimization", "output angle"))
        self.label_34.setText(_translate("Optimization", "tangential force angle"))
        self.label_21.setText(_translate("Optimization", "max/min calculation"))
        self.label_29.setText(_translate("Optimization", "Diff|External|"))
        self.label_35.setText(_translate("Optimization", "Diff|External|"))
        self.textEdit_3.setHtml(_translate("Optimization", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">ED</p></body></html>"))
        self.textEdit_4.setHtml(_translate("Optimization", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">CD</p></body></html>"))

class Ui_alphaNumeric(object):
    def setupUi(self, alphaNumeric):
        alphaNumeric.setObjectName("alphaNumeric")
        alphaNumeric.resize(906, 765)
        self.centralwidget = QtWidgets.QWidget(alphaNumeric)
        self.centralwidget.setObjectName("centralwidget")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(10, 0, 461, 31))
        self.label_2.setObjectName("label_2")
        self.frame = QtWidgets.QFrame(self.centralwidget)
        self.frame.setGeometry(QtCore.QRect(0, 80, 841, 441))
        self.frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame.setObjectName("frame")
        self.label_4 = QtWidgets.QLabel(self.frame)
        self.label_4.setGeometry(QtCore.QRect(20, 10, 151, 16))
        self.label_4.setObjectName("label_4")
        self.tableAlfa = QtWidgets.QTableWidget(self.frame)
        self.tableAlfa.setGeometry(QtCore.QRect(10, 30, 851, 421))
        self.tableAlfa.setObjectName("tableAlfa")
        self.tableAlfa.setColumnCount(15)
        self.tableAlfa.setRowCount(0)
        item = QtWidgets.QTableWidgetItem()
        self.tableAlfa.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableAlfa.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableAlfa.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableAlfa.setHorizontalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableAlfa.setHorizontalHeaderItem(4, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableAlfa.setHorizontalHeaderItem(5, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableAlfa.setHorizontalHeaderItem(6, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableAlfa.setHorizontalHeaderItem(7, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableAlfa.setHorizontalHeaderItem(8, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableAlfa.setHorizontalHeaderItem(9, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableAlfa.setHorizontalHeaderItem(10, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableAlfa.setHorizontalHeaderItem(11, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableAlfa.setHorizontalHeaderItem(12, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableAlfa.setHorizontalHeaderItem(13, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableAlfa.setHorizontalHeaderItem(14, item)
        self.tableAlfa.horizontalHeader().setDefaultSectionSize(55)
        self.layoutWidget = QtWidgets.QWidget(self.centralwidget)
        self.layoutWidget.setGeometry(QtCore.QRect(10, 40, 555, 25))
        self.layoutWidget.setObjectName("layoutWidget")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.layoutWidget)
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label = QtWidgets.QLabel(self.layoutWidget)
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)
        self.Textstep = QtWidgets.QLineEdit(self.layoutWidget)
        self.Textstep.setObjectName("Textstep")
        self.horizontalLayout.addWidget(self.Textstep)
        self.label_3 = QtWidgets.QLabel(self.layoutWidget)
        self.label_3.setObjectName("label_3")
        self.horizontalLayout.addWidget(self.label_3)
        self.TableButton = QtWidgets.QPushButton(self.layoutWidget)
        self.TableButton.setObjectName("TableButton")
        self.horizontalLayout.addWidget(self.TableButton)
        self.ExtremeButton = QtWidgets.QPushButton(self.layoutWidget)
        self.ExtremeButton.setObjectName("ExtremeButton")
        self.horizontalLayout.addWidget(self.ExtremeButton)
        self.layoutWidget.raise_()
        self.label_2.raise_()
        self.frame.raise_()
        alphaNumeric.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(alphaNumeric)
        self.statusbar.setObjectName("statusbar")
        alphaNumeric.setStatusBar(self.statusbar)

        self.retranslateUi(alphaNumeric)
        QtCore.QMetaObject.connectSlotsByName(alphaNumeric)

    def retranslateUi(self, alphaNumeric):
        _translate = QtCore.QCoreApplication.translate
        alphaNumeric.setWindowTitle(_translate("alphaNumeric", "MainWindow"))
        self.label_2.setText(_translate("alphaNumeric", "<html><head/><body><p><span style=\" font-size:12pt; font-weight:600;\">alphanumeric output</span></p></body></html>"))
        self.label_4.setText(_translate("alphaNumeric", "<html><head/><body><p><span style=\" font-size:12pt; font-weight:600;\">Tabular Data</span></p></body></html>"))
        item = self.tableAlfa.horizontalHeaderItem(0)
        item.setText(_translate("alphaNumeric", "ALFA"))
        item = self.tableAlfa.horizontalHeaderItem(1)
        item.setText(_translate("alphaNumeric", "Beta"))
        item = self.tableAlfa.horizontalHeaderItem(2)
        item.setText(_translate("alphaNumeric", "Beta_S"))
        item = self.tableAlfa.horizontalHeaderItem(3)
        item.setText(_translate("alphaNumeric", "Beta_SS"))
        item = self.tableAlfa.horizontalHeaderItem(4)
        item.setText(_translate("alphaNumeric", "NYS_T"))
        item = self.tableAlfa.horizontalHeaderItem(5)
        item.setText(_translate("alphaNumeric", "NYS_A"))
        item = self.tableAlfa.horizontalHeaderItem(6)
        item.setText(_translate("alphaNumeric", "NYK_T"))
        item = self.tableAlfa.horizontalHeaderItem(7)
        item.setText(_translate("alphaNumeric", "NYK_A"))
        item = self.tableAlfa.horizontalHeaderItem(8)
        item.setText(_translate("alphaNumeric", "Gama"))
        item = self.tableAlfa.horizontalHeaderItem(9)
        item.setText(_translate("alphaNumeric", "Gama_S"))
        item = self.tableAlfa.horizontalHeaderItem(10)
        item.setText(_translate("alphaNumeric", "Gama_SS"))
        item = self.tableAlfa.horizontalHeaderItem(11)
        item.setText(_translate("alphaNumeric", "NYS_T"))
        item = self.tableAlfa.horizontalHeaderItem(12)
        item.setText(_translate("alphaNumeric", "NYS_A"))
        item = self.tableAlfa.horizontalHeaderItem(13)
        item.setText(_translate("alphaNumeric", "NYK_T"))
        item = self.tableAlfa.horizontalHeaderItem(14)
        item.setText(_translate("alphaNumeric", "NYK_A"))
        self.label.setText(_translate("alphaNumeric", "Stepwidth for tables"))
        self.Textstep.setText(_translate("alphaNumeric", "30"))
        self.label_3.setText(_translate("alphaNumeric", "°"))
        self.TableButton.setText(_translate("alphaNumeric", "tabular data"))
        self.ExtremeButton.setText(_translate("alphaNumeric", "Extreme values"))



class Ui_extreme(object):
    def setupUi(self, extreme):
        extreme.setObjectName("extreme")
        extreme.resize(877, 726)
        self.centralwidget = QtWidgets.QWidget(extreme)
        self.centralwidget.setObjectName("centralwidget")
        self.frame_2 = QtWidgets.QFrame(self.centralwidget)
        self.frame_2.setGeometry(QtCore.QRect(0, 10, 861, 331))
        self.frame_2.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_2.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_2.setObjectName("frame_2")
        self.label_5 = QtWidgets.QLabel(self.frame_2)
        self.label_5.setGeometry(QtCore.QRect(0, 0, 171, 16))
        self.label_5.setObjectName("label_5")
        self.tableExtreme1 = QtWidgets.QTableWidget(self.frame_2)
        self.tableExtreme1.setGeometry(QtCore.QRect(0, 20, 841, 61))
        self.tableExtreme1.setObjectName("tableExtreme1")
        self.tableExtreme1.setColumnCount(4)
        self.tableExtreme1.setRowCount(1)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme1.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme1.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme1.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme1.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme1.setHorizontalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme1.setItem(0, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme1.setItem(0, 2, item)
        self.tableExtreme1.horizontalHeader().setDefaultSectionSize(195)
        self.tableExtreme2 = QtWidgets.QTableWidget(self.frame_2)
        self.tableExtreme2.setGeometry(QtCore.QRect(0, 110, 841, 91))
        self.tableExtreme2.setObjectName("tableExtreme2")
        self.tableExtreme2.setColumnCount(10)
        self.tableExtreme2.setRowCount(2)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme2.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme2.setVerticalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme2.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme2.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme2.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme2.setHorizontalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme2.setHorizontalHeaderItem(4, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme2.setHorizontalHeaderItem(5, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme2.setHorizontalHeaderItem(6, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme2.setHorizontalHeaderItem(7, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme2.setHorizontalHeaderItem(8, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme2.setHorizontalHeaderItem(9, item)
        item = QtWidgets.QTableWidgetItem()
        brush = QtGui.QBrush(QtGui.QColor(255, 85, 0))
        brush.setStyle(QtCore.Qt.NoBrush)
        item.setForeground(brush)
        self.tableExtreme2.setItem(0, 6, item)
        self.tableExtreme2.horizontalHeader().setDefaultSectionSize(80)
        self.label_6 = QtWidgets.QLabel(self.frame_2)
        self.label_6.setGeometry(QtCore.QRect(520, 90, 321, 20))
        self.label_6.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_6.setFrameShape(QtWidgets.QFrame.WinPanel)
        self.label_6.setObjectName("label_6")
        self.tableExtreme3 = QtWidgets.QTableWidget(self.frame_2)
        self.tableExtreme3.setGeometry(QtCore.QRect(0, 230, 841, 91))
        self.tableExtreme3.setObjectName("tableExtreme3")
        self.tableExtreme3.setColumnCount(10)
        self.tableExtreme3.setRowCount(2)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme3.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme3.setVerticalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme3.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme3.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme3.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme3.setHorizontalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme3.setHorizontalHeaderItem(4, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme3.setHorizontalHeaderItem(5, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme3.setHorizontalHeaderItem(6, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme3.setHorizontalHeaderItem(7, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme3.setHorizontalHeaderItem(8, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme3.setHorizontalHeaderItem(9, item)
        item = QtWidgets.QTableWidgetItem()
        brush = QtGui.QBrush(QtGui.QColor(255, 85, 0))
        brush.setStyle(QtCore.Qt.NoBrush)
        item.setForeground(brush)
        self.tableExtreme3.setItem(0, 6, item)
        self.tableExtreme3.horizontalHeader().setDefaultSectionSize(80)
        self.label_7 = QtWidgets.QLabel(self.frame_2)
        self.label_7.setGeometry(QtCore.QRect(520, 210, 321, 20))
        self.label_7.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_7.setFrameShape(QtWidgets.QFrame.WinPanel)
        self.label_7.setObjectName("label_7")
        self.frame_3 = QtWidgets.QFrame(self.centralwidget)
        self.frame_3.setGeometry(QtCore.QRect(0, 360, 861, 331))
        self.frame_3.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_3.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_3.setObjectName("frame_3")
        self.label_8 = QtWidgets.QLabel(self.frame_3)
        self.label_8.setGeometry(QtCore.QRect(0, 0, 171, 16))
        self.label_8.setObjectName("label_8")
        self.tableExtreme1_2 = QtWidgets.QTableWidget(self.frame_3)
        self.tableExtreme1_2.setGeometry(QtCore.QRect(0, 20, 841, 61))
        self.tableExtreme1_2.setObjectName("tableExtreme1_2")
        self.tableExtreme1_2.setColumnCount(4)
        self.tableExtreme1_2.setRowCount(1)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme1_2.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme1_2.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme1_2.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme1_2.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme1_2.setHorizontalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme1_2.setItem(0, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme1_2.setItem(0, 2, item)
        self.tableExtreme1_2.horizontalHeader().setDefaultSectionSize(195)
        self.tableExtreme2_2 = QtWidgets.QTableWidget(self.frame_3)
        self.tableExtreme2_2.setGeometry(QtCore.QRect(0, 110, 841, 91))
        self.tableExtreme2_2.setObjectName("tableExtreme2_2")
        self.tableExtreme2_2.setColumnCount(10)
        self.tableExtreme2_2.setRowCount(2)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme2_2.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme2_2.setVerticalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme2_2.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme2_2.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme2_2.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme2_2.setHorizontalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme2_2.setHorizontalHeaderItem(4, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme2_2.setHorizontalHeaderItem(5, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme2_2.setHorizontalHeaderItem(6, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme2_2.setHorizontalHeaderItem(7, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme2_2.setHorizontalHeaderItem(8, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme2_2.setHorizontalHeaderItem(9, item)
        item = QtWidgets.QTableWidgetItem()
        brush = QtGui.QBrush(QtGui.QColor(255, 85, 0))
        brush.setStyle(QtCore.Qt.NoBrush)
        item.setForeground(brush)
        self.tableExtreme2_2.setItem(0, 6, item)
        self.tableExtreme2_2.horizontalHeader().setDefaultSectionSize(80)
        self.label_9 = QtWidgets.QLabel(self.frame_3)
        self.label_9.setGeometry(QtCore.QRect(520, 90, 321, 20))
        self.label_9.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_9.setFrameShape(QtWidgets.QFrame.WinPanel)
        self.label_9.setObjectName("label_9")
        self.tableExtreme3_2 = QtWidgets.QTableWidget(self.frame_3)
        self.tableExtreme3_2.setGeometry(QtCore.QRect(0, 230, 841, 91))
        self.tableExtreme3_2.setObjectName("tableExtreme3_2")
        self.tableExtreme3_2.setColumnCount(10)
        self.tableExtreme3_2.setRowCount(2)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme3_2.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme3_2.setVerticalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme3_2.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme3_2.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme3_2.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme3_2.setHorizontalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme3_2.setHorizontalHeaderItem(4, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme3_2.setHorizontalHeaderItem(5, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme3_2.setHorizontalHeaderItem(6, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme3_2.setHorizontalHeaderItem(7, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme3_2.setHorizontalHeaderItem(8, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableExtreme3_2.setHorizontalHeaderItem(9, item)
        item = QtWidgets.QTableWidgetItem()
        brush = QtGui.QBrush(QtGui.QColor(255, 85, 0))
        brush.setStyle(QtCore.Qt.NoBrush)
        item.setForeground(brush)
        self.tableExtreme3_2.setItem(0, 6, item)
        self.tableExtreme3_2.horizontalHeader().setDefaultSectionSize(80)
        self.label_10 = QtWidgets.QLabel(self.frame_3)
        self.label_10.setGeometry(QtCore.QRect(520, 210, 321, 20))
        self.label_10.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_10.setFrameShape(QtWidgets.QFrame.WinPanel)
        self.label_10.setObjectName("label_10")
        extreme.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(extreme)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 877, 23))
        self.menubar.setObjectName("menubar")
        extreme.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(extreme)
        self.statusbar.setObjectName("statusbar")
        extreme.setStatusBar(self.statusbar)

        self.retranslateUi(extreme)
        QtCore.QMetaObject.connectSlotsByName(extreme)

    def retranslateUi(self, extreme):
        _translate = QtCore.QCoreApplication.translate
        extreme.setWindowTitle(_translate("extreme", "MainWindow"))
        self.label_5.setText(_translate("extreme", "Coupled Link No.2(cranks)"))
        item = self.tableExtreme1.verticalHeaderItem(0)
        item.setText(_translate("extreme", "Values"))
        item = self.tableExtreme1.horizontalHeaderItem(0)
        item.setText(_translate("extreme", "Target wipping angle:"))
        item = self.tableExtreme1.horizontalHeaderItem(1)
        item.setText(_translate("extreme", "Calculated wipping angle:"))
        item = self.tableExtreme1.horizontalHeaderItem(2)
        item.setText(_translate("extreme", "Park Alpha-Abs"))
        item = self.tableExtreme1.horizontalHeaderItem(3)
        item.setText(_translate("extreme", "theoret Alpha-Abs"))
        __sortingEnabled = self.tableExtreme1.isSortingEnabled()
        self.tableExtreme1.setSortingEnabled(False)
        self.tableExtreme1.setSortingEnabled(__sortingEnabled)
        item = self.tableExtreme2.verticalHeaderItem(0)
        item.setText(_translate("extreme", "Max"))
        item = self.tableExtreme2.verticalHeaderItem(1)
        item.setText(_translate("extreme", "Min"))
        item = self.tableExtreme2.horizontalHeaderItem(0)
        item.setText(_translate("extreme", "Alfa"))
        item = self.tableExtreme2.horizontalHeaderItem(1)
        item.setText(_translate("extreme", "Beta"))
        item = self.tableExtreme2.horizontalHeaderItem(2)
        item.setText(_translate("extreme", "Alfa"))
        item = self.tableExtreme2.horizontalHeaderItem(3)
        item.setText(_translate("extreme", "Beta_S"))
        item = self.tableExtreme2.horizontalHeaderItem(4)
        item.setText(_translate("extreme", "Alfa"))
        item = self.tableExtreme2.horizontalHeaderItem(5)
        item.setText(_translate("extreme", "Beta_SS"))
        item = self.tableExtreme2.horizontalHeaderItem(6)
        item.setText(_translate("extreme", "Alfa"))
        item = self.tableExtreme2.horizontalHeaderItem(7)
        item.setText(_translate("extreme", "NYS_T"))
        item = self.tableExtreme2.horizontalHeaderItem(8)
        item.setText(_translate("extreme", "Alfa"))
        item = self.tableExtreme2.horizontalHeaderItem(9)
        item.setText(_translate("extreme", "NYS_A"))
        __sortingEnabled = self.tableExtreme2.isSortingEnabled()
        self.tableExtreme2.setSortingEnabled(False)
        self.tableExtreme2.setSortingEnabled(__sortingEnabled)
        self.label_6.setText(_translate("extreme", "<html><head/><body><p align=\"center\"><span style=\" font-size:10pt; font-weight:600;\">output crank</span></p></body></html>"))
        item = self.tableExtreme3.verticalHeaderItem(0)
        item.setText(_translate("extreme", "Max"))
        item = self.tableExtreme3.verticalHeaderItem(1)
        item.setText(_translate("extreme", "Min"))
        item = self.tableExtreme3.horizontalHeaderItem(6)
        item.setText(_translate("extreme", "Alfa"))
        item = self.tableExtreme3.horizontalHeaderItem(7)
        item.setText(_translate("extreme", "NYS_T"))
        item = self.tableExtreme3.horizontalHeaderItem(8)
        item.setText(_translate("extreme", "Alfa"))
        item = self.tableExtreme3.horizontalHeaderItem(9)
        item.setText(_translate("extreme", "NYS_A"))
        __sortingEnabled = self.tableExtreme3.isSortingEnabled()
        self.tableExtreme3.setSortingEnabled(False)
        self.tableExtreme3.setSortingEnabled(__sortingEnabled)
        self.label_7.setText(_translate("extreme", "<html><head/><body><p align=\"center\"><span style=\" font-size:10pt; font-weight:600;\">actuation crank</span></p></body></html>"))
        self.label_8.setText(_translate("extreme", "Coupled Link No.3(cranks)"))
        item = self.tableExtreme1_2.verticalHeaderItem(0)
        item.setText(_translate("extreme", "Values"))
        item = self.tableExtreme1_2.horizontalHeaderItem(0)
        item.setText(_translate("extreme", "Target wipping angle:"))
        item = self.tableExtreme1_2.horizontalHeaderItem(1)
        item.setText(_translate("extreme", "Calculated wipping angle:"))
        item = self.tableExtreme1_2.horizontalHeaderItem(2)
        item.setText(_translate("extreme", "Park Alpha-Abs"))
        item = self.tableExtreme1_2.horizontalHeaderItem(3)
        item.setText(_translate("extreme", "theoret Alpha-Abs"))
        __sortingEnabled = self.tableExtreme1_2.isSortingEnabled()
        self.tableExtreme1_2.setSortingEnabled(False)
        self.tableExtreme1_2.setSortingEnabled(__sortingEnabled)
        item = self.tableExtreme2_2.verticalHeaderItem(0)
        item.setText(_translate("extreme", "Max"))
        item = self.tableExtreme2_2.verticalHeaderItem(1)
        item.setText(_translate("extreme", "Min"))
        item = self.tableExtreme2_2.horizontalHeaderItem(0)
        item.setText(_translate("extreme", "Alfa"))
        item = self.tableExtreme2_2.horizontalHeaderItem(1)
        item.setText(_translate("extreme", "Beta"))
        item = self.tableExtreme2_2.horizontalHeaderItem(2)
        item.setText(_translate("extreme", "Alfa"))
        item = self.tableExtreme2_2.horizontalHeaderItem(3)
        item.setText(_translate("extreme", "Beta_S"))
        item = self.tableExtreme2_2.horizontalHeaderItem(4)
        item.setText(_translate("extreme", "Alfa"))
        item = self.tableExtreme2_2.horizontalHeaderItem(5)
        item.setText(_translate("extreme", "Beta_SS"))
        item = self.tableExtreme2_2.horizontalHeaderItem(6)
        item.setText(_translate("extreme", "Alfa"))
        item = self.tableExtreme2_2.horizontalHeaderItem(7)
        item.setText(_translate("extreme", "NYS_T"))
        item = self.tableExtreme2_2.horizontalHeaderItem(8)
        item.setText(_translate("extreme", "Alfa"))
        item = self.tableExtreme2_2.horizontalHeaderItem(9)
        item.setText(_translate("extreme", "NYS_A"))
        __sortingEnabled = self.tableExtreme2_2.isSortingEnabled()
        self.tableExtreme2_2.setSortingEnabled(False)
        self.tableExtreme2_2.setSortingEnabled(__sortingEnabled)
        self.label_9.setText(_translate("extreme", "<html><head/><body><p align=\"center\"><span style=\" font-size:10pt; font-weight:600;\">output crank</span></p></body></html>"))
        item = self.tableExtreme3_2.verticalHeaderItem(0)
        item.setText(_translate("extreme", "Max"))
        item = self.tableExtreme3_2.verticalHeaderItem(1)
        item.setText(_translate("extreme", "Min"))
        item = self.tableExtreme3_2.horizontalHeaderItem(6)
        item.setText(_translate("extreme", "Alfa"))
        item = self.tableExtreme3_2.horizontalHeaderItem(7)
        item.setText(_translate("extreme", "NYS_T"))
        item = self.tableExtreme3_2.horizontalHeaderItem(8)
        item.setText(_translate("extreme", "Alfa"))
        item = self.tableExtreme3_2.horizontalHeaderItem(9)
        item.setText(_translate("extreme", "NYS_A"))
        __sortingEnabled = self.tableExtreme3_2.isSortingEnabled()
        self.tableExtreme3_2.setSortingEnabled(False)
        self.tableExtreme3_2.setSortingEnabled(__sortingEnabled)
        self.label_10.setText(_translate("extreme", "<html><head/><body><p align=\"center\"><span style=\" font-size:10pt; font-weight:600;\">actuation crank</span></p></body></html>"))
class Ui_Tolerance(object):
    def setupUi(self, Tolerance):
        Tolerance.setObjectName("Tolerance")
        Tolerance.resize(800, 600)
        self.centralwidget = QtWidgets.QWidget(Tolerance)
        self.centralwidget.setObjectName("centralwidget")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(10, 10, 161, 31))
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(10, 40, 141, 31))
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(10, 70, 141, 31))
        self.label_3.setObjectName("label_3")
        self.comboLink = QtWidgets.QComboBox(self.centralwidget)
        self.comboLink.setGeometry(QtCore.QRect(120, 70, 69, 22))
        self.comboLink.setObjectName("comboLink")
        self.comboLink.addItem("")
        self.comboLink.addItem("")
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(230, 70, 141, 31))
        self.label_4.setObjectName("label_4")
        self.comboObj = QtWidgets.QComboBox(self.centralwidget)
        self.comboObj.setGeometry(QtCore.QRect(370, 70, 91, 22))
        self.comboObj.setObjectName("comboObj")
        self.comboObj.addItem("")
        self.comboObj.addItem("")
        self.ButtonCal = QtWidgets.QPushButton(self.centralwidget)
        self.ButtonCal.setGeometry(QtCore.QRect(530, 70, 75, 23))
        self.ButtonCal.setObjectName("ButtonCal")
        self.tableTolerance = QtWidgets.QTableWidget(self.centralwidget)
        self.tableTolerance.setGeometry(QtCore.QRect(10, 100, 631, 401))
        self.tableTolerance.setObjectName("tableTolerance")
        self.tableTolerance.setColumnCount(5)
        self.tableTolerance.setRowCount(1)
        item = QtWidgets.QTableWidgetItem()
        self.tableTolerance.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableTolerance.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableTolerance.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableTolerance.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableTolerance.setHorizontalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableTolerance.setHorizontalHeaderItem(4, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableTolerance.setItem(0, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableTolerance.setItem(0, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableTolerance.setItem(0, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableTolerance.setItem(0, 3, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableTolerance.setItem(0, 4, item)
        Tolerance.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(Tolerance)
        self.statusbar.setObjectName("statusbar")
        Tolerance.setStatusBar(self.statusbar)

        self.retranslateUi(Tolerance)
        QtCore.QMetaObject.connectSlotsByName(Tolerance)

    def retranslateUi(self, Tolerance):
        _translate = QtCore.QCoreApplication.translate
        Tolerance.setWindowTitle(_translate("Tolerance", "MainWindow"))
        self.label.setText(_translate("Tolerance", "<html><head/><body><p><span style=\" font-weight:600;\">Tolerance Calculation</span></p></body></html>"))
        self.label_2.setText(_translate("Tolerance", "<html><head/><body><p><span style=\" font-weight:600;\">maximum value</span></p></body></html>"))
        self.label_3.setText(_translate("Tolerance", "<html><head/><body><p>Link</p></body></html>"))
        self.comboLink.setItemText(0, _translate("Tolerance", "Master"))
        self.comboLink.setItemText(1, _translate("Tolerance", "Slave"))
        self.label_4.setText(_translate("Tolerance", "<html><head/><body><p>objective Function</p></body></html>"))
        self.comboObj.setItemText(0, _translate("Tolerance", "outputAngle"))
        self.comboObj.setItemText(1, _translate("Tolerance", "tangent force angle"))
        self.ButtonCal.setText(_translate("Tolerance", "Calculate"))
        item = self.tableTolerance.verticalHeaderItem(0)
        item.setText(_translate("Tolerance", "num"))
        item = self.tableTolerance.horizontalHeaderItem(0)
        item.setText(_translate("Tolerance", "N2"))
        item = self.tableTolerance.horizontalHeaderItem(1)
        item.setText(_translate("Tolerance", "Parameter"))
        item = self.tableTolerance.horizontalHeaderItem(2)
        item.setText(_translate("Tolerance", "Dimension"))
        item = self.tableTolerance.horizontalHeaderItem(3)
        item.setText(_translate("Tolerance", "Function Value"))
        item = self.tableTolerance.horizontalHeaderItem(4)
        item.setText(_translate("Tolerance", "Deviation"))
        __sortingEnabled = self.tableTolerance.isSortingEnabled()
        self.tableTolerance.setSortingEnabled(False)
        self.tableTolerance.setSortingEnabled(__sortingEnabled)
 
class MainWindow(QtWidgets.QTabWidget, Ui_MainWindow):
    def __init__(self,parent=None):
        super(MainWindow,self).__init__(parent)
        self.setWindowTitle("Kinematics  Calculation of wiper linkage")
        self.resize(1000, 1000)
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

# =============================================================================
        self.WindowInput.LoadButton.clicked.connect(self.WindowInput.LoadClicked)
        self.WindowInput.LoadButton.clicked.connect(self.WindowOpt.OptClicked)
        self.WindowInput.LoadButton.clicked.connect(self.WindowInput.UpdateOptValue)


        self.WindowInput.OutputButton.clicked.connect(self.WindowInput.close)
        self.WindowInput.OutputButton.clicked.connect(self.WindowAlpha.show)
        self.WindowInput.OutputButton.clicked.connect(self.resetAlpha)
        
        self.WindowAlpha.ExtremeButton.clicked.connect(self.WindowExtreme.ExtremeTable)
        self.WindowAlpha.ExtremeButton.clicked.connect(self.WindowAlpha.close)
        self.WindowAlpha.ExtremeButton.clicked.connect(self.WindowExtreme.show)
        self.WindowAlpha.ExtremeButton.clicked.connect(self.reset2)
        
        self.WindowInput.ToleranceButton.clicked.connect(self.WindowInput.close)
        self.WindowInput.ToleranceButton.clicked.connect(self.WindowTolerance.show)
        self.WindowInput.ToleranceButton.clicked.connect(self.resetTolerance)
        
        
        self.WindowInput.SaveButton.clicked.connect(self.writeMainCordinate)
        self.WindowInput.SaveButton.clicked.connect(self.writeDetailCordinate)
        self.WindowInput.SaveButton.clicked.connect(self.WindowInput.updateNewsheet)
        self.WindowInput.SaveButton.clicked.connect(self.saveExcel)


    def saveExcel(self):
        gs.wb.save(gs.excel_out)
        print('save completed')
    def writeMainCordinate(self):
        outM1=Func.Output(gs.BC,gs.CD,gs.ED,gs.xm1,gs.A,gs.B,gs.E,gs.F,KBEW=gs.KBEW)  #Master -UWL,[alpha,beta,NYS_T,NYS_A,NYK_T,NYK_A,C[0][0],C[0][1],C[0][2],Db[0][0],Db[0][1],Db[0][2]]
        outM2=Func.Output(gs.BC,gs.CD,gs.ED,gs.xm2,gs.A,gs.B,gs.E,gs.F,KBEW=gs.KBEW) #Master-OWL [alpha,beta,NYS_T,NYS_A,NYK_T,NYK_A,C[0][0],C[0][1],C[0][2],Db[0][0],Db[0][1],Db[0][2]]
        outS3=Func.Output(gs.BC2,gs.CD2,gs.ED2,gs.xm12,gs.A2,gs.B2,gs.E2,gs.F2,KBEW=gs.KBEW2) #slave UWL
        outS4=Func.Output(gs.BC2,gs.CD2,gs.ED2,gs.xm22,gs.A2,gs.B2,gs.E2,gs.F2,KBEW=gs.KBEW2) #slave OWL
        gs.listWrite = []
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
            gs.listWrite =[y for x in array for y in x]
            num = len(gs.listWrite)
    
            for i in range(num):
                Func.write(gs.sheetDesign1,i+1,2,'%.4f'%gs.listWrite[i])
            gs.wb1.save(filename=gs.excel_design1)
        
        else:
            if gs.Master =='driver side':
                array = [gs.A,gs.Ap,gs.F,gs.Fp,gs.F2,gs.Fp2 ,gs.B, gs.E, gs.B2, gs.E2,outM1[6:12],outS3[6:12],
                outM1[6:12],outS3[6:12],outM2[6:12],outS4[6:12]]
            else:

                array = [gs.A,gs.Ap,gs.F2,gs.Fp2,gs.F,gs.Fp ,gs.B2, gs.E2, gs.B, gs.E,outS3[6:12],outM1[6:12],
                outS3[6:12],outM1[6:12],outS4[6:12],outM2[6:12]]
            gs.listWrite =[y for x in array for y in x]
            num = len(gs.listWrite)
    
            for i in range(num):
                Func.write(gs.sheetDesign1,i+1,2,'%.4f'%gs.listWrite[i])
            gs.wb1.save(filename=gs.excel_design1)

    def writeDetailCordinate (self):
            alphaList = np.linspace(gs.xm1,gs.xm1+360,12,endpoint =False)
            gs.listWrite2 =[]
            for alpha in alphaList:
                if gs.MechanicType =='Center' :
                    outM=   Func.Output(gs.BC,gs.CD,gs.ED,alpha,gs.A,gs.B,gs.E,gs.F,KBEW=gs.KBEW)  #
                    outS =  Func.Output(gs.BC2,gs.CD2,gs.ED2,alpha,gs.A2,gs.B2,gs.E2,gs.F2,KBEW=gs.KBEW2)  #
                else:
                    outM=   Func.Output(gs.BC, gs.CD,  gs.ED, alpha, gs.A, gs.B , gs.E,  gs.F, KBEW=gs.KBEW)  #
                    alpha2 =outM[1]+gs.Delta2 
                    outS =  Func.Output(gs.BC2,gs.CD2, gs.ED2,alpha2,gs.A2,gs.B2, gs.E2, gs.F2, KBEW=gs.KBEW2)  #
                if gs.Master =='driver side':                     
                    gs.listWrite2.extend(outM[6:12])
                    gs.listWrite2.extend(outS[6:12])
                else:
                    gs.listWrite2.extend(outS[6:12])
                    gs.listWrite2.extend(outM[6:12])
            num = len(gs.listWrite2)
            for i in range(num):
                Func.write(gs.sheetDesign2,i+1,2,'%.4f'%gs.listWrite2[i])

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
class MainWindow_Tolerance(QtWidgets.QMainWindow,Ui_Tolerance):
    def __init__(self,parent=None):
        super(MainWindow_Tolerance,self).__init__(parent)
        self.setupUi(self)
        self.ButtonCal.clicked.connect(self.Tolerance)
    def Tolerance(self):
        noCrank = self.comboLink.currentText() #Master,Slave
        obj = self.comboObj.currentText() #wipping angle,NYS_T

        listToleranceStrall = ["BC", "ED", 'Delta','CD',"F_X", "F_Y", "F_Z","Fp_X", "Fp_Y", "Fp_Z",'FE',"A_X", "A_Y", "A_Z","Ap_X", "Ap_Y", "Ap_Z",'Distance',
                             "BC2", "ED2", 'Delta2','CD2',"F_X2", "F_Y2", "F_Z2", "Fp_X2", "Fp_Y2", "Fp_Z2",'FE2',"A_X2", "A_Y2", "A_Z2","Ap_X2", "Ap_Y2", "Ap_Z2",'Distance2'] #TBD:DElta,Distance
        numTolerance = int(len(listToleranceStrall)/2)
        if  obj =='outputAngle':
                t = 0
                index = 1 # to write out
        elif obj =='tangent force angle':
                t = 1
                index =4 # to write out
        else:
                print('please configure objective function')
                
        if noCrank=='Master':
                listToleranceStr = listToleranceStrall[0:numTolerance]
                Base = [eval('gs.'+st) for st in listToleranceStr]
                Target = Func.Tolerance(Base,gs.xm1,gs.xm2)[t]
                n = 0
        elif noCrank == 'Slave':
                listToleranceStr = listToleranceStrall[numTolerance:]
                Base = [eval('gs.'+st) for st in listToleranceStr]
                Target = Func.Tolerance(Base,gs.xm12,gs.xm22)[t]
                n=1
        else:
                print('please configure noCrank')

        
        errorPList = [ 'gs.errorP.'+t for t in listToleranceStr]
        errorPositive = [eval(t) for t in errorPList]
        errorNList = [ 'gs.errorN.'+t for t in listToleranceStr]
        errorNegative = [eval(t) for t in errorNList]
        w2ErrorList = []  # store error
        kpiList = []  # store kpi dimension
        w2List = [] # store value
        
        
        ToleranceValue = Base.copy()  # to be changed
        if n==0: #master
                for i in range(numTolerance):
                        Func.updateTolerance(ToleranceValue, w2ErrorList, w2List, kpiList, i,  t,Target, errorPositive[i],
                                errorNegative[i], gs.xm1, gs.xm2)
                arrayi = ToleranceValue.copy()
                print(ToleranceValue)
                for i in range(numTolerance):
                        arrayi[i] = Base[i]
                        Func.updateTolerance(arrayi, w2ErrorList, w2List, kpiList, i, t,Target, errorPositive[i],
                                errorNegative[i], gs.xm1, gs.xm2)
        elif n==1:#slave
                for i in range(numTolerance):
                        Func.updateTolerance(ToleranceValue, w2ErrorList, w2List, kpiList, i, t,Target,
                                             errorPositive[i],
                                             errorNegative[i], gs.xm12, gs.xm22)
                
                arrayi = ToleranceValue.copy()
                for i in range(numTolerance):
                        arrayi[i] = Base[i]
                        Func.updateTolerance(arrayi,w2ErrorList, w2List, kpiList, i, t,Target,
                                             errorPositive[i],
                                             errorNegative[i], gs.xm12, gs.xm22)
        Parameter = np.tile(listToleranceStr, 2)
        #w2ErrorList=np.array(w2ErrorList).cumsum()
        # start wrtie to ui/excel
        self.tableTolerance.setRowCount(2*numTolerance)
        for i in range(2*numTolerance):
                self.tableTolerance.setItem(i, 0, Qitem(str(2*(n+1))))
                self.tableTolerance.setItem(i, 1, Qitem(Parameter[i]))
                self.tableTolerance.setItem(i, 2, Qitem('%.4f'%kpiList[i]))
                self.tableTolerance.setItem(i, 3, Qitem('%.4f'%w2List[i]))
                self.tableTolerance.setItem(i, 4, Qitem('%.4f'%w2ErrorList[i]))
        Func.write(gs.sheet6,3,4,(n+2))
        Func.write(gs.sheet6,3,6,index)
        startColumn = 2
        startRow = 10
        for i in range(2*numTolerance):
                Func.write(gs.sheet6, startRow + i, startColumn, 2*(n+1))  # TBD:N2
                Func.write(gs.sheet6, startRow + i, startColumn + 1, Parameter[i])  # parameter
                Func.write(gs.sheet6, startRow + i, startColumn + 2, '%.4f'%kpiList[i])
                Func.write(gs.sheet6, startRow + i, startColumn + 3, '%.4f'%w2List[i])
                Func.write(gs.sheet6, startRow + i, startColumn + 4, '%.4f'%w2ErrorList[i])

        print('Tolerance calculation completed')

class MainWindow_extreme(QtWidgets.QMainWindow,Ui_extreme):
    def __init__(self,parent=None):
        super(MainWindow_extreme,self).__init__(parent)
        self.setupUi(self)
    def ExtremeTable(self):
        if gs.MechanicType =='Center':
                UWL1 = gs.xm1+90
                UWL2 = gs.xm12+90
        else:
                UWL1 = gs.xm1 + 90
                UWL2 = gs.xm1 + 90
        if gs.DriveType =='Standard':
                Park = gs.xm1+90
        else:
                Park = gs.xm1+90 #TBD
        self.tableExtreme1.setItem(0, 2, Qitem('%.2f' % Park))
        self.tableExtreme1.setItem(0, 3, Qitem('%.2f' % UWL1))
        self.tableExtreme1_2.setItem(0, 2, Qitem('%.2f' % Park))
        self.tableExtreme1_2.setItem(0, 3, Qitem('%.2f' % UWL2))

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

        Func.write(gs.sheet5, 3, 11, '%.2f' % Park)
        Func.write(gs.sheet5, 4, 11, '%.2f' % UWL1)
        Func.write(gs.sheet5, 20, 11, '%.2f' % Park)
        Func.write(gs.sheet5, 21, 11, '%.2f' % UWL2)

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

class MainWindow_alphaNumeric(QtWidgets.QMainWindow,Ui_alphaNumeric):
    def __init__(self,parent=None):
        super(MainWindow_alphaNumeric,self).__init__(parent)
        self.setupUi(self)
        
        self.TableButton.clicked.connect(self.alphaTable)
    def alphaTable(self):
        
        print('-----------------------start map calculation-------------------------')
        time_start = time.time()
        #constrain:
        #NYS-A ef CD
        #==============================================================================

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

        #plot
        x=out['alpha']
        plt.figure(1)
        ax=plt.subplot(2,2,1)
        plt.plot(x,out['beta'],'k--',label='Beta')
        plt.plot(x,out2['beta2'],'r--',label='Gamma')
        plt.legend()

        ax=plt.subplot(2,2,2)
        plt.plot(x,out['beta_s'],'k--',label='B-S')
        plt.plot(x,out2['beta_s2'],'r:',label='G-S')
        plt.legend()

        ax=plt.subplot(2,2,3)
        l3=plt.plot(x,out['beta_ss'],'k--',label='B-SS')
        l4=plt.plot(x,out2['beta_ss2'],'r:',label='G-SS')
        plt.legend()
        plt.savefig('.\\output\\pics\\'+gs.ProjectName+'_Angle'+gs.styleTime+'.jpg')

        plt.figure(2)
        ax=plt.subplot(2,2,1)
        plt.plot(x,out['NYS_T'],'k--',label='NYS_T')
        plt.plot(x,out2['NYS_T2'],'r:',label='NYS_T2')
        plt.legend()

        ax=plt.subplot(2,2,2)
        plt.plot(x,out['NYK_T'],'k--',label='NYK_T')
        plt.plot(x,out2['NYK_T2'],'r:',label='NYK_T2')
        plt.legend()

        ax=plt.subplot(2,2,3)
        plt.plot(x,out['NYS_A'],'k--',label='NYS_A')
        plt.plot(x,out2['NYS_A2'],'r:',label='NYS_A2')
        plt.legend()

        ax=plt.subplot(2,2,4)
        plt.plot(x,out['NYK_A'],'k--',label='NYK_A')
        plt.plot(x,out2['NYK_A2'],'r:',label='NYK_A2')
        plt.legend()
        plt.savefig('.\\output\\pics\\'+gs.ProjectName+'_NY'+gs.styleTime+'.jpg')

        #end plot

        #end plot
        cols =['alpha','beta','beta_s','beta_ss','NYS_T','NYS_A','NYK_T','NYK_A',]
        cols2=['beta2','beta_s2','beta_ss2','NYS_T2','NYS_A2','NYK_T2','NYK_A2']
        colsCordinate=['Cx','Cy','Cz','Dx','Dy','Dz']
        colsCordinate2=['Cx2','Cy2','Cz2','Dx2','Dy2','Dz2']
        # =============================================================================
        outKPI = out.loc[:,cols]
        outKPI2 = out2.loc[:,cols2]
        outCordinate = out.loc[:,colsCordinate]
        outCordinate2=out2.loc[:,colsCordinate2]
        # =============================================================================
        outKPI['NYS_T']=outKPI['NYS_T'].astype('float64')
        outKPI['NYS_A']=outKPI['NYS_A'].astype('float64')
        outKPI['NYK_T']=outKPI['NYK_T'].astype('float64')
        outKPI['NYK_A']=outKPI['NYK_A'].astype('float64')
        outKPI2['NYS_T2']=outKPI2['NYS_T2'].astype('float64')
        outKPI2['NYS_A2']=outKPI2['NYS_A2'].astype('float64')
        outKPI2['NYK_T2']=outKPI2['NYK_T2'].astype('float64')
        outKPI2['NYK_A2']=outKPI2['NYK_A2'].astype('float64')

        outKPIall=pd.concat([outKPI,outKPI2],axis=1)
        outall = pd.concat([out,out2],axis=1)

        # =============================================================================
        gs.maxArray=[]
        gs.minArray=[]
        gs.maxArray2=[]
        gs.minArray2=[]
        # =============================================================================
        # =============================================================================

        for i in range(1,len(cols)):
            listT = Func.getMax(outKPIall,i)
            gs.maxArray.append(listT[0])
            gs.maxArray.append(listT[1])
            listT = Func.getMin(outKPIall,i)
            gs.minArray.append(listT[0])
            gs.minArray.append(listT[1])

        for i in range(len(cols),len(cols)+len(cols2)):
            listT = Func.getMax(outKPIall,i)
            gs.maxArray2.append(listT[0])
            gs.maxArray2.append(listT[1])
            listT = Func.getMin(outKPIall,i)
            gs.minArray2.append(listT[0])
            gs.minArray2.append(listT[1])
            

        gs.w2cal=gs.maxArray[1]-gs.minArray[1]
        gs.w3cal=gs.maxArray2[1]-gs.minArray2[1]
        print('w2=%.2f'%gs.w2cal+'\tw3=%.2f'%gs.w3cal)
        time_end=time.time()

        print(' map time cost：%.4f'%(time_end-time_start)+'s')
        startRow = 10
        startCol = 1
        self.tableAlfa.setRowCount(gs.num)
        for i in range(gs.num):
            for j in range(outKPIall.shape[1]):
                self.tableAlfa.setItem(i,j,Qitem('%.2f'%outKPIall.iloc[i,j]))
                Func.write(gs.sheet4,startRow+i , startCol+j , outKPIall.iloc[i,j])
        Func.write(gs.sheet4,3,6,gs.step)

        #outCordinate.to_excel('outputCordinate.xlsx')
        
        #Range=[1.2,50,8,90,8]# 90 TBD
        #write to excel

    #animation
        matplotlib.matplotlib_fname()

        number=gs.num
        ListC=[]
        ListD=[]
        ListC2=[]
        ListD2=[]
        
        
        Atemp=np.array([gs.A]*number).T
        Btemp=np.array([gs.B]*number).T
        Etemp=np.array([gs.E]*number).T
        Ftemp=np.array([gs.F]*number).T
        A2temp=np.array([gs.A2]*number).T
        B2temp=np.array([gs.B2]*number).T
        E2temp=np.array([gs.E2]*number).T
        F2temp=np.array([gs.F2]*number).T
        Ctemp= np.array((outCordinate['Cx'],outCordinate['Cy'],outCordinate['Cz']))
        Dtemp= np.array((outCordinate['Dx'],outCordinate['Dy'],outCordinate['Dz']))
        C2temp= np.array((outCordinate2['Cx2'],outCordinate2['Cy2'],outCordinate2['Cz2']))
        D2temp= np.array((outCordinate2['Dx2'],outCordinate2['Dy2'],outCordinate2['Dz2']))
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
        

        xrange = xmax-xmin
        yrange = ymax-ymin
        zrange = zmax-zmin


        ab=gs.B-gs.A
        lines = [ax2.plot([data[i][0,0],data[i][3,0]],[data[i][1,0],data[i][4,0]],[data[i][2,0],data[i][5,0]],label=Label[i],color=colors[i],linestyle=linestyles[i])[0] for i  in range(10)]
        gs.ab = [ab[0]/xrange,ab[1]/yrange,ab[2]/zrange] #scale axis equal
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
        return gs.line_ani

    
class MainWindow_Opt(QtWidgets.QMainWindow,Ui_Optimization):
    def __init__(self,parent=None):
        super(MainWindow_Opt,self).__init__(parent)
        self.setupUi(self)
        self.pushButton.clicked.connect(self.OptClicked)

        
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
        if gs.MechanicType =='Center':
            Equal1 = repr(equal1[0]).replace('xCrank', 'model.outCrank').replace('xLink',
                                                                                 'model.link')  # link length equation for master link
            Equal2 = repr(equal2).replace('xCrank', 'model.outCrank').replace('sqrt',
                                                                              'pe.sqrt')  # bc parrel bd equation for master link
            Equal12 = repr(equal12[0]).replace('xCrank', 'model.outCrank2').replace('xLink', 'model.link2')
                                                                                                                       # link length equation for slave link
            Equal22 = repr(equal22).replace('xCrank', 'model.outCrank2').replace('sqrt',
                                                                              'pe.sqrt')  # bc parrel bd equation for slave link
                                                                              
            F_M11 = Equal1.replace('sin(xs)', 'model.sxs1').replace('cos(xs)', 'model.cxs1').replace('sin(xm)',
                                                                                                     'model.sxm1').replace(
                'cos(xm)', 'model.cxm1')  # 11 means equation 1 for  position 1
            F_M12 = Equal1.replace('sin(xs)', 'model.sxs2').replace('cos(xs)', 'model.cxs2').replace('sin(xm)',
                                                                                                     'model.sxm2').replace(
                'cos(xm)', 'model.cxm2')
            
            F_M11n = Equal1.replace('sin(xs)', 'model.sxs1n').replace('cos(xs)', 'model.cxs1n').replace('sin(xm)',
                                                                                                     'model.sxm1n').replace(
                'cos(xm)', 'model.cxm1n')  #
            
            F_M21 = Equal2.replace('sin(xs)', 'model.sxs1').replace('cos(xs)', 'model.cxs1').replace('sin(xm)',
                                                                                                     'model.sxm1').replace(
                'cos(xm)', 'model.cxm1')
            F_M22 = Equal2.replace('sin(xs)', 'model.sxs2').replace('cos(xs)', 'model.cxs2').replace('sin(xm)',
                                                                                                     'model.sxm2').replace(
                'cos(xm)', 'model.cxm2')
            F_M21n = Equal2.replace('sin(xs)', 'model.sxs1n').replace('cos(xs)', 'model.cxs1n').replace('sin(xm)',
                                                                                                     'model.sxm1n').replace(
                'cos(xm)', 'model.cxm1n')
            
            F_S11 = Equal12.replace('sin(xs)', 'model.sxs12').replace('cos(xs)', 'model.cxs12').replace('sin(xm)',
                                                                                                          'model.sxm12').replace(
                'cos(xm)', 'model.cxm12')
            F_S12 = Equal12.replace('sin(xs)', 'model.sxs22').replace('cos(xs)', 'model.cxs22').replace('sin(xm)',
                                                                                                          'model.sxm22').replace(
                'cos(xm)', 'model.cxm22')  # two position start and OWL
            F_S21 = Equal22.replace('sin(xs)', 'model.sxs12').replace('cos(xs)', 'model.cxs12').replace('sin(xm)',
                                                                                                     'model.sxm12').replace(
                'cos(xm)', 'model.cxm12')
            F_S22 = Equal22.replace('sin(xs)', 'model.sxs22').replace('cos(xs)', 'model.cxs22').replace('sin(xm)',
                                                                                                     'model.sxm22').replace(
                'cos(xm)', 'model.cxm22')
                
            
            F_M1NYK_T = repr(NYK_T).replace('sin(xm)', 'model.sxm1').replace('cos(xm)', 'model.cxm1').replace('acos',
                                                                                                              'pe.acos').replace(
                'xCrank', 'model.outCrank').replace('acos', 'pe.acos').replace('sin(xm)', 'model.sxm1').replace('cos(xm)',
                                                                                                                'model.cxm1').replace(
                'sqrt', 'pe.sqrt').replace('sin(xs)', 'model.sxs1').replace('cos(xs)', 'model.cxs1')
            F_M2NYK_T = repr(NYK_T).replace('sin(xm)', 'model.sxm2').replace('cos(xm)', 'model.cxm2').replace('acos',
                                                                                                              'pe.acos').replace(
                'xCrank', 'model.outCrank').replace('acos', 'pe.acos').replace('sin(xm)', 'model.sxm1').replace('cos(xm)',
                                                                                                                'model.cxm1').replace(
                'sqrt', 'pe.sqrt').replace('sin(xs)', 'model.sxs2').replace('cos(xs)', 'model.cxs2')
            F_M1NYS_T = repr(NYS_T).replace('sin(xs)', 'model.sxs1').replace('cos(xs)', 'model.cxs1').replace('xCrank',
                                                                                                              'model.outCrank').replace(
                'acos', 'pe.acos').replace('sin(xm)', 'model.sxm1').replace('cos(xm)', 'model.cxm1').replace('sqrt', 'pe.sqrt')
            F_M2NYS_T = repr(NYS_T).replace('sin(xs)', 'model.sxs2').replace('cos(xs)', 'model.cxs2').replace('xCrank',
                                                                                                              'model.outCrank').replace(
                'acos', 'pe.acos').replace('sin(xm)', 'model.sxm2').replace('cos(xm)', 'model.cxm2').replace('sqrt', 'pe.sqrt')
            F_S1NYS_T = repr(NYS_T2).replace('sin(xs)', 'model.sxs12').replace('cos(xs)', 'model.cxs12').replace('xCrank',
                                                                                                                   'model.outCrank2').replace(
                'acos', 'pe.acos').replace('sin(xm)', 'model.sxm12').replace('cos(xm)', 'model.cxm12').replace('sqrt',
                                                                                                                 'pe.sqrt').replace(
                'BC', 'model.BC2')
            F_S2NYS_T = repr(NYS_T2).replace('sin(xs)', 'model.sxs22').replace('cos(xs)', 'model.cxs22').replace('xCrank',
                                                                                                                   'model.outCrank2').replace(
                'acos', 'pe.acos').replace('sin(xm)', 'model.sxm22').replace('cos(xm)', 'model.cxm22').replace('sqrt',
                                                                                                                 'pe.sqrt').replace(
                'BC', 'model.BC2')
            F_S1NYK_T = repr(NYK_T2).replace('sin(xm)', 'model.sxm12').replace('cos(xm)', 'model.cxm12').replace('sin(xs)',
                                                                                                                   'model.sxs12').replace(
                'cos(xs)', 'model.cxs12').replace('acos', 'pe.acos').replace('xCrank', 'model.outCrank2').replace('sqrt',
                                                                                                                    'pe.sqrt').replace(
                'BC', 'model.BC2')
            F_S2NYK_T = repr(NYK_T2).replace('sin(xm)', 'model.sxm22').replace('cos(xm)', 'model.cxm22').replace('sin(xs)',
                                                                                                                   'model.sxs22').replace(
                'cos(xs)', 'model.cxs22').replace('acos', 'pe.acos').replace('xCrank', 'model.outCrank2').replace('sqrt',
                                                                                                                    'pe.sqrt').replace(
                'BC', 'model.BC2')

            model = pe.ConcreteModel()
            model.link = pe.Var(initialize= gs.CD, bounds=(4.4 * gs.BC, 10 * gs.BC))  # length of link
            model.outCrank = pe.Var(initialize= gs.ED, bounds=(0.25 * gs.BC, 4 * gs.BC))  # length of output crank
            model.sxm1 = pe.Var(initialize=0.5, bounds=(-1, 1))
            model.cxm1 = pe.Var(initialize=math.sqrt(3) / 2, bounds=(-1, 1))
            model.sxm2 = pe.Var(initialize=-0.5, bounds=(-1, 1))
            model.cxm2 = pe.Var(initialize=math.sqrt(3) / 2, bounds=(-1, 1))
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
            
            model.BC2 = pe.Var(initialize=gs.BC, bounds=(gs.BC, gs.BC))
            
            # --- master link
            # 1n represents inline position
            # 1 represents starts position ,depends on measured from
            # 2 represents OWL position
            # wiping angel should be angle from 1 to 2
            model.Con = pe.ConstraintList()
            model.obj = pe.Objective(
                expr=((model.cxs2 * model.cxs1 + model.sxs2 * model.sxs1 - math.cos(gs.w2)) ** 2)  # minmize wipping angle error
            +(model.cxs22 * model.cxs12 + model.sxs22 * model.sxs12 - math.cos(gs.w3)) ** 2)  # the lowest wipping angle requirement
            # =============================================================================
            model.Con.add ((model.cxs2 * model.cxs1 + model.sxs2 * model.sxs1 - math.cos(gs.w2)) ** 2<=0.01)
            model.Con.add((model.cxs22 * model.cxs12 + model.sxs22 * model.sxs12 - math.cos(gs.w3)) ** 2<=0.01)
            if Clock:
                model.Con.add(model.sxs2*model.cxs1-model.sxs1*model.cxs2>=0)
            else:
                model.Con.add(model.sxs2*model.cxs1-model.sxs1*model.cxs2<=0)
            if Clock2:
                model.Con.add(model.sxs22*model.cxs12-model.sxs12*model.cxs22>=0)
            else:
                print('pls specicy clock')
                
            model.Con.add(eval(F_M11) == 0)
            model.Con.add(eval(F_M12) == 0)  # physical requirement
            model.Con.add(1000 * (eval(F_M21)**2 - 1) == 0)
            model.Con.add(1000 * (eval(F_M22)**2 - 1) == 0)  #
            model.Con.add(eval(F_S11) == 0)
            model.Con.add(eval(F_S12) == 0)  # physical requirement
            model.Con.add(1000 * (eval(F_S21)**2 - 1) == 0)
            model.Con.add(1000 * (eval(F_S22)**2- 1) == 0)  #
            
            model.Con.add(expr=(model.sxm2 ** 2 + model.cxm2 ** 2 - 1) == 0)
            model.Con.add(expr=(model.sxm1 ** 2 + model.cxm1 ** 2 - 1) == 0)
            model.Con.add(expr=(model.sxs1 ** 2 + model.cxs1 ** 2 - 1) == 0)
            model.Con.add(expr=(model.sxs2 ** 2 + model.cxs2 ** 2 - 1) == 0)
            
            model.Con.add(expr=(model.sxm22 ** 2 + model.cxm22 ** 2 - 1) == 0)
            model.Con.add(expr=(model.sxm12 ** 2 + model.cxm12 ** 2 - 1) == 0)
            model.Con.add(expr=(model.sxs12 ** 2 + model.cxs12 ** 2 - 1) == 0)
            model.Con.add(expr=(model.sxs22 ** 2 + model.cxs22 ** 2 - 1) == 0)
            # =============================================================================
            model.Con.add((eval(F_M1NYS_T)) + (eval(F_M2NYS_T)) == 0)
            model.Con.add((eval(F_S1NYS_T)) + (eval(F_S2NYS_T)) == 0)
            # =============================================================================
            # =============================================================================
            # model.Con.add(eval(F_M1NYS_T) ** 2 <= (math.cos(40 * math.pi / 180)) ** 2)  # NYS_T<50
            # model.Con.add(eval(F_S1NYS_T) ** 2 <= (math.cos(40 * math.pi / 180)) ** 2)  # NYS_T<50
            model.Con.add(expr=model.cxm1*model.cxm12+model.sxm1*model.sxm12>=math.cos(8*math.pi/180))
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
            
            gs.xm1 = Func.getDegree(cxm1, sxm1)
            gs.xm2 = Func.getDegree(cxm2, sxm2)
            gs.xs1 = Func.getDegree(cxs1, sxs1)
            gs.xs2 = Func.getDegree(cxs2, sxs2)
            
            
            gs.CD = (pe.value((model.link)))
            gs.ED = (pe.value(model.outCrank))
            M1NYS_T = Func.degree(math.pi / 2 - math.acos(pe.value(M1NYS_Tv)))
            M2NYS_T = Func.degree(math.pi / 2 - math.acos(pe.value(M2NYS_Tv)))
            
            sxm12 = pe.value(model.sxm12)
            cxm12 = pe.value(model.cxm12)
            sxm22 = pe.value(model.sxm22)
            cxm22 = pe.value(model.cxm22)
            sxs12 = pe.value(model.sxs12)
            cxs12 = pe.value(model.cxs12)
            sxs22 = pe.value(model.sxs22)
            cxs22 = pe.value(model.cxs22)
            # =============================================================================
            
            # =============================================================================
            gs.xs22 = Func.getDegree(cxs22, sxs22)
            gs.xs12 = Func.getDegree(cxs12, sxs12)
            gs.xm22 = Func.getDegree(cxm22, sxm22)
            gs.xm12 = Func.getDegree(cxm12, sxm12)
            gs.CD2 = (pe.value((model.link2)))
            gs.ED2 = (pe.value(model.outCrank2))
            
            # =============================================================================
            S1NYS_T = Func.degree(math.pi / 2 - math.acos(pe.value(S1NYS_Tv)))
            S2NYS_T = Func.degree(math.pi / 2 - math.acos(pe.value(S2NYS_Tv)))

        else: # Master-Slave
                    #--- master link
            Equal1 = repr(equal1[0]).replace('xCrank','model.outCrank').replace('xLink','model.link') # link length equation for master link
            Equal2 = repr(equal2).replace('xCrank','model.outCrank').replace('sqrt','pe.sqrt') # bc parrel bd equation for master link
            Equal12 = repr(equal12[0]).replace('xCrank','model2.outCrank2').replace('xLink','model2.link2').replace('BC','model2.bc2') #link length equation for slave link

            F_M11=Equal1.replace('sin(xs)','model.sxs1').replace('cos(xs)','model.cxs1').replace('sin(xm)','model.sxm1').replace('cos(xm)','model.cxm1') # 11 means equation 1 for start position
            F_M12=Equal1.replace('sin(xs)','model.sxs2').replace('cos(xs)','model.cxs2').replace('sin(xm)','model.sxm2').replace('cos(xm)','model.cxm2')
            F_M13=Equal1.replace('sin(xs)','model.sxs3').replace('cos(xs)','model.cxs3').replace('sin(xm)','model.sxm3').replace('cos(xm)','model.cxm3') # 11 means equation 1 for start position

            F_M21=Equal2.replace('sin(xs)','model.sxs1').replace('cos(xs)','model.cxs1').replace('sin(xm)','model.sxm1').replace('cos(xm)','model.cxm1')
            F_M22=Equal2.replace('sin(xs)','model.sxs2').replace('cos(xs)','model.cxs2').replace('sin(xm)','model.sxm2').replace('cos(xm)','model.cxm2')
            F_M23=Equal2.replace('sin(xs)','model.sxs3').replace('cos(xs)','model.cxs3').replace('sin(xm)','model.sxm3').replace('cos(xm)','model.cxm3')

            F_S11=Equal12.replace('sin(xs)','model2.sxs12').replace('cos(xs)','model2.cxs12').replace('sin(xm)','model2.sxm12').replace('cos(xm)','model2.cxm12')
            F_S12=Equal12.replace('sin(xs)','model2.sxs22').replace('cos(xs)','model2.cxs22').replace('sin(xm)','model2.sxm22').replace('cos(xm)','model2.cxm22') # two position start and OWL

            F_M1NYK_T= repr(NYK_T).replace('sin(xm)','model.sxm1').replace('cos(xm)','model.cxm1').replace('acos','pe.acos').replace('xCrank','model.outCrank').replace('acos','pe.acos').replace('sin(xm)','model.sxm1').replace('cos(xm)','model.cxm1').replace('sqrt','pe.sqrt').replace('sin(xs)','model.sxs1').replace('cos(xs)','model.cxs1')
            F_M2NYK_T= repr(NYK_T).replace('sin(xm)','model.sxm2').replace('cos(xm)','model.cxm2').replace('acos','pe.acos').replace('xCrank','model.outCrank').replace('acos','pe.acos').replace('sin(xm)','model.sxm1').replace('cos(xm)','model.cxm1').replace('sqrt','pe.sqrt').replace('sin(xs)','model.sxs2').replace('cos(xs)','model.cxs2')
            F_M1NYS_T= repr(NYS_T).replace('sin(xs)','model.sxs1').replace('cos(xs)','model.cxs1').replace('xCrank','model.outCrank').replace('acos','pe.acos').replace('sin(xm)','model.sxm1').replace('cos(xm)','model.cxm1').replace('sqrt','pe.sqrt')
            F_M2NYS_T= repr(NYS_T).replace('sin(xs)','model.sxs2').replace('cos(xs)','model.cxs2').replace('xCrank','model.outCrank').replace('acos','pe.acos').replace('sin(xm)','model.sxm2').replace('cos(xm)','model.cxm2').replace('sqrt','pe.sqrt')
            F_S1NYS_T= repr(NYS_T2).replace('sin(xs)','model2.sxs12').replace('cos(xs)','model2.cxs12').replace('xCrank','model2.outCrank2').replace('acos','pe.acos').replace('sin(xm)','model2.sxm12').replace('cos(xm)','model2.cxm12').replace('sqrt','pe.sqrt').replace('BC','model2.bc2')
            F_S2NYS_T= repr(NYS_T2).replace('sin(xs)','model2.sxs22').replace('cos(xs)','model2.cxs22').replace('xCrank','model2.outCrank2').replace('acos','pe.acos').replace('sin(xm)','model2.sxm22').replace('cos(xm)','model2.cxm22').replace('sqrt','pe.sqrt').replace('BC','model2.bc2')
            F_S1NYK_T= repr(NYK_T2).replace('sin(xm)','model2.sxm12').replace('cos(xm)','model2.cxm12').replace('sin(xs)','model2.sxs12').replace('cos(xs)','model2.cxs12').replace('acos','pe.acos').replace('xCrank','model2.outCrank2').replace('sqrt','pe.sqrt').replace('BC','model2.bc2')
            F_S2NYK_T= repr(NYK_T2).replace('sin(xm)','model2.sxm22').replace('cos(xm)','model2.cxm22').replace('sin(xs)','model2.sxs22').replace('cos(xs)','model2.cxs22').replace('acos','pe.acos').replace('xCrank','model2.outCrank2').replace('sqrt','pe.sqrt').replace('BC','model2.bc2')
            
            model = pe.ConcreteModel()
            model.link = pe.Var(initialize = gs.CD,bounds=(3*gs.BC,10*gs.BC) ) #length of link 
            model.outCrank = pe.Var(initialize =gs.ED,bounds = (0.5*gs.BC,2*gs.BC)) # length of output crank 
            # =============================================================================
            # =============================================================================
            
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
            model.obj=pe.Objective(expr=(model.cxs2*model.cxs1+model.sxs2*model.sxs1-math.cos(gs.w2))**2) # minmize wipping angle error
            
            model.Con.add(expr=(model.cxs2 * model.cxs1 + model.sxs2 * model.sxs1 - math.cos(
                    gs.w2)) ** 2 <= 0.001)  # the lowest wipping angle requirement

            if (Clock==1):
                model.Con.add(model.sxs2*model.cxs1-model.cxs2*model.sxs1>=0)
            elif(Clock==-1) :
                model.Con.add(model.sxs2*model.cxs1-model.cxs2*model.sxs1<=0)
            else:
                print(gs.KBEW)

            model.Con.add(eval(F_M11)==0) # meet basic physical constrain
            model.Con.add(eval(F_M12)==0) # meet basic physical constrain
            #TBD: MechanicType: Center,Master-Slave
            # =============================================================================
            if (gs.DriveType=='Standard'):
                # standard system, 1 and 2 are in inlien position
                model.Con.add(1000*(eval(F_M22)**2-1)==0) #
                model.Con.add(1000*(eval(F_M21)**2-1)==0)
                
            elif (gs.DriveType=='Reversing'): #TBD
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

                    
            else:
                print('please configure if is standard system ')
            model.Con.add(expr=(model.sxm2**2+model.cxm2**2-1)==0)
            model.Con.add(expr=(model.sxm1**2+model.cxm1**2-1)==0)
            model.Con.add(expr=(model.sxs1**2+model.cxs1**2-1)==0)
            model.Con.add(expr=(model.sxs2**2+model.cxs2**2-1)==0)
            model.Con.add((eval(F_M1NYS_T))+(eval(F_M2NYS_T))==0)
            model.Con.add(eval(F_M1NYS_T)**2<=(math.cos(40*math.pi/360))**2)#NYS_T<50
            tt1=eval(F_M1NYS_T)
            tt2=eval(F_M2NYS_T)
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

            gs.xs1=Func.getDegree(cxs1,sxs1)
            gs.xm1=Func.getDegree(cxm1,sxm1)
            gs.xm2=Func.getDegree(cxm2,sxm2)
            gs.xs2=Func.getDegree(cxs2,sxs2)
            if gs.DriveType=='Reversing':
                sxm3=pe.value(model.sxm3)
                cxm3=pe.value(model.cxm3)
                sxs3=pe.value(model.sxs3)
                cxs3=pe.value(model.cxs3)
                gs.xs3=Func.getDegree(cxs3,sxs3)
                gs.xm3=Func.getDegree(cxm3,sxm3)


            M1NYS_Tv = eval(F_M1NYS_T)
            M2NYS_Tv = eval(F_M2NYS_T)
            M1NYK_Tv = eval(F_M1NYK_T)
            M2NYK_Tv = eval(F_M2NYK_T)
            
            
            # =============================================================================

            gs.CD=  (pe.value((model.link)))
            gs.ED=(pe.value(model.outCrank))
            M1NYS_T=Func.degree(math.pi/2-math.acos(pe.value(M1NYS_Tv)))
            M2NYS_T=Func.degree(math.pi/2-math.acos(pe.value(M2NYS_Tv)))
            M1NYK_T=Func.degree(math.pi/2-math.acos(pe.value(M1NYK_Tv)))
            M2NYK_T=Func.degree(math.pi/2-math.acos(pe.value(M2NYK_Tv)))

            if gs.DriveType=='Reversing':
                
                # Equal1inline = repr(equal1[0]).replace('xCrank','gs.ED').replace('xLink','gs.CD').replace('xm','angle[0]').replace('xs','angle[1]').replace('sin','sp.sin').replace('cos','sp.cos').replace('sqrt','sp.sqrt').replace('BC','gs.BC') # link length equation for master link
                # Equal2inline = repr(equal2).replace('xCrank','gs.ED').replace('xLink','gs.CD').replace('xm','angle[0]').replace('xs','angle[1]').replace('sin','sp.sin').replace('cos','sp.cos').replace('sqrt','sp.sqrt').replace('BC','gs.BC')
                # EqualInline = [Equal1inline,Equal2inline]
                # def Finline(angle):
                #     F1 = eval(Equal1inline)
                #     F2 = 1000*(eval(Equal2inline)-1)
                #     return([F1,F2])
                # if gs.KBEW =='+x':
                #     [xim1,xis1]= op.fsolve(Finline,[0,0])
                # else :
                #     [xim1,xis1] = op.fsolve(Finline,[math.pi,math.pi])
                # #[xim2,xis2]= op.fsolve(Finline,[math.pi,math.pi])
                # gs.Alfa = gs.xm2-gs.xm1
                # gs.offsetAngle = gs.xm1-Func.degree(xim1)
                gs.Alfa = gs.xm2-gs.xm1
                gs.offsetAngle = gs.xm1-gs.xm3
            else:
                gs.offsetAngle = 0
                gs.Alfa = 360

            gs.w2opt = Func.wipAngle(gs.xs2,gs.xs1,gs.KBEW)
            print('CD=%.4f'% gs.CD +'\tED=%.4f'% gs.ED +'\tNYS_T1=%.4f'% M1NYS_T +'\tNYS_T2=%.4f'% M2NYS_T)
            print('NYK_T1={:.4f},  NYK_T2= {:.4f} ' .format(M1NYK_T,M2NYK_T))
            print('w2=%.4f'%gs.w2opt+'\tw2Target=%.4f'%(180/math.pi*gs.w2))
            if gs.DriveType=='Reversing':
                print('offsetAngle = {:.4f}, Reversing angle={:.4f}'.format(gs.offsetAngle,gs.Alfa))

            self.TableOpt.setItem(0,0,Qitem('%.4f'%gs.ED))
            self.TableOpt.setItem(0,1,Qitem('%.4f'%gs.CD))
            self.TableOpt.setItem(0,2,Qitem('%.4f'%gs.xs1))
            self.TableOpt.setItem(0,3,Qitem('%.4f'%gs.xs2))
            self.TableOpt.setItem(0,4,Qitem('%.4f'%gs.w2opt))
            self.TableOpt.setItem(0,5,Qitem('%.4f'%(M1NYS_T)))
            self.TableOpt.setItem(0,6,Qitem('%.4f'%(M2NYS_T)))
            self.TableOpt.setItem(0,7,Qitem('%.4f'%(M1NYS_T+M2NYS_T)))
            Func.write(gs.sheet2, 2, 12, gs.startFrom)
            Func.write(gs.sheet2,8,6,gs.w2Target)
            Func.write(gs.sheet2,16,2,gs.ED)
            Func.write(gs.sheet2,16,3,gs.CD)
            Func.write(gs.sheet2,16,5,'%.4f'% gs.xs1)
            Func.write(gs.sheet2,16,6,'%.4f'% gs.xs2)
            Func.write(gs.sheet2,16,7,'%.4f'% gs.w2opt)
            Func.write(gs.sheet2,16,9,'%.4f'%M1NYS_T)
            Func.write(gs.sheet2,16,10,'%.4f'%M2NYS_T)
            Func.write(gs.sheet2,16,11,'%.4f'%(M1NYS_T+M2NYS_T))


            print('-------------------------After rounding---------------------------')
            gs.CD=round(gs.CD,1)
            gs.ED=round(gs.ED,1)
            gs.Alfa = round(gs.Alfa,1)
            gs.offsetAngle = round(gs.offsetAngle,1)
            if gs.DriveType=='Reversing':
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
                gs.xm2 = gs.xm1+gs.Alfa
            else:
                gs.offsetAngle = 0
                gs.Alfa = 360
            

            gs.Ra_Rb = float(gs.BC/gs.CD)
            outM1=Func.Output(gs.BC,gs.CD,gs.ED,gs.xm1,gs.A,gs.B,gs.E,gs.F,KBEW=gs.KBEW)  #Master -UWL,[alpha,beta,NYS_T,NYS_A,NYK_T,NYK_A,C[0][0],C[0][1],C[0][2],Db[0][0],Db[0][1],Db[0][2]]
            outM2=Func.Output(gs.BC,gs.CD,gs.ED,gs.xm2,gs.A,gs.B,gs.E,gs.F,KBEW= gs.KBEW) #Master-OWL [alpha,beta,NYS_T,NYS_A,NYK_T,NYK_A,C[0][0],C[0][1],C[0][2],Db[0][0],Db[0][1],Db[0][2]]

            gs.xs1 = outM1[1]
            gs.xs2 = outM2[1]
            cxs1 = math.cos(Func.rad(gs.xs1))
            cxs2 = math.cos(Func.rad(gs.xs2))
            sxs1 = math.sin(Func.rad(gs.xs1))
            sxs2 = math.sin(Func.rad(gs.xs2))
            gs.w2opt = Func.wipAngle(gs.xs2,gs.xs1,gs.KBEW)


                
            
            M1NYS_T=outM1[2]
            M2NYS_T=outM2[2]
            M1NYK_T=outM1[4]
            M2NYK_T=outM2[4]
            Func.write(gs.sheet2,17,2,gs.ED)
            Func.write(gs.sheet2,17,3,gs.CD)
            Func.write(gs.sheet2,17,5,'%.4f'%outM1[1])
            Func.write(gs.sheet2,17,6,'%.4f'%outM2[1])
            Func.write(gs.sheet2,17,7,'%.4f'%gs.w2opt)
            Func.write(gs.sheet2,17,9,'%.4f'%M1NYS_T)
            Func.write(gs.sheet2,17,10,'%.4f'%M2NYS_T)
            Func.write(gs.sheet2,17,11,'%.4f'%(M1NYS_T+M2NYS_T))
            Func.write(gs.sheet2,19,3,'%.4f'%gs.w2opt)
            Func.write(gs.sheet2,19,9,'%.4f'%M1NYS_T)
            Func.write(gs.sheet2,19,10,'%.4f'%M2NYS_T)
            self.TableOpt.setItem(1,0,Qitem('%.1f'%gs.ED))
            self.TableOpt.setItem(1,1,Qitem('%.1f'%gs.CD))
            self.TableOpt.setItem(1,2,Qitem('%.2f'%outM1[1]))
            self.TableOpt.setItem(1,3,Qitem('%.2f'%outM2[1]))
            self.TableOpt.setItem(1,4,Qitem('%.2f'%gs.w2opt))
            self.TableOpt.setItem(1,5,Qitem('%.2f'%(M1NYS_T)))
            self.TableOpt.setItem(1,6,Qitem('%.2f'%(M2NYS_T)))
            self.TableOpt.setItem(1,7,Qitem('%.2f'%(M1NYS_T+M2NYS_T)))
            self.TableOpt.setItem(2,1,Qitem('%.2f'%gs.w2opt))
            self.TableOpt.setItem(2,6,Qitem('%.2f'%M1NYS_T))
            self.TableOpt.setItem(2,7,Qitem('%.2f'%M2NYS_T))
            print('After Rounding:CD=%.1f'% gs.CD +'\tED=%.1f'% gs.ED +'\tNYS_T1=%.4f'% M1NYS_T +'\tNYS_T2=%.4f'% M2NYS_T)
            print('NYK_T1={:.4f},  NYK_T2= {:.4f} ' .format(M1NYK_T,M2NYK_T))
            print('w2=%.4f'%gs.w2opt+'\tw2Target=%.4f'%(180/math.pi*gs.w2))
            if gs.DriveType=='Reversing':
                print('offsetAngle = {:.4f}, Reversing angle={:.4f}'.format(gs.offsetAngle,gs.Alfa))
            #=============================================================================

            #slave  link
            print('------------------------------slave link info------------------------')
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
                
            # model2.sxs12 =  pe.Var(initialize =0,bounds=(-1,1)) 
            # model2.cxs12 = pe.Var(initialize = 1,bounds=(-1,1))
            # model2.sxs22 =  pe.Var(initialize =0,bounds=(-1,1)) 
            # model2.cxs22 = pe.Var(initialize = -1,bounds=(-1,1))
            model2.bc2 = pe.Var(initialize =gs.BC,bounds=(0.25*BC,4*BC) ) #link length
            model2.Con = pe.ConstraintList()
            model2.obj=pe.ObjectiveList()
            # =============================================================================
            model2.obj.add(expr=(100*(model2.cxs22*model2.cxs12+model2.sxs22*model2.sxs12-math.cos(gs.w3))**2+(eval(F_S1NYS_T)+eval(F_S2NYS_T))**2+(eval(F_S1NYK_T)+eval(F_S2NYK_T))**2)) # minize wipping angel 
            # =============================================================================
            #==============================================================================
            # model2.Con.add(expr=1000*(model2.cxs22*model2.cxs12+model2.sxs22*model2.sxs12-math.cos(w3))>=0.0) #wipping angle requirement 
            # model2.Con.add(eval(F12NYS_T)+eval(F22NYS_T)>=0)
            # model2.Con.add(eval(F12NYK_T)+eval(F22NYK_T)>=0)
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
            t1=eval(F_S1NYS_T)
            t2=eval(F_S2NYS_T)
            t3=eval(F_S1NYK_T)
            t4=eval(F_S2NYK_T)
            #==============================================================================
            # model2.Con.add(eval(F12NYS_T)+eval(F22NYS_T)==0)
            # #==============================================================================
            # #==============================================================================
            # model2.Con.add(eval(F12NYK_T)+eval(F22NYK_T)==0)
            #==============================================================================
            # =============================================================================
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
            gs.BC2=pe.value(model2.bc2)
            gs.xs22=Func.getDegree(cxs22,sxs22)
            gs.xs12=Func.getDegree(cxs12,sxs12)
            gs.xm22=Func.getDegree(cxm22,sxm22)
            gs.xm12=Func.getDegree(cxm12,sxm12)
            gs.Delta2=Func.getDegree(cDelta2,sDelta2)
            gs.CD2=  (pe.value((model2.link2)))
            gs.ED2=(pe.value(model2.outCrank2))
            S1NYS_Tv = eval(F_S1NYS_T)
            S2NYS_Tv = eval(F_S2NYS_T)
            S1NYK_Tv = eval(F_S1NYK_T)
            S2NYK_Tv = eval(F_S2NYK_T)
            S1NYS_T=Func.degree(math.pi/2-math.acos(pe.value(S1NYS_Tv)))
            S2NYS_T=Func.degree(math.pi/2-math.acos(pe.value(S2NYS_Tv)))
            S1NYK_T=Func.degree(math.pi/2-math.acos(pe.value(S1NYK_Tv)))
            S2NYK_T=Func.degree(math.pi/2-math.acos(pe.value(S2NYK_Tv)))


        # write and rounding 
        
        gs.w3opt = Func.wipAngle(gs.xs22,gs.xs12,gs.KBEW2)

        print ('----------------------------slave link info-----------------------------')
        print( 'BC2=%.4f'%gs.BC2+'CD2=%.4f'% gs.CD2 + '\tED2=%.4f'% gs.ED2
              +'\nDelta=%.4f'%gs.Delta2 +'\tD=%.4f'%gs.Distance2)
        print('NYS_T1=%.4f'%S1NYS_T+'\tNYS_T2=%.4f'%S2NYS_T)
        print('NYK_T1=%.4f'%S1NYK_T+'\tNYK_T2=%.4f'%S2NYK_T)
        print('w3=%.4f'%gs.w3opt+'\tw3Target=%.4f'%(180/math.pi*gs.w3))
        

        if gs.MechanicType =='Center':
            Func.write(gs.sheet2,37,2,gs.ED2)
            self.TableOpt2.setItem(0,0,Qitem('%.4f'%gs.ED2))
        else:
            Func.write(gs.sheet2,37,2,gs.BC2)
            self.TableOpt2.setItem(0,0,Qitem('%.4f'%gs.BC2))


        self.TableOpt2.setItem(0,1,Qitem('%.4f'%gs.CD2))
        self.TableOpt2.setItem(0,2,Qitem('%.4f'%gs.xs12))
        self.TableOpt2.setItem(0,3,Qitem('%.4f'%gs.xs22))
        self.TableOpt2.setItem(0,4,Qitem('%.4f'%gs.w3opt))
        self.TableOpt2.setItem(0,5,Qitem('%.4f'%S1NYS_T))
        self.TableOpt2.setItem(0,6,Qitem('%.4f'%S2NYS_T))
        self.TableOpt2.setItem(0,7,Qitem('%.4f'%(S1NYS_T+S2NYS_T)))
        Func.write(gs.sheet2,29,6,gs.w3Target)
        Func.write(gs.sheet2,37,3,gs.CD2)
        Func.write(gs.sheet2,37,5,gs.xs12)
        Func.write(gs.sheet2,37,6,gs.xs22)
        Func.write(gs.sheet2,37,7,gs.w3opt)
        Func.write(gs.sheet2,37,9,S1NYS_T)
        Func.write(gs.sheet2,37,10,S2NYS_T)
        Func.write(gs.sheet2,37,11,(S1NYS_T+S2NYS_T))

        ###rounding
        gs.CD2=round(gs.CD2,1)
        if gs.MechanicType =='Center':
            gs.ED2= round(gs.ED2,1)
        else:
            gs.BC2=round(gs.BC2,1)
            
        gs.Ra_Rb2=float(gs.BC2/gs.CD2)
        outS3=Func.Output(gs.BC2,gs.CD2,gs.ED2,gs.xm12,gs.A2,gs.B2,gs.E2,gs.F2,KBEW=gs.KBEW2) #slave UWL
        outS4=Func.Output(gs.BC2,gs.CD2,gs.ED2,gs.xm22,gs.A2,gs.B2,gs.E2,gs.F2,KBEW=gs.KBEW2) #slave OWL
        #[alpha,beta,NYS_T,NYS_A,NYK_T,NYK_A,C[0][0],C[0][1],C[0][2],D[0][0],D[0][1],D[0][2]]
        gs.w3opt=Func.angleDiff(outS4[1],outS3[1])
        S1NYS_T=outS3[2]
        S2NYS_T=outS4[2]
        self.TextW2.setText('%.4f'%gs.w2Target)
        self.TextW3.setText('%.4f'%gs.w3Target)
        if gs.MechanicType =='Center':
            Func.write(gs.sheet2,38,2,gs.ED2)
            Func.write(gs.sheet2, 31, 6, 'ED')
            self.TableOpt2.setItem(1,0,Qitem('%.1f'%gs.ED2))
            self.textEdit_3.setText('ED')
        else:
            Func.write(gs.sheet2,38,2,gs.BC2)
            Func.write(gs.sheet2, 31, 6, 'BC')
            self.textEdit_3.setText('BC')
            self.TableOpt2.setItem(1,0,Qitem('%.1f'%gs.BC2))
        Func.write(gs.sheet2,38,3,gs.CD2)
        Func.write(gs.sheet2,38,5,'%.2f'%outS3[1])
        Func.write(gs.sheet2,38,6,'%.2f'%outS4[1])
        Func.write(gs.sheet2,38,7,'%.2f'%gs.w3opt)
        Func.write(gs.sheet2,38,9,'%.2f'%S1NYS_T)
        Func.write(gs.sheet2,38,10,'%.2f'%S2NYS_T)
        Func.write(gs.sheet2,38,11,'%.2f'%(S1NYS_T+S2NYS_T))
        Func.write(gs.sheet2,40,3,'%.2f'%gs.w3opt)
        Func.write(gs.sheet2,40,9,'%.2f'%S1NYS_T)
        Func.write(gs.sheet2,40,10,'%.2f'%S2NYS_T)

        self.TableOpt2.setItem(1,1,Qitem('%.1f'%gs.CD2))
        self.TableOpt2.setItem(1,2,Qitem('%.2f'%outS3[1]))
        self.TableOpt2.setItem(1,3,Qitem('%.2f'%outS4[1]))
        self.TableOpt2.setItem(1,4,Qitem('%.2f'%gs.w3opt))
        self.TableOpt2.setItem(1,5,Qitem('%.2f'%S1NYS_T))
        self.TableOpt2.setItem(1,6,Qitem('%.2f'%S2NYS_T))
        self.TableOpt2.setItem(1,7,Qitem('%.2f'%(S1NYS_T+S2NYS_T)))
        self.TableOpt2.setItem(2,1,Qitem('%.2f'%gs.w3opt))
        self.TableOpt2.setItem(2,6,Qitem('%.2f'%S1NYS_T))
        self.TableOpt2.setItem(2,7,Qitem('%.2f'%S2NYS_T))

        print('After Rounding:'+'BC2=%.1f'%gs.BC2+'\tCD2=%.1f'%gs.CD2
              +'\nDelta=%.4f'%gs.Delta2+'\tED2=%.1f'%gs.ED2)

        print('NYS_T1=%.4f'%S1NYS_T+'\tNYS_T2=%.4f'%S2NYS_T)
        print('NYK_T1=%.4f'%S1NYK_T+'\tNYK_T2=%.4f'%S2NYK_T)
        print('w3=%.4f'%gs.w3opt+'\tw3Target=%.4f'%(180/math.pi*gs.w3))
        outM1=[float(s) for s in outM1]
        outM2=[float(s) for s in outM2]
        outS3=[float(s) for s in outS3]
        outS4=[float(s) for s in outS4]
        # write to output sheet

        if gs.startFrom == 'APS1':
            gs.zeroAngle = gs.xm1-gs.APS1
            gs.offsetAngle = gs.offsetAngle-gs.APS1
            gs.Alfa = gs.Alfa+gs.APS1
        elif gs.startFrom =='APS2':
            gs.zeroAngle = gs.xm1-gs.APS2
            gs.offsetAngle = gs.offsetAngle-gs.APS2
            gs.Alfa = gs.Alfa+gs.APS2
        else:
            gs.zeroAngle = gs.xm1
            gs.Alfa = gs.Alfa
            gs.offsetAngle = gs.offsetAngle
        
        print('--------optimization completed------------')


class MainWindow_Input(QtWidgets.QMainWindow,Ui_Input):
    def __init__(self,parent=None):
        super(MainWindow_Input,self).__init__(parent)
        self.setupUi(self)
        self.initiInput()
# =============================================================================
#     def blankState(self):
#         self.tableMasterCranklInfo.item(4,1).setflags()
#         self.tableMasterCranklInfo.item(5,1).setflags((ItemFlags) 0)
# =============================================================================

    def UpdateOptValue(self):
        print('update opt value')

        Func.writeNewValueGUI(self.tableMasterCranklInfo,4,0,gs.CD)
        Func.writeNewValueGUI(self.tableMasterCranklInfo,5,0,gs.ED)
        Func.writeNewValueGUI(self.tableMasterCrankInfo2,0,1,'%.4f'%gs.Ra_Rb)
        Func.writeNewValueGUI(self.tableMasterCrankInfo2,0,2,'%.4f'%gs.w2opt)
        Func.writeNewValueGUI(self.tableSlaveCrankInfo2,0,1,'%.4f'%gs.Ra_Rb2)
        Func.writeNewValueGUI(self.tableSlaveCrankInfo2,0,2,'%.4f'%gs.w3opt)
        Func.writeNewValueGUI(self.tableMasterCranklInfo,1,0,'%.4f'%gs.Delta)
        Func.writeNewValueGUI(self.tableSlaveCrankInfo,1,0,'%.4f'%gs.Delta2)
        #self.tableMasterCranklInfo.setItem(4, 0 ,Qitem(str(gs.CD)))

        if gs.DriveType == 'Reversing':
            Func.writeNewValueGUI(self.tableMotorInfo,0,0,'%.4f'%gs.Alfa)
            Func.writeNewValueGUI(self.tableMotorInfo,1,0,'%.4f'%gs.offsetAngle)



        if gs.MechanicType =='Center':
            Func.writeNewValueGUI(self.tableSlaveCrankInfo,4,0,gs.CD2)
            Func.writeNewValueGUI(self.tableSlaveCrankInfo,5,0,gs.ED2)

        else :
            Func.writeNewValueGUI(self.tableSlaveCrankInfo,4,0,gs.CD2)
            Func.writeNewValueGUI(self.tableSlaveCrankInfo,3,0,gs.BC2)
       
            
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
        self.comboBoxNo2.setCurrentText   (Func.read(gs.sheet1,35,7).strip())
        self.comboBoxNo3.setCurrentText   (Func.read(gs.sheet1,52,7).strip())
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
        gs.w2actual = float(Func.read(gs.sheet1,49,17))
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

        gs.Master =  QtWidgets.QComboBox.currentText(self.comboBoxNo2)
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
        Func.write(gs.sheet1,49,17,gs.w2actual)

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

        gs.excel_out=os.getcwd()+'\\output\\'+gs.ProjectName+'Report_'+gs.styleTime+'.xlsx'
        
        [gs.excel_design1, gs.excel_design2] = gs.DesignPath[gs.MechanicType]
        gs.wb1 =  xl.load_workbook( gs.excel_design1 )
        gs.sheetDesign1 = gs.wb1['Sheet1']
        gs.wb2 =  xl.load_workbook( gs.excel_design2 )
        gs.sheetDesign2 = gs.wb2['Sheet1']
        # gs.excel_design1=os.getcwd()+'\\output\\'+gs.ProjectName+'_DesignTable1_'+gs.styleTime+'.xlsx'
        # gs.excel_design2=os.getcwd()+'\\output\\'+gs.ProjectName+'_DesignTable2_'+gs.styleTime+'.xlsx'
        gs.wb.save(gs.template_path)
        print('save input configuration')
                # pre-treatment
        degree=[gs.w2,gs.w3,gs.Alfa,gs.errorP.Alfa ,gs.errorN.Alfa ,gs.Offset,gs.errorP.Offset ,gs.errorN.Offset,gs.APS1,
        gs.APS2,gs.IPL,gs.UWL,gs.ALP,gs.SP ,gs.KMP ]
        for i in range(len(degree)):
            if type(degree[i])==str:
                print('%s is  not a number ! will be set to 0'%degree[i])
                degree[i]=0
                
        
        radius = [x*math.pi/180 for x in degree]
        [gs.w2,gs.w3,gs.Alfa,gs.errorP.Alfa ,gs.errorN.Alfa ,gs.Offset,gs.errorP.Offset ,gs.errorN.Offset,gs.APS1,
        gs.APS2,gs.IPL,gs.UWL,gs.ALP,gs.SP ,gs.KMP ]=radius
         
        gs.A = np.array([gs.A_X,gs.A_Y,gs.A_Z])
        gs.Ap = np.array([gs.Ap_X,gs.Ap_Y,gs.Ap_Z])
        gs.F = np.array([gs.F_X,gs.F_Y,gs.F_Z])
        gs.Fp = np.array([gs.Fp_X,gs.Fp_Y,gs.Fp_Z])
        gs.F2 = np.array([gs.F_X2,gs.F_Y2,gs.F_Z2])
        gs.Fp2 = np.array([gs.Fp_X2,gs.Fp_Y2,gs.Fp_Z2])
        gs.B = Func.point(gs.A,gs.Ap,gs.Distance)
        gs.E = Func.point(gs.F,gs.Fp,gs.FE)
        gs.E2 = Func.point(gs.F2,gs.Fp2,gs.FE2)
        
        if gs.MechanicType =='Center':
            gs.A2 = gs.A.copy()
            gs.Ap2 = gs.Ap.copy()
            gs.A_X2 = gs.A_X
            gs.A_Y2 = gs.A_Y
            gs.A_Z2 = gs.A_Z
            gs.Ap_X2 = gs.Ap_X
            gs.Ap_Y2 = gs.Ap_Y
            gs.Ap_Z2 = gs.Ap_Z
            gs.errorP.A_X2 = gs.errorP.A_X
            gs.errorP.A_Y2 = gs.errorP.A_Y
            gs.errorP.A_Z2 = gs.errorP.A_Z
            gs.errorP.Ap_X2 = gs.errorP.Ap_X
            gs.errorP.Ap_Y2 = gs.errorP.Ap_Y
            gs.errorP.Ap_Z2 = gs.errorP.Ap_Z
            gs.errorN.A_X2 = gs.errorN.A_X
            gs.errorN.A_Y2 = gs.errorN.A_Y
            gs.errorN.A_Z2 = gs.errorN.A_Z
            gs.errorN.Ap_X2 = gs.errorN.Ap_X
            gs.errorN.Ap_Y2 = gs.errorN.Ap_Y
            gs.errorN.Ap_Z2 = gs.errorN.Ap_Z
            gs.Distance2=(gs.Distance+gs.Distance2)
            gs.B2 = Func.point(gs.A2,gs.Ap2,gs.Distance2)
        else:
            gs.A2 = gs.E.copy()
            gs.Ap2 = gs.F.copy()
            gs.A_X2 = gs.E_X
            gs.A_Y2 = gs.E_Y
            gs.A_Z2 = gs.E_Z
            gs.Ap_X2 = gs.F_X
            gs.Ap_Y2 = gs.F_Y
            gs.Ap_Z2 = gs.F_Z
            gs.errorP.A_X2 = gs.errorP.Fp_X
            gs.errorP.A_Y2 = gs.errorP.Fp_Y
            gs.errorP.A_Z2 = gs.errorP.Fp_Z
            gs.errorP.Ap_X2 = gs.errorP.F_X
            gs.errorP.Ap_Y2 = gs.errorP.F_Y
            gs.errorP.Ap_Z2 = gs.errorP.F_Z
            gs.errorN.A_X2 = gs.errorN.Fp_X
            gs.errorN.A_Y2 = gs.errorN.Fp_Y
            gs.errorN.A_Z2 = gs.errorN.Fp_Z
            gs.errorN.Ap_X2 = gs.errorN.F_X
            gs.errorN.Ap_Y2 = gs.errorN.F_Y
            gs.errorN.Ap_Z2 = gs.errorN.F_Z
            gs.B2 = Func.point(gs.A2,gs.Ap2,gs.Distance2)
            
        # else:TBD

        #==============================================================================

        
        # =============================================================================
        if not gs.BC2:
            logger.warning('please give a correct BC2 value')
        #==============================================================================
        # os.system("pause")
        #==============================================================================
        if gs.startFrom=='EPS':
            gs.startAngle=gs.Offset
        elif gs.startFrom=="APS1":
            gs.startAngle=gs.Offset+gs.APS1
        elif gs.startFrom=='Park':
            logger.warning('TO be defined what park position mean')
        else:
            logger.warning('pls configure where measured from')
            
        gs.Alfa=np.sign(gs.isClock)*gs.Alfa
        print('load comple')
        print('load completed')

import uuid
#import hashlib
import time
import base64
import os
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
    if (mac == macL)&(second<=deadLine) :# ok
        print('ok')
        return 1
    else:
        print ('pls update license')
        return 0

# generate license
mac = get_mac_address()
deadLine = str(time.mktime(time.strptime("2019-01-01","%Y-%m-%d")))  #deadline date
createLicense(mac,deadLine)

# interpret license

# m = hashlib.md5()
# m.update(mac.encode('utf8'))
# macMD5 = m.hexdigest()



if __name__ == "__main__":

        
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)  
    rq = time.strftime('%Y%m%d%H%M', time.localtime(time.time()))
    log_path = os.getcwd() + '/Logs/'
    if not os.path.exists(log_path):
        os.makedirs(log_path)

    log_name = log_path + rq + '.log'
    logfile = log_name
    fh = logging.FileHandler(logfile, mode='w')
    fh.setLevel(logging.DEBUG) 

    formatter = logging.Formatter("%(asctime)s - %(filename)s[line:%(lineno)d] - %(levelname)s: %(message)s")
    fh.setFormatter(formatter)
    logger.addHandler(fh)

    timeStamp=int(time.time())
    timeArray=time.localtime(timeStamp)
    gs.styleTime=time.strftime("%Y%m%d_%H%M%S",timeArray)

    class error(object):
        Offset=0
    gs.errorP = error()
    gs.errorN = error()
    flag = checkLicense()
    if flag == 1 :
        app = QtWidgets.QApplication(sys.argv)
        mainWindow = MainWindow()
        mainWindow.show()
    
    else:
        print('pls update license')



    sys.exit(app.exec_())
    sys.exit(0)
# =============================================================================



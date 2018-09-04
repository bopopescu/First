import sys
import random

import matplotlib
import matplotlib.pyplot as plt
matplotlib.use("Qt5Agg")

from PyQt5 import QtCore
from PyQt5.QtWidgets import QApplication, QMainWindow, QMenu, QVBoxLayout, QSizePolicy, QMessageBox, QWidget

from numpy import arange, sin, pi
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import time
import Func
import parameters
class MyMplCanvas(FigureCanvas):
    """这是一个窗口部件，即QWidget（当然也是FigureCanvasAgg）"""
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = fig.add_subplot(111)
        # 每次plot()调用的时候，我们希望原来的坐标轴被清除(所以False)
        #self.axes.hold(False)

        self.compute_initial_figure()

        #
        FigureCanvas.__init__(self, fig)
        self.setParent(parent)

        FigureCanvas.setSizePolicy(self,
                                   QSizePolicy.Expanding,
                                   QSizePolicy.Expanding)
        FigureCanvas.updateGeometry(self)

    def compute_initial_figure(self):
        pass

class MyStaticMplCanvas(MyMplCanvas):
    """静态画布：一条正弦线"""
    def first_figure(self):
            matplotlib.matplotlib_fname()
            x=gs.outKPIall['alpha']
            fig = plt.figure(1)
            ax1 = fig.add_subplot(2,2,1)
            ax1.plot(x,gs.outKPIall['beta'],'k--',label='Beta')
            ax1.plot(x,gs.outKPIall['beta2'],'r--',label='Gamma')
            
            ax2 = fig.add_subplot(2,2,2)
            ax2.plot(x,gs.outKPIall['beta_s'],'k--',label='B-S')
            ax2.plot(x,gs.outKPIall['beta_s2'],'r:',label='G-S')
            
            ax3 = fig.add_subplot(2,2,3)
            ax3.plot(x,gs.outKPIall['beta_ss'],'k--',label='B-SS')
            ax3.plot(x,gs.outKPIall['beta_ss2'],'r:',label='G-SS')
            ax1.legend()
            ax2.legend()
            ax3.legend()
            fig.show()
            fig.savefig('.\\output\\pics\\'+gs.ProjectName+'_Angle'+gs.styleTime+'.jpg')

            #plt.figure(2)

    def second_figure(self):
            x=gs.outKPIall['alpha']
            fig2,axs = plt.subplots(2,2)
            axs[0,0].plot(x,gs.outKPIall['NYS_T'],'k--',label='NYS_T')
            axs[0,0].plot(x,gs.outKPIall['NYS_T2'],'r:',label='NYS_T2')

            axs[1,0].plot(x,gs.outKPIall['NYK_T'],'k--',label='NYK_T')
            axs[1,0].plot(x,gs.outKPIall['NYK_T2'],'r:',label='NYK_T2')

            axs[0,1].plot(x,gs.outKPIall['NYS_A'],'k--',label='NYS_A')
            axs[0,1].plot(x,gs.outKPIall['NYS_A2'],'r:',label='NYS_A2')

            axs[1,1].plot(x,gs.outKPIall['NYK_A'],'k--',label='NYK_A')
            axs[1,1].plot(x,gs.outKPIall['NYK_A2'],'r:',label='NYK_A2')
            for row in axs:
                for col in row:
                    col.legend()

            fig2.show()
            fig2.savefig('.\\output\\pics\\'+gs.ProjectName+'_NY'+gs.styleTime+'.jpg')


class MyDynamicMplCanvas(MyMplCanvas):
    """动态画布：每秒自动更新，更换一条折线。"""
    def __init__(self, *args, **kwargs):
        MyMplCanvas.__init__(self, *args, **kwargs)
        timer = QtCore.QTimer(self)
        timer.timeout.connect(self.update_figure)
        timer.start(1000)

    def compute_initial_figure(self):
        self.axes.plot([0, 1, 2, 3], [1, 2, 0, 4], 'r')

    def update_figure(self):
        # 构建4个随机整数，位于闭区间[0, 10]
        l = [random.randint(0, 10) for i in range(4)]

        self.axes.plot([0, 1, 2, 3], l, 'r')
        self.draw()

class ApplicationWindow(QMainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.setWindowTitle("程序主窗口")

        self.file_menu = QMenu('&File', self)
        self.file_menu.addAction('&Quit', self.fileQuit,
                                 QtCore.Qt.CTRL + QtCore.Qt.Key_Q)
        self.menuBar().addMenu(self.file_menu)

        self.help_menu = QMenu('&Help', self)
        self.menuBar().addSeparator()
        self.menuBar().addMenu(self.help_menu)

        self.help_menu.addAction('&About', self.about)

        self.main_widget = QWidget(self)

        l = QVBoxLayout(self.main_widget)
        sc = MyStaticMplCanvas(self.main_widget, width=5, height=4, dpi=100)
        dc = MyDynamicMplCanvas(self.main_widget, width=5, height=4, dpi=100)
        sc.first_figure()
        sc.second_figure()
        l.addWidget(sc)
        l.addWidget(dc)

        self.main_widget.setFocus()
        self.setCentralWidget(self.main_widget)
        # 状态条显示2秒
        self.statusBar().showMessage(str(time.time())+'   k5la')

    def fileQuit(self):
        self.close()

    def closeEvent(self, ce):
        self.fileQuit()

    def about(self):
        QMessageBox.about(self, "About",
        """embedding_in_qt5.py example
        Copyright 2015 BoxControL

        This program is a simple example of a Qt5 application embedding matplotlib
        canvases. It is base on example from matplolib documentation, and initially was
        developed from Florent Rougon and Darren Dale.

        http://matplotlib.org/examples/user_interfaces/embedding_in_qt4.html

        It may be used and modified with no restriction; raw copies as well as
        modified versions may be distributed without limitation.
        """
        )

if __name__ == '__main__':
    app = QApplication(sys.argv)
    outputs= parameters.Outputs()
    gs = parameters.GS()
    outjson='output.json'
    Func.LoadJson(outjson,outputs,gs)
    aw = ApplicationWindow()
    aw.setWindowTitle("PyQt5 与 Matplotlib 例子")
    aw.show()
    #sys.exit(qApp.exec_())
    app.exec_()
import sys
import random

import matplotlib
import matplotlib.animation as animation
import matplotlib.pyplot as plt
import mpl_toolkits.mplot3d.axes3d as p3
from mpl_toolkits.mplot3d import proj3d
import numpy as np
matplotlib.use("Qt5Agg")

from PyQt5 import QtCore
from PyQt5.QtWidgets import QApplication, QMainWindow, QMenu, QVBoxLayout, QSizePolicy, QMessageBox, QWidget,QTabWidget

from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import time
import math
import Func
import parameters
matplotlib.matplotlib_fname()
class MyMplCanvas(FigureCanvas):
    """这是一个窗口部件，即QWidget（当然也是FigureCanvasAgg）"""
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        self.fig = plt.Figure(figsize=(width, height), dpi=dpi)
        #self.axes = fig.add_subplot(111)
        # 每次plot()调用的时候，我们希望原来的坐标轴被清除(所以False)
        #self.axes.hold(False)

        self.compute_initial_figure()

        #
        FigureCanvas.__init__(self, self.fig)
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
            
            x=gs.outKPIall['alpha']
            
            ax1 = self.fig.add_subplot(2,2,1)
            ax1.plot(x,gs.outKPIall['beta'],'k--',label='Beta')
            ax1.plot(x,gs.outKPIall['beta2'],'r--',label='Gamma')
            ax1.set_title('wiping angle v.s motor angle')
            ax2 = self.fig.add_subplot(2,2,2)
            ax2.plot(x,gs.outKPIall['beta_s'],'k--',label='B-S')
            ax2.plot(x,gs.outKPIall['beta_s2'],'r:',label='G-S')
            ax2.set_title('wiping angle velocity v.s motor angle')
            ax3 = self.fig.add_subplot(2,2,3)
            ax3.plot(x,gs.outKPIall['beta_ss'],'k--',label='B-SS')
            ax3.plot(x,gs.outKPIall['beta_ss2'],'r:',label='G-SS')
            ax3.set_title('wiping angel accelaration v.s motor angle')
            ax1.legend()
            ax2.legend()
            ax3.legend()
           
            self.fig.savefig('.\\output\\pics\\'+gs.ProjectName+'_Angle'+gs.styleTime+'.jpg')

            #plt.figure(2)

    def second_figure(self):
            x=gs.outKPIall['alpha']
            axs1 = self.fig.add_subplot(2,2,1)
            axs1.plot(x,gs.outKPIall['NYS_T'],'k--',label='NYS_T')
            axs1.plot(x,gs.outKPIall['NYS_T2'],'r:',label='NYS_T2')
            axs1.set_title('NYS_T v.s motro angle')
            axs2 = self.fig.add_subplot(2,2,2)
            axs2.plot(x,gs.outKPIall['NYK_T'],'k--',label='NYK_T')
            axs2.plot(x,gs.outKPIall['NYK_T2'],'r:',label='NYK_T2')
            axs2.set_title('NYK_T v.s motro angle')
            axs3 = self.fig.add_subplot(2,2,3)
            axs3.plot(x,gs.outKPIall['NYS_A'],'k--',label='NYS_A')
            axs3.plot(x,gs.outKPIall['NYS_A2'],'r:',label='NYS_A2')
            axs3.set_title('NYS_A v.s motro angle')
            axs4 = self.fig.add_subplot(2,2,4)
            axs4.plot(x,gs.outKPIall['NYK_A'],'k--',label='NYK_A')
            axs4.plot(x,gs.outKPIall['NYK_A2'],'r:',label='NYK_A2')
            axs4.set_title('NYK_A v.s motro angle')
            axs1.legend()
            axs2.legend()
            axs3.legend()
            axs4.legend()
            self.fig.savefig('.\\output\\pics\\'+gs.ProjectName+'_NY'+gs.styleTime+'.jpg')


         # temporialy test
    @staticmethod
    def generateData(): # generate data to animate
            number = gs.num 
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
            return data
                
    @staticmethod
    def getMinMax(data): # get min/max az,el for 3d plot
            cordinateMin = np.max(data[0],axis=1)
            cordinateMax = np.max(data[0],axis=1)
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
            
            abScale = [ab[0]/xxrange,ab[1]/yrange,ab[2]/zrange] #scale axis equal
            az=-90-Func.degree(math.atan(abScale[0]/abScale[1]))
            el=-np.sign(ab[1])*Func.degree(math.atan(abScale[2]/math.sqrt(abScale[0]**2+abScale[1]**2)))
            return [[xmin,xmax,ymin,ymax,zmin,zmax],[az,el]]

         #TBD
        # initial point
    @staticmethod
    def orthogonal_proj(zfront, zback):
            a = (zfront+zback)/(zfront-zback)
            b = -2*(zfront*zback)/(zfront-zback)
            return np.array([[1,0,0,0],
                                [0,1,0,0],
                                [0,0,a,b],
                                [0,0, -0.0001,zback]])

    def animationSteady(self):
            proj3d.persp_transformation = self.orthogonal_proj
            data= self.generateData()
            [[xmin,xmax,ymin,ymax,zmin,zmax],[az,el]]= self.getMinMax(data)
            ax2 = p3.Axes3D(self.fig)
            ax2.set_xlabel('x')
            ax2.set_ylabel('y')
            ax2.set_zlabel('z')
            Label=['AB','BC','CD','DE','EF','AB2','BC2','CD2','DE2','EF2']
            colors ='bgcmkbgcmk'
            linestyles=['-','-','-','-','-','--','--','--','--','--']

            ax2.set_xlim3d(xmin,xmax)
            ax2.set_ylim3d(ymin,ymax)
            ax2.set_zlim3d(zmin,zmax)
            
            ax2.set_title('initial position')
            #ax2.axis('scaled')
            ax2.view_init(azim=az, elev=el)
            lines = [ax2.plot([data[i][0,0],data[i][3,0]],[data[i][1,0],data[i][4,0]],[data[i][2,0],data[i][5,0]],label=Label[i],color=colors[i],linestyle=linestyles[i])[0] for i  in range(10)]
            self.fig.legend()
        # Attaching 3D axis to the figure
    def animation(self):

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
            number = 12
            
            proj3d.persp_transformation = self.orthogonal_proj
            data= self.generateData()
            [[xmin,xmax,ymin,ymax,zmin,zmax],[az,el]]= self.getMinMax(data)
            ax = p3.Axes3D(self.fig)
            gs.pause =False
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

            # Creating the Animation object
            Label=['AB','BC','CD','DE','EF','AB2','BC2','CD2','DE2','EF2']
            colors ='bgcmkbgcmk'
            linestyles=['-','-','-','-','-','--','--','--','--','--']
            lines = [ax.plot([data[i][0,0],data[i][3,0]],[data[i][1,0],data[i][4,0]],[data[i][2,0],data[i][5,0]],label=Label[i],color=colors[i],linestyle=linestyles[i])[0] for i  in range(10)]
            #start of each line
            self.fig.canvas.mpl_connect('button_press_event', onClick)
            gs.line_ani=animation.FuncAnimation(self.fig, update_lines, number, fargs=(data, lines,ax),interval=int(3600/number), blit=True)
            #==============================================================================
            # plt.rcParams['animation.ffmpeg_path']='G:\\wai\\ffmpeg\\bin\\ffmpeg.exe'
            #==============================================================================
            self.fig.legend()
            #plt.show()
            #return gs.line_ani 

class ApplicationWindow(QTabWidget):
    def __init__(self):
        QMainWindow.__init__(self)
        self.showMaximized()
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.setWindowTitle("plot&anmiation")

        self.widget1 = QWidget(self)
        self.widget2 = QWidget(self)
        self.widget3 = QWidget(self)
        self.widget4 = QWidget(self)
    
        self.addTab(self.widget1,u"Wiping angle curve")
        self.addTab(self.widget2,u"NY angle curve")
        self.addTab(self.widget3,u"initial position ")
        self.addTab(self.widget4,u"animation ")
        
        


        
        sc1 = MyStaticMplCanvas(self.widget1, width=8, height=6, dpi=100)
        sc2 = MyStaticMplCanvas(self.widget2, width=8, height=6, dpi=100)
        sc3 = MyStaticMplCanvas(self.widget3, width=8, height=6, dpi=100)
        sc4 = MyStaticMplCanvas(self.widget4, width=8, height=6, dpi=100)
        
        sc1.first_figure()
        sc2.second_figure()
        sc3.animationSteady()
        sc4.animation()

        l1 = QVBoxLayout(self.widget1)
        l2 = QVBoxLayout(self.widget2)
        l3 = QVBoxLayout(self.widget3)
        l4 = QVBoxLayout(self.widget4)
        
        l1.addWidget(sc1)
        l2.addWidget(sc2)
        l3.addWidget(sc3)
        l4.addWidget(sc4)

        self.widget1.setFocus()
        #self.settCentralWidget(self.widget1)
        # 状态条显示2秒
        #self.statusBar().showMessage(str(time.time())+'   k5la')

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
    aw.show()
    #sys.exit(qApp.exec_())
    app.exec_()
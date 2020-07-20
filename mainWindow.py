#!/usr/bin/env python

import sys

from PyQt5.QtCore import QSize
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets import QMainWindow
from PyQt5.QtWidgets import QMenuBar

from processGui import ProcessGui

class MainWindow(QMainWindow):

    def __init__(self, title,email,mdp,server,port):
        super(MainWindow, self).__init__()
        self.setWindowTitle(title) 
        
        icon = QIcon()
        icon.addFile(sys._MEIPASS+'/icon.ico', QSize(256,256))
        self.setWindowIcon(icon)
        
        #menuBar = QMenuBar(self)
        
        #menuFile = menuBar.addMenu('File')
        #menuEdit = menuBar.addMenu('Edit')
        #menuHelp = menuBar.addMenu('Help')

        #self.setMenuBar(menuBar)
        
        processGui = ProcessGui(title,email,mdp,server,port)
        self.setCentralWidget(processGui)
        
  
if __name__ == '__main__':
	
    app = QApplication(sys.argv)
    email = "ali.zakaria@ensea.fr"
    mdp = "###########"
    
    gui = MainWindow("OneEmail",email,mdp,"smtp2.ensea.fr",587)     
    gui.show()
    
    sys.exit(app.exec_())

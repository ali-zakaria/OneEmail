#!/usr/bin/env python
# coding: utf-8

from __future__ import unicode_literals
import sys

import urllib2

import smtplib

from PyQt5.QtCore import QSize
from PyQt5.QtGui import QFont
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtWidgets import QInputDialog
from PyQt5.QtWidgets import QHBoxLayout
from PyQt5.QtWidgets import QLineEdit
from PyQt5.QtWidgets import QPushButton
from PyQt5.QtWidgets import QVBoxLayout
from PyQt5.QtWidgets import QWidget
from PyQt5.QtWidgets import QLabel

from mainWindow import MainWindow

class WelcomeGui(QWidget):

    def __init__(self, title):
        super(WelcomeGui, self).__init__()
        self.title = title
        self.setWindowTitle(title) 
        icon = QIcon()
        icon.addFile('icon.ico', QSize(256,256))
        self.setWindowIcon(icon)        
        self.setFixedSize(300, 150)
        self.vlayout = QVBoxLayout(self)
        font = QFont("Helvetica", 9, QFont.Bold)
        
        self.Hlayout_email = QHBoxLayout()
        email_label = QLabel("Email")
        email_label.setFont(font)
        self.Hlayout_email.addSpacing(10)
        self.Hlayout_email.addWidget(email_label)
        self.Hlayout_email.addSpacing(55)
        self.email_display = QLineEdit()
        self.Hlayout_email.addWidget(self.email_display)
        self.Hlayout_email.addSpacing(10)
        self.vlayout.addLayout(self.Hlayout_email)
        
        self.Hlayout_mdp = QHBoxLayout()
        mdp_label = QLabel("Mot de passe")
        mdp_label.setFont(font)
        self.Hlayout_mdp.addSpacing(10)
        self.Hlayout_mdp.addWidget(mdp_label)
        self.Hlayout_mdp.addSpacing(10)
        
        self.mdp_display = QLineEdit()
        self.mdp_display.setEchoMode(QLineEdit.Password)
        self.Hlayout_mdp.addWidget(self.mdp_display)
        self.Hlayout_mdp.addSpacing(10)
        
        self.vlayout.addLayout(self.Hlayout_mdp)
        self.vlayout.addSpacing(10)
        
        self.Hlayout_button = QHBoxLayout()
        self.conbutton = QPushButton('Se connecter', self)
        self.conbutton.setDefault(True)
        self.conbutton.clicked.connect(self.connect_event)
        self.Hlayout_button.addSpacing(80)
        self.Hlayout_button.addWidget(self.conbutton)
        self.Hlayout_button.addSpacing(80)
        self.vlayout.addLayout(self.Hlayout_button)        

    def internet_on(self):
        try:
            urllib2.urlopen('http://216.58.192.142', timeout=1)
            return True
        except urllib2.URLError as err:
            return False
    
    def warning(self, titleMsg):
        QMessageBox(QMessageBox.Warning,self.title,titleMsg,QMessageBox.NoButton,self).exec_()
    
    def connect_event(self, event):
        if self.internet_on() :
            self.connect()
        else :
            self.warning('Veuillez vérifier votre connexion internet')

    def connect(self):
        email = self.email_display.text()
        mdp = self.mdp_display.text()

        serverSMTP = ""
        serverPORT = ""
        
        a = email.find("@")
        b = email.rfind(".")
        server_name = email[a+1:b]
        
        if a<1 or b-a<2:
            self.warning('Entrer un email valide SVP !')
            
        elif len(mdp)<6:
            self.warning('Entrer un mot de passe valide SVP !')
            
        else:
            
            if server_name == "outlook" :
                serverPORT = "587"
                serverSMTP = 'smtp-mail.outlook.com'

            elif server_name == "gmail" :
                serverSMTP = 'smtp.gmail.fr'
                serverPORT = "587"
            
            elif server_name == "ensea" :
                serverSMTP = 'smtp2.ensea.fr'
                serverPORT = "587"
            
            else :
                serverSMTP = QInputDialog.getText(self,self.title,"Serveur :")[0]
                serverPORT = QInputDialog.getText(self,self.title,"Port :")[0]         

            if len(serverSMTP)>0 and len(serverPORT)>0 :
                serverPORT = int(serverPORT)
                try:
                    server = smtplib.SMTP(serverSMTP, serverPORT)
                    server.starttls()
                    server.login(email, mdp)
                except smtplib.SMTPServerDisconnected:
                    self.warning('La combinaison email/mot de passe est incorrecte.')
                except smtplib.SMTPAuthenticationError:
                    self.warning('Entrer un mot de passe valide SVP !')
                else:
                    self.mainWindow = MainWindow(self.title,email,mdp,serverSMTP,serverPORT)
                    self.mainWindow.show()
                    self.hide()
    
  
if __name__ == '__main__':
	
    app = QApplication(sys.argv)
    
    gui = WelcomeGui("OneEmail")     
    gui.show()
    
    sys.exit(app.exec_())

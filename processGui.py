#!/usr/bin/env python
# coding: utf-8

from __future__ import unicode_literals
import sys

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.header import Header
from email import encoders

from PyQt5.QtCore import QSize
from PyQt5.QtCore import QObject
from PyQt5.QtGui import QFont
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QAction
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtWidgets import QMenu
from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets import QHBoxLayout
from PyQt5.QtWidgets import QLineEdit
from PyQt5.QtWidgets import QPushButton
from PyQt5.QtWidgets import QVBoxLayout
from PyQt5.QtWidgets import QWidget
from PyQt5.QtWidgets import QLabel
from PyQt5.QtWidgets import QTextEdit
from PyQt5.QtWidgets import QProgressDialog

class ProcessGui(QWidget):

    def __init__(self, title, email, mdp,serverSMTP,serverPORT):
        super(ProcessGui, self).__init__()
        self.title = title
        self.email = email
        self.mdp = mdp
        self.serverSMTP = serverSMTP
        self.serverPORT = serverPORT
        
        self.setWindowTitle(self.title)
        
        icon = QIcon()
        icon.addFile('icon.ico', QSize(256,256))
        self.setWindowIcon(icon)
		
        self.setFixedSize(450, 450)
        self.vlayout = QVBoxLayout()
        self.setLayout(self.vlayout)
        
        font = QFont("Helvetica", 9, QFont.Bold)
        Hlayout_dest = QHBoxLayout()
        dest_label = QLabel("À :")
        dest_label.setFont(font)
        Hlayout_dest.addWidget(dest_label)
        Hlayout_dest.addSpacing(20)
        
        self.dest_file = QTextEdit()
        self.dest_file.setText("Écrivez les emails de vos destinataires ligne par ligne !")
        self.dest_file.textChanged.connect(self.onTextChanged)
        self.dest_file.cursorPositionChanged.connect(self.onTextChanged)
        self.dest_file.setFont(QFont("Helvetica", italic = True))
        self.dest_file.setFixedSize(367,45)
        Hlayout_dest.addWidget(self.dest_file)
        Hlayout_dest.addSpacing(10)
        
        self.vlayout.addLayout(Hlayout_dest)
        
        Hlayout_obj = QHBoxLayout()
        obj_label = QLabel("Sujet :")
        obj_label.setFont(font)
        Hlayout_obj.addWidget(obj_label)
        Hlayout_obj.addSpacing(10)
        
        self.obj_display = QLineEdit()
        Hlayout_obj.addWidget(self.obj_display)
        Hlayout_obj.addSpacing(10)
        
        self.vlayout.addLayout(Hlayout_obj)
        
        Hlayout_pj = QHBoxLayout()
        self.pjButton = QPushButton('Joindre :', self)
        self.pjButton.clicked.connect(self.pj_event)
        self.attachments = []
        Hlayout_pj.addWidget(self.pjButton)
        Hlayout_pj.addSpacing(10)
        
        self.pj_display = QLineEdit()
        self.pj_display.setReadOnly(True)
        Hlayout_pj.addWidget(self.pj_display)
        Hlayout_pj.addSpacing(10)
        
        self.vlayout.addLayout(Hlayout_pj)
        self.vlayout.addSpacing(10)
        
        Hlayout_bar = QHBoxLayout()
       
        self.fontFamilyMenu = QPushButton("Tahoma")
        menu = QMenu()
        FontFamily = ['Arial','Arial Black','Calibri','Comic Sans MS','Corbel','Courrier New','Elephant','Georgia','Segoe Script','Tahoma','Times New Roman','Verdana']
        
        for family in FontFamily:
            action = QAction(family,self)
            action.setObjectName(family)
            font = QFont()
            font.setFamily(family)
            action.setFont(font)
            action.triggered.connect(self.updateFontFamily)
            menu.addAction(action)
        
        self.fontFamilyMenu.setMenu(menu)
        
        self.fontSizeMenu = QPushButton("8pt")
        self.fontSizeMenu.setFixedSize(60,23)
        menu = QMenu()
        FontSize = [8,9,10,11,12,13,14,16,18,24,36,48]
        
        for size in FontSize:
            action = QAction(str(size) + 'pt',self)
            action.setObjectName(str(size))
            action.triggered.connect(self.updateFontSize)
            menu.addAction(action)
        
        self.fontSizeMenu.setMenu(menu)
        
        Hlayout_bar.addWidget(self.fontFamilyMenu)
        Hlayout_bar.addWidget(self.fontSizeMenu)
        
        self.boldButton = QPushButton("B")
        self.boldButton.setCheckable(True)
        font = QFont()
        font.setBold(True)
        self.boldButton.setFont(font)
        self.boldButton.setObjectName("B")
        self.boldButton.setFixedSize(30,23)
        self.boldButton.clicked.connect(self.font_event)
        self.italicButton = QPushButton("I")
        self.italicButton.setCheckable(True)
        font = QFont()
        font.setItalic(True)
        self.italicButton.setFont(font)
        self.italicButton.setObjectName("I")
        self.italicButton.setFixedSize(30,23)
        self.italicButton.clicked.connect(self.font_event)
        self.underlinedButton = QPushButton("U")
        self.underlinedButton.setCheckable(True)
        font = QFont()
        font.setUnderline(True)
        self.underlinedButton.setFont(font)
        self.underlinedButton.setObjectName("U")
        self.underlinedButton.setFixedSize(30,23)
        self.underlinedButton.clicked.connect(self.font_event)
        
        Hlayout_bar.addWidget(self.boldButton) 
        Hlayout_bar.addWidget(self.italicButton) 
        Hlayout_bar.addWidget(self.underlinedButton) 
        Hlayout_bar.addSpacing(135)
        
        self.vlayout.addLayout(Hlayout_bar)
        
        Hlayout_body = QHBoxLayout()
        self.body = QTextEdit()
        self.fontBodyFamily = "Tahoma"
        self.fontBodyPointSize = 8
        self.body.setFontFamily(self.fontBodyFamily)
        self.body.setFontPointSize(self.fontBodyPointSize)
        self.body.cursorPositionChanged.connect(self.updateFont)
        Hlayout_body.addWidget(self.body)
        Hlayout_body.addSpacing(10)
        
        self.vlayout.addLayout(Hlayout_body)
        self.vlayout.addSpacing(10)
        
        Hlayout_button = QHBoxLayout()
        self.sendButton = QPushButton('Envoyer', self)
        self.sendButton.clicked.connect(self.send_event)
        Hlayout_button.addSpacing(100)
        Hlayout_button.addWidget(self.sendButton)
        Hlayout_button.addSpacing(100)
        self.vlayout.addLayout(Hlayout_button)
        
    def onTextChanged(self):
        self.dest_file.textChanged.disconnect()
        self.dest_file.cursorPositionChanged.disconnect()
        self.dest_file.setFont(QFont())
        self.dest_file.setText('')
        
    def specialCases(self, body, mail):
        
        i = body.find('%prenom')
        
        while i != -1:
            a = mail.find(".")
            name = mail[0:1].upper() + mail[1:a]
            body = body[:i] + name + body[i+7:]
            i = body.find('%prenom')
        
        i = body.find('%nom')
        
        while i != -1:
            a = mail.find(".")
            b = mail.find("@")
            name = mail[a+1:a+2].upper() + mail[a+2:b]
            body = body[:i] + name + body[i+4:]
            i = body.find('%nom')

        return body
        
    def information(self, titleMsg):
        QMessageBox(QMessageBox.Information,self.title,titleMsg,QMessageBox.NoButton,self).exec_()

    def warning(self, titleMsg):
        QMessageBox(QMessageBox.Warning,self.title,titleMsg,QMessageBox.NoButton,self).exec_()

    def updateFont(self):        
        #print(self.fontBodyFamily + " " + str(self.fontBodyPointSize))
        if self.body.fontWeight() == 75:
            self.boldButton.setChecked(True)
        else:
            self.boldButton.setChecked(False)
        
        if self.body.fontItalic():
            self.italicButton.setChecked(True)
        else:
            self.italicButton.setChecked(False)
        
        if self.body.fontUnderline():
            self.underlinedButton.setChecked(True)
        else:
            self.underlinedButton.setChecked(False)
        
        if len(self.body.toPlainText()) > 0:
            #print("true")
            #self.fontFamilyMenu.setText(self.fontBodyFamily)
            #self.fontSizeMenu.setText(str(self.fontBodyPointSize)+'pt')
            self.body.setFontFamily(self.fontBodyFamily)
            self.body.setFontPointSize(self.fontBodyPointSize)   
            self.fontBodyPointSize = int(self.body.fontPointSize())
            self.fontBodyFamily = self.body.fontFamily()          	
    
    def updateFontFamily(self, event):
        obj = QObject()
        sender = obj.sender()
        name = sender.objectName()
        self.body.setFontFamily(name)
        self.fontFamilyMenu.setText(name)
        self.fontBodyFamily = self.fontFamilyMenu.text()
    
    def updateFontSize(self, event):
        obj = QObject()
        sender = obj.sender()
        name = sender.objectName()
        self.body.setFontPointSize(int(name))
        self.fontSizeMenu.setText(name+'pt')
        self.fontBodyPointSize = int(name)

    
    def font_event(self, event):
        obj = QObject()
        sender = obj.sender()
        name = sender.objectName()
        if name == "B":
            if self.body.fontWeight() == 50:
                self.body.setFontWeight(75)
            else :
                self.body.setFontWeight(50)
        elif name == "I":
            if self.body.fontItalic():
                self.body.setFontItalic(False)
            else :
                self.body.setFontItalic(True)
        elif name == "U":
            if self.body.fontUnderline():
                self.body.setFontUnderline(False)
            else :
                self.body.setFontUnderline(True)	
    
    def pj_event(self, event):
        self.pj()
    
    def pj(self):
        filenames = QFileDialog.getOpenFileNames(self, 'Select one or more files to attach')[:-1]
        for files in filenames:
            for filename in files:
                print(filename)
                filename = str(filename)
                self.attachments.append(filename)
                a = filename.rfind('/')
                self.pj_display.setText(self.pj_display.text() + "\"" + filename[a+1:] +"\";")
		
    def send_event(self, event):
        self.send()

    def send(self):        
        fromaddr = self.email 
        password = self.mdp

        locked = False
        
        try:
            mails = str(self.dest_file.toPlainText()).splitlines()
        except UnicodeEncodeError as err:
            self.warning('Une adresse mail destinataire comporte un caractère invalide !')
            locked = True
        
        if locked == False:
            for mail in mails:
                print(mail)
                a = mail.find("@")
                b = mail.rfind(".")
                if a<1 or b-a<2:
                    self.warning('L\'adresse \"'+str(mail)+'\" de votre destinataire est invalide !')
                    locked = True

        if locked == False:
            if len(self.dest_file.toPlainText()) == 0:
                self.warning('Veuillez rentrer au moins un destinataire !')
                locked = True

        if locked == False:
            count = 0
            
            self.progressBar = QProgressDialog("Envoi en cours...","cancel",0,len(mails),self)
            self.progressBar.setCancelButton(None)
            self.progressBar.setWindowTitle(self.title)

            for mail in mails:
                msg = MIMEMultipart()
                msg['From'] = fromaddr
                msg['To'] = mail
                subject = self.obj_display.text()
                subject = self.specialCases(subject,mail)
                msg['Subject'] = Header(subject, 'utf-8')
                body = self.body.toPlainText()
                body = self.specialCases(body,mail)
                msg.attach(MIMEText(body.encode('utf-8'), 'plain', 'utf-8'))

                print(self.attachments)

                for filename in self.attachments:
                    attachment = open(filename, "rb")
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload((attachment).read())
                    encoders.encode_base64(part)
                    a = filename.rfind('/')
                    part.add_header('Content-Disposition', "attachment; filename= %s" % filename[a+1:])
                    msg.attach(part)

                server = smtplib.SMTP(self.serverSMTP, self.serverPORT)
                server.starttls()
                server.login(fromaddr, password)
                text = msg.as_string()
                print(mail)
                server.sendmail(fromaddr, mail, text)
                server.quit()
                count = count + 1
                self.progressBar.setValue(count)
                self.progressBar.show()
                self.progressBar.activateWindow()
                QApplication.processEvents()
                print(mail+" Success ! "+str(count))
            
            self.progressBar.hide()
            self.information('Votre email a été envoyé à vos '+str(count)+' destinataires avec succès')
            self.dest_file.setText('')
            self.obj_display.setText('') 
            self.pj_display.setText('')
            self.body.setText('')
            self.attachments = []
          
if __name__ == '__main__':
	
    app = QApplication(sys.argv)
    
    email = "ali.zakaria@ensea.fr"
    mdp = "Code2173+"
    processGui = ProcessGui("OneEmail",email,mdp,"smtp2.ensea.fr",587)
    processGui.show()

    sys.exit(app.exec_())
    

#!/usr/bin/env python

import sys
import PyQt5
from welcomeGui import WelcomeGui

  
if __name__ == '__main__':
	
    app = PyQt5.QtWidgets.QApplication(sys.argv)
    
    gui = WelcomeGui("OneEmail")     
    gui.show()

    sys.exit(app.exec_())
    

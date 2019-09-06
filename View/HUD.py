import sys
from PySide.QtCore import *
from PySide.QtGui import *
from Model import MainModule

#Layout Open
app = QApplication(sys.argv)
widget = QWidget()

inputPath = "none"
TargetPath = "none"

#Initialize
widget.setGeometry(250, 200, 550, 135)
widget.setFixedSize(550, 135)
widget.setWindowTitle('XN-1000 Excel Convert (ver2.0)')

#Widget
headerText = QLabel(widget)
headerText.setText("Note: select CSV file")

inputFile = QFileDialog(widget)
inputTextField = QLineEdit(widget)
inputTextField.setText("XN_00-22 (Build 63)_SAMPLE.csv")
inputBth = QPushButton('Select File', widget)
labelStatus = QLabel(widget)
labelStatus.setText("Please select input file.")
procressBth = QPushButton('Process', widget)
procressBth.setDisabled(1)
msg = QMessageBox()
msg.setIcon(QMessageBox.Information)
msg.setText("<Blank>")

#Locate
headerText.setGeometry(10,0,380,30)
inputBth.setGeometry(10,30,100,30)
inputTextField.setGeometry(120,30,420,30)
procressBth.setGeometry(10, 65, 530, 45)
labelStatus.setGeometry(10, 110, 380, 20)
inputFile.setGeometry(300, 200, 400, 300)
msg.setGeometry(300,300,200,100)


def browseFile():
    global inputPath
    inputPath = inputFile.getOpenFileName()[0]
    inputTextField.setText(inputPath)
    inputTextField.setDisabled(1)
    changeLabelStatus()
    print(inputPath)
    
def changeLabelStatus():
    if inputPath!= "none":
        labelStatus.setText("you already select excel file")
        procressBth.setDisabled(0)
    elif inputPath == "none" or inputPath == '':
        labelStatus.setText("please select excel file.")
        procressBth.setDisabled(1)
    else:
        labelStatus.setText("some things wrong")
        
def Execute():
    Status = MainModule.Runtime(inputPath)
    labelStatus.setText("Process Done")
     
    if Status == "Conversion Successful.":
        msg.setIcon(QMessageBox.Information)
    else: 
        msg.setIcon(QMessageBox.Critical)
    msg.setText(Status)
    msg.show()
    
#Listener
widget.connect(inputBth,SIGNAL('clicked()'), browseFile)
widget.connect(procressBth,SIGNAL('clicked()'), Execute)

#Layout Close
widget.show()
sys.exit(app.exec_())
        